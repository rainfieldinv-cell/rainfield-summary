# -*- coding: utf-8 -*-
"""요약본 생성기 자립용 헬퍼.

원래는 옆 저장소(rainfield-im)의 frame_builders(141KB)·content_parser(57KB) 에서
함수 하나씩만 끌어다 썼다. 요약본은 IM 자동화 폴더 안에서 독립적으로 돌아야 하므로
**필요한 함수만** 여기로 옮겼다.
  - replace_text_keep_runs : 글상자 서식(폰트·크기·색)을 유지한 채 글자만 교체
  - extract_page_images    : PDF 한 페이지의 이미지 추출(작은 아이콘·반복 로고 제외)
"""
from collections import Counter

_IMG_MIN_PX = 100 * 100        # 이보다 작으면 아이콘·불릿으로 보고 버림
_REPEAT_PAGE_MIN = 3           # 이 페이지 수 이상에 나오면 머리말 로고·워터마크


def replace_text_keep_runs(tf, new_text: str):
    """텍스트프레임의 글자만 바꾸고 서식은 그대로 둔다.

    첫 run 의 서식을 유지한 채 텍스트를 넣고 나머지 run 은 지운다.
    여러 줄(\\n)이면 줄바꿈을 살려 단락을 만든다.
    """
    lines = str(new_text if new_text is not None else "").split("\n")
    paras = tf.paragraphs
    if not paras:
        tf.text = "\n".join(lines)
        return

    p0 = paras[0]
    if p0.runs:
        p0.runs[0].text = lines[0]
        for r in p0.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        # ★run 이 없는 빈 칸: 그냥 p0.text 를 쓰면 서식(글자 크기·폰트)을 잃어
        #   글씨가 기본 18pt 로 튀어 표를 뭉갠다. 같은 표의 다른 칸 서식을 물려받는다.
        p0.text = lines[0]
        try:
            from copy import deepcopy
            src = None
            for p in paras[1:]:
                if p.runs:
                    src = p.runs[0]
                    break
            if src is None:                       # 같은 표의 다른 셀에서 찾기
                tbl = getattr(getattr(tf, "_parent", None), "_parent", None)
                holder = getattr(tbl, "_tbl", None)
                if holder is not None:
                    for rpr in holder.iter():
                        if rpr.tag.endswith('}rPr'):
                            for r in p0.runs:
                                r._r.insert(0, deepcopy(rpr))
                            break
            elif p0.runs:
                p0.runs[0]._r.insert(0, deepcopy(src._r.get_or_add_rPr()))
        except Exception:
            pass

    for p in paras[1:]:                       # 남은 단락 제거
        p._p.getparent().remove(p._p)

    for ln in lines[1:]:                      # 둘째 줄부터 새 단락(서식 복사)
        import copy as _copy
        new_p = _copy.deepcopy(p0._p)
        p0._p.getparent().append(new_p)
        from pptx.text.text import _Paragraph
        np = _Paragraph(new_p, p0._parent)
        if np.runs:
            np.runs[0].text = ln
            for r in np.runs[1:]:
                r._r.getparent().remove(r._r)


def repeated_xrefs(doc, min_pages: int = _REPEAT_PAGE_MIN) -> set:
    """여러 페이지에 반복 등장하는 이미지 = 발행사 로고·워터마크."""
    use = Counter()
    try:
        for pg in doc:
            for x in {im[0] for im in pg.get_images(full=True)}:
                use[x] += 1
    except Exception:
        return set()
    if doc.page_count < min_pages:
        return set()
    return {x for x, c in use.items() if c >= min_pages}


def extract_page_images(doc, page_idx: int, skip_xrefs=None) -> list:
    """PDF 한 페이지의 이미지를 [{xref,width,height,data}] 로 반환.

    작은 아이콘(_IMG_MIN_PX 미만)과 skip_xrefs(로고·워터마크)는 제외한다.
    투명 마스크(SMask)를 합쳐 '보이는 그대로' 만든다(안 하면 투명부가 검게 나옴).
    """
    import fitz
    page = doc[page_idx]
    out, seen = [], set()
    skip_xrefs = skip_xrefs or set()
    for img in page.get_images(full=True):
        xref = img[0]
        smask = img[1] if len(img) > 1 else 0
        if xref in seen or xref in skip_xrefs:
            continue
        seen.add(xref)
        try:
            pix = fitz.Pixmap(doc, xref)
            if smask:
                try:
                    if pix.alpha:
                        pix = fitz.Pixmap(pix, 0)
                    pix = fitz.Pixmap(pix, fitz.Pixmap(doc, smask))
                except Exception:
                    pass
            if (pix.n - pix.alpha) >= 4 or (
                    pix.colorspace and pix.colorspace.name not in ("DeviceRGB", "DeviceGray")):
                pix = fitz.Pixmap(fitz.csRGB, pix)
            if pix.width * pix.height < _IMG_MIN_PX:
                continue
            out.append({"xref": xref, "width": pix.width, "height": pix.height,
                        "data": pix.tobytes("png")})
        except Exception:
            continue
    return out


def page_png(pdf_bytes: bytes, page_no: int, zoom: float = 2.0):
    """PDF 한 페이지를 PNG bytes 로 렌더(1부터 세는 페이지 번호). 실패 시 None."""
    try:
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if not (1 <= page_no <= doc.page_count):
            doc.close()
            return None
        pix = doc[page_no - 1].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        png = pix.tobytes("png")
        doc.close()
        return png
    except Exception:
        return None


def page_count(pdf_bytes: bytes) -> int:
    try:
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        n = doc.page_count
        doc.close()
        return n
    except Exception:
        return 0


def find_es_pages(pdf_bytes: bytes) -> list:
    """'Executive Summary'(또는 하이라이트) 가 있는 페이지 번호 목록(1부터)."""
    KEYS = ("executive summary", "executive", "투자 포인트", "투자포인트",
            "핵심 요약", "highlight")
    out = []
    try:
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for i, pg in enumerate(doc, start=1):
            head = (pg.get_text("text") or "")[:400].lower()
            if any(k in head for k in KEYS):
                out.append(i)
        doc.close()
    except Exception:
        pass
    return out


def pptx_slide_png(pptx_path: str, slide_no: int = 1, w: int = 1600, h: int = 1108):
    """PPTX 한 장을 PNG bytes 로 (PowerPoint COM). 실패 시 (None, 사유).

    이 PC엔 LibreOffice 가 없어 PowerPoint 로 내보낸다.
    """
    import os as _os
    import tempfile as _tf
    try:
        import win32com.client
    except Exception:
        return None, "pywin32(win32com)가 없어 미리보기를 만들 수 없습니다."
    out = _os.path.join(_tf.mkdtemp(prefix="rf_hl_"), "s.png")
    app = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            pres = app.Presentations.Open(_os.path.abspath(pptx_path), WithWindow=False)
        except Exception:
            pres = app.Presentations.Open(_os.path.abspath(pptx_path))
        n = pres.Slides.Count
        if not (1 <= slide_no <= n):
            pres.Close()
            return None, f"슬라이드 {slide_no} 없음(총 {n}장)"
        pres.Slides(slide_no).Export(out, "PNG", w, h)
        pres.Close()
        with open(out, "rb") as f:
            return f.read(), None
    except Exception as e:
        return None, f"PowerPoint 변환 실패: {e}"
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass


def mean_brightness(data: bytes) -> float:
    """이미지 평균 밝기(0~255). 사진(조감도)은 어둡고, 도면·지도는 흰 바탕이라 밝다."""
    try:
        from PIL import Image
        import io as _io
        im = Image.open(_io.BytesIO(data)).convert("RGB")
        im.thumbnail((80, 80))
        px = list(im.getdata())
        if not px:
            return 255.0
        return sum((r + g + b) / 3 for r, g, b in px) / len(px)
    except Exception:
        return 255.0
