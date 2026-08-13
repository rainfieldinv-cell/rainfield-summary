"""
extractors.py
─────────────────────────────────────────────────────────
PDF 및 Word 파일에서 텍스트와 이미지를 추출하는 함수 모음.
app.py 에서 import 해서 사용합니다.
─────────────────────────────────────────────────────────
"""

import io
import fitz          # PyMuPDF — PDF 파싱
from docx import Document        # python-docx — Word 파싱
from PIL import Image            # Pillow — 이미지 변환


# ─────────────────────────────────────────────
# [설정값]
# 여기를 수정하면 동작 기준이 바뀝니다
# ─────────────────────────────────────────────

# 이미지의 짧은 쪽(min dimension)이 이 값 미만이면 로고/배너로 간주해 제외
# 세로가 좁은 배너나 가로가 좁은 로고를 걸러냅니다
MIN_IMAGE_SIZE = 150

# 전체 텍스트 글자 수가 이 미만이면 "스캔 PDF" 경고를 띄움
SCAN_PDF_TEXT_THRESHOLD = 200


# ─────────────────────────────────────────────
# [PDF 추출 함수]
# ─────────────────────────────────────────────
def extract_from_pdf(file_bytes: bytes) -> dict:
    """
    PDF 파일 바이트에서 텍스트와 이미지를 추출합니다.

    반환값 구조:
    {
        'file_type': 'pdf',
        'total_pages': int,
        'pages_text': [페이지1텍스트, 페이지2텍스트, ...],
        'images': [
            {'index': 0, 'page': 1, 'pil_image': PIL.Image, 'width': int, 'height': int},
            ...
        ],
        'full_text': str,
        'warnings': [경고문자열, ...]   # 처리 중 발생한 경고 메시지
    }

    예외:
        ValueError — 비밀번호 걸린 PDF
        RuntimeError — 손상된 파일 또는 파싱 불가
    """
    warnings = []

    # PDF 파일 열기
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        raise RuntimeError("파일을 읽을 수 없습니다. 파일이 손상되었거나 지원하지 않는 형식입니다.") from e

    # 비밀번호 걸린 PDF 감지
    if doc.is_encrypted:
        doc.close()
        raise ValueError("비밀번호가 걸린 PDF는 처리할 수 없습니다.")

    total_pages = doc.page_count
    pages_text = []
    images = []
    image_index = 0

    for page_num in range(total_pages):
        page = doc[page_num]

        # ── 텍스트 추출 ──
        text = page.get_text("text")
        pages_text.append(text)

        # ── 이미지 추출 ──
        try:
            img_list = page.get_images(full=True)
            for img_info in img_list:
                xref = img_info[0]  # 이미지 참조 번호
                try:
                    pil_img = _extract_image_hires(doc, page, xref)
                    w, h = pil_img.size

                    # 너무 작은 이미지(로고/아이콘 추정)는 제외
                    # 짧은 쪽이 MIN_IMAGE_SIZE 미만이면 로고/배너로 간주
                    if min(w, h) < MIN_IMAGE_SIZE:
                        continue

                    images.append({
                        "index": image_index,
                        "page": page_num + 1,   # 1부터 시작하는 페이지 번호
                        "pil_image": pil_img,
                        "width": w,
                        "height": h,
                    })
                    image_index += 1

                except Exception:
                    # 개별 이미지 추출 실패 — 해당 이미지만 건너뜀
                    continue

        except Exception:
            # 페이지 이미지 목록 조회 실패 — 텍스트는 이미 있으니 경고만
            warnings.append(f"{page_num + 1}페이지 이미지 추출 중 오류가 발생했습니다.")

    doc.close()

    full_text = "\n".join(pages_text)

    # 스캔 PDF 경고 (텍스트가 거의 없는 경우)
    # 여기를 수정하면 경고 기준 글자 수가 바뀝니다
    if len(full_text.strip()) < SCAN_PDF_TEXT_THRESHOLD:
        warnings.append(
            "텍스트 추출 결과가 매우 적습니다. 스캔 PDF는 정확도가 떨어질 수 있습니다."
        )

    if image_index == 0 and len(warnings) > 0:
        warnings.append("이미지 추출 중 일부 오류가 있었습니다.")

    return {
        "file_type": "pdf",
        "total_pages": total_pages,
        "pages_text": pages_text,
        "images": images,
        "full_text": full_text,
        "warnings": warnings,
    }


# ─────────────────────────────────────────────
# [고해상도 이미지 추출 헬퍼]
# ─────────────────────────────────────────────
def _extract_image_hires(doc, page, xref, min_px: int = 1200):
    """
    PDF에서 이미지를 고해상도로 추출합니다.

    1단계: doc.extract_image(xref) 로 원본 이미지 추출
    2단계: 원본이 min_px 미만이면, 페이지에서 해당 이미지 영역을
           3배율(216 dpi)로 렌더링해서 선명한 버전으로 교체합니다.

    여기서 min_px 를 수정하면 고해상도 렌더링 기준이 바뀝니다.
    """
    # ── 1단계: 원본 embedded 이미지 추출 ──────────────────
    base_image = doc.extract_image(xref)
    img_bytes  = base_image["image"]
    pil_img    = Image.open(io.BytesIO(img_bytes)).convert("RGB")

    # 이미 충분한 해상도면 바로 반환
    if max(pil_img.width, pil_img.height) >= min_px:
        return pil_img

    # ── 2단계: 이미지 bbox 를 3배율로 렌더링 ──────────────
    try:
        # 페이지 내 해당 이미지의 위치(rect) 조회
        rects = page.get_image_rects(xref)
        if rects:
            clip    = fitz.Rect(rects[0])
            zoom    = 3.0
            pixmap  = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip)
            hires   = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
            # 렌더링 결과가 원본보다 크면 교체
            if max(hires.width, hires.height) > max(pil_img.width, pil_img.height):
                return hires
    except Exception:
        pass  # 실패 시 원본 사용

    return pil_img


# ─────────────────────────────────────────────
# [Word 추출 함수]
# ─────────────────────────────────────────────
def extract_from_docx(file_bytes: bytes) -> dict:
    """
    Word(.docx) 파일 바이트에서 텍스트와 이미지를 추출합니다.

    Word는 '페이지' 개념이 명확하지 않으므로
    전체 텍스트를 하나의 페이지로 처리합니다.
    (STEP 5에서 AI가 의미 단위로 분리할 예정)
    """
    warnings = []

    try:
        doc = Document(io.BytesIO(file_bytes))
    except Exception as e:
        raise RuntimeError("파일을 읽을 수 없습니다. 파일이 손상되었거나 지원하지 않는 형식입니다.") from e

    # ── 텍스트 추출: 모든 문단을 합칩니다 ──
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    full_text = "\n".join(paragraphs)

    # ── 이미지 추출: 파일 내부의 관련 파트에서 이미지 찾기 ──
    images = []
    image_index = 0

    try:
        for rel in doc.part.rels.values():
            # 이미지 관계(relationship)만 필터링
            if "image" in rel.reltype:
                try:
                    img_bytes = rel.target_part.blob
                    pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
                    w, h = pil_img.size

                    # 너무 작은 이미지 제외 (짧은 쪽 기준)
                    if min(w, h) < MIN_IMAGE_SIZE:
                        continue

                    images.append({
                        "index": image_index,
                        "page": 1,          # Word는 페이지 구분 없이 1로 통일
                        "pil_image": pil_img,
                        "width": w,
                        "height": h,
                    })
                    image_index += 1

                except Exception:
                    continue

    except Exception:
        warnings.append("이미지 추출 중 일부 오류가 있었습니다.")

    # 스캔/이미지 전용 Word 경고
    if len(full_text.strip()) < SCAN_PDF_TEXT_THRESHOLD:
        warnings.append(
            "텍스트 추출 결과가 매우 적습니다. 내용을 확인해주세요."
        )

    return {
        "file_type": "docx",
        "total_pages": 1,           # Word는 1페이지로 통일
        "pages_text": [full_text],  # 전체를 하나의 페이지로 처리
        "images": images,
        "full_text": full_text,
        "warnings": warnings,
    }


# ─────────────────────────────────────────────
# [사업명 자동 감지 함수]
# ─────────────────────────────────────────────
def detect_business_name(pages_text: list) -> str:
    """
    추출된 텍스트의 앞 3페이지에서 사업명을 자동으로 찾아 반환합니다.
    찾지 못하면 빈 문자열을 반환합니다.

    인식 패턴 (여기를 수정하면 인식 패턴이 바뀝니다):
    - OOO 도시개발사업
    - OOO 개발사업
    - OOO 신축공사
    - OOO 리모델링
    - OOO 토지담보대출
    - OOO 주택사업
    - OOO 분양사업
    """
    import re

    # 앞 3페이지만 검사 (여기를 수정하면 검사 페이지 수가 바뀝니다)
    search_pages = pages_text[:3]
    search_text = "\n".join(search_pages)

    # 사업명 패턴 목록 (여기에 패턴을 추가하면 더 많은 유형을 인식합니다)
    patterns = [
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*도시개발사업",
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*개발사업",
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*신축공사",
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*리모델링",
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*토지담보대출",
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*주택사업",
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*분양사업",
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*재개발사업",
        r"[가-힣\w\s·\(\)（）\-]{2,30}?\s*재건축사업",
    ]

    for pattern in patterns:
        match = re.search(pattern, search_text)
        if match:
            # 앞뒤 공백 제거 후 반환
            return match.group(0).strip()

    return ""


# ─────────────────────────────────────────────
# [파일 종류 판별 함수]
# ─────────────────────────────────────────────
def get_file_type(filename: str) -> str:
    """
    파일 이름에서 확장자를 보고 'pdf' 또는 'docx'를 반환합니다.
    지원하지 않는 형식이면 None을 반환합니다.
    """
    lower = filename.lower()
    if lower.endswith(".pdf"):
        return "pdf"
    elif lower.endswith(".docx"):
        return "docx"
    return None
