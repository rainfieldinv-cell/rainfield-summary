# -*- coding: utf-8 -*-
"""요약본 자동 생성기 (로컬 웹앱) — 단계별 진행.

기존 IM 변환기와 같은 방식으로 한 화면에 하나씩:
  1 원본 업로드  →  2 하이라이트  →  3 내용  →  4 생성
"""
import os, sys, tempfile
import streamlit as st

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from summary_pipeline import (extract_summary, build_summary,      # noqa: E402
                              build_highlight_preview)
try:      # 배포 직후 옛 모듈이 남아 있어도 앱 전체가 죽지 않게 한다
    from summary_pipeline import extra_blocks_of                   # noqa: E402
except ImportError:
    def extra_blocks_of(_data):
        return []
from engine_bits import (page_png, page_count, find_es_pages,      # noqa: E402
                         pptx_slide_png)

STEP_NAMES = ["원본 업로드", "하이라이트", "내용", "생성"]

st.set_page_config(page_title="Rainfield 요약본 생성기", layout="wide")


# ──────────────────────────────────────────────────
# 스텝퍼 (ui_components.render_stepper 와 같은 모양)
# ──────────────────────────────────────────────────
def render_stepper(current: int):
    st.markdown("""
    <style>
    .stp-wrap{display:flex;align-items:center;justify-content:center;flex-wrap:wrap;
              gap:2px;padding:12px 0 16px 0;}
    .stp-item{display:flex;flex-direction:column;align-items:center;min-width:82px;max-width:110px;}
    .stp-circle{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;
                justify-content:center;font-size:13px;font-weight:bold;margin-bottom:4px;}
    .stp-done{background:#1a7f37;color:#fff;}
    .stp-active{background:#08377C;color:#fff;box-shadow:0 0 0 3px #cdd4e6;}
    .stp-todo{background:#e5e7eb;color:#9ca3af;}
    .stp-label{font-size:11px;text-align:center;line-height:1.3;word-break:keep-all;}
    .stp-l-done{color:#1a7f37;font-weight:600;}
    .stp-l-active{color:#08377C;font-weight:700;}
    .stp-l-todo{color:#9ca3af;}
    .stp-conn{width:26px;height:2px;margin-bottom:20px;}
    .stp-c-done{background:#1a7f37;} .stp-c-todo{background:#e5e7eb;}
    </style>
    """, unsafe_allow_html=True)
    html = '<div class="stp-wrap">'
    for i, name in enumerate(STEP_NAMES):
        n = i + 1
        if n < current:
            cc, lc, tx, conn = "stp-done", "stp-l-done", "✓", "stp-c-done"
        elif n == current:
            cc, lc, tx, conn = "stp-active", "stp-l-active", str(n), "stp-c-todo"
        else:
            cc, lc, tx, conn = "stp-todo", "stp-l-todo", str(n), "stp-c-todo"
        html += (f'<div class="stp-item"><div class="stp-circle {cc}">{tx}</div>'
                 f'<div class="stp-label {lc}">{name}</div></div>')
        if n < len(STEP_NAMES):
            html += f'<div class="stp-conn {conn}"></div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)
    st.markdown("---")


def _goto(n):
    st.session_state["step"] = n
    st.rerun()


st.session_state.setdefault("step", 1)
step = st.session_state["step"]

# 기존 IM 변환기와 같은 회사 톤(네이비 08377C + 화이트)
st.markdown("""
<style>
/* ★화면 전체 폭을 쓴다. 전엔 max-width:1200px 로 묶어 놔서
   원본 이미지가 작아 글씨가 안 보이고 카드도 좁게 눌렸다. */
.block-container{padding-top:1.6rem;max-width:100%;
                 padding-left:2.2rem;padding-right:2.2rem;}
h1,h2,h3{color:#08377C;}
/* ★버튼은 전부 남색으로 통일(기본 IM 변환기와 동일) */
div.stButton>button{border-radius:4px;border:1px solid #08377C;color:#08377C;
                    background:#fff;font-weight:700;}
div.stButton>button:hover{background:#eef3fa;color:#08377C;border-color:#08377C;}
div.stButton>button[kind="primary"]{background:#1A2B5E;color:#fff;border-color:#1A2B5E;}
div.stButton>button[kind="primary"]:hover{background:#08377C;border-color:#08377C;}
div.stDownloadButton>button{background:#08377C;color:#fff;border:1px solid #08377C;
                            border-radius:4px;font-weight:700;}
div.stDownloadButton>button:hover{background:#0063A1;border-color:#0063A1;}
hr{border-color:#dfe6f0;}
</style>
""", unsafe_allow_html=True)

st.title("Rainfield 요약본 자동 생성기")
st.caption("원본 IM(PDF) → 4~5장 요약본")
render_stepper(step)
data = st.session_state.get("data")


# ──────────────────────────────────────────────────
# 1단계 — 원본 업로드
# ──────────────────────────────────────────────────
if step == 1:
    st.markdown("### 1단계 · 원본 IM 업로드")
    st.caption("PDF를 올리면 원본을 읽어 핵심 데이터를 자동으로 뽑습니다. "
               "처음 보는 원본은 1분 남짓 걸리고, 같은 원본은 즉시 나옵니다.")
    up = st.file_uploader("원본 IM PDF", type=["pdf"])

    if up is not None and st.session_state.get("_pdf_name") != up.name:
        st.session_state["_pdf_bytes"] = up.getvalue()
        st.session_state["_pdf_name"] = up.name
        st.session_state.pop("ppt", None)
        with st.spinner("원본을 읽어 핵심 데이터를 추출하는 중..."):
            try:
                d = extract_summary(up.getvalue())
                st.session_state["data"] = d
                hl = [dict(h) for h in (d.get("highlights") or [])][:3]
                while len(hl) < 3:
                    hl.append({"title": "", "subtitle": "", "bullets": []})
                st.session_state["hl"] = hl
                # 다른 PDF를 올렸는데 앞 원본의 카드 글이 남지 않도록 위젯도 갈아끼운다.
                for _i in range(3):
                    _sub = (hl[_i].get("subtitle") or "").strip()
                    st.session_state[f"ht_{_i}"] = hl[_i].get("title", "") or ""
                    st.session_state[f"hs_{_i}"] = _sub
                    st.session_state[f"hb_{_i}"] = "\n".join(hl[_i].get("bullets") or [])
                    st.session_state[f"hu_{_i}"] = bool(_sub)
                st.session_state.pop("hl_prev", None)
            except Exception as e:
                import traceback
                st.error(f"추출 실패: {e}")
                st.code(traceback.format_exc())

    data = st.session_state.get("data")
    if data:
        st.success("추출 완료")
        c1, c2 = st.columns(2)
        c1.metric("딜명", (data.get("deal_name") or "-")[:40])
        c2.metric("날짜", data.get("date_ko") or "-")
        _found = [k for k in ("사모사채개요", "담보대출조건", "사업일정", "사업개요",
                              "투입에쿼티", "법인개요", "재무제표") if data.get(k)]
        st.caption("찾은 항목 : " + (", ".join(_found) if _found else "없음"))
        st.markdown("---")
        _b1, _b2, _b3 = st.columns([2, 4, 2])
        if _b3.button("다음 단계 →", type="primary", use_container_width=True):
            _goto(2)


# ──────────────────────────────────────────────────
# 2단계 — 하이라이트
# ──────────────────────────────────────────────────
elif step == 2:
    st.markdown("## 2단계. 하이라이트")
    st.caption("표지 다음에 들어가는 '핵심 3가지' 카드를 만드는 단계입니다. "
               "왼쪽 원본을 보면서 오른쪽 카드를 고치세요.")
    if not data:
        st.warning("먼저 1단계에서 원본을 올려주세요.")
    else:
        _pdf = st.session_state.get("_pdf_bytes")

        # ★위젯 값을 카드 데이터에 강제로 맞춘다.
        #   예전엔 키를 pop 하고 text_input(value=...) 으로 되돌렸는데,
        #   Streamlit 이 위젯 값을 자체 저장소에도 들고 있어 제목·내용이 되살아났다
        #   ('카드 3개 비우기' 를 눌러도 부제목만 사라지던 원인).
        #   → 세션 상태의 위젯 키에 직접 써 넣으면 확실하게 반영된다.
        def _sync_card_widgets(cards):
            for _i in range(3):
                _h = cards[_i]
                _sub = (_h.get("subtitle") or "").strip()
                st.session_state[f"ht_{_i}"] = _h.get("title", "") or ""
                st.session_state[f"hs_{_i}"] = _sub
                st.session_state[f"hb_{_i}"] = "\n".join(_h.get("bullets") or [])
                st.session_state[f"hu_{_i}"] = bool(_sub)

        hb1, hb2, _ = st.columns([2, 2, 3])
        if hb1.button("🤖 Executive Summary 자동 추출", use_container_width=True,
                      help="IM의 'Executive Summary' 항목만 읽어 3개 카드로 정리합니다"
                           "(그 뒤 직접 수정)."):
            src = data.get("highlights") or []
            hl = [dict(h) for h in src][:3]
            while len(hl) < 3:
                hl.append({"title": "", "subtitle": "", "bullets": []})
            st.session_state["hl"] = hl
            _sync_card_widgets(hl)
            st.rerun()
        if hb2.button("🧹 카드 3개 비우기", use_container_width=True,
                      help="제목·부제목·내용을 카드 3개 모두 비웁니다."):
            blank = [{"title": "", "subtitle": "", "bullets": []} for _ in range(3)]
            st.session_state["hl"] = blank
            _sync_card_widgets(blank)
            st.session_state.pop("hl_prev", None)      # 옛 미리보기 이미지도 버린다
            st.rerun()

        # 자동 추출 결과 안내(기본 IM 변환기와 동일한 초록 배너)
        _pg = (st.session_state.get("es_pages_str") or "").strip()
        if any((h.get("title") or "").strip() for h in st.session_state["hl"]):
            st.success(f"'Executive Summary' 항목({_pg}p)을 3개 카드로 정리했어요. "
                       f"확인·수정하세요.")

        img_col, form_col = st.columns([1.4, 1])

        # ── 왼쪽: 원본 하이라이트 페이지 실제 이미지 ──────
        with img_col:
            st.markdown("#### 📄 원본 Executive Summary")
            if not _pdf:
                st.warning("원본 PDF가 없습니다. 1단계에서 올려주세요.")
            else:
                npage = page_count(_pdf)
                det = find_es_pages(_pdf)
                st.session_state.setdefault(
                    "es_pages_str", ",".join(str(p) for p in det) if det else "2")
                raw = st.text_input(
                    "Executive Summary 페이지 (여러 장이면 2-5 · 한 장이면 3)",
                    key="es_pages_str",
                    help="자동으로 찾은 값이에요. ES가 여러 페이지면 여기서 범위를 고치세요(예: 2-5).")
                pages_sel = []
                for tok in str(raw or "").replace(" ", "").split(","):
                    if "-" in tok:
                        try:
                            a, b = tok.split("-")[:2]
                            pages_sel += list(range(int(a), int(b) + 1))
                        except Exception:
                            pass
                    elif tok.isdigit():
                        pages_sel.append(int(tok))
                pages_sel = [p for p in pages_sel if 1 <= p <= npage]
                if det:
                    st.caption(f"자동 감지: {', '.join(str(p) for p in det)}p "
                               f"(필요하면 위에서 조정)")
                if not pages_sel:
                    st.warning("표시할 페이지가 없습니다. 위에 페이지 번호를 적어주세요.")
                else:
                    idx = min(max(0, st.session_state.get("es_idx", 0)), len(pages_sel) - 1)
                    if len(pages_sel) > 1:
                        p1, p2, p3 = st.columns([1, 2, 1])
                        if p1.button("◀ 이전", use_container_width=True, disabled=(idx <= 0)):
                            st.session_state["es_idx"] = idx - 1
                            st.rerun()
                        p2.markdown(
                            f"<div style='text-align:center;padding-top:6px;'>"
                            f"<b>{pages_sel[idx]}</b>p ({idx+1}/{len(pages_sel)})</div>",
                            unsafe_allow_html=True)
                        if p3.button("다음 ▶", use_container_width=True,
                                     disabled=(idx >= len(pages_sel) - 1)):
                            st.session_state["es_idx"] = idx + 1
                            st.rerun()
                    # zoom 3.0 = 고해상도 렌더(작은 글씨도 읽히게)
                    png = page_png(_pdf, pages_sel[idx], zoom=3.0)
                    if png:
                        st.image(png, use_container_width=True)   # 원본은 그대로(스크롤 X)
                    else:
                        st.caption(f"⚠️ {pages_sel[idx]}p 렌더 실패")

        # ── 오른쪽: 카드 3개 편집 ──────────────────────
        with form_col:
            st.markdown("#### ✅ 하이라이트 카드 3개")
            st.caption("카드 간격은 다음 단계에서 일정하게 배치되고, "
                       "높이는 내용 양에 따라 자동 조절됩니다.")
            # ★스크롤 상자 없이 카드가 그대로 아래로 흐른다(기본 IM 변환기와 동일).
            for i in range(3):
                h = st.session_state["hl"][i]
                with st.container(border=True):
                    st.markdown(f"**카드 {i + 1}**")
                    # ★값은 항상 세션 상태의 위젯 키로만 오간다(value= 를 쓰지 않는다).
                    #   둘을 섞으면 어느 쪽이 이기는지가 갈려 '비우기'가 안 먹었다.
                    _sub0 = (h.get("subtitle") or "").strip()
                    st.session_state.setdefault(f"ht_{i}", h.get("title", "") or "")
                    st.session_state.setdefault(f"hs_{i}", _sub0)
                    st.session_state.setdefault(f"hb_{i}",
                                                "\n".join(h.get("bullets") or []))
                    st.session_state.setdefault(f"hu_{i}", bool(_sub0))

                    h["title"] = st.text_input(
                        "✓ 제목", key=f"ht_{i}", placeholder="예: 낮은 인허가 리스크")
                    # ★부제목은 '선택사항' — 기존 IM 자동화와 동일하게 체크박스로 켠다.
                    h["use_sub"] = st.checkbox("부제목 넣기", key=f"hu_{i}")
                    if h["use_sub"]:
                        h["subtitle"] = st.text_input(
                            "부제목 (하늘색 줄)", key=f"hs_{i}",
                            placeholder="예: 주요 심의·평가 완료 단계")
                    else:
                        h["subtitle"] = ""
                    h["bullets"] = [b for b in st.text_area(
                        "내용", key=f"hb_{i}", height=110,
                        placeholder="카드 본문 (여러 줄 입력 가능)"
                        ).split("\n") if b.strip()]

        # ── 완성본 미리보기 (실제 슬라이드를 만들어 이미지로) ──
        st.markdown("---")
        if st.checkbox("🖼️ 하이라이트 완성본 미리보기", key="hl_preview_on",
                       help="작성한 카드 3개로 실제 하이라이트 슬라이드를 만들어 "
                            "이미지로 보여줍니다.(다운로드 아님 · 화면 확인용)"):
            if st.button("완성본 이미지 생성/갱신", key="hl_prev_gen"):
                with st.spinner("완성본 슬라이드를 만드는 중..."):
                    try:
                        _d = dict(data)
                        _d["highlights"] = st.session_state["hl"]
                        _p = tempfile.NamedTemporaryFile(suffix=".pptx",
                                                         delete=False).name
                        build_highlight_preview(_d, _p)
                        img, err = pptx_slide_png(_p, 1)
                        st.session_state["hl_prev"] = {"img": img, "err": err}
                    except Exception as e:
                        st.session_state["hl_prev"] = {"img": None, "err": str(e)}
            _r = st.session_state.get("hl_prev")
            if not _r:
                st.info("‘완성본 이미지 생성/갱신’을 누르면 완성된 하이라이트가 여기에 표시됩니다.")
            elif _r.get("err"):
                st.error(f"이미지 생성 실패 — {_r['err']}")
            elif _r.get("img"):
                st.markdown("#### ✅ 완성본 하이라이트")
                _i1, _i2, _i3 = st.columns([1, 3, 1])
                with _i2:
                    st.image(_r["img"], use_container_width=True)


        # 하단 상태 표시(기본 IM 변환기와 동일)
        _n = sum(1 for h in st.session_state["hl"]
                 if (h.get("title") or "").strip() or (h.get("bullets") or []))
        _m = sum(1 for h in st.session_state["hl"] if (h.get("subtitle") or "").strip())
        st.success(f"저장됨 · 작성한 카드 {_n}/3개 · 부제목 사용 {_m}개")
    st.markdown("---")
    b1, _sp, b2 = st.columns([2, 4, 2])
    if b1.button("← 이전 단계", use_container_width=True):
        _goto(1)
    if b2.button("다음 단계 →", type="primary", use_container_width=True,
                 disabled=not data):
        _goto(3)


# ──────────────────────────────────────────────────
# 3단계 — 내용
# ──────────────────────────────────────────────────
elif step == 3:
    st.markdown("### 3단계 · 내용 배치")
    st.caption("A4 한 장을 좌·우 두 단으로 나눴습니다. 각 칸에 **무엇을 넣을지** 고르면 "
               "원본에서 그 내용을 찾아 채웁니다.")
    if not data:
        st.warning("먼저 1단계에서 원본을 올려주세요.")
    else:
        # 원본에서 찾은 것만 고를 수 있게
        _loan = data.get("담보대출조건") or {}
        AVAIL = {
            "사모사채개요": bool(data.get("사모사채개요") or _loan),
            "담보대출조건": bool(_loan),
            "대출조건표":   bool(_loan.get("tranches")),   # Tr.A/B 금액·LTV·금리 표(따로 배치 가능)
            "사업일정":     bool(data.get("사업일정")),
            "법인개요":     bool(data.get("법인개요")),
            "재무제표":     bool(data.get("재무제표")),
            "조감도":       bool(data.get("이미지_있음", True)),
        }
        # ★목록은 '기본 7항목' 을 위에, 원본에서 찾은 그 밖의 내용을 그 아래에.
        _SEP = "──── 원본의 그 밖의 내용 ────"      # 고르면 (비움)으로 친다
        EXTRA = [t for t, _p in extra_blocks_of(data)]
        OPTS = ["(비움)"] + [k for k, v in AVAIL.items() if v]
        if EXTRA:
            OPTS += [_SEP] + EXTRA
        MISSING = [k for k, v in AVAIL.items() if not v]

        # ── 페이지 수 먼저 고른다(1장이냐 2장이냐에 따라 아래 배치가 달라짐) ──
        st.session_state.setdefault("pages", 1)
        npages = st.radio("핵심요약 페이지 수", [1, 2], horizontal=True,
                          index=st.session_state["pages"] - 1,
                          format_func=lambda x: f"{x}페이지")
        st.session_state["pages"] = npages
        st.caption("한 장에 안 들어가면 2페이지로 나누세요. 생성 후 넘치면 경고가 뜹니다.")
        st.markdown("---")
        if npages == 2:
            st.markdown("#### 1페이지 (첫째 장)")

        st.session_state.setdefault("slots", {
            "left":  ["사모사채개요", "담보대출조건", "대출조건표"],
            "right": ["조감도", "사업일정", "법인개요", "재무제표"],
        })
        st.session_state.setdefault("slots2", {"left": ["법인개요", "재무제표"],
                                               "right": []})
        S = st.session_state["slots"]
        S2 = st.session_state["slots2"]

        def _clear_slot_widgets(tag, side, n):
            """칸을 옮기거나 지우면 selectbox 위젯 값이 옛 자리에 남아 되살아난다.
               해당 단의 위젯 상태를 지워 목록 값이 그대로 반영되게 한다."""
            for j in range(n + 2):
                st.session_state.pop(f"sl{tag}_{side}_{j}", None)

        # ★칸을 하나씩 추가·수정·삭제·이동한다(칸 수를 한꺼번에 정하지 않음).
        def _slot_editor(store, tag, side, title):
            st.markdown(f"**{title}**")
            seq = store.setdefault(side, [])
            for i in range(len(seq)):
                c1, c2, c3, c4 = st.columns([6, 1, 1, 1])
                cur = seq[i] if seq[i] in OPTS else "(비움)"
                _pick = c1.selectbox(
                    f"{i + 1}번 칸", OPTS, index=OPTS.index(cur),
                    key=f"sl{tag}_{side}_{i}", label_visibility="collapsed")
                seq[i] = "(비움)" if _pick == _SEP else _pick    # 구분선은 고를 수 없다
                if c2.button("↑", key=f"up{tag}_{side}_{i}", disabled=(i == 0),
                             help="위로"):
                    seq[i - 1], seq[i] = seq[i], seq[i - 1]
                    _clear_slot_widgets(tag, side, len(seq))
                    st.rerun()
                if c3.button("↓", key=f"dn{tag}_{side}_{i}",
                             disabled=(i == len(seq) - 1), help="아래로"):
                    seq[i + 1], seq[i] = seq[i], seq[i + 1]
                    _clear_slot_widgets(tag, side, len(seq))
                    st.rerun()
                if c4.button("✕", key=f"rm{tag}_{side}_{i}", help="이 칸 삭제"):
                    seq.pop(i)
                    _clear_slot_widgets(tag, side, len(seq) + 1)
                    st.rerun()
            if st.button("＋ 칸 추가", key=f"add{tag}_{side}",
                         disabled=(len(seq) >= 6), use_container_width=True):
                seq.append("(비움)")
                st.rerun()

        cL, cR = st.columns(2)
        with cL:
            _slot_editor(S, "", "left", "왼쪽 단")
        with cR:
            _slot_editor(S, "", "right", "오른쪽 단")
        nL, nR = len(S["left"]), len(S["right"])

        # ── A4 미리보기 (가로 10.83 x 7.5 비율) ──────────
        def _cells(side, n):
            out = ""
            for i in range(n):
                v = S[side][i] if i < len(S[side]) else "(비움)"
                empty = (v == "(비움)")
                out += (f'<div class="cell{" empty" if empty else ""}">'
                        f'{"" if empty else v}</div>')
            return out

        st.markdown("""
        <style>
        .a4{width:100%;max-width:900px;aspect-ratio:10.83/7.5;border:1px solid #c9d2e3;
            background:#fff;display:flex;gap:8px;padding:10px;box-sizing:border-box;
            margin:6px 0 2px 0;}
        .col{flex:1;display:flex;flex-direction:column;gap:8px;}
        .cell{flex:1;border:1px solid #08377C;background:#eef3fa;color:#08377C;
              border-radius:3px;display:flex;align-items:center;justify-content:center;
              font-size:13px;font-weight:700;text-align:center;}
        .cell.empty{border:1px dashed #c9d2e3;background:#fafbfc;color:#c9d2e3;}
        </style>
        """, unsafe_allow_html=True)
        st.markdown(
            f'<div class="a4"><div class="col">{_cells("left", nL)}</div>'
            f'<div class="col">{_cells("right", nR)}</div></div>',
            unsafe_allow_html=True)
        st.caption("실제 높이는 내용 양에 맞춰 자동으로 조절됩니다. "
                   "위 그림은 어디에 무엇이 들어가는지 순서만 보여줍니다.")

        if MISSING:
            st.info("원본에서 못 찾아 고를 수 없는 항목 : " + ", ".join(MISSING))

        # ── 2페이지면 두 번째 장도 배치 ────────────────
        if npages == 2:
            st.markdown("---")
            st.markdown("#### 2페이지 (둘째 장)")
            d1, d2 = st.columns(2)
            with d1:
                _slot_editor(S2, "2", "left", "왼쪽 단")
            with d2:
                _slot_editor(S2, "2", "right", "오른쪽 단")
            nL2, nR2 = len(S2["left"]), len(S2["right"])

            def _cells2(side, n):
                out = ""
                for i in range(n):
                    v = S2[side][i] if i < len(S2[side]) else "(비움)"
                    empty = (v == "(비움)")
                    out += (f'<div class="cell{" empty" if empty else ""}">'
                            f'{"" if empty else v}</div>')
                return out or '<div class="cell empty"></div>'

            st.markdown(
                f'<div class="a4"><div class="col">{_cells2("left", nL2)}</div>'
                f'<div class="col">{_cells2("right", nR2)}</div></div>',
                unsafe_allow_html=True)
            _dup = set(x for x in S["left"] + S["right"] if x != "(비움)") & \
                   set(x for x in S2["left"] + S2["right"] if x != "(비움)")
            if _dup:
                st.warning("1·2페이지에 같은 항목이 있습니다 : " + ", ".join(_dup)
                           + " — 한쪽에서 빼주세요(둘 다 만들면 내용이 중복됩니다).")
    st.markdown("---")
    b1, _sp2, b2 = st.columns([2, 4, 2])
    if b1.button("← 이전 단계", use_container_width=True):
        _goto(2)
    if b2.button("다음 단계 →", type="primary", use_container_width=True, disabled=not data):
        _goto(4)


# ──────────────────────────────────────────────────
# 4단계 — 생성
# ──────────────────────────────────────────────────
elif step == 4:
    st.markdown("### 4단계 · 생성")
    if not data:
        st.warning("먼저 1단계에서 원본을 올려주세요.")
    else:
        st.write(f"**{data.get('deal_name','')}**  ·  {data.get('date_ko','')}  ·  "
                 f"핵심요약 {st.session_state.get('pages', 1)}페이지")
        if st.button("생성", type="primary"):
            data["highlights"] = st.session_state.get("hl") or []
            # 3단계에서 정한 칸 배치 → 빌더가 이 순서대로 쌓는다('(비움)' 은 제외)
            _s = st.session_state.get("slots") or {}
            _s2 = st.session_state.get("slots2") or {}
            data["_slots"] = {
                side: [v for v in _s.get(side, []) if v and v != "(비움)"]
                for side in ("left", "right")
            }
            data["_slots2"] = {
                side: [v for v in _s2.get(side, []) if v and v != "(비움)"]
                for side in ("left", "right")
            }
            with st.spinner("요약본을 만드는 중..."):
                try:
                    tf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
                    tf.write(st.session_state["_pdf_bytes"]); tf.close()
                    out = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False).name
                    build_summary(data, tf.name, out,
                                  pages=st.session_state.get("pages", 1))
                    with open(out, "rb") as f:
                        st.session_state["ppt"] = f.read()
                except Exception as e:
                    import traceback
                    st.session_state.pop("ppt", None)
                    st.error(f"생성 실패: {e}")
                    st.code(traceback.format_exc())

        # ★다운로드 버튼은 생성 블록 '바깥'에서 세션 값으로 그린다.
        #   버튼을 누르면 Streamlit 이 페이지를 다시 그리는데, 결과가 블록 안
        #   지역변수면 그때 사라져서 다시 생성해야 하는 문제가 생긴다.
        if st.session_state.get("ppt"):
            st.success("생성 완료")
            _nm = (data.get("deal_name") or "요약본")[:40]
            st.download_button(
                "다운로드", data=st.session_state["ppt"],
                file_name=f"[Rainfield]_{_nm}_요약본.pptx",
                mime="application/vnd.openxmlformats-officedocument."
                     "presentationml.presentation",
                type="primary", key="dl")
    st.markdown("---")
    _p1, _p2, _p3 = st.columns([2, 4, 2])
    if _p1.button("← 이전 단계", use_container_width=True):
        _goto(3)
