"""
claude_api.py
─────────────────────────────────────────────────────────
Claude API 공통 통신 모듈.
슬라이드별 프롬프트 로직은 이 모듈에 포함하지 않습니다.

주요 기능:
  - API 키를 .env / 환경변수에서 로드
  - temperature=0 고정 (재현 가능성)
  - JSON 출력 강제 + 파싱 실패 시 재시도
  - 숫자 환각 검증 (PDF 원문 숫자 집합과 대조)
  - SHA256 기반 디스크 캐시 (비용 절감)
  - 토큰 사용량·비용 로깅
─────────────────────────────────────────────────────────
"""

import json
import os
import re
import hashlib
from pathlib import Path

# ─────────────────────────────────────────────
# 설정 상수
# ─────────────────────────────────────────────
CLAUDE_MODEL  = "claude-sonnet-4-6"
MAX_TOKENS    = 8192   # ★큰 표(수십 행)가 잘려 뒷행·각주가 사라지는 것 방지(4096→8192)
# ★캐시 위치는 '현재 작업 디렉터리'가 아니라 **이 저장소(rainfield-im) 고정**이다.
#   전엔 Path(".claude_cache") 라 실행 위치가 바뀔 때마다 그 폴더에 빈 캐시가 새로 생겨
#   ①진짜 캐시(250개)를 못 찾아 매번 LLM 재호출(느리고 비용 발생, 결과도 달라짐)
#   ②엉뚱한 폴더가 지저분하게 생김 — 두 문제가 반복됐다.
CACHE_DIR = Path(os.environ.get("RAINFIELD_CACHE_DIR")
                 or (Path(__file__).resolve().parent.parent / ".claude_cache"))
CACHE_DIR.mkdir(parents=True, exist_ok=True)

# Claude Sonnet 4.6 가격 (USD per 1M tokens)
_PRICE_INPUT  = 3.0
_PRICE_OUTPUT = 15.0


# ─────────────────────────────────────────────
# 1. API 클라이언트 초기화
# ─────────────────────────────────────────────
def get_client():
    """
    환경변수 ANTHROPIC_API_KEY에서 키를 로드해 Anthropic 클라이언트를 반환합니다.
    .env 파일이 있으면 자동 로드 (python-dotenv).
    키가 없으면 ValueError 발생.
    """
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass

    from anthropic import Anthropic

    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        raise ValueError(
            "ANTHROPIC_API_KEY가 설정되지 않았습니다.\n"
            "방법 1: 프로젝트 루트에 .env 파일 생성 후 ANTHROPIC_API_KEY=sk-ant-... 입력\n"
            "방법 2: 환경변수로 직접 설정 (export ANTHROPIC_API_KEY=sk-ant-...)"
        )
    return Anthropic(api_key=api_key)


# ─────────────────────────────────────────────
# 2. 캐시 키 생성
# ─────────────────────────────────────────────
def make_cache_key(slide_num: int, pdf_text: str, prompt_version: str = "v1",
                   system_prompt: str = "") -> str:
    """
    slide_num + prompt_version + 시스템 프롬프트 본문 + pdf_text 의 SHA256 해시.
    같은 PDF + 같은 슬라이드 + 같은 프롬프트면 항상 동일한 키.

    ★system_prompt 를 키에 포함하는 이유:
      예전엔 prompt_version 문자열만 키에 넣어서, 프롬프트 '내용'을 고쳐도 버전을 안 올리면
      캐시가 옛 결과를 그대로 돌려줬다. 프롬프트를 바꿨는데 결과가 하나도 안 바뀌어
      한참 헤맨 적이 있다(2026-08-07). 이제 본문이 바뀌면 키가 자동으로 갈린다.
    """
    sp = hashlib.sha256((system_prompt or "").encode("utf-8")).hexdigest()[:16]
    raw = f"{slide_num}|{prompt_version}|{sp}|{pdf_text}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


# ─────────────────────────────────────────────
# 3. 핵심 API 호출 함수
# ─────────────────────────────────────────────
def call_claude(
    system_prompt: str,
    user_prompt: str,
    slide_num: int,
    pdf_context: str,
    use_cache: bool = True,
    max_retries: int = 2,
    prompt_version: str = "v1",
) -> dict:
    """
    Claude API 호출 → JSON 응답 파싱 → dict 반환.

    동작 흐름:
      1. 캐시 키 생성. 캐시 히트 시 즉시 반환.
      2. messages.create 호출 (temperature=0).
      3. 응답에서 JSON 블록 파싱.
      4. 파싱 실패 시 JSON 재출력 요청 메시지 추가 후 재시도.
      5. 성공 시 캐시 저장.
      6. 토큰 수·추정 비용 출력.

    Returns
    -------
    {
        "ok": bool,
        "data": dict | None,
        "raw_text": str,
        "usage": {"input_tokens": int, "output_tokens": int, "estimated_cost_usd": float},
        "cached": bool,
        "error": str | None,
    }
    """
    cache_key  = make_cache_key(slide_num, pdf_context, prompt_version,
                                system_prompt=system_prompt)
    cache_file = CACHE_DIR / f"{cache_key}.json"

    # ── 캐시 히트 ────────────────────────────────────────────
    if use_cache and cache_file.exists():
        cached = json.loads(cache_file.read_text(encoding="utf-8"))
        cached["cached"] = True
        print(f"[claude_api] slide={slide_num} CACHE HIT ({cache_key[:12]}…)")
        return cached

    # ── API 호출 ─────────────────────────────────────────────
    client   = get_client()
    messages = [{"role": "user", "content": user_prompt}]

    total_input  = 0
    total_output = 0
    raw_text     = ""
    data         = None
    error_msg    = None

    for attempt in range(max_retries + 1):
        try:
            response = client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=MAX_TOKENS,
                temperature=0,
                system=system_prompt,
                messages=messages,
            )
        except Exception as exc:
            error_msg = str(exc)
            break

        raw_text      = response.content[0].text
        total_input  += response.usage.input_tokens
        total_output += response.usage.output_tokens

        parsed = _parse_json_from_text(raw_text)
        if parsed is not None:
            data = parsed
            break

        # 파싱 실패 → 재시도 메시지 추가
        if attempt < max_retries:
            print(f"[claude_api] slide={slide_num} JSON 파싱 실패, 재시도 {attempt + 1}/{max_retries}")
            messages.append({"role": "assistant", "content": raw_text})
            messages.append({
                "role": "user",
                "content": (
                    "응답을 반드시 순수한 JSON 형식으로만 출력하세요. "
                    "코드 블록(```json ... ```) 또는 중괄호로 시작하는 JSON만 출력하고 "
                    "다른 텍스트는 포함하지 마세요."
                ),
            })
        else:
            error_msg = "JSON 파싱 실패 — 모든 재시도 소진"

    cost = estimate_cost(total_input, total_output)
    usage = {
        "input_tokens":       total_input,
        "output_tokens":      total_output,
        "estimated_cost_usd": cost,
    }
    print(
        f"[claude_api] slide={slide_num} "
        f"in={total_input} out={total_output} "
        f"cost=${cost:.4f} USD"
    )

    result = {
        "ok":       data is not None,
        "data":     data,
        "raw_text": raw_text,
        "usage":    usage,
        "cached":   False,
        "error":    error_msg,
    }

    # ── 캐시 저장 ─────────────────────────────────────────────
    if use_cache and data is not None:
        cache_file.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    return result


# ─────────────────────────────────────────────
# 4. 숫자 환각 검증
# ─────────────────────────────────────────────
_NUM_RE = re.compile(r"\d[\d,\.]*\d|\d")   # 콤마·소수 포함 숫자


def _normalize_number(s: str) -> str:
    """콤마·공백을 제거하고 불필요한 소수점 0 정리. 비교용 정규화."""
    return s.replace(",", "").replace(" ", "").rstrip("0").rstrip(".")


def verify_numbers_in_pdf(claude_output: dict, pdf_text: str) -> dict:
    """
    Claude 출력 dict 안의 모든 숫자가 PDF 원문에 존재하는지 검증합니다.

    검증 규칙:
      - 콤마 단위 구분자 차이 무시 ("1640" == "1,640")
      - 단독 1~2자리 숫자(1, 10 등)는 너무 흔하므로 검증 제외
      - 날짜 형식(2026.08 등), 금액, 비율, 면적 모두 포함

    Returns
    -------
    {
        "ok": bool,
        "hallucinated_numbers": list[str],
        "verified_count": int,
        "total_count": int,
    }
    """
    # PDF 원문 숫자 집합 구축
    pdf_nums = {
        _normalize_number(m)
        for m in _NUM_RE.findall(pdf_text)
        if len(m.replace(",", "").replace(".", "")) >= 3  # 3자리 이상만
    }

    # Claude 출력 전체 텍스트화
    output_text = json.dumps(claude_output, ensure_ascii=False)
    claude_nums = [m for m in _NUM_RE.findall(output_text)
                   if len(m.replace(",", "").replace(".", "")) >= 3]

    hallucinated = []
    for num in claude_nums:
        norm = _normalize_number(num)
        if norm not in pdf_nums:
            hallucinated.append(num)

    total     = len(claude_nums)
    verified  = total - len(hallucinated)
    return {
        "ok":                   len(hallucinated) == 0,
        "hallucinated_numbers": hallucinated,
        "verified_count":       verified,
        "total_count":          total,
    }


# ─────────────────────────────────────────────
# 5. 비용 추정
# ─────────────────────────────────────────────
def estimate_cost(input_tokens: int, output_tokens: int) -> float:
    """
    Claude Sonnet 4.6 기준 비용 추정 (USD).
      input : $3 / 1M tokens
      output: $15 / 1M tokens
    """
    return (input_tokens * _PRICE_INPUT + output_tokens * _PRICE_OUTPUT) / 1_000_000


# ─────────────────────────────────────────────
# 6. 캐시 관리
# ─────────────────────────────────────────────
def clear_cache() -> int:
    """캐시 디렉토리의 모든 .json 파일을 삭제하고 삭제 개수를 반환합니다."""
    count = 0
    for f in CACHE_DIR.glob("*.json"):
        f.unlink()
        count += 1
    print(f"[claude_api] 캐시 {count}개 파일 삭제 완료")
    return count


def get_cache_stats() -> dict:
    """캐시 파일 개수와 총 용량(MB)을 반환합니다."""
    files     = list(CACHE_DIR.glob("*.json"))
    total_mb  = sum(f.stat().st_size for f in files) / (1024 * 1024)
    return {"file_count": len(files), "total_mb": round(total_mb, 3)}


# ─────────────────────────────────────────────
# 내부 헬퍼
# ─────────────────────────────────────────────
def _parse_json_from_text(text: str) -> "dict | None":
    """
    텍스트에서 JSON 블록을 추출해 파싱합니다.
    ```json ... ``` 블록 → 중괄호 블록 → 전체 텍스트 순으로 시도.
    """
    # 1) ```json ... ``` 코드 블록
    m = re.search(r"```json\s*([\s\S]+?)\s*```", text)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass

    # 2) 중괄호 블록 (첫 { ~ 마지막 })
    start = text.find("{")
    end   = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        try:
            return json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            pass

    # 3) 전체 텍스트 시도
    try:
        return json.loads(text.strip())
    except json.JSONDecodeError:
        return None
