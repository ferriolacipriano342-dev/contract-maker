"""증빙서류(신분증·건축물대장·등기부등본) AI 판독 및 계약서 자동 채우기 모듈.

- Pillow 기반 EXIF 자동 회전 보정 및 고해상도 이미지 리사이즈
- Gemini API (`google-genai`) `response_mime_type="application/json"` + Pydantic `response_schema` 강제
- 실패 원인(API 키 미설정, API 오류, 응답 차단/거절, JSON 파싱 실패, 빈 추출) 및 AI 원본 응답 투명 반환
"""

from __future__ import annotations

import io
import json
import os
import re
import time
import tomllib
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import TimeoutError as FuturesTimeoutError
from pathlib import Path
from typing import Any, Callable

from PIL import Image, ImageOps
from pydantic import BaseModel, Field

BASE_DIR = Path(__file__).resolve().parent.parent
SECRETS_TOML_PATH = BASE_DIR / ".streamlit" / "secrets.toml"

MAX_IMAGE_DIMENSION = 1024
JPEG_QUALITY = 80
DEFAULT_API_TIMEOUT_SECONDS = 60
RETRY_DELAY_SECONDS = 1.5
MAX_OCR_OUTPUT_TOKENS = 1024
MAX_FALLBACK_MODELS = 8
GEMINI_OCR_MODEL = "gemini-2.5-flash"
PREFERRED_OCR_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-2.5-flash-lite",
    "gemini-3.8-flash",
    "gemini-flash-latest",
    "gemini-2.5-pro",
]
DEFAULT_FALLBACK_MODELS = PREFERRED_OCR_MODELS[1:]
_EXCLUDED_MODEL_KEYWORDS = (
    "embed",
    "tts",
    "imagen",
    "veo",
    "aqa",
    "audio",
    "live",
    "robotics",
    "computer-use",
    "deep-research",
    "image-generation",
    "native-audio",
    "gemma",
)
SUPPORTED_IMAGE_MIMES = {"image/jpeg", "image/jpg", "image/png", "image/webp"}
SUPPORTED_MIMES = SUPPORTED_IMAGE_MIMES | {"application/pdf"}


class ExtractedDocumentFields(BaseModel):
    """Gemini가 반드시 반환해야 하는 고정 JSON 스키마."""

    id_name: str = Field(default="", description="신분증에 기재된 성명")
    id_number: str = Field(default="", description="신분증에 기재된 주민등록번호 또는 생년월일/식별번호")
    id_address: str = Field(default="", description="신분증에 기재된 주소")
    property_address: str = Field(default="", description="토지대장·건축물대장·등기부등본 표제부에 기재된 소재지(지번)")
    land_category: str = Field(default="", description="토지 지목 (예: 전, 답, 대, 임야)")
    land_share_ratio: str = Field(default="", description="대지권비율 (예: 33101/55.3748)")
    land_area: str = Field(default="", description="토지 또는 대지권 면적 숫자만 (단위 제외)")
    land_zoning: str = Field(default="", description="토지 용도지역·지구 (예: 계획관리지역, 농림지역, 자연녹지지역)")
    building_structure: str = Field(default="", description="건물 주구조 (예: 철근콘크리트)")
    building_use: str = Field(default="", description="건물 용도 (예: 공동주택, 노유자시설)")
    building_area: str = Field(default="", description="건물/전유부분 면적 숫자만 (단위 제외)")
    lease_part: str = Field(default="", description="동·호수 또는 대상 부분 (예: 108동 401호 전부)")
    lease_area: str = Field(default="", description="전용/계약 면적 숫자만 (단위 제외)")
    facility_name: str = Field(default="", description="요양원/시설 명칭 (서류에 있는 경우)")
    store_name: str = Field(default="", description="상호/점포명 (서류에 있는 경우)")
    max_bond_amount: str = Field(default="", description="등기부등본 을구 근저당권 채권최고액 합계 숫자만")
    owner_name: str = Field(default="", description="등기부등본 갑구 소유자 성명")
    owner_rrn: str = Field(default="", description="등기부등본 갑구 소유자 주민등록번호/생년월일")
    owner_address: str = Field(default="", description="등기부등본 갑구 소유자 주소")


def _read_key_from_secrets_file() -> str:
    """프로젝트의 .streamlit/secrets.toml에서 GEMINI_API_KEY를 실시간으로 읽는다 (재시작 불필요)."""
    if not SECRETS_TOML_PATH.exists():
        return ""
    try:
        with SECRETS_TOML_PATH.open("rb") as f:
            parsed = tomllib.load(f)
        val = parsed.get("GEMINI_API_KEY", "")
        return str(val).strip() if val else ""
    except Exception:
        return ""


def get_gemini_api_key(secrets_obj: Any = None) -> str:
    """실시간 .streamlit/secrets.toml -> st.secrets -> 환경변수 순으로 GEMINI_API_KEY를 조회한다."""
    file_key = _read_key_from_secrets_file()
    if file_key:
        return file_key
    if secrets_obj is not None:
        try:
            key = secrets_obj.get("GEMINI_API_KEY", "")
            if key and str(key).strip():
                return str(key).strip()
        except Exception:
            pass
    return os.environ.get("GEMINI_API_KEY", "").strip()


def mask_api_key(api_key: str) -> str:
    """화면 표시용 마스킹 문자열 (예: AIza••••1234)."""
    k = api_key.strip()
    if len(k) <= 8:
        return "•" * len(k)
    return f"{k[:4]}••••••{k[-4:]}"


def save_gemini_api_key_to_secrets(api_key: str) -> tuple[bool, str]:
    """입력한 Gemini API 키를 .streamlit/secrets.toml에 안전하게 저장(기존 APP_PIN 등 보존)한다."""
    cleaned = api_key.strip().strip('"').strip("'")
    if not cleaned:
        return False, "API 키를 입력해 주세요."

    SECRETS_TOML_PATH.parent.mkdir(parents=True, exist_ok=True)
    escaped = cleaned.replace("\\", "\\\\").replace('"', '\\"')
    new_line = f'GEMINI_API_KEY = "{escaped}"'

    lines: list[str] = []
    if SECRETS_TOML_PATH.exists():
        lines = SECRETS_TOML_PATH.read_text(encoding="utf-8").splitlines()

    replaced = False
    out_lines: list[str] = []
    for line in lines:
        if re.match(r"^\s*#?\s*GEMINI_API_KEY\s*=", line):
            if not replaced:
                out_lines.append(new_line)
                replaced = True
        else:
            out_lines.append(line)

    if not replaced:
        if not out_lines:
            out_lines.append('APP_PIN = "0000"')
        out_lines.append(new_line)

    content = "\n".join(out_lines) + "\n"
    os.environ["GEMINI_API_KEY"] = cleaned
    try:
        SECRETS_TOML_PATH.write_text(content, encoding="utf-8")
        return True, f"Gemini API 키가 `.streamlit/secrets.toml`에 저장되어 즉시 활성화되었습니다. ({mask_api_key(cleaned)})"
    except OSError:
        return (
            True,
            f"현재 세션에 Gemini API 키가 즉시 활성화되었습니다. ({mask_api_key(cleaned)}) "
            "클라우드 영구 저장은 Streamlit Cloud 대시보드 [Settings → Secrets]에 등록해 주세요.",
        )


def delete_gemini_api_key_from_secrets() -> None:
    """저장된 GEMINI_API_KEY를 .streamlit/secrets.toml과 환경변수에서 제거한다."""
    os.environ.pop("GEMINI_API_KEY", None)
    if not SECRETS_TOML_PATH.exists():
        return
    lines = SECRETS_TOML_PATH.read_text(encoding="utf-8").splitlines()
    kept = [line for line in lines if not re.match(r"^\s*GEMINI_API_KEY\s*=", line)]
    SECRETS_TOML_PATH.write_text(("\n".join(kept) + "\n") if kept else "", encoding="utf-8")


def preprocess_document_bytes(
    data: bytes,
    mime_type: str,
    filename: str = "",
    max_dim: int = MAX_IMAGE_DIMENSION,
) -> tuple[bytes, str, str]:
    """이미지의 EXIF 회전을 자동 보정하고 최대 해상도로 리사이즈해 JPEG로 정규화한다.

    PDF 파일은 원본 그대로 반환한다.
    """
    norm_mime = (mime_type or "").lower().strip()
    if norm_mime == "image/jpg":
        norm_mime = "image/jpeg"
    if not norm_mime and filename.lower().endswith(".pdf"):
        norm_mime = "application/pdf"
    if not norm_mime:
        norm_mime = "image/jpeg"

    if norm_mime == "application/pdf":
        return data, "application/pdf", f"{filename}: PDF 원본 전송 ({len(data) // 1024}KB)"

    if norm_mime not in SUPPORTED_IMAGE_MIMES:
        return data, norm_mime, f"{filename}: 미지원 MIME({norm_mime})"

    try:
        with Image.open(io.BytesIO(data)) as raw_img:
            orig_size = raw_img.size
            exif = raw_img.getexif()
            orientation = exif.get(274, 1) if exif else 1
            img = ImageOps.exif_transpose(raw_img)

            # 알파 채널이 있는 PNG 등은 흰색 배경 RGB로 변환
            if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
                bg = Image.new("RGB", img.size, (255, 255, 255))
                rgba = img.convert("RGBA")
                bg.paste(rgba, mask=rgba.split()[-1])
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")

            resized = False
            if max(img.size) > max_dim:
                img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
                resized = True

            out_buf = io.BytesIO()
            img.save(out_buf, format="JPEG", quality=JPEG_QUALITY, optimize=True)
            processed = out_buf.getvalue()

            note_parts = [f"{orig_size[0]}x{orig_size[1]}"]
            if orientation != 1:
                note_parts.append(f"EXIF 회전보정({orientation})")
            if resized:
                note_parts.append(f"리사이즈→{img.size[0]}x{img.size[1]}")
            note_parts.append(f"{max(1, len(processed) // 1024)}KB")
            return processed, "image/jpeg", f"{filename}: " + " / ".join(note_parts)
    except Exception as exc:
        return data, norm_mime, f"{filename}: Pillow 보정 건너뜀({type(exc).__name__}), 원본 전송"


_OCR_PROMPT = """당신은 대한민국 공인중개사 사무소의 계약서 작성 보조 OCR 엔진입니다.
사용자가 본인의 부동산·토지·요양원 계약서 작성을 위해 직접 제출한 신분증(주민등록증·운전면허증·사업자등록증) 및 공적장부(등기부등본·토지대장·건축물대장·토지이용계획확인원) 이미지/PDF에서 글자를 정확히 판독하여 지정된 JSON 스키마에 채워 넣으세요.

규칙:
1. 신분증(주민등록증, 운전면허증 등)에 기재된 성명, 주민등록번호(또는 생년월일/식별번호), 주소는 반드시 `id_name`, `id_number`, `id_address` 필드에 넣으세요.
   - 주민등록증의 주소는 맨 아래 최신 전입 주소(가장 마지막 줄에 기재된 주소)를 우선 추출하세요.
   - 성명 옆 괄호 안 한자는 제외하고 한글 성명만 추출하세요 (예: "홍길동(洪吉童)" -> "홍길동").
2. 등기부등본(토지·건물·집합건물)·토지대장·건축물대장의 **상단 `[토지]`/`[건물]`/`[집합건물]` 표시 및 표제부의 `소재지번`**(예: "경기도 양평군 양서면 목왕리 123-4")은 반드시 **`property_address`**에 빠짐없이 전체 주소로 넣으세요.
   - 표제부/토지대장의 지목(`전`, `답`, `대`, `임야` 등)은 `land_category`, 토지면적(㎡ 숫자만)은 `land_area`, 용도지역·지구는 `land_zoning`, 대지권비율은 `land_share_ratio`, 건물 구조·용도·면적은 `building_structure`·`building_use`·`building_area`, 동·호수는 `lease_part`, 전용면적은 `lease_area`에 넣으세요.
   - 등기부등본 을구의 말소되지 않은 유효 근저당권 채권최고액 합계는 `max_bond_amount`에 넣으세요.
3. 등기부등본 갑구의 현재 최종 소유자 정보(성명, 주민등록번호/생년월일, 주소)는 `owner_name`, `owner_rrn`, `owner_address`에 넣으세요.
4. 서류에 없는 항목은 추측하지 말고 빈 문자열("")로 두세요.
5. 면적은 단위(㎡) 없이 숫자만, 채권최고액은 쉼표/원 없이 순수 숫자 문자열만 반환하세요.
"""


def _parse_json_strict(raw_text: str) -> dict[str, Any]:
    text = raw_text.strip()
    fence_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.DOTALL)
    if fence_match:
        text = fence_match.group(1)
    else:
        brace_match = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if brace_match:
            text = brace_match.group(0)
    parsed = json.loads(text)
    if not isinstance(parsed, dict):
        raise ValueError("AI 응답이 JSON 객체(dict) 형태가 아닙니다.")
    return parsed


def map_extracted_to_form_fields(
    raw_json: dict[str, Any],
    id_target_role: str = "party1",
    doc_categories: list[str] | None = None,
) -> dict[str, str]:
    """AI JSON 결과를 계약 유형 무관 100% 결정론적 canonical form 키로 매핑한다.

    - id_target_role == 'party1': 매도인·임대인·양도인 (lessor_name, lessor_rrn, lessor_address)
    - id_target_role == 'party2': 매수인·임차인·양수인 (lessee_name, lessee_rrn, lessee_address)
    """
    cleaned = {
        k: str(v).strip()
        for k, v in raw_json.items()
        if v is not None and str(v).strip() != ""
    }
    mapped: dict[str, str] = {}
    categories = set(doc_categories or [])
    only_id_docs = bool(categories) and categories.issubset({"신분증"})

    # 서류 내용 자체에 등기부·토지대장·건축물대장 항목이 포함되어 있는지 판별
    # (사용자가 분류 드롭다운을 '신분증'으로 둔 채 등기부등본/토지대장을 올려도 소재지가 누락되지 않게 방지)
    has_registry_signals = any(
        cleaned.get(k)
        for k in (
            "land_category",
            "land_area",
            "land_zoning",
            "land_share_ratio",
            "building_structure",
            "building_area",
            "owner_name",
            "max_bond_amount",
        )
    )

    # 1) 신분증 정보 -> 선택된 당사자(party1: 매도인·임대인·양도인 / party2: 매수인·임차인·양수인)
    prefix = "lessor" if id_target_role == "party1" else "lessee"

    id_name = cleaned.get("id_name") or cleaned.get(f"{prefix}_name") or cleaned.get("name", "")
    id_number = cleaned.get("id_number") or cleaned.get(f"{prefix}_rrn") or cleaned.get("rrn", "")
    id_address = cleaned.get("id_address") or cleaned.get(f"{prefix}_address", "")

    # 순수 신분증만 올렸고 등기부 항목이 전혀 없는데 AI가 주소를 property_address에 넣은 경우에만 당사자 주소로 보정
    pure_id_only = only_id_docs and not has_registry_signals
    if pure_id_only and not id_address and cleaned.get("property_address"):
        id_address = cleaned["property_address"]

    if id_name:
        mapped[f"{prefix}_name"] = id_name
    if id_number:
        mapped[f"{prefix}_rrn"] = id_number
    if id_address:
        mapped[f"{prefix}_address"] = id_address

    # 2) 등기부등본 갑구 소유자 정보 -> party1(매도인·임대인·양도인)이 비어 있으면 자동 보완
    if cleaned.get("owner_name") and "lessor_name" not in mapped:
        mapped["lessor_name"] = cleaned["owner_name"]
    if cleaned.get("owner_rrn") and "lessor_rrn" not in mapped:
        mapped["lessor_rrn"] = cleaned["owner_rrn"]
    if cleaned.get("owner_address") and "lessor_address" not in mapped:
        mapped["lessor_address"] = cleaned["owner_address"]

    # 3) 부동산/토지/시설 공적장부 정보 (등기부 표제부 소재지번 -> address 빠짐없이 매핑)
    prop_addr = cleaned.get("property_address") or cleaned.get("address", "")
    if prop_addr and not (pure_id_only and prop_addr == id_address):
        mapped["address"] = prop_addr

    for field in (
        "land_category",
        "land_share_ratio",
        "land_area",
        "land_zoning",
        "building_structure",
        "building_use",
        "building_area",
        "lease_part",
        "lease_area",
        "facility_name",
        "store_name",
        "max_bond_amount",
    ):
        if cleaned.get(field):
            mapped[field] = cleaned[field]

    return mapped


def extract_contract_fields_from_documents(
    documents: list[dict[str, Any]],
    api_key: str,
    id_target_role: str = "party1",
    timeout_seconds: int = DEFAULT_API_TIMEOUT_SECONDS,
    on_progress: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    """증빙서류(이미지/PDF)를 Pillow 전처리 후 Gemini API(60초 타임아웃·다중 모델 자동 우회)로 판독한다."""
    step_log: list[str] = []

    def _emit(msg: str) -> None:
        step_log.append(msg)
        if on_progress is not None:
            try:
                on_progress(msg)
            except Exception:
                pass

    result: dict[str, Any] = {
        "ok": False,
        "mapped_fields": {},
        "raw_json": {},
        "raw_response_text": "",
        "error_type": "",
        "message": "",
        "preprocess_notes": [],
        "step_log": step_log,
        "id_target_role": id_target_role,
    }

    if not api_key:
        result["error_type"] = "NO_API_KEY"
        result["message"] = (
            "GEMINI_API_KEY가 설정되지 않았습니다. 좌측 사이드바 하단의 "
            "`🔑 Gemini API 키 설정`에서 키를 저장해 주세요."
        )
        return result

    if not documents:
        result["error_type"] = "NO_DOCUMENTS"
        result["message"] = "분석할 증빙서류가 없습니다. 먼저 신분증·건축물대장·등기부등본을 촬영하거나 업로드해 주세요."
        return result

    try:
        from google import genai
        from google.genai import types
    except ImportError as exc:
        result["error_type"] = "SDK_MISSING"
        result["message"] = f"`google-genai` 패키지가 설치되어 있지 않습니다. (`pip install google-genai`) — {exc}"
        return result

    # 1단계: Pillow 이미지 EXIF 회전 보정 및 경량화(1024px / JPEG Q80)
    _emit("1) 이미지 보정 중 (EXIF 회전 보정 · 1024px/Q80 경량화)...")
    parts: list[Any] = [_OCR_PROMPT]
    categories: list[str] = []
    attached_count = 0
    total_payload_bytes = 0

    for doc in documents:
        raw_bytes = doc.get("data")
        if not raw_bytes:
            continue
        fname = str(doc.get("filename", "document"))
        cat = str(doc.get("category", "기타"))
        proc_bytes, proc_mime, note = preprocess_document_bytes(
            raw_bytes,
            mime_type=str(doc.get("mime_type", "")),
            filename=fname,
            max_dim=MAX_IMAGE_DIMENSION,
        )
        result["preprocess_notes"].append(note)
        if proc_mime not in SUPPORTED_MIMES:
            continue

        categories.append(cat)
        total_payload_bytes += len(proc_bytes)
        parts.append(f"[증빙서류 분류: {cat} | 파일명: {fname}]")
        parts.append(types.Part.from_bytes(data=proc_bytes, mime_type=proc_mime))
        attached_count += 1

    if attached_count == 0:
        result["error_type"] = "NO_DOCUMENTS"
        result["message"] = "분석 가능한 이미지(JPG/PNG/WEBP) 또는 PDF 파일이 없습니다."
        return result

    _emit(f"✅ 1) 이미지 보정 완료 ({attached_count}건, 전송 용량 {max(1, total_payload_bytes // 1024)}KB)")

    # 2단계: 고속 안정 모델 우선 + 다중 모델 자동 우회(Fallback) 전송
    # (gemini-2.5-flash -> gemini-2.0-flash -> gemini-2.5-flash-lite -> gemini-3.8-flash 등)
    _emit(f"2) Gemini API({GEMINI_OCR_MODEL}) 전송 중... (최대 {timeout_seconds}초 대기)")
    started_at = time.monotonic()

    def _make_client(limit_sec: float) -> Any:
        http_opts = types.HttpOptions(
            timeout=int(limit_sec * 1000),
            retry_options=types.HttpRetryOptions(attempts=1),
        )
        return genai.Client(api_key=api_key, http_options=http_opts), http_opts

    def _call_model(model_name: str, limit_sec: float) -> Any:
        client, http_opts = _make_client(limit_sec)
        return client.models.generate_content(
            model=model_name,
            contents=parts,
            config=types.GenerateContentConfig(
                temperature=0.0,
                max_output_tokens=MAX_OCR_OUTPUT_TOKENS,
                response_mime_type="application/json",
                response_schema=ExtractedDocumentFields,
                http_options=http_opts,
            ),
        )

    def _run_with_hard_timeout(fn: Callable[[], Any], limit_sec: float) -> Any:
        executor = ThreadPoolExecutor(max_workers=1)
        future = executor.submit(fn)
        try:
            return future.result(timeout=max(1.0, limit_sec))
        finally:
            executor.shutdown(wait=False, cancel_futures=True)

    def _is_fallback_trigger(exc: Exception) -> tuple[bool, str, bool]:
        """재시도·우회 가능 오류 여부, 한글 사유 태그, 대기(sleep) 필요 여부를 반환한다."""
        msg = f"{type(exc).__name__}: {exc}".upper()
        if "504" in msg or "DEADLINE_EXCEEDED" in msg or "DEADLINE EXCEEDED" in msg or "TIMEOUT" in msg or "TIMED OUT" in msg:
            return True, "504 시간초과", True
        if "503" in msg or "UNAVAILABLE" in msg or "OVERLOADED" in msg:
            return True, "503 혼잡", True
        if "429" in msg or "RESOURCE_EXHAUSTED" in msg or "RATE LIMIT" in msg or "QUOTA" in msg:
            return True, "429 한도초과", True
        if "502" in msg or "BAD_GATEWAY" in msg or "500" in msg or "INTERNAL" in msg:
            return True, "5xx 서버지연", True
        if "404" in msg or "NOT_FOUND" in msg or "NOT FOUND" in msg:
            return True, "404 미지원", False
        return False, "", False

    def _discover_live_fallback_models(exclude: set[str]) -> list[str]:
        """client.models.list()로 현재 API 키에서 실제 호출 가능한 추가 gemini OCR 모델을 동적 조회한다."""
        discovered: list[str] = []
        try:
            def _list_models() -> list[Any]:
                client, _ = _make_client(8.0)
                return list(client.models.list())

            raw_models = _run_with_hard_timeout(_list_models, 8.0)
            for m in raw_models:
                raw_name = str(getattr(m, "name", "") or "")
                short_name = raw_name.removeprefix("models/").strip()
                lower = short_name.lower()
                if not lower.startswith("gemini-"):
                    continue
                if any(kw in lower for kw in _EXCLUDED_MODEL_KEYWORDS):
                    continue
                actions = getattr(m, "supported_actions", None) or []
                if actions and not any("generatecontent" in str(a).lower() for a in actions):
                    continue
                if short_name not in exclude and short_name not in discovered:
                    discovered.append(short_name)
        except Exception as list_exc:
            _emit(f"ℹ️ 추가 가용 모델 목록 조회 건너뜀({type(list_exc).__name__})")

        def _rank(name: str) -> tuple[int, str]:
            lo = name.lower()
            if "2.5-flash" in lo and "lite" not in lo:
                tier = 0
            elif "2.0-flash" in lo:
                tier = 1
            elif "flash-lite" in lo:
                tier = 2
            elif "flash" in lo:
                tier = 3
            elif "pro" in lo:
                tier = 4
            else:
                tier = 5
            return (tier, name)

        discovered.sort(key=_rank)
        return discovered[:MAX_FALLBACK_MODELS]

    candidate_queue: list[str] = list(PREFERRED_OCR_MODELS)
    tried_models: set[str] = set()
    attempt_errors: list[str] = []
    discovered_once = False
    first_fail_reason = ""
    used_model = GEMINI_OCR_MODEL
    response = None

    idx = 0
    while idx < len(candidate_queue):
        model_name = candidate_queue[idx]
        idx += 1
        if model_name in tried_models:
            continue
        tried_models.add(model_name)

        if len(tried_models) > 1:
            _emit(f"🔄 대체 모델 [{model_name}]로 자동 전환하여 전송 중... (최대 {timeout_seconds}초)")

        model_start = time.monotonic()
        try:
            response = _run_with_hard_timeout(
                lambda m=model_name: _call_model(m, float(timeout_seconds)),
                float(timeout_seconds),
            )
            used_model = model_name
            break
        except FuturesTimeoutError:
            m_elapsed = time.monotonic() - model_start
            attempt_errors.append(f"{model_name}: TIMEOUT({m_elapsed:.1f}초)")
            if not first_fail_reason:
                first_fail_reason = "시간초과"
            if idx >= len(candidate_queue) and not discovered_once:
                discovered_once = True
                extras = _discover_live_fallback_models(tried_models)
                if extras:
                    _emit("📋 추가 가용 모델 확보: " + ", ".join(extras))
                    candidate_queue.extend(extras)
            if idx < len(candidate_queue):
                _emit(f"⏱️ {model_name} 응답 시간 초과({m_elapsed:.1f}초) → {RETRY_DELAY_SECONDS:.1f}초 후 다음 모델로 자동 재시도합니다.")
                time.sleep(RETRY_DELAY_SECONDS)
            continue
        except Exception as exc:
            should_fallback, reason_tag, should_wait = _is_fallback_trigger(exc)
            attempt_errors.append(f"{model_name}: {reason_tag or type(exc).__name__} ({exc})")
            if should_fallback:
                if not first_fail_reason:
                    first_fail_reason = reason_tag
                if idx >= len(candidate_queue) and not discovered_once:
                    discovered_once = True
                    _emit("🔍 현재 API 키에서 추가 호출 가능한 Gemini 모델 목록 조회 중...")
                    extras = _discover_live_fallback_models(tried_models)
                    if extras:
                        _emit("📋 추가 가용 모델 확보: " + ", ".join(extras))
                        candidate_queue.extend(extras)
                if idx < len(candidate_queue):
                    if should_wait:
                        _emit(f"⚠️ {model_name} {reason_tag} → {RETRY_DELAY_SECONDS:.1f}초 후 다음 모델로 자동 재시도합니다.")
                        time.sleep(RETRY_DELAY_SECONDS)
                    else:
                        _emit(f"⚠️ {model_name} {reason_tag} → 다음 모델로 즉시 전환합니다.")
                continue
            # 인증 오류(401/403) 등 우회 불가능한 치명 오류는 즉시 중단
            result["error_type"] = "API_ERROR"
            result["message"] = f"Gemini API 호출 중 오류가 발생했습니다({model_name}): {type(exc).__name__}: {exc}"
            result["raw_response_text"] = "\n".join(attempt_errors)
            return result

    if response is None:
        result["error_type"] = "ALL_MODELS_FAILED"
        result["message"] = (
            f"시도한 모든 Gemini 모델({len(tried_models)}개: {', '.join(tried_models)})이 "
            "시간초과(504), 혼잡(503/429) 또는 미지원(404)으로 실패했습니다. 잠시 후 다시 시도해 주세요."
        )
        result["raw_response_text"] = "\n".join(attempt_errors)
        return result

    elapsed = time.monotonic() - started_at
    result["used_model"] = used_model
    if used_model != GEMINI_OCR_MODEL:
        _emit(
            f"✅ {GEMINI_OCR_MODEL} {first_fail_reason or '혼잡'} → {used_model}로 자동 전환 성공 ({elapsed:.1f}초 소요)"
        )
    else:
        _emit(f"✅ 2) Gemini API({used_model}) 응답 수신 완료 ({elapsed:.1f}초 소요)")

    # 3단계: 응답 검증 및 JSON 파싱
    _emit("3) 응답 파싱 중...")
    raw_text = ""
    try:
        raw_text = response.text or ""
    except Exception as exc:
        raw_text = f"[response.text 접근 예외: {exc}]"

    candidates = getattr(response, "candidates", None) or []
    prompt_feedback = getattr(response, "prompt_feedback", None)
    finish_reasons = [str(getattr(c, "finish_reason", "")) for c in candidates]

    if not raw_text.strip() or any(
        fr and ("SAFETY" in fr.upper() or "RECITATION" in fr.upper() or "PROHIBITED" in fr.upper() or "BLOCK" in fr.upper())
        for fr in finish_reasons
    ):
        result["error_type"] = "BLOCKED_RESPONSE"
        result["message"] = (
            "Gemini 모델이 응답을 반환하지 않았거나 정책/안전 필터에 의해 차단되었습니다. "
            f"(finish_reason={finish_reasons}, prompt_feedback={prompt_feedback})"
        )
        result["raw_response_text"] = (
            f"raw_text={raw_text!r}\nfinish_reasons={finish_reasons}\nprompt_feedback={prompt_feedback}"
        )
        return result

    result["raw_response_text"] = raw_text

    try:
        parsed_json = _parse_json_strict(raw_text)
        result["raw_json"] = parsed_json
    except Exception as exc:
        result["error_type"] = "JSON_PARSE_ERROR"
        result["message"] = f"AI 응답을 JSON으로 파싱하지 못했습니다: {type(exc).__name__}: {exc}"
        return result

    mapped = map_extracted_to_form_fields(
        parsed_json,
        id_target_role=id_target_role,
        doc_categories=categories,
    )
    result["mapped_fields"] = mapped

    if not mapped:
        result["error_type"] = "EMPTY_EXTRACTION"
        result["message"] = (
            "AI 응답은 정상 수신되었으나 서류에서 계약서로 옮길 텍스트를 추출하지 못했습니다. "
            "사진의 초점·밝기·해상도를 확인하고 아래 AI 원본 응답을 점검해 주세요."
        )
        return result

    _emit(f"✅ 3) 응답 파싱 완료 ({len(mapped)}개 항목 추출)")
    result["ok"] = True
    switch_note = f" [{GEMINI_OCR_MODEL} → {used_model} 자동 전환]" if used_model != GEMINI_OCR_MODEL else f" [{used_model}]"
    result["message"] = (
        f"증빙서류 {attached_count}건에서 총 {len(mapped)}개 항목을 판독해 입력란에 자동 반영했습니다.{switch_note} ({elapsed:.1f}초)"
    )
    return result
