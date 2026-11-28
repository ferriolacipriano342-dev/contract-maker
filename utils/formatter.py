"""금액 한글 변환, 날짜/기간 계산, 6대 계약 유형별 금액 교차검증 및 특약 별지 분리 유틸리티."""

from __future__ import annotations

import math
import re
from datetime import date, datetime
from typing import Any

# A4 1장 본문 특약란 최대 허용 줄 수 / 항목 수 (초과 시 별지 자동 분리)
MAX_INLINE_CLAUSE_LINES = 10
MAX_INLINE_CLAUSE_COUNT = 5
CHARS_PER_LINE = 46

# 모던·디지털 A4 서식 카테고리별 시그니처 컬러 테마
DOC_SIGNATURE_PALETTES: dict[str, dict[str, str]] = {
    "요양시설": {
        "primary": "#1E3A5F",
        "accent": "#B48A3E",
        "tint": "#EDF2F8",
        "money_bg": "#FBF8F1",
        "doc_category": "요양시설 양도·양수",
    },
    "주거": {
        "primary": "#1D4ED8",
        "accent": "#3B82F6",
        "tint": "#EFF6FF",
        "money_bg": "#EFF6FF",
        "doc_category": "주거용 부동산",
    },
    "상가": {
        "primary": "#047857",
        "accent": "#10B981",
        "tint": "#ECFDF5",
        "money_bg": "#ECFDF5",
        "doc_category": "상업용 부동산",
    },
    "토지": {
        "primary": "#15803D",
        "accent": "#22C55E",
        "tint": "#F0FDF4",
        "money_bg": "#F0FDF4",
        "doc_category": "토지·농지 매매",
    },
}

_DIGITS = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"]
_SUB_UNITS = ["", "십", "백", "천"]
_GROUP_UNITS = ["", "만", "억", "조", "경"]

# 계약 유형별 메타데이터 (요양시설 1순위 + 동종 부동산 순서 정렬, 당사자 호칭, 주요 대금 라벨, 서식 파일명)
CONTRACT_TYPE_META: dict[str, dict[str, str]] = {
    "nursing_home_transfer": {
        "label": "🏥 [요양시설] 요양원 양도·양수",
        "short_label": "요양원 양도·양수",
        "category": "요양시설",
        "form_title": "노 인 요 양 시 설  양 도 · 양 수  계 약 서",
        "party1": "양도인",
        "party2": "양수인",
        "amount_label": "총 양도대금",
        "template_name": "nursing_home_transfer.html",
    },
    "apartment_sale": {
        "label": "🏢 [주거] 아파트 매매",
        "short_label": "아파트 매매",
        "category": "주거",
        "form_title": "아 파 트  매 매  계 약 서",
        "party1": "매도인",
        "party2": "매수인",
        "amount_label": "매매대금",
        "template_name": "real_estate_sale.html",
    },
    "real_estate_sale": {
        "label": "🏠 [주거] 주택 매매",
        "short_label": "주택 매매",
        "category": "주거",
        "form_title": "주 택  매 매  계 약 서",
        "party1": "매도인",
        "party2": "매수인",
        "amount_label": "매매대금",
        "template_name": "real_estate_sale.html",
    },
    "housing_jeonse": {
        "label": "🏠 [주거] 주택 전세",
        "short_label": "주택 전세",
        "category": "주거",
        "form_title": "주 택  전 세  계 약 서",
        "party1": "임대인",
        "party2": "임차인",
        "amount_label": "보증금",
        "template_name": "housing_jeonse.html",
    },
    "housing_wolse": {
        "label": "🏠 [주거] 주택 월세",
        "short_label": "주택 월세",
        "category": "주거",
        "form_title": "주 택  임 대 차 (월 세)  계 약 서",
        "party1": "임대인",
        "party2": "임차인",
        "amount_label": "보증금",
        "template_name": "housing_wolse.html",
    },
    "commercial_sale": {
        "label": "🏪 [상가] 상가 매매",
        "short_label": "상가 매매",
        "category": "상가",
        "form_title": "상 가 건 물  매 매  계 약 서",
        "party1": "매도인",
        "party2": "매수인",
        "amount_label": "매매대금",
        "template_name": "real_estate_sale.html",
    },
    "commercial_lease": {
        "label": "🏪 [상가] 상가 임대차",
        "short_label": "상가 임대차",
        "category": "상가",
        "form_title": "상 가 건 물  임 대 차  계 약 서",
        "party1": "임대인",
        "party2": "임차인",
        "amount_label": "보증금",
        "template_name": "commercial_lease.html",
    },
    "premium_rights": {
        "label": "🏪 [상가] 권리금 계약",
        "short_label": "권리금 계약",
        "category": "상가",
        "form_title": "상 가  권 리 금  계 약 서",
        "party1": "양도인",
        "party2": "양수인",
        "amount_label": "권리금 총액",
        "template_name": "premium_rights.html",
    },
    "farmland_sale": {
        "label": "🌾 [토지] 농지 매매",
        "short_label": "농지 매매",
        "category": "토지",
        "form_title": "농 지  매 매  계 약 서",
        "party1": "매도인",
        "party2": "매수인",
        "amount_label": "매매대금",
        "template_name": "land_sale.html",
    },
    "land_sale": {
        "label": "⛰️ [토지] 토지 매매",
        "short_label": "토지 매매",
        "category": "토지",
        "form_title": "토 지  매 매  계 약 서",
        "party1": "매도인",
        "party2": "매수인",
        "amount_label": "매매대금",
        "template_name": "land_sale.html",
    },
}

SALE_BUILDING_TYPES = ("apartment_sale", "real_estate_sale", "commercial_sale")
SALE_LAND_TYPES = ("farmland_sale", "land_sale")


def get_type_meta(contract_type: str) -> dict[str, str]:
    return CONTRACT_TYPE_META.get(contract_type, CONTRACT_TYPE_META["nursing_home_transfer"])


def parse_int(value: Any, default: int = 0) -> int:
    """문자열/숫자에서 정수를 안전하게 추출한다."""
    if value is None or isinstance(value, bool):
        return default
    if isinstance(value, (int, float)):
        return int(round(value))
    cleaned = re.sub(r"[^0-9\-]", "", str(value))
    if not cleaned or cleaned == "-":
        return default
    try:
        return int(cleaned)
    except ValueError:
        return default


def clean_str(value: Any) -> str:
    """None이나 비문자열을 안전하게 문자열로 정제한다."""
    if value is None:
        return ""
    return str(value).strip()


def number_to_korean(amount: int, zero_text: str = "") -> str:
    """정수 금액을 계약서 표준 한글 금액 문자열로 변환한다.

    예시:
        160000000 -> "일억육천만"
        16000000  -> "일천육백만"
        144000000 -> "일억사천사백만"
        0         -> zero_text (기본값 "")
    """
    amount = parse_int(amount, 0)
    if amount == 0:
        return zero_text

    prefix = "마이너스 " if amount < 0 else ""
    n = abs(amount)

    groups: list[str] = []
    group_idx = 0

    while n > 0:
        chunk = n % 10000
        n //= 10000
        if chunk > 0:
            chunk_str = _four_digits_to_korean(chunk)
            unit = _GROUP_UNITS[group_idx] if group_idx < len(_GROUP_UNITS) else ""
            groups.append(chunk_str + unit)
        group_idx += 1

    return prefix + "".join(reversed(groups))


def _four_digits_to_korean(chunk: int) -> str:
    """1~9999 범위 정수를 법률/계약서 표기(일천, 일백 등 '일' 유지)로 변환한다."""
    parts: list[str] = []
    for pos in range(3, -1, -1):
        divisor = 10**pos
        digit = (chunk // divisor) % 10
        if digit > 0:
            parts.append(_DIGITS[digit] + _SUB_UNITS[pos])
    return "".join(parts)


def format_comma(amount: int) -> str:
    """천 단위 콤마 포맷 (예: 160000000 -> '160,000,000')."""
    return f"{parse_int(amount):,}"


def format_won(amount: int) -> str:
    """원화 기호 포함 포맷 (예: 160000000 -> '₩ 160,000,000')."""
    val = parse_int(amount)
    return f"₩ {val:,}"


def calculate_balance(
    deposit: int,
    down_payment: int,
    interim_payment: int,
    loan_assumption: int = 0,
    lease_deposit_assumption: int = 0,
    escrow_amount: int = 0,
) -> int:
    """유형별 공제/승계/유보 항목을 반영한 잔금을 계산한다."""
    return (
        parse_int(deposit)
        - parse_int(down_payment)
        - parse_int(interim_payment)
        - parse_int(loan_assumption)
        - parse_int(lease_deposit_assumption)
        - parse_int(escrow_amount)
    )


def parse_float(value: Any) -> float | None:
    """면적 문자열(예: '1,000.5 ㎡')에서 양수 실수를 안전하게 추출한다."""
    if value is None:
        return None
    cleaned = re.sub(r"[^0-9.]", "", str(value))
    if not cleaned:
        return None
    try:
        val = float(cleaned)
        return val if val > 0 else None
    except ValueError:
        return None


def sqm_to_pyeong(area_value: Any) -> float | None:
    """제곱미터(㎡)를 평(3.305785㎡ 기준)으로 환산한다."""
    sqm = parse_float(area_value)
    if sqm is None:
        return None
    return sqm / 3.305785


def format_pyeong(area_value: Any) -> str:
    """㎡ 입력값을 '302.5평' 문자열로 반환한다."""
    py = sqm_to_pyeong(area_value)
    if py is None:
        return ""
    return f"{py:,.1f}평"


def format_price_per_pyeong(total_amount: int, area_value: Any) -> str:
    """총 매매대금과 면적(㎡)으로 3.3㎡(1평)당 참고 단가를 계산한다."""
    amt = parse_int(total_amount)
    py = sqm_to_pyeong(area_value)
    if amt <= 0 or py is None or py <= 0:
        return ""
    unit_price = int(round(amt / py))
    return f"₩ {unit_price:,} / 평"


def validate_amounts(
    deposit: int,
    down_payment: int,
    interim_payment: int,
    balance: int,
    contract_type: str = "housing_jeonse",
    loan_assumption: int = 0,
    lease_deposit_assumption: int = 0,
    escrow_amount: int = 0,
) -> tuple[bool, str]:
    """계약 유형별 총액과 세부 지급액(계약금·중도금·승계액/유보금·잔금)의 합계를 교차 검증한다."""
    meta = get_type_meta(contract_type)
    amt_name = meta["amount_label"]

    dep = parse_int(deposit)
    down = parse_int(down_payment)
    interim = parse_int(interim_payment)
    bal = parse_int(balance)
    loan = parse_int(loan_assumption) if contract_type in (SALE_BUILDING_TYPES + SALE_LAND_TYPES) else 0
    lease_dep = parse_int(lease_deposit_assumption) if contract_type in SALE_BUILDING_TYPES else 0
    escrow = parse_int(escrow_amount) if contract_type == "nursing_home_transfer" else 0

    if dep <= 0:
        return False, f"{amt_name}은 0원보다 커야 합니다. (현재 0원)"

    if any(x < 0 for x in (down, interim, bal, loan, lease_dep, escrow)):
        return (
            False,
            f"금액 오류: 세부 지급 항목 및 잔금은 음수가 될 수 없습니다. (현재 잔금: {bal:,}원)",
        )

    total = down + interim + bal + loan + lease_dep + escrow
    if total != dep:
        diff = total - dep
        parts = [f"계약금({down:,})"]
        if loan > 0:
            parts.append(f"융자승계({loan:,})")
        if lease_dep > 0:
            parts.append(f"임대보증금승계({lease_dep:,})")
        if interim > 0:
            parts.append(f"중도금({interim:,})")
        if escrow > 0:
            parts.append(f"유보금({escrow:,})")
        parts.append(f"잔금({bal:,})")
        expr = " + ".join(parts)
        return (
            False,
            f"금액 교차검증 불일치: {expr} = {total:,}원 ≠ {amt_name}({dep:,}원) [차액 {diff:+,}원]",
        )

    return True, f"금액 교차검증 정상 ({amt_name} {dep:,}원 일치)"


def parse_date(value: Any) -> date | None:
    """date/datetime/ISO 문자열을 date 객체로 변환한다."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y년 %m월 %d일"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def format_date_parts(value: Any, zero_pad: bool = False) -> tuple[str, str, str]:
    """날짜를 (년, 월, 일) 튜플로 분리한다. 미입력 시 ('', '', '')."""
    d = parse_date(value)
    if d is None:
        return "", "", ""
    if zero_pad:
        return f"{d.year}", f"{d.month:02d}", f"{d.day:02d}"
    return f"{d.year}", f"{d.month}", f"{d.day}"


def format_date_kr(value: Any, zero_pad: bool = True) -> str:
    """날짜를 '2026년 03월 28일' 형식으로 포맷한다. 미입력 시 수기용 '20    년    월    일'."""
    y, m, d = format_date_parts(value, zero_pad=zero_pad)
    if not y:
        return "20    년    월    일"
    return f"{y}년 {m}월 {d}일"


def calculate_lease_duration(start_value: Any, end_value: Any) -> str:
    """임대차 존속기간을 계산한다 (예: '24개월', '12개월 15일')."""
    start = parse_date(start_value)
    end = parse_date(end_value)
    if start is None or end is None or end < start:
        return ""

    months = (end.year - start.year) * 12 + (end.month - start.month)
    day_diff = end.day - start.day

    if day_diff in (-1, 0):
        return f"{months}개월" if months > 0 else "1일"

    if day_diff < -1:
        months -= 1
        approx_days = 30 + day_diff + 1
        if months <= 0:
            return f"{(end - start).days}일"
        return f"{months}개월 {approx_days}일"

    if months == 0:
        return f"{(end - start).days}일"
    return f"{months}개월 {day_diff}일"


def substitute_clause_variables(
    template: str,
    values: dict[str, Any],
    variable_meta: dict[str, dict[str, Any]] | None = None,
) -> str:
    """특약 본문 내 {변수명}을 안전하게 치환한다."""
    variable_meta = variable_meta or {}

    def _replace(match: re.Match[str]) -> str:
        key = match.group(1).strip()
        raw_val = values.get(key, "")
        meta = variable_meta.get(key, {})
        var_type = meta.get("type", "text")

        if raw_val is None or str(raw_val).strip() == "":
            return f"[{key}]"

        if var_type == "currency":
            digits = re.sub(r"[^0-9]", "", str(raw_val))
            if digits:
                return f"{int(digits):,}"
            return str(raw_val).strip()

        if var_type == "date":
            parsed = parse_date(raw_val)
            if parsed is not None:
                return format_date_kr(parsed, zero_pad=True)
            return str(raw_val).strip()

        return str(raw_val).strip()

    return re.sub(r"\{([^{}]+)\}", _replace, template)


def _visual_width(line: str) -> float:
    width = 0.0
    for ch in line:
        width += 1.0 if ord(ch) > 127 else 0.55
    return width


def estimate_clause_lines(text: str, chars_per_line: int = CHARS_PER_LINE) -> int:
    """A4 본문 특약란에서 차지하는 예상 줄 수를 계산한다."""
    if not text or not text.strip():
        return 0
    total = 0
    for raw_line in text.splitlines():
        if not raw_line.strip():
            total += 1
            continue
        w = _visual_width(raw_line)
        total += max(1, math.ceil(w / chars_per_line))
    return total


def resolve_smart_clause_placeholders(text: str, form_data: dict[str, Any]) -> str:
    """특약 문구 내 {잔금일}, {계약일}, {매매대금}, {보증금} 등 스마트 변수를 본문 입력값으로 실시간 자동 치환한다."""
    if not text or "{" not in text:
        return text

    dep = parse_int(form_data.get("deposit", 0))
    down = parse_int(form_data.get("down_payment", 0))
    interim = parse_int(form_data.get("interim_payment", 0))
    bal = parse_int(form_data.get("balance", 0))
    rent = parse_int(form_data.get("monthly_rent", 0))

    blank_mode = bool(form_data.get("blank_form_mode", False))
    bal_date_str = "20  년  월  일" if blank_mode else format_date_kr(form_data.get("balance_date"), zero_pad=True)
    con_date_str = "20  년  월  일" if blank_mode else format_date_kr(form_data.get("contract_date"), zero_pad=True)
    hand_date_str = "20  년  월  일" if blank_mode else format_date_kr(form_data.get("handover_date"), zero_pad=True)
    end_date_str = "20  년  월  일" if blank_mode else format_date_kr(form_data.get("lease_end_date"), zero_pad=True)

    dep_fmt = f"{dep:,}" if dep > 0 else ""
    smart_map: dict[str, str] = {
        "잔금일": bal_date_str,
        "잔금지급일": bal_date_str,
        "계약일": con_date_str,
        "계약작성일": con_date_str,
        "인도일": hand_date_str,
        "만료일": end_date_str,
        "매매대금": dep_fmt,
        "보증금": dep_fmt,
        "양도대금": dep_fmt,
        "권리금": dep_fmt,
        "계약금": f"{down:,}" if down > 0 else "",
        "중도금": f"{interim:,}" if interim > 0 else "",
        "잔금": f"{bal:,}" if bal > 0 else "",
        "월세": f"{rent:,}" if rent > 0 else "",
        "월차임": f"{rent:,}" if rent > 0 else "",
        "소재지": clean_str(form_data.get("address")),
        "임대인": clean_str(form_data.get("lessor_name")),
        "매도인": clean_str(form_data.get("lessor_name")),
        "양도인": clean_str(form_data.get("lessor_name")),
        "임차인": clean_str(form_data.get("lessee_name")),
        "매수인": clean_str(form_data.get("lessee_name")),
        "양수인": clean_str(form_data.get("lessee_name")),
    }

    def _sub(match: re.Match[str]) -> str:
        key = match.group(1).strip()
        if key in smart_map and smart_map[key]:
            return smart_map[key]
        return match.group(0)

    return re.sub(r"\{([^{}]+)\}", _sub, text)


def format_numbered_clauses(
    clauses: list[str],
    attachment_note: str = "",
    deposit_account: str = "",
) -> tuple[str, int, bool]:
    """선택된 특약 리스트를 번호 목록 문자열로 조합하고 별지 분리 여부(5항 초과 또는 10줄 초과)를 판정한다."""
    lines: list[str] = []
    idx = 1
    clause_count = 0
    for clause in clauses:
        cleaned = clause.strip()
        if not cleaned:
            continue
        clause_count += 1
        if re.match(r"^\d+[\.\)]\s+", cleaned):
            lines.append(cleaned)
        else:
            lines.append(f"{idx}. {cleaned}")
        idx += 1

    if attachment_note.strip():
        lines.append(f"   첨부 : {attachment_note.strip()}")
    if deposit_account.strip():
        lines.append(f"   입금계좌 : {deposit_account.strip()}")

    full_text = "\n".join(lines)
    est_lines = estimate_clause_lines(full_text)
    is_annex = (clause_count > MAX_INLINE_CLAUSE_COUNT) or (est_lines > MAX_INLINE_CLAUSE_LINES)
    return full_text, est_lines, is_annex


def build_contract_context(form_data: dict[str, Any]) -> dict[str, Any]:
    """UI 입력 데이터를 6대 계약 유형 HTML/DOCX/HWPX 공통 렌더링 컨텍스트로 정규화한다."""
    contract_type = clean_str(form_data.get("contract_type")) or "housing_jeonse"
    meta = get_type_meta(contract_type)

    deposit = parse_int(form_data.get("deposit", 0))
    down_payment = parse_int(form_data.get("down_payment", 0))
    interim_payment = parse_int(form_data.get("interim_payment", 0))
    balance = parse_int(form_data.get("balance", 0))

    # 유형별 확장 금액 필드
    loan_assumption = (
        parse_int(form_data.get("loan_assumption", 0))
        if contract_type in (SALE_BUILDING_TYPES + SALE_LAND_TYPES)
        else 0
    )
    lease_deposit_assumption = (
        parse_int(form_data.get("lease_deposit_assumption", 0))
        if contract_type in SALE_BUILDING_TYPES
        else 0
    )
    monthly_rent = parse_int(form_data.get("monthly_rent", 0))
    management_fee = parse_int(form_data.get("management_fee", 0))
    escrow_amount = (
        parse_int(form_data.get("escrow_amount", 0))
        if contract_type == "nursing_home_transfer"
        else 0
    )

    is_valid, validation_msg = validate_amounts(
        deposit=deposit,
        down_payment=down_payment,
        interim_payment=interim_payment,
        balance=balance,
        contract_type=contract_type,
        loan_assumption=loan_assumption,
        lease_deposit_assumption=lease_deposit_assumption,
        escrow_amount=escrow_amount,
    )

    blank_form_mode = bool(form_data.get("blank_form_mode", False))

    if blank_form_mode:
        interim_y, interim_m, interim_d = "20  ", "", ""
        balance_y, balance_m, balance_d = "20  ", "", ""
        contract_y, contract_m, contract_d = "", "", ""
        handover_date = None
        lease_end_date = None
        lease_duration = ""
    else:
        interim_y, interim_m, interim_d = format_date_parts(form_data.get("interim_date"))
        balance_y, balance_m, balance_d = format_date_parts(form_data.get("balance_date"))
        contract_y, contract_m, contract_d = format_date_parts(form_data.get("contract_date"))
        handover_date = parse_date(form_data.get("handover_date"))
        lease_end_date = parse_date(form_data.get("lease_end_date"))
        lease_duration = calculate_lease_duration(handover_date, lease_end_date)

    extra_parcels = "" if blank_form_mode else clean_str(form_data.get("extra_parcels"))
    if blank_form_mode:
        rendered_clauses: list[str] = []
        attachment_note = ""
        deposit_account = ""
    else:
        raw_clauses = list(form_data.get("rendered_clauses", []))
        rendered_clauses = [
            resolve_smart_clause_placeholders(str(c), form_data)
            for c in raw_clauses
            if str(c).strip()
        ]
        if contract_type in SALE_LAND_TYPES and extra_parcels:
            rendered_clauses.append(f"[다필지 추가 대상 토지]\n{extra_parcels}")
        attachment_note = clean_str(form_data.get("attachment_note"))
        deposit_account = clean_str(form_data.get("deposit_account"))

    clause_count = len([c for c in rendered_clauses if c.strip()])
    land_area_str = clean_str(form_data.get("land_area"))
    land_area_pyeong = format_pyeong(land_area_str)
    price_per_pyeong = format_price_per_pyeong(deposit, land_area_str)

    full_clauses_text, estimated_lines, is_annex = format_numbered_clauses(
        rendered_clauses,
        attachment_note=attachment_note,
        deposit_account=deposit_account,
    )

    if is_annex:
        body_clauses_text = f"※ 상세 특약은 첨부 [별지] 특약사항(총 {clause_count}항) 참조"
        body_clauses_lines = body_clauses_text.splitlines()
        annex_clauses_lines = full_clauses_text.splitlines()
    else:
        body_clauses_text = full_clauses_text
        body_clauses_lines = full_clauses_text.splitlines() if full_clauses_text else []
        annex_clauses_lines = []

    use_co_broker = bool(form_data.get("use_co_broker", False))
    color_mode = "bw" if clean_str(form_data.get("color_mode")) == "bw" else "color"
    cat_key = meta.get("category", "주거")
    sig_pal = DOC_SIGNATURE_PALETTES.get(cat_key, DOC_SIGNATURE_PALETTES["주거"])
    doc_id = clean_str(form_data.get("doc_id")) or "DOC-STANDARD"

    context: dict[str, Any] = {
        # 메타 / 상태 / 모던 디지털 테마
        "doc_id": doc_id,
        "doc_category": sig_pal["doc_category"],
        "doc_primary": sig_pal["primary"],
        "doc_accent": sig_pal["accent"],
        "doc_tint": sig_pal["tint"],
        "doc_money_bg": sig_pal["money_bg"],
        "clause_count": clause_count,
        "contract_type": contract_type,
        "contract_type_label": clean_str(form_data.get("contract_type_label")) or meta["label"],
        "form_title": clean_str(form_data.get("form_title")) or meta["form_title"],
        "party1_label": meta["party1"],
        "party2_label": meta["party2"],
        "amount_label": meta["amount_label"],
        "color_mode": color_mode,
        "is_color_mode": color_mode == "color",
        "draft_mode": bool(form_data.get("draft_mode", False)),
        "blank_form_mode": blank_form_mode,
        "watermark_text": "사전 검토용 초안 (계약 효력 없음)",
        "is_valid": is_valid,
        "validation_msg": validation_msg,
        # ① 부동산·시설·점포 표시
        "address": clean_str(form_data.get("address")),
        "land_category": clean_str(form_data.get("land_category")),
        "land_share_ratio": clean_str(form_data.get("land_share_ratio")),
        "land_area": land_area_str,
        "land_area_pyeong": land_area_pyeong,
        "price_per_pyeong": price_per_pyeong,
        "land_zoning": clean_str(form_data.get("land_zoning")),
        "land_status": clean_str(form_data.get("land_status")),
        "extra_parcels": extra_parcels,
        "building_structure": clean_str(form_data.get("building_structure")),
        "building_use": clean_str(form_data.get("building_use")),
        "building_area": clean_str(form_data.get("building_area")),
        "lease_part": clean_str(form_data.get("lease_part")),
        "lease_area": clean_str(form_data.get("lease_area")),
        # 유형별 전용 표시 항목
        "business_type": clean_str(form_data.get("business_type")),
        "facility_name": clean_str(form_data.get("facility_name")),
        "facility_capacity": clean_str(form_data.get("facility_capacity")),
        "facility_reg_no": clean_str(form_data.get("facility_reg_no")),
        "store_name": clean_str(form_data.get("store_name")),
        "new_lease_summary": clean_str(form_data.get("new_lease_summary")),
        # ② 계약내용 및 금액
        "deposit": deposit,
        "deposit_kr": number_to_korean(deposit, zero_text=""),
        "deposit_won": format_won(deposit) if deposit > 0 else "",
        "deposit_comma": format_comma(deposit) if deposit > 0 else "",
        "down_payment": down_payment,
        "down_payment_kr": number_to_korean(down_payment, zero_text=""),
        "down_payment_comma": format_comma(down_payment) if down_payment > 0 else "",
        "down_receipt_person": clean_str(form_data.get("down_receipt_person")),
        "interim_payment": interim_payment,
        "interim_payment_kr": number_to_korean(interim_payment, zero_text=""),
        "interim_payment_comma": format_comma(interim_payment) if interim_payment > 0 else "",
        "interim_year": interim_y,
        "interim_month": interim_m,
        "interim_day": interim_d,
        "balance": balance,
        "balance_kr": number_to_korean(balance, zero_text=""),
        "balance_comma": format_comma(balance) if balance > 0 else "",
        "balance_year": balance_y,
        "balance_month": balance_m,
        "balance_day": balance_d,
        # 매매 승계 항목
        "loan_assumption": loan_assumption,
        "loan_assumption_kr": number_to_korean(loan_assumption, zero_text=""),
        "loan_assumption_comma": format_comma(loan_assumption) if loan_assumption > 0 else "",
        "lease_deposit_assumption": lease_deposit_assumption,
        "lease_deposit_assumption_kr": number_to_korean(lease_deposit_assumption, zero_text=""),
        "lease_deposit_assumption_comma": (
            format_comma(lease_deposit_assumption) if lease_deposit_assumption > 0 else ""
        ),
        # 월세/상가 차임·관리비·부가세
        "monthly_rent": monthly_rent,
        "monthly_rent_kr": number_to_korean(monthly_rent, zero_text=""),
        "monthly_rent_comma": format_comma(monthly_rent) if monthly_rent > 0 else "",
        "monthly_rent_won": format_won(monthly_rent) if monthly_rent > 0 else "",
        "monthly_rent_day": clean_str(form_data.get("monthly_rent_day")),
        "rent_pay_type": "" if blank_form_mode else (clean_str(form_data.get("rent_pay_type")) or "선불"),
        "vat_type": "" if blank_form_mode else (clean_str(form_data.get("vat_type")) or "부가가치세 별도"),
        "management_fee": management_fee,
        "management_fee_kr": number_to_korean(management_fee, zero_text=""),
        "management_fee_comma": format_comma(management_fee) if management_fee > 0 else "",
        "management_fee_note": clean_str(form_data.get("management_fee_note")),
        # 요양원 유보금(에스크로)
        "escrow_amount": escrow_amount,
        "escrow_amount_kr": number_to_korean(escrow_amount, zero_text=""),
        "escrow_amount_comma": format_comma(escrow_amount) if escrow_amount > 0 else "",
        "escrow_period": clean_str(form_data.get("escrow_period")),
        "escrow_condition": clean_str(form_data.get("escrow_condition")),
        # 날짜/기간
        "handover_date_kr": format_date_kr(handover_date, zero_pad=True),
        "lease_end_date_kr": format_date_kr(lease_end_date, zero_pad=True),
        "lease_duration": lease_duration,
        "contract_date_kr": (
            f"{contract_y}년 {contract_m}월 {contract_d}일" if contract_y else "20    년    월    일"
        ),
        # ③ 특약사항 및 별지
        "full_clauses_text": full_clauses_text,
        "body_clauses_text": body_clauses_text,
        "body_clauses_lines": body_clauses_lines,
        "annex_clauses_lines": annex_clauses_lines,
        "estimated_clause_lines": estimated_lines,
        "is_annex": is_annex,
        "attachment_note": attachment_note,
        "deposit_account": deposit_account,
        # ④ 당사자 인적사항 (1:1 정확 매핑)
        "lessor_address": clean_str(form_data.get("lessor_address")),
        "lessor_rrn": clean_str(form_data.get("lessor_rrn")),
        "lessor_phone1": clean_str(form_data.get("lessor_phone1")),
        "lessor_phone2": clean_str(form_data.get("lessor_phone2")),
        "lessor_name": clean_str(form_data.get("lessor_name")),
        "lessee_address": clean_str(form_data.get("lessee_address")),
        "lessee_rrn": clean_str(form_data.get("lessee_rrn")),
        "lessee_phone1": clean_str(form_data.get("lessee_phone1")),
        "lessee_phone2": clean_str(form_data.get("lessee_phone2")),
        "lessee_name": clean_str(form_data.get("lessee_name")),
        # ⑤ 개업공인중개사 (주 중개사 + 공동중개사)
        "broker1_office": clean_str(form_data.get("broker1_office")),
        "broker1_address": clean_str(form_data.get("broker1_address")),
        "broker1_rep": clean_str(form_data.get("broker1_rep")),
        "broker1_reg_no": clean_str(form_data.get("broker1_reg_no")),
        "broker1_phone": clean_str(form_data.get("broker1_phone")),
        "broker1_associate": clean_str(form_data.get("broker1_associate")),
        "use_co_broker": use_co_broker,
        "broker2_office": clean_str(form_data.get("broker2_office")) if use_co_broker else "",
        "broker2_address": clean_str(form_data.get("broker2_address")) if use_co_broker else "",
        "broker2_rep": clean_str(form_data.get("broker2_rep")) if use_co_broker else "",
        "broker2_reg_no": clean_str(form_data.get("broker2_reg_no")) if use_co_broker else "",
        "broker2_phone": clean_str(form_data.get("broker2_phone")) if use_co_broker else "",
        "broker2_associate": clean_str(form_data.get("broker2_associate")) if use_co_broker else "",
    }

    return context
