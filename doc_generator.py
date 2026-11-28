"""통합 문서 생성 엔진 (6대 계약 유형 정식 HTML/PDF, 6대 유형 컬러/흑백 DOCX, 사용자 업로드 HWPX).

- HTML/PDF: 6대 계약 유형별 정식 A4 1장 + 2페이지 별지 (컬러/흑백 모드 + 주요사항 큰 글씨 강조).
- DOCX: 6대 계약 유형별 컬러/흑백 Word(.docx) 서식 (MS Word 및 아래아한글 공용, 주요사항 큰 글씨 강조).
- HWPX: 사용자가 아래아한글에서 직접 제작해 업로드한 `.hwpx` 템플릿이 등록된 경우에만 활성화.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, select_autoescape

from utils.formatter import CONTRACT_TYPE_META, get_type_meta
from utils.hwpx_helper import (
    extract_hwpx_placeholders,
    render_hwpx,
    validate_hwpx_bytes,
)

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = BASE_DIR / "templates"
HTML_TEMPLATE_DIR = TEMPLATES_DIR / "html"
DOCX_TEMPLATE_DIR = TEMPLATES_DIR / "docx"
HWPX_TEMPLATE_DIR = TEMPLATES_DIR / "hwpx"

HWPX_PLACEHOLDER_GUIDE: list[tuple[str, str]] = [
    ("{{소재지}} / {{address}}", "부동산/시설/점포 소재지"),
    ("{{시설명}} / {{facility_name}}", "요양원 시설 명칭"),
    ("{{정원}} / {{facility_capacity}}", "요양원 인가정원/현원"),
    ("{{기관기호}} / {{facility_reg_no}}", "장기요양기관 기호"),
    ("{{상호}} / {{store_name}}", "권리금 대상 상호(점포명)"),
    ("{{업종}} / {{business_type}}", "상가/권리금 영업 업종"),
    ("{{신규임대차조건}} / {{new_lease_summary}}", "권리금 신규 임대차 조건 요약"),
    ("{{보증금_한글}} / {{보증금_숫자}}", "보증금/매매대금/양도대금/권리금 한글·원화"),
    ("{{계약금_한글}} / {{영수자}}", "계약금 한글 및 영수자"),
    ("{{융자승계_한글}} / {{임대보증금승계_한글}}", "매매 융자승계·임대보증금승계 한글"),
    ("{{중도금_한글}} / {{잔금_한글}}", "중도금·잔금 한글"),
    ("{{월차임_한글}} / {{관리비_한글}} / {{부가세구분}}", "월세·상가 월차임·관리비·부가세"),
    ("{{유보금_한글}} / {{유보기간}}", "요양원 에스크로 유보금·유보기간"),
    ("{{인도일}} / {{만료일}} / {{존속기간}} / {{계약일}}", "주요 날짜 및 존속기간"),
    ("{{특약사항}} / {{특약사항_전문}}", "본문 특약사항 / 별지 전체 특약 전문"),
    ("{{임대인_주소}} / {{임대인_주민번호}} / {{임대인_전화1}} / {{임대인_성명}}", "임대인(매도인/양도인) 정보"),
    ("{{임차인_주소}} / {{임차인_주민번호}} / {{임차인_전화1}} / {{임차인_성명}}", "임차인(매수인/양수인) 정보"),
    ("{{중개사1_상호}} / {{중개사1_대표}} / {{중개사1_등록번호}} / {{중개사1_전화}}", "주 개업공인중개사 정보"),
    ("{{중개사2_상호}} / {{중개사2_대표}} / {{중개사2_등록번호}} / {{중개사2_전화}}", "공동중개사 정보"),
]

_jinja_env = Environment(
    loader=FileSystemLoader(str(HTML_TEMPLATE_DIR)),
    autoescape=select_autoescape(["html", "xml"]),
)


def _palette(color_mode: str) -> dict[str, Any]:
    """컬러 / 흑백 출력 모드별 Word(.docx) 색상 팔레트."""
    if color_mode == "bw":
        return {
            "outer_border": "000000",
            "inner_border": "000000",
            "header_fill": "F2F2F2",
            "title_rgb": (0, 0, 0),
            "label_rgb": (0, 0, 0),
            "money_rgb": (0, 0, 0),
            "date_rgb": (0, 0, 0),
            "addr_rgb": (0, 0, 0),
            "body_rgb": (0, 0, 0),
        }
    return {
        "outer_border": "1B365D",
        "inner_border": "64748B",
        "header_fill": "E8F0FA",
        "title_rgb": (27, 54, 93),
        "label_rgb": (21, 48, 87),
        "money_rgb": (29, 78, 216),
        "date_rgb": (153, 27, 27),
        "addr_rgb": (15, 23, 42),
        "body_rgb": (17, 24, 39),
    }


def _set_cell_shading(cell: Any, fill_hex: str) -> None:
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def _set_table_borders(table: Any, outer_hex: str = "1B365D", inner_hex: str = "64748B") -> None:
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="10" w:space="0" w:color="{outer_hex}"/>'
        f'<w:left w:val="single" w:sz="10" w:space="0" w:color="{outer_hex}"/>'
        f'<w:bottom w:val="single" w:sz="10" w:space="0" w:color="{outer_hex}"/>'
        f'<w:right w:val="single" w:sz="10" w:space="0" w:color="{outer_hex}"/>'
        f'<w:insideH w:val="single" w:sz="6" w:space="0" w:color="{inner_hex}"/>'
        f'<w:insideV w:val="single" w:sz="6" w:space="0" w:color="{inner_hex}"/>'
        "</w:tblBorders>"
    )
    table._tbl.tblPr.append(borders)


def _apply_korean_font(run: Any, font_name: str = "맑은 고딕") -> None:
    """Windows(맑은 고딕)·Linux/Mac·아래아한글 모두에서 한글 폰트가 깨지지 않도록 eastAsia 폰트를 지정한다."""
    from docx.oxml.ns import qn

    run.font.name = font_name
    r_pr = run._r.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        from docx.oxml import OxmlElement

        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)
    r_fonts.set(qn("w:cs"), font_name)


def _style_cell_text(
    cell: Any,
    text: str,
    bold: bool = False,
    size_pt: float = 7.7,
    align: str = "left",
    shade: str | None = None,
    rgb: tuple[int, int, int] = (0, 0, 0),
) -> None:
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.space_before = Pt(0.6)
    p.paragraph_format.space_after = Pt(0.6)
    run = p.add_run(text)
    run.bold = bold
    _apply_korean_font(run, "맑은 고딕")
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(*rgb)
    if shade:
        _set_cell_shading(cell, shade)


def _docx_intro_and_articles(contract_type: str) -> tuple[str, str, list[str]]:
    if contract_type in ("land_sale", "farmland_sale"):
        obj_name = "농지" if contract_type == "farmland_sale" else "토지"
        return (
            f"아래 {obj_name}에 대하여 매도인과 매수인은 상호 합의하에 다음과 같이 {obj_name} 매매계약을 체결한다.",
            "2. 거래금액 및 지급시기  제1조 [매매대금 및 지급방법] 매수인은 매매대금을 아래와 같이 지불하기로 한다.",
            [
                "위 제1조의 매매대금은 개업공인중개사의 입회하에 지불하기로 한다.",
                f"제2조 [소유권 이전 및 {obj_name} 인도] 매도인은 잔금 수령과 동시에 소유권이전등기 서류를 교부하고 등기절차에 협력하며, {obj_name} 인도일은 {{{{ handover_date_kr }}}}로 한다.",
                "제3조 [제한물권 등의 소멸] 매도인은 승계 합의된 융자금 외 저당권·지상권·지역권·압류 및 미납 제세공과금을 잔금일까지 제거하여 완전한 소유권을 이전한다.",
                f"제4조 [제세공과금 및 수익의 귀속] {obj_name} 수익 귀속과 공과금·부담금 정산은 인도일을 기준으로 하며, 지방세 납부의무는 지방세법 규정에 따른다.",
                "제5조 [계약의 해제] 중도금(없을 때는 잔금) 지급 전까지 매도인은 계약금 배액 상환, 매수인은 계약금 포기로 계약을 해제할 수 있다.",
                "제6조 [채무불이행·손해배상·중개보수] 불이행 시 서면 최고 후 해제 및 계약금 기준 손해배상하며, 중개보수는 쌍방이 각각 지급하고 확인·설명서 및 공제증서 사본을 교부한다.",
            ],
        )

    if contract_type in ("real_estate_sale", "apartment_sale", "commercial_sale"):
        target_name = (
            "아파트"
            if contract_type == "apartment_sale"
            else ("상가건물" if contract_type == "commercial_sale" else "부동산")
        )
        return (
            f"아래 {target_name}에 대하여 매도인과 매수인은 합의하에 다음과 같이 매매계약을 체결한다.",
            "2. 계약내용  제1조 [매매대금 및 지급시기] 매수인은 매매대금을 아래와 같이 지불하기로 한다.",
            [
                "위 제1조의 매매대금은 개업공인중개사의 입회하에 지불하기로 한다.",
                f"제2조 [소유권 이전 등] 매도인은 잔금 수령과 동시에 소유권이전등기 서류를 교부하며, {target_name} 인도일은 {{{{ handover_date_kr }}}}로 한다.",
                "제3조 [제한물권 등의 소멸] 매도인은 승계 합의된 권리 외 저당권·지상권·임차권 및 제세공과금 미납을 잔금일까지 제거하여 완전한 소유권을 이전한다.",
                "제4조 [지방세 등] 부동산 수익 귀속과 제세공과금 부담은 인도일을 기준으로 하며, 지방세 납부의무는 지방세법 규정에 의한다.",
                "제5조 [계약의 해제] 중도금(없을 때는 잔금) 지급 전까지 매도인은 계약금 배액 상환, 매수인은 계약금 포기로 계약을 해제할 수 있다.",
                "제6조 [채무불이행·손해배상·중개보수] 불이행 시 서면 최고 후 해제 및 계약금 기준 손해배상하며, 중개보수는 쌍방이 각각 지급하고 확인·설명서 및 공제증서 사본을 교부한다.",
            ],
        )

    if contract_type == "housing_wolse":
        return (
            "아래 부동산에 대하여 임대인과 임차인은 합의하에 다음과 같이 주택 임대차(월세) 계약을 체결한다.",
            "2. 계약내용  제1조 [보증금·차임 및 관리비] 위 부동산의 보증금, 월 차임 및 관리비를 아래와 같이 약정한다.",
            [
                "위 제1조의 보증금은 개업공인중개사의 입회하에 지불하기로 한다.",
                "제2조 [존속기간] 임대인은 {{ handover_date_kr }} 까지 인도하며, 임대차기간은 인도일로부터 {{ lease_end_date_kr }} 까지로 한다.({{ lease_duration }})",
                "제3조 [용도변경 및 전대 등] 임차인은 임대인의 동의 없이 용도·구조를 변경하거나 전대·임차권 양도·담보제공을 하지 못한다.",
                "제4조 [계약의 종료] 종료 시 임차인은 원상회복하여 반환하고, 임대인은 연체차임·미납관리비·손해배상금을 공제한 잔여 보증금을 반환한다.",
                "제5조 [계약의 해제 및 2기 차임연체 해지] 중도금(없을 때는 잔금) 전 계약금 포기/배액 상환 해제 가능하며, 2기 차임 연체 시 임대인은 계약을 해지할 수 있다.",
                "제6조 [채무불이행·중개보수·확인설명서] 불이행 시 계약금 기준 손해배상하며, 중개보수는 쌍방이 각각 지급하고 확인·설명서 및 공제증서 사본을 교부한다.",
            ],
        )

    if contract_type == "commercial_lease":
        return (
            "아래 상가건물에 대하여 임대인과 임차인은 합의하에 다음과 같이 상가건물 임대차계약을 체결한다.",
            "2. 계약내용  제1조 [보증금·차임(부가세)·관리비] 보증금, 월 차임 및 관리비를 아래와 같이 약정한다.",
            [
                "위 제1조의 보증금은 개업공인중개사의 입회하에 지불하며, 세금계산서는 법령에 따라 발행한다.",
                "제2조 [존속기간] 임대인은 {{ handover_date_kr }} 까지 인도하며, 임대차기간은 인도일로부터 {{ lease_end_date_kr }} 까지로 한다.({{ lease_duration }})",
                "제3조 [용도·업종 제한 및 인허가] 임차인은 동의 없이 용도·구조를 변경하거나 전대·양도할 수 없으며, 영업 인허가 및 소방·위생 기준 충족은 임차인 책임으로 한다.",
                "제4조 [원상회복 및 계약의 종료] 종료 시 원상회복하여 명도하고, 임대인은 연체차임·관리비·공과금을 공제한 보증금 잔액을 반환한다.",
                "제5조 [계약의 해제 및 3기 차임연체 해지] 상가건물 임대차보호법에 따라 3기의 차임액 연체 시 임대인은 계약을 해지할 수 있다.",
                "제6조 [계약갱신요구권·권리금·중개보수] 계약갱신요구권 및 권리금 회수기회 보호는 상가임대차보호법에 따르며, 중개보수는 쌍방 각각 지급하고 확인·설명서 및 공제증서 사본을 교부한다.",
            ],
        )

    if contract_type == "nursing_home_transfer":
        return (
            "아래 노인요양시설(장기요양기관)에 대하여 양도인과 양수인은 합의하에 다음과 같이 양도·양수 계약을 체결한다.",
            "2. 양도대금 및 유보금(에스크로)  제1조 [총 양도대금 및 지급방법] 아래와 같이 약정한다.",
            [
                "총 양도대금 = 계약금 + 중도금 + 잔금 + 유보금(에스크로)이며, 모든 지급은 개업공인중개사 입회하에 진행한다.",
                "제2조 [지위승계 및 인허가 변경 협조] 양도인은 양도기준일({{ handover_date_kr }})까지 노인의료복지시설 설치신고 변경 및 장기요양기관 지정승계에 적극 협조한다.",
                "제3조 [행정처분·부당청구 면책 및 유보금 정산] 양도기준일 이전 운영기간의 현지조사·부당청구 환수·과징금·업무정지 책임은 전액 양도인에게 귀속되며, 양수인은 유보금에서 우선 공제·정산할 수 있다.",
                "제4조 [종사자 노무·퇴직금 및 4대보험 정산] 양도기준일 이전 전체 종사자의 급여·퇴직충당금·연차수당·4대보험료는 양도인이 전액 정산 완료한다.",
                "제5조 [수급자 승계·시설비품 인계·채무면책] 양도기준일 이전 본인부담금·공단급여는 양도인 귀속이며, 시설·비품 일체를 현 상태로 인계하고 이전 거래처·리스·공과금·조세 채무는 양도인이 변제한다.",
                "제6조 [해제·손해배상·중개보수] 중도금 전 계약금 포기/배액 상환 해제 가능하며, 불이행 시 계약금 기준 배상하고 중개보수는 쌍방 각각 지급한다.",
            ],
        )

    if contract_type == "premium_rights":
        return (
            "아래 상가 점포의 시설·집기비품 및 영업권(권리금)에 대하여 양도인과 양수인은 합의하에 다음과 같이 권리금 계약을 체결한다.",
            "2. 권리금 총액 및 지급시기  제1조 [권리금 총액 및 지급방법] 아래와 같이 약정한다.",
            [
                "본 권리금의 지급 및 시설·영업권 인계는 개업공인중개사의 입회하에 진행한다.",
                "제2조 [시설·비품 인도 및 영업권 양도] 잔금 수령과 동시에 시설·인테리어·집기비품 일체 및 영업권을 현 상태로 인도({{ handover_date_kr }})하며, 인허가 지위승계에 적극 협조한다.",
                "제3조 [신규 임대차계약 체결 조건부 효력] 건물 임대인과 신규 임대차계약이 양수인 귀책 없이 성립되지 않을 경우 본 계약은 자동 해제되고 수령 금원 전액을 즉시 반환한다.",
                "제4조 [기존 채무·공과금 및 제세 정산] 잔금일 이전 차임·관리비·공과금·세금·거래처 외상 및 종업원 임금은 양도인이 전액 정산한다.",
                "제5조 [경업금지 의무] 양도인은 상법 제41조 및 특약에 따라 인근 동일 상권 내 동종·유사 업종을 영위하지 않기로 한다.",
                "제6조 [해제·손해배상·용역보수] 중도금 전 계약금 포기/배액 상환 해제 가능하며, 불이행 시 계약금을 기준으로 배상하고 알선 보수는 별도 약정에 따른다.",
            ],
        )

    return (
        "아래 부동산에 대하여 임대인과 임차인은 합의하에 다음과 같이 임대차계약을 체결한다.",
        "2. 계약내용  제1조 [보증금 및 지불시기] 위 부동산의 임대차에 관하여 보증금과 지불시기를 다음과 같이 약정한다.",
        [
            "위 제1항의 보증금은 개업공인중개사의 입회하에 지불하기로 한다.",
            "제2조 [존속기간] 임대인은 {{ handover_date_kr }} 까지 임차인에게 인도하며, 임대차기간은 인도일로부터 {{ lease_end_date_kr }} 까지로 한다.({{ lease_duration }})",
            "제3조 [용도변경 및 전대 등] 임차인은 임대인의 동의 없이는 위 부동산의 용도나 구조 등을 변경하거나 전대·임차권 양도 또는 담보제공을 하지 못하며 임대차 목적 이외의 용도에 사용할 수 없다.",
            "제4조 [계약의 종료] 임대차 계약이 종료된 경우 임차인은 원상으로 회복하여 반환하고, 임대인은 연체임대료·손해배상 금액을 제하고 보증금 잔액을 반환한다.",
            "제5조 [계약의 해제] 임차인이 중도금(중도금이 없을 때는 잔금)을 지불하기 전까지 임대인은 계약금의 배액을 상환하고 임차인은 계약금을 포기하고 이 계약을 해제할 수 있다.",
            "제6조 [채무불이행과 손해배상] 당사자 일방이 불이행 시 서면 최고 후 계약을 해제할 수 있으며, 별도 약정이 없는 한 계약금을 손해배상의 기준으로 본다.",
            "제7조 [중개보수 및 확인설명서] 중개보수는 계약 당사자 쌍방이 각각 지급하며, 개업공인중개사는 중개대상물 확인·설명서 및 업무보증관계증서(공제증서 등) 사본을 교부한다.",
        ],
    )


def get_docx_template_path(contract_type: str, color_mode: str = "color") -> Path:
    safe_type = contract_type if contract_type in CONTRACT_TYPE_META else "housing_jeonse"
    safe_mode = "bw" if color_mode == "bw" else "color"
    return DOCX_TEMPLATE_DIR / f"{safe_type}_{safe_mode}.docx"


def create_docx_template(
    contract_type: str,
    color_mode: str = "color",
    output_path: Path | None = None,
) -> Path:
    """6대 계약 유형별 컬러/흑백 정식 A4 1장 + 별지 Word(.docx) 템플릿 생성 (주요사항 큰 글씨 강조)."""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Mm, Pt, RGBColor

    meta = get_type_meta(contract_type)
    pal = _palette(color_mode)
    target = output_path or get_docx_template_path(contract_type, color_mode)
    target.parent.mkdir(parents=True, exist_ok=True)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(8.5)
    section.bottom_margin = Mm(8.5)
    section.left_margin = Mm(10)
    section.right_margin = Mm(10)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "맑은 고딕"
    normal_style.font.size = Pt(7.8)
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    ns_rpr = normal_style._element.get_or_add_rPr()
    ns_rfonts = ns_rpr.rFonts
    if ns_rfonts is None:
        ns_rfonts = OxmlElement("w:rFonts")
        ns_rpr.insert(0, ns_rfonts)
    ns_rfonts.set(qn("w:eastAsia"), "맑은 고딕")

    # 초안 배너
    p_draft = doc.add_paragraph()
    p_draft.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_draft.paragraph_format.space_after = Pt(1)
    r_draft = p_draft.add_run("{{ draft_banner }}")
    r_draft.bold = True
    r_draft.font.size = Pt(8.5)
    r_draft.font.color.rgb = RGBColor(*pal["date_rgb"])

    # 제목 (큰 글씨 강조: 17pt Extra Bold)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("{{ form_title }}")
    r_title.bold = True
    r_title.font.size = Pt(17.0)
    r_title.font.color.rgb = RGBColor(*pal["title_rgb"])

    intro_text, pay_heading, articles = _docx_intro_and_articles(contract_type)
    p_intro = doc.add_paragraph(intro_text)
    p_intro.paragraph_format.space_after = Pt(1.5)
    p_intro.runs[0].font.size = Pt(7.3)
    p_intro.runs[0].font.color.rgb = RGBColor(*pal["body_rgb"])

    # ① 대상 표시 표 (소재지·시설명·상호 9.5pt Bold 강조)
    p_s1 = doc.add_paragraph()
    p_s1.paragraph_format.space_after = Pt(1)
    sec1_title = (
        "1. 양도·양수 대상 요양시설의 표시"
        if contract_type == "nursing_home_transfer"
        else (
            "1. 대상 점포 및 신규 임대차 조건 요약"
            if contract_type == "premium_rights"
            else (
                "1. 농지의 표시"
                if contract_type == "farmland_sale"
                else (
                    "1. 토지의 표시"
                    if contract_type == "land_sale"
                    else (
                        "1. 상가건물의 표시"
                        if contract_type in ("commercial_lease", "commercial_sale")
                        else "1. 부동산의 표시"
                    )
                )
            )
        )
    )
    r_s1 = p_s1.add_run(sec1_title)
    r_s1.bold = True
    r_s1.font.size = Pt(8.5)
    r_s1.font.color.rgb = RGBColor(*pal["title_rgb"])

    h_fill = pal["header_fill"]
    l_rgb = pal["label_rgb"]
    a_rgb = pal["addr_rgb"]
    b_rgb = pal["body_rgb"]

    if contract_type in ("land_sale", "farmland_sale"):
        t1 = doc.add_table(rows=4, cols=4)
        _set_table_borders(t1, pal["outer_border"], pal["inner_border"])
        t1.cell(0, 1).merge(t1.cell(0, 3))
        _style_cell_text(t1.cell(0, 0), "소재지 (지번)", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(0, 1), "{{ address }}", bold=True, size_pt=9.5, rgb=a_rgb)
        _style_cell_text(t1.cell(1, 0), "지  목", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 1), "{{ land_category }}", bold=True, rgb=b_rgb)
        _style_cell_text(t1.cell(1, 2), "공부상 면적", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 3), "{{ land_area }} ㎡{% if land_area_pyeong %} (약 {{ land_area_pyeong }}){% endif %}", bold=True, align="right", rgb=b_rgb)
        _style_cell_text(t1.cell(2, 0), "용도지역·지구", bold=True, size_pt=7.3, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 1), "{{ land_zoning }}", rgb=b_rgb)
        _style_cell_text(t1.cell(2, 2), "토지 현황", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 3), "{{ land_status }}", rgb=b_rgb)
        _style_cell_text(t1.cell(3, 0), "매 매 대 상", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(3, 1), "{{ lease_part if lease_part else ('' if blank_form_mode else '토지 전부') }}", bold=True, rgb=b_rgb)
        _style_cell_text(t1.cell(3, 2), "평당 참고단가", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(3, 3), "{{ price_per_pyeong }}", align="right", rgb=b_rgb)
    elif contract_type == "nursing_home_transfer":
        t1 = doc.add_table(rows=4, cols=4)
        _set_table_borders(t1, pal["outer_border"], pal["inner_border"])
        t1.cell(1, 1).merge(t1.cell(1, 3))
        t1.cell(3, 1).merge(t1.cell(3, 3))
        _style_cell_text(t1.cell(0, 0), "시 설 명 칭", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(0, 1), "{{ facility_name }}", bold=True, size_pt=9.5, rgb=a_rgb)
        _style_cell_text(t1.cell(0, 2), "장기요양기관기호", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(0, 3), "{{ facility_reg_no }}", rgb=b_rgb)
        _style_cell_text(t1.cell(1, 0), "소 재 지", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 1), "{{ address }}", bold=True, size_pt=9.5, rgb=a_rgb)
        _style_cell_text(t1.cell(2, 0), "인가정원/현원", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 1), "{{ facility_capacity }}", bold=True, rgb=b_rgb)
        _style_cell_text(t1.cell(2, 2), "구조·용도/면적", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 3), "{{ building_structure }} ({{ building_use }}) / {{ lease_area }} ㎡", rgb=b_rgb)
        _style_cell_text(t1.cell(3, 0), "양 도 범 위", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(3, 1), "{{ lease_part }}", rgb=b_rgb)
    elif contract_type == "premium_rights":
        t1 = doc.add_table(rows=4, cols=4)
        _set_table_borders(t1, pal["outer_border"], pal["inner_border"])
        t1.cell(1, 1).merge(t1.cell(1, 3))
        t1.cell(3, 1).merge(t1.cell(3, 3))
        _style_cell_text(t1.cell(0, 0), "상호(점포명)", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(0, 1), "{{ store_name }}", bold=True, size_pt=9.5, rgb=a_rgb)
        _style_cell_text(t1.cell(0, 2), "영업 업종", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(0, 3), "{{ business_type }}", rgb=b_rgb)
        _style_cell_text(t1.cell(1, 0), "소 재 지", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 1), "{{ address }}", bold=True, size_pt=9.5, rgb=a_rgb)
        _style_cell_text(t1.cell(2, 0), "점포 층·호수", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 1), "{{ lease_part }}", rgb=b_rgb)
        _style_cell_text(t1.cell(2, 2), "점포 면적", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 3), "{{ lease_area }} ㎡", align="right", rgb=b_rgb)
        _style_cell_text(t1.cell(3, 0), "신규 임대차 조건", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(3, 1), "{{ new_lease_summary }}", bold=True, rgb=b_rgb)
    elif contract_type == "commercial_lease":
        t1 = doc.add_table(rows=4, cols=7)
        _set_table_borders(t1, pal["outer_border"], pal["inner_border"])
        t1.cell(0, 1).merge(t1.cell(0, 6))
        t1.cell(2, 1).merge(t1.cell(2, 4))
        t1.cell(3, 1).merge(t1.cell(3, 6))
        _style_cell_text(t1.cell(0, 0), "소 재 지", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(0, 1), "{{ address }}", bold=True, size_pt=9.5, rgb=a_rgb)
        _style_cell_text(t1.cell(1, 0), "건 물", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 1), "구조", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 2), "{{ building_structure }}", rgb=b_rgb)
        _style_cell_text(t1.cell(1, 3), "용도", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 4), "{{ building_use }}", rgb=b_rgb)
        _style_cell_text(t1.cell(1, 5), "면적", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 6), "{{ building_area }} ㎡", align="right", rgb=b_rgb)
        _style_cell_text(t1.cell(2, 0), "임대할부분", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 1), "{{ lease_part }}", bold=True, rgb=b_rgb)
        _style_cell_text(t1.cell(2, 5), "계약면적", bold=True, size_pt=7.1, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 6), "{{ lease_area }} ㎡", align="right", rgb=b_rgb)
        _style_cell_text(t1.cell(3, 0), "영업 업종", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(3, 1), "{{ business_type }}", rgb=b_rgb)
    else:
        t1 = doc.add_table(rows=4, cols=7)
        _set_table_borders(t1, pal["outer_border"], pal["inner_border"])
        t1.cell(0, 1).merge(t1.cell(0, 6))
        t1.cell(3, 1).merge(t1.cell(3, 4))
        _style_cell_text(t1.cell(0, 0), "소 재 지", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(0, 1), "{{ address }}", bold=True, size_pt=9.5, rgb=a_rgb)
        _style_cell_text(t1.cell(1, 0), "토 지", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 1), "지목", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 2), "{{ land_category }}", rgb=b_rgb)
        _style_cell_text(t1.cell(1, 3), "대지권비율", bold=True, size_pt=7.0, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 4), "{{ land_share_ratio }}", rgb=b_rgb)
        _style_cell_text(t1.cell(1, 5), "면적", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(1, 6), "{{ land_area }} ㎡", align="right", rgb=b_rgb)
        _style_cell_text(t1.cell(2, 0), "건 물", bold=True, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 1), "구조", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 2), "{{ building_structure }}", rgb=b_rgb)
        _style_cell_text(t1.cell(2, 3), "용도", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 4), "{{ building_use }}", rgb=b_rgb)
        _style_cell_text(t1.cell(2, 5), "면적", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(2, 6), "{{ building_area }} ㎡", align="right", rgb=b_rgb)
        part_lbl = "매매대상" if contract_type in ("real_estate_sale", "apartment_sale", "commercial_sale") else "임대할부분"
        _style_cell_text(t1.cell(3, 0), part_lbl, bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(3, 1), "{{ lease_part }}", bold=True, rgb=b_rgb)
        _style_cell_text(t1.cell(3, 5), "면적", bold=True, size_pt=7.2, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t1.cell(3, 6), "{{ lease_area }} ㎡", align="right", rgb=b_rgb)

    # ② 대금 및 지급시기 표 (핵심 금액 9.5pt Bold, 잔금일 9.0pt Bold 강조)
    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(2)
    p_s2.paragraph_format.space_after = Pt(1)
    r_s2 = p_s2.add_run(pay_heading)
    r_s2.bold = True
    r_s2.font.size = Pt(7.8)
    r_s2.font.color.rgb = RGBColor(*pal["title_rgb"])

    # (라벨, 한글금액표현, 우측설명표현, 영수자표현, 핵심금액강조여부, 날짜강조여부)
    pay_rows: list[tuple[str, str, str, str | None, bool, bool]] = [
        (meta["amount_label"], "ㅡ금 {{ deposit_kr }}", "원정 ({{ deposit_won }})", None, True, False),
        ("계 약 금", "ㅡ금 {{ down_payment_kr }}", "원정은 계약시에 지불하고 영수함.", "{{ down_receipt_person }} (인)", True, False),
    ]
    if contract_type in ("real_estate_sale", "apartment_sale", "commercial_sale"):
        pay_rows.append(("융자승계", "ㅡ금 {{ loan_assumption_kr }}", "원정은 매수인이 금융기관 대출 채무를 승계함.", None, False, False))
        pay_rows.append(("보증금승계", "ㅡ금 {{ lease_deposit_assumption_kr }}", "원정은 현 임차인의 임대차보증금을 승계함.", None, False, False))
    elif contract_type in ("land_sale", "farmland_sale"):
        pay_rows.append(("융자승계", "ㅡ금 {{ loan_assumption_kr }}", "원정은 매수인이 금융기관 근저당 채무를 승계함.", None, False, False))

    pay_rows.append(("중 도 금", "ㅡ금 {{ interim_payment_kr }}", "원정은 {{ interim_year }} 년 {{ interim_month }} 월 {{ interim_day }} 일에 지불하며,", None, False, False))
    pay_rows.append(("잔  금", "ㅡ금 {{ balance_kr }}", "원정은 {{ balance_year }}년 {{ balance_month }}월 {{ balance_day }}일에 지불한다.", None, True, True))

    if contract_type == "housing_wolse":
        pay_rows.append(("월 차 임", "ㅡ금 {{ monthly_rent_kr }}", "원정은 매월 {{ monthly_rent_day }}일 ({{ rent_pay_type }}) 지급한다.", None, True, True))
        pay_rows.append(("관 리 비", "ㅡ금 {{ management_fee_kr }}", "원정 ({{ management_fee_note }})", None, False, False))
    elif contract_type == "commercial_lease":
        pay_rows.append(("월 차 임", "ㅡ금 {{ monthly_rent_kr }}", "원정 ({{ vat_type }}), 매월 {{ monthly_rent_day }}일 ({{ rent_pay_type }}) 지급", None, True, True))
        pay_rows.append(("관 리 비", "ㅡ금 {{ management_fee_kr }}", "원정 ({{ management_fee_note }})", None, False, False))
    elif contract_type == "nursing_home_transfer":
        pay_rows.append(("유 보 금", "ㅡ금 {{ escrow_amount_kr }}", "원정은 잔금일로부터 {{ escrow_period }}간 에스크로 유보 ({{ escrow_condition }})", None, True, True))

    t2 = doc.add_table(rows=len(pay_rows), cols=5)
    _set_table_borders(t2, pal["outer_border"], pal["inner_border"])
    for r_idx, (lbl, kr_expr, desc_expr, receipt_expr, emph_money, emph_date) in enumerate(pay_rows):
        _style_cell_text(t2.cell(r_idx, 0), lbl, bold=True, size_pt=7.5, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(
            t2.cell(r_idx, 1),
            kr_expr,
            bold=True,
            size_pt=9.5 if emph_money else 7.8,
            rgb=pal["money_rgb"] if emph_money else b_rgb,
        )
        if receipt_expr is not None:
            _style_cell_text(t2.cell(r_idx, 2), desc_expr, size_pt=7.3, rgb=b_rgb)
            _style_cell_text(t2.cell(r_idx, 3), "영수자", bold=True, size_pt=7.1, align="center", shade=h_fill, rgb=l_rgb)
            _style_cell_text(t2.cell(r_idx, 4), receipt_expr, size_pt=7.3, align="right", rgb=pal["date_rgb"])
        else:
            t2.cell(r_idx, 2).merge(t2.cell(r_idx, 4))
            desc_size = 9.5 if r_idx == 0 else (9.0 if emph_date else 7.4)
            desc_rgb = pal["money_rgb"] if r_idx == 0 else (pal["date_rgb"] if emph_date else b_rgb)
            _style_cell_text(t2.cell(r_idx, 2), desc_expr, bold=(r_idx == 0 or emph_date), size_pt=desc_size, rgb=desc_rgb)

    # 표준 조항 (제2조 존속기간/인도일은 굵게 강조)
    for idx, art in enumerate(articles):
        pa = doc.add_paragraph(art)
        pa.paragraph_format.space_before = Pt(0.5)
        pa.paragraph_format.space_after = Pt(0.5)
        r_art = pa.runs[0]
        if idx == 1:  # 제2조 존속기간/인도일 강조
            r_art.bold = True
            r_art.font.size = Pt(8.3)
            r_art.font.color.rgb = RGBColor(*pal["date_rgb"])
        else:
            r_art.font.size = Pt(6.95)
            r_art.font.color.rgb = RGBColor(*b_rgb)

    # 특약사항
    p_sp_h = doc.add_paragraph()
    p_sp_h.paragraph_format.space_before = Pt(2)
    p_sp_h.paragraph_format.space_after = Pt(0.8)
    r_sp_h = p_sp_h.add_run("특약사항")
    r_sp_h.bold = True
    r_sp_h.font.size = Pt(8.4)
    r_sp_h.font.color.rgb = RGBColor(*pal["title_rgb"])

    p_sp_b = doc.add_paragraph("{{ body_clauses_text }}")
    p_sp_b.paragraph_format.space_after = Pt(2.5)
    p_sp_b.runs[0].font.size = Pt(7.7)
    p_sp_b.runs[0].font.color.rgb = RGBColor(*b_rgb)

    # 마감 문구 및 계약일
    p1_lbl = meta["party1"]
    p2_lbl = meta["party2"]
    p_close = doc.add_paragraph(
        f"본 계약에 대하여 {p1_lbl}과 {p2_lbl}은 이의 없음을 확인하고 각각 서명 또는 날인 후 {p1_lbl}, {p2_lbl}, 개업공인중개사가 각 1통씩 보관한다."
    )
    p_close.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_close.paragraph_format.space_after = Pt(1.2)
    p_close.runs[0].font.size = Pt(7.0)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date.paragraph_format.space_after = Pt(2)
    r_date = p_date.add_run("{{ contract_date_kr }}")
    r_date.bold = True
    r_date.font.size = Pt(9.3)
    r_date.font.color.rgb = RGBColor(*pal["title_rgb"])

    # 당사자 및 개업공인중개사 날인표 (9행 6열 — 여유로운 행 높이 및 가독성 강화)
    t3 = doc.add_table(rows=9, cols=6)
    _set_table_borders(t3, pal["outer_border"], pal["inner_border"])
    for row in t3.rows:
        row.height = Mm(6.8)

    t3.cell(0, 0).merge(t3.cell(1, 0))
    t3.cell(0, 2).merge(t3.cell(0, 5))
    _style_cell_text(t3.cell(0, 0), p1_lbl, bold=True, size_pt=8.0, align="center", shade=h_fill, rgb=l_rgb)
    _style_cell_text(t3.cell(0, 1), "주소", bold=True, size_pt=7.8, align="center", shade=h_fill, rgb=l_rgb)
    _style_cell_text(t3.cell(0, 2), "{{ lessor_address }}", size_pt=8.0, rgb=b_rgb)
    _style_cell_text(t3.cell(1, 1), "주민(사업자)번호", bold=True, size_pt=7.4, align="center", shade=h_fill, rgb=l_rgb)
    _style_cell_text(t3.cell(1, 2), "{{ lessor_rrn }}", size_pt=8.0, rgb=b_rgb)
    _style_cell_text(t3.cell(1, 3), "전화", bold=True, size_pt=7.6, align="center", shade=h_fill, rgb=l_rgb)
    _style_cell_text(t3.cell(1, 4), "{{ lessor_phone1 }}{% if lessor_phone2 %} / {{ lessor_phone2 }}{% endif %}", size_pt=7.8, rgb=b_rgb)
    _style_cell_text(t3.cell(1, 5), "성명: {{ lessor_name }} (인)", bold=True, size_pt=8.0, align="right", rgb=pal["date_rgb"])

    t3.cell(2, 0).merge(t3.cell(3, 0))
    t3.cell(2, 2).merge(t3.cell(2, 5))
    _style_cell_text(t3.cell(2, 0), p2_lbl, bold=True, size_pt=8.0, align="center", shade=h_fill, rgb=l_rgb)
    _style_cell_text(t3.cell(2, 1), "주소", bold=True, size_pt=7.8, align="center", shade=h_fill, rgb=l_rgb)
    _style_cell_text(t3.cell(2, 2), "{{ lessee_address }}", size_pt=8.0, rgb=b_rgb)
    _style_cell_text(t3.cell(3, 1), "주민(사업자)번호", bold=True, size_pt=7.4, align="center", shade=h_fill, rgb=l_rgb)
    _style_cell_text(t3.cell(3, 2), "{{ lessee_rrn }}", size_pt=8.0, rgb=b_rgb)
    _style_cell_text(t3.cell(3, 3), "전화", bold=True, size_pt=7.6, align="center", shade=h_fill, rgb=l_rgb)
    _style_cell_text(t3.cell(3, 4), "{{ lessee_phone1 }}{% if lessee_phone2 %} / {{ lessee_phone2 }}{% endif %}", size_pt=7.8, rgb=b_rgb)
    _style_cell_text(t3.cell(3, 5), "성명: {{ lessee_name }} (인)", bold=True, size_pt=8.0, align="right", rgb=pal["date_rgb"])

    t3.cell(4, 0).merge(t3.cell(8, 0))
    _style_cell_text(t3.cell(4, 0), "개업\n공인\n중개사", bold=True, size_pt=7.5, align="center", shade=h_fill, rgb=l_rgb)

    broker_rows_meta = [
        (4, "사무소명칭", "{{ broker1_office }}", "{{ broker2_office }}", "left", b_rgb),
        (5, "사무소소재지", "{{ broker1_address }}", "{{ broker2_address }}", "left", b_rgb),
        (6, "대표", "{{ broker1_rep }} (인)", "{{ broker2_rep }} (인)", "right", pal["date_rgb"]),
        (7, "등록번호/전화", "{{ broker1_reg_no }}{% if broker1_phone %} / {{ broker1_phone }}{% endif %}", "{{ broker2_reg_no }}{% if broker2_phone %} / {{ broker2_phone }}{% endif %}", "left", b_rgb),
        (8, "소속공인중개사", "{{ broker1_associate }} (인)", "{{ broker2_associate }} (인)", "right", b_rgb),
    ]
    for r_idx, lbl, v1, v2, align, v_rgb in broker_rows_meta:
        t3.cell(r_idx, 4).merge(t3.cell(r_idx, 5))
        _style_cell_text(t3.cell(r_idx, 1), lbl, bold=True, size_pt=7.4, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t3.cell(r_idx, 2), v1, size_pt=7.7, align=align, rgb=v_rgb)
        _style_cell_text(t3.cell(r_idx, 3), lbl, bold=True, size_pt=7.4, align="center", shade=h_fill, rgb=l_rgb)
        _style_cell_text(t3.cell(r_idx, 4), v2, size_pt=7.7, align=align, rgb=v_rgb)

    # 조건부 2페이지 [별지] 특약사항
    doc.add_paragraph("{%p if is_annex %}")
    doc.add_page_break()

    p_ax_t = doc.add_paragraph()
    p_ax_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ax_t = p_ax_t.add_run("[별지] 특약사항")
    r_ax_t.bold = True
    r_ax_t.font.size = Pt(15)
    r_ax_t.font.color.rgb = RGBColor(*pal["title_rgb"])

    p_ax_meta = doc.add_paragraph("대상 소재지 : {{ address }}   |   계약일자 : {{ contract_date_kr }}")
    p_ax_meta.runs[0].bold = True
    p_ax_meta.runs[0].font.size = Pt(9.2)
    p_ax_meta.runs[0].font.color.rgb = RGBColor(*a_rgb)

    p_ax_body = doc.add_paragraph("{{ full_clauses_text }}")
    p_ax_body.runs[0].font.size = Pt(8.6)
    p_ax_body.runs[0].font.color.rgb = RGBColor(*b_rgb)

    p_ax_note = doc.add_paragraph(
        "위 특약사항은 본 계약서의 일부로서 효력을 가지며, 당사자는 이를 확인하고 간인 및 서명·날인한다."
    )
    p_ax_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ax_note.runs[0].font.size = Pt(8.0)

    p_ax_sign = doc.add_paragraph(
        f"{p1_lbl} : {{{{ lessor_name }}}} (인)     /     {p2_lbl} : {{{{ lessee_name }}}} (인)     /     개업공인중개사 : {{{{ broker1_rep }}}} (인)"
    )
    p_ax_sign.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ax_sign.runs[0].bold = True
    p_ax_sign.runs[0].font.size = Pt(8.5)
    p_ax_sign.runs[0].font.color.rgb = RGBColor(*pal["title_rgb"])

    doc.add_paragraph("{%p endif %}")

    doc.save(str(target))
    return target


def ensure_templates(force_rebuild_docx: bool = False) -> None:
    """6대 계약 유형별 컬러/흑백 정식 DOCX 템플릿(총 12종)을 보장한다."""
    DOCX_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    HWPX_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)

    for ctype in CONTRACT_TYPE_META:
        for mode in ("color", "bw"):
            tpl_path = get_docx_template_path(ctype, mode)
            if force_rebuild_docx or not tpl_path.exists():
                create_docx_template(ctype, mode, tpl_path)


# ---------------------------------------------------------------------------
# 사용자 업로드 커스텀 HWPX 템플릿 관리
# ---------------------------------------------------------------------------

def get_custom_hwpx_path(contract_type: str) -> Path:
    safe_type = "".join(ch for ch in contract_type if ch.isalnum() or ch == "_")
    return HWPX_TEMPLATE_DIR / f"custom_{safe_type}.hwpx"


def has_custom_hwpx_template(contract_type: str) -> bool:
    return get_custom_hwpx_path(contract_type).exists()


def save_custom_hwpx_template(contract_type: str, data: bytes) -> tuple[bool, str, list[str]]:
    HWPX_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    ok, errors = validate_hwpx_bytes(data)
    if not ok:
        return False, "HWPX 규격 검증 실패: " + "; ".join(errors[:3]), []

    target = get_custom_hwpx_path(contract_type)
    target.write_bytes(data)
    placeholders = extract_hwpx_placeholders(data)
    return True, f"사용자 한글 템플릿이 등록되었습니다. ({target.name})", placeholders


def reset_custom_hwpx_template(contract_type: str) -> None:
    custom = get_custom_hwpx_path(contract_type)
    if custom.exists():
        custom.unlink()


def render_html(context: dict[str, Any], include_print_toolbar: bool = False) -> str:
    """6대 계약 유형별 정식 HTML/CSS A4 계약서를 컬러/흑백 모드로 렌더링한다."""
    contract_type = str(context.get("contract_type", "housing_jeonse"))
    meta = get_type_meta(contract_type)
    template = _jinja_env.get_template(meta["template_name"])
    render_ctx = dict(context)
    render_ctx["include_print_toolbar"] = include_print_toolbar
    return template.render(**render_ctx)


def generate_pdf(context: dict[str, Any]) -> tuple[bytes | None, str]:
    html_str = render_html(context, include_print_toolbar=False)
    try:
        from weasyprint import HTML  # type: ignore

        pdf_bytes = HTML(string=html_str, base_url=str(BASE_DIR)).write_pdf()
        return pdf_bytes, "WeasyPrint PDF 생성 완료"
    except Exception as exc:
        return (
            None,
            f"서버 PDF 엔진 미구성({type(exc).__name__}): "
            "PDF 인쇄용 HTML을 내려받아 브라우저 인쇄(Ctrl+P → PDF로 저장)를 이용해 주세요.",
        )


def generate_docx(context: dict[str, Any]) -> bytes:
    """선택된 계약 유형과 컬러/흑백 모드에 맞는 정식 .docx를 생성한다."""
    from docxtpl import DocxTemplate

    contract_type = str(context.get("contract_type", "housing_jeonse"))
    color_mode = "bw" if context.get("color_mode") == "bw" else "color"
    tpl_path = get_docx_template_path(contract_type, color_mode)
    if not tpl_path.exists():
        create_docx_template(contract_type, color_mode, tpl_path)

    tpl = DocxTemplate(str(tpl_path))
    render_ctx = dict(context)
    render_ctx["draft_banner"] = (
        f"[ {context.get('watermark_text', '사전 검토용 초안')} ]"
        if context.get("draft_mode")
        else ""
    )

    tpl.render(render_ctx)
    buf = io.BytesIO()
    tpl.save(buf)
    return buf.getvalue()


def _build_hwpx_mapping(context: dict[str, Any]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for key, val in context.items():
        if isinstance(val, (str, int, float)):
            mapping[key] = str(val)
        elif val is None:
            mapping[key] = ""

    mapping["draft_banner"] = (
        f"[{context.get('watermark_text', '사전 검토용 초안')}] "
        if context.get("draft_mode")
        else ""
    )

    kr_aliases = {
        "계약서제목": mapping.get("form_title", ""),
        "초안배너": mapping["draft_banner"],
        "소재지": mapping.get("address", ""),
        "지목": mapping.get("land_category", ""),
        "대지권비율": mapping.get("land_share_ratio", ""),
        "토지면적": mapping.get("land_area", ""),
        "토지평수": mapping.get("land_area_pyeong", ""),
        "평당단가": mapping.get("price_per_pyeong", ""),
        "용도지역": mapping.get("land_zoning", ""),
        "토지현황": mapping.get("land_status", ""),
        "추가필지": mapping.get("extra_parcels", ""),
        "건물구조": mapping.get("building_structure", ""),
        "건물용도": mapping.get("building_use", ""),
        "건물면적": mapping.get("building_area", ""),
        "임대할부분": mapping.get("lease_part", ""),
        "임대면적": mapping.get("lease_area", ""),
        "시설명": mapping.get("facility_name", ""),
        "정원": mapping.get("facility_capacity", ""),
        "기관기호": mapping.get("facility_reg_no", ""),
        "상호": mapping.get("store_name", ""),
        "업종": mapping.get("business_type", ""),
        "신규임대차조건": mapping.get("new_lease_summary", ""),
        "보증금_한글": mapping.get("deposit_kr", ""),
        "보증금_숫자": mapping.get("deposit_won", ""),
        "보증금_콤마": mapping.get("deposit_comma", ""),
        "계약금_한글": mapping.get("down_payment_kr", ""),
        "계약금_숫자": mapping.get("down_payment_comma", ""),
        "영수자": mapping.get("down_receipt_person", ""),
        "융자승계_한글": mapping.get("loan_assumption_kr", ""),
        "임대보증금승계_한글": mapping.get("lease_deposit_assumption_kr", ""),
        "중도금_한글": mapping.get("interim_payment_kr", ""),
        "잔금_한글": mapping.get("balance_kr", ""),
        "잔금일_년": mapping.get("balance_year", ""),
        "잔금일_월": mapping.get("balance_month", ""),
        "잔금일_일": mapping.get("balance_day", ""),
        "월차임_한글": mapping.get("monthly_rent_kr", ""),
        "관리비_한글": mapping.get("management_fee_kr", ""),
        "부가세구분": mapping.get("vat_type", ""),
        "유보금_한글": mapping.get("escrow_amount_kr", ""),
        "유보기간": mapping.get("escrow_period", ""),
        "인도일": mapping.get("handover_date_kr", ""),
        "만료일": mapping.get("lease_end_date_kr", ""),
        "존속기간": mapping.get("lease_duration", ""),
        "계약일": mapping.get("contract_date_kr", ""),
        "특약사항": mapping.get("body_clauses_text", ""),
        "특약사항_전문": mapping.get("full_clauses_text", ""),
        "임대인_주소": mapping.get("lessor_address", ""),
        "임대인_주민번호": mapping.get("lessor_rrn", ""),
        "임대인_전화1": mapping.get("lessor_phone1", ""),
        "임대인_전화2": mapping.get("lessor_phone2", ""),
        "임대인_성명": mapping.get("lessor_name", ""),
        "임차인_주소": mapping.get("lessee_address", ""),
        "임차인_주민번호": mapping.get("lessee_rrn", ""),
        "임차인_전화1": mapping.get("lessee_phone1", ""),
        "임차인_전화2": mapping.get("lessee_phone2", ""),
        "임차인_성명": mapping.get("lessee_name", ""),
        "중개사1_상호": mapping.get("broker1_office", ""),
        "중개사1_소재지": mapping.get("broker1_address", ""),
        "중개사1_대표": mapping.get("broker1_rep", ""),
        "중개사1_등록번호": mapping.get("broker1_reg_no", ""),
        "중개사1_전화": mapping.get("broker1_phone", ""),
        "중개사1_소속": mapping.get("broker1_associate", ""),
        "중개사2_상호": mapping.get("broker2_office", ""),
        "중개사2_소재지": mapping.get("broker2_address", ""),
        "중개사2_대표": mapping.get("broker2_rep", ""),
        "중개사2_등록번호": mapping.get("broker2_reg_no", ""),
        "중개사2_전화": mapping.get("broker2_phone", ""),
        "중개사2_소속": mapping.get("broker2_associate", ""),
    }
    mapping.update(kr_aliases)
    return mapping


def generate_hwpx(context: dict[str, Any]) -> bytes:
    """사용자가 업로드한 커스텀 .hwpx 템플릿이 있을 때만 치환본을 생성한다."""
    contract_type = str(context.get("contract_type", "housing_jeonse"))
    custom_path = get_custom_hwpx_path(contract_type)
    if not custom_path.exists():
        raise FileNotFoundError(
            f"'{contract_type}' 유형에 등록된 사용자 한글(.hwpx) 템플릿이 없습니다."
        )

    mapping = _build_hwpx_mapping(context)
    return render_hwpx(
        template_path=custom_path,
        mapping=mapping,
        annex_enabled=bool(context.get("is_annex", False)),
    )
