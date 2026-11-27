"""한컴 표준 OWPML(KS X 6101) 호환 HWPX 생성·치환·검증 엔진.

- 기본 템플릿 생성: `python-hwpx`의 공식 한컴 표준 스켈레톤(mimetype 비압축, version.xml,
  settings.xml, META-INF/container.xml·container.rdf·manifest.xml, Contents/content.hpf·header.xml·section0.xml)
  기반으로 아래아한글에서 규격 오류 없이 열리는 `.hwpx` 생성.
- 치환 엔진: `lxml.etree` 기반으로 기존 네임스페이스·속성을 100% 보존하면서
  1) 분절된 <hp:run><hp:t> 태그 병합
  2) 특수문자(&, <, >) 안전 이스케이프 및 XML 제어문자 제거
  3) 줄바꿈(\\n) -> 개별 <hp:p> 문단 분리 및 구 레이아웃 캐시(<hp:linesegarray>) 정리
  4) 별지 구간([[ANNEX_START]]~[[ANNEX_END]]) 조건부 제거/페이지 나누기
- 사용자 업로드 템플릿: 아래아한글에서 직접 만든 `.hwpx` 파일의 규격 검증 및 `{{태그}}` 자동 인식 지원.
"""

from __future__ import annotations

import copy
import html
import io
import re
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree as ET

HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HP = f"{{{HP_NS}}}"

PLACEHOLDER_RE = re.compile(r"\{\{([^{}]+)\}\}")
ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

ANNEX_START_MARKER = "[[ANNEX_START]]"
ANNEX_END_MARKER = "[[ANNEX_END]]"


def sanitize_xml_text(value: Any) -> str:
    """XML 1.0 허용 불가 제어 문자를 제거한다."""
    if value is None:
        return ""
    return ILLEGAL_XML_RE.sub("", str(value))


def escape_xml_string(value: Any) -> str:
    """문자열 기반 XML 조합 시 &, <, > 특수문자를 html.escape()로 이스케이프한다."""
    return html.escape(sanitize_xml_text(value), quote=False)


def _is_simple_text_run(el: ET._Element) -> bool:
    """단일 <hp:t>만 포함하는 단순 텍스트 <hp:run>인지 판별한다."""
    if el.tag != f"{HP}run":
        return False
    children = list(el)
    return len(children) == 1 and children[0].tag == f"{HP}t" and len(list(children[0])) == 0


def _get_run_text(run_el: ET._Element) -> str:
    t_node = run_el.find(f"{HP}t")
    return t_node.text if (t_node is not None and t_node.text) else ""


def merge_split_runs_in_tree(root: ET._Element) -> None:
    """동일 문단(<hp:p>) 내 쪼개진 <hp:run><hp:t> 태그를 병합해 {{변수}} 분절을 복구한다.

    아래아한글에서 저장한 `.hwpx` 파일은 `{{보증금_한글}}`이 여러 `<hp:run>`으로
    쪼개지는 경우가 빈번하므로, 연속된 동일 서식 단순 텍스트 run을 하나로 합친다.
    """
    for p in root.iter(f"{HP}p"):
        children = list(p)
        i = 0
        while i < len(children):
            el = children[i]
            if _is_simple_text_run(el):
                char_pr = el.get("charPrIDRef", "0")
                buf_texts = [_get_run_text(el)]
                j = i + 1
                while j < len(children):
                    nxt = children[j]
                    if _is_simple_text_run(nxt) and nxt.get("charPrIDRef", "0") == char_pr:
                        buf_texts.append(_get_run_text(nxt))
                        j += 1
                    else:
                        break
                if j > i + 1:
                    t_node = el.find(f"{HP}t")
                    if t_node is not None:
                        t_node.text = "".join(buf_texts)
                    for rm in children[i + 1 : j]:
                        p.remove(rm)
                i = j
            else:
                i += 1


def _clear_linesegarray(p_el: ET._Element) -> None:
    """텍스트가 변경된 문단의 캐시된 <hp:linesegarray>를 제거해 한글이 줄바꿈을 재계산하게 한다."""
    for lsa in list(p_el.findall(f"./{HP}linesegarray")):
        p_el.remove(lsa)


def _paragraph_plain_text(p_el: ET._Element) -> str:
    return "".join(t.text or "" for t in p_el.findall(f"./{HP}run/{HP}t")).strip()


def _replace_in_text(text: str, mapping: dict[str, str]) -> tuple[str, bool]:
    changed = False

    def _sub(m: re.Match[str]) -> str:
        nonlocal changed
        key = m.group(1).strip()
        if key in mapping:
            changed = True
            return sanitize_xml_text(mapping[key])
        return m.group(0)

    out = PLACEHOLDER_RE.sub(_sub, text)
    return out, changed


def _expand_paragraph_if_multiline(
    parent: ET._Element,
    p_el: ET._Element,
    mapping: dict[str, str],
) -> None:
    """문단 내 치환 값에 줄바꿈(\\n)이 포함되면 개별 <hp:p> 문단으로 안전하게 분리한다."""
    text_nodes = list(p_el.findall(f"./{HP}run/{HP}t"))
    if not text_nodes:
        return

    replaced_values: list[str] = []
    any_changed = False
    for t_node in text_nodes:
        new_val, changed = _replace_in_text(t_node.text or "", mapping)
        replaced_values.append(new_val)
        if changed:
            any_changed = True

    if not any_changed:
        return

    has_newline = any("\n" in val for val in replaced_values)
    if not has_newline:
        for t_node, val in zip(text_nodes, replaced_values):
            t_node.text = val
        _clear_linesegarray(p_el)
        return

    # 줄바꿈이 있는 경우: 원본 <hp:p>를 복제해 줄별 독립 <hp:p> 문단으로 전개
    # 단, 첫 문단에 <hp:secPr> 등 구역 설정이 붙어있으면 첫 문단에만 유지
    full_replaced = "".join(replaced_values)
    lines = full_replaced.split("\n")

    try:
        idx = list(parent).index(p_el)
    except ValueError:
        for t_node, val in zip(text_nodes, replaced_values):
            t_node.text = val.replace("\n", " ")
        _clear_linesegarray(p_el)
        return

    first_run = p_el.find(f"./{HP}run")
    char_pr = first_run.get("charPrIDRef", "0") if first_run is not None else "0"

    new_paragraphs: list[ET._Element] = []
    for line_idx, line in enumerate(lines):
        new_p = copy.deepcopy(p_el)
        _clear_linesegarray(new_p)
        if line_idx > 0:
            new_p.set("pageBreak", "0")
            # 후속 문단에서는 secPr/ctrl 등 구역 정의 run 제거
            for run in list(new_p.findall(f"./{HP}run")):
                new_p.remove(run)
        else:
            for run in list(new_p.findall(f"./{HP}run")):
                if _is_simple_text_run(run):
                    new_p.remove(run)
        new_run = ET.SubElement(new_p, f"{HP}run", {"charPrIDRef": char_pr})
        new_t = ET.SubElement(new_run, f"{HP}t")
        new_t.text = line
        new_paragraphs.append(new_p)

    parent.remove(p_el)
    for offset, np in enumerate(new_paragraphs):
        parent.insert(idx + offset, np)


def _handle_annex_block(container: ET._Element, annex_enabled: bool) -> None:
    """[[ANNEX_START]] ~ [[ANNEX_END]] 마커 구간을 처리한다 (숫자형 hp:p/@id 규격 유지)."""
    children = list(container)
    in_annex = False
    first_annex_p_seen = False

    for child in children:
        if child.tag != f"{HP}p":
            continue
        txt = _paragraph_plain_text(child)
        if txt == ANNEX_START_MARKER:
            in_annex = True
            container.remove(child)
            continue
        if txt == ANNEX_END_MARKER:
            in_annex = False
            container.remove(child)
            continue
        if in_annex:
            if not annex_enabled:
                container.remove(child)
            elif not first_annex_p_seen:
                child.set("pageBreak", "1")
                first_annex_p_seen = True


def _walk_and_replace(container: ET._Element, mapping: dict[str, str], annex_enabled: bool) -> None:
    """재귀적으로 모든 컨테이너(<hs:sec>, <hp:subList>)의 <hp:p>를 치환한다."""
    _handle_annex_block(container, annex_enabled=annex_enabled)

    for child in list(container):
        if child.tag == f"{HP}p":
            # 표(<hp:tbl>) 내부 셀(<hp:subList>)을 먼저 재귀 처리
            for sub_list in child.findall(f".//{HP}subList"):
                _walk_and_replace(sub_list, mapping, annex_enabled)
            _expand_paragraph_if_multiline(container, child, mapping)
        else:
            _walk_and_replace(child, mapping, annex_enabled)


def process_section_xml(
    xml_bytes: bytes,
    mapping: dict[str, str],
    annex_enabled: bool = False,
) -> bytes:
    """section0.xml 바이트를 파싱해 태그 병합·이스케이프·문단 분리를 수행한다."""
    safe_mapping = {k: sanitize_xml_text(v) for k, v in mapping.items()}

    parser = ET.XMLParser(remove_blank_text=False, resolve_entities=False)
    root = ET.fromstring(xml_bytes, parser=parser)
    merge_split_runs_in_tree(root)
    _walk_and_replace(root, safe_mapping, annex_enabled=annex_enabled)

    return ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def render_hwpx(
    template_path: Path,
    mapping: dict[str, str],
    annex_enabled: bool = False,
) -> bytes:
    """HWPX 템플릿 파일을 읽어 컨텍스트를 치환한 .hwpx 바이트를 반환한다."""
    out_buf = io.BytesIO()
    with zipfile.ZipFile(template_path, "r") as zin:
        entries = zin.namelist()
        ordered = ["mimetype"] + [e for e in entries if e != "mimetype"]

        with zipfile.ZipFile(out_buf, "w") as zout:
            for name in ordered:
                if name not in entries:
                    continue
                data = zin.read(name)
                if name == "mimetype":
                    zout.writestr(name, data, compress_type=zipfile.ZIP_STORED)
                elif name.startswith("Contents/section") and name.endswith(".xml"):
                    processed = process_section_xml(
                        data,
                        mapping=mapping,
                        annex_enabled=annex_enabled,
                    )
                    zout.writestr(name, processed, compress_type=zipfile.ZIP_DEFLATED)
                else:
                    zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)

    return out_buf.getvalue()


def validate_hwpx_bytes(data: bytes) -> tuple[bool, list[str]]:
    """HWPX 바이트의 ZIP 패키징(mimetype 비압축 등) 및 한컴 OWPML XSD 규격을 검증한다."""
    errors: list[str] = []
    try:
        with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
            infos = zf.infolist()
            if not infos:
                return False, ["빈 ZIP 파일입니다."]
            if infos[0].filename != "mimetype":
                errors.append(f"첫 번째 ZIP 엔트리가 'mimetype'이 아닙니다: {infos[0].filename}")
            elif infos[0].compress_type != zipfile.ZIP_STORED:
                errors.append("mimetype 엔트리가 비압축(ZIP_STORED)으로 저장되지 않았습니다.")
            else:
                mt = zf.read("mimetype").decode("utf-8", errors="ignore").strip()
                if mt != "application/hwp+zip":
                    errors.append(f"mimetype 내용이 표준과 다릅니다: {mt}")

            names = set(zf.namelist())
            for req in ("version.xml", "META-INF/container.xml", "Contents/header.xml", "Contents/section0.xml"):
                if req not in names:
                    errors.append(f"필수 HWPX 구성 파일 누락: {req}")
    except zipfile.BadZipFile:
        return False, ["유효한 HWPX(ZIP) 파일이 아닙니다."]

    if errors:
        return False, errors

    # python-hwpx 공식 OWPML XSD 스키마 검증
    try:
        from hwpx import HwpxDocument

        doc = HwpxDocument.open(io.BytesIO(data))
        report = doc.validate()
        if getattr(report, "issues", ()):
            for issue in report.issues:
                errors.append(str(issue))
    except Exception as exc:
        errors.append(f"OWPML 스키마 검증 실패: {exc}")

    return len(errors) == 0, errors


def extract_hwpx_placeholders(data: bytes) -> list[str]:
    """업로드된 .hwpx 파일 내부(분절 run 병합 후)에서 인식된 {{태그}} 목록을 추출한다."""
    found: list[str] = []
    seen: set[str] = set()
    try:
        with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
            for name in zf.namelist():
                if name.startswith("Contents/section") and name.endswith(".xml"):
                    parser = ET.XMLParser(remove_blank_text=False, resolve_entities=False)
                    root = ET.fromstring(zf.read(name), parser=parser)
                    merge_split_runs_in_tree(root)
                    for t_node in root.iter(f"{HP}t"):
                        for m in PLACEHOLDER_RE.finditer(t_node.text or ""):
                            key = m.group(1).strip()
                            if key not in seen:
                                seen.add(key)
                                found.append(key)
    except Exception:
        pass
    return found


# ---------------------------------------------------------------------------
# 한컴 공식 XSD 규격 준수 초기 샘플 HWPX 템플릿 생성기 (python-hwpx 기반)
# ---------------------------------------------------------------------------

def create_default_hwpx_template(output_path: Path) -> None:
    """python-hwpx 공식 스켈레톤을 사용해 아래아한글에서 정상 열리는 A4 전세 계약서 템플릿을 생성한다."""
    from hwpx import HwpxDocument

    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = HwpxDocument.new()
    doc.page.setup(
        paper_size="A4",
        margin_left_mm=10.0,
        margin_right_mm=10.0,
        margin_top_mm=10.0,
        margin_bottom_mm=10.0,
        header_margin_mm=4.0,
        footer_margin_mm=4.0,
    )

    # 서식 정의 (문자/문단 스타일)
    c_title = doc.styles.ensure_run(bold=True, size=15, font="맑은 고딕")
    c_h2 = doc.styles.ensure_run(bold=True, size=8.5, font="맑은 고딕")
    c_body = doc.styles.ensure_run(bold=False, size=8.0, font="맑은 고딕")
    c_small = doc.styles.ensure_run(bold=False, size=7.2, font="맑은 고딕")
    c_bold_sm = doc.styles.ensure_run(bold=True, size=7.5, font="맑은 고딕")

    p_left = doc.styles.ensure_paragraph_format(alignment="LEFT", line_spacing=125)
    p_center = doc.styles.ensure_paragraph_format(alignment="CENTER", line_spacing=125)

    # 기본 빈 첫 문단에 제목 배치 (빈 여백 줄 방지)
    first_p = doc.paragraphs[0]
    first_p.text = "{{draft_banner}}{{form_title}}"
    first_p.para_pr_id_ref = p_center
    first_p.char_pr_id_ref = c_title

    doc.add_paragraph(
        "아래 부동산에 대하여 임대인과 임차인은 합의하에 다음과 같이 임대차계약을 체결한다.",
        para_pr_id_ref=p_left,
        char_pr_id_ref=c_small,
    )

    # 1. 부동산의 표시
    doc.add_paragraph("1. 부동산의 표시", para_pr_id_ref=p_left, char_pr_id_ref=c_h2)

    W = 53000  # A4 190mm 가용 폭
    t1 = doc.add_table(4, 7, width=W)
    t1.set_column_widths([11, 7, 22, 9, 25, 7, 19])
    t1.set_borders(line_type="SOLID", border_width_mm=0.12, color="#000000", inner=True)

    t1.merge_cells(0, 1, 0, 6)
    t1.merge_cells(3, 1, 3, 4)

    t1_cells = [
        (0, 0, "소 재 지", True), (0, 1, "{{address}}", False),
        (1, 0, "토 지", True), (1, 1, "지목", True), (1, 2, "{{land_category}}", False),
        (1, 3, "대지권비율", True), (1, 4, "{{land_share_ratio}}", False),
        (1, 5, "면적", True), (1, 6, "{{land_area}} ㎡", False),
        (2, 0, "건 물", True), (2, 1, "구조", True), (2, 2, "{{building_structure}}", False),
        (2, 3, "용도", True), (2, 4, "{{building_use}}", False),
        (2, 5, "면적", True), (2, 6, "{{building_area}} ㎡", False),
        (3, 0, "임대할부분", True), (3, 1, "{{lease_part}}", False),
        (3, 5, "면적", True), (3, 6, "{{lease_area}} ㎡", False),
    ]
    for r, c, txt, shaded in t1_cells:
        t1.set_cell_text(r, c, txt)
        if shaded:
            t1.set_cell_shading(r, c, "#EFEFEF")

    # 2. 계약내용
    doc.add_paragraph(
        "2. 계약내용  제1조 [보증금 및 지불시기] 위 부동산의 임대차에 관하여 보증금과 지불시기를 다음과 같이 약정한다.",
        para_pr_id_ref=p_left,
        char_pr_id_ref=c_bold_sm,
    )

    t2 = doc.add_table(4, 5, width=W)
    t2.set_column_widths([11, 31, 33, 9, 16])
    t2.set_borders(line_type="SOLID", border_width_mm=0.12, color="#000000", inner=True)

    t2.merge_cells(0, 2, 0, 4)
    t2.merge_cells(2, 2, 2, 4)
    t2.merge_cells(3, 2, 3, 4)

    t2_cells = [
        (0, 0, "보 증 금", True), (0, 1, "ㅡ금 {{deposit_kr}}", False),
        (0, 2, "원정 ({{deposit_won}})", False),
        (1, 0, "계 약 금", True), (1, 1, "ㅡ금 {{down_payment_kr}}", False),
        (1, 2, "원정은 계약시에 지불하고 영수함.", False),
        (1, 3, "영수자", True), (1, 4, "{{down_receipt_person}} (인)", False),
        (2, 0, "중 도 금", True), (2, 1, "ㅡ금 {{interim_payment_kr}}", False),
        (2, 2, "원정은 {{interim_year}} 년 {{interim_month}} 월 {{interim_day}} 일에 지불하며,", False),
        (3, 0, "잔  금", True), (3, 1, "ㅡ금 {{balance_kr}}", False),
        (3, 2, "원정은 {{balance_year}} 년 {{balance_month}} 월 {{balance_day}} 일에 지불한다.", False),
    ]
    for r, c, txt, shaded in t2_cells:
        t2.set_cell_text(r, c, txt)
        if shaded:
            t2.set_cell_shading(r, c, "#EFEFEF")

    articles = [
        "위 제1항의 보증금은 개업공인중개사의 입회하에 지불하기로 한다.",
        "제2조 [존속기간] 임대인은 위 부동산을 임대차 목적대로 사용·수익할 수 있는 상태로 {{handover_date_kr}} 까지 임차인에게 인도하며, 임대차기간은 인도일로부터 {{lease_end_date_kr}} 까지로 한다.({{lease_duration}})",
        "제3조 [용도변경 및 전대 등] 임차인은 임대인의 동의 없이는 위 부동산의 용도나 구조 등을 변경하거나 전대·임차권 양도 또는 담보제공을 하지 못하며 임대차 목적 이외의 용도에 사용할 수 없다.",
        "제4조 [계약의 종료] 임대차 계약이 종료된 경우 임차인은 위 부동산을 원상으로 회복하여 임대인에게 반환한다. 이러한 경우 임대인은 보증금을 임차인에게 반환하고, 연체임대료 또는 손해배상 금액이 있을 때는 이들을 제하고 그 잔액을 반환한다.",
        "제5조 [계약의 해제] 임차인이 임대인에게 중도금(중도금이 없을 때는 잔금)을 지불하기 전까지 임대인은 계약금의 배액을 상환하고 임차인은 계약금을 포기하고 이 계약을 해제할 수 있다.",
        "제6조 [채무불이행과 손해배상] 임대인 또는 임차인이 본 계약상의 내용에 대하여 불이행이 있을 경우 그 상대방은 불이행한 자에 대하여 서면으로 최고하고 계약을 해제할 수 있다. 그리고 계약 당사자는 계약해제에 따른 손해배상을 각각 상대방에게 청구할 수 있으며, 손해배상에 대하여 별도의 약정이 없는 한 계약금을 손해배상의 기준으로 본다.",
        "제7조 [중개보수] 개업공인중개사는 임대인 또는 임차인의 본 계약 불이행에 대하여 책임을 지지 않는다. 또한 중개보수는 본 계약 체결에 따라 계약 당사자 쌍방이 각각 지불하며, 개업공인중개사의 고의나 과실 없이 본 계약이 무효·취소 또는 해제되어도 중개보수는 지급한다. 공동 중개인 경우에 임대인과 임차인은 자신이 중개 의뢰한 개업공인중개사에게 각각 중개보수를 지급한다.",
        "제8조 [확인설명서 등 교부] 개업공인중개사는 중개대상물 확인·설명서를 작성하고 업무보증관계증서(공제증서 등) 사본을 첨부하여 거래계약서와 동시에 거래당사자 쌍방에게 교부한다.",
    ]
    for art in articles:
        doc.add_paragraph(art, para_pr_id_ref=p_left, char_pr_id_ref=c_small)

    # 특약사항
    doc.add_paragraph("특약사항", para_pr_id_ref=p_left, char_pr_id_ref=c_h2)
    doc.add_paragraph("{{body_clauses_text}}", para_pr_id_ref=p_left, char_pr_id_ref=c_body)

    doc.add_paragraph(
        "본 계약에 대하여 임대인과 임차인은 이의 없음을 확인하고 각각 서명 또는 날인 후 임대인, 임차인, 개업공인중개사가 각 1통씩 보관한다.",
        para_pr_id_ref=p_center,
        char_pr_id_ref=c_small,
    )
    doc.add_paragraph("{{contract_date_kr}}", para_pr_id_ref=p_center, char_pr_id_ref=c_h2)

    # 당사자 및 개업공인중개사 표 (9행 6열)
    t3 = doc.add_table(9, 6, width=W)
    t3.set_column_widths([6, 13, 31, 13, 19, 18])
    t3.set_borders(line_type="SOLID", border_width_mm=0.12, color="#000000", inner=True)

    t3.merge_cells(0, 0, 1, 0)
    t3.merge_cells(0, 2, 0, 5)
    t3.merge_cells(2, 0, 3, 0)
    t3.merge_cells(2, 2, 2, 5)
    t3.merge_cells(4, 0, 8, 0)
    for r_idx in range(4, 9):
        t3.merge_cells(r_idx, 4, r_idx, 5)

    t3_cells = [
        (0, 0, "임대인", True), (0, 1, "주 소", True), (0, 2, "{{lessor_address}}", False),
        (1, 1, "주민등록번호", True), (1, 2, "{{lessor_rrn}}", False),
        (1, 3, "전화", True), (1, 4, "{{lessor_phone1}} / {{lessor_phone2}}", False),
        (1, 5, "성명: {{lessor_name}} (인)", False),
        (2, 0, "임차인", True), (2, 1, "주 소", True), (2, 2, "{{lessee_address}}", False),
        (3, 1, "주민등록번호", True), (3, 2, "{{lessee_rrn}}", False),
        (3, 3, "전화", True), (3, 4, "{{lessee_phone1}} / {{lessee_phone2}}", False),
        (3, 5, "성명: {{lessee_name}} (인)", False),
        (4, 0, "개업공인중개사", True),
        (4, 1, "사무소명칭", True), (4, 2, "{{broker1_office}}", False),
        (4, 3, "사무소명칭", True), (4, 4, "{{broker2_office}}", False),
        (5, 1, "사무소소재지", True), (5, 2, "{{broker1_address}}", False),
        (5, 3, "사무소소재지", True), (5, 4, "{{broker2_address}}", False),
        (6, 1, "대 표", True), (6, 2, "{{broker1_rep}} (인)", False),
        (6, 3, "대 표", True), (6, 4, "{{broker2_rep}} (인)", False),
        (7, 1, "등록번호/전화", True), (7, 2, "{{broker1_reg_no}} / {{broker1_phone}}", False),
        (7, 3, "등록번호/전화", True), (7, 4, "{{broker2_reg_no}} / {{broker2_phone}}", False),
        (8, 1, "소속공인중개사", True), (8, 2, "{{broker1_associate}} (인)", False),
        (8, 3, "소속공인중개사", True), (8, 4, "{{broker2_associate}} (인)", False),
    ]
    for r, c, txt, shaded in t3_cells:
        t3.set_cell_text(r, c, txt)
        if shaded:
            t3.set_cell_shading(r, c, "#EFEFEF")

    # 조건부 2페이지 [별지] 특약사항 (마커 문단으로 제어, hp:p/@id는 숫자 규격 유지)
    doc.add_paragraph(ANNEX_START_MARKER, para_pr_id_ref=p_left, char_pr_id_ref=c_small)
    doc.add_paragraph("[별 지] 특 약 사 항", para_pr_id_ref=p_center, char_pr_id_ref=c_title)
    doc.add_paragraph(
        "부동산 소재지 : {{address}}   |   계약일자 : {{contract_date_kr}}",
        para_pr_id_ref=p_left,
        char_pr_id_ref=c_h2,
    )
    doc.add_paragraph("{{full_clauses_text}}", para_pr_id_ref=p_left, char_pr_id_ref=c_body)
    doc.add_paragraph(
        "위 특약사항은 본 계약서의 일부로서 효력을 가지며, 당사자는 이를 확인하고 간인 및 서명·날인한다.",
        para_pr_id_ref=p_center,
        char_pr_id_ref=c_small,
    )
    doc.add_paragraph(
        "임대인 : {{lessor_name}} (인)   /   임차인 : {{lessee_name}} (인)   /   개업공인중개사 : {{broker1_rep}} (인)",
        para_pr_id_ref=p_center,
        char_pr_id_ref=c_h2,
    )
    doc.add_paragraph(ANNEX_END_MARKER, para_pr_id_ref=p_left, char_pr_id_ref=c_small)

    # 최종 OWPML XSD 스키마 검증 후 저장
    report = doc.validate()
    if getattr(report, "issues", ()):
        raise RuntimeError(f"HWPX 기본 템플릿 XSD 검증 실패: {report.issues}")

    doc.save(str(output_path))
