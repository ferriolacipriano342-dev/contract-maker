"""부동산·요양원 통합 계약서 작성 및 출력 웹 앱 (3단계: 실무 완성형).

- 6대 정식 A4 서식(전세·매매·월세·상가·요양원·권리금) + 주요사항 큰 글씨 강조 + 컬러/흑백 출력 모드
- 로컬 SQLite DB(data/contracts.db) 임시저장 / 최종 누적 저장 / 사이드바 목록·검색·복사·삭제
- 증빙서류 카메라 촬영 & 갤러리 업로드 보관(data/attachments/{계약ID}/) + Gemini AI 자동 채우기
"""

from __future__ import annotations

import hmac
import json
from datetime import date, timedelta
from pathlib import Path
from typing import Any

import streamlit as st
import streamlit.components.v1 as components

from utils.ai_extractor import (
    delete_gemini_api_key_from_secrets,
    extract_contract_fields_from_documents,
    generate_legal_clauses_from_text,
    get_gemini_api_key,
    mask_api_key,
    save_gemini_api_key_to_secrets,
)
from utils.doc_generator import (
    HWPX_PLACEHOLDER_GUIDE,
    ensure_templates,
    generate_docx,
    generate_hwpx,
    generate_pdf,
    has_custom_hwpx_template,
    render_html,
    reset_custom_hwpx_template,
    save_custom_hwpx_template,
)
from utils.formatter import (
    CONTRACT_TYPE_META,
    MAX_INLINE_CLAUSE_COUNT,
    MAX_INLINE_CLAUSE_LINES,
    build_contract_context,
    calculate_balance,
    calculate_lease_duration,
    format_date_kr,
    format_price_per_pyeong,
    format_pyeong,
    format_won,
    get_type_meta,
    number_to_korean,
    resolve_smart_clause_placeholders,
    substitute_clause_variables,
)
from utils.storage import (
    ATTACHMENT_CATEGORIES,
    delete_attachment,
    delete_contract,
    ensure_contract_record_exists,
    generate_contract_id,
    init_db,
    list_attachments,
    list_contracts,
    load_contract,
    read_attachment_bytes,
    save_attachment,
    save_contract,
)

BASE_DIR = Path(__file__).resolve().parent
CLAUSES_PATH = BASE_DIR / "data" / "clauses.json"

# 계약 유형별 시그니처 카테고리 테마 컬러 (A4 디지털 서식과 UI 선택 카드 100% 통일)
CONTRACT_TYPE_THEMES: dict[str, dict[str, str]] = {
    "nursing_home_transfer": {
        "category": "🏥 요양시설",
        "primary": "#1E3A5F",
        "hover": "#152B47",
        "soft_bg": "#F1F5F9",
        "badge_bg": "#F5ECD7",
    },
    "apartment_sale": {
        "category": "🏢 주거 · 아파트",
        "primary": "#1D4ED8",
        "hover": "#1E40AF",
        "soft_bg": "#EFF6FF",
        "badge_bg": "#DBEAFE",
    },
    "real_estate_sale": {
        "category": "🏠 주거 · 주택",
        "primary": "#1D4ED8",
        "hover": "#1E40AF",
        "soft_bg": "#EFF6FF",
        "badge_bg": "#DBEAFE",
    },
    "housing_jeonse": {
        "category": "🏠 주거 · 전세",
        "primary": "#1D4ED8",
        "hover": "#1E40AF",
        "soft_bg": "#EFF6FF",
        "badge_bg": "#DBEAFE",
    },
    "housing_wolse": {
        "category": "🏠 주거 · 월세",
        "primary": "#1D4ED8",
        "hover": "#1E40AF",
        "soft_bg": "#EFF6FF",
        "badge_bg": "#DBEAFE",
    },
    "commercial_sale": {
        "category": "🏪 상가 · 매매",
        "primary": "#047857",
        "hover": "#065F46",
        "soft_bg": "#ECFDF5",
        "badge_bg": "#D1FAE5",
    },
    "commercial_lease": {
        "category": "🏪 상가 · 임대차",
        "primary": "#047857",
        "hover": "#065F46",
        "soft_bg": "#ECFDF5",
        "badge_bg": "#D1FAE5",
    },
    "premium_rights": {
        "category": "🏪 상가 · 권리금",
        "primary": "#047857",
        "hover": "#065F46",
        "soft_bg": "#ECFDF5",
        "badge_bg": "#D1FAE5",
    },
    "farmland_sale": {
        "category": "🌾 토지 · 농지",
        "primary": "#15803D",
        "hover": "#166534",
        "soft_bg": "#F0FDF4",
        "badge_bg": "#DCFCE7",
    },
    "land_sale": {
        "category": "⛰️ 토지 · 일반",
        "primary": "#15803D",
        "hover": "#166534",
        "soft_bg": "#F0FDF4",
        "badge_bg": "#DCFCE7",
    },
}

st.set_page_config(
    page_title="통합계약서 — 부동산·요양원 전문 작성기",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .block-container {
        padding-top: 0.45rem !important;
        padding-bottom: 1.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
      }
      .app-header-bar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 0.75rem;
        flex-wrap: wrap;
        padding: 0.7rem 1.1rem;
        border-radius: 12px;
        background: linear-gradient(135deg, #0b1c3d 0%, #163368 55%, #1d4ed8 100%);
        border-bottom: 3.5px solid #f59e0b;
        box-shadow: 0 5px 16px rgba(15, 23, 42, 0.16);
        color: #f8fafc;
        margin-bottom: 0.55rem;
      }
      .app-header-left {
        display: flex;
        align-items: center;
        gap: 0.75rem;
      }
      .app-header-icon {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 42px;
        height: 42px;
        border-radius: 10px;
        background: rgba(255, 255, 255, 0.14);
        border: 1.5px solid rgba(251, 191, 36, 0.65);
        font-size: 1.45rem;
        flex-shrink: 0;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.25);
      }
      .app-header-title {
        margin: 0;
        font-size: 1.65rem;
        font-weight: 900;
        letter-spacing: 0.08em;
        line-height: 1.15;
        color: #ffffff;
      }
      .app-header-sub {
        margin: 0.18rem 0 0 0;
        font-size: 0.84rem;
        font-weight: 500;
        color: #dbeafe;
        letter-spacing: 0.01em;
      }
      .app-header-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 14px;
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.14);
        border: 1px solid rgba(255, 255, 255, 0.35);
        color: #ffffff;
        font-size: 0.85rem;
        font-weight: 700;
        backdrop-filter: blur(4px);
      }
      .type-card-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        flex-wrap: wrap;
        margin-bottom: 0.38rem;
      }
      .type-active-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 3px 11px;
        border-radius: 999px;
        color: #ffffff;
        font-size: 0.8rem;
        font-weight: 800;
        letter-spacing: 0.02em;
      }
      .type-hint-badge {
        display: inline-flex;
        align-items: center;
        padding: 2px 9px;
        border-radius: 999px;
        font-size: 0.76rem;
        font-weight: 700;
      }
      .kr-amount-hint {
        font-size: 0.82rem;
        color: #1d4ed8;
        font-weight: 600;
        margin-top: -0.4rem;
        margin-bottom: 0.45rem;
      }
      .doc-id-pill {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        background: #f1f5f9;
        border: 1px solid #cbd5e1;
        font-size: 0.82rem;
        font-weight: 600;
        color: #0f172a;
      }
      .sb-card {
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.5rem 0.6rem;
        margin-bottom: 0.4rem;
        background: #ffffff;
      }
      .st-key-font_size_bar {
        background: #f8fafc;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 0.25rem 0.75rem;
        margin-bottom: 0.5rem;
      }
      .st-key-font_size_bar [data-testid="stRadio"],
      .st-key-preview_toggle_box [data-testid="stRadio"] {
        display: flex;
        flex-direction: row;
        align-items: center;
        flex-wrap: wrap;
        gap: 0.65rem;
      }
      .st-key-font_size_bar [data-testid="stWidgetLabel"] {
        margin-bottom: 0 !important;
        font-weight: 700 !important;
        color: #1e293b !important;
      }
      /* 연속 카메라 촬영 스튜디오: 세로로 시원하게 확대 + 군더더기 제거 + 강조 저장 버튼 */
      .cam-count-badge {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 8px;
        width: 100%;
        padding: 0.65rem 1rem;
        border-radius: 12px;
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        color: #ffffff;
        font-size: 1.08rem;
        font-weight: 800;
        letter-spacing: 0.02em;
        margin-bottom: 0.55rem;
        box-shadow: 0 3px 10px rgba(15, 23, 42, 0.14);
      }
      .cam-count-pill {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 999px;
        background: #f59e0b;
        color: #0f172a;
        font-weight: 900;
      }
      .st-key-camera_studio [data-testid="stCameraInput"] {
        width: 100% !important;
      }
      .st-key-camera_studio [data-testid="stCameraInput"] > div {
        width: 100% !important;
      }
      .st-key-camera_studio [data-testid="stCameraInput"] video,
      .st-key-camera_studio [data-testid="stCameraInput"] img {
        width: 100% !important;
        min-height: 440px !important;
        height: 56vh !important;
        max-height: 68vh !important;
        object-fit: cover !important;
        border-radius: 12px !important;
        background: #0f172a !important;
      }
      .st-key-camera_studio [data-testid="stCameraInputButton"] {
        width: 100% !important;
        min-height: 50px !important;
        font-size: 16px !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
        margin-top: 0.35rem !important;
      }
      .st-key-cam_add_btn_wrap .stButton > button {
        min-height: 54px !important;
        border-radius: 12px !important;
        border: 2px solid #2563eb !important;
        background: #eff6ff !important;
        color: #1d4ed8 !important;
        font-weight: 800 !important;
      }
      .st-key-cam_add_btn_wrap .stButton > button p {
        font-size: 16.5px !important;
        font-weight: 800 !important;
        color: #1d4ed8 !important;
      }
      .st-key-cam_add_btn_wrap .stButton > button:disabled {
        opacity: 0.48 !important;
      }
      .st-key-cam_finish_btn_wrap .stButton > button {
        min-height: 56px !important;
        border-radius: 12px !important;
        border: none !important;
        background: #16a34a !important;
        color: #ffffff !important;
        font-weight: 900 !important;
        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.3) !important;
      }
      .st-key-cam_finish_btn_wrap .stButton > button:hover {
        background: #15803d !important;
      }
      .st-key-cam_finish_btn_wrap .stButton > button p {
        font-size: 17px !important;
        font-weight: 900 !important;
        color: #ffffff !important;
      }
      .st-key-cam_finish_btn_wrap .stButton > button:disabled {
        background: #94a3b8 !important;
        box-shadow: none !important;
      }
      .st-key-cam_thumb_strip [data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 0.35rem !important;
      }
      .st-key-cam_thumb_strip [data-testid="stColumn"] {
        flex: 1 1 0% !important;
        min-width: 0 !important;
        width: auto !important;
      }
      .st-key-cam_thumb_strip img {
        border-radius: 8px !important;
        border: 1.5px solid #cbd5e1 !important;
        height: 72px !important;
        object-fit: cover !important;
        width: 100% !important;
      }
      .st-key-cam_thumb_strip .stButton > button {
        min-height: 28px !important;
        height: 28px !important;
        padding: 0 4px !important;
        border-radius: 6px !important;
        background: #fee2e2 !important;
        border: 1px solid #fca5a5 !important;
        color: #b91c1c !important;
      }
      .st-key-cam_thumb_strip .stButton > button p {
        font-size: 12px !important;
        font-weight: 800 !important;
        color: #b91c1c !important;
      }
      /* 전문가용 특약 순서 관리(▲/▼/삭제) 버튼 콤팩트 가로 고정 */
      .st-key-clause_order_list [data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        align-items: center !important;
        gap: 0.28rem !important;
      }
      .st-key-clause_order_list .stButton > button {
        min-height: 32px !important;
        height: 32px !important;
        padding: 0 6px !important;
        border-radius: 6px !important;
      }
      .st-key-clause_order_list .stButton > button p {
        font-size: 13px !important;
        font-weight: 800 !important;
        line-height: 1 !important;
      }
      .clause-order-card {
        border: 1px solid #cbd5e1;
        border-left: 3.5px solid #1e3a5f;
        border-radius: 8px;
        padding: 0.4rem 0.6rem;
        background: #f8fafc;
        font-size: 0.85rem;
        line-height: 1.42;
        color: #0f172a;
      }
      /* Sticky 동작 보장을 위해 워크스페이스 래퍼 overflow: visible 명시 */
      .st-key-workspace_split,
      .st-key-workspace_split [data-testid="stVerticalBlock"],
      .st-key-workspace_split [data-testid="stHorizontalBlock"]:has(.st-key-form_pane) {
        overflow: visible !important;
      }
      /* 데스크탑: 모바일 전용 탭 숨김 + 좌측 콤팩트 입력 폼 + 우측 계약서 미리보기 Sticky 고정 */
      @media (min-width: 769px) {
        .st-key-mobile_view_switch {
          display: none !important;
        }
        .st-key-workspace_split [data-testid="stHorizontalBlock"]:has(.st-key-form_pane) {
          align-items: flex-start !important;
        }
        .st-key-workspace_split [data-testid="stColumn"]:has(.st-key-preview_pane) {
          position: sticky !important;
          top: 1rem !important;
          max-height: 92vh !important;
          overflow-y: auto !important;
          align-self: flex-start !important;
          padding-right: 4px;
        }
        .st-key-form_pane [data-testid="stVerticalBlock"] {
          gap: 0.42rem !important;
        }
        .st-key-form_pane [data-testid="stExpander"] summary {
          padding: 0.45rem 0.7rem !important;
        }
        .st-key-form_pane [data-testid="stExpanderDetails"] {
          padding: 0.55rem 0.65rem 0.65rem 0.65rem !important;
        }
        .st-key-form_pane .stTextInput input,
        .st-key-form_pane .stNumberInput input,
        .st-key-form_pane .stDateInput input {
          padding-top: 0.28rem !important;
          padding-bottom: 0.28rem !important;
          min-height: 36px;
        }
        .st-key-form_pane [data-testid="stWidgetLabel"] {
          margin-bottom: 0.12rem !important;
        }
        .st-key-form_pane hr {
          margin: 0.45rem 0 !important;
        }
        .st-key-preview_iframe_box iframe {
          border: 1px solid #cbd5e1 !important;
          border-radius: 10px !important;
          background: #e9edf2 !important;
          box-shadow: 0 4px 14px rgba(15, 23, 42, 0.08) !important;
          touch-action: pan-x pan-y pinch-zoom !important;
        }
      }
      @media (max-width: 768px) {
        html, body, .stApp {
          -webkit-text-size-adjust: 100%;
          touch-action: pan-x pan-y pinch-zoom !important;
        }
        .block-container {
          padding-left: 0.45rem !important;
          padding-right: 0.45rem !important;
          padding-top: 0.35rem !important;
        }
        .app-header-bar {
          padding: 0.5rem 0.75rem !important;
          margin-bottom: 0.35rem !important;
          gap: 0.4rem !important;
        }
        .app-header-icon {
          width: 34px !important;
          height: 34px !important;
          font-size: 1.15rem !important;
        }
        .app-header-title {
          font-size: 1.28rem !important;
        }
        .app-header-sub {
          display: none !important;
        }
        .app-header-pill {
          padding: 3px 9px !important;
          font-size: 12.5px !important;
        }
        .st-key-font_size_bar {
          padding: 0.2rem 0.55rem !important;
          margin-bottom: 0.35rem !important;
        }
        .st-key-contract_type_card {
          padding: 0.4rem 0.55rem 0.45rem 0.55rem !important;
        }
        .st-key-top_controls_row [data-testid="stVerticalBlock"] {
          gap: 0.3rem !important;
        }
        /* 모바일 상단 버튼 다이어트: 잠금/임시저장/최종저장 3등분 1줄 + 예시/초기화 2등분 1줄 강제 */
        .st-key-action_buttons_row [data-testid="stHorizontalBlock"],
        .st-key-quick_tools_row [data-testid="stHorizontalBlock"],
        .st-key-mobile_view_switch [data-testid="stHorizontalBlock"] {
          display: flex !important;
          flex-direction: row !important;
          flex-wrap: nowrap !important;
          gap: 0.35rem !important;
        }
        .st-key-action_buttons_row [data-testid="stColumn"],
        .st-key-quick_tools_row [data-testid="stColumn"],
        .st-key-mobile_view_switch [data-testid="stColumn"] {
          flex: 1 1 0% !important;
          min-width: 0 !important;
          width: auto !important;
        }
        .st-key-action_buttons_row .stButton > button,
        .st-key-quick_tools_row .stButton > button {
          min-height: 36px !important;
          height: 36px !important;
          padding: 0.15rem 0.3rem !important;
          white-space: nowrap !important;
          border-radius: 8px !important;
        }
        .st-key-action_buttons_row .stButton > button p,
        .st-key-quick_tools_row .stButton > button p {
          font-size: 13px !important;
          line-height: 1.1 !important;
          white-space: nowrap !important;
        }
        /* 모바일 전용 [✏️ 계약서 입력 / 📄 계약서 크게 보기] 전환 바 */
        .st-key-mobile_view_switch {
          position: sticky !important;
          top: 0 !important;
          z-index: 950 !important;
          background: rgba(248, 250, 252, 0.96) !important;
          backdrop-filter: blur(6px);
          border: 1.5px solid #cbd5e1 !important;
          border-radius: 10px !important;
          padding: 0.3rem !important;
          margin: 0.35rem 0 0.5rem 0 !important;
          box-shadow: 0 4px 12px rgba(15, 23, 42, 0.1) !important;
        }
        .st-key-mobile_view_switch .stButton > button {
          min-height: 44px !important;
          height: 44px !important;
          padding: 0.25rem 0.4rem !important;
          border-radius: 8px !important;
          white-space: nowrap !important;
        }
        .st-key-mobile_view_switch .stButton > button p {
          font-size: 15px !important;
          font-weight: 800 !important;
          line-height: 1.15 !important;
          white-space: nowrap !important;
        }
        /* 모바일 워크스페이스: 선택한 탭만 100% 전체 폭·높이로 시원하게 표시 */
        .st-key-workspace_split [data-testid="stHorizontalBlock"]:has(.st-key-form_pane) {
          display: flex !important;
          flex-direction: column !important;
          flex-wrap: nowrap !important;
          gap: 0.5rem !important;
        }
        .st-key-workspace_split [data-testid="stColumn"]:has(.st-key-form_pane),
        .st-key-workspace_split [data-testid="stColumn"]:has(.st-key-preview_pane) {
          width: 100% !important;
          min-width: 100% !important;
          flex: none !important;
        }
        .st-key-preview_pane,
        .st-key-preview_iframe_box,
        .st-key-preview_iframe_box iframe {
          touch-action: pan-x pan-y pinch-zoom !important;
        }
        .st-key-preview_iframe_box iframe {
          height: 83vh !important;
          min-height: 520px !important;
          width: 100% !important;
          border: 1px solid #cbd5e1 !important;
          border-radius: 10px !important;
          background: #e9edf2 !important;
        }
        .app-header-sub,
        .app-header-pill {
          font-size: 15px !important;
        }
        .stApp p,
        .stApp li,
        .stMarkdown,
        .stMarkdown p,
        .stMarkdown li,
        [data-testid="stWidgetLabel"],
        [data-testid="stWidgetLabel"] p,
        [data-testid="stExpander"] summary,
        [data-testid="stExpander"] summary p,
        [data-testid="stAlert"] p,
        .stRadio label,
        .stRadio label p,
        .stCheckbox label,
        .stCheckbox label p,
        [data-baseweb="tab"],
        [data-baseweb="tab"] p {
          font-size: 17px !important;
          line-height: 1.5 !important;
        }
        .stTextInput input,
        .stNumberInput input,
        .stTextArea textarea,
        .stDateInput input,
        [data-baseweb="input"] input,
        [data-baseweb="textarea"] textarea,
        [data-baseweb="select"],
        [data-baseweb="select"] span,
        [data-baseweb="select"] div,
        [data-baseweb="select"] input,
        .st-key-contract_type [data-baseweb="select"] span,
        .st-key-contract_type [data-baseweb="select"] div[value],
        .st-key-contract_type [data-baseweb="select"] input,
        [data-baseweb="menu"] li,
        [data-baseweb="popover"] li {
          font-size: 18px !important;
          line-height: 1.45 !important;
        }
        .stTextInput input,
        .stNumberInput input,
        .stDateInput input,
        [data-baseweb="select"] > div {
          min-height: 46px !important;
        }
        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stFormSubmitButton"] > button,
        .stButton > button p,
        .stDownloadButton > button p,
        [data-testid="stFormSubmitButton"] > button p {
          font-size: 17.5px !important;
          font-weight: 700 !important;
        }
        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stFormSubmitButton"] > button {
          min-height: 48px !important;
          padding: 0.55rem 0.9rem !important;
        }
        .kr-amount-hint,
        .doc-id-pill,
        .type-active-badge,
        .type-hint-badge,
        [data-testid="stCaptionContainer"],
        [data-testid="stCaptionContainer"] p,
        [data-testid="stFileUploader"] small {
          font-size: 15.5px !important;
          line-height: 1.45 !important;
        }
      }
    </style>
    """,
    unsafe_allow_html=True,
)


FONT_SCALE_PRESETS: dict[str, dict[str, dict[str, str]]] = {
    "보통": {
        "desktop": {},
        "mobile": {
            "body": "17px",
            "input": "18px",
            "button": "17.5px",
            "btn_height": "48px",
            "input_height": "46px",
            "caption": "15.5px",
        },
    },
    "크게": {
        "desktop": {
            "body": "17px",
            "input": "17.5px",
            "button": "16.5px",
            "btn_height": "44px",
            "input_height": "44px",
            "caption": "15px",
        },
        "mobile": {
            "body": "19.5px",
            "input": "20px",
            "button": "19.5px",
            "btn_height": "52px",
            "input_height": "50px",
            "caption": "17px",
        },
    },
    "아주 크게": {
        "desktop": {
            "body": "20px",
            "input": "20.5px",
            "button": "19px",
            "btn_height": "50px",
            "input_height": "50px",
            "caption": "17px",
        },
        "mobile": {
            "body": "22px",
            "input": "22.5px",
            "button": "21.5px",
            "btn_height": "56px",
            "input_height": "54px",
            "caption": "19px",
        },
    },
}


def _build_font_rules(cfg: dict[str, str]) -> str:
    if not cfg:
        return ""
    return f"""
        .stApp p,
        .stApp li,
        .stMarkdown,
        .stMarkdown p,
        .stMarkdown li,
        [data-testid="stWidgetLabel"],
        [data-testid="stWidgetLabel"] p,
        [data-testid="stExpander"] summary,
        [data-testid="stExpander"] summary p,
        [data-testid="stAlert"] p,
        .stRadio label,
        .stRadio label p,
        .stCheckbox label,
        .stCheckbox label p,
        [data-baseweb="tab"],
        [data-baseweb="tab"] p {{
          font-size: {cfg['body']} !important;
          line-height: 1.5 !important;
        }}
        .stTextInput input,
        .stNumberInput input,
        .stTextArea textarea,
        .stDateInput input,
        [data-baseweb="input"] input,
        [data-baseweb="textarea"] textarea,
        [data-baseweb="select"],
        [data-baseweb="select"] span,
        [data-baseweb="select"] div,
        [data-baseweb="select"] input,
        .st-key-contract_type [data-baseweb="select"] span,
        .st-key-contract_type [data-baseweb="select"] div[value],
        .st-key-contract_type [data-baseweb="select"] input,
        [data-baseweb="menu"] li,
        [data-baseweb="popover"] li {{
          font-size: {cfg['input']} !important;
          line-height: 1.45 !important;
        }}
        .stTextInput input,
        .stNumberInput input,
        .stDateInput input,
        [data-baseweb="select"] > div {{
          min-height: {cfg['input_height']} !important;
        }}
        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stFormSubmitButton"] > button,
        .stButton > button p,
        .stDownloadButton > button p,
        [data-testid="stFormSubmitButton"] > button p {{
          font-size: {cfg['button']} !important;
          font-weight: 700 !important;
        }}
        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stFormSubmitButton"] > button {{
          min-height: {cfg['btn_height']} !important;
        }}
        .kr-amount-hint,
        .doc-id-pill,
        .type-active-badge,
        .type-hint-badge,
        [data-testid="stCaptionContainer"],
        [data-testid="stCaptionContainer"] p,
        [data-testid="stFileUploader"] small {{
          font-size: {cfg['caption']} !important;
          line-height: 1.45 !important;
        }}
        .app-header-sub,
        .app-header-pill {{
          font-size: {cfg['caption']} !important;
        }}
    """


def inject_font_scale_css(font_scale: str) -> None:
    """사용자가 선택한 글자 크기('보통'/'크게'/'아주 크게')와 모바일 확대 스타일을 즉시 적용한다."""
    preset = FONT_SCALE_PRESETS.get(font_scale, FONT_SCALE_PRESETS["보통"])
    desktop_css = _build_font_rules(preset["desktop"])
    mobile_css = _build_font_rules(preset["mobile"])
    st.markdown(
        f"""
        <style>
          {desktop_css}
          @media (max-width: 768px) {{
            {mobile_css}
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def inject_mobile_pinch_zoom_meta() -> None:
    """상위 문서의 viewport 메타태그에 두 손가락 핀치 줌(확대/축소) 허용을 강제 적용한다."""
    components.html(
        """
        <script>
          (function () {
            var content = "width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes";
            function patchDoc(doc) {
              if (!doc || !doc.head) return;
              var meta = doc.querySelector('meta[name="viewport"]');
              if (!meta) {
                meta = doc.createElement('meta');
                meta.name = 'viewport';
                doc.head.appendChild(meta);
              }
              if (meta.getAttribute('content') !== content) {
                meta.setAttribute('content', content);
              }
            }
            try { patchDoc(window.parent.document); } catch (e) {}
            try { patchDoc(window.top.document); } catch (e) {}
          })();
        </script>
        """,
        height=0,
        width=0,
    )


def inject_mobile_tab_css(active_tab: str) -> None:
    """모바일(<=768px)에서만 선택한 탭('input' 또는 'preview')이 전체 화면을 차지하도록 제어한다."""
    if active_tab == "preview":
        mobile_rule = """
            .st-key-workspace_split [data-testid="stColumn"]:has(.st-key-form_pane) {
              display: none !important;
            }
            .st-key-workspace_split [data-testid="stColumn"]:has(.st-key-preview_pane) {
              display: block !important;
            }
        """
    else:
        mobile_rule = """
            .st-key-workspace_split [data-testid="stColumn"]:has(.st-key-preview_pane) {
              display: none !important;
            }
            .st-key-workspace_split [data-testid="stColumn"]:has(.st-key-form_pane) {
              display: block !important;
            }
        """
    st.markdown(
        f"""
        <style>
          @media (max-width: 768px) {{
            {mobile_rule}
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _enhance_live_preview_html(printable_html: str) -> str:
    """실시간 A4 미리보기 iframe 전용 핀치줌·확대/축소 버튼([축소 -][화면맞춤][크게 100%][더 크게 130%]) 최적화.

    다운로드 및 인쇄용 원본 printable_html은 변경하지 않는다.
    """
    enhanced = printable_html.replace(
        '<meta name="viewport" content="width=device-width, initial-scale=1.0" />',
        '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes" />',
    )

    zoom_toolbar_html = """<div class="zoom-controls">
      <button type="button" id="btnZoomOut" title="축소">축소 -</button>
      <button type="button" id="btnZoomFit" title="화면 폭에 맞춤">화면맞춤</button>
      <button type="button" id="btnZoom100" title="A4 100% 크기로 크게 보기">크게 100%</button>
      <button type="button" id="btnZoom130" title="A4 130% 크기로 더 크게 보기">더 크게 130%</button>
      <button type="button" class="print-btn" onclick="window.print()">🖨️ 인쇄</button>
    </div>"""
    enhanced = enhanced.replace(
        '<button onclick="window.print()">🖨️ 인쇄 / PDF로 저장</button>',
        zoom_toolbar_html,
    )

    # .preview-wrapper를 독립 4방향 스크롤 뷰포트로 감싸 100%/130% 확대 시에도 상단 툴바 고정 + 부드러운 좌우/상하 스와이프 보장
    enhanced = enhanced.replace(
        '<div class="preview-wrapper">',
        '<div id="previewScrollViewport" class="preview-scroll-viewport"><div class="preview-wrapper">',
        1,
    )

    extra_css = """
    <style>
      @media screen {
        html, body {
          height: 100% !important;
          margin: 0 !important;
          padding: 0 !important;
          overflow: hidden !important;
          display: flex !important;
          flex-direction: column !important;
          touch-action: pan-x pan-y pinch-zoom !important;
        }
        .toolbar {
          flex-shrink: 0 !important;
          position: relative !important;
          top: 0 !important;
          left: 0 !important;
          width: 100% !important;
          padding: 6px 8px !important;
          gap: 6px !important;
          z-index: 300 !important;
        }
        .toolbar-meta {
          gap: 5px !important;
          font-size: 11.5px !important;
        }
        .zoom-controls {
          display: flex;
          align-items: center;
          gap: 5px;
          flex-wrap: wrap;
        }
        .zoom-controls button {
          background: #334155 !important;
          color: #f8fafc !important;
          border: 1px solid #64748b !important;
          border-radius: 6px !important;
          padding: 5px 9px !important;
          font-size: 12.5px !important;
          font-weight: 700 !important;
          min-height: 32px !important;
          cursor: pointer;
          white-space: nowrap;
        }
        .zoom-controls button:hover {
          background: #475569 !important;
        }
        .zoom-controls button.active {
          background: #f59e0b !important;
          color: #0f172a !important;
          border-color: #fbbf24 !important;
        }
        .zoom-controls button.print-btn {
          background: #2563eb !important;
          border-color: #3b82f6 !important;
        }
        .preview-scroll-viewport {
          flex: 1 1 auto !important;
          width: 100% !important;
          overflow: auto !important;
          -webkit-overflow-scrolling: touch !important;
          touch-action: pan-x pan-y pinch-zoom !important;
          background: #e9edf2;
        }
        .preview-wrapper {
          display: flex !important;
          flex-direction: column !important;
          align-items: flex-start !important;
          padding: 8px 6px 24px 6px !important;
          width: max-content !important;
          min-width: 100% !important;
          box-sizing: border-box !important;
          touch-action: pan-x pan-y pinch-zoom !important;
        }
        .page-scale-slot {
          margin: 0 auto !important;
          touch-action: pan-x pan-y pinch-zoom !important;
        }
        .a4-page {
          -webkit-font-smoothing: antialiased;
          text-rendering: optimizeLegibility;
          touch-action: pan-x pan-y pinch-zoom !important;
        }
      }
      @media print {
        html, body {
          height: auto !important;
          overflow: visible !important;
          display: block !important;
        }
        .preview-scroll-viewport {
          height: auto !important;
          overflow: visible !important;
        }
      }
    </style>
    </head>"""
    if "</head>" in enhanced:
        enhanced = enhanced.replace("</head>", extra_css, 1)

    zoom_script = """</div>
  <script>
    (function () {
      var STORE_KEY = 'contract_preview_zoom_v2';
      var zoomMode = 'fit';
      var customScale = 1.0;

      try {
        var saved = window.sessionStorage.getItem(STORE_KEY);
        if (saved) {
          var parsed = JSON.parse(saved);
          if (parsed && parsed.mode) {
            zoomMode = parsed.mode;
            if (typeof parsed.scale === 'number') customScale = parsed.scale;
          }
        }
      } catch (e) {}

      function saveZoomState() {
        try {
          window.sessionStorage.setItem(STORE_KEY, JSON.stringify({ mode: zoomMode, scale: customScale }));
        } catch (e) {}
      }

      function computeFitScale(rawW) {
        var vp = document.getElementById('previewScrollViewport');
        var cw = vp ? vp.clientWidth : document.documentElement.clientWidth;
        var availableW = Math.max(260, cw - 14);
        var maxFit = cw >= 700 ? 1.38 : 1.0;
        return Math.min(maxFit, availableW / rawW);
      }

      function updateActiveButtons(appliedScale) {
        var bFit = document.getElementById('btnZoomFit');
        var b100 = document.getElementById('btnZoom100');
        var b130 = document.getElementById('btnZoom130');
        [bFit, b100, b130].forEach(function (b) { if (b) b.classList.remove('active'); });
        if (zoomMode === 'fit' && bFit) {
          bFit.classList.add('active');
        } else if (zoomMode === 'fixed' && Math.abs(customScale - 1.0) < 0.02 && b100) {
          b100.classList.add('active');
        } else if (zoomMode === 'fixed' && Math.abs(customScale - 1.3) < 0.02 && b130) {
          b130.classList.add('active');
        }
        var badge = document.getElementById('zoomBadge');
        if (badge) {
          var pct = Math.round(appliedScale * 100);
          badge.textContent = (zoomMode === 'fit' ? '화면맞춤 ' : '확대 ') + pct + '%';
        }
      }

      function applyZoom() {
        var slots = document.querySelectorAll('.page-scale-slot');
        var appliedScale = 1.0;
        slots.forEach(function (slot) {
          var page = slot.querySelector('.a4-page');
          if (!page) return;
          page.style.transform = 'none';
          var rawW = page.offsetWidth || 794;
          var rawH = page.offsetHeight || 1123;
          var scale = (zoomMode === 'fit') ? computeFitScale(rawW) : Math.max(0.4, Math.min(2.0, customScale));
          appliedScale = scale;
          page.style.transform = 'scale(' + scale + ')';
          page.style.transformOrigin = 'top left';
          slot.style.width = Math.ceil(rawW * scale) + 'px';
          slot.style.height = Math.ceil(rawH * scale) + 'px';
        });
        updateActiveButtons(appliedScale);
      }

      function resetForPrint() {
        document.querySelectorAll('.page-scale-slot').forEach(function (slot) {
          var page = slot.querySelector('.a4-page');
          if (page) page.style.transform = 'none';
          slot.style.width = '100%';
          slot.style.height = 'auto';
        });
      }

      function bindZoomControls() {
        var bOut = document.getElementById('btnZoomOut');
        var bFit = document.getElementById('btnZoomFit');
        var b100 = document.getElementById('btnZoom100');
        var b130 = document.getElementById('btnZoom130');
        if (bOut) {
          bOut.addEventListener('click', function () {
            var current = (zoomMode === 'fit') ? computeFitScale(794) : customScale;
            zoomMode = 'fixed';
            customScale = Math.max(0.4, Math.round((current - 0.15) * 100) / 100);
            saveZoomState();
            applyZoom();
          });
        }
        if (bFit) {
          bFit.addEventListener('click', function () {
            zoomMode = 'fit';
            saveZoomState();
            applyZoom();
          });
        }
        if (b100) {
          b100.addEventListener('click', function () {
            zoomMode = 'fixed';
            customScale = 1.0;
            saveZoomState();
            applyZoom();
          });
        }
        if (b130) {
          b130.addEventListener('click', function () {
            zoomMode = 'fixed';
            customScale = 1.3;
            saveZoomState();
            applyZoom();
          });
        }
      }

      window.addEventListener('DOMContentLoaded', function () { bindZoomControls(); applyZoom(); });
      window.addEventListener('load', applyZoom);
      window.addEventListener('resize', function () { if (zoomMode === 'fit') applyZoom(); });
      window.addEventListener('beforeprint', resetForPrint);
      window.addEventListener('afterprint', applyZoom);
      var vpEl = document.getElementById('previewScrollViewport');
      if (window.ResizeObserver && vpEl) {
        new ResizeObserver(function () { if (zoomMode === 'fit') applyZoom(); }).observe(vpEl);
      }
      bindZoomControls();
      applyZoom();
    })();
  </script>"""

    script_start = enhanced.find("<script>")
    script_end = enhanced.rfind("</script>")
    if script_start != -1 and script_end != -1:
        enhanced = enhanced[:script_start] + zoom_script + enhanced[script_end + len("</script>"):]
    return enhanced


@st.cache_data
def _load_clauses_library_cached(file_mtime: float) -> dict[str, Any]:
    with CLAUSES_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def load_clauses_library() -> dict[str, Any]:
    """data/clauses.json 수정 시각(mtime)을 캐시 키로 사용해 변경 즉시 드롭다운에 반영한다."""
    mtime = CLAUSES_PATH.stat().st_mtime if CLAUSES_PATH.exists() else 0.0
    return _load_clauses_library_cached(mtime)


def get_ordered_type_keys(clauses_lib: dict[str, Any]) -> list[str]:
    """CONTRACT_TYPE_META 기준 7대 정식 계약 유형 순서를 보장한다."""
    ordered = [k for k in CONTRACT_TYPE_META if k in clauses_lib]
    for k in clauses_lib:
        if k not in ordered:
            ordered.append(k)
    return ordered


def get_configured_password() -> str:
    """st.secrets['APP_PASSWORD']에서 접속 비밀번호를 읽는다. 미설정 시 빈 문자열('') 반환."""
    try:
        pw = st.secrets.get("APP_PASSWORD")
        return str(pw).strip() if pw is not None else ""
    except Exception:
        return ""


def render_password_lock_screen(expected_password: str) -> None:
    """외부 접속 차단용 중앙 로그인 잠금 화면 (비밀번호 미인증 시 본 화면·사이드바 완전 차단)."""
    st.markdown(
        """
        <style>
          [data-testid="stSidebar"],
          [data-testid="collapsedControl"] {
            display: none !important;
          }
          .lock-card-header {
            text-align: center;
            padding: 1.4rem 1.2rem 1.1rem 1.2rem;
            border-radius: 14px 14px 0 0;
            background: linear-gradient(135deg, #0b1c3d 0%, #163368 60%, #1d4ed8 100%);
            border-bottom: 3.5px solid #f59e0b;
            color: #ffffff;
            margin-top: 4rem;
          }
          .lock-card-header h2 {
            margin: 0.35rem 0 0.2rem 0;
            font-size: 1.7rem;
            font-weight: 900;
            letter-spacing: 0.08em;
            color: #ffffff;
          }
          .lock-card-header p {
            margin: 0;
            font-size: 0.85rem;
            color: #dbeafe;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )

    _, center, _ = st.columns([1.15, 1.1, 1.15])
    with center:
        st.markdown(
            """
            <div class="lock-card-header">
              <div style="font-size: 2rem;">🔐</div>
              <h2>통합계약서</h2>
              <p>인가된 사용자만 접속할 수 있습니다.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("password_login_form", clear_on_submit=False):
            entered = st.text_input(
                "비밀번호",
                type="password",
                placeholder="접속 비밀번호를 입력하세요",
            )
            submitted = st.form_submit_button("로그인", use_container_width=True, type="primary")
            if submitted:
                if hmac.compare_digest(entered, expected_password):
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("비밀번호가 올바르지 않습니다.")

        st.caption("🛡️ 비밀번호는 Streamlit Secrets(`APP_PASSWORD`)에서 안전하게 검증됩니다.")


def _empty_form_values() -> dict[str, Any]:
    """모든 텍스트는 '', 숫자는 0으로 완전히 비워진 초기 폼 사전."""
    today = date.today()
    balance_dt = today + timedelta(days=45)
    lease_end_dt = date(balance_dt.year + 2, balance_dt.month, balance_dt.day)

    return {
        "address": "",
        "land_category": "",
        "land_share_ratio": "",
        "land_area": "",
        "land_zoning": "",
        "land_status": "",
        "extra_parcels": "",
        "building_structure": "",
        "building_use": "",
        "building_area": "",
        "lease_part": "",
        "lease_area": "",
        "business_type": "",
        "facility_name": "",
        "facility_capacity": "",
        "facility_reg_no": "",
        "store_name": "",
        "new_lease_summary": "",
        "deposit": 0,
        "down_payment": 0,
        "interim_payment": 0,
        "balance": 0,
        "loan_assumption": 0,
        "lease_deposit_assumption": 0,
        "monthly_rent": 0,
        "monthly_rent_day": "",
        "rent_pay_type": "선불",
        "vat_type": "부가가치세 별도",
        "management_fee": 0,
        "management_fee_note": "",
        "escrow_amount": 0,
        "escrow_period": "",
        "escrow_condition": "",
        "auto_balance": True,
        "down_receipt_person": "",
        "use_interim_date": False,
        "interim_date": today + timedelta(days=20),
        "balance_date": balance_dt,
        "handover_date": balance_dt,
        "lease_end_date": lease_end_dt,
        "contract_date": today,
        "attachment_note": "",
        "deposit_account": "",
        "lessor_address": "",
        "lessor_rrn": "",
        "lessor_phone1": "",
        "lessor_phone2": "",
        "lessor_name": "",
        "lessee_address": "",
        "lessee_rrn": "",
        "lessee_phone1": "",
        "lessee_phone2": "",
        "lessee_name": "",
        "broker1_office": "",
        "broker1_address": "",
        "broker1_rep": "",
        "broker1_reg_no": "",
        "broker1_phone": "",
        "broker1_associate": "",
        "use_co_broker": False,
        "broker2_office": "",
        "broker2_address": "",
        "broker2_rep": "",
        "broker2_reg_no": "",
        "broker2_phone": "",
        "broker2_associate": "",
    }


def _sample_form_values(contract_type: str) -> dict[str, Any]:
    """계약 유형별 비식별 실무 예시 데이터."""
    data = _empty_form_values()
    today = date.today()
    balance_dt = today + timedelta(days=45)
    lease_end_dt = date(balance_dt.year + 2, balance_dt.month, balance_dt.day)

    data.update(
        {
            "down_receipt_person": "홍길동",
            "balance_date": balance_dt,
            "handover_date": balance_dt,
            "lease_end_date": lease_end_dt,
            "contract_date": today,
            "attachment_note": "확인설명서 및 공제증서",
            "deposit_account": "기업 010-0000-0000 홍길동 (예시)",
            "lessor_address": "인천광역시 서구 가좌로 123, 201동 604호 (예시)",
            "lessor_rrn": "750101-1******",
            "lessor_phone1": "010-0000-1111",
            "lessor_phone2": "",
            "lessor_name": "홍길동",
            "lessee_address": "경기도 남양주시 오남읍 진건오남로 781-17, 103동 603호 (예시)",
            "lessee_rrn": "830723-1******",
            "lessee_phone1": "010-0000-2222",
            "lessee_phone2": "",
            "lessee_name": "김철수",
            "broker1_office": "한빛공인중개사사무소",
            "broker1_address": "경기도 남양주시 오남읍 진건오남로 756길 상가 102호",
            "broker1_rep": "박중개",
            "broker1_reg_no": "가3622-08-0000",
            "broker1_phone": "031-575-0000",
            "broker1_associate": "",
            "use_co_broker": True,
            "broker2_office": "신한공인중개사사무소",
            "broker2_address": "서울특별시 중랑구 망우로 307 B106호",
            "broker2_rep": "이공동",
            "broker2_reg_no": "11260-2015-00000",
            "broker2_phone": "010-0000-3333",
            "broker2_associate": "",
        }
    )

    if contract_type == "real_estate_sale":
        data.update(
            {
                "address": "경기도 남양주시 다산동 1234 다산자이아파트 205동 1203호",
                "land_category": "대",
                "land_share_ratio": "45120/52.4100",
                "land_area": "52.41",
                "building_structure": "철근콘크리트",
                "building_use": "공동주택(아파트)",
                "building_area": "84.9700",
                "lease_part": "205동 1203호 전부",
                "lease_area": "84.9700",
                "deposit": 750_000_000,
                "down_payment": 75_000_000,
                "loan_assumption": 200_000_000,
                "lease_deposit_assumption": 0,
                "interim_payment": 100_000_000,
                "use_interim_date": True,
                "balance": 375_000_000,
            }
        )
    elif contract_type == "housing_wolse":
        data.update(
            {
                "address": "서울특별시 중랑구 면목동 567-8 현대빌라 302호",
                "land_category": "대",
                "land_share_ratio": "280/24.5",
                "land_area": "24.50",
                "building_structure": "철근콘크리트",
                "building_use": "다세대주택",
                "building_area": "49.80",
                "lease_part": "302호 전부",
                "lease_area": "49.80",
                "deposit": 30_000_000,
                "down_payment": 3_000_000,
                "interim_payment": 0,
                "balance": 27_000_000,
                "monthly_rent": 850_000,
                "monthly_rent_day": "25",
                "rent_pay_type": "선불",
                "management_fee": 70_000,
                "management_fee_note": "공용전기·청소·인터넷 포함, 수도·가스 별도",
            }
        )
    elif contract_type == "commercial_lease":
        data.update(
            {
                "address": "경기도 구리시 인창동 670-1 센트럴프라자 1층 104호",
                "building_structure": "철근콘크리트",
                "building_use": "제1종근린생활시설",
                "building_area": "1250.40",
                "lease_part": "1층 104호 전부",
                "lease_area": "56.20",
                "business_type": "커피전문점(휴게음식점)",
                "deposit": 50_000_000,
                "down_payment": 5_000_000,
                "interim_payment": 0,
                "balance": 45_000_000,
                "monthly_rent": 2_800_000,
                "monthly_rent_day": "10",
                "rent_pay_type": "선불",
                "vat_type": "부가가치세 별도",
                "management_fee": 350_000,
                "management_fee_note": "관리규약 실비 정산, 부가세 별도",
            }
        )
    elif contract_type == "nursing_home_transfer":
        data.update(
            {
                "facility_name": "행복한실버노인요양원 (예시)",
                "facility_reg_no": "2-41360-00000",
                "address": "경기도 남양주시 진접읍 경복대로 123, 3~4층",
                "facility_capacity": "인가정원 29인 / 현원 27인",
                "building_structure": "철근콘크리트",
                "building_use": "노유자시설(노인의료복지시설)",
                "building_area": "680.50",
                "lease_part": "시설·비품·장기요양기관 지정 지위 및 운영권 일체",
                "lease_area": "680.50",
                "deposit": 350_000_000,
                "down_payment": 35_000_000,
                "interim_payment": 100_000_000,
                "use_interim_date": True,
                "escrow_amount": 50_000_000,
                "escrow_period": "6개월",
                "escrow_condition": "현지조사·부당청구 환수·행정처분 대비 정산 후 지급",
                "balance": 165_000_000,
            }
        )
    elif contract_type == "premium_rights":
        data.update(
            {
                "store_name": "오남수제버거 (예시)",
                "business_type": "일반음식점",
                "address": "경기도 남양주시 오남읍 진건오남로 756, 1층 105호",
                "lease_part": "1층 105호 점포 전부",
                "lease_area": "62.80",
                "new_lease_summary": "보증금 4,000만원 / 월세 220만원(VAT별도) / 관리비 실비 / 24개월",
                "deposit": 60_000_000,
                "down_payment": 6_000_000,
                "interim_payment": 0,
                "balance": 54_000_000,
            }
        )
    elif contract_type == "apartment_sale":
        data.update(
            {
                "address": "경기도 남양주시 다산동 6123 다산자이아파트 203동 1402호",
                "land_category": "대",
                "land_share_ratio": "45210/48.2510",
                "land_area": "48.25",
                "building_structure": "철근콘크리트",
                "building_use": "공동주택(아파트)",
                "building_area": "84.9720",
                "lease_part": "203동 1402호 전부",
                "lease_area": "84.9720",
                "deposit": 680_000_000,
                "down_payment": 68_000_000,
                "loan_assumption": 200_000_000,
                "lease_deposit_assumption": 0,
                "interim_payment": 100_000_000,
                "use_interim_date": True,
                "balance": 312_000_000,
            }
        )
    elif contract_type == "commercial_sale":
        data.update(
            {
                "address": "경기도 구리시 인창동 670-1 센트럴프라자 1층 104호",
                "land_category": "대",
                "land_share_ratio": "1850/31.20",
                "land_area": "31.20",
                "building_structure": "철근콘크리트",
                "building_use": "제1종근린생활시설",
                "building_area": "56.20",
                "lease_part": "1층 104호 구분상가 전부",
                "lease_area": "56.20",
                "vat_type": "사업포괄양수도(부가세 없음)",
                "deposit": 750_000_000,
                "down_payment": 75_000_000,
                "loan_assumption": 300_000_000,
                "lease_deposit_assumption": 50_000_000,
                "interim_payment": 0,
                "balance": 325_000_000,
            }
        )
    elif contract_type == "farmland_sale":
        data.update(
            {
                "address": "경기도 양평군 용문면 다문리 456-2",
                "land_category": "답",
                "land_area": "1320.0",
                "land_zoning": "농림지역 (농업진흥구역)",
                "land_status": "논 (경작 중)",
                "lease_part": "농지 전부",
                "extra_parcels": "",
                "deposit": 420_000_000,
                "down_payment": 42_000_000,
                "loan_assumption": 0,
                "interim_payment": 0,
                "balance": 378_000_000,
            }
        )
    elif contract_type == "land_sale":
        data.update(
            {
                "address": "경기도 양평군 양서면 목왕리 123-4",
                "land_category": "전",
                "land_area": "1000.0",
                "land_zoning": "계획관리지역",
                "land_status": "밭 (경작 중)",
                "lease_part": "토지 전부",
                "extra_parcels": "",
                "deposit": 500_000_000,
                "down_payment": 50_000_000,
                "loan_assumption": 100_000_000,
                "interim_payment": 0,
                "balance": 350_000_000,
            }
        )
    else:
        data.update(
            {
                "address": "경기도 남양주시 오남읍 오남리 801 오남아이파크아파트 108동 401호",
                "land_category": "대",
                "land_share_ratio": "33101/55.3748",
                "land_area": "32.01",
                "building_structure": "철근콘크리트",
                "building_use": "공동주택",
                "building_area": "84.9580",
                "lease_part": "108동 401호 전부",
                "lease_area": "84.9580",
                "deposit": 160_000_000,
                "down_payment": 16_000_000,
                "interim_payment": 0,
                "balance": 144_000_000,
            }
        )
    return data


def _purge_old_widget_keys() -> None:
    prefixes = ("f_", "chk_", "txt_", "var_", "_prev_bind_", "new_custom_clause_", "custom_edit_")
    for k in list(st.session_state.keys()):
        if k.startswith(prefixes):
            del st.session_state[k]


def ensure_state_initialized() -> None:
    if "_form_nonce" not in st.session_state:
        st.session_state["_form_nonce"] = 0
    if "_form_data" not in st.session_state:
        st.session_state["_form_data"] = _empty_form_values()
    if "current_contract_id" not in st.session_state:
        st.session_state["current_contract_id"] = generate_contract_id()
    if "current_contract_status" not in st.session_state:
        st.session_state["current_contract_status"] = "unsaved"
    if "contract_type" not in st.session_state:
        st.session_state["contract_type"] = "nursing_home_transfer"
    if "color_mode" not in st.session_state:
        st.session_state["color_mode"] = "color"
    if "ui_font_scale" not in st.session_state:
        st.session_state["ui_font_scale"] = "보통"
    if "mobile_active_tab" not in st.session_state:
        st.session_state["mobile_active_tab"] = "input"
    if "blank_form_mode" not in st.session_state:
        st.session_state["blank_form_mode"] = True
    if "_pending_blank_form_mode" in st.session_state:
        st.session_state["blank_form_mode"] = bool(st.session_state.pop("_pending_blank_form_mode"))
    if "draft_mode" not in st.session_state:
        st.session_state["draft_mode"] = False
    if "custom_clauses_by_type" not in st.session_state:
        st.session_state["custom_clauses_by_type"] = {}
    if "clause_order_by_type" not in st.session_state:
        st.session_state["clause_order_by_type"] = {}
    if "_clause_overrides" not in st.session_state:
        st.session_state["_clause_overrides"] = {}
    if "_ai_clause_feedback" not in st.session_state:
        st.session_state["_ai_clause_feedback"] = None
    if "_expand_att_section" not in st.session_state:
        st.session_state["_expand_att_section"] = False
    if "_expand_party_section" not in st.session_state:
        st.session_state["_expand_party_section"] = False
    if "ai_last_result" not in st.session_state:
        st.session_state["ai_last_result"] = None
    if "cam_step" not in st.session_state:
        st.session_state["cam_step"] = 0
    if "captured_photos" not in st.session_state:
        st.session_state["captured_photos"] = []
    if "camera_mode_open" not in st.session_state:
        st.session_state["camera_mode_open"] = True
    if "_cam_just_saved_msg" not in st.session_state:
        st.session_state["_cam_just_saved_msg"] = ""


def _consume_pending_autofill() -> None:
    """위젯 렌더링 전 최상단에서 대기 버퍼(pending_autofill)를 소비해 st.session_state에 강제 동기화한다.

    6대 계약 유형 전체의 실제 위젯 키(`_wkey(field)`)와 충돌 없이 즉시 반영된다.
    """
    pending = st.session_state.pop("pending_autofill", None)
    if not isinstance(pending, dict):
        return

    extracted: dict[str, Any] = pending.get("mapped_fields", {})
    contract_type: str = str(pending.get("contract_type") or st.session_state.get("contract_type", "housing_jeonse"))

    form = dict(st.session_state.get("_form_data", _empty_form_values()))
    direct_fields = (
        "address",
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
        "lessor_name",
        "lessor_rrn",
        "lessor_address",
        "lessee_name",
        "lessee_rrn",
        "lessee_address",
    )

    filled_keys: list[str] = []
    party_filled = False
    for field in direct_fields:
        val = str(extracted.get(field, "")).strip()
        if val:
            form[field] = val
            filled_keys.append(field)
            if field.startswith(("lessor_", "lessee_")):
                party_filled = True

    max_bond = str(extracted.get("max_bond_amount", "")).strip()
    overrides = dict(st.session_state.get("_clause_overrides", {}))
    if max_bond:
        filled_keys.append("max_bond_amount")
        bond_clause_map = {
            "land_sale": "land_req_04",
            "farmland_sale": "farm_req_04",
            "apartment_sale": "apt_req_02",
            "commercial_sale": "comm_sale_req_02",
            "real_estate_sale": "sale_req_02",
        }
        target_cid = bond_clause_map.get(contract_type, "jeonse_req_02")
        entry = dict(overrides.get(target_cid, {}))
        entry_vars = dict(entry.get("variables", {}))
        entry_vars["채권최고액"] = max_bond
        entry["variables"] = entry_vars
        overrides[target_cid] = entry

    _purge_old_widget_keys()
    new_nonce = int(st.session_state.get("_form_nonce", 0)) + 1
    st.session_state["_form_nonce"] = new_nonce
    st.session_state["_form_data"] = form
    st.session_state["_clause_overrides"] = overrides
    st.session_state["_expand_att_section"] = True
    if filled_keys:
        st.session_state["blank_form_mode"] = False
    if party_filled:
        st.session_state["_expand_party_section"] = True

    # 새 nonce 위젯 키에도 즉시 선반영하여 첫 렌더링부터 100% 동기화 보장
    for field, val in form.items():
        st.session_state[f"f_{new_nonce}_{field}"] = val

    if isinstance(st.session_state.get("ai_last_result"), dict):
        st.session_state["ai_last_result"]["applied_keys"] = filled_keys


def start_new_contract(keep_type: bool = True) -> None:
    """새 계약서 작성 시작 (새 ID 발급 + 폼/특약 완전 초기화 + 빈 양식 모드 활성화)."""
    _purge_old_widget_keys()
    st.session_state["_form_nonce"] = int(st.session_state.get("_form_nonce", 0)) + 1
    st.session_state["_form_data"] = _empty_form_values()
    st.session_state["current_contract_id"] = generate_contract_id()
    st.session_state["current_contract_status"] = "unsaved"
    st.session_state["_pending_blank_form_mode"] = True
    st.session_state["custom_clauses_by_type"] = {}
    st.session_state["clause_order_by_type"] = {}
    st.session_state["_clause_overrides"] = {}
    st.session_state["_ai_clause_feedback"] = None
    st.session_state["ai_last_result"] = None
    st.session_state["_expand_party_section"] = False
    st.session_state["captured_photos"] = []
    st.session_state["cam_step"] = int(st.session_state.get("cam_step", 0)) + 1
    st.session_state["camera_mode_open"] = True
    st.session_state["_cam_just_saved_msg"] = ""
    if not keep_type:
        st.session_state["contract_type"] = "nursing_home_transfer"


def reset_form_to_empty() -> None:
    """[입력 초기화]: 현재 문서 ID는 유지하고 입력값만 전부 '' / 0으로 초기화 + 빈 양식 모드 활성화."""
    _purge_old_widget_keys()
    st.session_state["_form_nonce"] = int(st.session_state.get("_form_nonce", 0)) + 1
    st.session_state["_form_data"] = _empty_form_values()
    st.session_state["_pending_blank_form_mode"] = True
    st.session_state["custom_clauses_by_type"] = {}
    st.session_state["clause_order_by_type"] = {}
    st.session_state["_clause_overrides"] = {}
    st.session_state["_ai_clause_feedback"] = None
    st.session_state["ai_last_result"] = None
    st.session_state["_expand_party_section"] = False


def apply_sample_to_form(contract_type: str) -> None:
    _purge_old_widget_keys()
    st.session_state["_form_nonce"] = int(st.session_state.get("_form_nonce", 0)) + 1
    st.session_state["_form_data"] = _sample_form_values(contract_type)
    st.session_state["_pending_blank_form_mode"] = False
    st.session_state["custom_clauses_by_type"] = {}
    st.session_state["clause_order_by_type"] = {}
    st.session_state["_clause_overrides"] = {}
    st.session_state["_ai_clause_feedback"] = None


def apply_loaded_payload(
    contract_id: str,
    contract_type: str,
    status: str,
    payload: dict[str, Any],
) -> None:
    """DB에서 불러온 계약서 데이터를 세션과 위젯에 정확히 복원한다."""
    _purge_old_widget_keys()
    st.session_state["_form_nonce"] = int(st.session_state.get("_form_nonce", 0)) + 1
    merged = _empty_form_values()
    merged.update(payload.get("form_data", {}))
    st.session_state["_form_data"] = merged
    st.session_state["current_contract_id"] = contract_id
    st.session_state["current_contract_status"] = status
    st.session_state["contract_type"] = contract_type
    st.session_state["color_mode"] = payload.get("color_mode", "color")
    st.session_state["draft_mode"] = bool(payload.get("draft_mode", False))
    st.session_state["_pending_blank_form_mode"] = False
    st.session_state["custom_clauses_by_type"] = payload.get("custom_clauses_by_type", {})
    st.session_state["clause_order_by_type"] = {}
    st.session_state["_clause_overrides"] = payload.get("clause_state", {})
    st.session_state["_ai_clause_feedback"] = None
    st.session_state["captured_photos"] = []
    st.session_state["cam_step"] = int(st.session_state.get("cam_step", 0)) + 1
    st.session_state["camera_mode_open"] = True
    st.session_state["_cam_just_saved_msg"] = ""


def _wkey(field: str) -> str:
    nonce = st.session_state["_form_nonce"]
    key = f"f_{nonce}_{field}"
    if key not in st.session_state:
        st.session_state[key] = st.session_state["_form_data"][field]
    return key


def _capture_current_clause_state(contract_type: str, type_cfg: dict[str, Any]) -> dict[str, Any]:
    """현재 선택된 특약 체크 상태·수정 문구·변수 입력값을 저장용 사전으로 직렬화한다."""
    nonce = st.session_state["_form_nonce"]
    state: dict[str, Any] = {}
    for clause in type_cfg.get("clauses", []):
        cid = clause["id"]
        chk_key = f"chk_{nonce}_{contract_type}_{cid}"
        txt_key = f"txt_{nonce}_{contract_type}_{cid}"
        vars_dict: dict[str, str] = {}
        for var_name in clause.get("variables", {}):
            vkey = f"var_{nonce}_{contract_type}_{cid}_{var_name}"
            if vkey in st.session_state:
                vars_dict[var_name] = str(st.session_state[vkey])
        state[cid] = {
            "checked": bool(st.session_state.get(chk_key, clause.get("required", False))),
            "template": str(st.session_state.get(txt_key, clause["template"])),
            "variables": vars_dict,
        }
    return state


def _build_storage_payload(
    live_form: dict[str, Any],
    contract_type: str,
    type_cfg: dict[str, Any],
) -> dict[str, Any]:
    return {
        "form_data": dict(live_form),
        "clause_state": _capture_current_clause_state(contract_type, type_cfg),
        "custom_clauses_by_type": st.session_state.get("custom_clauses_by_type", {}),
        "color_mode": st.session_state.get("color_mode", "color"),
        "draft_mode": bool(st.session_state.get("draft_mode", False)),
    }


def _queue_pending_autofill(
    ai_result: dict[str, Any],
    contract_type: str,
) -> None:
    """AI 추출 결과를 대기 버퍼(pending_autofill)에 넣고 rerun으로 안전하게 반영한다."""
    st.session_state["ai_last_result"] = ai_result
    st.session_state["_expand_att_section"] = True
    if ai_result.get("ok") and ai_result.get("mapped_fields"):
        st.session_state["pending_autofill"] = {
            "mapped_fields": dict(ai_result["mapped_fields"]),
            "contract_type": contract_type,
        }


def _field_display_name(field: str, p1_role: str, p2_role: str) -> str:
    labels = {
        "address": "소재지",
        "land_category": "지목",
        "land_share_ratio": "대지권비율",
        "land_area": "토지면적",
        "land_zoning": "용도지역·지구",
        "building_structure": "건물 구조",
        "building_use": "건물 용도",
        "building_area": "건물 면적",
        "lease_part": "대상 부분",
        "lease_area": "전용/계약 면적",
        "facility_name": "요양시설 명칭",
        "store_name": "상호(점포명)",
        "max_bond_amount": "채권최고액(특약 연동)",
        "lessor_name": f"{p1_role} 성명",
        "lessor_rrn": f"{p1_role} 주민(사업자)번호",
        "lessor_address": f"{p1_role} 주소",
        "lessee_name": f"{p2_role} 성명",
        "lessee_rrn": f"{p2_role} 주민(사업자)번호",
        "lessee_address": f"{p2_role} 주소",
    }
    return labels.get(field, field)


def render_sidebar_saved_contracts(clauses_lib: dict[str, Any]) -> None:
    """좌측 사이드바: 저장된 계약서 목록 (유형 필터, 상태 필터, 검색, 불러오기, 복사, 삭제)."""
    with st.sidebar:
        st.markdown("## 📂 저장된 계약서 목록")
        if st.button("➕ 새 계약서 작성", use_container_width=True, type="primary"):
            start_new_contract(keep_type=True)
            st.rerun()

        st.markdown("---")
        type_options = ["all"] + get_ordered_type_keys(clauses_lib)
        type_filter = st.selectbox(
            "계약 유형 필터",
            options=type_options,
            format_func=lambda k: "전체 유형" if k == "all" else clauses_lib[k]["label"],
            key="sb_type_filter",
        )
        status_filter = st.selectbox(
            "저장 상태",
            options=["all", "draft", "completed"],
            format_func=lambda s: {"all": "전체 상태", "draft": "🟡 임시저장", "completed": "🟢 최종 완료"}[s],
            key="sb_status_filter",
        )
        search_q = st.text_input(
            "검색 (소재지·시설명·당사자)",
            key="sb_search_query",
            placeholder="예: 오남, 홍길동, 요양원",
        )

        rows = list_contracts(
            contract_type_filter=type_filter,
            search_query=search_q,
            status_filter=status_filter,
        )
        st.caption(f"총 **{len(rows)}건** 저장됨 (`data/contracts.db`)")

        if not rows:
            st.info("저장된 계약서가 없습니다. 상단의 [임시저장] 또는 [최종 누적 저장]을 이용해 보관하세요.")
        else:
            active_id = st.session_state.get("current_contract_id", "")
            for row in rows:
                cid = row["id"]
                is_active = cid == active_id
                status_badge = "🟢 완료" if row["status"] == "completed" else "🟡 임시저장"
                att_cnt = int(row.get("attachment_count", 0) or 0)
                att_badge = f" · 📎 {att_cnt}" if att_cnt > 0 else ""
                amt = int(row.get("deposit_amount", 0) or 0)
                amt_txt = f"{amt:,}원" if amt > 0 else "금액 미입력"

                with st.container(border=True):
                    active_mark = " 👉 **[편집중]**" if is_active else ""
                    st.markdown(f"**{status_badge}** · `{row['contract_type_label']}`{att_badge}{active_mark}")
                    st.markdown(f"{row['title']}")
                    st.caption(f"{amt_txt} | 수정: {row['updated_at'][:16]}")

                    b1, b2, b3 = st.columns([1.2, 1.2, 0.9])
                    with b1:
                        if st.button("열기", key=f"sb_open_{cid}", use_container_width=True):
                            meta_row, payload, _ = load_contract(cid)
                            if meta_row and payload:
                                apply_loaded_payload(
                                    contract_id=cid,
                                    contract_type=meta_row["contract_type"],
                                    status=meta_row["status"],
                                    payload=payload,
                                )
                                st.toast(f"계약서를 불러왔습니다: {meta_row['title']}")
                                st.rerun()
                    with b2:
                        if st.button("복사", key=f"sb_copy_{cid}", use_container_width=True, help="복사하여 새 계약 작성"):
                            meta_row, payload, _ = load_contract(cid)
                            if meta_row and payload:
                                new_cid = generate_contract_id()
                                apply_loaded_payload(
                                    contract_id=new_cid,
                                    contract_type=meta_row["contract_type"],
                                    status="unsaved",
                                    payload=payload,
                                )
                                st.toast(f"복사하여 새 계약({new_cid})을 시작했습니다.")
                                st.rerun()
                    with b3:
                        if st.button("삭제", key=f"sb_del_{cid}", use_container_width=True):
                            delete_contract(cid)
                            if cid == active_id:
                                start_new_contract(keep_type=True)
                            st.toast("저장된 계약서 및 증빙서류를 삭제했습니다.")
                            st.rerun()

        # 사이드바 최하단: Gemini API 키 설정 접이식 메뉴
        st.markdown("---")
        current_key = get_gemini_api_key(st.secrets)
        key_status_badge = f"활성화됨 ({mask_api_key(current_key)})" if current_key else "미설정"
        with st.expander(f"🔑 Gemini API 키 설정 ({key_status_badge})", expanded=False):
            if current_key:
                st.success(f"✅ 현재 설정됨: `{mask_api_key(current_key)}`")
            else:
                st.caption("Google AI Studio에서 발급받은 Gemini API 키를 붙여넣고 [저장]을 누르세요.")

            key_input = st.text_input(
                "Gemini API 키",
                type="password",
                placeholder="AIza...",
                key="sb_gemini_api_key_input",
            )
            k1, k2 = st.columns([1.4, 1.0])
            with k1:
                if st.button("💾 저장", key="sb_save_gemini_key_btn", type="primary", use_container_width=True):
                    ok_save, save_msg = save_gemini_api_key_to_secrets(key_input)
                    if ok_save:
                        # 이전 NO_API_KEY 오류 배너가 남아있으면 즉시 제거
                        last_ai = st.session_state.get("ai_last_result")
                        if isinstance(last_ai, dict) and last_ai.get("error_type") == "NO_API_KEY":
                            st.session_state["ai_last_result"] = None
                        st.session_state["sb_gemini_key_feedback"] = ("success", save_msg)
                        st.rerun()
                    else:
                        st.session_state["sb_gemini_key_feedback"] = ("error", save_msg)
                        st.rerun()
            with k2:
                if st.button(
                    "🗑️ 삭제",
                    key="sb_del_gemini_key_btn",
                    disabled=not current_key,
                    use_container_width=True,
                ):
                    delete_gemini_api_key_from_secrets()
                    st.session_state["sb_gemini_key_feedback"] = ("info", "저장된 Gemini API 키가 삭제되었습니다.")
                    st.rerun()

            fb = st.session_state.pop("sb_gemini_key_feedback", None)
            if fb:
                lvl, fb_msg = fb
                if lvl == "success":
                    st.success(fb_msg)
                elif lvl == "error":
                    st.error(fb_msg)
                else:
                    st.info(fb_msg)
            st.caption("저장 경로: `.streamlit/secrets.toml` (PC 로컬 전용)")


def render_attachments_and_ai_panel(
    contract_id: str,
    contract_type: str,
    type_cfg: dict[str, Any],
    live_form: dict[str, Any],
    p1_role: str,
    p2_role: str,
) -> None:
    """입력 폼 최상단: 증빙서류 카메라 촬영·갤러리 업로드 보관 및 AI 자동 채우기."""
    existing_atts = list_attachments(contract_id)
    panel_title = (
        f"📎 증빙서류 보관 및 AI 자동 채우기 (보관 중: {len(existing_atts)}건)"
        if existing_atts
        else "📎 증빙서류 보관 및 AI 자동 채우기 (신분증·건축물대장·등기부등본)"
    )

    with st.expander(panel_title, expanded=bool(st.session_state.get("_expand_att_section", False))):
        c1, c2 = st.columns([0.9, 1.6])
        with c1:
            att_cat = st.selectbox(
                "증빙서류 분류",
                options=ATTACHMENT_CATEGORIES,
                key="att_category_select",
            )
        with c2:
            id_role = st.radio(
                "신분증 적용 대상",
                options=["party1", "party2"],
                format_func=lambda r: (
                    f"매도인·임대인·양도인 ({p1_role})"
                    if r == "party1"
                    else f"매수인·임차인·양수인 ({p2_role})"
                ),
                horizontal=True,
                key="ai_id_role",
            )

        tab_upload, tab_cam = st.tabs(["🖼️ 갤러리 / 파일 업로드 (JPG·PNG·PDF)", "📷 카메라 직접 촬영"])

        with tab_upload:
            uploaded_files = st.file_uploader(
                "증빙서류 파일 선택 (다중 선택 가능)",
                type=["jpg", "jpeg", "png", "pdf"],
                accept_multiple_files=True,
                key=f"att_uploader_{contract_id}_{len(existing_atts)}",
            )
            if st.button(
                "📥 선택 파일 보관함에 저장",
                disabled=not uploaded_files,
                use_container_width=True,
                key="save_uploaded_atts_btn",
            ):
                payload = _build_storage_payload(live_form, contract_type, type_cfg)
                ensure_contract_record_exists(contract_id, contract_type, type_cfg["label"], payload)
                if st.session_state.get("current_contract_status") == "unsaved":
                    st.session_state["current_contract_status"] = "draft"
                for uf in uploaded_files or []:
                    mime = uf.type or (
                        "application/pdf" if uf.name.lower().endswith(".pdf") else "image/jpeg"
                    )
                    save_attachment(contract_id, att_cat, uf.name, uf.getvalue(), mime)
                st.session_state["_expand_att_section"] = True
                st.success(f"{len(uploaded_files or [])}건의 증빙서류가 `data/attachments/{contract_id}/`에 보관되었습니다.")
                st.rerun()

        with tab_cam:
            captured_photos: list[dict[str, Any]] = st.session_state.setdefault("captured_photos", [])
            cam_step: int = int(st.session_state.get("cam_step", 0))
            camera_open: bool = bool(st.session_state.get("camera_mode_open", True))
            cam_photo = None

            just_saved_msg = st.session_state.get("_cam_just_saved_msg", "")
            if just_saved_msg and not camera_open:
                st.success(just_saved_msg)

            if not camera_open:
                if st.button(
                    "📷 카메라 촬영 화면 다시 열기",
                    use_container_width=True,
                    type="primary",
                    key="reopen_cam_studio_btn",
                ):
                    st.session_state["camera_mode_open"] = True
                    st.session_state["_cam_just_saved_msg"] = ""
                    st.session_state["_expand_att_section"] = True
                    st.rerun()
            else:
                with st.container(key="camera_studio"):
                    # 상단 큼직한 촬영 장수 배지
                    st.markdown(
                        f'<div class="cam-count-badge">📷 현재 촬영된 서류: '
                        f'<span class="cam-count-pill">{len(captured_photos)}장</span></div>',
                        unsafe_allow_html=True,
                    )

                    # 촬영된 사진 썸네일 + 삭제(X) 버튼 (최대 4장씩 가로 정렬)
                    if captured_photos:
                        with st.container(key="cam_thumb_strip"):
                            per_row = 4
                            for row_start in range(0, len(captured_photos), per_row):
                                row_items = captured_photos[row_start : row_start + per_row]
                                t_cols = st.columns(per_row, gap="small")
                                for offset, item in enumerate(row_items):
                                    abs_idx = row_start + offset
                                    with t_cols[offset]:
                                        st.image(item["bytes"], use_container_width=True)
                                        if st.button(
                                            f"✕ #{abs_idx + 1}",
                                            key=f"del_cap_photo_{item['id']}",
                                            use_container_width=True,
                                        ):
                                            st.session_state["captured_photos"] = [
                                                p for p in captured_photos if p["id"] != item["id"]
                                            ]
                                            st.session_state["_expand_att_section"] = True
                                            st.rerun()

                    # 세로로 크고 군더더기 없는 카메라 뷰파인더
                    cam_photo = st.camera_input(
                        "서류 촬영",
                        key=f"cam_{cam_step}",
                        label_visibility="collapsed",
                    )

                    total_ready_photos = len(captured_photos) + (1 if cam_photo is not None else 0)

                    # 하단 단순·큼직한 버튼 2개
                    with st.container(key="cam_add_btn_wrap"):
                        if st.button(
                            "➕ 사진 추가 / 연속 촬영",
                            disabled=cam_photo is None,
                            use_container_width=True,
                            key="cam_add_next_btn",
                        ):
                            if cam_photo is not None:
                                next_num = len(captured_photos) + 1
                                captured_photos.append(
                                    {
                                        "id": f"cap_{cam_step}_{next_num}",
                                        "bytes": cam_photo.getvalue(),
                                        "mime": cam_photo.type or "image/jpeg",
                                        "category": att_cat,
                                        "filename": f"camera_{att_cat}_{next_num}.jpg",
                                    }
                                )
                                st.session_state["captured_photos"] = captured_photos
                                st.session_state["cam_step"] = cam_step + 1
                                st.session_state["_cam_just_saved_msg"] = ""
                                st.session_state["_expand_att_section"] = True
                                st.rerun()

                    with st.container(key="cam_finish_btn_wrap"):
                        if st.button(
                            f"✅ 촬영 마무리 및 저장 (총 {total_ready_photos}장)",
                            disabled=total_ready_photos == 0,
                            use_container_width=True,
                            key="cam_finish_save_btn",
                        ):
                            photos_to_save = list(captured_photos)
                            if cam_photo is not None:
                                next_num = len(photos_to_save) + 1
                                photos_to_save.append(
                                    {
                                        "id": f"cap_{cam_step}_{next_num}",
                                        "bytes": cam_photo.getvalue(),
                                        "mime": cam_photo.type or "image/jpeg",
                                        "category": att_cat,
                                        "filename": f"camera_{att_cat}_{next_num}.jpg",
                                    }
                                )
                            if photos_to_save:
                                payload = _build_storage_payload(live_form, contract_type, type_cfg)
                                ensure_contract_record_exists(contract_id, contract_type, type_cfg["label"], payload)
                                if st.session_state.get("current_contract_status") == "unsaved":
                                    st.session_state["current_contract_status"] = "draft"
                                for p_item in photos_to_save:
                                    save_attachment(
                                        contract_id,
                                        p_item["category"],
                                        p_item["filename"],
                                        p_item["bytes"],
                                        p_item["mime"],
                                    )
                                st.session_state["captured_photos"] = []
                                st.session_state["cam_step"] = cam_step + 1
                                st.session_state["camera_mode_open"] = False
                                st.session_state["_cam_just_saved_msg"] = (
                                    f"✅ 촬영한 증빙서류 총 {len(photos_to_save)}장이 보관함에 저장되었습니다. "
                                    "아래 [✨ AI로 서류 읽어 자동 채우기]를 눌러 바로 반영하세요."
                                )
                                st.session_state["_expand_att_section"] = True
                                st.rerun()

        # 보관된 증빙서류 열람 및 다운로드
        if existing_atts:
            st.markdown("---")
            st.markdown(f"**🗂️ 보관된 증빙서류 ({len(existing_atts)}건)**")
            for att in existing_atts:
                att_id = att["id"]
                raw_bytes = read_attachment_bytes(att["stored_path"])
                row_c1, row_c2, row_c3 = st.columns([2.4, 1.0, 0.7])
                with row_c1:
                    kb = max(1, int(att["file_size"]) // 1024)
                    st.markdown(f"`[{att['category']}]` **{att['original_filename']}** ({kb:,} KB)")
                    if raw_bytes and att["mime_type"].startswith("image/"):
                        with st.expander("미리보기", expanded=False):
                            st.image(raw_bytes, use_container_width=True)
                with row_c2:
                    if raw_bytes:
                        st.download_button(
                            "⬇️ 열기/저장",
                            data=raw_bytes,
                            file_name=att["original_filename"],
                            mime=att["mime_type"],
                            key=f"att_dl_{att_id}",
                            use_container_width=True,
                        )
                with row_c3:
                    if st.button("삭제", key=f"att_del_{att_id}", use_container_width=True):
                        delete_attachment(att_id)
                        st.session_state["_expand_att_section"] = True
                        st.rerun()

        # AI 자동 채우기 (Gemini API 연동 — 미저장 업로드/촬영 파일도 1클릭 자동 포함·보관)
        st.markdown("---")
        pending_captured = list(st.session_state.get("captured_photos", []))
        pending_upload_cnt = len(uploaded_files or []) + len(pending_captured) + (1 if cam_photo is not None else 0)
        total_available = len(existing_atts) + pending_upload_cnt
        api_key = get_gemini_api_key(st.secrets)

        if not api_key:
            st.warning(
                "🔑 **GEMINI_API_KEY 미설정**: 좌측 사이드바 하단의 **`🔑 Gemini API 키 설정`**에 "
                "API 키를 붙여넣고 **[💾 저장]**을 누르시면 즉시 AI 자동 채우기가 활성화됩니다."
            )

        if st.button(
            f"✨ AI로 서류 읽어 자동 채우기 (대상: {p1_role if id_role == 'party1' else p2_role} / 서류 {total_available}건)",
            disabled=total_available == 0,
            type="primary",
            use_container_width=True,
            key="ai_autofill_btn",
        ):
            # 아직 보관함에 저장하지 않은 업로드/연속촬영 파일이 있으면 먼저 자동 보관
            if pending_upload_cnt > 0:
                payload = _build_storage_payload(live_form, contract_type, type_cfg)
                ensure_contract_record_exists(contract_id, contract_type, type_cfg["label"], payload)
                if st.session_state.get("current_contract_status") == "unsaved":
                    st.session_state["current_contract_status"] = "draft"
                for uf in uploaded_files or []:
                    mime = uf.type or (
                        "application/pdf" if uf.name.lower().endswith(".pdf") else "image/jpeg"
                    )
                    save_attachment(contract_id, att_cat, uf.name, uf.getvalue(), mime)
                for p_item in pending_captured:
                    save_attachment(
                        contract_id,
                        p_item["category"],
                        p_item["filename"],
                        p_item["bytes"],
                        p_item["mime"],
                    )
                if cam_photo is not None:
                    save_attachment(
                        contract_id,
                        att_cat,
                        f"camera_{att_cat}_{len(pending_captured) + 1}.jpg",
                        cam_photo.getvalue(),
                        cam_photo.type or "image/jpeg",
                    )
                if pending_captured or cam_photo is not None:
                    st.session_state["captured_photos"] = []
                    st.session_state["cam_step"] = int(st.session_state.get("cam_step", 0)) + 1
                    st.session_state["camera_mode_open"] = False
                existing_atts = list_attachments(contract_id)

            docs_for_ai = []
            for att in existing_atts:
                b = read_attachment_bytes(att["stored_path"])
                if b:
                    docs_for_ai.append(
                        {
                            "category": att["category"],
                            "filename": att["original_filename"],
                            "mime_type": att["mime_type"],
                            "data": b,
                        }
                    )

            with st.status("🤖 AI 증빙서류 자동 채우기 실행 중...", expanded=True) as status_box:
                def _on_ai_step(msg: str) -> None:
                    status_box.write(msg)
                    status_box.update(label=msg)

                ai_res = extract_contract_fields_from_documents(
                    documents=docs_for_ai,
                    api_key=api_key,
                    id_target_role=id_role,
                    timeout_seconds=60,
                    on_progress=_on_ai_step,
                )
                if ai_res.get("ok"):
                    status_box.update(label="✅ 1) 이미지 보정 → 2) API 전송 → 3) 응답 파싱 완료!", state="complete")
                else:
                    status_box.update(
                        label=f"🚨 중단됨 [{ai_res.get('error_type', 'ERROR')}]",
                        state="error",
                    )
            _queue_pending_autofill(ai_res, contract_type)
            st.rerun()

        # 최근 AI 자동 채우기 진단 결과 투명 표시 (단계 로그·성공/실패·원본 응답 모두 유지)
        last_ai = st.session_state.get("ai_last_result")
        if isinstance(last_ai, dict):
            if last_ai.get("ok"):
                applied_keys = last_ai.get("applied_keys") or list(last_ai.get("mapped_fields", {}).keys())
                pretty_names = ", ".join(
                    _field_display_name(k, p1_role, p2_role) for k in applied_keys
                )
                st.success(f"✅ {last_ai.get('message', '')}\n\n**반영 항목**: {pretty_names}")
                with st.expander("🔍 AI 판독 단계 로그 · 원본 JSON · 이미지 보정 내역 보기", expanded=False):
                    if last_ai.get("step_log"):
                        st.markdown("\n".join(f"- {s}" for s in last_ai["step_log"]))
                    if last_ai.get("preprocess_notes"):
                        st.caption("이미지 전처리: " + " | ".join(last_ai["preprocess_notes"]))
                    st.json(last_ai.get("mapped_fields", {}))
                    st.code(last_ai.get("raw_response_text", ""), language="json")
            else:
                err_code = last_ai.get("error_type", "ERROR")
                st.error(f"🚨 AI 자동 채우기 실패 [{err_code}]: {last_ai.get('message', '알 수 없는 오류')}")
                if last_ai.get("step_log"):
                    st.markdown("**진행 단계 로그:**\n" + "\n".join(f"- {s}" for s in last_ai["step_log"]))
                if last_ai.get("preprocess_notes"):
                    st.caption("이미지 전처리 내역: " + " | ".join(last_ai["preprocess_notes"]))
                raw_txt = last_ai.get("raw_response_text", "")
                if raw_txt:
                    st.markdown("**🔍 AI 원본 응답 / 오류 상세:**")
                    st.code(raw_txt, language="json")


CLAUSE_FILTER_GROUPS = ["전체", "필수 기본", "대출·융자", "하자·시설", "세금·비용", "행정·승계"]


def _classify_clause_group(clause: dict[str, Any]) -> str:
    """특약 메타데이터(category/title/required)를 상황별 표준 5대 카테고리로 분류한다."""
    raw = f"{clause.get('category', '')} {clause.get('title', '')}"
    if any(k in raw for k in ("근저당", "융자", "대출", "보증보험", "에스크로", "유보", "권리변동", "권리말소")):
        return "대출·융자"
    if any(k in raw for k in ("하자", "시설", "수선", "지상물", "분묘", "매설", "경작", "농작물", "비품", "생활", "원상복구", "위반건축물")):
        return "하자·시설"
    if any(k in raw for k in ("세금", "부가세", "포괄양수도", "관리비", "정산", "체납", "보수", "비용", "직불금", "인건비")):
        return "세금·비용"
    if any(k in raw for k in ("농취증", "인허가", "허가", "승계", "명도", "임대차", "지정", "행정", "도로", "맹지", "경계", "측량", "대항력", "전입", "영업", "경업")):
        return "행정·승계"
    if clause.get("required") or any(k in raw for k in ("기본", "현황", "법령")):
        return "필수 기본"
    return "행정·승계"


def render_clause_selector(
    contract_type: str,
    type_cfg: dict[str, Any],
    live_form: dict[str, Any],
) -> list[str]:
    """③ 전문가용 정밀 특약사항 시스템 (상황별 카테고리 필터, 스마트 변수 자동 치환, AI 법률 특약 생성, 순서 이동 ▲/▼ 및 개별 삭제)."""
    nonce = st.session_state["_form_nonce"]
    overrides: dict[str, Any] = st.session_state.get("_clause_overrides", {})
    clauses_def: list[dict[str, Any]] = type_cfg.get("clauses", [])
    meta_info = get_type_meta(contract_type)
    p1_role = meta_info["party1"]
    p2_role = meta_info["party2"]

    # 직전 실행에서 예약된 위젯 초기화 / 체크 해제 안전 처리 (위젯 생성 전 실행)
    for k_clear in st.session_state.pop("_pending_clear_keys", []):
        st.session_state[k_clear] = ""
    for k_uncheck in st.session_state.pop("_pending_uncheck_keys", []):
        st.session_state[k_uncheck] = False

    st.caption(
        "💡 스마트 변수 자동 치환 지원: `{잔금일}`, `{계약일}`, `{인도일}`, `{매매대금}`, `{보증금}`, `{계약금}`, `{잔금}` 등이 본문 입력값과 실시간 연동됩니다."
    )

    # 1) 모든 프리셋 특약의 세션 키 초기화 및 카테고리 분류
    preset_by_id: dict[str, dict[str, Any]] = {}
    available_groups: list[str] = ["전체"]
    for clause in clauses_def:
        cid = clause["id"]
        preset_by_id[cid] = clause
        grp = _classify_clause_group(clause)
        if grp not in available_groups:
            available_groups.append(grp)
        saved_c = overrides.get(cid, {})
        chk_key = f"chk_{nonce}_{contract_type}_{cid}"
        txt_key = f"txt_{nonce}_{contract_type}_{cid}"
        if chk_key not in st.session_state:
            st.session_state[chk_key] = bool(saved_c.get("checked", clause.get("required", False)))
        if txt_key not in st.session_state:
            st.session_state[txt_key] = str(saved_c.get("template", clause["template"]))

    # 카테고리 정렬 순서 유지
    ordered_groups = [g for g in CLAUSE_FILTER_GROUPS if g in available_groups]
    filter_key = f"clause_cat_filter_{contract_type}"
    if st.session_state.get(filter_key) not in ordered_groups:
        st.session_state[filter_key] = "전체"

    selected_group = st.radio(
        "📂 상황별 특약 카테고리 필터",
        options=ordered_groups,
        horizontal=True,
        key=filter_key,
    )

    # 2) 필터링된 체크박스 라이브러리 렌더링 + 전체 체크 항목 텍스트 실시간 해석
    resolved_presets: dict[str, str] = {}
    for clause in clauses_def:
        cid = clause["id"]
        saved_c = overrides.get(cid, {})
        chk_key = f"chk_{nonce}_{contract_type}_{cid}"
        txt_key = f"txt_{nonce}_{contract_type}_{cid}"
        grp = _classify_clause_group(clause)
        show_here = selected_group == "전체" or grp == selected_group

        if show_here:
            badge = "🔴 필수" if clause.get("required") else "🔵 선택"
            label = f"{badge} [{grp}] {clause['title']}"
            checked = st.checkbox(label, key=chk_key)
        else:
            checked = bool(st.session_state.get(chk_key, False))

        if checked:
            var_meta: dict[str, dict[str, Any]] = clause.get("variables", {})
            saved_vars: dict[str, Any] = saved_c.get("variables", {})
            var_values: dict[str, Any] = {}

            if var_meta:
                if show_here:
                    var_cols = st.columns(min(len(var_meta), 2))
                for idx, (var_name, vmeta) in enumerate(var_meta.items()):
                    vkey = f"var_{nonce}_{contract_type}_{cid}_{var_name}"
                    bind_field = vmeta.get("bind_field")
                    default_val = saved_vars.get(var_name, vmeta.get("default", ""))

                    if bind_field and bind_field in live_form and var_name not in saved_vars:
                        bound_val = live_form[bind_field]
                        if vmeta.get("type") == "date":
                            default_val = format_date_kr(bound_val, zero_pad=True)
                        else:
                            default_val = str(bound_val) if bound_val else ""

                    if vkey not in st.session_state:
                        st.session_state[vkey] = default_val
                    elif bind_field and vmeta.get("type") == "date":
                        bound_str = format_date_kr(live_form.get(bind_field), zero_pad=True)
                        prev_bind_key = f"_prev_bind_{vkey}"
                        if st.session_state.get(prev_bind_key) != bound_str:
                            st.session_state[vkey] = bound_str
                            st.session_state[prev_bind_key] = bound_str

                    if show_here:
                        with var_cols[idx % len(var_cols)]:
                            var_values[var_name] = st.text_input(
                                vmeta.get("label", var_name),
                                key=vkey,
                                placeholder=f"{{{var_name}}}",
                            )
                    else:
                        var_values[var_name] = st.session_state.get(vkey, default_val)

            if show_here:
                with st.expander("✏️ 특약 문구 직접 수정", expanded=False):
                    edited_template = st.text_area(
                        "특약 원문 ({변수명} 유지 가능)",
                        key=txt_key,
                        height=78,
                        label_visibility="collapsed",
                    )
            else:
                edited_template = str(st.session_state.get(txt_key, clause["template"]))

            resolved = substitute_clause_variables(edited_template, var_values, var_meta)
            resolved = resolve_smart_clause_placeholders(resolved, live_form)
            resolved_presets[cid] = resolved

    # 3) ✨ AI 맞춤 법률 특약 생성기
    st.markdown("---")
    st.markdown("**✨ AI 맞춤 법률 특약 생성기**")
    ai_note_key = f"ai_clause_note_{nonce}_{contract_type}"
    ai_note_text = st.text_input(
        "현장 합의 내용 (구어체 입력)",
        key=ai_note_key,
        placeholder="예: 잔금 전에 매도인이 도배·장판 새로 해주고, 입주 후 6개월 내 누수 생기면 수리비 부담",
        label_visibility="collapsed",
    )
    custom_map: dict[str, list[str]] = st.session_state["custom_clauses_by_type"]
    custom_list: list[str] = custom_map.setdefault(contract_type, [])

    if st.button(
        "✨ AI 법률 특약으로 변환·추가",
        key=f"ai_clause_gen_btn_{nonce}_{contract_type}",
        type="primary",
        use_container_width=True,
    ):
        api_key = get_gemini_api_key(st.secrets)
        with st.spinner("🤖 AI가 현장 합의를 정식 법률 특약 문구로 다듬고 있습니다..."):
            gen_res = generate_legal_clauses_from_text(
                user_note=ai_note_text,
                contract_type_label=type_cfg["label"],
                api_key=api_key,
                party1_label=p1_role,
                party2_label=p2_role,
                timeout_seconds=60,
            )
        if gen_res.get("ok"):
            for new_c in gen_res.get("clauses", []):
                custom_list.append(new_c)
            st.session_state["_pending_clear_keys"] = [ai_note_key]
            st.session_state["_ai_clause_feedback"] = ("success", gen_res.get("message", "AI 특약이 추가되었습니다."))
            st.rerun()
        else:
            st.session_state["_ai_clause_feedback"] = ("error", gen_res.get("message", "AI 특약 생성 실패"))
            st.rerun()

    ai_fb = st.session_state.get("_ai_clause_feedback")
    if isinstance(ai_fb, tuple) and len(ai_fb) == 2:
        if ai_fb[0] == "success":
            st.success(ai_fb[1])
        else:
            st.error(ai_fb[1])

    # 4) 수동 사용자 정의 특약 직접 추가
    st.markdown("**➕ 사용자 정의 특약 직접 추가**")
    c_in, c_btn = st.columns([4, 1])
    new_clause_key = f"new_custom_clause_{nonce}_{contract_type}"
    with c_in:
        new_clause_text = st.text_input(
            "새 특약 내용 입력",
            key=new_clause_key,
            placeholder="예: 잔금일({잔금일}) 전까지 도배 및 장판을 신규 시공한다.",
            label_visibility="collapsed",
        )
    with c_btn:
        if st.button("추가", key=f"add_btn_{nonce}_{contract_type}", use_container_width=True):
            if new_clause_text.strip():
                custom_list.append(new_clause_text.strip())
                st.session_state["_pending_clear_keys"] = [new_clause_key]
                st.rerun()

    # 5) 적용된 특약 통합 순서 관리 (▲ 위로 / ▼ 아래로 / 🗑️ 개별 삭제)
    order_map: dict[str, list[str]] = st.session_state.setdefault("clause_order_by_type", {})
    order_tokens: list[str] = list(order_map.get(contract_type, []))

    valid_tokens: list[str] = [f"p:{cid}" for cid in resolved_presets] + [
        f"c:{i}" for i, txt in enumerate(custom_list) if str(txt).strip()
    ]
    valid_set = set(valid_tokens)
    # 기존 순서를 최대한 보존하되, 제거된 항목은 빼고 새로 추가된 항목은 뒤에 추가
    order_tokens = [tok for tok in order_tokens if tok in valid_set]
    for tok in valid_tokens:
        if tok not in order_tokens:
            order_tokens.append(tok)
    order_map[contract_type] = order_tokens

    rendered_list: list[str] = []
    if order_tokens:
        st.markdown("---")
        annex_hint = (
            f" · 📄 **{len(order_tokens)}항 (5항 초과 → 2페이지 [별지] 자동 분리)**"
            if len(order_tokens) > MAX_INLINE_CLAUSE_COUNT
            else f" · A4 1페이지 내 표시 (5항 초과 시 별지 자동 분리)"
        )
        st.markdown(f"**📋 적용된 특약 순서 관리 (총 {len(order_tokens)}항)**")
        st.caption(annex_hint.lstrip(" · "))

        with st.container(key="clause_order_list"):
            for pos, tok in enumerate(order_tokens):
                kind, raw_id = tok.split(":", 1)
                if kind == "p":
                    raw_text = resolved_presets.get(raw_id, "")
                    tag_badge = "표준"
                else:
                    c_idx = int(raw_id)
                    raw_text = resolve_smart_clause_placeholders(str(custom_list[c_idx]), live_form)
                    tag_badge = "맞춤/AI"

                rendered_list.append(raw_text)
                oc_txt, oc_up, oc_dn, oc_del = st.columns([5.8, 0.75, 0.75, 0.85], gap="small")
                with oc_txt:
                    st.markdown(
                        f'<div class="clause-order-card"><strong>{pos + 1}항 [{tag_badge}]</strong> {raw_text}</div>',
                        unsafe_allow_html=True,
                    )
                with oc_up:
                    if st.button(
                        "▲",
                        key=f"cl_up_{nonce}_{contract_type}_{pos}_{tok}",
                        disabled=(pos == 0),
                        use_container_width=True,
                        help="위로 이동",
                    ):
                        order_tokens[pos - 1], order_tokens[pos] = order_tokens[pos], order_tokens[pos - 1]
                        order_map[contract_type] = order_tokens
                        st.rerun()
                with oc_dn:
                    if st.button(
                        "▼",
                        key=f"cl_dn_{nonce}_{contract_type}_{pos}_{tok}",
                        disabled=(pos == len(order_tokens) - 1),
                        use_container_width=True,
                        help="아래로 이동",
                    ):
                        order_tokens[pos + 1], order_tokens[pos] = order_tokens[pos], order_tokens[pos + 1]
                        order_map[contract_type] = order_tokens
                        st.rerun()
                with oc_del:
                    if st.button(
                        "🗑️",
                        key=f"cl_del_{nonce}_{contract_type}_{pos}_{tok}",
                        use_container_width=True,
                        help="이 특약 삭제",
                    ):
                        if kind == "p":
                            chk_key = f"chk_{nonce}_{contract_type}_{raw_id}"
                            st.session_state.setdefault("_pending_uncheck_keys", []).append(chk_key)
                            order_tokens.pop(pos)
                            order_map[contract_type] = order_tokens
                        else:
                            del_idx = int(raw_id)
                            custom_list.pop(del_idx)
                            # 커스텀 인덱스 재정렬
                            new_order: list[str] = []
                            for t in order_tokens:
                                tk, ti = t.split(":", 1)
                                if tk == "p":
                                    new_order.append(t)
                                else:
                                    old_i = int(ti)
                                    if old_i < del_idx:
                                        new_order.append(f"c:{old_i}")
                                    elif old_i > del_idx:
                                        new_order.append(f"c:{old_i - 1}")
                            order_map[contract_type] = new_order
                        st.rerun()

    st.markdown("---")
    a1, a2 = st.columns(2)
    with a1:
        live_form["attachment_note"] = st.text_input(
            "첨부 서류 표기",
            key=_wkey("attachment_note"),
            placeholder="예: 확인설명서 및 공제증서",
        )
    with a2:
        live_form["deposit_account"] = st.text_input(
            "입금계좌",
            key=_wkey("deposit_account"),
            placeholder="예: 기업 010-0000-0000 홍길동",
        )

    return rendered_list


def main() -> None:
    inject_mobile_pinch_zoom_meta()
    configured_password = get_configured_password()
    if configured_password and not st.session_state.get("authenticated", False):
        render_password_lock_screen(configured_password)
        st.stop()
        return

    init_db()
    ensure_templates()
    ensure_state_initialized()
    _consume_pending_autofill()
    clauses_lib = load_clauses_library()

    # 좌측 사이드바: 저장된 계약서 목록
    render_sidebar_saved_contracts(clauses_lib)

    type_keys = get_ordered_type_keys(clauses_lib)
    type_labels = {k: clauses_lib[k]["label"] for k in type_keys}
    active_type_key = st.session_state.get("contract_type", "nursing_home_transfer")
    if active_type_key not in CONTRACT_TYPE_THEMES:
        active_type_key = "nursing_home_transfer"
    theme = CONTRACT_TYPE_THEMES[active_type_key]
    active_type_label = type_labels.get(active_type_key, "🏥 [요양시설] 요양원 양도·양수")
    active_cat_label = theme.get("category", "계약서")

    # 선택된 계약 유형 카테고리에 따른 전용 고급 선택 카드·드롭다운 동적 테마 CSS (19px 굵은 글씨 강조)
    st.markdown(
        f"""
        <style>
          .st-key-contract_type_card {{
            border: 3px solid {theme['primary']} !important;
            border-left: 7px solid {theme['primary']} !important;
            background: linear-gradient(135deg, {theme['soft_bg']} 0%, #ffffff 100%) !important;
            border-radius: 14px !important;
            padding: 0.7rem 0.9rem 0.8rem 0.9rem !important;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.12) !important;
          }}
          .st-key-contract_type_card .type-active-badge {{
            font-size: 0.9rem !important;
            font-weight: 900 !important;
            padding: 4px 12px !important;
            letter-spacing: 0.03em !important;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.14) !important;
          }}
          .st-key-contract_type_card .type-hint-badge {{
            font-size: 0.82rem !important;
            font-weight: 800 !important;
            padding: 3px 10px !important;
          }}
          .st-key-contract_type [data-baseweb="select"] > div {{
            border: 2.5px solid {theme['primary']} !important;
            border-radius: 11px !important;
            background-color: #ffffff !important;
            box-shadow: 0 3px 10px rgba(15, 23, 42, 0.1) !important;
            min-height: 54px !important;
            padding: 4px 8px 4px 10px !important;
            cursor: pointer !important;
            transition: all 0.16s ease !important;
          }}
          .st-key-contract_type [data-baseweb="select"] > div:hover {{
            border-color: {theme['hover']} !important;
            background-color: {theme['soft_bg']} !important;
            box-shadow: 0 5px 15px rgba(15, 23, 42, 0.18) !important;
          }}
          .st-key-contract_type [data-baseweb="select"] span,
          .st-key-contract_type [data-baseweb="select"] div[value],
          .st-key-contract_type [data-baseweb="select"] input {{
            font-size: 19px !important;
            font-weight: 900 !important;
            color: {theme['primary']} !important;
            letter-spacing: -0.01em !important;
          }}
          .st-key-contract_type [data-baseweb="select"] svg {{
            background-color: {theme['primary']} !important;
            color: #ffffff !important;
            fill: #ffffff !important;
            width: 34px !important;
            height: 34px !important;
            padding: 6px !important;
            border-radius: 9px !important;
            transition: background-color 0.16s ease, transform 0.16s ease !important;
          }}
          .st-key-contract_type [data-baseweb="select"] > div:hover svg {{
            background-color: {theme['hover']} !important;
            transform: scale(1.06) !important;
          }}
          @media (max-width: 768px) {{
            .st-key-contract_type [data-baseweb="select"] span,
            .st-key-contract_type [data-baseweb="select"] div[value],
            .st-key-contract_type [data-baseweb="select"] input {{
              font-size: 19px !important;
              font-weight: 900 !important;
            }}
            .st-key-contract_type [data-baseweb="select"] > div {{
              min-height: 54px !important;
            }}
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="app-header-bar">
          <div class="app-header-left">
            <div class="app-header-icon">📑</div>
            <div>
              <h1 class="app-header-title">통합계약서</h1>
              <p class="app-header-sub">요양시설 · 아파트·주택 · 상가 · 농지·토지 전문 계약서 작성기</p>
            </div>
          </div>
          <div class="app-header-pill" style="border-color:{theme['badge_bg']};">
            <span style="display:inline-block;width:11px;height:11px;border-radius:50%;background:{theme['primary']};border:2px solid #fff;"></span>
            <span>{active_type_label}</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ==================== 화면 글자 크기 조절 바 (모바일·PC 공용) ====================
    with st.container(key="font_size_bar"):
        ui_font_scale = st.radio(
            "🔤 글자 크기",
            options=["보통", "크게", "아주 크게"],
            horizontal=True,
            key="ui_font_scale",
        )
    inject_font_scale_css(ui_font_scale)

    # ==================== 상단 제어 바 (1행: 강화된 계약유형 카드 · 색상모드 · 빈 양식/초안 · 예시/초기화) ====================
    with st.container(key="top_controls_row"):
        ctrl1, ctrl2, ctrl3, ctrl4 = st.columns([2.75, 1.35, 1.45, 1.35], gap="small")
        with ctrl1:
            with st.container(key="contract_type_card"):
                st.markdown(
                    f"""
                    <div class="type-card-header">
                      <span class="type-active-badge" style="background:{theme['primary']};">{active_cat_label}</span>
                      <span class="type-hint-badge" style="background:{theme['badge_bg']}; color:{theme['primary']}; border:1.5px solid {theme['primary']};">▼ 계약 유형 변경</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                selected_type = st.selectbox(
                    "계약 유형 선택",
                    options=type_keys,
                    format_func=lambda k: type_labels[k],
                    key="contract_type",
                    label_visibility="collapsed",
                )
        with ctrl2:
            color_mode = st.radio(
                "출력 색상 모드",
                options=["color", "bw"],
                format_func=lambda m: "🎨 컬러" if m == "color" else "🖨️ 흑백",
                horizontal=True,
                key="color_mode",
            )
        with ctrl3:
            blank_form_mode = st.toggle(
                "📄 빈 양식 모드 (날짜·특약 비우기)",
                key="blank_form_mode",
                help="켜면 날짜를 '20  년  월  일' 빈칸으로 두고 기본 특약 문구도 모두 비워 수기용 완전 빈 양식으로 출력합니다.",
            )
            draft_mode = st.toggle(
                "초안 워터마크",
                key="draft_mode",
                help="미리보기 및 출력물 배경에 '사전 검토용 초안 (계약 효력 없음)' 표시",
            )
        with ctrl4:
            with st.container(key="quick_tools_row"):
                qt1, qt2 = st.columns(2, gap="small")
                with qt1:
                    if st.button("📋 예시", use_container_width=True, help="선택한 계약 유형의 예시 데이터를 채웁니다."):
                        apply_sample_to_form(selected_type)
                        st.rerun()
                with qt2:
                    if st.button("🧹 초기화", use_container_width=True, help="모든 입력값을 빈칸/0원으로 초기화하고 빈 양식 모드로 전환합니다."):
                        reset_form_to_empty()
                        st.rerun()

    type_cfg = clauses_lib[selected_type]
    meta = get_type_meta(selected_type)
    p1_role = meta["party1"]
    p2_role = meta["party2"]
    amt_name = meta["amount_label"]
    current_cid = st.session_state["current_contract_id"]
    current_status = st.session_state.get("current_contract_status", "unsaved")

    # 상단 제어 바 (2행: 문서 ID + 모바일에서도 가로 3등분 1줄 [🔒 잠금 | 💾 임시저장 | ✅ 최종 저장])
    status_label_map = {"unsaved": "⚪ 미저장 새 문서", "draft": "🟡 임시저장됨", "completed": "🟢 최종 저장 완료"}
    save_bar_left, save_bar_right = st.columns([2.2, 2.8], gap="small")
    with save_bar_left:
        st.markdown(
            f'<div class="doc-id-pill">문서 ID: {current_cid} &nbsp;|&nbsp; {status_label_map.get(current_status, "⚪ 미저장")}</div>',
            unsafe_allow_html=True,
        )
    with save_bar_right:
        with st.container(key="action_buttons_row"):
            act1, act2, act3 = st.columns(3, gap="small")
            with act1:
                st.button(
                    "🔒 잠금",
                    use_container_width=True,
                    disabled=not bool(configured_password),
                    on_click=lambda: st.session_state.__setitem__("authenticated", False),
                    help=(
                        "화면을 즉시 잠그고 비밀번호 입력 화면으로 돌아갑니다."
                        if configured_password
                        else "APP_PASSWORD가 설정되지 않은 로컬 환경에서는 비밀번호 잠금이 비활성화되어 있습니다."
                    ),
                )
            with act2:
                save_draft_clicked = st.button("💾 임시저장", use_container_width=True)
            with act3:
                save_final_clicked = st.button("✅ 최종 저장", use_container_width=True, type="primary")

    live_form: dict[str, Any] = dict(st.session_state["_form_data"])

    # ==================== 모바일 전용 [✏️ 계약서 입력 / 📄 계약서 크게 보기] 전환 바 ====================
    active_mob_tab = st.session_state.get("mobile_active_tab", "input")
    with st.container(key="mobile_view_switch"):
        mt1, mt2 = st.columns(2, gap="small")
        with mt1:
            if st.button(
                "✏️ 계약서 입력",
                key="btn_mob_tab_input",
                type="primary" if active_mob_tab == "input" else "secondary",
                use_container_width=True,
            ):
                st.session_state["mobile_active_tab"] = "input"
                st.rerun()
        with mt2:
            if st.button(
                "📄 계약서 크게 보기",
                key="btn_mob_tab_preview",
                type="primary" if active_mob_tab == "preview" else "secondary",
                use_container_width=True,
            ):
                st.session_state["mobile_active_tab"] = "preview"
                st.rerun()
    inject_mobile_tab_css(active_mob_tab)

    # ==================== 좌측 콤팩트 입력 폼(3.5) / 우측 대형 실시간 A4 미리보기(6.5) ====================
    workspace_box = st.container(key="workspace_split")
    with workspace_box:
        col_form, col_preview = st.columns([3.5, 6.5], gap="medium")

    with col_form, st.container(key="form_pane"):
        # 0. 최상단: 증빙서류 촬영·갤러리 업로드 보관 및 AI 자동 채우기
        render_attachments_and_ai_panel(
            contract_id=current_cid,
            contract_type=selected_type,
            type_cfg=type_cfg,
            live_form=live_form,
            p1_role=p1_role,
            p2_role=p2_role,
        )

        # ① 대상 부동산 / 토지 / 농지 / 아파트 / 상가 / 시설 / 점포의 표시
        sec1_label_map = {
            "nursing_home_transfer": "① 양도·양수 대상 요양시설의 표시",
            "apartment_sale": "① 아파트의 표시 (동·호수·대지권·면적)",
            "commercial_sale": "① 상가건물의 표시 (소재지·구조·호수·면적)",
            "commercial_lease": "① 상가건물의 표시",
            "premium_rights": "① 대상 점포 및 신규 임대차 조건",
            "farmland_sale": "① 농지의 표시 (지번·지목·면적·용도지역)",
            "land_sale": "① 토지의 표시 (지번·지목·면적·용도지역)",
        }
        sec1_label = sec1_label_map.get(selected_type, "① 부동산의 표시")
        with st.expander(sec1_label, expanded=True):
            if selected_type in ("land_sale", "farmland_sale"):
                is_farm = selected_type == "farmland_sale"
                live_form["address"] = st.text_input(
                    "농지 소재지 (지번)" if is_farm else "토지 소재지 (지번)",
                    key=_wkey("address"),
                    placeholder="예: 경기도 양평군 용문면 다문리 456-2" if is_farm else "예: 경기도 양평군 양서면 목왕리 123-4",
                )
                ld1, ld2 = st.columns(2)
                with ld1:
                    live_form["land_category"] = st.text_input(
                        "지목",
                        key=_wkey("land_category"),
                        placeholder="전, 답, 과수원 등" if is_farm else "전, 답, 대, 임야 등",
                    )
                with ld2:
                    live_form["land_area"] = st.text_input(
                        "공부상 면적 (㎡)", key=_wkey("land_area"), placeholder="예: 1000.0"
                    )
                    py_txt = format_pyeong(live_form["land_area"])
                    if py_txt:
                        st.markdown(
                            f'<div class="kr-amount-hint">약 {py_txt} (3.3058㎡ 기준 자동 환산)</div>',
                            unsafe_allow_html=True,
                        )

                ld3, ld4 = st.columns(2)
                with ld3:
                    live_form["lease_part"] = st.text_input(
                        "매매 대상 (전부 / 지분율)",
                        key=_wkey("lease_part"),
                        placeholder="예: 농지 전부 (또는 지분 1/2)" if is_farm else "예: 토지 전부 (또는 지분 1/2)",
                    )
                with ld4:
                    live_form["land_zoning"] = st.text_input(
                        "용도지역·지구",
                        key=_wkey("land_zoning"),
                        placeholder="예: 농림지역(농업진흥구역), 계획관리지역" if is_farm else "예: 계획관리지역, 농림지역",
                    )

                live_form["land_status"] = st.text_input(
                    "경작 현황" if is_farm else "토지 현황",
                    key=_wkey("land_status"),
                    placeholder="예: 논(경작 중), 밭, 비닐하우스 등" if is_farm else "예: 경작 중(전), 나대지, 임야 등",
                )

                live_form["extra_parcels"] = st.text_area(
                    "다필지(추가 필지) 목록 메모 (선택)",
                    key=_wkey("extra_parcels"),
                    height=68,
                    placeholder="메인 필지 외 추가 필지가 있으면 지번·지목·면적을 입력하세요. 입력 시 특약/별지에 자동 반영됩니다.",
                )

            elif selected_type == "nursing_home_transfer":
                n1, n2 = st.columns([1.4, 1.0])
                with n1:
                    live_form["facility_name"] = st.text_input(
                        "요양시설 명칭", key=_wkey("facility_name"), placeholder="예: 행복한실버노인요양원"
                    )
                with n2:
                    live_form["facility_reg_no"] = st.text_input(
                        "장기요양기관 기호", key=_wkey("facility_reg_no"), placeholder="예: 2-41360-00000"
                    )
                live_form["address"] = st.text_input(
                    "시설 소재지", key=_wkey("address"), placeholder="예: 경기도 남양주시 ..."
                )
                n3, n4 = st.columns([1.2, 1.2])
                with n3:
                    live_form["facility_capacity"] = st.text_input(
                        "인가 정원 / 현원", key=_wkey("facility_capacity"), placeholder="예: 인가정원 29인 / 현원 27인"
                    )
                with n4:
                    live_form["lease_area"] = st.text_input(
                        "시설 전용면적 (㎡)", key=_wkey("lease_area"), placeholder="예: 680.50"
                    )
                n5, n6 = st.columns(2)
                with n5:
                    live_form["building_structure"] = st.text_input(
                        "건물 구조", key=_wkey("building_structure"), placeholder="철근콘크리트"
                    )
                with n6:
                    live_form["building_use"] = st.text_input(
                        "건물 용도", key=_wkey("building_use"), placeholder="노유자시설(노인의료복지시설)"
                    )
                live_form["lease_part"] = st.text_input(
                    "양도 범위",
                    key=_wkey("lease_part"),
                    placeholder="예: 시설·비품·장기요양기관 지정 지위 및 운영권 일체",
                )

            elif selected_type == "premium_rights":
                pr1, pr2 = st.columns(2)
                with pr1:
                    live_form["store_name"] = st.text_input(
                        "상호 (점포명)", key=_wkey("store_name"), placeholder="예: 오남수제버거"
                    )
                with pr2:
                    live_form["business_type"] = st.text_input(
                        "영업 업종", key=_wkey("business_type"), placeholder="예: 일반음식점"
                    )
                live_form["address"] = st.text_input(
                    "점포 소재지", key=_wkey("address"), placeholder="예: 경기도 남양주시 ..."
                )
                pr3, pr4 = st.columns([1.6, 1.0])
                with pr3:
                    live_form["lease_part"] = st.text_input(
                        "점포 층·호수", key=_wkey("lease_part"), placeholder="예: 1층 105호 점포 전부"
                    )
                with pr4:
                    live_form["lease_area"] = st.text_input(
                        "점포 면적 (㎡)", key=_wkey("lease_area"), placeholder="예: 62.80"
                    )
                live_form["new_lease_summary"] = st.text_input(
                    "신규 임대차 조건 요약",
                    key=_wkey("new_lease_summary"),
                    placeholder="예: 보증금 4,000만원 / 월세 220만원(VAT별도) / 24개월",
                )

            elif selected_type == "commercial_lease":
                live_form["address"] = st.text_input(
                    "상가 소재지", key=_wkey("address"), placeholder="예: 경기도 구리시 ..."
                )
                c1, c2 = st.columns(2)
                with c1:
                    live_form["building_structure"] = st.text_input(
                        "건물 구조", key=_wkey("building_structure"), placeholder="철근콘크리트"
                    )
                with c2:
                    live_form["building_use"] = st.text_input(
                        "건축물 용도", key=_wkey("building_use"), placeholder="제1종근린생활시설"
                    )
                c3, c4 = st.columns(2)
                with c3:
                    live_form["building_area"] = st.text_input(
                        "건물 연면적 (㎡)", key=_wkey("building_area"), placeholder="예: 1250.40"
                    )
                with c4:
                    live_form["lease_area"] = st.text_input(
                        "계약 면적 (㎡)", key=_wkey("lease_area"), placeholder="예: 56.20"
                    )
                c5, c6 = st.columns(2)
                with c5:
                    live_form["lease_part"] = st.text_input(
                        "임대할 부분 (층·호수)", key=_wkey("lease_part"), placeholder="예: 1층 104호 전부"
                    )
                with c6:
                    live_form["business_type"] = st.text_input(
                        "임차인 영업 업종/용도", key=_wkey("business_type"), placeholder="예: 커피전문점(휴게음식점)"
                    )

            else:
                live_form["address"] = st.text_input(
                    "소재지 (도로명/지번 및 동·호수)",
                    key=_wkey("address"),
                    placeholder="예: 경기도 남양주시 오남읍 ...",
                )
                r1c1, r1c2 = st.columns(2)
                with r1c1:
                    live_form["land_category"] = st.text_input(
                        "지목", key=_wkey("land_category"), placeholder="예: 대"
                    )
                with r1c2:
                    live_form["land_share_ratio"] = st.text_input(
                        "대지권비율", key=_wkey("land_share_ratio"), placeholder="예: 33101/55.3748"
                    )

                r2c1, r2c2 = st.columns(2)
                with r2c1:
                    live_form["land_area"] = st.text_input(
                        "토지면적 (㎡)", key=_wkey("land_area"), placeholder="예: 32.01"
                    )
                with r2c2:
                    live_form["building_area"] = st.text_input(
                        "건물 면적 (㎡)", key=_wkey("building_area"), placeholder="예: 84.9580"
                    )

                r3c1, r3c2 = st.columns(2)
                with r3c1:
                    live_form["building_structure"] = st.text_input(
                        "건물 구조", key=_wkey("building_structure"), placeholder="예: 철근콘크리트"
                    )
                with r3c2:
                    live_form["building_use"] = st.text_input(
                        "건물 용도", key=_wkey("building_use"), placeholder="예: 공동주택"
                    )

                part_lbl = (
                    "매매대상 (동·호수)"
                    if selected_type == "apartment_sale"
                    else (
                        "매매대상 (층·호수)"
                        if selected_type == "commercial_sale"
                        else ("매매대상 부분" if selected_type == "real_estate_sale" else "임대할 부분")
                    )
                )
                area_lbl = (
                    "매매 면적 (㎡)"
                    if selected_type in ("apartment_sale", "real_estate_sale", "commercial_sale")
                    else "임대 면적 (㎡)"
                )
                r4c1, r4c2 = st.columns([1.4, 1.0])
                with r4c1:
                    live_form["lease_part"] = st.text_input(
                        part_lbl, key=_wkey("lease_part"), placeholder="예: 108동 401호 전부"
                    )
                with r4c2:
                    live_form["lease_area"] = st.text_input(
                        area_lbl, key=_wkey("lease_area"), placeholder="예: 84.9580"
                    )

        # ② 계약내용 및 금액
        with st.expander(f"② 계약내용 및 {amt_name}", expanded=True):
            a1, a2 = st.columns(2)
            with a1:
                deposit_val = st.number_input(
                    f"{amt_name}(원)", min_value=0, step=1_000_000, key=_wkey("deposit")
                )
                live_form["deposit"] = deposit_val
                py_unit_hint = (
                    f" · 평당 {format_price_per_pyeong(deposit_val, live_form.get('land_area'))}"
                    if selected_type in ("land_sale", "farmland_sale")
                    and format_price_per_pyeong(deposit_val, live_form.get("land_area"))
                    else ""
                )
                if deposit_val > 0:
                    st.markdown(
                        f'<div class="kr-amount-hint">금 {number_to_korean(deposit_val)}원정 ({format_won(deposit_val)}{py_unit_hint})</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        '<div class="kr-amount-hint" style="color:#64748b;">미입력 (출력 시 수기용 빈칸)</div>',
                        unsafe_allow_html=True,
                    )
            with a2:
                down_val = st.number_input(
                    "계약금(원)", min_value=0, step=1_000_000, key=_wkey("down_payment")
                )
                live_form["down_payment"] = down_val
                if down_val > 0:
                    st.markdown(
                        f'<div class="kr-amount-hint">금 {number_to_korean(down_val)}원정</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        '<div class="kr-amount-hint" style="color:#64748b;">미입력 (출력 시 수기용 빈칸)</div>',
                        unsafe_allow_html=True,
                    )

            loan_val = 0
            lease_dep_val = 0
            escrow_val = 0

            if selected_type in ("land_sale", "farmland_sale"):
                loan_val = st.number_input(
                    "융자금 승계 — 근저당권(원)",
                    min_value=0,
                    step=1_000_000,
                    key=_wkey("loan_assumption"),
                )
                live_form["loan_assumption"] = loan_val
                if loan_val > 0:
                    st.markdown(
                        f'<div class="kr-amount-hint">금 {number_to_korean(loan_val)}원정</div>',
                        unsafe_allow_html=True,
                    )

            if selected_type in ("real_estate_sale", "apartment_sale", "commercial_sale"):
                s1, s2 = st.columns(2)
                with s1:
                    loan_val = st.number_input(
                        "융자금 승계(원)", min_value=0, step=1_000_000, key=_wkey("loan_assumption")
                    )
                    live_form["loan_assumption"] = loan_val
                    if loan_val > 0:
                        st.markdown(
                            f'<div class="kr-amount-hint">금 {number_to_korean(loan_val)}원정</div>',
                            unsafe_allow_html=True,
                        )
                with s2:
                    lease_dep_val = st.number_input(
                        "임대차보증금 승계(원)",
                        min_value=0,
                        step=1_000_000,
                        key=_wkey("lease_deposit_assumption"),
                    )
                    live_form["lease_deposit_assumption"] = lease_dep_val
                    if lease_dep_val > 0:
                        st.markdown(
                            f'<div class="kr-amount-hint">금 {number_to_korean(lease_dep_val)}원정</div>',
                            unsafe_allow_html=True,
                        )
                if selected_type == "commercial_sale":
                    vat_options = ["사업포괄양수도(부가세 없음)", "부가가치세 별도", "부가가치세 포함", "면세/해당없음"]
                    vat_key = _wkey("vat_type")
                    if st.session_state.get(vat_key) not in vat_options:
                        st.session_state[vat_key] = vat_options[0]
                    live_form["vat_type"] = st.selectbox(
                        "건물분 부가가치세 구분",
                        options=vat_options,
                        key=vat_key,
                    )

            if selected_type == "nursing_home_transfer":
                e1, e2 = st.columns([1.3, 1.0])
                with e1:
                    escrow_val = st.number_input(
                        "유보금(에스크로)(원)", min_value=0, step=1_000_000, key=_wkey("escrow_amount")
                    )
                    live_form["escrow_amount"] = escrow_val
                    if escrow_val > 0:
                        st.markdown(
                            f'<div class="kr-amount-hint">금 {number_to_korean(escrow_val)}원정</div>',
                            unsafe_allow_html=True,
                        )
                with e2:
                    live_form["escrow_period"] = st.text_input(
                        "유보기간", key=_wkey("escrow_period"), placeholder="예: 6개월"
                    )
                live_form["escrow_condition"] = st.text_input(
                    "유보 정산 조건",
                    key=_wkey("escrow_condition"),
                    placeholder="현지조사·부당청구 환수·행정처분 대비",
                )

            a3, a4 = st.columns(2)
            with a3:
                interim_val = st.number_input(
                    "중도금(원)", min_value=0, step=1_000_000, key=_wkey("interim_payment")
                )
                live_form["interim_payment"] = interim_val
                if interim_val > 0:
                    st.markdown(
                        f'<div class="kr-amount-hint">금 {number_to_korean(interim_val)}원정</div>',
                        unsafe_allow_html=True,
                    )
            with a4:
                auto_bal = st.checkbox("잔금 자동 계산", key=_wkey("auto_balance"))
                live_form["auto_balance"] = auto_bal
                if auto_bal:
                    computed_bal = calculate_balance(
                        deposit=deposit_val,
                        down_payment=down_val,
                        interim_payment=interim_val,
                        loan_assumption=loan_val,
                        lease_deposit_assumption=lease_dep_val,
                        escrow_amount=escrow_val,
                    )
                    st.number_input(
                        "잔금(원) [자동계산]",
                        value=computed_bal,
                        disabled=True,
                        key=f"auto_bal_view_{st.session_state['_form_nonce']}_{selected_type}",
                    )
                    balance_val = computed_bal
                else:
                    balance_val = st.number_input(
                        "잔금(원) [수동입력]", step=1_000_000, key=_wkey("balance")
                    )
                live_form["balance"] = balance_val
                if balance_val > 0:
                    st.markdown(
                        f'<div class="kr-amount-hint">금 {number_to_korean(balance_val)}원정</div>',
                        unsafe_allow_html=True,
                    )
                elif balance_val < 0:
                    st.markdown(
                        f'<div class="kr-amount-hint" style="color:#b91c1c;">잔금 음수 ({balance_val:,}원)</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        '<div class="kr-amount-hint" style="color:#64748b;">미입력 (출력 시 수기용 빈칸)</div>',
                        unsafe_allow_html=True,
                    )

            if selected_type in ("housing_wolse", "commercial_lease"):
                st.markdown("---")
                m1, m2 = st.columns(2)
                with m1:
                    rent_val = st.number_input(
                        "월 차임(원)", min_value=0, step=50_000, key=_wkey("monthly_rent")
                    )
                    live_form["monthly_rent"] = rent_val
                    if rent_val > 0:
                        st.markdown(
                            f'<div class="kr-amount-hint">월 금 {number_to_korean(rent_val)}원정</div>',
                            unsafe_allow_html=True,
                        )
                with m2:
                    live_form["monthly_rent_day"] = st.text_input(
                        "매월 지급일", key=_wkey("monthly_rent_day"), placeholder="예: 25"
                    )

                m3, m4 = st.columns(2)
                with m3:
                    live_form["rent_pay_type"] = st.selectbox(
                        "선불/후불", options=["선불", "후불"], key=_wkey("rent_pay_type")
                    )
                with m4:
                    live_form["management_fee"] = st.number_input(
                        "월 관리비(원)", min_value=0, step=10_000, key=_wkey("management_fee")
                    )

                if selected_type == "commercial_lease":
                    mf1, mf2 = st.columns(2)
                    with mf1:
                        live_form["vat_type"] = st.selectbox(
                            "부가가치세 구분",
                            options=["부가가치세 별도", "부가가치세 포함", "면세/해당없음"],
                            key=_wkey("vat_type"),
                        )
                    with mf2:
                        live_form["management_fee_note"] = st.text_input(
                            "관리비 포함/별도 내역",
                            key=_wkey("management_fee_note"),
                            placeholder="예: 공과금 실비 정산, 부가세 별도",
                        )
                else:
                    live_form["management_fee_note"] = st.text_input(
                        "관리비 포함/별도 내역",
                        key=_wkey("management_fee_note"),
                        placeholder="예: 공용전기·청소 포함, 수도·가스 별도",
                    )

            st.markdown("---")
            d1, d2 = st.columns(2)
            with d1:
                live_form["down_receipt_person"] = st.text_input(
                    "계약금 영수자", key=_wkey("down_receipt_person"), placeholder=f"{p1_role} 성명"
                )
            with d2:
                live_form["balance_date"] = st.date_input("잔금 지급일", key=_wkey("balance_date"))

            d3, d4 = st.columns(2)
            with d3:
                use_interim_dt = st.checkbox("중도금 지급일 기입", key=_wkey("use_interim_date"))
                live_form["use_interim_date"] = use_interim_dt
                if use_interim_dt:
                    live_form["interim_date"] = st.date_input("중도금 지급일", key=_wkey("interim_date"))
            with d4:
                live_form["contract_date"] = st.date_input("계약 작성일", key=_wkey("contract_date"))

            if selected_type in ("housing_jeonse", "housing_wolse", "commercial_lease"):
                p1, p2 = st.columns(2)
                with p1:
                    live_form["handover_date"] = st.date_input("인도일 (존속 시작)", key=_wkey("handover_date"))
                with p2:
                    live_form["lease_end_date"] = st.date_input("임대차 만료일", key=_wkey("lease_end_date"))
                dur_str = calculate_lease_duration(live_form["handover_date"], live_form["lease_end_date"])
                if dur_str:
                    st.caption(f"⏱️ 자동 계산된 존속기간: **{dur_str}**")
            else:
                handover_lbl = (
                    "양도기준일 (인수인계일)"
                    if selected_type == "nursing_home_transfer"
                    else "인도일 (명도·이전일)"
                )
                live_form["handover_date"] = st.date_input(handover_lbl, key=_wkey("handover_date"))

        # ③ 전문가용 정밀 특약사항 시스템
        with st.expander("③ 정밀 특약사항 · AI 법률 특약 생성기", expanded=not blank_form_mode):
            if blank_form_mode:
                st.info(
                    "📄 현재 **[빈 양식 모드]**가 켜져 있어 특약사항이 깨끗한 수기용 빈칸으로 출력됩니다."
                )
                if st.button(
                    "📝 특약 작성 시작 (빈 양식 모드 끄기)",
                    key="disable_blank_for_clauses_btn",
                    use_container_width=True,
                ):
                    st.session_state["_pending_blank_form_mode"] = False
                    st.rerun()
                rendered_clauses = []
            else:
                rendered_clauses = render_clause_selector(selected_type, type_cfg, live_form)

        # ④ 당사자 인적사항 (AI로 당사자 정보 추출 시 자동 펼침)
        with st.expander(
            f"④ 당사자 인적사항 ({p1_role} / {p2_role})",
            expanded=bool(st.session_state.get("_expand_party_section", False)),
        ):
            st.markdown(f"**{p1_role}**")
            live_form["lessor_address"] = st.text_input(f"{p1_role} 주소", key=_wkey("lessor_address"))
            l1, l2 = st.columns(2)
            with l1:
                live_form["lessor_name"] = st.text_input(f"{p1_role} 성명", key=_wkey("lessor_name"))
            with l2:
                live_form["lessor_rrn"] = st.text_input(
                    f"{p1_role} 주민(사업자)번호", key=_wkey("lessor_rrn"), placeholder="000000-0000000"
                )
            l3, l4 = st.columns(2)
            with l3:
                live_form["lessor_phone1"] = st.text_input(
                    f"{p1_role} 전화1", key=_wkey("lessor_phone1"), placeholder="010-0000-0000"
                )
            with l4:
                live_form["lessor_phone2"] = st.text_input(f"{p1_role} 전화2", key=_wkey("lessor_phone2"))

            st.markdown("---")
            st.markdown(f"**{p2_role}**")
            live_form["lessee_address"] = st.text_input(f"{p2_role} 주소", key=_wkey("lessee_address"))
            t1, t2 = st.columns(2)
            with t1:
                live_form["lessee_name"] = st.text_input(f"{p2_role} 성명", key=_wkey("lessee_name"))
            with t2:
                live_form["lessee_rrn"] = st.text_input(
                    f"{p2_role} 주민(사업자)번호", key=_wkey("lessee_rrn"), placeholder="000000-0000000"
                )
            t3, t4 = st.columns(2)
            with t3:
                live_form["lessee_phone1"] = st.text_input(
                    f"{p2_role} 전화1", key=_wkey("lessee_phone1"), placeholder="010-0000-0000"
                )
            with t4:
                live_form["lessee_phone2"] = st.text_input(f"{p2_role} 전화2", key=_wkey("lessee_phone2"))

        # ⑤ 개업공인중개사 정보
        with st.expander("⑤ 개업공인중개사 정보 (공동중개 지원)", expanded=False):
            st.markdown("**주 개업공인중개사 (좌측란)**")
            b1, b2 = st.columns(2)
            with b1:
                live_form["broker1_office"] = st.text_input("사무소 명칭", key=_wkey("broker1_office"))
                live_form["broker1_rep"] = st.text_input("대표 성명", key=_wkey("broker1_rep"))
                live_form["broker1_reg_no"] = st.text_input("등록번호", key=_wkey("broker1_reg_no"))
            with b2:
                live_form["broker1_address"] = st.text_input("사무소 소재지", key=_wkey("broker1_address"))
                live_form["broker1_phone"] = st.text_input("전화번호", key=_wkey("broker1_phone"))
                live_form["broker1_associate"] = st.text_input(
                    "소속공인중개사 (선택)", key=_wkey("broker1_associate")
                )

            st.markdown("---")
            use_co = st.checkbox("공동중개사 입력 (우측란 활성화)", key=_wkey("use_co_broker"))
            live_form["use_co_broker"] = use_co
            if use_co:
                cb1, cb2 = st.columns(2)
                with cb1:
                    live_form["broker2_office"] = st.text_input("공동중개 사무소 명칭", key=_wkey("broker2_office"))
                    live_form["broker2_rep"] = st.text_input("공동중개 대표 성명", key=_wkey("broker2_rep"))
                    live_form["broker2_reg_no"] = st.text_input("공동중개 등록번호", key=_wkey("broker2_reg_no"))
                with cb2:
                    live_form["broker2_address"] = st.text_input("공동중개 사무소 소재지", key=_wkey("broker2_address"))
                    live_form["broker2_phone"] = st.text_input("공동중개 전화번호", key=_wkey("broker2_phone"))
                    live_form["broker2_associate"] = st.text_input(
                        "공동중개 소속공인중개사", key=_wkey("broker2_associate")
                    )

    # canonical 세션 동기화
    st.session_state["_form_data"] = dict(live_form)

    # 임시저장 / 최종 누적 저장 처리 (live_form 및 특약 상태 수집 직후 실행)
    if save_draft_clicked or save_final_clicked:
        target_status = "completed" if save_final_clicked else "draft"
        storage_payload = _build_storage_payload(live_form, selected_type, type_cfg)
        save_contract(
            contract_id=current_cid,
            contract_type=selected_type,
            contract_type_label=type_cfg["label"],
            status=target_status,
            payload=storage_payload,
        )
        st.session_state["current_contract_status"] = target_status
        msg = (
            f"✅ 최종 누적 저장 완료 ({current_cid})"
            if target_status == "completed"
            else f"💾 임시저장 완료 ({current_cid}) — 브라우저를 닫아도 이어서 작성할 수 있습니다."
        )
        st.toast(msg)
        st.rerun()

    # ==================== 공통 컨텍스트 빌드 및 교차검증 ====================
    form_payload = dict(live_form)
    form_payload["doc_id"] = current_cid
    form_payload["contract_type"] = selected_type
    form_payload["contract_type_label"] = type_cfg["label"]
    form_payload["form_title"] = type_cfg["form_title"]
    form_payload["color_mode"] = color_mode
    form_payload["draft_mode"] = draft_mode
    form_payload["blank_form_mode"] = blank_form_mode
    form_payload["interim_date"] = live_form["interim_date"] if live_form.get("use_interim_date") else None
    form_payload["rendered_clauses"] = [] if blank_form_mode else rendered_clauses

    context = build_contract_context(form_payload)
    is_valid = context["is_valid"]

    # ==================== 우측 실시간 A4 미리보기 및 출력 제어 ====================
    with col_preview, st.container(key="preview_pane"):
        with st.container(key="preview_actions_box"):
            is_blank_amounts = (
                context["deposit"] == 0
                and context["down_payment"] == 0
                and context["interim_payment"] == 0
                and context["balance"] == 0
            )
            if blank_form_mode:
                st.info(
                    "📄 **[완전 빈 양식 모드]** 날짜(`20  년  월  일`)·특약·금액이 모두 비워져 수기 작성용으로 바로 인쇄·다운로드할 수 있습니다."
                )
            elif is_blank_amounts:
                st.info("📝 금액 미입력 상태입니다. 이 상태에서도 바로 인쇄·다운로드할 수 있습니다.")
            elif is_valid:
                st.success(f"✅ {context['validation_msg']}")
            else:
                st.warning(f"⚠️ {context['validation_msg']} (참고용 안내 — 이 상태에서도 인쇄·다운로드 가능)")

            est_lines = context["estimated_clause_lines"]
            c_cnt = context["clause_count"]
            mode_badge = "🎨 컬러 인쇄용" if color_mode == "color" else "🖨️ 흑백 인쇄용"
            if context["is_annex"]:
                st.info(
                    f"{mode_badge} · 📄 특약 **총 {c_cnt}항 / {est_lines}줄** (한도 {MAX_INLINE_CLAUSE_COUNT}항·{MAX_INLINE_CLAUSE_LINES}줄 초과) → "
                    f"1페이지 **`※ 상세 특약은 첨부 [별지] 특약사항(총 {c_cnt}항) 참조`** 및 2페이지 [별지] 자동 생성됨"
                )
            else:
                st.caption(
                    f"{mode_badge} · 📏 특약: **{c_cnt}항 ({est_lines}/{MAX_INLINE_CLAUSE_LINES}줄)** (A4 1장 내 출력 · 5항 초과 시 별지 자동 분리)"
                )

            date_tag = str(live_form.get("contract_date") or date.today()).replace("-", "")
            color_tag = "_컬러" if color_mode == "color" else "_흑백"
            draft_tag = "_초안" if draft_mode else ""
            clean_type_name = meta.get("short_label", type_cfg["label"]).replace(" ", "_")
            base_filename = f"{clean_type_name}_계약서_{date_tag}{color_tag}{draft_tag}"

            printable_html = render_html(context, include_print_toolbar=True)
            has_user_hwpx = has_custom_hwpx_template(selected_type)

            # 모든 인쇄·다운로드 버튼 상시 활성화 (빈 양식·미저장 상태에서도 즉시 출력 가능)
            dl_cols = st.columns(3 if has_user_hwpx else 2)

            with dl_cols[0]:
                pdf_bytes, pdf_msg = generate_pdf(context)
                if pdf_bytes is not None:
                    st.download_button(
                        "📄 PDF 다운로드 (A4)",
                        data=pdf_bytes,
                        file_name=f"{base_filename}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        type="primary",
                    )
                else:
                    st.download_button(
                        "🖨️ PDF 인쇄용 HTML (A4 100%)",
                        data=printable_html.encode("utf-8"),
                        file_name=f"{base_filename}_인쇄용.html",
                        mime="text/html",
                        use_container_width=True,
                        type="primary",
                        help=pdf_msg,
                    )

            with dl_cols[1]:
                docx_bytes = generate_docx(context)
                st.download_button(
                    "📝 Word (.docx · 한글에서도 열림)",
                    data=docx_bytes,
                    file_name=f"{base_filename}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

            if has_user_hwpx:
                with dl_cols[2]:
                    hwpx_bytes = generate_hwpx(context)
                    st.download_button(
                        "📘 사용자 한글 (.hwpx)",
                        data=hwpx_bytes,
                        file_name=f"{base_filename}.hwpx",
                        mime="application/hwp+zip",
                        use_container_width=True,
                    )

            hwpx_panel_title = (
                f"📘 사용자 한글(.hwpx) 템플릿 등록됨 — {type_cfg['label']} 전용 .hwpx 활성화"
                if has_user_hwpx
                else "📘 사용자 한글(.hwpx) 템플릿 업로드 (선택 · 미등록 시 Word .docx 사용 권장)"
            )
            with st.expander(hwpx_panel_title, expanded=False):
                st.caption(
                    "💡 **실무 팁**: 위 **Word(.docx)** 파일은 MS Word뿐 아니라 **아래아한글에서도 바로 열리고 편집**되며, "
                    "선택한 컬러/흑백 모드가 그대로 적용됩니다. 사무소 전용 `.hwpx` 서식이 있을 경우 여기에 업로드해 사용하세요."
                )
                uploaded_hwpx = st.file_uploader(
                    f"{type_cfg['label']} 전용 .hwpx 파일 선택",
                    type=["hwpx"],
                    key=f"hwpx_uploader_{selected_type}",
                )
                u_col1, u_col2 = st.columns(2)
                with u_col1:
                    if st.button(
                        "📤 사용자 .hwpx 등록",
                        disabled=uploaded_hwpx is None,
                        use_container_width=True,
                        key=f"hwpx_save_btn_{selected_type}",
                    ):
                        if uploaded_hwpx is not None:
                            ok_up, up_msg, found_tags = save_custom_hwpx_template(
                                selected_type, uploaded_hwpx.getvalue()
                            )
                            if ok_up:
                                st.success(up_msg)
                                if found_tags:
                                    st.info("인식된 치환 태그: " + ", ".join(f"`{{{{{t}}}}}`" for t in found_tags[:20]))
                                st.rerun()
                            else:
                                st.error(up_msg)
                with u_col2:
                    if st.button(
                        "🗑️ 등록된 .hwpx 해제",
                        disabled=not has_user_hwpx,
                        use_container_width=True,
                        key=f"hwpx_reset_btn_{selected_type}",
                    ):
                        reset_custom_hwpx_template(selected_type)
                        st.success("사용자 .hwpx 템플릿이 해제되었습니다.")
                        st.rerun()

                if st.checkbox("지원 치환 태그 목록 보기", key=f"show_tags_{selected_type}"):
                    tag_md = "\n".join(f"- `{tag}` : {desc}" for tag, desc in HWPX_PLACEHOLDER_GUIDE)
                    st.markdown(tag_md)

        # 실시간 A4 미리보기 iframe (확대/축소 [축소 -][화면맞춤][크게 100%][더 크게 130%] + 4방향 터치 스크롤 지원)
        with st.container(key="preview_iframe_box"):
            live_preview_html = _enhance_live_preview_html(printable_html)
            preview_height = 2450 if context["is_annex"] else 1380
            if hasattr(st, "iframe"):
                st.iframe(live_preview_html, height=preview_height)
            else:
                components.html(live_preview_html, height=preview_height, scrolling=True)


if __name__ == "__main__":
    main()
