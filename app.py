"""검수시트 자동 생성기 - Streamlit UI."""
from __future__ import annotations

import html as _html
import io
import json
import re
from datetime import date

import streamlit as st
import streamlit.components.v1 as components


def _sanitize_filename(name: str) -> str:
    """파일명 유효성 보정: 공백 트리밍 + 금지문자 제거 + 100자 제한."""
    name = (name or "").strip()
    if not name:
        return ""
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    if len(name) > 100:
        name = name[:100]
    return name

from converter import (
    CHECKSHEET_HEADERS,
    DAYS,
    WEEKDAYS,
    COL_A,
    COL_TIER,
    COL_MAPPING_TYPE,
    COL_EPISODE_NUMBER,
    COL_FORMAL_NUMBER,
    COL_NEW_SUMMARY,
    PREVIEW_COL_INDICES,
    PREVIEW_COL_WIDTHS,
    parse_csv,
    build_for_day,
    build_xlsx,
    build_ordered_entries,
    select_day_range,
    check_jongyeong_alerts,
)

st.set_page_config(page_title="검수시트 자동 생성기", page_icon="🎬", layout="wide")


def _password_gate() -> None:
    """팀 공용 비밀번호 게이트. Streamlit Secrets의 'app_password' 값과 비교."""
    if st.session_state.get("authed"):
        return

    try:
        correct = st.secrets.get("app_password", "")
    except Exception:
        correct = ""
    if not correct:
        # secrets에 비밀번호가 설정되지 않은 경우 — 코드에 fallback 두지 않음(보안)
        st.error(
            "비밀번호가 설정되어 있지 않습니다. 관리자가 Streamlit Cloud Settings → Secrets에서 "
            "`app_password` 값을 입력하거나, 로컬 실행 시 `.streamlit/secrets.toml`에 추가해주세요."
        )
        st.stop()

    st.markdown(
        """
        <style>
        .gate-wrap { max-width: 420px; margin: 60px auto 16px auto; padding: 32px 28px;
                     background: #ffffff; border: 1px solid #e2e8f0; border-radius: 14px;
                     box-shadow: 0 6px 18px rgba(15,23,42,0.05); text-align: center; }
        .gate-title { font-size: 18px; font-weight: 700; color: #0f172a; margin-bottom: 6px; }
        .gate-sub { font-size: 13px; color: #64748b; margin-bottom: 18px; }
        </style>
        <div class="gate-wrap">
            <div class="gate-title">🔒 검수시트 자동 생성기</div>
            <div class="gate-sub">팀 비밀번호를 입력하세요</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    cc = st.columns([1, 2, 1])
    with cc[1]:
        with st.form("auth_form", clear_on_submit=False):
            pwd = st.text_input("비밀번호", type="password", label_visibility="collapsed", placeholder="비밀번호")
            submitted = st.form_submit_button("입장", type="primary", use_container_width=True)
            if submitted:
                # secrets가 숫자로 파싱돼도 (따옴표 없이) 비교가 되도록 양쪽을 문자열로 정규화
                if pwd and pwd.strip() == str(correct).strip():
                    st.session_state["authed"] = True
                    st.rerun()
                else:
                    st.error("비밀번호가 일치하지 않습니다.")
    st.stop()


_password_gate()

st.markdown(
    """
    <style>
    /* 다운로드 카드 영역(SaaS 톤) — Streamlit container(border=True)로 감싼 후 내부 텍스트만 스타일 */
    .dl-card-title {
        font-size: 15.5px;
        font-weight: 600;
        color: #0f172a;
        margin: 0 0 6px 0;
        font-family: inherit;
    }
    .dl-card-sub {
        font-size: 13.5px;
        color: #475569;
        line-height: 1.55;
        margin: 0 0 10px 0;
        font-family: inherit;
    }
    .dl-card-filename {
        font-size: 13px;
        color: #94a3b8;
        margin: 0;
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
    }
    /* 다운로드 버튼 — '변환 미리보기'(Streamlit primary) 톤을 색만 초록으로. 솔리드 + 그림자 거의 없음. */
    div[data-testid="stDownloadButton"] > button {
        background: #00b46e !important;
        color: #ffffff !important;
        border: 1px solid #00b46e !important;
        padding: 6px 18px !important;
        border-radius: 8px !important;
        font-size: 13px !important;
        font-weight: 600 !important;
        font-family: inherit !important;
        letter-spacing: 0 !important;
        box-shadow: none !important;
        transition: background 0.15s ease, border-color 0.15s ease;
        width: auto !important;
        min-width: 0 !important;
        height: 38px !important;
        line-height: 1.2 !important;
        white-space: nowrap !important;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background: #009d61 !important;
        border-color: #009d61 !important;
        box-shadow: none !important;
    }
    div[data-testid="stDownloadButton"] > button:active {
        background: #008555 !important;
        border-color: #008555 !important;
    }
    div[data-testid="stDownloadButton"] > button p {
        margin: 0 !important;
        font-size: 13px !important;
        font-weight: 600 !important;
    }
    /* 시작 요일 셀렉트박스와 변환 미리보기 버튼 정렬용 */
    .label-spacer {
        height: 28px;
    }
    /* 셀렉트박스 줄 아래 hint 리스트 */
    .hint-list {
        margin-top: 6px;
        font-size: 12.5px;
        color: #64748b;
        line-height: 1.85;
    }
    .hint-list > div {
        padding-left: 2px;
    }
    /* 카드 컨테이너 — Streamlit 1.50의 st.container(border=True) 기본 보더를 그대로 사용해
       위쪽 'CSV 업로드' 카드와 시각적 일관성 유지. 커스텀 shadow/border 덮어씌우지 않음 */
    /* 다운로드 카드 강조용 그린 좌측 보더 (다운로드 액션 카드 전용 markdown 클래스로 적용) */
    .dl-accent {
        border-left: 3px solid #059669;
        padding-left: 12px;
        margin-left: -2px;
        margin-bottom: 8px;
    }
    /* 파일명 입력 — 항상 보이는 부드러운 박스, .xlsx 접미를 박스 안에 통합(::after).
       다운로드 카드(.dl-accent 마커) 내부 stTextInput만 정확히 타게팅.
       wrapper 자체에도 max-width 360px → ::after(right:14px)가 input 박스 안 우측에 정확히 위치 */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stTextInput"] {
        position: relative;
        margin: 0 !important;
        max-width: 360px !important;
        width: 100% !important;
    }
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stTextInput"]::after {
        content: '.xlsx';
        position: absolute;
        right: 14px;
        top: 50%;
        transform: translateY(-50%);
        color: #94a3b8;
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
        font-size: 13px;
        line-height: 1;
        pointer-events: none;
        z-index: 2;
    }
    /* BaseWeb input 래퍼들의 기본 배경/그림자 전부 transparent — input 자체만 보이게 */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stTextInput"] > div,
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stTextInput"] [data-baseweb="input"],
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stTextInput"] [data-baseweb="base-input"] {
        background: transparent !important;
        background-color: transparent !important;
        box-shadow: none !important;
        border: none !important;
    }
    /* 파일명 input — 거의 흰색(#f9f9f9) + 얇은 회색 보더만. 그림자/음영 없음 */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stTextInput"] input {
        background: #f9f9f9 !important;
        background-color: #f9f9f9 !important;
        border: 1px solid #e5e7eb !important;
        border-radius: 8px !important;
        box-shadow: none !important;
        color: #0f172a !important;
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace !important;
        font-size: 13px !important;
        padding: 4px 56px 4px 14px !important;
        height: 38px !important;
        width: 100% !important;
        max-width: 360px !important;
        transition: border-color 0.15s ease;
        cursor: text;
    }
    /* 카드 안의 가로 행 컬럼 사이 간격 좁히기 (Streamlit 기본 1rem → 0.5rem) */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stHorizontalBlock"] {
        gap: 0.5rem !important;
    }
    /* 호버 — 배경은 그대로, 보더만 살짝 어두워짐 */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stTextInput"] input:hover {
        border-color: #d1d5db !important;
        box-shadow: none !important;
    }
    /* 포커스 — 흰 배경 + 인디고 보더만 (글로우 없이) */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stTextInput"] input:focus {
        background: #ffffff !important;
        background-color: #ffffff !important;
        border-color: #6366f1 !important;
        box-shadow: none !important;
        outline: none !important;
    }

    /* 카드 행 미니 라벨 — '파일명' / '복사'. 13px 통일, input/버튼(38px)과 수직 정렬 */
    .fn-static-label {
        font-size: 13px;
        color: #64748b;
        font-weight: 500;
        letter-spacing: 0;
        line-height: 38px;
        white-space: nowrap;
        font-family: inherit;
    }
    /* 체크박스 라벨도 13px 통일 (다운로드 카드 내) */
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stCheckbox"] label,
    [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .dl-accent) [data-testid="stCheckbox"] label p {
        font-size: 13px !important;
    }
    /* 카드 내 행간 여백 */
    .card-row-gap {
        height: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🎬 검수시트 자동 생성기")
st.caption(
    "주간편성표 CSV를 업로드하고 요일 범위를 선택하면, 검수시트 엑셀과 동일한 컬럼 구조로 변환된 미리보기·xlsx 다운로드가 제공됩니다."
)

# 편성표 입력에서 기대하는 컬럼 헤더 (20컬럼 풀셋). 19컬럼(끝의 "창고" 누락)도 허용.
EXPECTED_HEADERS = [
    "비고", "예약작", "오픈일", "요일", "시간", "Tier", "id", "code", "Title",
    "지난주 e/n", "지난주 f/n", "이번주 e/n", "이번주 f/n", "오픈 에피 수",
    "편성비고", "WX/동시방영", "최종회차", "CP bill", "원본경로", "창고",
]
REQUIRED_HEADERS_MIN = set(EXPECTED_HEADERS) - {"창고"}  # 창고는 옵셔널


def _validate_headers(df) -> tuple[bool, str]:
    """편성표 입력 헤더 검증. (ok, message) 반환."""
    cols = list(df.columns)
    n = len(cols)
    if n not in (19, 20):
        return False, (
            f"컬럼 개수가 {n}개입니다. 편성표는 19개(창고 제외) 또는 20개 컬럼이어야 해요.\n"
            f"기대 헤더: {', '.join(EXPECTED_HEADERS)}"
        )
    missing = REQUIRED_HEADERS_MIN - set(cols)
    if missing:
        return False, f"필수 컬럼 누락: {', '.join(sorted(missing))}"
    return True, ""


with st.container(border=True):
    input_tabs = st.tabs(["📂 CSV 업로드", "📋 텍스트 붙여넣기"])
    with input_tabs[0]:
        uploaded = st.file_uploader("편성표 CSV 업로드", type=["csv"], label_visibility="collapsed")
    with input_tabs[1]:
        pasted_text = st.text_area(
            "편성표 텍스트 붙여넣기",
            height=180,
            key="pasted_text",
            placeholder=(
                "구글시트/엑셀에서 헤더 행 포함하여 셀들을 복사한 후 여기 붙여넣으세요.\n"
                "탭 구분(TSV) 자동 감지. 쉼표 구분(CSV)도 OK."
            ),
            label_visibility="collapsed",
        )

    cols = st.columns([2, 2, 2])
    with cols[0]:
        start_day = st.selectbox("시작 요일", WEEKDAYS, index=0)
    with cols[1]:
        end_day = st.selectbox("끝 요일", WEEKDAYS, index=WEEKDAYS.index(start_day))
    with cols[2]:
        # 셀렉트박스 라벨 자리만큼 여백을 줘서 버튼이 셀렉트박스 입력 영역과 같은 줄에 정렬되게
        st.markdown('<div class="label-spacer"></div>', unsafe_allow_html=True)
        run = st.button("변환 미리보기", type="primary", use_container_width=True)
    # 셀렉트박스 아래에 안내문 — 줄별로
    st.markdown(
        """
        <div class="hint-list">
            <div>· 월~금만 검수시트로 만듭니다. 토/일 편성은 자동으로 금요일에 합쳐져요.</div>
            <div>· 한 요일만 보고 싶으면 시작·끝을 같게 두세요.</div>
            <div>· EBS 섹션은 금요일에 자동 포함됩니다.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _esc(v) -> str:
    return _html.escape(str(v) if v is not None else "")


def _column_letter(idx: int) -> str:
    s = ""
    n = idx
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def _row_bg(category: str, no_fill: bool = False, over_final: bool = False) -> str:
    """검수시트 템플릿(Arial 10pt) 색 톤에 맞춘 미리보기 행 배경."""
    if over_final:
        return "#cfe2f3"  # 종영 의심 — 옅은 파랑 (템플릿 동일)
    if no_fill:
        return "#ffffff"
    if category == "header":
        return "#d9d9d9"  # 요일 헤더 회색
    if category == "header_holiday":
        return "#00ff00"  # 휴일 형광 초록 (템플릿)
    if category == "reserved":
        return "#ead1dc"  # 예약작 — 연한 자홍색 3 (xlsx와 동기화)
    if category == "cancelled":
        return "#e6b8af"  # 결방/홀드백 — 연한 붉은 딸기색 3 (xlsx와 동기화)
    if category == "ebs":
        return "#ffffff"  # EBS 콘텐츠 행은 음영 없음
    if category == "holiday_header":
        return "#d9d9d9"  # 연휴지연 그룹 헤더 (요일과 같은 회색)
    if category == "holiday_body":
        return "#cccccc"  # 연휴지연 콘텐츠 (더 진한 회색)
    if category == "prev_day":
        return "#ead1dc"  # 전날 셋팅 = 예약작과 동일 (xlsx와 동기화)
    return "#ffffff"


def _cell_html(val: str, col_idx: int, category: str) -> str:
    if col_idx == COL_A:
        if val == "종영":
            return f"<span style='color:#222;font-weight:700;background:#fff292;padding:1px 4px;border-radius:3px'>{_esc(val)}</span>"
        if category == "holiday_header":
            # 연휴지연 그룹 헤더 — 빨간 굵게로 강조 (유지)
            return f"<span style='font-weight:700;color:#ff0000'>{_esc(val)}</span>"
        # 요일 헤더(header, header_holiday)는 일반 굵기 — 음영(녹색/회색)으로 충분히 구분됨
        return _esc(val)
    if col_idx == COL_MAPPING_TYPE and val:
        badge_color = "#1976d2" if val == "basic" else "#9c27b0"
        return (
            f"<span style='background:{badge_color};color:white;padding:1px 6px;"
            f"border-radius:8px;font-size:10px'>{_esc(val)}</span>"
        )
    return _esc(val)


def _render_preview(headers, results_by_day: list[tuple[str, dict]]) -> str:
    """여러 요일 결과를 한 표에 이어붙여 미리보기 (매핑된 컬럼만 표시).
    table-layout: fixed + <colgroup>으로 컬럼 폭을 고정하고, 신규 요약처럼 긴 헤더 셀은
    colspan으로 옆 셀들 위에 시각적으로 흘러나오게 처리(구글 스프레드시트 셀 오버플로우와 동일).
    """
    indices = PREVIEW_COL_INDICES
    new_summary_pos = indices.index(COL_NEW_SUMMARY)  # colspan 시작 위치
    table_width = sum(PREVIEW_COL_WIDTHS.get(i, 100) for i in indices)
    cell_base = "border:1px solid #ddd;padding:4px 8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"

    parts = ["<div style='overflow-x:auto;border:1px solid #ccc'>"]
    # 검수시트 템플릿 통일: Arial 10pt
    parts.append(
        f"<table style='border-collapse:collapse;font-size:13px;"
        f"font-family:Arial,Helvetica,sans-serif;"
        f"table-layout:fixed;width:{table_width}px'>"
    )
    # 컬럼 폭 고정
    parts.append("<colgroup>")
    for i in indices:
        w = PREVIEW_COL_WIDTHS.get(i, 100)
        parts.append(f"<col style='width:{w}px'>")
    parts.append("</colgroup>")
    # 헤더1: 알파벳
    parts.append("<thead><tr style='background:#bdbdbd'>")
    for i in indices:
        parts.append(
            "<th style='border:1px solid #777;padding:2px 6px;font-size:10px;color:#222;"
            "overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>"
            + _esc(_column_letter(i))
            + "</th>"
        )
    parts.append("</tr><tr style='background:#e8e8e8'>")
    for i in indices:
        h = headers[i] if headers[i] else "(구분)"
        parts.append(
            "<th style='border:1px solid #999;padding:4px 8px;font-size:11px;"
            "font-weight:600;text-align:left;overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>"
            + _esc(h)
            + "</th>"
        )
    parts.append("</tr></thead><tbody>")

    def emit_row(entry):
        cat = entry["category"]
        no_fill = bool(entry.get("no_fill", False))
        is_over = bool(entry.get("over_final"))
        bg = _row_bg(cat, no_fill, is_over)
        parts.append(f"<tr style='background:{bg}'>")

        # 요일 헤더 행에서 C열(원본검수 = COL_NEW_SUMMARY)의 신규 요약은 길어서
        # 옆 빈 셀들 위로 흘러나오게 colspan 처리 (구글 시트 셀 오버플로우 흉내)
        is_overflow_header = (
            cat in ("header", "header_holiday")
            and bool(entry["values"][COL_NEW_SUMMARY])
        )

        # A열은 ellipsis 없이 전체 텍스트 표시 (휴일 라벨, 연휴지연에피 비고 텍스트가 길어도 잘리지 않음)
        a_cell_style = "border:1px solid #ddd;padding:4px 8px;white-space:nowrap;overflow:visible;"

        if is_overflow_header:
            for n, i in enumerate(indices):
                v = entry["values"][i]
                cell = _cell_html(v, i, cat)
                if n < new_summary_pos:
                    style = a_cell_style if i == COL_A else cell_base
                    parts.append(f"<td style='{style}'>{cell}</td>")
                elif n == new_summary_pos:
                    remaining = len(indices) - n
                    parts.append(
                        f"<td colspan='{remaining}' "
                        f"style='border:1px solid #ddd;padding:4px 8px;"
                        f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>"
                        f"{cell}</td>"
                    )
                    break
        else:
            for i in indices:
                v = entry["values"][i]
                cell = _cell_html(v, i, cat)
                style = a_cell_style if i == COL_A else cell_base
                parts.append(f"<td style='{style}'>{cell}</td>")
        parts.append("</tr>")

    for d, result in results_by_day:
        emit_row(result["header_row"])
        for r in result["body"]:
            emit_row(r)
        for r in result["cancelled"]:
            emit_row(r)
        for r in result.get("holiday_block", []):
            emit_row(r)
        if d == "금" and result["ebs_section"]:
            parts.append("<tr style='background:#d9d9d9'>")
            for k, _i in enumerate(indices):
                label = "EBS" if k == 0 else ""
                parts.append(
                    f"<td style='{cell_base}font-weight:700'>{_esc(label)}</td>"
                )
            parts.append("</tr>")
            for r in result["ebs_section"]:
                emit_row(r)

    parts.append("</tbody></table></div>")
    return "".join(parts)


def _render_clipboard_html(headers, results_by_day, include_header: bool = False) -> str:
    """클립보드 복사용 HTML 생성 — 검수시트 70컬럼 전체.
    구글시트에 붙여넣었을 때 서식(음영/굵기/정렬)이 유지되도록 인라인 스타일 사용.
    원칙: 매핑 컬러 제거 · 테두리 제거 · A열 종영 스타일 제거 · 연휴지연 헤더만 빨간 굵게 유지 ·
          EBS 행 일반 굵기 · Tier 좌측 정렬 · '결방' 셀 우측 정렬.
    """
    indices = list(range(len(headers)))  # 70컬럼 전체

    parts = [
        "<table style='border-collapse:collapse;"
        "font-family:Arial,Helvetica,sans-serif;font-size:10pt'>"
    ]

    if include_header:
        parts.append("<tr style='background:#d9d9d9'>")
        for i in indices:
            h = headers[i] if headers[i] else ""
            parts.append(
                f"<td style='padding:2px 6px;font-weight:600'>{_esc(h)}</td>"
            )
        parts.append("</tr>")

    def emit_clip_row(entry, *, override_bg: str | None = None, force_bold_a: bool = False) -> None:
        cat = entry["category"]
        no_fill = bool(entry.get("no_fill", False))
        is_over = bool(entry.get("over_final"))
        bg = override_bg if override_bg is not None else _row_bg(cat, no_fill, is_over)
        bg_style = f"background:{bg};" if bg and bg != "#ffffff" else ""
        parts.append(f"<tr style='{bg_style}'>")

        for i in indices:
            v = entry["values"][i] if i < len(entry["values"]) else ""
            v_str = str(v) if v is not None else ""

            # 셀 컨텐츠 — 매핑 컬러 배지 제거. A열 종영 셀도 일반 텍스트로 (볼드·노란음영 제거).
            # 단, 연휴지연 헤더의 A열만 빨간 굵게 유지.
            if i == COL_A and cat == "holiday_header":
                content = f"<span style='font-weight:700;color:#ff0000'>{_esc(v_str)}</span>"
            elif i == COL_A and force_bold_a:
                # EBS 배너 행 A열 — 클립보드에서는 일반 굵기 (요청: EBS 볼드 → 일반)
                content = _esc(v_str)
            else:
                content = _esc(v_str)

            # 정렬: Tier 좌측 / "결방"만 정확히 일치할 때 우측 / 그 외 기본
            align_style = ""
            if i == COL_TIER:
                align_style = "text-align:left;"
            if v_str == "결방":
                align_style = "text-align:right;"

            parts.append(
                f"<td style='padding:2px 6px;{align_style}'>{content}</td>"
            )
        parts.append("</tr>")

    for d, result in results_by_day:
        emit_clip_row(result["header_row"])
        for r in result["body"]:
            emit_clip_row(r)
        for r in result["cancelled"]:
            emit_clip_row(r)
        for r in result.get("holiday_block", []):
            emit_clip_row(r)
        if d == "금" and result["ebs_section"]:
            # EBS 배너 — 회색 배경, 일반 굵기로 "EBS"만 표시
            parts.append("<tr style='background:#d9d9d9'>")
            for i in indices:
                label = "EBS" if i == COL_A else ""
                parts.append(f"<td style='padding:2px 6px'>{_esc(label)}</td>")
            parts.append("</tr>")
            for r in result["ebs_section"]:
                emit_clip_row(r)

    parts.append("</table>")
    return "".join(parts)


def _clipboard_button(
    html_content: str,
    *,
    button_label: str = "📋 클립보드에 복사",
    key: str = "clip",
    height: int = 52,
) -> None:
    """클립보드 복사 — 텍스트 링크 스타일 + 클릭 시 아래에 '✅ 복사됨!' 페이드 피드백.

    구현 메모: navigator.clipboard.write()는 iframe sandbox에서 권한 문제로
    조용히 실패하는 경우가 있어, 호환성이 가장 좋은 execCommand('copy') +
    contentEditable div + copy 이벤트(text/html + text/plain) 방식을 사용한다.
    """
    html_js = json.dumps(html_content)
    safe_key = re.sub(r"[^A-Za-z0-9_]", "_", key)
    component_html = f"""
    <style>
      html, body {{
        margin: 0;
        padding: 0;
        background: transparent;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      }}
      .clip-link-{safe_key} {{
        display: inline-block;
        color: #6366f1;
        font-size: 13px;
        font-weight: 500;
        cursor: pointer;
        padding: 6px 0;
        user-select: none;
        white-space: nowrap;
        transition: color 0.15s ease;
      }}
      .clip-link-{safe_key}:hover {{
        color: #4f46e5;
        text-decoration: underline;
      }}
      .clip-status-{safe_key} {{
        display: block;
        font-size: 12.5px;
        font-weight: 600;
        margin-top: 2px;
        opacity: 0;
        transition: opacity 0.3s ease;
        pointer-events: none;
      }}
      .clip-status-{safe_key}.show {{ opacity: 1; }}
    </style>
    <span class="clip-link-{safe_key}" id="clip-link-{safe_key}"
          onclick="copyHtml_{safe_key}()">{_html.escape(button_label)}</span>
    <div class="clip-status-{safe_key}" id="clip-status-{safe_key}"></div>
    <script>
    (function() {{
      const html_payload = {html_js};
      window.copyHtml_{safe_key} = function() {{
        const div = document.createElement('div');
        div.contentEditable = 'true';
        div.innerHTML = html_payload;
        div.style.position = 'fixed';
        div.style.left = '-99999px';
        div.style.top = '0';
        div.style.opacity = '0';
        document.body.appendChild(div);

        const onCopy = function(e) {{
          e.clipboardData.setData('text/html', html_payload);
          e.clipboardData.setData('text/plain', div.innerText);
          e.preventDefault();
        }};
        document.addEventListener('copy', onCopy);

        const range = document.createRange();
        range.selectNodeContents(div);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);

        let ok = false;
        try {{
          ok = document.execCommand('copy');
        }} catch (err) {{
          ok = false;
        }}

        document.removeEventListener('copy', onCopy);
        sel.removeAllRanges();
        document.body.removeChild(div);

        const status = document.getElementById('clip-status-{safe_key}');
        if (ok) {{
          status.innerText = '✅ 복사됨!';
          status.style.color = '#059669';
        }} else {{
          status.innerText = '⚠️ 복사 실패 — 다시 시도해주세요';
          status.style.color = '#dc2626';
        }}
        status.classList.add('show');
        clearTimeout(window._statusTimer_{safe_key});
        window._statusTimer_{safe_key} = setTimeout(function() {{
          status.classList.remove('show');
        }}, 2500);
      }};
    }})();
    </script>
    """
    components.html(component_html, height=height)


# 변환 미리보기 버튼이 눌리면 결과를 session_state에 캐시 — 다운로드 버튼 클릭으로
# Streamlit이 rerun을 트리거해도 미리보기가 그대로 유지되도록.
if run:
    text_input = (pasted_text or "").strip()
    if uploaded is None and not text_input:
        st.error("CSV 파일을 업로드하거나 텍스트를 붙여넣어주세요.")
        st.session_state.pop("preview_active", None)
    else:
        df_now = None
        try:
            if text_input:
                # 첫 줄에 탭이 있으면 TSV(구글시트 복사), 아니면 CSV
                first_line = text_input.split("\n", 1)[0]
                sep = "\t" if "\t" in first_line else ","
                df_now = parse_csv(io.StringIO(text_input), sep=sep)
            else:
                df_now = parse_csv(uploaded)
        except Exception as e:
            st.error(f"편성표 파싱 실패: {e}")
            st.stop()

        ok, msg = _validate_headers(df_now)
        if not ok:
            st.error(f"편성표 헤더 검증 실패\n\n{msg}")
            st.stop()

        days_now = select_day_range(start_day, end_day)
        if not days_now:
            st.error("요일 선택이 잘못되었습니다.")
            st.stop()
        st.session_state["cached_df"] = df_now
        st.session_state["cached_days"] = days_now
        st.session_state["preview_active"] = True

# preview_active가 True이면 캐시된 df/days로 미리보기·다운로드 렌더링
if st.session_state.get("preview_active") and st.session_state.get("cached_df") is not None:
    df = st.session_state["cached_df"]
    days = st.session_state["cached_days"]
    if True:  # 들여쓰기 유지를 위한 래퍼

        with st.expander(f"📥 업로드된 편성표 (전체 {len(df)}행) — 클릭해서 펼치기"):
            st.dataframe(df, use_container_width=True)

        # 종영 의심 알림 (편성표 전체 기준) — 콜아웃 박스
        alerts = check_jongyeong_alerts(df)
        if alerts:
            items_html = "".join(
                f"<li style='margin:6px 0'>"
                f"<span style='color:#d32f2f;font-weight:700'>확인 필요</span> "
                f"<strong>{_esc(a['title'])}</strong> "
                f"<strong>{a['final']}회차</strong>를 끝으로 종영인 것으로 확인됨"
                f"</li>"
                for a in alerts
            )
            st.markdown(
                f"""
                <div style='
                    background:#fff5f5;
                    border:1px solid #ffcdd2;
                    border-left:6px solid #d32f2f;
                    border-radius:12px;
                    padding:18px 22px;
                    margin:14px 0;
                    box-shadow:0 4px 14px rgba(211,47,47,0.10);
                '>
                    <div style='font-size:16.5px;font-weight:800;color:#b71c1c;margin-bottom:8px;'>
                        종영 의심 항목
                    </div>
                    <div style='font-size:13.5px;color:#555;margin-bottom:12px;line-height:1.6;'>
                        편성표의 <em>최종회차</em>보다 이번주 회차가 더 큽니다. 미리보기와 엑셀 시트 맨 위에 동일하게 표시되며,
                        해당 회차 셀은 <span style='color:#1565c0;font-weight:700'>파란색</span>으로 강조됩니다.
                    </div>
                    <ul style='margin:0;padding-left:22px;line-height:1.85;font-size:14.5px;color:#222;'>
                        {items_html}
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )

        # 요일별 결과 모으기
        results_by_day: list[tuple[str, dict]] = [(d, build_for_day(df, d)) for d in days]

        # 통계
        total_new = sum(r["new_count"] for _, r in results_by_day)
        total_body = sum(len(r["body"]) for _, r in results_by_day)
        total_cancel = sum(len(r["cancelled"]) for _, r in results_by_day)
        total_ebs = sum(len(r["ebs_section"]) for d, r in results_by_day if d == "금")
        stats = st.columns(5)
        stats[0].metric("선택 요일 수", len(days))
        stats[1].metric("신규 합계", total_new, help="요일별 헤더 행 C열에 합쳐서 표시됩니다.")
        stats[2].metric("본문 행 합계", total_body)
        stats[3].metric("결방/홀드백 합계", total_cancel)
        stats[4].metric("EBS 섹션", total_ebs)

        range_label = f"{days[0]}-{days[-1]}요일" if len(days) > 1 else f"{days[0]}요일"
        st.markdown(f"### 📋 {range_label} 검수시트 미리보기")
        st.caption(
            "데이터가 채워지는 컬럼만 표시했어요(공란 컬럼은 숨김). 다운로드 파일은 검수시트 70개 컬럼 그대로 나옵니다. "
            "🟪 핑크 = 예약작 · 🟥 빨강 = 결방/홀드백(요일 하단) · 노란 셀 = 종영 (A열) · "
            "회색 헤더 = 요일 시작 행. 비고 'A열'에 '전날 셋팅' 적혀 있으면 자동으로 전날에 포함됩니다."
        )

        st.markdown(
            _render_preview(CHECKSHEET_HEADERS, results_by_day),
            unsafe_allow_html=True,
        )

        try:
            xlsx_bytes = build_xlsx(df, days)
            today = date.today().isoformat()
            filename = f"검수시트_{today}_{range_label}.xlsx"
            default_base = f"검수시트_{today}_{range_label}"
            # 요일 범위가 바뀌면 파일명 base를 새 default로 리셋
            if st.session_state.get("filename_range_label") != range_label:
                st.session_state["filename_base"] = default_base
                st.session_state["filename_range_label"] = range_label

            # 미리보기 표와 다운로드 카드 사이 여백
            st.markdown('<div style="height:32px"></div>', unsafe_allow_html=True)

            with st.container(border=True):
                st.markdown(
                    """
                    <div class="dl-accent">
                        <div class="dl-card-title">검수시트 준비 완료</div>
                        <div class="dl-card-sub">xlsx로 다운로드하거나 구글시트로 바로 복사할 수 있어요. 클립보드 복사는 셀 색·정렬·굵기까지 유지됩니다.</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # 파일명 행 — 라벨 / [input + .xlsx(::after로 내부 통합)] / 콤팩트 다운로드 버튼 / 버퍼
                fn_cols = st.columns([0.5, 3.2, 1.3, 2.2])
                with fn_cols[0]:
                    st.markdown('<div class="fn-static-label">파일명</div>', unsafe_allow_html=True)
                with fn_cols[1]:
                    st.text_input(
                        "파일명",
                        key="filename_base",
                        label_visibility="collapsed",
                        help="파일명을 수정할 수 있어요. Enter 또는 바깥 클릭으로 저장됩니다.",
                    )

                edited_base = _sanitize_filename(st.session_state.get("filename_base", "")) or default_base
                filename = f"{edited_base}.xlsx"

                with fn_cols[2]:
                    # use_container_width 제거 — CSS의 width:auto와 합쳐져 콘텐츠 너비로 콤팩트
                    st.download_button(
                        label="엑셀 다운로드",
                        data=xlsx_bytes,
                        file_name=filename,
                        mime="application/vnd.openxlsxformats-officedocument.spreadsheetml.sheet",
                    )

                # 행간 여백
                st.markdown('<div class="card-row-gap"></div>', unsafe_allow_html=True)

                # 클립보드 복사 행 — 라벨 / 텍스트 링크 / 헤더 포함 체크박스 (정말 바로 옆에 붙도록)
                # include_header는 체크박스가 뒤에 렌더되므로 session_state에서 먼저 읽어 링크 HTML에 반영
                include_header = bool(st.session_state.get("clip_include_header", False))
                clip_html = _render_clipboard_html(
                    CHECKSHEET_HEADERS, results_by_day, include_header=include_header
                )

                # 카드 안 가로 행 gap을 8px로 줄였으므로 컬럼 비율도 타이트하게
                clip_cols = st.columns([0.5, 1.7, 1.3, 3.7])
                with clip_cols[0]:
                    st.markdown('<div class="fn-static-label">복사</div>', unsafe_allow_html=True)
                with clip_cols[1]:
                    _clipboard_button(
                        clip_html,
                        button_label="📋 클립보드에 복사",
                        key=f"clip_{range_label}_{int(include_header)}",
                    )
                with clip_cols[2]:
                    st.checkbox(
                        "헤더 포함",
                        value=False,
                        key="clip_include_header",
                        help="체크하면 검수시트 컬럼 헤더 1행도 함께 복사돼요. 보통은 해제(데이터만 복사).",
                    )

        except Exception as e:
            st.error(f"엑셀 생성 실패: {e}")
