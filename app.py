"""검수시트 자동 생성기 - Streamlit UI."""
import html as _html
from datetime import date

import streamlit as st

from converter import (
    CHECKSHEET_HEADERS,
    DAYS,
    COL_A,
    COL_MAPPING_TYPE,
    COL_EPISODE_NUMBER,
    COL_FORMAL_NUMBER,
    PREVIEW_COL_INDICES,
    parse_csv,
    build_for_day,
    build_xlsx,
    select_day_range,
    check_jongyeong_alerts,
)

st.set_page_config(page_title="검수시트 자동 생성기", page_icon="🎬", layout="wide")

st.markdown(
    """
    <style>
    /* 다운로드 버튼: 단정한 다크 슬레이트 솔리드 */
    div[data-testid="stDownloadButton"] > button {
        background: #111827 !important;
        color: #ffffff !important;
        border: 1px solid #111827 !important;
        padding: 12px 24px !important;
        border-radius: 8px !important;
        font-size: 14px !important;
        font-weight: 500 !important;
        font-family: inherit !important;
        letter-spacing: 0 !important;
        box-shadow: none !important;
        transition: background 0.15s ease, border-color 0.15s ease;
        width: 100% !important;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background: #000000 !important;
        border-color: #000000 !important;
    }
    div[data-testid="stDownloadButton"] > button:active {
        background: #1f2937 !important;
    }
    div[data-testid="stDownloadButton"] {
        margin-top: 6px;
    }
    /* 콜아웃 카드 */
    .download-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 14px 18px;
        margin-top: 14px;
        margin-bottom: 8px;
        font-family: inherit;
    }
    .download-card .download-title {
        font-size: 14px;
        font-weight: 600;
        color: #0f172a;
        margin-bottom: 4px;
        font-family: inherit;
    }
    .download-card .download-sub {
        font-size: 12.5px;
        color: #475569;
        line-height: 1.55;
        font-family: inherit;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🎬 검수시트 자동 생성기")
st.caption(
    "주간편성표 CSV를 업로드하고 요일 범위를 선택하면, 검수시트 엑셀과 동일한 컬럼 구조로 변환된 미리보기·xlsx 다운로드가 제공됩니다."
)

with st.container(border=True):
    uploaded = st.file_uploader("편성표 CSV 업로드", type=["csv"])
    cols = st.columns([2, 2, 2, 3])
    with cols[0]:
        start_day = st.selectbox("시작 요일", DAYS, index=0)
    with cols[1]:
        end_day = st.selectbox("끝 요일", DAYS, index=DAYS.index(start_day))
    with cols[2]:
        run = st.button("변환 미리보기", type="primary", use_container_width=True)
    with cols[3]:
        st.caption(
            "월~일 순서 고정. 한 요일만 보고 싶으면 시작·끝을 같게 두세요. "
            "범위 안에 금요일이 있으면 EBS 섹션도 자동 포함됩니다."
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
    if over_final:
        return "#bfdbfe"  # 종영 의심 — 행 전체 옅은 파랑 (다른 음영 모두 덮음)
    if no_fill:
        return "#ffffff"
    if category == "header":
        return "#dcdcdc"
    if category == "header_holiday":
        return "#a8f2a8"  # 형광 초록 (휴일 요일 헤더)
    if category == "reserved":
        return "#ffd6e5"
    if category == "cancelled":
        return "#ffd0d0"
    if category == "ebs":
        return "#ececec"
    if category == "holiday_header":
        return "#dcdcdc"  # 그룹 헤더 (요일 헤더와 같은 회색)
    if category == "holiday_body":
        return "#9c9c9c"  # 콘텐츠 (더 진한 회색)
    if category == "prev_day":
        return "#fff4e0"  # 전날 셋팅 (옅은 살구색 — 일반과 결방 사이 구분용)
    return "#ffffff"


def _cell_html(val: str, col_idx: int, category: str) -> str:
    if col_idx == COL_A:
        if val == "종영":
            return f"<span style='color:#222;font-weight:600;background:#fff3a3;padding:1px 4px;border-radius:3px'>{_esc(val)}</span>"
        if category == "header":
            return f"<span style='font-weight:600'>{_esc(val)}</span>"
        return _esc(val)
    if col_idx == COL_MAPPING_TYPE and val:
        badge_color = "#1976d2" if val == "basic" else "#9c27b0"
        return (
            f"<span style='background:{badge_color};color:white;padding:1px 6px;"
            f"border-radius:8px;font-size:10px'>{_esc(val)}</span>"
        )
    return _esc(val)


def _render_preview(headers, results_by_day: list[tuple[str, dict]]) -> str:
    """여러 요일 결과를 한 표에 이어붙여 미리보기 (매핑된 컬럼만 표시)."""
    indices = PREVIEW_COL_INDICES
    parts = ["<div style='overflow-x:auto;border:1px solid #ccc'>"]
    parts.append(
        "<table style='border-collapse:collapse;font-size:12px;"
        "font-family:-apple-system,BlinkMacSystemFont,sans-serif;white-space:nowrap'>"
    )
    parts.append("<thead><tr style='background:#bdbdbd'>")
    for i in indices:
        parts.append(
            "<th style='border:1px solid #777;padding:2px 6px;font-size:10px;color:#222'>"
            + _esc(_column_letter(i))
            + "</th>"
        )
    parts.append("</tr><tr style='background:#e8e8e8'>")
    for i in indices:
        h = headers[i] if headers[i] else "(구분)"
        parts.append(
            "<th style='border:1px solid #999;padding:4px 8px;font-size:11px;"
            "font-weight:600;text-align:left'>" + _esc(h) + "</th>"
        )
    parts.append("</tr></thead><tbody>")

    def emit_row(entry):
        cat = entry["category"]
        no_fill = bool(entry.get("no_fill", False))
        is_over = bool(entry.get("over_final"))
        bg = _row_bg(cat, no_fill, is_over)
        parts.append(f"<tr style='background:{bg}'>")
        for i in indices:
            v = entry["values"][i]
            cell = _cell_html(v, i, cat)
            parts.append(f"<td style='border:1px solid #ddd;padding:4px 8px'>{cell}</td>")
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
            parts.append("<tr style='background:#bdbdbd'>")
            for k, _i in enumerate(indices):
                label = "📺 EBS" if k == 0 else ""
                parts.append(
                    f"<td style='border:1px solid #999;padding:3px 8px;font-size:11px;font-weight:600'>{_esc(label)}</td>"
                )
            parts.append("</tr>")
            for r in result["ebs_section"]:
                emit_row(r)

    parts.append("</tbody></table></div>")
    return "".join(parts)


if run:
    if uploaded is None:
        st.error("먼저 편성표 CSV 파일을 업로드해주세요.")
    else:
        try:
            df = parse_csv(uploaded)
        except Exception as e:
            st.error(f"CSV 파싱 실패: {e}")
            st.stop()

        days = select_day_range(start_day, end_day)
        if not days:
            st.error("요일 선택이 잘못되었습니다.")
            st.stop()

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
                    <div style='font-size:15px;font-weight:800;color:#b71c1c;margin-bottom:6px;'>
                        종영 의심 항목
                    </div>
                    <div style='font-size:12px;color:#555;margin-bottom:10px;line-height:1.6;'>
                        편성표의 <em>최종회차</em>보다 이번주 회차가 더 큽니다. 미리보기와 엑셀 시트 맨 위에 동일하게 표시되며,
                        해당 회차 셀은 <span style='color:#1565c0;font-weight:700'>파란색</span>으로 강조됩니다.
                    </div>
                    <ul style='margin:0;padding-left:22px;line-height:1.8;font-size:13.5px;color:#222;'>
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

        range_label = f"{days[0]}~{days[-1]}요일" if len(days) > 1 else f"{days[0]}요일"
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

        st.info("💡 행 단위 복사·붙여넣기는 아래 다운로드한 엑셀 파일에서 진행하세요. 미리보기는 데이터 검토용입니다.")

        try:
            xlsx_bytes = build_xlsx(df, days)
            today = date.today().isoformat()
            filename = f"검수시트_{today}_{range_label}.xlsx"
            st.markdown(
                """
                <div class="download-card">
                    <div class="download-title">검수시트 엑셀 준비 완료</div>
                    <div class="download-sub">
                        선택한 요일 범위가 한 시트에 이어 출력됩니다(요일 헤더 + 예약작 + 연휴지연편성 + 일반 + 전날 셋팅 + 결방, 금요일 포함 시 EBS 섹션). 행 단위로 그대로 복사해 검수시트에 붙여넣으면 돼요.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            dl_cols = st.columns([1, 2, 1])
            with dl_cols[1]:
                st.download_button(
                    label=f"{filename}  다운로드",
                    data=xlsx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxlsxformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        except Exception as e:
            st.error(f"엑셀 생성 실패: {e}")
