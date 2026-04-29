"""편성표 CSV → 검수시트 미리보기 변환 로직."""
from __future__ import annotations

import re
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

DAYS = ["월", "화", "수", "목", "금", "토", "일"]
DAY_LABEL = {d: d + "요일" for d in DAYS}
DAY_PREV = {"월": "일", "화": "월", "수": "화", "목": "수", "금": "목", "토": "금", "일": "토"}

# 검수시트 엑셀의 컬럼 헤더 (검수시트 예시 - 시트1.csv 1행 기준).
# 첫 컬럼은 빈 헤더(요일 라벨/종영 표시 들어가는 자리). 동일 이름 컬럼("담당자")이 중복 등장하지만
# 위치가 의미를 갖기 때문에 list 그대로 유지하고 인덱스로만 접근한다.
CHECKSHEET_HEADERS: list[str] = [
    "",                # 0  A열 (구분/요일/종영)
    "Tier",            # 1  B
    "원본검수",        # 2  C  ← 요일 헤더 행에서 신규 요약 텍스트가 들어가는 자리
    "담당자",          # 3  D
    "왓챠검수",        # 4  E
    "담당자",          # 5  F
    "series id",       # 6
    "series code",     # 7
    "series title",    # 8
    "season id",       # 9
    "season code",     # 10
    "season title",    # 11
    "mapping_type",    # 12
    "episode id",      # 13
    "episode code",    # 14
    "episode number",  # 15
    "formal_number",   # 16
    "episode title",   # 17
    "cp",              # 18
    "service_start",   # 19
    "service_end",     # 20
    "지원",
    "구분",
    "원소재화질",
    "서비스화질",
    "화질코멘트",
    "화질패널티",
    "성인물/키즈",
    "연령등급",
    "방통위/채널",
    "방통위/방영일",
    "영등위/관람등급번호",
    "주제",
    "선정성",
    "폭력성",
    "대사",
    "공포",
    "약물",
    "모방위험",
    "소재종류",
    "서비스 화면비",
    "화면비 비고",
    "음향",
    "자막",
    "더빙",
    "파일경로",        # 45
    "start offset(ms)",
    "end offset(ms)",
    "SoundTrack ( | 로 구분)",
    "VolumeOffset",
    "SoundLanguage ( | 로 구분)",
    "SDAR (교정화면비)",
    "SrcBucket",
    "비기닝(시간)",
    "엔딩(시간)",
    "beginning_sec",
    "opening_beginning_sec",
    "opening_ending_sec",
    "ending_sec",
    "nuvus_id",
    "비고 ",
    "백드롭",
    "짧은줄거리",
    "시즌 가로th",
    "시즌 세로th",
    "시즌 로고",
    "시리즈 가로th",
    "시리즈 세로th",
    "시리즈 로고",
    "담당자",          # 69 (마지막)
]

# 검수시트 컬럼 인덱스 상수
COL_A = 0
COL_TIER = 1
COL_NEW_SUMMARY = 2  # C열 - 요일 헤더 행에서 신규 요약 텍스트 위치
COL_SEASON_ID = 9
COL_SEASON_CODE = 10
COL_SEASON_TITLE = 11
COL_MAPPING_TYPE = 12
COL_EPISODE_NUMBER = 15
COL_FORMAL_NUMBER = 16
COL_CP = 18
COL_FILE_PATH = 45

# 미리보기에 표시할 컬럼 인덱스 (실제로 데이터가 채워지는 자리만).
PREVIEW_COL_INDICES = [
    COL_A,
    COL_TIER,
    COL_NEW_SUMMARY,
    COL_SEASON_ID,
    COL_SEASON_CODE,
    COL_SEASON_TITLE,
    COL_MAPPING_TYPE,
    COL_EPISODE_NUMBER,
    COL_FORMAL_NUMBER,
    COL_CP,
    COL_FILE_PATH,
]


def _empty_row() -> list[str]:
    return [""] * len(CHECKSHEET_HEADERS)


def parse_csv(file: Any) -> pd.DataFrame:
    df = pd.read_csv(file, dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _strip_mapping_suffix(title: str) -> tuple[str, str]:
    if title.endswith(" - 자막"):
        return title[: -len(" - 자막")], "basic"
    if title.endswith(" - 더빙"):
        return title[: -len(" - 더빙")], "dubbing"
    return title, ""


def _is_irregular(row: pd.Series) -> bool:
    yoil = str(row.get("요일", "") or "").strip()
    sigan = str(row.get("시간", "") or "").strip()
    return yoil == "비정기" or sigan == "비정기"


def _is_friday_ebs_special(row: pd.Series) -> bool:
    cp = str(row.get("CP bill", "") or "")
    yoil = str(row.get("요일", "") or "").strip()
    return cp.startswith("EBS") and yoil == ""


def _classify(row: pd.Series) -> str:
    bigo = str(row.get("편성비고", "") or "")
    bigo_a = str(row.get("비고", "") or "")
    e_n = str(row.get("이번주 e/n", "") or "").strip()
    sigan = str(row.get("시간", "") or "").strip()
    if any(k in bigo for k in ["결방", "홀드백"]) or e_n in {"결방", "홀드백"}:
        return "cancelled"
    if "신규" in bigo:
        return "new"
    if "전날 셋팅" in bigo_a or "전날셋팅" in bigo_a:
        return "prev_day"
    if sigan and sigan != "18:00":
        return "reserved"
    return "normal"


def _parse_hour(sigan: str) -> int:
    m = re.match(r"^(\d{1,2}):", sigan.strip())
    return int(m.group(1)) if m else -1


def _display_day(row: pd.Series, prev_normal_day: str | None = None) -> str:
    yoil = str(row.get("요일", "") or "").strip()
    sigan = str(row.get("시간", "") or "").strip()
    bigo = str(row.get("비고", "") or "").strip()
    # 연휴지연편성: 편성표 csv 위치상 직전의 일반 요일에 따라가도록
    if "연휴지연" in bigo:
        return prev_normal_day or yoil
    # 전날 셋팅: 강제로 전날
    if "전날 셋팅" in bigo or "전날셋팅" in bigo:
        return DAY_PREV.get(yoil, yoil)
    # 예약작 시간 < 12 → 전날
    cat = _classify(row)
    if cat == "reserved":
        h = _parse_hour(sigan)
        if 0 <= h < 12:
            return DAY_PREV.get(yoil, yoil)
    return yoil


def _compute_holiday_info(df: pd.DataFrame) -> tuple[set[str], dict[str, str]]:
    """편성표 비고에 '(기존 X에피)' 표기로 X 요일을 휴일로 식별.
    Returns (휴일 요일 set, 요일 → 'M/D' 매핑)."""
    holiday_days: set[str] = set()
    pat = re.compile(r"\(\s*기존\s*([월화수목금토일])\s*에피\s*\)")
    for _, row in df.iterrows():
        bigo = str(row.get("비고", "") or "")
        if "연휴지연" not in bigo:
            continue
        m = pat.search(bigo)
        if m:
            holiday_days.add(m.group(1))

    day_to_date: dict[str, str] = {}
    if "요일" in df.columns and "오픈일" in df.columns:
        for d in holiday_days:
            rows = df[df["요일"].astype(str).str.strip() == d]
            for _, row in rows.iterrows():
                opendate = str(row.get("오픈일", "") or "").strip()
                if not opendate:
                    continue
                nums = re.findall(r"\d+", opendate)
                if len(nums) >= 3:
                    day_to_date[d] = f"{int(nums[1])}/{int(nums[2])}"
                elif len(nums) == 2:
                    day_to_date[d] = f"{int(nums[0])}/{int(nums[1])}"
                if d in day_to_date:
                    break
    return holiday_days, day_to_date


def _compute_prev_normal_days(df: pd.DataFrame) -> list[str | None]:
    """편성표 csv 순서대로 각 행 직전의 '일반 요일'(연휴지연편성 X) 추적."""
    out: list[str | None] = []
    last: str | None = None
    for _, row in df.iterrows():
        bigo = str(row.get("비고", "") or "").strip()
        yoil = str(row.get("요일", "") or "").strip()
        if "연휴지연" not in bigo and yoil in DAYS:
            last = yoil
        out.append(last)
    return out


def check_jongyeong_alerts(df: pd.DataFrame) -> list[dict]:
    """`최종회차`보다 `이번주 e/n`이 더 큰 행 → 미고지 종영 가능성 알림.

    이미 편성비고에 '종영'이 적힌 행은 이미 처리됐다고 보고 알림 X.
    """
    alerts: list[dict] = []
    if "최종회차" not in df.columns:
        return alerts
    for _, row in df.iterrows():
        bigo = str(row.get("편성비고", "") or "")
        if "종영" in bigo:
            continue
        final_str = str(row.get("최종회차", "") or "")
        nums = re.findall(r"\d+", final_str)
        if not nums:
            continue
        # "13 (38-50화)" 같은 표기에서 마지막 회차 번호는 가장 큰 숫자(50)
        final_ep = max(int(n) for n in nums)
        e_n_str = str(row.get("이번주 e/n", "") or "")
        e_nums = re.findall(r"\d+", e_n_str)
        if not e_nums:
            continue
        max_e = max(int(n) for n in e_nums)
        if max_e > final_ep:
            title_raw = str(row.get("Title", "") or "").strip()
            title_clean, _mt = _strip_mapping_suffix(title_raw)
            alerts.append(
                {
                    "title": title_clean,
                    "final": final_ep,
                    "current_max": max_e,
                    "message": f"확인 필요 {title_clean} {final_ep}회차를 끝으로 종영인 것으로 확인됨",
                }
            )
    return alerts


def _format_new_summary(new_items: list[dict]) -> str:
    """`신규: 타이틀 e/n화 (CP bill) / 타이틀 e/n화 (CP bill) / ...` 형식 문자열."""
    if not new_items:
        return ""
    parts = []
    for r in new_items:
        title = r["Title"]
        e_n = str(r.get("이번주 e/n", "") or "").strip()
        cp = str(r.get("CP bill", "") or "").strip()
        ep = f"{e_n}화" if e_n else ""
        item = " ".join(s for s in [title, ep] if s)
        if cp:
            item = f"{item} ({cp})"
        parts.append(item)
    return "신규: " + " / ".join(parts)


def _split_episode_pairs(e_n: str, f_n: str) -> list[tuple[str, str]]:
    """`이번주 e/n`이 'X-Y' 범위면 X, X+1, ..., Y로 분할하고, f/n도 같은 개수로 짝지음."""
    e = str(e_n).strip()
    f = str(f_n).strip()
    m = re.match(r"^(\d+)-(\d+)$", e)
    if not m:
        return [(e, f)]
    a, b = int(m.group(1)), int(m.group(2))
    if b < a or (b - a) > 60:
        return [(e, f)]
    e_list = list(range(a, b + 1))
    fm = re.match(r"^(\d+)-(\d+)$", f)
    if fm:
        fa, fb = int(fm.group(1)), int(fm.group(2))
        f_list = list(range(fa, fb + 1))
        if len(f_list) != len(e_list):
            f_list = [""] * len(e_list)
    elif f == "":
        f_list = [""] * len(e_list)
    else:
        f_list = [f] + [""] * (len(e_list) - 1)
    return [(str(x), str(y)) for x, y in zip(e_list, f_list)]


def _calc_over_final(row: pd.Series, e_val: str, already_jongyeong: bool) -> bool:
    """이 분할 회차가 편성표 `최종회차` 숫자보다 큰지 (이미 '종영' 처리된 행은 제외)."""
    if already_jongyeong:
        return False
    final_str = str(row.get("최종회차", "") or "")
    final_nums = re.findall(r"\d+", final_str)
    if not final_nums:
        return False
    final_ep_val = max(int(n) for n in final_nums)
    e_nums = re.findall(r"\d+", str(e_val))
    if not e_nums:
        return False
    return max(int(n) for n in e_nums) > final_ep_val


def _row_values_for(row: pd.Series, e_val: str, f_val: str, jongyeong: bool) -> list[str]:
    vals = _empty_row()
    title_raw = str(row.get("Title", "") or "")
    title_clean, mt = _strip_mapping_suffix(title_raw)
    if jongyeong:
        vals[COL_A] = "종영"
    vals[COL_TIER] = str(row.get("Tier", "") or "")
    vals[COL_SEASON_ID] = str(row.get("id", "") or "")
    vals[COL_SEASON_CODE] = str(row.get("code", "") or "")
    vals[COL_SEASON_TITLE] = title_clean
    vals[COL_MAPPING_TYPE] = mt
    vals[COL_EPISODE_NUMBER] = e_val
    vals[COL_FORMAL_NUMBER] = f_val
    vals[COL_CP] = str(row.get("CP bill", "") or "")
    vals[COL_FILE_PATH] = str(row.get("원본경로", "") or "")
    return vals


def build_for_day(df: pd.DataFrame, target_day: str) -> dict:
    """선택된 요일 한 개에 대해 검수시트 미리보기 데이터 구조 반환."""
    df = df.copy()
    if "요일" not in df.columns:
        raise ValueError("CSV에 '요일' 컬럼이 없습니다. 헤더를 확인해주세요.")

    df = df[~df.apply(_is_irregular, axis=1)].reset_index(drop=True)

    ebs_mask = df.apply(_is_friday_ebs_special, axis=1)
    ebs_df = df[ebs_mask].copy().reset_index(drop=True)
    df = df[~ebs_mask].copy().reset_index(drop=True)

    prev_normal_days = _compute_prev_normal_days(df)
    holiday_days, day_to_date = _compute_holiday_info(df)
    is_target_holiday = target_day in holiday_days

    new_items: list[dict] = []
    reserved_rows: list[dict] = []
    # 연휴지연편성: 비고 텍스트별로 그룹화 → 출력 시 한 헤더 + 그 아래 콘텐츠
    holiday_groups: dict[str, list[dict]] = {}
    holiday_group_order: list[str] = []
    normal_rows: list[dict] = []
    prev_day_rows: list[dict] = []  # 비고 '전날 셋팅' 행 — 전날 맨 아래(결방 위)
    cancelled: list[dict] = []

    for idx, row in df.iterrows():
        cat = _classify(row)
        is_jongyeong = "종영" in str(row.get("편성비고", "") or "")
        bigo = str(row.get("비고", "") or "").strip()
        is_holiday = "연휴지연" in bigo
        prev = prev_normal_days[idx] if idx < len(prev_normal_days) else None
        display_d = _display_day(row, prev_normal_day=prev)

        if cat == "new":
            if display_d != target_day:
                continue
            title_raw = str(row.get("Title", "") or "")
            title_clean, _mt = _strip_mapping_suffix(title_raw)
            new_items.append(
                {
                    "Title": title_clean,
                    "이번주 e/n": str(row.get("이번주 e/n", "") or ""),
                    "CP bill": str(row.get("CP bill", "") or ""),
                }
            )
            continue

        if display_d != target_day:
            continue

        e_n = str(row.get("이번주 e/n", "") or "")
        f_n = str(row.get("이번주 f/n", "") or "")
        pairs = _split_episode_pairs(e_n, f_n)

        if is_holiday:
            # 비고 텍스트별로 그룹 헤더 + 콘텐츠 두 톤 회색
            if bigo not in holiday_groups:
                holiday_groups[bigo] = []
                holiday_group_order.append(bigo)
            is_split = len(pairs) > 1
            group_id = f"holi-{idx}" if is_split else None
            for i, (e_val, f_val) in enumerate(pairs):
                jongyeong_here = is_jongyeong and i == len(pairs) - 1
                vals = _row_values_for(row, e_val, f_val, False)
                vals[COL_A] = ""  # 콘텐츠 행 A열은 비워둠 (헤더 행에만 비고 텍스트)
                holiday_groups[bigo].append(
                    {
                        "values": vals,
                        "category": "holiday_body",
                        "time": str(row.get("시간", "") or ""),
                        "jongyeong": False,
                        "group_id": group_id,
                        "split_idx": i,
                        "split_total": len(pairs),
                        "over_final": _calc_over_final(row, e_val, is_jongyeong),
                    }
                )
            continue

        is_split = len(pairs) > 1
        group_id = f"grp-{idx}" if is_split else None
        for i, (e_val, f_val) in enumerate(pairs):
            jongyeong_here = is_jongyeong and i == len(pairs) - 1
            vals = _row_values_for(row, e_val, f_val, jongyeong_here)
            entry = {
                "values": vals,
                "category": cat,
                "time": str(row.get("시간", "") or ""),
                "jongyeong": jongyeong_here,
                "group_id": group_id,
                "split_idx": i,
                "split_total": len(pairs),
                "over_final": _calc_over_final(row, e_val, is_jongyeong),
            }
            if cat == "cancelled":
                cancelled.append(entry)
            elif cat == "reserved":
                reserved_rows.append(entry)
            elif cat == "prev_day":
                prev_day_rows.append(entry)
            else:
                normal_rows.append(entry)

    reserved_rows.sort(key=lambda r: _parse_hour(r["time"]))

    # 출력 순서: 예약작 → 일반 → 전날 셋팅(결방 위)
    # (결방 → 연휴지연 그룹은 build_xlsx / 미리보기에서 cancelled 뒤에 별도로 이어 붙임)
    body: list[dict] = reserved_rows + normal_rows + prev_day_rows

    # 연휴지연 그룹 — 비고 텍스트별로 (헤더 1행 + 콘텐츠들)
    holiday_block: list[dict] = []
    for bigo_text in holiday_group_order:
        h_vals = _empty_row()
        h_vals[COL_A] = bigo_text
        holiday_block.append(
            {
                "values": h_vals,
                "category": "holiday_header",
                "time": "",
                "jongyeong": False,
                "over_final": False,
            }
        )
        holiday_block.extend(holiday_groups[bigo_text])

    # 휴일 요일이면 그 요일 콘텐츠는 음영 X (요일 헤더 행만 형광 초록)
    if is_target_holiday:
        for entry in body + cancelled + holiday_block:
            entry["no_fill"] = True

    # 요일 헤더 행
    header_vals = _empty_row()
    if is_target_holiday:
        date_str = day_to_date.get(target_day, "")
        if date_str:
            header_vals[COL_A] = f"({DAY_LABEL[target_day]} {date_str} 휴일)"
        else:
            header_vals[COL_A] = f"({DAY_LABEL[target_day]} 휴일)"
    else:
        header_vals[COL_A] = DAY_LABEL.get(target_day, target_day)
    header_vals[COL_NEW_SUMMARY] = _format_new_summary(new_items)
    header_row = {
        "values": header_vals,
        "category": "header_holiday" if is_target_holiday else "header",
        "new_count": len(new_items),
    }

    # EBS 섹션 (금요일 한정)
    ebs_section: list[dict] = []
    if target_day == "금":
        for _, row in ebs_df.iterrows():
            title_raw = str(row.get("Title", "") or "")
            title_clean, mt = _strip_mapping_suffix(title_raw)
            vals = _empty_row()
            vals[COL_A] = str(row.get("편성비고", "") or "")
            vals[COL_TIER] = str(row.get("Tier", "") or "")
            vals[COL_SEASON_ID] = str(row.get("id", "") or "")
            vals[COL_SEASON_CODE] = str(row.get("code", "") or "")
            vals[COL_SEASON_TITLE] = title_clean
            vals[COL_MAPPING_TYPE] = mt
            vals[COL_CP] = str(row.get("CP bill", "") or "")
            # 원본경로/episode number/formal_number 빈칸 (스펙)
            ebs_section.append({"values": vals, "category": "ebs"})

    return {
        "headers": CHECKSHEET_HEADERS,
        "header_row": header_row,
        "body": body,
        "cancelled": cancelled,
        "holiday_block": holiday_block,
        "ebs_section": ebs_section,
        "new_count": len(new_items),
    }


# ── XLSX 출력 ────────────────────────────────────────────────────────────

# 행 카테고리별 음영 (구글 시트 검수시트와 자연스럽게 어울리는 톤)
_FILL_HEADER_DAY = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")  # 회색
_FILL_RESERVED = PatternFill(start_color="FFFFD6E5", end_color="FFFFD6E5", fill_type="solid")    # 핑크
_FILL_CANCELLED = PatternFill(start_color="FFFFD0D0", end_color="FFFFD0D0", fill_type="solid")   # 빨강
_FILL_EBS = PatternFill(start_color="FFECECEC", end_color="FFECECEC", fill_type="solid")
_FILL_EBS_BANNER = PatternFill(start_color="FFBDBDBD", end_color="FFBDBDBD", fill_type="solid")
_FILL_JONGYEONG = PatternFill(start_color="FFFFF3A3", end_color="FFFFF3A3", fill_type="solid")   # 노랑
_FILL_HOLIDAY_HEADER = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")  # 연휴지연 그룹 헤더 (요일 헤더와 같은 회색)
_FILL_HOLIDAY_BODY = PatternFill(start_color="FF9C9C9C", end_color="FF9C9C9C", fill_type="solid")    # 연휴지연 콘텐츠 (더 진한 회색)
_FILL_HEADER_HOLIDAY = PatternFill(start_color="FFA8F2A8", end_color="FFA8F2A8", fill_type="solid")  # 형광 초록 (휴일 요일 헤더)
_FILL_PREV_DAY = PatternFill(start_color="FFFFF4E0", end_color="FFFFF4E0", fill_type="solid")        # 전날 셋팅(옅은 살구)
_FILL_OVER_FINAL = PatternFill(start_color="FFBFDBFE", end_color="FFBFDBFE", fill_type="solid")      # 종영 의심 (행 전체 옅은 파랑)
_FILL_ALERT = PatternFill(start_color="FFFFE082", end_color="FFFFE082", fill_type="solid")           # 알림(노란-주황)


def _fill_for(category: str) -> PatternFill | None:
    if category == "header":
        return _FILL_HEADER_DAY
    if category == "reserved":
        return _FILL_RESERVED
    if category == "cancelled":
        return _FILL_CANCELLED
    if category == "ebs":
        return _FILL_EBS
    if category == "holiday_header":
        return _FILL_HOLIDAY_HEADER
    if category == "holiday_body":
        return _FILL_HOLIDAY_BODY
    if category == "header_holiday":
        return _FILL_HEADER_HOLIDAY
    if category == "prev_day":
        return _FILL_PREV_DAY
    return None


def build_xlsx(df: pd.DataFrame, days: list[str] | str) -> bytes:
    """선택한 요일(여러 개 가능)의 검수시트를 한 시트에 이어쓴 .xlsx 바이트.

    - days가 단일 요일 문자열이면 그 요일만, 리스트면 리스트 순서대로 이어 출력.
    - 각 요일마다 [요일 헤더 행 + 본문 + 결방] 순서. 금요일이면 그 뒤에 EBS 배너+섹션.
    """
    if isinstance(days, str):
        days = [days]
    if not days:
        days = ["월"]

    wb = Workbook()
    ws = wb.active
    ws.title = f"{days[0]}~{days[-1]}요일" if len(days) > 1 else f"{days[0]}요일"

    # 시트 상단에 종영 의심 알림 (있으면)
    alerts = check_jongyeong_alerts(df)
    for alert in alerts:
        a_vals = [""] * len(CHECKSHEET_HEADERS)
        a_vals[0] = "확인 필요"
        a_vals[2] = alert["message"]
        ws.append(a_vals)
        row_idx = ws.max_row
        for col_idx in range(1, len(CHECKSHEET_HEADERS) + 1):
            ws.cell(row=row_idx, column=col_idx).fill = _FILL_ALERT
        ws.cell(row=row_idx, column=1).font = Font(bold=True, color="FFC62828")
        ws.cell(row=row_idx, column=COL_NEW_SUMMARY + 1).alignment = Alignment(
            wrap_text=True, vertical="center"
        )

    def write_entry(entry: dict) -> None:
        ws.append(entry["values"])
        row_idx = ws.max_row
        if not entry.get("no_fill"):
            # over_final이면 카테고리 음영을 덮고 행 전체를 옅은 파랑으로
            fill = _FILL_OVER_FINAL if entry.get("over_final") else _fill_for(entry["category"])
            if fill is not None:
                for col_idx in range(1, len(entry["values"]) + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = fill
        if entry.get("jongyeong") and entry["category"] != "holiday":
            ws.cell(row=row_idx, column=1).fill = _FILL_JONGYEONG
            ws.cell(row=row_idx, column=1).font = Font(bold=True)
        if entry["category"] in ("header", "header_holiday"):
            ws.cell(row=row_idx, column=1).font = Font(bold=True)
            ws.cell(row=row_idx, column=COL_NEW_SUMMARY + 1).alignment = Alignment(
                wrap_text=True, vertical="center"
            )

    for d in days:
        result = build_for_day(df, d)
        write_entry(result["header_row"])
        for r in result["body"]:
            write_entry(r)
        for r in result["cancelled"]:
            write_entry(r)
        for r in result["holiday_block"]:
            write_entry(r)
        if d == "금" and result["ebs_section"]:
            banner_vals = [""] * len(CHECKSHEET_HEADERS)
            banner_vals[0] = "📺 EBS"
            ws.append(banner_vals)
            banner_row = ws.max_row
            for col_idx in range(1, len(CHECKSHEET_HEADERS) + 1):
                ws.cell(row=banner_row, column=col_idx).fill = _FILL_EBS_BANNER
            ws.cell(row=banner_row, column=1).font = Font(bold=True)
            for r in result["ebs_section"]:
                write_entry(r)

    width_map = {
        COL_A + 1: 10,
        COL_TIER + 1: 6,
        COL_SEASON_ID + 1: 11,
        COL_SEASON_CODE + 1: 12,
        COL_SEASON_TITLE + 1: 38,
        COL_MAPPING_TYPE + 1: 11,
        COL_EPISODE_NUMBER + 1: 10,
        COL_FORMAL_NUMBER + 1: 10,
        COL_CP + 1: 18,
        COL_FILE_PATH + 1: 50,
        COL_NEW_SUMMARY + 1: 70,
    }
    for col, w in width_map.items():
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = w

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def select_day_range(start: str, end: str) -> list[str]:
    """월~일 고정 순서 안에서 start~end 구간을 추출. start가 end보다 뒤면 swap."""
    if start not in DAYS or end not in DAYS:
        return []
    si, ei = DAYS.index(start), DAYS.index(end)
    if si > ei:
        si, ei = ei, si
    return DAYS[si : ei + 1]
