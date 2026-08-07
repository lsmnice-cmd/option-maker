# -*- coding: utf-8 -*-
"""
신화 업무 도구 통합 앱
 - 거래처원장 비교 (서식 자동 인식 · 일자별 대조)
 - 상품 중량 및 옵션가 자동 생성기
 - 배송 달력 배너 생성기
"""
import io
import re
from collections import defaultdict
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import xlwt

st.set_page_config(page_title="신화 업무 도구", layout="wide")

# ─────────────────────────────────────────────────────────────
# 도구 선택
# ─────────────────────────────────────────────────────────────
if "tool" not in st.session_state:
    st.session_state.tool = None

TOOLS = {
    "ledger": "📑 거래처원장 비교",
    "option": "⚖️ 상품 중량 및 옵션가 자동 생성기",
    "banner": "🗓️ 배송 달력 배너 생성기",
}


def go_home():
    st.session_state.tool = None


if st.session_state.tool is None:
    st.title("신화 업무 도구")
    st.caption("사용할 도구를 선택하세요")
    st.markdown("---")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("📑 거래처원장 비교")
        st.write(
            "서식이 서로 다른 원장 두 개를 대조합니다.  \n"
            "상품명·상품코드는 무시하고 **일자별 중량·금액**이 맞는지 확인합니다."
        )
        if st.button("거래처원장 비교 열기", type="primary", use_container_width=True):
            st.session_state.tool = "ledger"
            st.rerun()

    with c2:
        st.subheader("⚖️ 상품 중량 및 옵션가 자동 생성기")
        st.write(
            "기준가와 단가를 입력해 중량별 옵션가를 자동 계산합니다.  \n"
            "**네이버 추가상품 서식**과 기존 표준 서식을 모두 지원합니다."
        )
        if st.button("옵션가 생성기 열기", type="primary", use_container_width=True):
            st.session_state.tool = "option"
            st.rerun()

    with c3:
        st.subheader("🗓️ 배송 달력 배너 생성기")
        st.write(
            "날짜를 클릭해 휴무·재개·마감일을 표시하고  \n"
            "스마트스토어용 **배송 안내 달력 배너**를 PNG/JPG로 저장합니다."
        )
        if st.button("배너 생성기 열기", type="primary", use_container_width=True):
            st.session_state.tool = "banner"
            st.rerun()

    st.stop()

with st.sidebar:
    st.markdown("### 도구")
    pick = st.radio("이동", list(TOOLS.keys()),
                    format_func=lambda k: TOOLS[k],
                    index=list(TOOLS.keys()).index(st.session_state.tool),
                    label_visibility="collapsed")
    if pick != st.session_state.tool:
        st.session_state.tool = pick
        st.rerun()
    st.markdown("---")
    st.button("🏠 메인 화면으로", on_click=go_home, use_container_width=True)


# ═════════════════════════════════════════════════════════════
# 도구 1 — 거래처원장 비교 (서식 자동 인식)
# ═════════════════════════════════════════════════════════════
SHINHWA_COLS = ["월일", "상품명", "원산지", "Box", "Kg",
                "매입단가", "매입공급가", "매입부가세", "매입합계",
                "매출단가", "매출공급가", "매출부가세", "매출합계",
                "지급액", "수금액", "미수금액", "X1", "X2", "X3"]

FORMAT_LABEL = {
    "shinhwa": "신화 표준 원장 (월/일 · 상품명 · 매입/매출)",
    "partner": "거래처 ERP 원장 (일자 · 적요 · 중량(kg) · 매출금액)",
}


def _cell(v):
    return "" if pd.isna(v) else str(v).strip()


def detect_format(raw):
    """헤더 행을 위에서부터 훑어 서식을 판별한다. → (서식코드, 헤더행index)"""
    for i in range(min(25, len(raw))):
        vals = [_cell(v) for v in raw.iloc[i].tolist()]
        joined = " ".join(vals)
        if "월/일" in joined and "상품명" in joined:
            return "shinhwa", i
        if "일자" in vals and any("중량" in v for v in vals):
            return "partner", i
    return None, None


def load_shinhwa(raw, hr):
    """신화 표준 원장 → 표준 컬럼(날짜·상품명·중량·단가·금액·수금·미수)"""
    ncol = min(raw.shape[1], len(SHINHWA_COLS))
    df = raw.iloc[hr + 1:, :ncol].reset_index(drop=True).copy()
    df.columns = SHINHWA_COLS[:ncol]
    df["엑셀행"] = df.index + hr + 2
    for c in SHINHWA_COLS:
        if c not in df.columns:
            df[c] = np.nan

    df["월일"] = df["월일"].astype(str).str.strip()
    df = df[df["월일"].str.match(r"\d{4}[/.\-]\d{1,2}[/.\-]\d{1,2}")].reset_index(drop=True)
    for c in ["Kg", "매입단가", "매입합계", "매출단가", "매출합계",
              "미수금액", "수금액", "지급액"]:
        df[c] = pd.to_numeric(df.get(c), errors="coerce")

    out = pd.DataFrame({
        "날짜": pd.to_datetime(df["월일"], errors="coerce").dt.strftime("%Y-%m-%d"),
        "상품명": df["상품명"].fillna("").astype(str).str.strip(),
        "중량": df["Kg"].fillna(0.0),
        "단가": df["매입단가"].fillna(df["매출단가"]).fillna(0.0),
        "금액": df["매입합계"].fillna(df["매출합계"]).fillna(0.0),
        "수금": df["수금액"].fillna(0.0) + df["지급액"].fillna(0.0),
        "미수": df["미수금액"],
        "엑셀행": df["엑셀행"],
    })
    return out.dropna(subset=["날짜"]).reset_index(drop=True)


def load_partner(raw, hr):
    """거래처 ERP 원장(일자·적요·중량(kg)·단가·매출금액·수금액·매입금액·출금액·잔액)"""
    hdr = [_cell(v) for v in raw.iloc[hr].tolist()]

    def find(*keys):
        for i, h in enumerate(hdr):          # 1) 완전 일치 우선
            if h in keys:
                return i
        for i, h in enumerate(hdr):          # 2) 부분 일치
            if any(k in h for k in keys):
                return i
        return None

    i_date, i_note = find("일자"), find("적요")
    i_w, i_p = find("중량(kg)", "중량"), find("단가")
    i_sale, i_buy = find("매출금액"), find("매입금액")
    i_in, i_out = find("수금액"), find("출금액")
    i_bal = find("잔액")

    df = raw.iloc[hr + 1:].reset_index(drop=True).copy()
    df["엑셀행"] = df.index + hr + 2

    d = pd.to_numeric(df[i_date], errors="coerce")
    keep = d.notna() & d.fillna(0).astype("int64").astype(str).str.match(r"\d{8}$")
    df, d = df[keep].copy(), d[keep].astype("int64").astype(str)

    def num(idx):
        if idx is None:
            return pd.Series(0.0, index=df.index)
        return pd.to_numeric(df[idx], errors="coerce").fillna(0.0)

    sale, buy, w = num(i_sale), num(i_buy), num(i_w)
    is_buy = (buy != 0) & (sale == 0)          # 반품/입고 → 부호 반전

    out = pd.DataFrame({
        "날짜": pd.to_datetime(d, format="%Y%m%d", errors="coerce")
                 .dt.strftime("%Y-%m-%d").values,
        "상품명": (df[i_note].fillna("").astype(str).str.strip().values
                if i_note is not None else ""),
        "중량": np.where(is_buy, -w, w),
        "단가": num(i_p).values,
        "금액": (sale - buy).values,
        "수금": (num(i_in) - num(i_out)).values,
        "미수": (pd.to_numeric(df[i_bal], errors="coerce").values
                if i_bal is not None else np.nan),
        "엑셀행": df["엑셀행"].values,
    })
    return out.dropna(subset=["날짜"]).reset_index(drop=True)


def ledger_load(file):
    """업로드된 원장 파일 → (서식코드, 표준화된 DataFrame)"""
    raw = pd.read_excel(file, sheet_name=0, header=None)
    fmt, hr = detect_format(raw)
    if fmt == "shinhwa":
        return fmt, load_shinhwa(raw, hr)
    if fmt == "partner":
        return fmt, load_partner(raw, hr)
    raise ValueError("원장 서식을 인식하지 못했습니다. "
                     "'월/일 + 상품명' 또는 '일자 + 중량' 헤더가 있어야 합니다.")


# ── 일자별 집계 대조 ─────────────────────────────────────────
def daily_frame(df):
    g = df.groupby("날짜").agg(중량=("중량", "sum"), 금액=("금액", "sum"),
                              수금=("수금", "sum"))
    live = df[(df["중량"] != 0) | (df["금액"] != 0)]
    g = g.join(live.groupby("날짜").size().rename("건수"), how="left")
    g["건수"] = g["건수"].fillna(0)
    return g


def compare_daily(a, b, tol_w, tol_amt, use_cash):
    da, db = daily_frame(a), daily_frame(b)
    sa, sb = set(da.index), set(db.index)
    m = da.join(db, how="outer", lsuffix="_A", rsuffix="_B").fillna(0.0)
    m["중량차"] = (m["중량_A"] - m["중량_B"]).round(3)
    m["금액차"] = (m["금액_A"] - m["금액_B"]).round(0)
    m["수금차"] = (m["수금_A"] - m["수금_B"]).round(0)
    m["건수차"] = m["건수_A"] - m["건수_B"]

    def verdict(row):
        d = row.name
        if d not in sb:
            return "B파일 없음"
        if d not in sa:
            return "A파일 없음"
        bad = []
        if abs(row["중량차"]) > tol_w:
            bad.append("중량")
        if abs(row["금액차"]) > tol_amt:
            bad.append("금액")
        if use_cash and abs(row["수금차"]) > tol_amt:
            bad.append("수금")
        return "일치" if not bad else " · ".join(bad) + " 차이"

    m["판정"] = m.apply(verdict, axis=1)
    m = m.reset_index().rename(columns={"index": "날짜"})
    return m[["날짜", "중량_A", "중량_B", "중량차", "금액_A", "금액_B", "금액차",
              "수금_A", "수금_B", "수금차", "건수_A", "건수_B", "건수차", "판정"]]


# ── 품목 짝짓기 (상품명 무시 · 숫자로 매칭) ──────────────────
PAIR_ORDER = {"A파일 없음": 0, "B파일 없음": 1, "중량·금액 다름": 2,
              "금액 다름": 3, "중량 다름": 4, "일치": 9}


def match_day(la, lb, tol_w, tol_amt):
    """같은 날짜의 A행·B행을 숫자만 보고 짝지어 (A행, B행, 구분) 목록을 만든다."""
    used = set()
    pairs = []

    def take(cond, score, label):
        for i, r in enumerate(la):
            if r.get("_done"):
                continue
            best, best_s = None, None
            for j, s in enumerate(lb):
                if j in used or not cond(r, s):
                    continue
                v = score(r, s)
                if best_s is None or v < best_s:
                    best, best_s = j, v
            if best is not None:
                used.add(best)
                r["_done"] = True
                pairs.append((r, lb[best], label))

    dw = lambda r, s: abs(r["중량"] - s["중량"])
    da = lambda r, s: abs(r["금액"] - s["금액"])
    dp = lambda r, s: abs(r["단가"] - s["단가"])

    # 1) 중량·금액 모두 허용오차 안 → 일치
    take(lambda r, s: dw(r, s) <= tol_w and da(r, s) <= tol_amt,
         lambda r, s: dw(r, s) + da(r, s) / 1e6, "일치")
    # 2) 금액은 같은데 중량이 다름
    take(lambda r, s: da(r, s) <= tol_amt,
         lambda r, s: dw(r, s), "중량 다름")
    # 3) 중량은 같은데 금액이 다름 (단가 오타·누락)
    take(lambda r, s: dw(r, s) <= tol_w,
         lambda r, s: da(r, s), "금액 다름")
    # 4) 단가가 같음 → 같은 품목인데 물량이 다름
    take(lambda r, s: dp(r, s) <= 1,
         lambda r, s: dw(r, s) + da(r, s) / 1e6, "중량·금액 다름")

    for r in la:
        if not r.get("_done"):
            pairs.append((r, None, "B파일 없음"))
    for j, s in enumerate(lb):
        if j not in used:
            pairs.append((None, s, "A파일 없음"))
    return pairs


def build_pairs(a, b, tol_w, tol_amt, overlap_only):
    """전체 기간을 날짜별로 짝지어 비교 표를 만든다."""
    item_a = a[(a["중량"] != 0) | (a["금액"] != 0)]
    item_b = b[(b["중량"] != 0) | (b["금액"] != 0)]
    dates = sorted(set(item_a["날짜"]) & set(item_b["날짜"])) if overlap_only \
        else sorted(set(item_a["날짜"]) | set(item_b["날짜"]))

    rows = []
    for d in dates:
        la = item_a[item_a["날짜"] == d].to_dict("records")
        lb = item_b[item_b["날짜"] == d].to_dict("records")
        for x, y, label in match_day(la, lb, tol_w, tol_amt):
            rows.append({
                "날짜": d, "구분": label,
                "A행": x["엑셀행"] if x else None,
                "A상품명": x["상품명"] if x else "── 없음 ──",
                "A중량": x["중량"] if x else None,
                "A단가": x["단가"] if x else None,
                "A금액": x["금액"] if x else None,
                "B행": y["엑셀행"] if y else None,
                "B상품명": y["상품명"] if y else "── 없음 ──",
                "B중량": y["중량"] if y else None,
                "B단가": y["단가"] if y else None,
                "B금액": y["금액"] if y else None,
                "중량차": round((x["중량"] if x else 0) - (y["중량"] if y else 0), 2),
                "금액차": round((x["금액"] if x else 0) - (y["금액"] if y else 0)),
            })
    df = pd.DataFrame(rows)
    if not df.empty:
        df["_o"] = df["구분"].map(PAIR_ORDER)
        df = df.sort_values(["날짜", "_o"]).drop(columns="_o").reset_index(drop=True)
    return df


# ── 표시 서식 ───────────────────────────────────────────────
DAILY_FMT = {
    "중량_A": "{:,.2f}", "중량_B": "{:,.2f}", "중량차": "{:,.2f}",
    "금액_A": "{:,.0f}", "금액_B": "{:,.0f}", "금액차": "{:,.0f}",
    "수금_A": "{:,.0f}", "수금_B": "{:,.0f}", "수금차": "{:,.0f}",
    "건수_A": "{:.0f}", "건수_B": "{:.0f}", "건수차": "{:.0f}",
}

PAIR_FMT = {
    "A행": "{:.0f}", "B행": "{:.0f}",
    "A중량": "{:,.2f}", "A단가": "{:,.0f}", "A금액": "{:,.0f}",
    "B중량": "{:,.2f}", "B단가": "{:,.0f}", "B금액": "{:,.0f}",
    "중량차": "{:,.2f}", "금액차": "{:,.0f}",
}

PAIR_COLOR = {
    "A파일 없음": "#ffd6d6", "B파일 없음": "#ffd6d6",
    "중량·금액 다름": "#ffe2c2", "금액 다름": "#fff3c4",
    "중량 다름": "#e8f0fe", "일치": "#f2f7f3",
}

LINE_FMT = {"중량": "{:,.2f}", "단가": "{:,.0f}", "금액": "{:,.0f}",
            "수금": "{:,.0f}", "엑셀행": "{:.0f}"}


def color_daily(row):
    v = str(row.get("판정", ""))
    if "없음" in v:
        c = "background-color:#ffd6d6"
    elif "차이" in v:
        c = "background-color:#fff3c4"
    else:
        c = "background-color:#e9f7ec"
    return [c] * len(row)


def color_pair(row):
    c = PAIR_COLOR.get(str(row.get("구분", "")), "")
    return [f"background-color:{c}" if c else ""] * len(row)


def run_ledger():
    st.title("📑 거래처원장 비교")
    st.caption("서식이 달라도 됩니다 · 상품명은 무시하고 **중량·단가·금액**만 보고 "
               "행끼리 짝지어, 어느 건이 어떻게 다른지 나란히 보여줍니다.")

    c1, c2 = st.columns(2)
    with c1:
        fa = st.file_uploader("A 파일 (예: 신화미트 원장)", type=["xlsx", "xls"], key="lg_a")
    with c2:
        fb = st.file_uploader("B 파일 (예: 거래처가 보내준 원장)", type=["xlsx", "xls"], key="lg_b")

    if not (fa and fb):
        st.info("👆 비교할 원장 파일 2개를 업로드하세요. "
                "신화 표준 원장과 거래처 ERP 원장 서식을 자동으로 구분합니다.")
        return

    try:
        fmt_a, a = ledger_load(fa)
        fmt_b, b = ledger_load(fb)
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        return

    if a.empty or b.empty:
        st.error("데이터 행을 찾지 못했습니다. 원장 양식인지 확인하세요.")
        return

    i1, i2 = st.columns(2)
    i1.success(f"**A 서식 인식:** {FORMAT_LABEL[fmt_a]}  \n"
               f"{len(a):,}행 · {a['날짜'].min()} ~ {a['날짜'].max()}")
    i2.success(f"**B 서식 인식:** {FORMAT_LABEL[fmt_b]}  \n"
               f"{len(b):,}행 · {b['날짜'].min()} ~ {b['날짜'].max()}")

    overlap = sorted(set(a["날짜"]) & set(b["날짜"]))
    if not overlap:
        st.warning("⚠️ 두 파일에 **겹치는 날짜가 하나도 없습니다.** "
                   "출력 기간이 서로 다른 원장인지 확인하세요.")

    st.markdown("---")
    o1, o2, o3 = st.columns([1, 1, 2])
    tol_w = o1.number_input("중량 허용오차 (kg)", min_value=0.0, value=0.01,
                            step=0.01, format="%.2f")
    tol_amt = o2.number_input("금액 허용오차 (원)", min_value=0, value=10, step=10)
    with o3:
        overlap_only = st.checkbox("겹치는 날짜만 대조", value=True)
        sort_amt = st.checkbox("금액차 큰 순으로 정렬", value=False)

    pairs = build_pairs(a, b, tol_w, tol_amt, overlap_only)
    diff = pairs[pairs["구분"] != "일치"] if not pairs.empty else pairs
    daily = compare_daily(a, b, tol_w, tol_amt, False)
    bad_days = daily[daily["판정"] != "일치"]

    st.markdown("---")
    m = st.columns(4)
    m[0].metric("대조한 건수", f"{len(pairs):,}건", f"차이 {len(diff):,}건")
    m[1].metric("차이 나는 날짜", f"{len(set(diff['날짜'])) if not diff.empty else 0}일")
    m[2].metric("중량 차이 합", f"{diff['중량차'].sum() if not diff.empty else 0:,.2f} kg")
    m[3].metric("금액 차이 합", f"{diff['금액차'].sum() if not diff.empty else 0:,.0f} 원")

    t1, t2, t3, t4 = st.tabs(["🎯 차이 품목 대조", "📅 일자별 대조",
                              "🔍 날짜별 전체 보기", "📊 요약"])

    with t1:
        if pairs.empty:
            st.info("대조할 데이터가 없습니다. (겹치는 날짜가 없을 수 있습니다)")
        elif diff.empty:
            st.success("✅ 모든 건이 허용오차 안에서 1:1로 맞습니다.")
        else:
            f1, f2 = st.columns([1, 3])
            day_opt = ["전체"] + sorted(set(diff["날짜"]))
            pick_day = f1.selectbox("날짜", day_opt)
            kinds = sorted(set(diff["구분"]), key=lambda k: PAIR_ORDER.get(k, 9))
            pick_kind = f2.multiselect("보고 싶은 유형", kinds, default=kinds)

            view = diff.copy()
            if pick_day != "전체":
                view = view[view["날짜"] == pick_day]
            if pick_kind:
                view = view[view["구분"].isin(pick_kind)]
            if sort_amt:
                view = view.reindex(view["금액차"].abs()
                                    .sort_values(ascending=False).index)

            st.dataframe(view.style.apply(color_pair, axis=1)
                         .format(PAIR_FMT, na_rep=""),
                         use_container_width=True, height=560, hide_index=True)
            st.caption(
                "🟥 **A/B파일 없음** — 한쪽에만 있는 건(누락·중복 입력) · "
                "🟧 **중량·금액 다름** — 단가만 같고 물량이 다름 · "
                "🟨 **금액 다름** — 중량은 같은데 금액이 다름(단가 차이) · "
                "🟦 **중량 다름** — 금액은 같은데 중량이 다름  \n"
                "상품명은 매칭에 쓰지 않고 참고 표시만 합니다. "
                "수금·입금만 있는 행은 이 표에서 제외됩니다."
            )

    with t2:
        if bad_days.empty:
            st.success("✅ 날짜별 중량·금액 합계가 모두 일치합니다.")
        else:
            st.dataframe(bad_days.style.apply(color_daily, axis=1)
                         .format(DAILY_FMT, na_rep=""),
                         use_container_width=True, height=520, hide_index=True)
            st.caption("🟥 한쪽 파일에 날짜 자체가 없음 · 🟨 중량/금액 합계가 다름")

    with t3:
        cand = overlap if overlap else sorted(set(a["날짜"]) | set(b["날짜"]))
        if not cand:
            st.info("표시할 날짜가 없습니다.")
        else:
            sel = st.selectbox("확인할 날짜", cand, key="detail_day")
            ra = a[a["날짜"] == sel][["엑셀행", "상품명", "중량", "단가", "금액", "수금"]]
            rb = b[b["날짜"] == sel][["엑셀행", "상품명", "중량", "단가", "금액", "수금"]]
            d1, d2 = st.columns(2)
            with d1:
                st.markdown(f"**A파일 — {len(ra)}건 · "
                            f"{ra['중량'].sum():,.2f}kg · {ra['금액'].sum():,.0f}원**")
                st.dataframe(ra.style.format(LINE_FMT, na_rep=""),
                             use_container_width=True, height=430, hide_index=True)
            with d2:
                st.markdown(f"**B파일 — {len(rb)}건 · "
                            f"{rb['중량'].sum():,.2f}kg · {rb['금액'].sum():,.0f}원**")
                st.dataframe(rb.style.format(LINE_FMT, na_rep=""),
                             use_container_width=True, height=430, hide_index=True)
            st.caption(f"차이 → 중량 {ra['중량'].sum() - rb['중량'].sum():,.2f}kg · "
                       f"금액 {ra['금액'].sum() - rb['금액'].sum():,.0f}원")

    with t4:
        def lastv(df):
            s = df["미수"].dropna()
            return float(s.iloc[-1]) if len(s) else 0.0

        kind_cnt = (diff["구분"].value_counts().to_dict() if not diff.empty else {})
        summary = pd.DataFrame([
            {"항목": "서식", "A파일": FORMAT_LABEL[fmt_a],
             "B파일": FORMAT_LABEL[fmt_b], "차이": ""},
            {"항목": "행 수", "A파일": f"{len(a):,}", "B파일": f"{len(b):,}",
             "차이": f"{len(a) - len(b):,}"},
            {"항목": "기간", "A파일": f"{a['날짜'].min()} ~ {a['날짜'].max()}",
             "B파일": f"{b['날짜'].min()} ~ {b['날짜'].max()}",
             "차이": f"겹치는 날짜 {len(overlap)}일"},
            {"항목": "중량 합계", "A파일": f"{a['중량'].sum():,.2f}",
             "B파일": f"{b['중량'].sum():,.2f}",
             "차이": f"{a['중량'].sum() - b['중량'].sum():,.2f}"},
            {"항목": "금액 합계", "A파일": f"{a['금액'].sum():,.0f}",
             "B파일": f"{b['금액'].sum():,.0f}",
             "차이": f"{a['금액'].sum() - b['금액'].sum():,.0f}"},
            {"항목": "수금 합계", "A파일": f"{a['수금'].sum():,.0f}",
             "B파일": f"{b['수금'].sum():,.0f}",
             "차이": f"{a['수금'].sum() - b['수금'].sum():,.0f}"},
            {"항목": "최종 잔액(미수)", "A파일": f"{lastv(a):,.0f}",
             "B파일": f"{lastv(b):,.0f}", "차이": f"{lastv(a) - lastv(b):,.0f}"},
        ] + [{"항목": f"[{k}]", "A파일": f"{v}건", "B파일": "", "차이": ""}
             for k, v in sorted(kind_cnt.items(),
                                key=lambda kv: PAIR_ORDER.get(kv[0], 9))])
        st.dataframe(summary, use_container_width=True, hide_index=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        (diff if not diff.empty else pd.DataFrame({"결과": ["차이 없음"]})) \
            .to_excel(w, sheet_name="차이 품목", index=False)
        if not pairs.empty:
            pairs.to_excel(w, sheet_name="전체 대조", index=False)
        daily.to_excel(w, sheet_name="일자별 대조", index=False)
        a.to_excel(w, sheet_name="A파일 정규화", index=False)
        b.to_excel(w, sheet_name="B파일 정규화", index=False)
    st.download_button("💾 비교결과 다운로드 (xlsx)", buf.getvalue(),
                       f"원장비교_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═════════════════════════════════════════════════════════════
# 도구 2 — 상품 중량 및 옵션가 자동 생성기
# ═════════════════════════════════════════════════════════════
def naver_to_internal(df_naver):
    rows = []
    for _, r in df_naver.iterrows():
        rows.append({
            '품목': str(r['추가상품명']),
            '중량': str(r['추가상품값']),
            '옵션가': r.get('추가상품가', 0),
            '재고수량': r.get('재고수량', 0),
            '관리코드': r.get('관리코드', ''),
            '사용여부': r.get('사용여부', 'Y'),
        })
    return pd.DataFrame(rows)


def internal_to_naver(df_internal, col_item_name, col_weight_name):
    return pd.DataFrame({
        '추가상품명': df_internal[col_item_name],
        '추가상품값': df_internal[col_weight_name],
        '추가상품가': df_internal['옵션가'],
        '재고수량': df_internal['재고수량'],
        '사용여부': df_internal.get('사용여부', 'Y'),
        '관리코드': df_internal.get('관리코드', ''),
    })


def run_option():
    ss = st.session_state
    ss.setdefault('processed_data', None)
    ss.setdefault('last_file_id', None)
    ss.setdefault('col_item_name', None)
    ss.setdefault('col_weight_name', None)
    ss.setdefault('history', [])
    ss.setdefault('global_base_price', 0)
    ss.setdefault('last_selected_item', None)
    ss.setdefault('reset_counter', 0)
    ss.setdefault('file_format', None)

    st.title("⚖️ 상품 중량 및 옵션가 자동 생성기")
    st.caption("다중 품목 지원 · 네이버 추가상품 서식 자동 인식")

    uploaded_file = st.file_uploader(
        "기존 양식 파일(xls, xlsx, csv) 또는 네이버 추가상품 파일을 업로드하세요",
        type=['xls', 'xlsx', 'csv'], key="opt_file")

    if uploaded_file:
        current_file_id = getattr(uploaded_file, 'file_id',
                                  uploaded_file.name + str(uploaded_file.size))
        if ss.last_file_id != current_file_id:
            try:
                for key in ['base_price', 'global_base_price_input']:
                    if key in ss:
                        del ss[key]
                ss.global_base_price = 0
                ss.last_selected_item = None
                ss.reset_counter += 1

                if uploaded_file.name.endswith('.csv'):
                    file_bytes = uploaded_file.read()
                    df = None
                    for enc in ['utf-8', 'cp949', 'euc-kr', 'utf-8-sig']:
                        try:
                            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
                            df.columns = df.columns.str.strip()
                            if ('품목 및 등급' in df.columns or '품목' in df.columns
                                    or '추가상품명' in df.columns):
                                break
                        except Exception:
                            continue
                else:
                    engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
                    df = pd.read_excel(uploaded_file, engine=engine)
                    df.columns = df.columns.str.strip()

                if df is None:
                    st.error("파일을 제대로 읽지 못했습니다.")
                    st.stop()

                if '추가상품명' in df.columns and '추가상품값' in df.columns:
                    ss.file_format = 'naver'
                    df = naver_to_internal(df)
                    col_name, col_weight = '품목', '중량'
                    st.info("📌 네이버 추가상품 서식으로 인식했습니다. "
                            "내부 변환 후 처리하며, 다운로드 시 네이버 서식으로 저장됩니다.")
                else:
                    ss.file_format = 'standard'
                    if '품목 및 등급' in df.columns:
                        col_name = '품목 및 등급'
                    elif '품목' in df.columns:
                        col_name = '품목'
                    else:
                        st.error("'품목 및 등급', '품목', 또는 '추가상품명' 열(A열)을 찾을 수 없습니다.")
                        st.stop()

                    if '중량' in df.columns:
                        col_weight = '중량'
                    elif '포장&중량' in df.columns:
                        col_weight = '포장&중량'
                    else:
                        st.error("'중량' 또는 '포장&중량' 열(B열)을 찾을 수 없습니다.")
                        st.stop()

                df['__sort_1'] = range(len(df))
                df['__sort_2'] = 0.0
                ss.col_item_name = col_name
                ss.col_weight_name = col_weight
                ss.processed_data = df.copy()
                ss.last_file_id = current_file_id
                ss.history = []
                st.success("파일이 성공적으로 로드되었습니다! 아래에서 기준가를 먼저 입력해주세요.")
            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
                st.stop()

    if ss.processed_data is None:
        st.info("👆 파일을 업로드하면 작업을 시작할 수 있습니다.")
        return

    st.markdown("---")
    st.subheader("⚡ 기준가 설정 (필수)")
    col_bp1, col_bp2 = st.columns([2, 3])
    with col_bp1:
        entered_base_price = st.number_input("🚨 기준가(원)를 입력하세요", min_value=0,
                                             value=ss.global_base_price, step=100,
                                             key="global_base_price_input")
    with col_bp2:
        if entered_base_price > 0:
            st.success(f"✅ 기준가 **{entered_base_price:,}원** 이 설정되었습니다. "
                       "아래에서 품목별 작업을 진행하세요.")
        else:
            st.warning("⚠️ 기준가를 입력해야 이후 작업을 진행할 수 있습니다.")

    ss.global_base_price = entered_base_price
    if ss.global_base_price == 0:
        st.info("👆 기준가를 입력하면 품목 선택 및 중량 관리 기능이 활성화됩니다.")
        return

    df = ss.processed_data
    col_item_name = ss.col_item_name
    col_weight_name = ss.col_weight_name

    st.markdown("---")
    col_title, col_undo = st.columns([3, 1])
    with col_title:
        fmt_label = "네이버 추가상품" if ss.file_format == 'naver' else "기존 표준"
        st.subheader(f"1. 수정할 품목 선택 및 단가 설정  ({fmt_label} 서식)")
    with col_undo:
        if st.button("⏪ 방금 한 작업 되돌리기 (Undo)", disabled=not ss.history):
            ss.processed_data = ss.history.pop()
            st.success("이전 상태로 되돌렸습니다!")
            st.rerun()

    unique_items = df[col_item_name].dropna().unique()
    selected_item = st.selectbox(f"A열({col_item_name})에서 수정할 항목을 선택하세요", unique_items)

    if ss.get('last_selected_item') != selected_item:
        ss.reset_counter += 1
        ss.last_selected_item = selected_item

    naver_price_match = re.search(r'kg\s*(\d{3,})', str(selected_item))
    std_price_match = re.search(r'(\d{1,3}(?:,\d{3})*|\d+)원', str(selected_item))
    if ss.file_format == 'naver' and naver_price_match:
        original_price_str = naver_price_match.group(0)
        current_price = int(naver_price_match.group(1))
    elif std_price_match:
        original_price_str = std_price_match.group(0)
        current_price = int(std_price_match.group(1).replace(',', ''))
    else:
        original_price_str, current_price = "", 0
        st.warning("⚠️ 선택하신 품목명에서 기준단가를 찾을 수 없습니다. "
                   "아래 팝업창에서 단가를 직접 입력해 주세요!")

    with st.popover("⚙️ 단가 입력하기 (클릭하여 팝업창 열기)", use_container_width=True):
        st.markdown("#### 단가 설정")
        new_price = st.number_input("단가(원) - 변경 시 자동 반영됩니다",
                                    value=current_price, step=100)
        st.divider()
        st.markdown("#### 🛡️ 계산 안전장치 (미리보기)")
        base_price = ss.global_base_price
        sample_opt = int((5.0 * new_price - base_price) / 10) * 10
        st.info(f"**적용될 계산 공식:** (중량 × 단가 **{new_price}**원) - 기준가 "
                f"**{base_price:,}**원\n\n"
                f"👉 **예시:** 중량이 5.0kg일 경우, 옵션가는 **{sample_opt}**원으로 책정됩니다.")

    st.markdown("---")
    st.subheader(f"2. {col_weight_name} 관리")

    item_rows_for_list = df[df[col_item_name] == selected_item].copy()
    if '재고수량' in item_rows_for_list.columns:
        item_rows_for_list['재고수량'] = pd.to_numeric(
            item_rows_for_list['재고수량'], errors='coerce').fillna(0)
        existing_stock = item_rows_for_list[item_rows_for_list['재고수량'] > 0]
    else:
        existing_stock = item_rows_for_list

    existing_weights_list = existing_stock[col_weight_name].astype(str).tolist()

    col_w1, col_w2 = st.columns(2)
    with col_w1:
        st.markdown(f"**기존 {col_weight_name} 리스트 (재고 0 제외)**")
        st.text_area("참고용입니다 (이곳에서 수정 불가)",
                     value="\n".join(existing_weights_list), height=200, disabled=True)
    with col_w2:
        st.markdown(f"**새로운 {col_weight_name} 리스트 추가**")
        weight_input = st.text_area("추가할 중량만 줄바꿈(Enter)으로 입력하세요.",
                                    height=200, key=f"weight_input_{ss.reset_counter}")

    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        btn_only_price = st.button("👉 새 중량 추가 없이 [단가/기준가만 일괄 변경]",
                                   use_container_width=True)
    with col_btn2:
        btn_add_weights = st.button("👉 새 중량 추가하고 [단가/기준가 일괄 변경]",
                                    type="primary", use_container_width=True)

    if btn_only_price or btn_add_weights:
        base_price = ss.global_base_price
        if base_price == 0:
            st.error("🚨 기준가를 입력해주세요!")
            st.stop()

        ss.history.append(df.copy())

        if original_price_str:
            if ss.file_format == 'naver':
                new_item_name = str(selected_item).replace(original_price_str, f"kg{new_price}")
            else:
                new_item_name = str(selected_item).replace(original_price_str, f"{new_price}원")
        else:
            new_item_name = str(selected_item)

        item_rows = df[df[col_item_name] == selected_item].copy()

        sample_b = item_rows[col_weight_name].iloc[0] if len(item_rows) > 0 else "0kg"
        num_match = re.search(r'(\d+\.?\d*)', str(sample_b))
        if num_match:
            prefix = str(sample_b)[:num_match.start()]
            suffix = str(sample_b)[num_match.end():]
        else:
            prefix, suffix = "", "kg"

        sample_e = (item_rows['관리코드'].iloc[0]
                    if len(item_rows) > 0 and '관리코드' in item_rows.columns else "0kg")
        num_match_e = re.search(r'(\d+\.?\d*)', str(sample_e))
        if num_match_e:
            prefix_e = str(sample_e)[:num_match_e.start()]
            suffix_e = str(sample_e)[num_match_e.end():]
        else:
            prefix_e, suffix_e = "", "kg"

        base_sort_1 = (item_rows['__sort_1'].min() if not item_rows.empty
                       else df['__sort_1'].max() + 1)

        if '재고수량' in item_rows.columns:
            item_rows['재고수량'] = pd.to_numeric(item_rows['재고수량'],
                                              errors='coerce').fillna(0)
            item_rows = item_rows[item_rows['재고수량'] > 0]
        else:
            item_rows['재고수량'] = 1.0

        def extract_num(text):
            m = re.search(r'(\d+\.?\d*)', str(text))
            return float(m.group(1)) if m else 0.0

        if not item_rows.empty:
            item_rows['numeric_weight'] = item_rows[col_weight_name].apply(extract_num)
            item_rows['옵션가'] = (item_rows['numeric_weight'] * new_price
                                - base_price).apply(lambda x: int(x / 10) * 10)
            item_rows[col_item_name] = new_item_name
            item_rows['__sort_1'] = base_sort_1
            item_rows['__sort_2'] = item_rows['numeric_weight']

        new_rows_data = []
        if btn_add_weights:
            for w_str in weight_input.strip().split('\n'):
                w_str = w_str.strip()
                if not w_str:
                    continue
                w_num_match = re.search(r'(\d+\.?\d*)', w_str)
                if w_num_match:
                    w_num = float(w_num_match.group(1))
                    opt_price = int((w_num * new_price - base_price) / 10) * 10
                    new_rows_data.append({
                        col_item_name: new_item_name,
                        col_weight_name: f"{prefix}{w_num}{suffix}",
                        "옵션가": opt_price,
                        "재고수량": 1.0,
                        "관리코드": f"{prefix_e}{w_num}{suffix_e}",
                        "사용여부": "Y",
                        "numeric_weight": w_num,
                        "__sort_1": base_sort_1,
                        "__sort_2": w_num,
                    })

        new_item_df = pd.DataFrame(new_rows_data)
        combined_df = (pd.concat([item_rows, new_item_df], ignore_index=True)
                       if not new_item_df.empty else item_rows)
        if not combined_df.empty:
            combined_df = combined_df.drop(columns=['numeric_weight'], errors='ignore')

        df_remaining = df[df[col_item_name] != selected_item]
        final_concat = pd.concat([df_remaining, combined_df], ignore_index=True)
        final_concat['재고수량'] = pd.to_numeric(final_concat['재고수량'],
                                             errors='coerce').fillna(0)

        group_cols = [col_item_name, col_weight_name, '옵션가']
        agg_dict = {'재고수량': 'sum'}
        for c in final_concat.columns:
            if c not in group_cols and c != '재고수량':
                agg_dict[c] = 'first'

        final_concat = final_concat.groupby(group_cols, as_index=False).agg(agg_dict)
        final_concat = final_concat.sort_values(
            by=['__sort_1', '__sort_2']).reset_index(drop=True)

        ss.processed_data = final_concat
        if btn_only_price:
            st.success(f"✅ '{new_item_name}' 기존 중량들의 단가/기준가가 안전하게 변경되었습니다!")
        else:
            st.success(f"✅ '{new_item_name}' 중량 추가 및 단가 일괄 적용이 완료되었습니다!")
        ss.reset_counter += 1
        st.rerun()

    st.markdown("---")
    st.subheader("3. 최종 결과물 확인 및 다운로드")

    display_df = ss.processed_data.drop(columns=['__sort_1', '__sort_2'], errors='ignore')
    if ss.file_format == 'naver':
        export_df = internal_to_naver(display_df, col_item_name, col_weight_name)
    else:
        export_df = display_df

    st.dataframe(export_df, use_container_width=True)

    xls_buffer = io.BytesIO()
    try:
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Sheet1')
        for col_idx, cname in enumerate(export_df.columns.tolist()):
            ws.write(0, col_idx, str(cname))
        for row_idx, row in enumerate(export_df.values):
            for col_idx, val in enumerate(row):
                if pd.isna(val):
                    val = ""
                elif not isinstance(val, (int, float)):
                    val = str(val)
                ws.write(row_idx + 1, col_idx, val)
        wb.save(xls_buffer)

        if not export_df.empty:
            if ss.file_format == 'naver':
                prefix_name = "supplementProduct"
            else:
                prefix_name = re.sub(r'[\\/*?:"<>|]', "",
                                     str(export_df[col_item_name].iloc[0]))
            final_filename = f"{prefix_name}_{datetime.now():%Y%m%d_%H%M%S}.xls"
        else:
            final_filename = "최종수정본_옵션조합.xls"

        st.download_button(f"💾 모든 변경사항 다운로드 ({final_filename})",
                           xls_buffer.getvalue(), final_filename,
                           "application/octet-stream")
    except Exception as e:
        st.error(f"엑셀 저장 중 오류가 발생했습니다: {e}")


# ═════════════════════════════════════════════════════════════
# 도구 3 — 배송 달력 배너 생성기 (HTML 내장)
# ═════════════════════════════════════════════════════════════
BANNER_HTML = r"""__CALENDAR_HTML__"""


def run_banner():
    st.title("🗓️ 배송 달력 배너 생성기")
    st.caption("날짜를 클릭해 휴무·재개·마감일을 표시하고 PNG/JPG로 저장합니다.")

    with st.expander("💡 사용법 / 저장이 안 될 때", expanded=False):
        st.markdown(
            "1. 왼쪽 패널에서 **연·월·테마·문구**를 설정합니다.  \n"
            "2. **표시 모드**(배송휴무 / 배송재개 / 전지역마감 / 오네마감 / 지우기)를 "
            "고른 뒤 달력의 날짜를 클릭합니다.  \n"
            "3. **PNG 저장 / JPG 저장** 버튼을 누르면 이미지가 만들어집니다.  \n\n"
            "🔸 브라우저 보안 정책 때문에 앱 안(iframe)에서는 자동 다운로드가 "
            "막힐 수 있습니다. 그럴 때는 버튼 아래에 나타나는 **결과 이미지를 "
            "우클릭 → '이미지를 다른 이름으로 저장'**(모바일은 길게 누르기) 하시면 됩니다.  \n"
            "🔸 아래 버튼으로 HTML 파일을 내려받아 브라우저에서 직접 열면 "
            "자동 다운로드가 항상 정상 동작합니다."
        )
        st.download_button(
            "📥 배너 생성기 HTML 따로 받기",
            BANNER_HTML.encode("utf-8"),
            "배송달력_배너생성기.html",
            "text/html",
        )

    components.html(BANNER_HTML, height=1500, scrolling=True)


# ═════════════════════════════════════════════════════════════
if st.session_state.tool == "ledger":
    run_ledger()
elif st.session_state.tool == "option":
    run_option()
else:
    run_banner()
