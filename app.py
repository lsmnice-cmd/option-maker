# -*- coding: utf-8 -*-
"""
신화 업무 도구 통합 앱  v1.6
 - 거래처원장 비교 (서식 자동 인식 · 일자별 대조)
 - 상품 중량 및 옵션가 자동 생성기
 - 배송 달력 배너 생성기

[v1.6 변경사항]
 · 달력 표시 모드 정리 — 전지역 배송 마감, 오네지역 배송 마감, 배송 휴무(방문수령·퀵착불), 배송 재개(수량제한) 제거.
 · 「13시 이전 주문건까지 출고」 모드 추가.
 · 배송 휴무일은 동그라미 대신 날짜 위에 큰 빨간 X 표시로 강조.

[v1.5 변경사항]
 · 달력 표시 모드 2종 추가 — 「배송 휴무 (방문수령 퀵착불 가능)」, 「배송 재개 (수량제한)」.
   날짜 칸에 두 줄로 표시되고 범례에도 자동 반영됩니다.

[v1.4 변경사항]
 · 비어 있던 배송 달력 배너 생성기 HTML을 다시 넣었습니다 (자리표시자 __CALENDAR_HTML__ → 실제 HTML).
 · 배너 저장 시 결과 이미지를 화면에도 표시해 iframe에서 다운로드가 막혀도 우클릭 저장이 가능합니다.

[v1.3 변경사항]
 · 단가 입력창을 팝업(클릭해서 열기)에서 **화면에 항상 보이는 입력창**으로 바꿨습니다.
 · 옆에 1.0 / 3.0 / 5.0kg 옵션가 미리보기를 함께 표시합니다.

[v1.2 변경사항]
 · 옵션명(품목명)을 입력창에서 직접 고쳐 적용할 수 있습니다.
   - 옵션명 안의 단가 표기를 고치면 그 단가가 우선 적용됩니다.
   - 단가 입력창만 바꾼 경우에는 옵션명의 단가 표기가 자동으로 새 단가로 바뀝니다.
 · [옵션명만 바꾸기] 버튼 추가 — 옵션가·중량은 그대로 두고 이름만 교체.
 · 전 품목 옵션명 일괄 찾아 바꾸기 기능 추가.

[v1.1 변경사항]
 · 기준가를 바꾸면 파일 안의 **모든 품목·모든 중량**의 옵션가를 한 번에 다시 계산합니다.
   (품목명에 적힌 kg단가를 읽어 "중량 × 단가 − 기준가" 로 재산출)
 · 자동 재계산 on/off 체크박스와 수동 [지금 다시 계산] 버튼 제공.
 · 재계산 결과 요약(변경 건수)과 단가를 못 읽은 품목 목록 표시.
 · 다운로드 직전에 "현재 기준가와 어긋난 행"이 있는지 검산해 경고 표시.
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

APP_VERSION = "v1.6"

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
    st.title(f"신화 업무 도구 {APP_VERSION}")
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
            "**기준가를 바꾸면 전 품목 옵션가가 한 번에 다시 계산됩니다.**"
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
    st.caption(APP_VERSION)


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


# ── 품목 짝짓기 (상품명 무시 · 숫자로 매칭 · 분할 기재 합산) ──
PAIR_ORDER = {"A파일 없음": 0, "B파일 없음": 1, "중량·금액 다름": 2,
              "금액 다름": 3, "중량 다름": 4, "분할 일치": 8, "일치": 9}

MATCHED = ("일치", "분할 일치")


def _agg(rows):
    """여러 행을 한 덩어리로 합쳐 표시용 dict를 만든다."""
    names = [str(r["상품명"]) for r in rows]
    if len(names) > 3:
        label = " + ".join(names[:3]) + f" 외 {len(names) - 3}건"
    else:
        label = " + ".join(names)
    prices = {round(float(r["단가"])) for r in rows}
    return {
        "행": ", ".join(str(int(r["엑셀행"])) for r in rows),
        "상품명": label,
        "중량": round(sum(float(r["중량"]) for r in rows), 3),
        "단가": prices.pop() if len(prices) == 1 else np.nan,
        "금액": round(sum(float(r["금액"]) for r in rows)),
        "건수": len(rows),
    }


def _verdict(dw, damt, tol_w, tol_amt, n_a, n_b):
    bad = []
    if abs(dw) > tol_w:
        bad.append("중량")
    if abs(damt) > tol_amt:
        bad.append("금액")
    if not bad:
        return "분할 일치" if (n_a > 1 or n_b > 1) else "일치"
    return "·".join(bad) + " 다름"


def match_day(la, lb, tol_w, tol_amt):
    """같은 날짜의 A행·B행을 숫자만 보고 짝짓는다. 1:N(분할 기재)도 묶어서 처리."""
    usedA, usedB, pairs = set(), set(), []

    # 1) 1:1 완전 일치 (중량·금액 모두 허용오차 안)
    for i, r in enumerate(la):
        best, bs = None, None
        for j, s in enumerate(lb):
            if j in usedB:
                continue
            dw = abs(r["중량"] - s["중량"])
            da = abs(r["금액"] - s["금액"])
            if dw <= tol_w and da <= tol_amt:
                v = dw + da / 1e6
                if bs is None or v < bs:
                    best, bs = j, v
        if best is not None:
            usedA.add(i)
            usedB.add(best)
            pairs.append((_agg([r]), _agg([lb[best]]), "일치"))

    # 2) 남은 행을 '같은 단가'끼리 묶어 합계로 대조 (한 건을 여러 줄로 나눠 적은 경우)
    def group(rows, used):
        g = defaultdict(list)
        for i, r in enumerate(rows):
            if i in used or pd.isna(r["단가"]):
                continue
            g[round(float(r["단가"]))].append(i)
        return g

    ga, gb = group(la, usedA), group(lb, usedB)
    for p in sorted(set(ga) & set(gb)):
        ra = [la[i] for i in ga[p]]
        rb = [lb[j] for j in gb[p]]
        A, B = _agg(ra), _agg(rb)
        label = _verdict(A["중량"] - B["중량"], A["금액"] - B["금액"],
                         tol_w, tol_amt, len(ra), len(rb))
        usedA.update(ga[p])
        usedB.update(gb[p])
        pairs.append((A, B, label))

    # 3) 단가가 서로 다른 경우 — 중량 또는 금액 한쪽만 같은 행끼리 1:1
    def one_to_one(cond, score, label):
        for i, r in enumerate(la):
            if i in usedA:
                continue
            best, bs = None, None
            for j, s in enumerate(lb):
                if j in usedB or not cond(r, s):
                    continue
                v = score(r, s)
                if bs is None or v < bs:
                    best, bs = j, v
            if best is not None:
                usedA.add(i)
                usedB.add(best)
                pairs.append((_agg([r]), _agg([lb[best]]), label))

    one_to_one(lambda r, s: abs(r["금액"] - s["금액"]) <= tol_amt,
               lambda r, s: abs(r["중량"] - s["중량"]), "중량 다름")
    one_to_one(lambda r, s: abs(r["중량"] - s["중량"]) <= tol_w,
               lambda r, s: abs(r["금액"] - s["금액"]), "금액 다름")

    # 4) 끝내 짝이 없는 행
    for i, r in enumerate(la):
        if i not in usedA:
            pairs.append((_agg([r]), None, "B파일 없음"))
    for j, s in enumerate(lb):
        if j not in usedB:
            pairs.append((None, _agg([s]), "A파일 없음"))
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
                "A행": x["행"] if x else "",
                "A상품명": x["상품명"] if x else "── 없음 ──",
                "A중량": x["중량"] if x else None,
                "A단가": x["단가"] if x else None,
                "A금액": x["금액"] if x else None,
                "B행": y["행"] if y else "",
                "B상품명": y["상품명"] if y else "── 없음 ──",
                "B중량": y["중량"] if y else None,
                "B단가": y["단가"] if y else None,
                "B금액": y["금액"] if y else None,
                "중량차": round((x["중량"] if x else 0) - (y["중량"] if y else 0), 2),
                "금액차": round((x["금액"] if x else 0) - (y["금액"] if y else 0)),
                "묶음": (f"A {x['건수']}건 ↔ B {y['건수']}건"
                       if x and y and (x["건수"] > 1 or y["건수"] > 1) else ""),
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
    "A중량": "{:,.2f}", "A단가": "{:,.0f}", "A금액": "{:,.0f}",
    "B중량": "{:,.2f}", "B단가": "{:,.0f}", "B금액": "{:,.0f}",
    "중량차": "{:,.2f}", "금액차": "{:,.0f}",
}

PAIR_COLOR = {
    "A파일 없음": "#ffd6d6", "B파일 없음": "#ffd6d6",
    "중량·금액 다름": "#ffe2c2", "금액 다름": "#fff3c4",
    "중량 다름": "#e8f0fe", "분할 일치": "#e6f4ea", "일치": "#f2f7f3",
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
        show_split = st.checkbox("나눠 적어서 맞은 건도 같이 보기", value=False)

    pairs = build_pairs(a, b, tol_w, tol_amt, overlap_only)
    if pairs.empty:
        diff = split_ok = pairs
    else:
        diff = pairs[~pairs["구분"].isin(MATCHED)]
        split_ok = pairs[pairs["구분"] == "분할 일치"]
    daily = compare_daily(a, b, tol_w, tol_amt, False)
    bad_days = daily[daily["판정"] != "일치"]

    st.markdown("---")
    m = st.columns(4)
    m[0].metric("대조한 건수", f"{len(pairs):,}건", f"차이 {len(diff):,}건")
    m[1].metric("나눠 적어 맞은 건", f"{len(split_ok):,}건",
                f"차이 나는 날짜 {len(set(diff['날짜'])) if not diff.empty else 0}일")
    m[2].metric("중량 차이 합", f"{diff['중량차'].sum() if not diff.empty else 0:,.2f} kg")
    m[3].metric("금액 차이 합", f"{diff['금액차'].sum() if not diff.empty else 0:,.0f} 원")

    t1, t2, t3, t4 = st.tabs(["🎯 차이 품목 대조", "📅 일자별 대조",
                              "🔍 날짜별 전체 보기", "📊 요약"])

    with t1:
        base = pairs[pairs["구분"] != "일치"] if show_split else diff
        if pairs.empty:
            st.info("대조할 데이터가 없습니다. (겹치는 날짜가 없을 수 있습니다)")
        elif base.empty:
            st.success("✅ 모든 건이 허용오차 안에서 맞습니다.")
        else:
            f1, f2 = st.columns([1, 3])
            day_opt = ["전체"] + sorted(set(base["날짜"]))
            pick_day = f1.selectbox("날짜", day_opt)
            kinds = sorted(set(base["구분"]), key=lambda k: PAIR_ORDER.get(k, 9))
            pick_kind = f2.multiselect("보고 싶은 유형", kinds, default=kinds)

            view = base.copy()
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
                "한쪽이 한 건을 여러 줄로 나눠 적은 경우, **같은 단가끼리 묶어서** "
                "합계로 대조합니다. `묶음` 열에 몇 건 대 몇 건인지 표시되고 "
                "`A행 / B행`에 해당 엑셀 행번호가 모두 들어갑니다.  \n"
                "🟥 **A/B파일 없음** 한쪽에만 있는 건 · "
                "🟧 **중량·금액 다름** 단가는 같은데 물량이 안 맞음 · "
                "🟨 **금액 다름** 중량은 같은데 금액이 다름 · "
                "🟦 **중량 다름** 금액은 같은데 중량이 다름 · "
                "🟩 **분할 일치** 나눠 적었지만 합계는 맞음  \n"
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


# ── 공통 계산 유틸 (v1.1) ────────────────────────────────────
def extract_num(text):
    """문자열에서 첫 번째 숫자를 실수로 뽑아낸다. (예: '팩 2.5kg' → 2.5)"""
    m = re.search(r'(\d+\.?\d*)', str(text))
    return float(m.group(1)) if m else 0.0


def extract_unit_price(item_name, file_format):
    """
    품목명에 적힌 kg당 단가를 읽어낸다.
      · 네이버 서식 : '...kg13500' 형태
      · 표준 서식   : '...13,500원' 형태
    → (단가, 품목명에서 매칭된 문자열).  못 찾으면 (None, "")
    """
    name = str(item_name)
    naver_m = re.search(r'kg\s*(\d{3,})', name)
    std_m = re.search(r'(\d{1,3}(?:,\d{3})*|\d+)원', name)

    if file_format == 'naver' and naver_m:
        return int(naver_m.group(1)), naver_m.group(0)
    if std_m:
        return int(std_m.group(1).replace(',', '')), std_m.group(0)
    if naver_m:
        return int(naver_m.group(1)), naver_m.group(0)
    return None, ""


def calc_option_price(weight, unit_price, base_price):
    """옵션가 = (중량 × 단가 − 기준가) 를 10원 단위로 절삭"""
    return int((weight * unit_price - base_price) / 10) * 10


def resolve_name_and_price(edited_name, origin_price, popover_price, file_format):
    """
    사용자가 고친 옵션명과 단가 입력창 값을 종합해 '최종 옵션명 / 적용 단가'를 정한다. (v1.2)

      1) 옵션명 안의 단가 표기를 직접 고친 경우 → 그 값을 단가로 채택 (옵션명 우선)
      2) 그 외에는 단가 입력창 값을 채택하고, 옵션명의 단가 표기를 새 값으로 자동 치환
    → (최종 옵션명, 적용 단가, 어디서 온 단가인지 설명)
    """
    edited_name = str(edited_name).strip()
    name_price, name_str = extract_unit_price(edited_name, file_format)

    if name_price is not None and origin_price is not None and name_price != origin_price:
        return edited_name, name_price, "옵션명에 적으신 단가"

    if name_price is not None and name_str:
        new_str = (f"kg{popover_price}" if name_str.lower().startswith("kg")
                   else f"{popover_price}원")
        return edited_name.replace(name_str, new_str), popover_price, "단가 입력창"

    return edited_name, popover_price, "단가 입력창 (옵션명에는 단가 표기가 없음)"


def recalc_all_options(df, col_item, col_weight, base_price, file_format):
    """
    파일 안의 모든 행을 현재 기준가로 다시 계산한다.
    품목명에서 단가를 못 읽은 행은 손대지 않고 그대로 둔다.
    → (재계산된 df, 값이 바뀐 행 수, 단가를 못 읽은 품목명 목록)
    """
    out = df.copy()
    if '옵션가' not in out.columns:
        out['옵션가'] = 0

    old = pd.to_numeric(out['옵션가'], errors='coerce')
    new_vals, skipped = [], []

    for idx, row in out.iterrows():
        unit, _ = extract_unit_price(row[col_item], file_format)
        if unit is None:
            new_vals.append(old.loc[idx] if pd.notna(old.loc[idx]) else 0)
            nm = str(row[col_item])
            if nm not in skipped:
                skipped.append(nm)
            continue
        new_vals.append(calc_option_price(extract_num(row[col_weight]),
                                          unit, base_price))

    out['옵션가'] = new_vals
    changed = int((old.fillna(-10 ** 9).astype(float)
                   != pd.Series(new_vals, index=out.index).astype(float)).sum())
    return out, changed, skipped


def count_mismatch(df, col_item, col_weight, base_price, file_format):
    """현재 기준가 기준으로 계산했을 때 옵션가가 어긋나는 행 수를 센다."""
    if df is None or df.empty:
        return 0
    bad = 0
    old = pd.to_numeric(df.get('옵션가'), errors='coerce')
    for idx, row in df.iterrows():
        unit, _ = extract_unit_price(row[col_item], file_format)
        if unit is None:
            continue
        want = calc_option_price(extract_num(row[col_weight]), unit, base_price)
        cur = old.loc[idx]
        if pd.isna(cur) or int(cur) != want:
            bad += 1
    return bad


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
    ss.setdefault('applied_base_price', None)      # v1.1 — 실제로 적용된 기준가
    ss.setdefault('auto_recalc', True)             # v1.1 — 자동 재계산 여부

    st.title("⚖️ 상품 중량 및 옵션가 자동 생성기")
    st.caption("다중 품목 지원 · 네이버 추가상품 서식 자동 인식 · "
               "기준가 변경 시 전 품목 일괄 재계산")

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
                ss.applied_base_price = None
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

    col_item_name = ss.col_item_name
    col_weight_name = ss.col_weight_name

    # ── v1.1 : 기준가 변경 시 전 품목 옵션가 일괄 재계산 ─────────
    st.markdown("##### 🔁 기준가 일괄 반영")
    r1, r2 = st.columns([3, 1])
    with r1:
        ss.auto_recalc = st.checkbox(
            "기준가를 바꾸면 **파일 안 모든 품목·모든 중량**의 옵션가를 자동으로 다시 계산합니다 "
            "(품목명에 적힌 kg단가 기준)",
            value=ss.auto_recalc, key="auto_recalc_cb")
    with r2:
        manual_recalc = st.button("🔁 지금 전체 다시 계산", use_container_width=True)

    auto_trigger = (ss.auto_recalc
                    and ss.applied_base_price != ss.global_base_price)

    if manual_recalc or auto_trigger:
        ss.history.append(ss.processed_data.copy())
        new_df, changed, skipped = recalc_all_options(
            ss.processed_data, col_item_name, col_weight_name,
            ss.global_base_price, ss.file_format)
        ss.processed_data = new_df
        ss.applied_base_price = ss.global_base_price

        st.success(
            f"✅ 기준가 **{ss.global_base_price:,}원** 기준으로 전체 "
            f"**{len(new_df):,}행**을 다시 계산했습니다. "
            f"(값이 바뀐 행 **{changed:,}건**)  \n"
            "계산식: `옵션가 = (중량 × 품목명의 kg단가 − 기준가)` · 10원 단위 절삭"
        )
        if skipped:
            head = " / ".join(skipped[:5])
            more = f" 외 {len(skipped) - 5}개" if len(skipped) > 5 else ""
            st.warning(
                f"⚠️ 품목명에서 kg단가를 읽지 못해 **건드리지 않은 품목**이 있습니다: "
                f"{head}{more}  \n"
                "해당 품목은 아래에서 직접 선택해 단가를 입력한 뒤 적용해 주세요."
            )

    df = ss.processed_data

    st.markdown("---")
    col_title, col_undo = st.columns([3, 1])
    with col_title:
        fmt_label = "네이버 추가상품" if ss.file_format == 'naver' else "기존 표준"
        st.subheader(f"1. 수정할 품목 선택 및 단가 설정  ({fmt_label} 서식)")
    with col_undo:
        if st.button("⏪ 방금 한 작업 되돌리기 (Undo)", disabled=not ss.history):
            ss.processed_data = ss.history.pop()
            ss.applied_base_price = None      # 되돌린 뒤 재계산 상태 초기화
            st.success("이전 상태로 되돌렸습니다!")
            st.rerun()

    unique_items = df[col_item_name].dropna().unique()
    selected_item = st.selectbox(f"A열({col_item_name})에서 수정할 항목을 선택하세요", unique_items)

    if ss.get('last_selected_item') != selected_item:
        ss.reset_counter += 1
        ss.last_selected_item = selected_item

    current_price, original_price_str = extract_unit_price(selected_item, ss.file_format)
    if current_price is None:
        current_price, original_price_str = 0, ""
        st.warning("⚠️ 선택하신 품목명에서 기준단가를 찾을 수 없습니다. "
                   "아래 팝업창에서 단가를 직접 입력해 주세요!")

    st.markdown("##### ⚙️ 단가 설정")
    p1, p2 = st.columns([1, 3])
    with p1:
        new_price = st.number_input("단가(원) — 바꾸면 바로 반영됩니다", min_value=0,
                                    value=int(current_price), step=100,
                                    key=f"unit_price_{ss.reset_counter}")
    with p2:
        base_price = ss.global_base_price
        st.markdown("<br>", unsafe_allow_html=True)
        st.info(
            f"**계산 공식:** (중량 × 단가 **{new_price:,}**원) − 기준가 "
            f"**{base_price:,}**원  \n"
            f"👉 **예시:** 1.0kg → **{calc_option_price(1.0, new_price, base_price):,}원** · "
            f"3.0kg → **{calc_option_price(3.0, new_price, base_price):,}원** · "
            f"5.0kg → **{calc_option_price(5.0, new_price, base_price):,}원**"
        )

    # ── v1.2 : 옵션명 직접 수정 ──────────────────────────────
    st.markdown("##### ✏️ 옵션명 수정")
    n1, n2 = st.columns([3, 1])
    with n1:
        edited_name = st.text_input(
            "옵션명(품목명)을 직접 고쳐 쓰실 수 있습니다. 아래 적용 버튼을 누르면 반영됩니다.",
            value=str(selected_item),
            key=f"item_name_edit_{ss.reset_counter}")
    with n2:
        st.markdown("<br>", unsafe_allow_html=True)
        btn_rename_only = st.button("✏️ 옵션명만 바꾸기", use_container_width=True)

    final_item_name, apply_price, price_src = resolve_name_and_price(
        edited_name, current_price, new_price, ss.file_format)

    if final_item_name != str(selected_item) or apply_price != current_price:
        st.info(f"적용될 옵션명 → **{final_item_name}**  \n"
                f"적용될 단가 → **{apply_price:,}원** ({price_src} 기준)")

    with st.expander("🔤 전 품목 옵션명 일괄 찾아 바꾸기"):
        f1, f2, f3 = st.columns([2, 2, 1])
        find_txt = f1.text_input("찾을 문구", key="rename_find")
        repl_txt = f2.text_input("바꿀 문구", key="rename_repl")
        f3.markdown("<br>", unsafe_allow_html=True)
        do_replace = f3.button("전체 바꾸기", use_container_width=True)
        hit = (df[col_item_name].astype(str)
               .str.contains(find_txt, regex=False).sum()) if find_txt else 0
        st.caption(f"현재 '{find_txt}' 가 들어간 행: **{hit:,}건**" if find_txt
                   else "찾을 문구를 입력하면 몇 건이 바뀌는지 미리 알려 드립니다. "
                        "단가 표기를 바꾸신 경우에는 위의 [🔁 지금 전체 다시 계산]을 "
                        "한 번 눌러 옵션가를 맞춰 주세요.")
        if do_replace and find_txt:
            ss.history.append(df.copy())
            new_df = df.copy()
            new_df[col_item_name] = (new_df[col_item_name].astype(str)
                                     .str.replace(find_txt, repl_txt, regex=False))
            ss.processed_data = new_df
            st.success(f"✅ {hit:,}건의 옵션명을 바꿨습니다.")
            st.rerun()

    if btn_rename_only:
        if not final_item_name:
            st.error("🚨 옵션명이 비어 있습니다.")
        elif final_item_name == str(selected_item):
            st.warning("옵션명이 그대로입니다. 바꿀 내용이 없습니다.")
        else:
            ss.history.append(df.copy())
            new_df = df.copy()
            mask = new_df[col_item_name] == selected_item
            new_df.loc[mask, col_item_name] = final_item_name
            ss.processed_data = new_df
            ss.last_selected_item = final_item_name
            st.success(f"✅ 옵션명을 '{final_item_name}' 로 바꿨습니다. "
                       f"({int(mask.sum()):,}행) · 옵션가는 그대로 두었습니다.")
            st.rerun()

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

        # v1.2 — 편집한 옵션명과 적용 단가를 사용
        new_item_name = final_item_name if final_item_name else str(selected_item)
        new_price = apply_price

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

        if not item_rows.empty:
            item_rows['numeric_weight'] = item_rows[col_weight_name].apply(extract_num)
            item_rows['옵션가'] = item_rows['numeric_weight'].apply(
                lambda w: calc_option_price(w, new_price, base_price))
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
                    opt_price = calc_option_price(w_num, new_price, base_price)
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

    # v1.1 — 현재 기준가와 어긋난 행이 남아 있는지 검산
    mismatch = count_mismatch(ss.processed_data, col_item_name, col_weight_name,
                              ss.global_base_price, ss.file_format)
    if mismatch:
        st.warning(f"⚠️ 현재 기준가({ss.global_base_price:,}원) 계산값과 다른 행이 "
                   f"**{mismatch:,}건** 있습니다. 위의 [🔁 지금 전체 다시 계산]을 "
                   "누르면 전부 맞춰집니다.")
    else:
        st.success(f"✅ 모든 행이 현재 기준가({ss.global_base_price:,}원) 기준으로 "
                   "정확히 계산되어 있습니다.")

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
BANNER_HTML = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>배송 달력 배너 생성기</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/pretendard/1.3.9/static/pretendard.min.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<style>
:root{
  --ink:#1c1c1e; --paper:#f5f4f0; --panel:#ffffff; --line:#e4e2db;
  --accent:#0e3b2e;
  --off:#c0392b; --off-soft:#fdecea;          /* 배송 휴무 */
  --resume:#1f5fbf; --resume-soft:#e8f0fc;    /* 배송 재개 */
  --cut13:#b45309; --cut13-soft:#fdf0e0;      /* 13시 이전 주문건까지 출고 */
}
*{box-sizing:border-box; margin:0; padding:0;}
body{
  font-family:'Pretendard Variable',Pretendard,-apple-system,'Noto Sans KR',sans-serif;
  background:var(--paper); color:var(--ink); min-height:100vh;
}
.app{max-width:1180px; margin:0 auto; padding:28px 20px 60px;}
.app-head{margin-bottom:22px;}
.app-head h1{font-size:22px; font-weight:800; letter-spacing:-0.02em;}
.app-head p{font-size:13px; color:#777; margin-top:4px;}
.layout{display:grid; grid-template-columns:340px 1fr; gap:26px; align-items:start;}
@media (max-width:920px){ .layout{grid-template-columns:1fr;} }

/* ---------- 컨트롤 패널 ---------- */
.panel{background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:20px;}
.panel + .panel{margin-top:16px;}
.panel h2{font-size:13px; font-weight:700; color:#999; letter-spacing:0.08em; margin-bottom:12px;}
.row{display:flex; gap:8px; margin-bottom:10px;}
.row:last-child{margin-bottom:0;}
label.field{display:block; font-size:12px; font-weight:600; color:#555; margin-bottom:5px;}
input[type=text], select{
  width:100%; padding:9px 11px; border:1px solid var(--line); border-radius:8px;
  font-size:14px; font-family:inherit; background:#fff;
}
input[type=text]:focus, select:focus{outline:2px solid var(--accent); outline-offset:-1px;}
.field-group{margin-bottom:14px;}
.field-group:last-child{margin-bottom:0;}

.mode-btns{display:grid; grid-template-columns:1fr 1fr; gap:6px;}
.mode-btn{
  padding:11px 6px; border-radius:9px; border:1.5px solid var(--line); background:#fff;
  font-size:13px; font-weight:700; cursor:pointer; font-family:inherit; transition:all .12s;
  line-height:1.25;
}
.mode-btn.off.active{border-color:var(--off); background:var(--off-soft); color:var(--off);}
.mode-btn.resume.active{border-color:var(--resume); background:var(--resume-soft); color:var(--resume);}
.mode-btn.cut13.active{border-color:var(--cut13); background:var(--cut13-soft); color:var(--cut13);}
.mode-btn.erase.active{border-color:#555; background:#f0f0f0; color:#333;}
.mode-btn small{display:block; font-size:10.5px; font-weight:600; opacity:.8; margin-top:2px;}
.hint{font-size:12px; color:#999; margin-top:10px; line-height:1.5;}

.week-list{display:flex; flex-direction:column; gap:6px;}
.week-item{
  display:flex; align-items:center; gap:10px; padding:9px 12px;
  border:1.5px solid var(--line); border-radius:9px; cursor:pointer;
  font-size:13px; font-weight:600; color:#666; background:#fff; transition:all .12s;
  user-select:none;
}
.week-item.on{border-color:var(--accent); background:#eef4f0; color:var(--accent);}
.week-item input{accent-color:var(--accent); width:15px; height:15px; cursor:pointer;}

.theme-btns{display:grid; grid-template-columns:repeat(4,1fr); gap:8px;}
.theme-btn{
  height:44px; border-radius:9px; cursor:pointer; border:2px solid transparent;
  display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:700; color:#fff;
  font-family:inherit;
}
.theme-btn.active{border-color:var(--ink); box-shadow:0 0 0 2px #fff inset;}

.export-btns{display:grid; grid-template-columns:1fr 1fr; gap:8px;}
.export-btn{
  padding:13px; border-radius:10px; border:none; cursor:pointer;
  font-size:14px; font-weight:800; font-family:inherit; transition:opacity .12s;
}
.export-btn:hover{opacity:.88;}
.export-btn.png{background:var(--ink); color:#fff;}
.export-btn.jpg{background:#fff; color:var(--ink); border:1.5px solid var(--ink);}
.size-note{font-size:12px; color:#999; margin-top:8px; text-align:center;}
.result-wrap{margin-top:12px; display:none;}
.result-wrap img{width:100%; border:1px solid var(--line); border-radius:8px;}
.result-wrap p{font-size:12px; color:#777; margin-top:6px; line-height:1.5;}

/* ---------- 배너 미리보기 ---------- */
.preview-wrap{overflow-x:auto;}
.banner{
  width:720px; margin:0 auto; background:var(--b-paper,#faf9f5);
  border:1px solid var(--line);
  --b-accent:#0e3b2e; --b-accent-text:#faf9f5; --b-paper:#faf9f5; --b-ink:#22221f;
}

.b-head{
  background:var(--b-accent); color:var(--b-accent-text);
  padding:38px 44px 32px; position:relative;
}
.b-head .en{
  font-size:12px; letter-spacing:0.34em; font-weight:600; opacity:.75; text-transform:uppercase;
}
.b-head .month-line{display:flex; align-items:flex-end; gap:16px; margin-top:8px;}
.b-head .month-num{font-size:76px; font-weight:800; line-height:0.9; letter-spacing:-0.03em;}
.b-head .month-unit{font-size:26px; font-weight:700; padding-bottom:6px;}
.b-head .title{
  font-size:21px; font-weight:600; margin-top:16px; letter-spacing:-0.01em;
  white-space:pre-wrap;
}
.b-head .stamp{
  position:absolute; right:44px; top:40px; width:86px; height:86px; border-radius:50%;
  border:1.5px solid currentColor; opacity:.85;
  display:flex; flex-direction:column; align-items:center; justify-content:center; gap:2px;
  font-size:12px; font-weight:700; letter-spacing:0.12em; text-align:center; line-height:1.3;
}
.stamp .s-small{font-size:10px; letter-spacing:0.2em; opacity:.8;}

.b-legend{
  display:flex; flex-wrap:wrap; gap:12px 20px; padding:18px 44px; border-bottom:1px solid var(--line);
  font-size:13px; font-weight:600; color:var(--b-ink); background:var(--b-paper);
}
.b-legend .item{display:flex; align-items:center; gap:7px;}
.dot{width:11px; height:11px; border-radius:50%; flex-shrink:0;}
.dot.off{background:transparent; width:13px; height:13px; position:relative; border-radius:0;}
.dot.off::before,.dot.off::after{content:''; position:absolute; left:50%; top:50%; width:15px; height:3px; background:var(--off); border-radius:2px;}
.dot.off::before{transform:translate(-50%,-50%) rotate(45deg);}
.dot.off::after{transform:translate(-50%,-50%) rotate(-45deg);}
.dot.resume{background:transparent; border:2.5px solid var(--resume); width:8px; height:8px;}
.dot.cut13{background:var(--cut13);}

.b-cal{padding:26px 44px 8px; background:var(--b-paper);}
.cal-grid{width:100%; border-collapse:collapse; table-layout:fixed;}
.cal-grid th{
  font-size:13px; font-weight:700; color:#8a8a82; padding-bottom:12px;
  letter-spacing:0.06em;
}
.cal-grid th.sun{color:var(--off);}
.cal-grid td{
  height:94px; vertical-align:top; text-align:center; cursor:pointer; position:relative;
  border-top:1px solid var(--line);
}
.cal-grid td:hover .num{outline:2px dashed #bbb; outline-offset:2px;}
.cal-grid td.empty{cursor:default;}
.cal-grid td.empty:hover .num{outline:none;}
.num{
  display:inline-flex; align-items:center; justify-content:center;
  width:38px; height:38px; border-radius:50%; margin-top:10px;
  font-size:16px; font-weight:600; color:var(--b-ink);
}
td.sun .num{color:var(--off);}
td.st-off .num{color:#fff; font-weight:800; position:relative; z-index:0; text-shadow:0 0 3px var(--off),0 0 3px var(--off),0 0 3px var(--off);}
td.st-off .num::before,td.st-off .num::after{
  content:''; position:absolute; left:50%; top:50%; width:48px; height:4.5px;
  background:var(--off); border-radius:3px; opacity:1; z-index:-1;
}
td.st-off .num::before{transform:translate(-50%,-50%) rotate(45deg);}
td.st-off .num::after{transform:translate(-50%,-50%) rotate(-45deg);}
td.st-resume .num{border:2.5px solid var(--resume); color:var(--resume); font-weight:700;}
td.st-cut13 .num{background:var(--cut13); color:#fff; font-weight:700;}
.tag{
  display:block; font-size:10.5px; font-weight:800; margin-top:4px; letter-spacing:-0.01em;
  line-height:1.2;
}
.tag.off{color:var(--off);}
.tag.resume{color:var(--resume);}
.tag.cut13{color:var(--cut13);}
.tag .sub{display:block; font-size:8.5px; font-weight:700; letter-spacing:-0.03em; margin-top:1px; opacity:.9;}

.b-foot{
  padding:20px 44px 30px; background:var(--b-paper);
  font-size:13.5px; line-height:1.65; color:#55554e; white-space:pre-wrap;
}
.b-foot .bar{width:26px; height:3px; background:var(--b-accent); margin-bottom:12px;}
</style>
</head>
<body>
<div class="app">
  <div class="app-head">
    <h1>배송 달력 배너 생성기</h1>
    <p>날짜를 클릭해 배송 휴무(X)·출고 마감·재개일을 표시하고, 보여줄 주만 골라 PNG/JPG로 내보내세요.</p>
  </div>

  <div class="layout">
    <div>
      <div class="panel">
        <h2>기본 설정</h2>
        <div class="field-group">
          <div class="row">
            <div style="flex:1">
              <label class="field">연도</label>
              <select id="selYear"></select>
            </div>
            <div style="flex:1">
              <label class="field">월</label>
              <select id="selMonth"></select>
            </div>
          </div>
        </div>
        <div class="field-group">
          <label class="field">제목</label>
          <input type="text" id="inpTitle" value="">
        </div>
        <div class="field-group">
          <label class="field">하단 안내 문구</label>
          <input type="text" id="inpNotice" value="배송 휴무 기간 주문 건은 배송 재개일부터 순차 발송됩니다. 신선식품 특성상 발송 일정을 꼭 확인해 주세요.">
        </div>
      </div>

      <div class="panel">
        <h2>날짜 표시 모드</h2>
        <div class="mode-btns">
          <button class="mode-btn off active" data-mode="off">배송 휴무</button>
          <button class="mode-btn resume" data-mode="resume">배송 재개</button>
          <button class="mode-btn cut13" data-mode="cut13">13시 이전 주문건까지 출고</button>
          <button class="mode-btn erase" data-mode="erase">지우기</button>
        </div>
        <p class="hint">모드를 선택한 뒤 오른쪽 달력에서 날짜를 클릭하면 표시됩니다. 같은 날짜를 다시 클릭하면 해제됩니다.</p>
      </div>

      <div class="panel">
        <h2>표시할 주 선택</h2>
        <div class="week-list" id="weekList"></div>
        <p class="hint">체크한 주만 배너에 표시됩니다. 필요한 주만 남기면 배너가 더 간결해집니다.</p>
      </div>

      <div class="panel">
        <h2>색상 테마</h2>
        <div class="theme-btns" id="themeBtns"></div>
      </div>

      <div class="panel">
        <h2>내보내기</h2>
        <div class="export-btns">
          <button class="export-btn png" id="btnPng">PNG 저장</button>
          <button class="export-btn jpg" id="btnJpg">JPG 저장</button>
        </div>
        <p class="size-note">2160px 폭 고해상도로 저장됩니다.</p>
        <div class="result-wrap" id="resultWrap">
          <img id="resultImg" alt="생성된 배너">
          <p>자동 다운로드가 막힌 환경이라면 위 이미지를 <b>우클릭 → 이미지를 다른 이름으로 저장</b>(모바일은 길게 누르기) 하세요.</p>
        </div>
      </div>
    </div>

    <div class="preview-wrap">
      <div class="banner" id="banner">
        <div class="b-head">
          <div class="en" id="bEn"></div>
          <div class="month-line">
            <span class="month-num" id="bMonthNum"></span>
            <span class="month-unit">월</span>
          </div>
          <div class="title" id="bTitle"></div>
          <div class="stamp"><span class="s-small">NOTICE</span><span>배송<br>안내</span></div>
        </div>
        <div class="b-legend" id="bLegend"></div>
        <div class="b-cal">
          <table class="cal-grid">
            <thead>
              <tr>
                <th class="sun">일</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th>토</th>
              </tr>
            </thead>
            <tbody id="calBody"></tbody>
          </table>
        </div>
        <div class="b-foot">
          <div class="bar"></div>
          <span id="bNotice"></span>
        </div>
      </div>
    </div>
  </div>
</div>

<script>
const MONTH_EN = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
const THEMES = [
  {name:'딥그린',  accent:'#0e3b2e', paper:'#faf9f5', text:'#faf9f5'},
  {name:'네이비',  accent:'#152847', paper:'#f8f9fb', text:'#f8f9fb'},
  {name:'버건디',  accent:'#5e1f26', paper:'#fbf8f6', text:'#fbf6f2'},
  {name:'차콜',    accent:'#26262a', paper:'#f7f7f5', text:'#f5f5f2'},
];
const MARK_DEF = {
  off:    {label:'배송 휴무',        tag:'배송휴무'},
  resume: {label:'배송 재개일',      tag:'배송재개'},
  cut13:  {label:'13시 이전 주문건까지 출고', tag:'13시 이전 주문', sub:'당일 출고'},
};

const state = {
  year: new Date().getFullYear(),
  month: new Date().getMonth(), // 0-based
  mode: 'off',
  theme: 0,
  marks: {},        // "YYYY-M-D" -> off|resume|cut13
  weeksOn: {},      // "YYYY-M" -> [true,...] 주별 표시 여부
};

/* --- 셀렉트 초기화 --- */
const selYear = document.getElementById('selYear');
const selMonth = document.getElementById('selMonth');
const thisYear = new Date().getFullYear();
for(let y = thisYear - 1; y <= thisYear + 2; y++){
  const o = document.createElement('option');
  o.value = y; o.textContent = y + '년';
  selYear.appendChild(o);
}
for(let m = 0; m < 12; m++){
  const o = document.createElement('option');
  o.value = m; o.textContent = (m+1) + '월';
  selMonth.appendChild(o);
}
selYear.value = state.year;
selMonth.value = state.month;

selYear.addEventListener('change', () => { state.year = +selYear.value; render(); });
selMonth.addEventListener('change', () => {
  state.month = +selMonth.value;
  const t = document.getElementById('inpTitle');
  if(/^\d+월 배송 일정 안내$/.test(t.value)) t.value = (state.month+1) + '월 배송 일정 안내';
  render();
});

document.getElementById('inpTitle').addEventListener('input', render);
document.getElementById('inpNotice').addEventListener('input', render);

/* --- 모드 버튼 --- */
document.querySelectorAll('.mode-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    state.mode = btn.dataset.mode;
    document.querySelectorAll('.mode-btn').forEach(b => b.classList.toggle('active', b === btn));
  });
});

/* --- 테마 버튼 --- */
const themeWrap = document.getElementById('themeBtns');
THEMES.forEach((t, i) => {
  const b = document.createElement('button');
  b.className = 'theme-btn' + (i === 0 ? ' active' : '');
  b.style.background = t.accent;
  b.textContent = t.name;
  b.addEventListener('click', () => {
    state.theme = i;
    themeWrap.querySelectorAll('.theme-btn').forEach((x, j) => x.classList.toggle('active', j === i));
    render();
  });
  themeWrap.appendChild(b);
});

function keyOf(d){ return state.year + '-' + state.month + '-' + d; }
function monthKey(){ return state.year + '-' + state.month; }

/* 이 달의 주 배열 생성: [[날짜 or 0(빈칸) × 7], ...] */
function buildWeeks(){
  const first = new Date(state.year, state.month, 1).getDay();
  const days = new Date(state.year, state.month + 1, 0).getDate();
  const weeks = [];
  let week = new Array(first).fill(0);
  for(let d = 1; d <= days; d++){
    week.push(d);
    if(week.length === 7){ weeks.push(week); week = []; }
  }
  if(week.length){ while(week.length < 7) week.push(0); weeks.push(week); }
  return weeks;
}

function getWeeksOn(weekCount){
  const k = monthKey();
  if(!state.weeksOn[k] || state.weeksOn[k].length !== weekCount){
    state.weeksOn[k] = new Array(weekCount).fill(true);
  }
  return state.weeksOn[k];
}

function onDayClick(d){
  const k = keyOf(d);
  if(state.mode === 'erase'){ delete state.marks[k]; }
  else if(state.marks[k] === state.mode){ delete state.marks[k]; }
  else { state.marks[k] = state.mode; }
  render();
}

function render(){
  const t = THEMES[state.theme];
  const banner = document.getElementById('banner');
  banner.style.setProperty('--b-accent', t.accent);
  banner.style.setProperty('--b-paper', t.paper);
  banner.style.setProperty('--b-accent-text', t.text);

  document.getElementById('bEn').textContent = MONTH_EN[state.month] + ' ' + state.year + ' · DELIVERY CALENDAR';
  document.getElementById('bMonthNum').textContent = state.month + 1;
  document.getElementById('bTitle').textContent = document.getElementById('inpTitle').value || '';
  document.getElementById('bNotice').textContent = document.getElementById('inpNotice').value || '';

  const weeks = buildWeeks();
  const on = getWeeksOn(weeks.length);

  /* --- 주 선택 체크박스 --- */
  const wl = document.getElementById('weekList');
  wl.innerHTML = '';
  weeks.forEach((week, i) => {
    const daysIn = week.filter(d => d > 0);
    const range = (state.month+1) + '/' + daysIn[0] + ' ~ ' + (state.month+1) + '/' + daysIn[daysIn.length-1];
    const item = document.createElement('label');
    item.className = 'week-item' + (on[i] ? ' on' : '');
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.checked = on[i];
    cb.addEventListener('change', () => { on[i] = cb.checked; render(); });
    item.appendChild(cb);
    const span = document.createElement('span');
    span.textContent = (i+1) + '주차  (' + range + ')';
    item.appendChild(span);
    wl.appendChild(item);
  });

  /* --- 범례: 실제 사용된 표시만 노출 --- */
  const used = new Set(Object.entries(state.marks)
    .filter(([k]) => k.startsWith(monthKey() + '-'))
    .map(([,v]) => v));
  const legend = document.getElementById('bLegend');
  legend.innerHTML = '';
  const order = ['off','cut13','resume'];
  const showTypes = order.filter(x => used.has(x));
  const finalTypes = showTypes.length ? showTypes : order; // 표시 전엔 전체 안내
  finalTypes.forEach(type => {
    const div = document.createElement('div');
    div.className = 'item';
    div.innerHTML = '<span class="dot ' + type + '"></span> ' + MARK_DEF[type].label;
    legend.appendChild(div);
  });

  /* --- 달력 (선택된 주만) --- */
  const body = document.getElementById('calBody');
  body.innerHTML = '';
  weeks.forEach((week, i) => {
    if(!on[i]) return;
    const tr = document.createElement('tr');
    week.forEach((d, dow) => {
      const td = document.createElement('td');
      if(d === 0){ td.className = 'empty'; tr.appendChild(td); return; }
      if(dow === 0) td.classList.add('sun');
      const st = state.marks[keyOf(d)];
      if(st) td.classList.add('st-' + st);
      let html = '<span class="num">' + d + '</span>';
      if(st){
        html += '<span class="tag ' + st + '">' + MARK_DEF[st].tag
              + (MARK_DEF[st].sub ? '<span class="sub">' + MARK_DEF[st].sub + '</span>' : '')
              + '</span>';
      }
      td.innerHTML = html;
      td.addEventListener('click', () => onDayClick(d));
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

/* --- 내보내기 --- */
async function exportImage(type){
  const banner = document.getElementById('banner');
  const canvas = await html2canvas(banner, {
    scale: 3,
    useCORS: true,
    backgroundColor: type === 'jpg' ? '#ffffff' : null,
  });
  const mime = type === 'jpg' ? 'image/jpeg' : 'image/png';
  const url = canvas.toDataURL(mime, 0.95);
  const fname = state.year + '년' + (state.month+1) + '월_배송달력.' + type;

  // 결과 이미지를 화면에도 표시 (iframe 등 자동 다운로드가 막힌 환경 대비)
  const img = document.getElementById('resultImg');
  img.src = url;
  document.getElementById('resultWrap').style.display = 'block';

  try{
    const a = document.createElement('a');
    a.href = url; a.download = fname;
    document.body.appendChild(a); a.click(); a.remove();
  }catch(e){ /* 다운로드 차단 시 이미지 우클릭 저장으로 대체 */ }
}
document.getElementById('btnPng').addEventListener('click', () => exportImage('png'));
document.getElementById('btnJpg').addEventListener('click', () => exportImage('jpg'));

document.getElementById('inpTitle').value = (new Date().getMonth()+1) + '월 배송 일정 안내';
render();
</script>
</body>
</html>
"""


def run_banner():
    st.title("🗓️ 배송 달력 배너 생성기")
    st.caption("날짜를 클릭해 휴무·재개·마감일을 표시하고 PNG/JPG로 저장합니다.")

    with st.expander("💡 사용법 / 저장이 안 될 때", expanded=False):
        st.markdown(
            "1. 왼쪽 패널에서 **연·월·테마·문구**를 설정합니다.  \n"
            "2. **표시 모드**(배송휴무 / 배송재개 / 13시 이전 주문건까지 출고 / 지우기)를 "
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
