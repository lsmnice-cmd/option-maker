# -*- coding: utf-8 -*-
"""
신화 업무 도구 통합 앱
 - 거래처원장 비교
 - 상품 중량 및 옵션가 자동 생성기
 - 배송 달력 배너 생성기
"""
import io
import re
from collections import defaultdict
from datetime import datetime

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
            "신화미트·신화푸드 거래처원장 두 파일을 대조해 틀린 행을 찾아냅니다.  \n"
            "상품명이 달라도 **중량·단가·금액·수금** 기준으로 비교합니다."
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
# 도구 1 — 거래처원장 비교
# ═════════════════════════════════════════════════════════════
LEDGER_COLS = ["월일", "상품명", "원산지", "Box", "Kg",
               "매입단가", "매입공급가", "매입부가세", "매입합계",
               "매출단가", "매출공급가", "매출부가세", "매출합계",
               "지급액", "수금액", "미수금액", "X1", "X2", "X3"]


def ledger_load(file):
    df = pd.read_excel(file, sheet_name=0, header=None, skiprows=5)
    df = df.iloc[:, :len(LEDGER_COLS)]
    df.columns = LEDGER_COLS[:df.shape[1]]
    df = df[df["월일"].notna()].copy()
    df["엑셀행"] = df.index + 6
    df["월일"] = df["월일"].astype(str).str.strip()
    df = df[df["월일"].str.match(r"\d{4}/\d{2}/\d{2}")].reset_index(drop=True)
    for c in ["Kg", "매입단가", "매입합계", "매출단가", "매출합계",
              "미수금액", "수금액", "지급액"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        else:
            df[c] = 0
    df["중량"] = df["Kg"].fillna(0)
    df["단가"] = df["매입단가"].fillna(df["매출단가"]).fillna(0)
    df["금액"] = df["매입합계"].fillna(df["매출합계"]).fillna(0)
    df["수금"] = df["수금액"].fillna(df["지급액"]).fillna(0)
    df["미수"] = df["미수금액"].abs()
    df["상품명"] = df["상품명"].fillna("").astype(str).str.strip()
    return df


def ledger_key(r):
    return (r["월일"], round(float(r["중량"]), 3), round(float(r["단가"]), 2),
            round(float(r["금액"]), 2), round(float(r["수금"]), 2))


def diff_text(x, y):
    if x is None:
        return "A파일 누락"
    if y is None:
        return "B파일 누락"
    d = [l for k, l in [("중량", "중량"), ("단가", "단가"), ("금액", "금액"), ("수금", "수금")]
         if abs(float(x[k]) - float(y[k])) > 0.001]
    return ", ".join(d) + " 다름" if d else "행 위치 차이"


def ledger_compare(a, b):
    bucket = defaultdict(list)
    for i, r in b.iterrows():
        bucket[ledger_key(r)].append(i)

    matched, ua = set(), []
    for i, r in a.iterrows():
        k = ledger_key(r)
        if bucket[k]:
            matched.add(bucket[k].pop(0))
        else:
            ua.append(i)
    ub = [i for i in b.index if i not in matched]

    ra, rb = a.loc[ua], b.loc[ub]

    rows = []
    for d in sorted(set(ra["월일"]) | set(rb["월일"])):
        la = ra[ra["월일"] == d].to_dict("records")
        lb = rb[rb["월일"] == d].to_dict("records")
        for i in range(max(len(la), len(lb))):
            x = la[i] if i < len(la) else None
            y = lb[i] if i < len(lb) else None
            rows.append({
                "월일": d,
                "A행": x["엑셀행"] if x else None, "A상품명": x["상품명"] if x else "── 없음 ──",
                "A중량": x["중량"] if x else None, "A단가": x["단가"] if x else None,
                "A금액": x["금액"] if x else None, "A수금": x["수금"] if x else None,
                "B행": y["엑셀행"] if y else None, "B상품명": y["상품명"] if y else "── 없음 ──",
                "B중량": y["중량"] if y else None, "B단가": y["단가"] if y else None,
                "B금액": y["금액"] if y else None, "B수금": y["수금"] if y else None,
                "차이": diff_text(x, y),
            })
    detail = pd.DataFrame(rows)

    def agg(df):
        return df.groupby("월일").agg(중량=("중량", "sum"), 금액=("금액", "sum"),
                                     건수=("월일", "size"))

    daily = agg(a).join(agg(b), how="outer", lsuffix="_A", rsuffix="_B").fillna(0)
    daily["중량차"] = (daily["중량_A"] - daily["중량_B"]).round(3)
    daily["금액차"] = daily["금액_A"] - daily["금액_B"]
    daily["건수차"] = daily["건수_A"] - daily["건수_B"]
    daily = daily[(daily["중량차"].abs() > 0.001) | (daily["금액차"].abs() > 0.01)
                  | (daily["건수차"] != 0)].reset_index()

    def last(df):
        s = df["미수"].dropna()
        return s.iloc[-1] if len(s) else 0

    summary = pd.DataFrame([
        {"항목": "행 수", "A파일": len(a), "B파일": len(b), "차이": len(a) - len(b)},
        {"항목": "중량합계", "A파일": round(a["중량"].sum(), 2),
         "B파일": round(b["중량"].sum(), 2),
         "차이": round(a["중량"].sum() - b["중량"].sum(), 2)},
        {"항목": "금액합계", "A파일": a["금액"].sum(), "B파일": b["금액"].sum(),
         "차이": a["금액"].sum() - b["금액"].sum()},
        {"항목": "최종미수금액", "A파일": last(a), "B파일": last(b),
         "차이": last(a) - last(b)},
        {"항목": "불일치 행수", "A파일": len(ra), "B파일": len(rb), "차이": ""},
    ])
    return summary, detail, daily


def highlight_row(row):
    d = str(row.get("차이", ""))
    if "누락" in d:
        c = "background-color:#ffd6d6"
    elif "다름" in d:
        c = "background-color:#fff3c4"
    elif "위치" in d:
        c = "background-color:#e8f0fe"
    else:
        c = ""
    return [c] * len(row)


LEDGER_FMT = {
    "A행": "{:.0f}", "B행": "{:.0f}",
    "A중량": "{:,.2f}", "A단가": "{:,.0f}", "A금액": "{:,.0f}", "A수금": "{:,.0f}",
    "B중량": "{:,.2f}", "B단가": "{:,.0f}", "B금액": "{:,.0f}", "B수금": "{:,.0f}",
}

DAILY_FMT = {
    "중량_A": "{:,.2f}", "중량_B": "{:,.2f}", "중량차": "{:,.2f}",
    "금액_A": "{:,.0f}", "금액_B": "{:,.0f}", "금액차": "{:,.0f}",
    "건수_A": "{:.0f}", "건수_B": "{:.0f}", "건수차": "{:.0f}",
}


def fmt_summary(v):
    if isinstance(v, (int, float)):
        return f"{v:,.2f}".rstrip("0").rstrip(".") if v % 1 else f"{v:,.0f}"
    return v


def run_ledger():
    st.title("📑 거래처원장 비교")
    st.caption("상품명(상품코드)은 달라도 무방하며, 중량·단가·금액·수금이 일치해야 합니다.")

    c1, c2 = st.columns(2)
    with c1:
        fa = st.file_uploader("A 파일 (예: 신화미트)", type=["xlsx", "xls"], key="lg_a")
    with c2:
        fb = st.file_uploader("B 파일 (예: 신화푸드)", type=["xlsx", "xls"], key="lg_b")

    if not (fa and fb):
        st.info("👆 비교할 원장 파일 2개를 업로드하세요.")
        return

    try:
        a, b = ledger_load(fa), ledger_load(fb)
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        return

    if a.empty or b.empty:
        st.error("데이터 행을 찾지 못했습니다. 거래처원장 양식인지 확인하세요.")
        return

    summary, detail, daily = ledger_compare(a, b)

    st.markdown("---")
    m = st.columns(4)
    m[0].metric("A 행수", f"{len(a):,}")
    m[1].metric("B 행수", f"{len(b):,}")
    m[2].metric("틀린 행", f"{len(detail):,}")
    gap = summary.loc[summary["항목"] == "최종미수금액", "차이"].iloc[0]
    m[3].metric("미수금액 차이", f"{gap:,.0f}원")

    t1, t2, t3 = st.tabs(["틀린행 대조", "날짜별 차이", "요약"])

    with t1:
        if detail.empty:
            st.success("✅ 모든 행이 일치합니다.")
        else:
            st.dataframe(
                detail.style.apply(highlight_row, axis=1).format(LEDGER_FMT, na_rep=""),
                use_container_width=True, height=520, hide_index=True)
            st.caption("🔴 한쪽 파일에 없는 행  ·  🟡 값이 다른 행  ·  🔵 위치만 다른 행")

    with t2:
        if daily.empty:
            st.success("✅ 날짜별 차이 없음")
        else:
            st.dataframe(daily.style.format(DAILY_FMT, na_rep=""),
                         use_container_width=True, hide_index=True)

    with t3:
        st.dataframe(summary.applymap(fmt_summary),
                     use_container_width=True, hide_index=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        detail.to_excel(w, sheet_name="틀린행 대조", index=False)
        daily.to_excel(w, sheet_name="날짜별 차이", index=False)
        summary.to_excel(w, sheet_name="요약", index=False)
    st.download_button("💾 비교결과 다운로드 (xlsx)", buf.getvalue(),
                       f"비교결과_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
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
BANNER_HTML = r"""
<!DOCTYPE html>
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
  --allcut:#b45309; --allcut-soft:#fdf0e0;    /* 전지역 배송 마감 */
  --onecut:#6d28d9; --onecut-soft:#f1eafd;    /* 오네지역 배송 마감 */
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
.mode-btn.allcut.active{border-color:var(--allcut); background:var(--allcut-soft); color:var(--allcut);}
.mode-btn.onecut.active{border-color:var(--onecut); background:var(--onecut-soft); color:var(--onecut);}
.mode-btn.erase.active{border-color:#555; background:#f0f0f0; color:#333;}
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
.export-result{display:none; margin-top:18px; background:var(--panel);
  border:1px solid var(--line); border-radius:14px; padding:16px;}
.export-result .ex-head{font-size:13px; font-weight:700; display:flex;
  align-items:center; gap:10px; flex-wrap:wrap;}
.export-result .ex-link{font-size:12px; font-weight:700; color:#fff;
  background:var(--ink); padding:5px 10px; border-radius:7px; text-decoration:none;}
.export-result .ex-hint{font-size:12px; color:#777; margin:8px 0 12px; line-height:1.6;}
.export-result img{width:100%; border:1px solid var(--line); border-radius:10px; display:block;}
.size-note{font-size:12px; color:#999; margin-top:8px; text-align:center;}

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
.dot.off{background:var(--off);}
.dot.resume{background:transparent; border:2.5px solid var(--resume); width:8px; height:8px;}
.dot.allcut{background:var(--allcut);}
.dot.onecut{background:transparent; border:2.5px solid var(--onecut); width:8px; height:8px;}

.b-cal{padding:26px 44px 8px; background:var(--b-paper);}
.cal-grid{width:100%; border-collapse:collapse; table-layout:fixed;}
.cal-grid th{
  font-size:13px; font-weight:700; color:#8a8a82; padding-bottom:12px;
  letter-spacing:0.06em;
}
.cal-grid th.sun{color:var(--off);}
.cal-grid td{
  height:84px; vertical-align:top; text-align:center; cursor:pointer; position:relative;
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
td.st-off .num{background:var(--off); color:#fff; font-weight:700;}
td.st-resume .num{border:2.5px solid var(--resume); color:var(--resume); font-weight:700;}
td.st-allcut .num{background:var(--allcut); color:#fff; font-weight:700;}
td.st-onecut .num{border:2.5px solid var(--onecut); color:var(--onecut); font-weight:700;}
.tag{
  display:block; font-size:10.5px; font-weight:800; margin-top:4px; letter-spacing:-0.01em;
  line-height:1.2;
}
.tag.off{color:var(--off);}
.tag.resume{color:var(--resume);}
.tag.allcut{color:var(--allcut);}
.tag.onecut{color:var(--onecut);}

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
    <p>날짜를 클릭해 배송 휴무·마감·재개일을 표시하고, 보여줄 주만 골라 PNG/JPG로 내보내세요.</p>
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
          <button class="mode-btn allcut" data-mode="allcut">전지역 배송 마감</button>
          <button class="mode-btn onecut" data-mode="onecut">오네지역 배송 마감</button>
          <button class="mode-btn erase" data-mode="erase" style="grid-column:1/-1">지우기</button>
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
      <div class="export-result" id="exportResult"></div>
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
  allcut: {label:'전지역 배송 마감', tag:'전지역마감'},
  onecut: {label:'오네지역 배송 마감', tag:'오네마감'},
};

const state = {
  year: new Date().getFullYear(),
  month: new Date().getMonth(), // 0-based
  mode: 'off',
  theme: 0,
  marks: {},        // "YYYY-M-D" -> off|resume|allcut|onecut
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
  const order = ['off','allcut','onecut','resume'];
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
      if(st) html += '<span class="tag ' + st + '">' + MARK_DEF[st].tag + '</span>';
      td.innerHTML = html;
      td.addEventListener('click', () => onDayClick(d));
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

/* --- 내보내기 (앱 내장 iframe 대응) --- */
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

  // (1) 일반 다운로드 시도 — 브라우저가 허용하면 바로 저장됨
  try{
    const a = document.createElement('a');
    a.href = url; a.download = fname;
    document.body.appendChild(a); a.click(); a.remove();
  }catch(e){ /* 샌드박스 iframe에서는 차단될 수 있음 */ }

  // (2) 차단 대비 — 결과 이미지를 화면에 표시해 직접 저장할 수 있게 함
  const box = document.getElementById('exportResult');
  box.innerHTML =
    '<div class="ex-head"><span>생성 완료 · ' + fname + '</span>' +
    '<a class="ex-link" href="' + url + '" download="' + fname + '" target="_blank">저장 / 새 창에서 열기</a></div>' +
    '<p class="ex-hint">자동 저장이 되지 않았다면 아래 이미지를 <b>우클릭 → 이미지를 다른 이름으로 저장</b>' +
    '(모바일은 길게 누르기) 하세요.</p>' +
    '<img src="' + url + '" alt="' + fname + '">';
  box.style.display = 'block';
  box.scrollIntoView({behavior:'smooth', block:'nearest'});
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
