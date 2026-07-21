# -*- coding: utf-8 -*-
"""
신화 업무 도구 통합 앱
 - 거래처원장 비교
 - 상품 중량 및 옵션가 자동 생성기
"""
import io
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
import streamlit as st
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
}


def go_home():
    st.session_state.tool = None


if st.session_state.tool is None:
    st.title("신화 업무 도구")
    st.caption("사용할 도구를 선택하세요")
    st.markdown("---")

    c1, c2 = st.columns(2)
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
                "A행": x["엑셀행"] if x else "", "A상품명": x["상품명"] if x else "── 없음 ──",
                "A중량": x["중량"] if x else "", "A단가": x["단가"] if x else "",
                "A금액": x["금액"] if x else "", "A수금": x["수금"] if x else "",
                "B행": y["엑셀행"] if y else "", "B상품명": y["상품명"] if y else "── 없음 ──",
                "B중량": y["중량"] if y else "", "B단가": y["단가"] if y else "",
                "B금액": y["금액"] if y else "", "B수금": y["수금"] if y else "",
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
            st.dataframe(detail.style.apply(highlight_row, axis=1),
                         use_container_width=True, height=520)
            st.caption("🔴 한쪽 파일에 없는 행  ·  🟡 값이 다른 행  ·  🔵 위치만 다른 행")

    with t2:
        st.dataframe(daily, use_container_width=True) if not daily.empty \
            else st.success("✅ 날짜별 차이 없음")

    with t3:
        st.dataframe(summary, use_container_width=True)

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
if st.session_state.tool == "ledger":
    run_ledger()
else:
    run_option()
