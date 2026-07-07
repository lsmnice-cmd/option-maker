import streamlit as st
import pandas as pd
import re
import io
import xlwt 
from datetime import datetime

st.set_page_config(layout="wide")
st.title("상품 중량 및 옵션가 자동 생성기 (다중 품목 지원)")

if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
    st.session_state.last_file_id = None
    st.session_state.col_item_name = None
    st.session_state.history = []
    st.session_state.global_base_price = 0
    st.session_state.file_format = None   # 💡 추가: 'standard' 또는 'naver'


# 💡 네이버 추가상품 서식을 내부 표준 서식으로 변환하는 함수
def naver_to_internal(df_naver):
    """
    네이버 추가상품 서식 -> 내부 표준 서식(품목/중량/옵션가/재고수량/관리코드/사용여부)
    추가상품명 -> 품목 (예: 냉장)프라임갈비본살kg47000)
    추가상품값 -> 품목명+중량 텍스트 (예: 냉장)프라임갈비본살 IBP원육3.17kg)
    추가상품가 -> 옵션가
    """
    rows = []
    for _, r in df_naver.iterrows():
        item_full = str(r['추가상품명'])          # A열 역할 (단가 포함 품목명)
        value_full = str(r['추가상품값'])         # 중량 텍스트가 들어있는 원본
        rows.append({
            '품목': item_full,
            '중량': value_full,                   # 중량 원본 텍스트를 그대로 보존
            '옵션가': r.get('추가상품가', 0),
            '재고수량': r.get('재고수량', 0),
            '관리코드': r.get('관리코드', ''),
            '사용여부': r.get('사용여부', 'Y'),
        })
    return pd.DataFrame(rows)


# 💡 내부 표준 서식을 네이버 추가상품 서식으로 되돌리는 함수 (다운로드용)
def internal_to_naver(df_internal, col_item_name):
    out = pd.DataFrame({
        '추가상품명': df_internal[col_item_name],
        '추가상품값': df_internal['중량'],
        '추가상품가': df_internal['옵션가'],
        '재고수량': df_internal['재고수량'],
        '사용여부': df_internal.get('사용여부', 'Y'),
        '관리코드': df_internal.get('관리코드', ''),
    })
    return out


uploaded_file = st.file_uploader("기존 양식 파일(xls, xlsx, csv) 또는 네이버 추가상품 파일을 업로드하세요", type=['xls', 'xlsx', 'csv'])

if uploaded_file:
    current_file_id = getattr(uploaded_file, 'file_id', uploaded_file.name + str(uploaded_file.size))
    
    if st.session_state.last_file_id != current_file_id:
        try:
            for key in ['base_price', 'weight_input', 'global_base_price_input']:
                if key in st.session_state:
                    del st.session_state[key]

            st.session_state.global_base_price = 0
                    
            if uploaded_file.name.endswith('.csv'):
                file_bytes = uploaded_file.read()
                encodings = ['utf-8', 'cp949', 'euc-kr', 'utf-8-sig']
                df = None
                for enc in encodings:
                    try:
                        df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
                        df.columns = df.columns.str.strip()
                        if '품목 및 등급' in df.columns or '품목' in df.columns or '추가상품명' in df.columns: break
                    except: continue
            else:
                if uploaded_file.name.endswith('.xls'):
                    df = pd.read_excel(uploaded_file, engine='xlrd')
                else:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                df.columns = df.columns.str.strip()
                
            if df is None:
                st.error("파일을 제대로 읽지 못했습니다.")
                st.stop()

            # 💡 서식 자동 감지: 네이버 추가상품 vs 기존 표준
            if '추가상품명' in df.columns and '추가상품값' in df.columns:
                st.session_state.file_format = 'naver'
                df = naver_to_internal(df)   # 내부 표준 서식으로 변환
                col_name = '품목'
                st.info("📌 네이버 추가상품 서식으로 인식했습니다. 내부적으로 변환하여 처리하며, 다운로드 시 네이버 서식으로 저장됩니다.")
            else:
                st.session_state.file_format = 'standard'
                if '품목 및 등급' in df.columns:
                    col_name = '품목 및 등급'
                elif '품목' in df.columns:
                    col_name = '품목'
                else:
                    st.error("파일에서 '품목 및 등급', '품목', 또는 '추가상품명' 열을 찾을 수 없습니다.")
                    st.stop()
                
            if '중량' not in df.columns:
                st.error("파일에서 '중량' 열을 찾을 수 없습니다.")
                st.stop()
                
            df['__sort_1'] = range(len(df))
            df['__sort_2'] = 0.0
            
            st.session_state.col_item_name = col_name
            st.session_state.processed_data = df.copy()
            st.session_state.last_file_id = current_file_id
            st.session_state.history = []
            st.success("파일이 성공적으로 로드되었습니다! 아래에서 기준가를 먼저 입력해주세요.")
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            st.stop()

if st.session_state.processed_data is not None:

    st.markdown("---")
    st.subheader("⚡ 기준가 설정 (필수)")

    col_bp1, col_bp2 = st.columns([2, 3])
    with col_bp1:
        entered_base_price = st.number_input(
            "🚨 기준가(원)를 입력하세요",
            min_value=0,
            value=st.session_state.global_base_price,
            step=100,
            key="global_base_price_input"
        )
    with col_bp2:
        if entered_base_price > 0:
            st.success(f"✅ 기준가 **{entered_base_price:,}원** 이 설정되었습니다. 아래에서 품목별 작업을 진행하세요.")
        else:
            st.warning("⚠️ 기준가를 입력해야 이후 작업을 진행할 수 있습니다.")

    st.session_state.global_base_price = entered_base_price

    if st.session_state.global_base_price == 0:
        st.info("👆 기준가를 입력하면 품목 선택 및 중량 관리 기능이 활성화됩니다.")
        st.stop()

if st.session_state.processed_data is not None:
    df = st.session_state.processed_data
    col_item_name = st.session_state.col_item_name
    
    st.markdown("---")
    
    col_title, col_undo = st.columns([3, 1])
    with col_title:
        fmt_label = "네이버 추가상품" if st.session_state.file_format == 'naver' else "기존 표준"
        st.subheader(f"1. 수정할 품목 선택 및 단가 설정  ({fmt_label} 서식)")
    with col_undo:
        can_undo = len(st.session_state.history) > 0
        if st.button("⏪ 방금 한 작업 되돌리기 (Undo)", disabled=not can_undo):
            st.session_state.processed_data = st.session_state.history.pop()
            st.success("이전 상태로 되돌렸습니다!")
            st.rerun()
    
    unique_items = df[col_item_name].dropna().unique()
    selected_item = st.selectbox(f"A열({col_item_name})에서 수정할 항목을 선택하세요", unique_items)
    
    match = re.search(r'(\d{1,3}(?:,\d{3})*|\d+)원?', str(selected_item))
    # 💡 네이버 서식은 'kg47000'처럼 '원' 없이 단가가 붙어있어 별도 패턴도 시도
    naver_price_match = re.search(r'kg\s*(\d{3,})', str(selected_item))
    if naver_price_match:
        original_price_str = naver_price_match.group(0)
        current_price = int(naver_price_match.group(1))
    elif re.search(r'(\d{1,3}(?:,\d{3})*|\d+)원', str(selected_item)):
        m2 = re.search(r'(\d{1,3}(?:,\d{3})*|\d+)원', str(selected_item))
        original_price_str = m2.group(0)
        current_price = int(m2.group(1).replace(',', ''))
    else:
        original_price_str = ""
        current_price = 0
        st.warning("⚠️ 선택하신 품목명에서 기준단가를 찾을 수 없습니다. 아래 팝업창에서 단가를 직접 입력해 주세요!")
    
    with st.popover("⚙️ 단가 입력하기 (클릭하여 팝업창 열기)", use_container_width=True):
        st.markdown("#### 단가 설정")
        new_price = st.number_input("단가(원) - 변경 시 자동 반영됩니다", value=current_price, step=100)
        
        st.divider()
        st.markdown("#### 🛡️ 계산 안전장치 (미리보기)")
        base_price = st.session_state.global_base_price
        sample_opt = int((5.0 * new_price - base_price) / 10) * 10
        st.info(f"**적용될 계산 공식:** (중량 × 단가 **{new_price}**원) - 기준가 **{base_price:,}**원\n\n"
                f"👉 **예시:** 중량이 5.0kg일 경우, 옵션가는 **{sample_opt}**원으로 책정됩니다.")

    st.markdown("---")
    st.subheader("2. 중량 관리")
    
    item_rows_for_list = df[df[col_item_name] == selected_item].copy()
    if '재고수량' in item_rows_for_list.columns:
        item_rows_for_list['재고수량'] = pd.to_numeric(item_rows_for_list['재고수량'], errors='coerce').fillna(0)
        existing_stock = item_rows_for_list[item_rows_for_list['재고수량'] > 0]
    else:
        existing_stock = item_rows_for_list
    
    existing_weights_list = existing_stock['중량'].astype(str).tolist()
    
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        st.markdown("**기존 중량 리스트 (재고 0 제외)**")
        st.text_area("참고용입니다 (이곳에서 수정 불가)", value="\n".join(existing_weights_list), height=200, disabled=True)
        
    with col_w2:
        st.markdown("**새로운 중량 리스트 추가**")
        weight_input = st.text_area("추가할 중량만 줄바꿈(Enter)으로 입력하세요.", height=200, key="weight_input")

    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        btn_only_price = st.button("👉 새 중량 추가 없이 [단가/기준가만 일괄 변경]", use_container_width=True)
    with col_btn2:
        btn_add_weights = st.button("👉 새 중량 추가하고 [단가/기준가 일괄 변경]", type="primary", use_container_width=True)
    
    if btn_only_price or btn_add_weights:
        base_price = st.session_state.global_base_price
        if base_price == 0:
            st.error("🚨 기준가를 입력해주세요!")
            st.stop()
            
        st.session_state.history.append(df.copy())

        # 💡 품목명 안의 단가 치환 (표준: OOO원 / 네이버: kgOOOOO)
        if original_price_str:
            if st.session_state.file_format == 'naver':
                new_item_name = str(selected_item).replace(original_price_str, f"kg{new_price}")
            else:
                new_item_name = str(selected_item).replace(original_price_str, f"{new_price}원")
        else:
            new_item_name = str(selected_item)
            
        item_rows = df[df[col_item_name] == selected_item].copy()
        
        sample_b = item_rows['중량'].iloc[0] if len(item_rows) > 0 else "0kg"
        # 💡 중량 텍스트 안의 숫자 위치를 찾아 prefix/suffix 분리 (kg 앞 숫자 기준)
        num_match = re.search(r'(\d+\.?\d*)\s*kg', str(sample_b))
        if num_match:
            prefix = str(sample_b)[:num_match.start(1)]
            suffix = str(sample_b)[num_match.end(1):]
        else:
            num_match = re.search(r'(\d+\.?\d*)', str(sample_b))
            if num_match:
                prefix = str(sample_b)[:num_match.start()]
                suffix = str(sample_b)[num_match.end():]
            else:
                prefix, suffix = "", "kg"
            
        sample_e = item_rows['관리코드'].iloc[0] if len(item_rows) > 0 and '관리코드' in item_rows.columns else "0kg"
        num_match_e = re.search(r'(\d+\.?\d*)\s*kg', str(sample_e))
        if num_match_e:
            prefix_e = str(sample_e)[:num_match_e.start(1)]
            suffix_e = str(sample_e)[num_match_e.end(1):]
        else:
            num_match_e = re.search(r'(\d+\.?\d*)', str(sample_e))
            if num_match_e:
                prefix_e = str(sample_e)[:num_match_e.start()]
                suffix_e = str(sample_e)[num_match_e.end():]
            else:
                prefix_e, suffix_e = "", "kg"
            
        base_sort_1 = item_rows['__sort_1'].min() if not item_rows.empty else df['__sort_1'].max() + 1
            
        if '재고수량' in item_rows.columns:
            item_rows['재고수량'] = pd.to_numeric(item_rows['재고수량'], errors='coerce').fillna(0)
            item_rows = item_rows[item_rows['재고수량'] > 0]
        else:
            item_rows['재고수량'] = 1.0
        
        def extract_num(text):
            m = re.search(r'(\d+\.?\d*)\s*kg', str(text))
            if m:
                return float(m.group(1))
            m = re.search(r'(\d+\.?\d*)', str(text))
            return float(m.group(1)) if m else 0.0
            
        if not item_rows.empty:
            item_rows['numeric_weight'] = item_rows['중량'].apply(extract_num)
            item_rows['옵션가'] = (item_rows['numeric_weight'] * new_price - base_price).apply(lambda x: int(x / 10) * 10)
            item_rows[col_item_name] = new_item_name
            # 💡 중량 텍스트 안의 품목명 앞부분도 새 품목명 규칙에 맞게 재조립하지 않고 원본 유지
            item_rows['__sort_1'] = base_sort_1
            item_rows['__sort_2'] = item_rows['numeric_weight']
            
        new_rows_data = []
        
        if btn_add_weights:
            weights = weight_input.strip().split('\n')
            for w_str in weights:
                w_str = w_str.strip()
                if not w_str: continue
                
                w_num_match = re.search(r'(\d+\.?\d*)', w_str)
                if w_num_match:
                    w_num = float(w_num_match.group(1))
                    opt_price = int((w_num * new_price - base_price) / 10) * 10
                    
                    formatted_weight = f"{prefix}{w_num}{suffix}"
                    formatted_code = f"{prefix_e}{w_num}{suffix_e}"
                    
                    new_rows_data.append({
                        col_item_name: new_item_name,
                        "중량": formatted_weight,
                        "옵션가": opt_price,
                        "재고수량": 1.0,
                        "관리코드": formatted_code,
                        "사용여부": "Y",
                        "numeric_weight": w_num,
                        "__sort_1": base_sort_1,
                        "__sort_2": w_num
                    })
                
        new_item_df = pd.DataFrame(new_rows_data)
        if not new_item_df.empty:
            combined_df = pd.concat([item_rows, new_item_df], ignore_index=True)
        else:
            combined_df = item_rows
            
        if not combined_df.empty:
            combined_df = combined_df.drop(columns=['numeric_weight'], errors='ignore')
            
        df_remaining = df[df[col_item_name] != selected_item]
        final_concat = pd.concat([df_remaining, combined_df], ignore_index=True)
        
        final_concat['재고수량'] = pd.to_numeric(final_concat['재고수량'], errors='coerce').fillna(0)
        group_cols = [col_item_name, '중량', '옵션가']
        
        agg_dict = {'재고수량': 'sum'}
        for c in final_concat.columns:
            if c not in group_cols and c != '재고수량':
                agg_dict[c] = 'first'
                
        final_concat = final_concat.groupby(group_cols, as_index=False).agg(agg_dict)
        final_concat = final_concat.sort_values(by=['__sort_1', '__sort_2']).reset_index(drop=True)
        
        st.session_state.processed_data = final_concat
        
        if btn_only_price:
            st.success(f"✅ '{new_item_name}' 기존 중량들의 단가/기준가가 안전하게 변경되었습니다!")
        else:
            st.success(f"✅ '{new_item_name}' 중량 추가 및 단가 일괄 적용이 완료되었습니다!")
            
        st.rerun()

    st.markdown("---")
    st.subheader("3. 최종 결과물 확인 및 다운로드")
    
    display_df = st.session_state.processed_data.drop(columns=['__sort_1', '__sort_2'], errors='ignore')

    # 💡 화면 표시 및 다운로드용 서식 결정
    if st.session_state.file_format == 'naver':
        export_df = internal_to_naver(display_df, col_item_name)
    else:
        export_df = display_df

    st.dataframe(export_df)
    
    xls_buffer = io.BytesIO()
    try:
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Sheet1')
        
        columns = export_df.columns.tolist()
        for col_idx, cname in enumerate(columns):
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
            if st.session_state.file_format == 'naver':
                first_item_name = str(export_df['추가상품명'].iloc[0])
            else:
                first_item_name = str(export_df[col_item_name].iloc[0])
            safe_item_name = re.sub(r'[\\/*?:"<>|]', "", first_item_name)
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            prefix_name = "supplementProduct" if st.session_state.file_format == 'naver' else safe_item_name
            final_filename = f"{prefix_name}_{current_time}.xls"
        else:
            final_filename = "최종수정본_옵션조합.xls"
        
        st.download_button(
            label=f"💾 모든 변경사항 다운로드 ({final_filename})",
            data=xls_buffer.getvalue(),
            file_name=final_filename,
            mime="application/octet-stream"
        )
    except Exception as e:
        st.error(f"엑셀 저장 중 오류가 발생했습니다: {e}")
