import streamlit as st
import pandas as pd
import re
import io
import xlwt 

st.set_page_config(layout="wide")
st.title("상품 중량 및 옵션가 자동 생성기 (다중 품목 지원)")

# 상태 유지를 위한 초기화
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
    st.session_state.file_name = None
    st.session_state.col_item_name = None
    st.session_state.history = [] 

# 1. 파일 업로드
uploaded_file = st.file_uploader("기존 양식 파일(xls, xlsx, csv)을 업로드하세요", type=['xls', 'xlsx', 'csv'])

if uploaded_file:
    if st.session_state.file_name != uploaded_file.name:
        try:
            if uploaded_file.name.endswith('.csv'):
                file_bytes = uploaded_file.read()
                encodings = ['utf-8', 'cp949', 'euc-kr', 'utf-8-sig']
                df = None
                for enc in encodings:
                    try:
                        df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
                        if '품목 및 등급' in df.columns or '품목' in df.columns: break
                    except: continue
            else:
                df = pd.read_excel(uploaded_file)
                
            if df is None:
                st.error("파일을 제대로 읽지 못했습니다.")
                st.stop()
                
            if '품목 및 등급' in df.columns:
                col_name = '품목 및 등급'
            elif '품목' in df.columns:
                col_name = '품목'
            else:
                st.error("파일에서 '품목 및 등급' 또는 '품목' 열(A열)을 찾을 수 없습니다.")
                st.stop()
                
            df['__sort_1'] = range(len(df))
            df['__sort_2'] = 0.0
            
            st.session_state.col_item_name = col_name
            st.session_state.processed_data = df.copy()
            st.session_state.file_name = uploaded_file.name
            st.session_state.history = [] 
            st.success("파일이 성공적으로 로드되었습니다! 아래에서 품목을 선택하고 작업을 진행하세요.")
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            st.stop()

# 2. 메인 작업 영역
if st.session_state.processed_data is not None:
    df = st.session_state.processed_data
    col_item_name = st.session_state.col_item_name
    
    st.markdown("---")
    
    col_title, col_undo = st.columns([3, 1])
    with col_title:
        st.subheader("1. 수정할 품목 선택 및 단가 설정")
    with col_undo:
        if st.session_state.history:
            if st.button("⏪ 방금 한 작업 되돌리기 (Undo)"):
                st.session_state.processed_data = st.session_state.history.pop()
                st.success("이전 상태로 되돌렸습니다!")
                st.rerun()
    
    unique_items = df[col_item_name].dropna().unique()
    selected_item = st.selectbox(f"A열({col_item_name})에서 수정할 항목을 선택하세요", unique_items)
    
    match = re.search(r'(\d{1,3}(?:,\d{3})*|\d+)원', str(selected_item)) 
    original_price_str = match.group(0) if match else ""
    current_price = int(match.group(1).replace(',', '')) if match else 0
        
    col1, col2 = st.columns(2)
    with col1:
        new_price = st.number_input("단가(원) - 변경 시 A열 이름과 옵션가에 자동 반영됩니다", value=current_price, step=100)
    with col2:
        base_price = st.number_input("기준가(원) 입력", value=52500, step=100)
    
    st.markdown("### 2. 새로운 중량 추가")
    weight_input = st.text_area("중량 리스트를 줄바꿈(Enter)으로 구분하여 입력하세요.", height=150)
    
    if st.button(f"👉 '{selected_item}' 작업 적용 및 임시 저장"):
        st.session_state.history.append(df.copy())
        
        if original_price_str:
            new_item_name = str(selected_item).replace(original_price_str, f"{new_price}원")
        else:
            new_item_name = str(selected_item)
            
        item_rows = df[df[col_item_name] == selected_item].copy()
        
        # 💡 [2번째 열] 중량 서식 추출
        sample_b = item_rows['중량'].iloc[0] if len(item_rows) > 0 else "0kg"
        num_match = re.search(r'(\d+\.?\d*)', str(sample_b))
        if num_match:
            prefix = str(sample_b)[:num_match.start()]
            suffix = str(sample_b)[num_match.end():]
        else:
            prefix, suffix = "", "kg"
            
        # 💡 [5번째 열] 관리코드 서식 추출 (요청하신 부분 추가!)
        sample_e = item_rows['관리코드'].iloc[0] if len(item_rows) > 0 and '관리코드' in item_rows.columns else "0kg"
        num_match_e = re.search(r'(\d+\.?\d*)', str(sample_e))
        if num_match_e:
            prefix_e = str(sample_e)[:num_match_e.start()]
            suffix_e = str(sample_e)[num_match_e.end():]
        else:
            prefix_e, suffix_e = "", "kg"
            
        base_sort_1 = item_rows['__sort_1'].min() if not item_rows.empty else df['__sort_1'].max() + 1
            
        item_rows['재고수량'] = pd.to_numeric(item_rows['재고수량'], errors='coerce').fillna(0)
        item_rows = item_rows[item_rows['재고수량'] > 0]
        
        def extract_num(text):
            m = re.search(r'(\d+\.?\d*)', str(text))
            return float(m.group(1)) if m else 0.0
            
        if not item_rows.empty:
            item_rows['numeric_weight'] = item_rows['중량'].apply(extract_num)
            item_rows['옵션가'] = (item_rows['numeric_weight'] * new_price - base_price).apply(lambda x: int(x / 10) * 10)
            item_rows[col_item_name] = new_item_name
            item_rows['__sort_1'] = base_sort_1
            item_rows['__sort_2'] = item_rows['numeric_weight']
            
        new_rows_data = []
        weights = weight_input.strip().split('\n')
        for w_str in weights:
            w_str = w_str.strip()
            if not w_str: continue
            
            w_num_match = re.search(r'(\d+\.?\d*)', w_str)
            if w_num_match:
                w_num = float(w_num_match.group(1))
                opt_price = int((w_num * new_price - base_price) / 10) * 10
                
                # 중량과 관리코드에 각각의 기존 서식을 입혀줌
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
        st.success(f"✅ '{new_item_name}' 작업 완료! (관리코드 서식도 똑같이 반영되었습니다)")
        st.rerun()

    st.markdown("---")
    st.subheader("3. 최종 결과물 확인 및 다운로드")
    
    display_df = st.session_state.processed_data.drop(columns=['__sort_1', '__sort_2'], errors='ignore')
    st.dataframe(display_df)
    
    xls_buffer = io.BytesIO()
    try:
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Sheet1')
        
        columns = display_df.columns.tolist()
        for col_idx, col_name in enumerate(columns):
            ws.write(0, col_idx, str(col_name))
            
        for row_idx, row in enumerate(display_df.values):
            for col_idx, val in enumerate(row):
                if pd.isna(val): 
                    val = ""
                elif not isinstance(val, (int, float)): 
                    val = str(val)
                ws.write(row_idx + 1, col_idx, val)
                
        wb.save(xls_buffer)
        
        st.download_button(
            label="💾 모든 변경사항 최종 다운로드 (업로드용 XLS 파일)",
            data=xls_buffer.getvalue(),
            file_name="최종수정본_옵션조합.xls",
            mime="application/vnd.ms-excel"
        )
    except Exception as e:
        st.error(f"엑셀 저장 중 오류가 발생했습니다: {e}")