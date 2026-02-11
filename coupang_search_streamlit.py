"""
쿠팡 입찰 검색 - Streamlit 웹 버전
설치 불필요, 브라우저에서 바로 실행
"""
import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

# 페이지 설정
st.set_page_config(
    page_title="쿠팡 입찰 검색 v6.0",
    page_icon="🔍",
    layout="wide"
)

# CSS 스타일
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem 0;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #2ca02c;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .stAlert {
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

def safe_to_float(value):
    """안전한 float 변환"""
    try:
        if value is None or pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value) if not pd.isna(value) else 0.0
        
        s = str(value).strip()
        if not s or s.lower() in ['', 'nan', 'none', 'nat']:
            return 0.0
        
        s = s.replace(',', '').replace('[', '').replace(']', '')
        s = s.replace('(', '').replace(')', '')
        
        parts = s.split()
        if parts:
            s = parts[0]
        
        return float(s) if s else 0.0
    except:
        return 0.0

def parse_advanced_search(search_text, text_series):
    """
    고급 검색 파싱 (간단 버전)
    
    지원:
    - "정확한 문구"
    - 단어1 AND 단어2
    - 단어1 OR 단어2
    - 단어1 NOT 단어2
    """
    if not search_text or not search_text.strip():
        return pd.Series([True] * len(text_series), index=text_series.index)
    
    search_text = search_text.strip()
    
    # AND 검색
    if ' AND ' in search_text:
        terms = [t.strip() for t in search_text.split(' AND ')]
        result = pd.Series([True] * len(text_series), index=text_series.index)
        
        for term in terms:
            # 따옴표 제거
            if term.startswith('"') and term.endswith('"'):
                term = term[1:-1]
                result = result & text_series.astype(str).str.contains(term, case=False, na=False, regex=False)
            else:
                result = result & text_series.astype(str).str.contains(term, case=False, na=False)
        
        return result
    
    # OR 검색
    elif ' OR ' in search_text:
        terms = [t.strip() for t in search_text.split(' OR ')]
        result = pd.Series([False] * len(text_series), index=text_series.index)
        
        for term in terms:
            if term.startswith('"') and term.endswith('"'):
                term = term[1:-1]
                result = result | text_series.astype(str).str.contains(term, case=False, na=False, regex=False)
            else:
                result = result | text_series.astype(str).str.contains(term, case=False, na=False)
        
        return result
    
    # NOT 검색
    elif ' NOT ' in search_text:
        parts = search_text.split(' NOT ', 1)
        include_term = parts[0].strip()
        exclude_term = parts[1].strip()
        
        # 포함
        if include_term.startswith('"') and include_term.endswith('"'):
            include_term = include_term[1:-1]
            result = text_series.astype(str).str.contains(include_term, case=False, na=False, regex=False)
        else:
            result = text_series.astype(str).str.contains(include_term, case=False, na=False)
        
        # 제외
        if exclude_term.startswith('"') and exclude_term.endswith('"'):
            exclude_term = exclude_term[1:-1]
            exclude = text_series.astype(str).str.contains(exclude_term, case=False, na=False, regex=False)
        else:
            exclude = text_series.astype(str).str.contains(exclude_term, case=False, na=False)
        
        return result & ~exclude
    
    # 단순 검색
    else:
        if search_text.startswith('"') and search_text.endswith('"'):
            search_text = search_text[1:-1]
            return text_series.astype(str).str.contains(search_text, case=False, na=False, regex=False)
        else:
            return text_series.astype(str).str.contains(search_text, case=False, na=False)

@st.cache_data
def load_excel_files(uploaded_files):
    """엑셀 파일 로딩"""
    all_table = []
    all_detail = []
    loaded_files = []
    
    xl_categories = {
        'XLA': '주방/유/홈/펫',
        'XLE': '식품', 
        'XLW': '대형가전/가구',
        'XLB': '가전',
        'XLC': '패션퍼스널/스포츠화장지'
    }
    
    file_id = 1
    for uploaded_file in uploaded_files:
        try:
            # 표 시트
            df_raw = pd.read_excel(uploaded_file, sheet_name='표', header=None, dtype=str)
            
            h0 = [str(x) if pd.notna(x) else '' for x in df_raw.iloc[0]]
            h1 = [str(x) if pd.notna(x) else '' for x in df_raw.iloc[1]]
            
            cols = []
            for i in range(len(h0)):
                if h0[i] in xl_categories.keys():
                    cols.append(h0[i])
                elif h1[i] and h1[i] != 'nan':
                    cols.append(h1[i])
                else:
                    cols.append(f'col_{i}')
            
            df = df_raw.iloc[2:].copy()
            df.columns = cols
            df = df.reset_index(drop=True)
            
            if 'col_7' in df.columns:
                df.rename(columns={'col_7': '원가율'}, inplace=True)
            
            df['파일명'] = uploaded_file.name
            df['파일ID'] = int(file_id)
            
            for c in df.columns:
                if c not in ['파일ID']:
                    df[c] = df[c].astype(str)
            
            all_table.append(df)
            
            # 상세품목
            uploaded_file.seek(0)  # 파일 포인터 리셋
            wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
            ws = wb['상세품목']
            
            header_row = 1
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
                if 'NO.' in [str(c) for c in row if c]:
                    header_row = i
                    break
            
            wb.close()
            
            uploaded_file.seek(0)
            df_detail = pd.read_excel(uploaded_file, sheet_name='상세품목', header=header_row-1, dtype=str)
            df_detail['파일명'] = uploaded_file.name
            df_detail['파일ID'] = int(file_id)
            all_detail.append(df_detail)
            
            loaded_files.append({'id': int(file_id), 'name': uploaded_file.name})
            file_id += 1
            
        except Exception as e:
            st.error(f"파일 로드 오류 ({uploaded_file.name}): {str(e)}")
            continue
    
    df_table = pd.concat(all_table, ignore_index=True, sort=False) if all_table else pd.DataFrame()
    df_detail = pd.concat(all_detail, ignore_index=True, sort=False) if all_detail else pd.DataFrame()
    
    return df_table, df_detail, loaded_files

def main():
    # 헤더
    st.markdown('<div class="main-header">🔍 쿠팡 입찰 검색 v6.0 WEB</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 파일 업로드
    st.markdown('<div class="section-header">📁 파일 업로드</div>', unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader(
        "엑셀 파일 선택 (여러 개 가능)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )
    
    if not uploaded_files:
        st.info("👆 엑셀 파일을 업로드하세요")
        return
    
    # 파일 로딩
    with st.spinner('파일 로딩 중...'):
        df_table, df_detail, loaded_files = load_excel_files(uploaded_files)
    
    if df_table.empty:
        st.error("파일을 로드할 수 없습니다!")
        return
    
    st.success(f"✅ {len(loaded_files)}개 파일 로드 완료 | 입찰: {df_table['NO.'].nunique()}개 | 상세품목: {len(df_detail):,}개")
    
    st.markdown("---")
    
    # 검색 조건
    st.markdown('<div class="section-header">🔍 검색 조건</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("기본 필터")
        
        # FC 선택
        fc_list = ['전체'] + sorted(df_table['FC'].dropna().unique().tolist())
        selected_fc = st.selectbox("FC", fc_list)
        
        # 원가율
        st.write("원가율 (%)")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            rate_min = st.number_input("최소", min_value=0.0, max_value=100.0, value=0.0, step=0.1, key='rate_min')
        with col_r2:
            rate_max = st.number_input("최대", min_value=0.0, max_value=100.0, value=100.0, step=0.1, key='rate_max')
        
        # XL 선택
        st.write("**XL 카테고리**")
        xl_selections = {}
        xl_categories = {
            'XLA': '주방/유/홈/펫',
            'XLE': '식품', 
            'XLW': '대형가전/가구',
            'XLB': '가전',
            'XLC': '패션/스포츠'
        }
        
        for xl_code, xl_name in xl_categories.items():
            xl_selections[xl_code] = st.checkbox(f"{xl_code} - {xl_name}", key=f"xl_{xl_code}")
    
    with col2:
        st.subheader("상세품목 검색")
        
        # 고급 검색 모드
        advanced_mode = st.checkbox("🔧 고급 검색 모드 (AND/OR/NOT 지원)")
        
        if advanced_mode:
            st.info("""
            **고급 검색 문법:**
            - `"정확한 문구"` - 따옴표로 정확히 검색
            - `라면 AND 매운맛` - 모두 포함
            - `라면 OR 우동` - 하나라도 포함
            - `라면 NOT 컵` - 라면 포함, 컵 제외
            
            **예시:** `패션의류 OR 스포츠 OR 생활용품`
            """)
        
        # 검색 필드
        cate2_search = st.text_input("CATE2 (대분류)", placeholder="예: 식품", key="cate2")
        desc_search = st.text_input("상품명", placeholder="예: 라면 AND 매운맛", key="desc")
        cate4_search = st.text_input("CATE4 (소분류)", placeholder="예: 봉지", key="cate4")
        cate5_search = st.text_input("CATE5 (세분류)", placeholder="예: 5입", key="cate5")
    
    st.markdown("---")
    
    # 검색 버튼
    col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 3])
    with col_btn1:
        search_clicked = st.button("🔍 검색", type="primary", use_container_width=True)
    with col_btn2:
        if st.button("🔄 초기화", use_container_width=True):
            st.rerun()
    
    if not search_clicked:
        return
    
    # 검색 실행
    with st.spinner('검색 중...'):
        df_result = df_table.copy()
        
        # FC 필터
        if selected_fc != '전체':
            df_result = df_result[df_result['FC'] == selected_fc]
        
        # 원가율 필터
        df_result['원가율_numeric'] = df_result['원가율'].apply(lambda x: safe_to_float(x) * 100)
        df_result = df_result[
            (df_result['원가율_numeric'] >= rate_min) &
            (df_result['원가율_numeric'] <= rate_max)
        ]
        
        # XL 필터
        selected_xl = [xl for xl, selected in xl_selections.items() if selected]
        if selected_xl:
            mask = pd.Series([False] * len(df_result), index=df_result.index)
            for xl in selected_xl:
                if xl in df_result.columns:
                    nums = df_result[xl].apply(safe_to_float)
                    mask = mask | (nums > 0)
            df_result = df_result[mask]
            
            # XL 합계 계산
            df_result['선택XL_합계'] = 0.0
            for xl in selected_xl:
                if xl in df_result.columns:
                    df_result['선택XL_합계'] += df_result[xl].apply(safe_to_float)
        
        # 키워드 검색
        if cate2_search or desc_search or cate4_search or cate5_search:
            matching_nos = []
            
            for _, row in df_result[['NO.', '파일ID']].drop_duplicates().iterrows():
                no, fid = row['NO.'], int(row['파일ID'])
                details = df_detail[(df_detail['NO.'] == no) & (df_detail['파일ID'] == fid)]
                
                if len(details) == 0:
                    continue
                
                # 각 필드별 검색
                mask = pd.Series([True] * len(details), index=details.index)
                
                if cate2_search and 'CATE2' in details.columns:
                    if advanced_mode:
                        m = parse_advanced_search(cate2_search, details['CATE2'])
                    else:
                        m = details['CATE2'].astype(str).str.contains(cate2_search, case=False, na=False)
                    mask = mask & m
                
                if desc_search and 'DESCRIPTION' in details.columns:
                    if advanced_mode:
                        m = parse_advanced_search(desc_search, details['DESCRIPTION'])
                    else:
                        m = details['DESCRIPTION'].astype(str).str.contains(desc_search, case=False, na=False)
                    mask = mask & m
                
                if cate4_search and 'CATE4' in details.columns:
                    if advanced_mode:
                        m = parse_advanced_search(cate4_search, details['CATE4'])
                    else:
                        m = details['CATE4'].astype(str).str.contains(cate4_search, case=False, na=False)
                    mask = mask & m
                
                if cate5_search and 'CATE5' in details.columns:
                    if advanced_mode:
                        m = parse_advanced_search(cate5_search, details['CATE5'])
                    else:
                        m = details['CATE5'].astype(str).str.contains(cate5_search, case=False, na=False)
                    mask = mask & m
                
                if mask.sum() > 0:
                    matching_nos.append((no, fid))
            
            # 매칭된 NO.만 남기기
            if matching_nos:
                mask_final = pd.Series([False] * len(df_result), index=df_result.index)
                for no, fid in matching_nos:
                    mask_final = mask_final | ((df_result['NO.'] == no) & (df_result['파일ID'] == fid))
                df_result = df_result[mask_final]
            else:
                df_result = df_result.iloc[0:0]  # 빈 DataFrame
    
    # 결과 표시
    st.markdown("---")
    st.markdown('<div class="section-header">📊 검색 결과</div>', unsafe_allow_html=True)
    
    st.write(f"**검색 결과: {len(df_result)}건**")
    
    if len(df_result) > 0:
        # 표시할 컬럼 선택
        display_cols = ['파일ID', 'NO.', 'FC', 'PLT', '원가율']
        for xl in ['XLA', 'XLE', 'XLW', 'XLB', 'XLC']:
            if xl in df_result.columns:
                display_cols.append(xl)
        
        if '선택XL_합계' in df_result.columns:
            display_cols.append('선택XL_합계')
        
        # 데이터 포맷팅
        df_display = df_result[display_cols].copy()
        
        # 숫자 포맷
        for col in ['XLA', 'XLE', 'XLW', 'XLB', 'XLC', '선택XL_합계']:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda x: f"{int(safe_to_float(x)):,}" if safe_to_float(x) > 0 else "-")
        
        df_display['원가율'] = df_display['원가율'].apply(lambda x: f"{safe_to_float(x)*100:.2f}%")
        
        # 테이블 표시
        st.dataframe(
            df_display,
            use_container_width=True,
            height=400
        )
        
        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='검색결과')
        
        st.download_button(
            label="📥 엑셀 다운로드",
            data=output.getvalue(),
            file_name=f"검색결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("검색 결과가 없습니다!")
    
    # 푸터
    st.markdown("---")
    st.caption(f"쿠팡 입찰 검색 v6.0 WEB | 마지막 업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

if __name__ == "__main__":
    main()
