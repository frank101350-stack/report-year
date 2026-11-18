import streamlit as st
import pandas as pd
import numpy as np
import re

# 1. 화면 기본 설정
st.set_page_config(page_title="Bảng điều khiển Doanh thu Bán hàng", layout="wide")

st.title("📈 Bảng điều khiển Doanh thu & Tốc độ tăng trưởng")
st.markdown("Vui lòng tải lên tệp Excel doanh thu hàng tháng dưới đây, kết quả phân tích sẽ được tính toán tự động.")

# 파일 업로드 컴포넌트
uploaded_file = st.file_uploader("Vui lòng tải lên tệp Excel (.xlsx) doanh thu mới nhất.", type=["xlsx"])

@st.cache_data
def load_data(file):
    df = pd.read_excel(file)
    df.columns = df.iloc[2]
    df = df.iloc[3:].reset_index(drop=True)
    
    df.columns = [str(c).strip() for c in df.columns]
    
    id_vars = ['CHAIN', 'AREA', 'SALE REP new']
    value_vars = [c for c in df.columns if 'ACTUAL' in str(c).upper()]
    
    df_melt = df.melt(id_vars=id_vars, value_vars=value_vars, var_name='Period', value_name='Sales')
    df_melt['Sales'] = pd.to_numeric(df_melt['Sales'], errors='coerce').fillna(0)
    
    def extract_month_year(period_str):
        match = re.search(r'([A-Za-z]+)\.?\s*(\d{4})', str(period_str))
        if match:
            return match.group(1).upper()[:3], int(match.group(2))
        return None, None
        
    df_melt['Month_Str'], df_melt['Year'] = zip(*df_melt['Period'].apply(extract_month_year))
    
    month_map = {'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6, 
                 'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12}
    df_melt['Month'] = df_melt['Month_Str'].map(month_map)
    
    df_melt = df_melt.dropna(subset=['Year', 'Month'])
    df_melt['Year'] = df_melt['Year'].astype(int)
    df_melt['Month'] = df_melt['Month'].astype(int)
    
    return df_melt

if uploaded_file is not None:
    try:
        df = load_data(uploaded_file)
        
        st.sidebar.header("🔍 Bộ lọc tìm kiếm")

        years = sorted([int(y) for y in df['Year'].unique()], reverse=True)
        years = ['Tất cả'] + years
        
        chains = ['Tất cả'] + sorted([str(x) for x in df['CHAIN'].dropna().unique()])
        areas = ['Tất cả'] + sorted([str(x) for x in df['AREA'].dropna().unique()])
        sales_reps = ['Tất cả'] + sorted([str(x) for x in df['SALE REP new'].dropna().unique()])

        selected_year = st.sidebar.selectbox("📅 Năm", years)
        selected_chain = st.sidebar.selectbox("🏢 Chuỗi (Cột A)", chains)
        selected_area = st.sidebar.selectbox("🌍 Khu vực (Cột E)", areas)
        selected_rep = st.sidebar.selectbox("👤 Nhân viên bán hàng (Cột F)", sales_reps)

        filtered_df = df.copy()
        if selected_chain != 'Tất cả':
            filtered_df = filtered_df[filtered_df['CHAIN'] == selected_chain]
        if selected_area != 'Tất cả':
            filtered_df = filtered_df[filtered_df['AREA'] == selected_area]
        if selected_rep != 'Tất cả':
            filtered_df = filtered_df[filtered_df['SALE REP new'] == selected_rep]

        monthly_sales = filtered_df.pivot_table(index='Month', columns='Year', values='Sales', aggfunc='sum').fillna(0)
        
        # ⭐️ [핵심 추가 1] 1~12월 데이터를 모두 더한 '합계' 행(Total)을 가장 아래에 추가합니다.
        monthly_sales.loc['Total'] = monthly_sales.sum()
        
        # 인덱스(행 이름)를 베트남어 월 이름과 합계(Tổng cộng)로 예쁘게 변경합니다.
        new_index = []
        for m in monthly_sales.index:
            if m == 'Total':
                new_index.append('Tổng cộng')
            else:
                new_index.append(f"Tháng {int(m)}")
        monthly_sales.index = new_index

        st.subheader("📊 Tình hình doanh thu hàng tháng (Tháng 1~12)")
        display_cols = [selected_year] if selected_year != 'Tất cả' else monthly_sales.columns
        st.dataframe(monthly_sales[display_cols].style.format("{:,.0f}"), use_container_width=True)

        st.divider()

        st.subheader("🚀 Phân tích tốc độ tăng trưởng doanh thu (%)")

        target_year = selected_year if selected_year != 'Tất cả' else monthly_sales.columns[-1]
        prev_year = target_year - 1

        if target_year in monthly_sales.columns and prev_year in monthly_sales.columns:
            
            # ⭐️ 합계 행을 제외한 순수 '월' 데이터만 뽑아서 YTD 비교를 수행합니다.
            months_only = [m for m in monthly_sales.index if m != 'Tổng cộng']
            valid_months = [m for m in months_only if monthly_sales.loc[m, target_year] > 0]
            
            current_year_total = monthly_sales.loc[valid_months, target_year].sum()
            prev_year_total_comparable = monthly_sales.loc[valid_months, prev_year].sum()
            
            annual_yoy = ((current_year_total / prev_year_total_comparable) - 1) * 100 if prev_year_total_comparable != 0 else 0
            
            months_str = ", ".join(valid_months) if len(valid_months) > 0 else "Không có dữ liệu"
            
            st.metric(label=f"🏆 Tổng doanh thu năm {target_year} (So với cùng kỳ: {months_str})", 
                      value=f"{current_year_total:,.0f}", 
                      delta=f"{annual_yoy:.2f}%")
                      
            growth_df = pd.DataFrame(index=monthly_sales.index)
            
            col_sales_target = f'Doanh thu năm {target_year}'
            col_mom = 'Tăng trưởng hàng tháng (MoM)'
            col_yoy = 'Tăng trưởng so với cùng kỳ năm trước (YoY)'
            
            growth_df[col_sales_target] = monthly_sales[target_year]
            
            # ⭐️ [핵심 추가 2] MoM 계산 (합계 행은 빼고 1~12월 끼리만 계산합니다)
            mom_temp = monthly_sales.loc[months_only, target_year].pct_change() * 100
            
            if 'Tháng 1' in monthly_sales.index and 'Tháng 12' in monthly_sales.index:
                prev_dec_sales = monthly_sales.loc['Tháng 12', prev_year]
                jan_sales = monthly_sales.loc['Tháng 1', target_year]
                if prev_dec_sales > 0:
                    mom_temp['Tháng 1'] = ((jan_sales / prev_dec_sales) - 1) * 100
            
            growth_df[col_mom] = mom_temp
            
            # YoY 계산 (기본 1~12월 비율 계산)
            growth_df[col_yoy] = ((monthly_sales[target_year] / monthly_sales[prev_year]) - 1) * 100
            
            # 빈칸 및 에러 처리 (기본값을 0으로 덮어씀)
            growth_df = growth_df.replace([np.inf, -np.inf], 0).fillna(0)
            
            # ⭐️ [핵심 추가 3] 합계(Tổng cộng) 행의 예외 처리
            growth_df.loc['Tổng cộng', col_mom] = np.nan # 합계의 MoM은 공란(NaN)으로 비워둡니다.
            growth_df.loc['Tổng cộng', col_yoy] = annual_yoy # 합계의 YoY는 공평한 YTD 성장률로 덮어씌웁니다!
            
            st.markdown(f"**Bảng chi tiết tốc độ tăng trưởng (Năm {target_year})**")
            
            # NaN 값(합계 행의 MoM)을 화면에 표시할 때만 0% 대신 깔끔한 빈칸으로 숨겨주는 포맷팅 마법
            st.dataframe(growth_df.style.format(
                formatter={
                    col_sales_target: '{:,.0f}',
                    col_mom: lambda x: '' if pd.isna(x) else f'{x:.2f}%',
                    col_yoy: '{:.2f}%'
                }
            ), use_container_width=True)

        else:
            st.info(f"Không có đủ dữ liệu của năm trước ({prev_year}) để tính toán tốc độ tăng trưởng cho năm {target_year} mà bạn đã chọn.")
            
    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi xử lý tệp. Vui lòng kiểm tra lại định dạng tệp. (Lỗi: {e})")
else:
    st.info("👆 Vui lòng kéo và thả tệp Excel cập nhật hàng tháng vào khu vực phía trên.")