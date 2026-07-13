import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# --- Настройка страницы ---
st.set_page_config(page_title="HR Analytics Dashboard", layout="wide")

# --- Главное название внутри красной рамки ---
st.markdown(
    """
    <div style="border: 3px solid red; padding: 15px; border-radius: 10px; text-align: center; margin-bottom: 25px;">
        <h1 style="margin: 0; padding: 0; font-size: 2.5rem; color: inherit;">HR Дашборд</h1>
    </div>
    """,
    unsafe_allow_html=True
)

# --- Боковая панель для загрузки ---
with st.sidebar:
    st.header("Загрузка данных из Excel")
    st.markdown("Загрузите Excel-файлы.")
    recruitment_file = st.file_uploader("1. Excel-файл - найм сотрудников", type=['xlsx'])
    training_file = st.file_uploader("2. Excel-файл - обучение персонала", type=['xlsx'])
    attestation_file = st.file_uploader("3. Excel-файл - аттестация сотрудников", type=['xlsx'])
    staff_file = st.file_uploader("4. Excel-файл - общее кол-во сотрудники", type=['xlsx'])

# --- Вспомогательные функции ---
def find_sheet_and_header(file_bytes, keywords):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    for sheet in xls.sheet_names:
        df_preview = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=15)
        for idx, row in df_preview.iterrows():
            row_str = [str(v).lower().strip() if v is not None and not (isinstance(v, float) and np.isnan(v)) else '' for v in row]
            if any(any(kw in val for kw in keywords) for val in row_str):
                return sheet, idx
    return xls.sheet_names[0], 0

def parse_test_score(val):
    if pd.isna(val):
        return np.nan
    val_str = str(val).replace(' ', '').replace('\n', '')
    if '/' not in val_str:
        return np.nan
    parts = val_str.split('/')
    if len(parts) != 2:
        return np.nan
    try:
        p1 = float(parts[0])
        p2 = float(parts[1])
        if p1 == 54:
            return (p2 / 54.0) * 100
        elif p2 == 54:
            return (p1 / 54.0) * 100
        else:
            return (p1 / p2) * 100 if p2 != 0 else np.nan
    except (ValueError, TypeError):
        return np.nan

def categorize_region(val):
    val_str = str(val).lower()
    if 'филиал' in val_str: return 'Филиал'
    if 'мхб' in val_str or 'цбо' in val_str: return 'МХБ / ЦБО'
    if 'го' in val_str or 'ho' in val_str or 'головной' in val_str or 'центральный' in val_str: return 'Головной офис'
    return 'Другое'

# --- Загрузка и кэширование данных ---
@st.cache_data(show_spinner=False)
def load_recruitment_data(file_bytes: bytes):
    sheet, header_idx = find_sheet_and_header(file_bytes, ['фио кандидата', 'вакансия'])
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, header=header_idx)
    df.columns = df.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()

    if 'ФИО кандидата' in df.columns: df = df.dropna(subset=['ФИО кандидата'])
    if 'Источник подбора' in df.columns: df['Источник подбора'] = df['Источник подбора'].replace({'Анкета': 'Анкета - референс'})

    if 'Дата регистрации' in df.columns: df['Дата регистрации'] = pd.to_datetime(df['Дата регистрации'], errors='coerce')
    if 'Дата стажировки' in df.columns: df['Дата стажировки'] = pd.to_datetime(df['Дата стажировки'], errors='coerce')
    if 'Дата завершения стажировки' in df.columns: df['Дата завершения стажировки'] = pd.to_datetime(df['Дата завершения стажировки'], errors='coerce')

    if 'Дни стажировки' not in df.columns and 'Дата стажировки' in df.columns and 'Дата завершения стажировки' in df.columns:
        df['Дни стажировки'] = (df['Дата завершения стажировки'] - df['Дата стажировки']).dt.days

    if 'Возраст' in df.columns: df['Возраст'] = pd.to_numeric(df['Возраст'], errors='coerce')
    if 'Тест' in df.columns: df['Успешность_теста_%'] = df['Тест'].apply(parse_test_score).astype(np.float32)
    return df

@st.cache_data(show_spinner=False)
def load_training_data(file_bytes: bytes):
    sheet, header_idx = find_sheet_and_header(file_bytes, ['фио сотрудника', 'название курса'])
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, header=header_idx)
    df.columns = df.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()
    if 'ФИО сотрудника' in df.columns: df = df.dropna(subset=['ФИО сотрудника'])
    
    cols_to_keep = ['ФИО сотрудника', 'Дата прохождения', 'Отдел', 'Должность',
                    'Название курса', 'Часы обучения', 'Результат теста (%)',
                    'План', 'Факт', 'Присутствие', 'Статус']
    existing_cols = [col for col in cols_to_keep if col in df.columns]
    df = df[existing_cols]
    
    if 'Дата прохождения' in df.columns: df['Дата прохождения'] = pd.to_datetime(df['Дата прохождения'], errors='coerce')
    if 'Часы обучения' in df.columns: df['Часы обучения'] = pd.to_numeric(df['Часы обучения'], errors='coerce', downcast='unsigned')
    return df

@st.cache_data(show_spinner=False)
def load_generic_data(file_bytes: bytes, keywords: list):
    sheet, header_idx = find_sheet_and_header(file_bytes, keywords)
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, header=header_idx)
    df.columns = df.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()
    return df

# --- Сохранение файлов в Session State ---
if recruitment_file: st.session_state['rec_bytes'] = recruitment_file.getvalue()
if training_file: st.session_state['tr_bytes'] = training_file.getvalue()
if attestation_file: st.session_state['att_bytes'] = attestation_file.getvalue()
if staff_file: st.session_state['staff_bytes'] = staff_file.getvalue()

df_recruitment = load_recruitment_data(st.session_state['rec_bytes']) if 'rec_bytes' in st.session_state else None
df_training = load_training_data(st.session_state['tr_bytes']) if 'tr_bytes' in st.session_state else None
df_attestation = load_generic_data(st.session_state['att_bytes'], ['фио', 'категория', 'аттестация', 'дарача', 'должность']) if 'att_bytes' in st.session_state else None
df_staff = load_generic_data(st.session_state['staff_bytes'], ['фио', 'дата', 'приема', 'подразделение']) if 'staff_bytes' in st.session_state else None

if not any([df_recruitment is not None, df_training is not None, df_attestation is not None, df_staff is not None]):
    st.info("Загрузите файлы Excel в боковой панели слева для начала работы.")
    st.stop()

# --- Вкладки ---
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Аналитика найма", "Аналитика стажёров", "Аттестация", "Корпоративное обучение", "Общий штат"])

# === ВКЛАДКА 1: НАЙМ ===
with tab1:
    if df_recruitment is not None:
        with st.form("recruitment_filters"):
            safe_dates_rec = df_recruitment['Дата регистрации'].dropna()
            min_date_rec = safe_dates_rec.min().date() if not safe_dates_rec.empty else pd.to_datetime('today').date()
            max_date_rec = safe_dates_rec.max().date() if not safe_dates_rec.empty else pd.to_datetime('today').date()
            date_range_rec = st.date_input("Период регистрации", [min_date_rec, max_date_rec], min_value=min_date_rec, max_value=max_date_rec)
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                vacs = df_recruitment['Вакансия'].dropna().unique() if 'Вакансия' in df_recruitment.columns else []
                selected_vacancies = st.multiselect("Выберите вакансии", vacs)
            with col2:
                typs = df_recruitment['Типы'].dropna().unique() if 'Типы' in df_recruitment.columns else []
                selected_types = st.multiselect("Тип подразделения", typs)
            with col3:
                srcs = df_recruitment['Источник подбора'].dropna().unique() if 'Источник подбора' in df_recruitment.columns else []
                selected_sources = st.multiselect("Источник привлечения", srcs)
            submit = st.form_submit_button("Применить фильтры ⚡")

        filtered_rec = df_recruitment.copy()
        if isinstance(date_range_rec, tuple) and len(date_range_rec) == 2:
            start_date, end_date = date_range_rec
            filtered_rec = filtered_rec[(filtered_rec['Дата регистрации'].dt.date >= start_date) & (filtered_rec['Дата регистрации'].dt.date <= end_date)]

        if selected_vacancies: filtered_rec = filtered_rec[filtered_rec['Вакансия'].isin(selected_vacancies)]
        if selected_types: filtered_rec = filtered_rec[filtered_rec['Типы'].isin(selected_types)]
        if selected_sources: filtered_rec = filtered_rec[filtered_rec['Источник подбора'].isin(selected_sources)]

        total_candidates = len(filtered_rec)
        hired_mask = filtered_rec['Принят на работу'].isin(['Штат', 'Стажёр', 'Волонтёр']) if 'Принят на работу' in filtered_rec.columns else pd.Series(False, index=filtered_rec.index)
        hired = hired_mask.sum()
        conv_rate = (hired / total_candidates * 100) if total_candidates > 0 else 0
        avg_age = filtered_rec[hired_mask]['Возраст'].mean() if 'Возраст' in filtered_rec.columns else np.nan
        
        if 'Пол' in filtered_rec.columns:
            hired_gender = filtered_rec[hired_mask]['Пол'].astype(str).str.strip().str.upper()
            m_count, f_count = hired_gender.str.startswith('М').sum(), hired_gender.str.startswith('Ж').sum()
            gender_text = f"М: {m_count} | Ж: {f_count}"
        else:
            gender_text = "Нет данных"

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Общее кол. кандидатов:", f"{total_candidates} чел.")
        m2.metric("Успешно трудоустроено", f"{hired} чел.")
        m3.metric("Эффективность подбора персонала:", f"{conv_rate:.1f}%")
        m4.metric("М / Ж (по найму)", gender_text)
        m5.metric("Ср. возраст (по найму)", f"{avg_age:.1f} лет" if not pd.isna(avg_age) else "Нет данных")

        st.divider()

        f_col1, f_col2 = st.columns(2)
        with f_col1:
            st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Общее кол-во заявок</b></div><br>", unsafe_allow_html=True)
            if 'Источник подбора' in filtered_rec.columns:
                src_traffic = filtered_rec['Источник подбора'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
                src_traffic.columns = ['Источник', 'Всего кандидатов']
                fig_source_vol = px.pie(src_traffic, names='Источник', values='Всего кандидатов', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_source_vol.update_traces(textposition='auto', textinfo='label+value')
                st.plotly_chart(fig_source_vol, use_container_width=True, key="fig_source_vol_tab1")

        with f_col2:
            st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Эффективность подбора персонала</b></div><br>", unsafe_allow_html=True)
            if 'Источник подбора' in filtered_rec.columns:
                src_df = filtered_rec.copy()
                src_df['Источник подбора'] = src_df['Источник подбора'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
                src_df['Is_Hired'] = hired_mask.astype(int)
                src_conv = src_df.groupby('Источник подбора').agg(Total=('Is_Hired', 'count'), Hired=('Is_Hired', 'sum')).reset_index()
                src_conv['Конверсия %'] = (src_conv['Hired'] / src_conv['Total'] * 100).round(1)
                src_conv = src_conv[src_conv['Total'] > 0].sort_values('Конверсия %', ascending=True)
                src_conv['Текст на графике'] = src_conv['Конверсия %'].astype(str) + '% (' + src_conv['Hired'].astype(str) + ' из ' + src_conv['Total'].astype(str) + ')'
                fig_src_conv = px.bar(src_conv, x='Конверсия %', y='Источник подбора', orientation='h', text='Текст на графике', color='Конверсия %', color_continuous_scale='Greens')
                fig_src_conv.update_traces(textposition='outside')
                fig_src_conv.update_layout(xaxis_range=[0, src_conv['Конверсия %'].max() * 1.3 if src_conv['Конверсия %'].max() > 0 else 100])
                st.plotly_chart(fig_src_conv, use_container_width=True, key="fig_src_conv_tab1")

        st.divider()

        # --- СТРУКТУРА КАНДИДАТОВ (ВСЕ ЗАЯВКИ) ---
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Общее количесвто кандидатов: Филиалы, МХБ и ГО</b></div><br>", unsafe_allow_html=True)
        r1, r2, r3 = st.columns(3)

        if 'Типы' in filtered_rec.columns and 'Все подразделения' in filtered_rec.columns:
            filtered_rec_copy = filtered_rec.copy()
            filtered_rec_copy['Типы_str'] = filtered_rec_copy['Типы'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            filtered_rec_copy['Подразделение_str'] = filtered_rec_copy['Все подразделения'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})

            with r1:
                st.markdown("**1. Общее количество кандидатов по ГО/Филиалу/ЦБО**")
                types_data = filtered_rec_copy['Типы_str'].value_counts(dropna=False).reset_index()
                types_data.columns = ['Тип', 'Количество']
                if not types_data.empty:
                    fig_types = px.pie(types_data, names='Тип', values='Количество', hole=0.4)
                    fig_types.update_traces(textposition='auto', textinfo='label+percent') 
                    fig_types.update_layout(showlegend=False)
                    st.plotly_chart(fig_types, use_container_width=True, key="fig_types_tab1")

            with r2:
                st.markdown("**2. Общее количество кандидатов по Филиалам**")
                branches_df = filtered_rec_copy[filtered_rec_copy['Типы_str'].str.contains('Филиал|филиал', na=False, case=False)]
                branch_data = branches_df['Подразделение_str'].value_counts(dropna=False).reset_index()
                branch_data.columns = ['Филиал', 'Количество']
                if not branch_data.empty:
                    fig_branch = px.pie(branch_data, names='Филиал', values='Количество', hole=0.4)
                    fig_branch.update_traces(textposition='auto', textinfo='label+value')
                    fig_branch.update_layout(showlegend=False)
                    st.plotly_chart(fig_branch, use_container_width=True, key="fig_branch_tab1")

            with r3:
                st.markdown("**3. Общее количество кандидатов по ЦБО**")
                mhb_df = filtered_rec_copy[filtered_rec_copy['Типы_str'].str.contains('МХБ|ЦБО|мхб|цбо', na=False, case=False)]
                mhb_data = mhb_df['Подразделение_str'].value_counts(dropna=False).reset_index()
                mhb_data.columns = ['МХБ/ЦБО', 'Количество']
                if not mhb_data.empty:
                    fig_mhb = px.pie(mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                    fig_mhb.update_traces(textposition='auto', textinfo='label+value')
                    fig_mhb.update_layout(showlegend=False)
                    st.plotly_chart(fig_mhb, use_container_width=True, key="fig_mhb_tab1")

        st.divider()

        # --- СТРУКТУРА УСПЕШНОГО НАЙМА (ТОЛЬКО ТРУДОУСТРОЕННЫЕ) ---
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Успешный найм кандидатов: Филиалы, МХБ и ГО (Только принятые)</b></div><br>", unsafe_allow_html=True)
        
        hr1, hr2, hr3 = st.columns(3)
        hired_df = filtered_rec[hired_mask].copy()

        if 'Типы' in hired_df.columns and 'Все подразделения' in hired_df.columns:
            hired_df['Типы_str'] = hired_df['Типы'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            hired_df['Подразделение_str'] = hired_df['Все подразделения'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            hired_df['Регион_calc'] = hired_df['Типы_str'].apply(categorize_region)

            with hr1:
                st.markdown("**1. Успешный найм кандидатов по ГО/Филиалу/ЦБО**")
                hired_reg_data = hired_df['Регион_calc'].value_counts().reset_index()
                hired_reg_data.columns = ['Регион', 'Количество']
                hired_reg_data = hired_reg_data[hired_reg_data['Количество'] > 0]
                if not hired_reg_data.empty:
                    fig_hired_reg = px.pie(hired_reg_data, names='Регион', values='Количество', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                    fig_hired_reg.update_traces(textposition='auto', textinfo='label+percent') 
                    fig_hired_reg.update_layout(showlegend=False)
                    st.plotly_chart(fig_hired_reg, use_container_width=True, key="fig_hired_reg_tab1")
                else:
                    st.info("Нет данных о найме")

            with hr2:
                st.markdown("**2. Успешный найм кандидатов по Филиалам**")
                hired_branches = hired_df[hired_df['Регион_calc'] == 'Филиал']
                hired_branch_data = hired_branches['Подразделение_str'].value_counts().reset_index()
                hired_branch_data.columns = ['Филиал', 'Количество']
                hired_branch_data = hired_branch_data[hired_branch_data['Количество'] > 0]
                if not hired_branch_data.empty:
                    fig_hired_branch = px.pie(hired_branch_data, names='Филиал', values='Количество', hole=0.4)
                    fig_hired_branch.update_traces(textposition='auto', textinfo='label+value')
                    fig_hired_branch.update_layout(showlegend=False)
                    st.plotly_chart(fig_hired_branch, use_container_width=True, key="fig_hired_branch_tab1")
                else:
                    st.info("В филиалах нет найма за этот период")

            with hr3:
                st.markdown("**3. Успешный найм кандидатов по ЦБО**")
                hired_mhb = hired_df[hired_df['Регион_calc'] == 'МХБ / ЦБО']
                hired_mhb_data = hired_mhb['Подразделение_str'].value_counts().reset_index()
                hired_mhb_data.columns = ['МХБ/ЦБО', 'Количество']
                hired_mhb_data = hired_mhb_data[hired_mhb_data['Количество'] > 0]
                if not hired_mhb_data.empty:
                    fig_hired_mhb = px.pie(hired_mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                    fig_hired_mhb.update_traces(textposition='auto', textinfo='label+value')
                    fig_hired_mhb.update_layout(showlegend=False)
                    st.plotly_chart(fig_hired_mhb, use_container_width=True, key="fig_hired_mhb_tab1")
                else:
                    st.info("В МХБ / ЦБО нет найма за этот период")

        st.divider()

        # --- ДИНАМИКА ПО МЕСЯЦАМ ---
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Регистрация кандидатов (по месяцам)</b></div><br>", unsafe_allow_html=True)
        if 'Дата регистрации' in filtered_rec.columns:
            trend_data = filtered_rec.dropna(subset=['Дата регистрации']).copy()
            if not trend_data.empty:
                trend_data['Период'] = trend_data['Дата регистрации'].dt.strftime('%Y-%m')
                trend_counts = trend_data['Период'].value_counts().reset_index().sort_values('Период')
                trend_counts.columns = ['Период', 'Кандидаты']
                
                fig_trend = px.line(trend_counts, x='Период', y='Кандидаты', markers=True, color_discrete_sequence=['#ff7f0e'])
                fig_trend.update_xaxes(type='category')
                st.plotly_chart(fig_trend, use_container_width=True, key="fig_trend_tab1")
            else:
                st.info("Нет данных по датам для построения графика.")
    else:
        st.warning("Файл по найму не загружен.")

# === ВКЛАДКА 2: СТАЖЁРЫ ===
with tab2:
    if df_recruitment is not None and 'Принят на работу' in df_recruitment.columns:
        interns_full = df_recruitment[df_recruitment['Принят на работу'].isin(['Стажёр'])].copy()
        
        with st.form("intern_filters"):
            safe_dates_int = interns_full['Дата регистрации'].dropna()
            min_date_int = safe_dates_int.min().date() if not safe_dates_int.empty else pd.to_datetime('today').date()
            max_date_int = safe_dates_int.max().date() if not safe_dates_int.empty else pd.to_datetime('today').date()
            date_range_int = st.date_input("Период регистрации стажёров", [min_date_int, max_date_int], min_value=min_date_int, max_value=max_date_int)
            submit_int = st.form_submit_button("Применить фильтры ⚡")

        filtered_int = interns_full.copy()
        if isinstance(date_range_int, tuple) and len(date_range_int) == 2:
            start_date_int, end_date_int = date_range_int
            filtered_int = filtered_int[(filtered_int['Дата регистрации'].dt.date >= start_date_int) & (filtered_int['Дата регистрации'].dt.date <= end_date_int)]

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Показатели программы стажировки</b></div><br>", unsafe_allow_html=True)
        
        total_interns = len(filtered_int)
        avg_intern_days = filtered_int['Дни стажировки'].mean() if 'Дни стажировки' in filtered_int.columns else np.nan
        avg_test_int = filtered_int['Успешность_теста_%'].mean() if 'Успешность_теста_%' in filtered_int.columns else np.nan
        avg_age_int = filtered_int['Возраст'].mean() if 'Возраст' in filtered_int.columns else np.nan

        im1, im2, im3, im4 = st.columns(4)
        im1.metric("Всего стажёров", f"{total_interns} чел.")
        im2.metric("Средний балл тестирования", f"{avg_test_int:.1f}%" if not pd.isna(avg_test_int) else "Нет данных")
        im3.metric("Средний возраст", f"{avg_age_int:.1f} лет" if not pd.isna(avg_age_int) else "Нет данных")
        im4.metric("Ср. дней стажировки", f"{avg_intern_days:.1f} дн." if not pd.isna(avg_intern_days) else "Нет данных")

        st.divider()
        
        # --- СРЕЗ СТАЖЁРОВ ПО ФИЛИАЛАМ И МХБ ---
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Стажёры: Филиалы, ЦБО, ГО</b></div><br>", unsafe_allow_html=True)
        
        int_r1, int_r2, int_r3 = st.columns(3)

        if 'Типы' in filtered_int.columns and 'Все подразделения' in filtered_int.columns:
            filtered_int['Типы_str'] = filtered_int['Типы'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            filtered_int['Подразделение_str'] = filtered_int['Все подразделения'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            filtered_int['Регион_calc'] = filtered_int['Типы_str'].apply(categorize_region)

            with int_r1:
                st.markdown("**1. Распределение стажёров: Филиалы, ЦБО, ГО **")
                int_reg_data = filtered_int['Регион_calc'].value_counts().reset_index()
                int_reg_data.columns = ['Регион', 'Количество']
                int_reg_data = int_reg_data[int_reg_data['Количество'] > 0]
                if not int_reg_data.empty:
                    fig_int_reg = px.pie(int_reg_data, names='Регион', values='Количество', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                    fig_int_reg.update_traces(textposition='auto', textinfo='label+percent')
                    fig_int_reg.update_layout(showlegend=False)
                    st.plotly_chart(fig_int_reg, use_container_width=True, key="fig_int_reg_tab2")
                else:
                    st.info("Нет данных")

            with int_r2:
                st.markdown("**2. Данные по филиалам**")
                int_branches = filtered_int[filtered_int['Регион_calc'] == 'Филиал']
                int_branch_data = int_branches['Подразделение_str'].value_counts().reset_index()
                int_branch_data.columns = ['Филиал', 'Количество']
                int_branch_data = int_branch_data[int_branch_data['Количество'] > 0]
                if not int_branch_data.empty:
                    fig_int_branch = px.pie(int_branch_data, names='Филиал', values='Количество', hole=0.4)
                    fig_int_branch.update_traces(textposition='auto', textinfo='label+value')
                    fig_int_branch.update_layout(showlegend=False)
                    st.plotly_chart(fig_int_branch, use_container_width=True, key="fig_int_branch_tab2")
                else:
                    st.info("В филиалах нет стажёров")

            with int_r3:
                st.markdown("**3. Данные по ЦБО**")
                int_mhb = filtered_int[filtered_int['Регион_calc'] == 'МХБ / ЦБО']
                int_mhb_data = int_mhb['Подразделение_str'].value_counts().reset_index()
                int_mhb_data.columns = ['МХБ/ЦБО', 'Количество']
                int_mhb_data = int_mhb_data[int_mhb_data['Количество'] > 0]
                if not int_mhb_data.empty:
                    fig_int_mhb = px.pie(int_mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                    fig_int_mhb.update_traces(textposition='auto', textinfo='label+value')
                    fig_int_mhb.update_layout(showlegend=False)
                    st.plotly_chart(fig_int_mhb, use_container_width=True, key="fig_int_mhb_tab2")
                else:
                    st.info("В ЦБО нет стажёров")

        st.divider()
        
        # --- ВАКАНСИИ (без статусов) ---
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Распределение стажёров по должностям</b></div><br>", unsafe_allow_html=True)
        
        if 'Вакансия' in filtered_int.columns:
            vac_counts = filtered_int['Вакансия'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
            vac_counts.columns = ['Вакансия', 'Количество']
            fig_vac_int = px.bar(vac_counts.head(15), x='Количество', y='Вакансия', orientation='h', text='Количество', color='Количество', color_continuous_scale='Purples')
            fig_vac_int.update_traces(textposition='outside')
            fig_vac_int.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_vac_int, use_container_width=True, key="fig_vac_int_tab2")

    else:
        st.warning("Нет данных по стажёрам. Загрузите файл по найму.")

# === ВКЛАДКА 3: АТТЕСТАЦИЯ ===
with tab3:
    if df_attestation is not None:
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Аттестация: ключевые показатели</b></div><br>", unsafe_allow_html=True)
        
        cols_lower = {c.lower(): c for c in df_attestation.columns}
        col_region = cols_lower.get('типы') or cols_lower.get('регион')
        col_subdiv = cols_lower.get('все подразделения') or cols_lower.get('подразделение') or cols_lower.get('отдел')

        df_att_filtered = df_attestation.copy()

        total_attested = len(df_att_filtered)
        st.metric("Количество аттестованных", f"{total_attested} чел.")
        st.divider()

        # --- СРЕЗ АТТЕСТОВАННЫХ ПО ФИЛИАЛАМ И МХБ ---
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Разбивка по подразделениям: Филиалы/ЦБО/ГО</b></div><br>", unsafe_allow_html=True)
        
        att_r1, att_r2, att_r3 = st.columns(3)
        
        if col_region and col_subdiv:
            df_att_filtered['Типы_str'] = df_att_filtered[col_region].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            df_att_filtered['Подразделение_str'] = df_att_filtered[col_subdiv].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            df_att_filtered['Регион_calc'] = df_att_filtered['Типы_str'].apply(categorize_region)

            with att_r1:
                st.markdown("**1. Аттестация сотрудников по подразделениям: Филиалы/ЦБО/ГО**")
                att_reg_data = df_att_filtered['Регион_calc'].value_counts().reset_index()
                att_reg_data.columns = ['Регион', 'Количество']
                att_reg_data = att_reg_data[att_reg_data['Количество'] > 0]
                if not att_reg_data.empty:
                    fig_att_reg = px.pie(att_reg_data, names='Регион', values='Количество', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set3)
                    fig_att_reg.update_traces(textposition='auto', textinfo='label+value')
                    fig_att_reg.update_layout(showlegend=False)
                    st.plotly_chart(fig_att_reg, use_container_width=True, key="fig_att_reg_tab3")
                else:
                    st.info("Нет данных")

            with att_r2:
                st.markdown("**2. Данные по Филиалам**")
                att_branches = df_att_filtered[df_att_filtered['Регион_calc'] == 'Филиал']
                att_branch_data = att_branches['Подразделение_str'].value_counts().reset_index()
                att_branch_data.columns = ['Филиал', 'Количество']
                att_branch_data = att_branch_data[att_branch_data['Количество'] > 0]
                if not att_branch_data.empty:
                    fig_att_branch = px.pie(att_branch_data, names='Филиал', values='Количество', hole=0.4)
                    fig_att_branch.update_traces(textposition='auto', textinfo='label+value')
                    fig_att_branch.update_layout(showlegend=False)
                    st.plotly_chart(fig_att_branch, use_container_width=True, key="fig_att_branch_tab3")
                else:
                    st.info("В филиалах нет данных")

            with att_r3:
                st.markdown("**3. Данные по ЦБО**")
                att_mhb = df_att_filtered[df_att_filtered['Регион_calc'] == 'МХБ / ЦБО']
                att_mhb_data = att_mhb['Подразделение_str'].value_counts().reset_index()
                att_mhb_data.columns = ['МХБ/ЦБО', 'Количество']
                att_mhb_data = att_mhb_data[att_mhb_data['Количество'] > 0]
                if not att_mhb_data.empty:
                    fig_att_mhb = px.pie(att_mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                    fig_att_mhb.update_traces(textposition='auto', textinfo='label+value')
                    fig_att_mhb.update_layout(showlegend=False)
                    st.plotly_chart(fig_att_mhb, use_container_width=True, key="fig_att_mhb_tab3")
                else:
                    st.info("В ЦБО нет данных")
        else:
            st.warning("В файле не найдены колонки для определения региона или подразделения (типы, регион, подразделение, отдел, все подразделения).")
            
    else:
        st.warning("Файл по аттестации не загружен. Добавьте его в боковой панели (слот 3).")

# === ВКЛАДКА 4: ОБУЧЕНИЕ ===
with tab4:
    if df_training is not None:
        with st.form("training_filters"):
            safe_dates_tr = df_training['Дата прохождения'].dropna()
            min_date_tr = safe_dates_tr.min().date() if not safe_dates_tr.empty else pd.to_datetime('today').date()
            max_date_tr = safe_dates_tr.max().date() if not safe_dates_tr.empty else pd.to_datetime('today').date()
            date_range_tr = st.date_input("🗓️ Период обучения", [min_date_tr, max_date_tr], min_value=min_date_tr, max_value=max_date_tr)
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                deps = df_training['Отдел'].dropna().unique() if 'Отдел' in df_training.columns else []
                selected_deps = st.multiselect("Выберите филиалы / отделы", deps)
            with col2:
                crs = df_training['Название курса'].dropna().unique() if 'Название курса' in df_training.columns else []
                selected_courses = st.multiselect("Выберите курсы", crs)
            submit_tr = st.form_submit_button("Применить фильтры ⚡")

        filtered_tr = df_training.copy()
        if isinstance(date_range_tr, tuple) and len(date_range_tr) == 2:
            filtered_tr = filtered_tr[(filtered_tr['Дата прохождения'].dt.date >= date_range_tr[0]) & (filtered_tr['Дата прохождения'].dt.date <= date_range_tr[1])]

        if selected_deps: filtered_tr = filtered_tr[filtered_tr['Отдел'].isin(selected_deps)]
        if selected_courses: filtered_tr = filtered_tr[filtered_tr['Название курса'].isin(selected_courses)]

        # --- РАСЧЕТ ПЛАНА И ФАКТА ПРИСУТСТВИЯ ---
        if 'План' in filtered_tr.columns and 'Факт' in filtered_tr.columns:
            filtered_tr['План_calc'] = pd.to_numeric(filtered_tr['План'], errors='coerce').fillna(0)
            filtered_tr['Факт_calc'] = pd.to_numeric(filtered_tr['Факт'], errors='coerce').fillna(0)
        else:
            filtered_tr['План_calc'] = 1
            if 'Присутствие' in filtered_tr.columns:
                filtered_tr['Факт_calc'] = filtered_tr['Присутствие'].astype(str).str.lower().isin(['да', '+', '1', 'присутствовал', 'был', 'true']).astype(int)
            elif 'Статус' in filtered_tr.columns:
                filtered_tr['Факт_calc'] = filtered_tr['Статус'].astype(str).str.lower().isin(['завершил', 'прошел', 'присутствовал', 'обучен', 'успешно']).astype(int)
            else:
                has_hours = pd.to_numeric(filtered_tr['Часы обучения'], errors='coerce').fillna(0) > 0
                has_score = filtered_tr['Результат теста (%)'].notna()
                filtered_tr['Факт_calc'] = (has_hours | has_score).astype(int)

        plan_total = filtered_tr['План_calc'].sum()
        fact_total = filtered_tr['Факт_calc'].sum()
        attendance_rate = (fact_total / plan_total * 100) if plan_total > 0 else 0

        total_hours = filtered_tr['Часы обучения'].sum() if 'Часы обучения' in filtered_tr.columns else 0
        avg_score = filtered_tr['Результат теста (%)'].mean() if 'Результат теста (%)' in filtered_tr.columns else np.nan

        t1, t2, t3, t4, t5 = st.columns(5)
        t1.metric("Плановое количество участников", f"{int(plan_total)} чел.")
        t2.metric("Приняли участие", f"{int(fact_total)} чел.")
        t3.metric("Процент явки", f"{attendance_rate:.1f}%")
        t4.metric("Общее количество часов:", f"{int(total_hours)} ч.")
        t5.metric("Ср. балл по тесту", f"{avg_score:.1f}%" if not pd.isna(avg_score) else "Нет тестов")

        st.divider()

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Плановая и фактическая вовлеченность (по курсам)</b></div><br>", unsafe_allow_html=True)
        if 'Название курса' in filtered_tr.columns:
            filtered_tr['Название курса_str'] = filtered_tr['Название курса'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            course_pf = filtered_tr.groupby('Название курса_str').agg(План=('План_calc', 'sum'), Факт=('Факт_calc', 'sum')).reset_index()
            course_pf = course_pf.rename(columns={'Название курса_str': 'Название курса'})
            course_pf = course_pf[course_pf['План'] > 0].sort_values('План', ascending=True).tail(15)
            course_pf_melted = course_pf.melt(id_vars='Название курса', value_vars=['План', 'Факт'], var_name='Показатель', value_name='Количество')
            
            fig_pf = px.bar(course_pf_melted, x='Количество', y='Название курса', color='Показатель', barmode='group',
                            text='Количество', color_discrete_map={'План': '#ff7f0e', 'Факт': '#2ca02c'})
            fig_pf.update_traces(textposition='outside')
            fig_pf.update_layout(yaxis={'categoryorder':'total ascending'}, legend_title_text='Легенда')
            st.plotly_chart(fig_pf, use_container_width=True, key="fig_pf_tab4")

    else:
        st.warning("Файл по обучению не загружен.")

# === ВКЛАДКА 5: ОБЩИЙ ШТАТ ===
with tab5:
    if df_staff is not None:
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Аналитика: общая численность сотрудников</b></div><br>", unsafe_allow_html=True)
        
        total_staff = len(df_staff)
        st.metric("Общая численность сотрудников", f"{total_staff} чел.")
        st.divider()

        # --- СТРУКТУРА ШТАТА ---
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Распределение штата:Филиалы/ЦБО/ГО</b></div><br>", unsafe_allow_html=True)
        
        cols_lower = {c.lower(): c for c in df_staff.columns}
        col_region = cols_lower.get('типы') or cols_lower.get('регион')
        col_subdiv = cols_lower.get('все подразделения') or cols_lower.get('подразделение') or cols_lower.get('отдел') or cols_lower.get('филиал')

        if col_subdiv or col_region:
            df_staff_filtered = df_staff.copy()
            
            # Умное определение Региона
            if col_region:
                df_staff_filtered['Типы_str'] = df_staff_filtered[col_region].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            elif col_subdiv:
                df_staff_filtered['Типы_str'] = df_staff_filtered[col_subdiv].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            else:
                df_staff_filtered['Типы_str'] = 'Не указано'
            
            df_staff_filtered['Регион_calc'] = df_staff_filtered['Типы_str'].apply(categorize_region)
            
            # Умное определение Подразделения
            if col_subdiv:
                df_staff_filtered['Подразделение_str'] = df_staff_filtered[col_subdiv].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            else:
                df_staff_filtered['Подразделение_str'] = df_staff_filtered['Типы_str']

            st_r1, st_r2, st_r3 = st.columns(3)

            with st_r1:
                st.markdown("**1. Общая численность сотрудников по подразделениям: Филиалы/ЦБО/ГО**")
                staff_reg_data = df_staff_filtered['Регион_calc'].value_counts().reset_index()
                staff_reg_data.columns = ['Регион', 'Количество']
                staff_reg_data = staff_reg_data[staff_reg_data['Количество'] > 0]
                if not staff_reg_data.empty:
                    fig_staff_reg = px.pie(staff_reg_data, names='Регион', values='Количество', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                    fig_staff_reg.update_traces(textposition='auto', textinfo='label+percent')
                    fig_staff_reg.update_layout(showlegend=False)
                    st.plotly_chart(fig_staff_reg, use_container_width=True, key="fig_staff_reg_tab5")
                else:
                    st.info("Нет данных")

            with st_r2:
                st.markdown("**2. Данные по Филиалам**")
                staff_branches = df_staff_filtered[df_staff_filtered['Регион_calc'] == 'Филиал']
                staff_branch_data = staff_branches['Подразделение_str'].value_counts().reset_index()
                staff_branch_data.columns = ['Филиал', 'Количество']
                staff_branch_data = staff_branch_data[staff_branch_data['Количество'] > 0]
                if not staff_branch_data.empty:
                    fig_staff_branch = px.pie(staff_branch_data, names='Филиал', values='Количество', hole=0.4)
                    fig_staff_branch.update_traces(textposition='auto', textinfo='label+value')
                    fig_staff_branch.update_layout(showlegend=False)
                    st.plotly_chart(fig_staff_branch, use_container_width=True, key="fig_staff_branch_tab5")
                else:
                    st.info("В филиалах нет данных")

            with st_r3:
                st.markdown("**3. Данные по ЦБО**")
                staff_mhb = df_staff_filtered[df_staff_filtered['Регион_calc'] == 'ЦБО']
                staff_mhb_data = staff_mhb['Подразделение_str'].value_counts().reset_index()
                staff_mhb_data.columns = ['ЦБО', 'Количество']
                staff_mhb_data = staff_mhb_data[staff_mhb_data['Количество'] > 0]
                if not staff_mhb_data.empty:
                    fig_staff_mhb = px.pie(staff_mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                    fig_staff_mhb.update_traces(textposition='auto', textinfo='label+value')
                    fig_staff_mhb.update_layout(showlegend=False)
                    st.plotly_chart(fig_staff_mhb, use_container_width=True, key="fig_staff_mhb_tab5")
                else:
                    st.info("В ЦБО нет данных")
        else:
            st.warning("Не удалось найти колонку с названиями подразделений (искали: подразделение, отдел, филиал, типы). Срез построить невозможно.")

        st.divider()

        # --- ДИНАМИКА ---
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Численность персонала: по месяцам</b></div><br>", unsafe_allow_html=True)
        
        date_cols = [c for c in df_staff.columns if 'дат' in c.lower() or 'период' in c.lower() or 'месяц' in c.lower()]
        
        if date_cols:
            target_date_col = date_cols[0]
            st.markdown(f"**График построен на основе колонки: `{target_date_col}`**")
            
            df_staff[target_date_col] = pd.to_datetime(df_staff[target_date_col], errors='coerce')
            trend_staff = df_staff.dropna(subset=[target_date_col]).copy()
            
            if not trend_staff.empty:
                trend_staff['Месяц_Год'] = trend_staff[target_date_col].dt.strftime('%Y-%m')
                staff_counts = trend_staff['Месяц_Год'].value_counts().reset_index().sort_values('Месяц_Год')
                staff_counts.columns = ['Период', 'Количество записей (сотрудников)']
                staff_counts['Период'] = staff_counts['Период'].astype(str)
                
                fig_staff_trend = px.line(staff_counts, x='Период', y='Количество записей (сотрудников)', markers=True, color_discrete_sequence=['#2ca02c'])
                fig_staff_trend.update_xaxes(type='category')
                fig_staff_trend.update_layout(xaxis_title="Месяц", yaxis_title="Численность", hovermode="x unified")
                st.plotly_chart(fig_staff_trend, use_container_width=True, key="fig_staff_trend_tab5")
            else:
                st.info(f"Не удалось распознать даты в колонке {target_date_col}.")
        else:
            st.warning("Не найдена колонка с датой (искал слова: дата, период, месяц) для построения графика по месяцам.")
    else:
        st.warning("Файл по общему штату/сотрудникам не загружен. Добавьте его в боковой панели (слот 4).")
