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
    st.header("📂 Загрузка файлов Excel")
    st.markdown("Перетащите сюда ваши файлы из Excel.")
    recruitment_file = st.file_uploader("1. Файл по найму (Приём)", type=['xlsx'])
    training_file = st.file_uploader("2. Файл по обучению", type=['xlsx'])
    attestation_file = st.file_uploader("3. Файл по аттестации", type=['xlsx'])
    staff_file = st.file_uploader("4. Файл по штату (Сотрудники)", type=['xlsx'])

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
df_attestation = load_generic_data(st.session_state['att_bytes'], ['фио', 'категория', 'аттестация', 'дарача']) if 'att_bytes' in st.session_state else None
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
        m1.metric("Всего кандидатов", f"{total_candidates} чел.")
        m2.metric("Успешно трудоустроено", f"{hired} чел.")
        m3.metric("Конверсия в найм", f"{conv_rate:.1f}%")
        m4.metric("М / Ж (найм)", gender_text)
        m5.metric("Ср. возраст (найм)", f"{avg_age:.1f} лет" if not pd.isna(avg_age) else "Нет данных")

        st.divider()

        f_col1, f_col2 = st.columns(2)
        with f_col1:
            st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Источники подбора (Общее кол-во заявок)</b></div><br>", unsafe_allow_html=True)
            if 'Источник подбора' in filtered_rec.columns:
                src_traffic = filtered_rec['Источник подбора'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
                src_traffic.columns = ['Источник', 'Всего кандидатов']
                fig_source_vol = px.pie(src_traffic, names='Источник', values='Всего кандидатов', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_source_vol.update_traces(textposition='inside', textinfo='label+value')
                st.plotly_chart(fig_source_vol, use_container_width=True)

        with f_col2:
            st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Источник найма (Конверсия)</b></div><br>", unsafe_allow_html=True)
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
                st.plotly_chart(fig_src_conv, use_container_width=True)

        st.divider()
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Динамика регистраций кандидатов</b></div><br>", unsafe_allow_html=True)
        if 'Дата регистрации' in filtered_rec.columns:
            trend_data = filtered_rec.dropna(subset=['Дата регистрации']).copy()
            trend_data['Период'] = trend_data['Дата регистрации'].dt.to_period('M')
            trend_counts = trend_data['Период'].value_counts().reset_index().sort_values('Период')
            trend_counts.columns = ['Период', 'Кандидаты']
            trend_counts['Период'] = trend_counts['Период'].astype(str) 
            fig_trend = px.line(trend_counts, x='Период', y='Кандидаты', markers=True, color_discrete_sequence=['#ff7f0e'])
            st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.warning("Файл по найму не загружен.")

# === ВКЛАДКА 2: СТАЖЁРЫ (ВЫДЕЛЕННЫЙ ДАШБОРД) ===
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

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Ключевые показатели по стажёрам</b></div><br>", unsafe_allow_html=True)
        
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
        
        # Разметка для диаграмм по стажерам
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Детальная разбивка стажёров (Регионы, Статусы, Должности)</b></div><br>", unsafe_allow_html=True)
        
        int_c1, int_c2 = st.columns(2)
        
        with int_c1:
            st.markdown("**Разбивка по регионам (Филиал / МХБ)**")
            if 'Типы' in filtered_int.columns:
                filtered_int['Регион'] = filtered_int['Типы'].apply(categorize_region)
                reg_counts = filtered_int['Регион'].value_counts().reset_index()
                reg_counts.columns = ['Регион', 'Количество']
                fig_reg_int = px.pie(reg_counts, names='Регион', values='Количество', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                fig_reg_int.update_traces(textposition='inside', textinfo='value+percent')
                st.plotly_chart(fig_reg_int, use_container_width=True)
            else:
                st.info("Нет колонки Типы/Регион")

        with int_c2:
            st.markdown("**Распределение по статусам**")
            if 'Статус' in filtered_int.columns:
                stat_counts = filtered_int['Статус'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
                stat_counts.columns = ['Статус', 'Количество']
                fig_stat_int = px.bar(stat_counts, x='Количество', y='Статус', orientation='h', text='Количество', color='Количество', color_continuous_scale='Blues')
                fig_stat_int.update_traces(textposition='outside')
                fig_stat_int.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig_stat_int, use_container_width=True)
            else:
                st.info("Нет колонки Статус")

        st.markdown("**Распределение стажёров по вакансиям (должностям)**")
        if 'Вакансия' in filtered_int.columns:
            vac_counts = filtered_int['Вакансия'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
            vac_counts.columns = ['Вакансия', 'Количество']
            fig_vac_int = px.bar(vac_counts.head(15), x='Количество', y='Вакансия', orientation='h', text='Количество', color='Количество', color_continuous_scale='Purples')
            fig_vac_int.update_traces(textposition='outside')
            fig_vac_int.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_vac_int, use_container_width=True)

    else:
        st.warning("Нет данных по стажёрам. Загрузите файл по найму.")

# === ВКЛАДКА 3: АТТЕСТАЦИЯ (НОВАЯ) ===
with tab3:
    if df_attestation is not None:
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Аналитика по Аттестации и Категориям (Дарача)</b></div><br>", unsafe_allow_html=True)
        
        # Поиск ключевых колонок в загруженном файле
        cols_lower = {c.lower(): c for c in df_attestation.columns}
        col_status = cols_lower.get('статус')
        col_region = cols_lower.get('типы') or cols_lower.get('регион')
        col_position = cols_lower.get('должность')
        col_category = cols_lower.get('категория') or cols_lower.get('дарача')

        total_attested = len(df_attestation)
        st.metric("📊 Общее количество прошедших аттестацию", f"{total_attested} сотрудников")
        st.divider()

        att_c1, att_c2 = st.columns(2)
        
        with att_c1:
            st.markdown("**Статус прохождения аттестации**")
            if col_status:
                stat_att = df_attestation[col_status].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
                stat_att.columns = ['Статус', 'Количество']
                fig_att_stat = px.pie(stat_att, names='Статус', values='Количество', hole=0.3, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_att_stat.update_traces(textposition='inside', textinfo='value+percent')
                st.plotly_chart(fig_att_stat, use_container_width=True)
            else:
                st.info("Колонка 'Статус' не найдена")

        with att_c2:
            st.markdown("**Регион (Филиал / МХБ)**")
            if col_region:
                df_attestation['Регион_calc'] = df_attestation[col_region].apply(categorize_region)
                reg_att = df_attestation['Регион_calc'].value_counts().reset_index()
                reg_att.columns = ['Регион', 'Количество']
                fig_att_reg = px.pie(reg_att, names='Регион', values='Количество', hole=0.3, color_discrete_sequence=px.colors.qualitative.Set3)
                fig_att_reg.update_traces(textposition='inside', textinfo='value+percent')
                st.plotly_chart(fig_att_reg, use_container_width=True)
            else:
                st.info("Колонка региона (Типы) не найдена")

        st.divider()
        att_c3, att_c4 = st.columns(2)

        with att_c3:
            st.markdown("**Результаты по Должностям (Топ-15)**")
            if col_position:
                pos_att = df_attestation[col_position].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
                pos_att.columns = ['Должность', 'Количество']
                fig_att_pos = px.bar(pos_att.head(15), x='Количество', y='Должность', orientation='h', text='Количество', color='Количество', color_continuous_scale='Teal')
                fig_att_pos.update_traces(textposition='outside')
                fig_att_pos.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig_att_pos, use_container_width=True)
            else:
                st.info("Колонка 'Должность' не найдена")

        with att_c4:
            st.markdown("**Присвоенная категория (Дарача)**")
            if col_category:
                cat_att = df_attestation[col_category].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
                cat_att.columns = ['Категория', 'Количество']
                fig_att_cat = px.bar(cat_att, x='Количество', y='Категория', orientation='h', text='Количество', color='Количество', color_continuous_scale='Oranges')
                fig_att_cat.update_traces(textposition='outside')
                fig_att_cat.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig_att_cat, use_container_width=True)
            else:
                st.info("Колонка 'Категория/Дарача' не найдена")
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

        unique_trained = filtered_tr['ФИО сотрудника'].nunique() if 'ФИО сотрудника' in filtered_tr.columns else len(filtered_tr)
        total_hours = filtered_tr['Часы обучения'].sum() if 'Часы обучения' in filtered_tr.columns else 0
        avg_score = filtered_tr['Результат теста (%)'].mean() if 'Результат теста (%)' in filtered_tr.columns else np.nan

        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Уникальных сотрудников", f"{unique_trained} чел.")
        t2.metric("Пройдено курсов (кол-во)", f"{len(filtered_tr)}")
        t3.metric("Суммарно часов обучения", f"{int(total_hours)} ч.")
        t4.metric("Средний результат теста", f"{avg_score:.1f}%" if not pd.isna(avg_score) else "Нет тестов")

        st.divider()

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Популярность образовательных курсов</b></div><br>", unsafe_allow_html=True)
        if 'Название курса' in filtered_tr.columns:
            course_data = filtered_tr['Название курса'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
            course_data.columns = ['Курс', 'Обучений']
            fig_courses = px.bar(course_data.head(15), x='Обучений', y='Курс', orientation='h', text='Обучений', color='Обучений', color_continuous_scale='Teal')
            fig_courses.update_traces(textposition='outside')
            fig_courses.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_courses, use_container_width=True)

    else:
        st.warning("Файл по обучению не загружен.")

# === ВКЛАДКА 5: ОБЩИЙ ШТАТ (НОВАЯ) ===
with tab5:
    if df_staff is not None:
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Динамика количества сотрудников компании по месяцам</b></div><br>", unsafe_allow_html=True)
        
        total_staff = len(df_staff)
        st.metric("🏢 Общее количество сотрудников (по файлу)", f"{total_staff} чел.")
        st.divider()

        # Автоматический поиск колонки с датой (Дата приема, Дата, Месяц, Период)
        date_cols = [c for c in df_staff.columns if 'дат' in c.lower() or 'период' in c.lower() or 'месяц' in c.lower()]
        
        if date_cols:
            target_date_col = date_cols[0]
            st.markdown(f"**График построен на основе колонки: `{target_date_col}`**")
            
            df_staff[target_date_col] = pd.to_datetime(df_staff[target_date_col], errors='coerce')
            trend_staff = df_staff.dropna(subset=[target_date_col]).copy()
            
            if not trend_staff.empty:
                trend_staff['Месяц_Год'] = trend_staff[target_date_col].dt.to_period('M')
                staff_counts = trend_staff['Месяц_Год'].value_counts().reset_index().sort_values('Месяц_Год')
                staff_counts.columns = ['Период', 'Количество записей (сотрудников)']
                staff_counts['Период'] = staff_counts['Период'].astype(str)
                
                fig_staff_trend = px.line(staff_counts, x='Период', y='Количество записей (сотрудников)', markers=True, color_discrete_sequence=['#2ca02c'])
                fig_staff_trend.update_layout(xaxis_title="Месяц", yaxis_title="Численность", hovermode="x unified")
                st.plotly_chart(fig_staff_trend, use_container_width=True)
            else:
                st.info(f"Не удалось распознать даты в колонке {target_date_col}.")
        else:
            st.warning("Не найдена колонка с датой (искал слова: дата, период, месяц) для построения графика по месяцам.")
    else:
        st.warning("Файл по общему штату/сотрудникам не загружен. Добавьте его в боковой панели (слот 4).")
