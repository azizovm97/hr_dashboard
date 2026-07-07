import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
from datetime import datetime

# --- Настройка страницы ---
st.set_page_config(page_title="HR Analytics Dashboard", layout="wide")

# --- Главное название внутри красной рамки ---
st.markdown(
    """
    <div style="border: 3px solid red; padding: 15px; border-radius: 10px; text-align: center; margin-bottom: 25px;">
        <h1 style="margin: 0; padding: 0; font-size: 2.5rem;">HR Дашборд</h1>
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

# --- Вспомогательные функции ---
def get_header_row_index(file_bytes, keywords):
    """Принимает bytes, возвращает индекс строки-заголовка."""
    df_preview = pd.read_excel(BytesIO(file_bytes), header=None, nrows=15)
    for idx, row in df_preview.iterrows():
        row_str = [str(v).lower().strip() if v is not None and not (isinstance(v, float) and np.isnan(v)) else '' for v in row]
        if any(any(kw in val for kw in keywords) for val in row_str):
            return idx
    return 0

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

# --- Загрузка и кэширование данных ---
@st.cache_data(show_spinner=False)
def load_recruitment_data(file_bytes: bytes):
    header_idx = get_header_row_index(file_bytes, ['фио кандидата', 'вакансия'])
    df = pd.read_excel(BytesIO(file_bytes), header=header_idx)
    df.columns = df.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()

    if 'ФИО кандидата' in df.columns:
        df = df.dropna(subset=['ФИО кандидата'])

    cols_to_keep = ['Дата регистрации', 'Пол', 'Возраст', 'Вакансия', 'Типы', 'Все подразделения',
                    'Тест', 'Отправлен за документами', 'Принят на работу',
                    'Источник подбора', 'Статус']
    existing_cols = [col for col in cols_to_keep if col in df.columns]
    df = df[existing_cols]

    cat_cols = ['Пол', 'Вакансия', 'Типы', 'Все подразделения', 'Отправлен за документами',
                'Принят на работу', 'Источник подбора', 'Статус']
    for col in cat_cols:
        if col in df.columns:
            df[col] = df[col].astype('category')

    if 'Дата регистрации' in df.columns:
        df['Дата регистрации'] = pd.to_datetime(df['Дата регистрации'], errors='coerce')
        
    if 'Возраст' in df.columns:
        df['Возраст'] = pd.to_numeric(df['Возраст'], errors='coerce')

    if 'Тест' in df.columns:
        df['Успешность_теста_%'] = df['Тест'].apply(parse_test_score).astype(np.float32)
    return df

@st.cache_data(show_spinner=False)
def load_training_data(file_bytes: bytes):
    header_idx = get_header_row_index(file_bytes, ['фио сотрудника', 'название курса'])
    df = pd.read_excel(BytesIO(file_bytes), header=header_idx)
    df.columns = df.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()

    if 'ФИО сотрудника' in df.columns:
        df = df.dropna(subset=['ФИО сотрудника'])

    cols_to_keep = ['ФИО сотрудника', 'Дата прохождения', 'Отдел', 'Должность',
                    'Название курса', 'Часы обучения', 'Результат теста (%)']
    existing_cols = [col for col in cols_to_keep if col in df.columns]
    df = df[existing_cols]

    cat_cols = ['Отдел', 'Должность', 'Название курса']
    for col in cat_cols:
        if col in df.columns:
            df[col] = df[col].astype('category')

    if 'Дата прохождения' in df.columns:
        df['Дата прохождения'] = pd.to_datetime(df['Дата прохождения'], errors='coerce')

    if 'Часы обучения' in df.columns:
        df['Часы обучения'] = pd.to_numeric(df['Часы обучения'], errors='coerce', downcast='unsigned')
    return df

# --- Инициализация переменных ---
df_recruitment = None
df_training = None

if recruitment_file is not None:
    recruitment_bytes = recruitment_file.read()
    df_recruitment = load_recruitment_data(recruitment_bytes)

if training_file is not None:
    training_bytes = training_file.read()
    df_training = load_training_data(training_bytes)

if recruitment_file is None and training_file is None:
    st.info("Загрузите файлы Excel в боковой панели слева для начала работы.")
    st.stop()

# --- Вкладки ---
tab1, tab2 = st.tabs(["Аналитика найма", "Корпоративное обучение"])

# === ВКЛАДКА 1: НАЙМ ===
with tab1:
    if df_recruitment is not None:
        with st.form("recruitment_filters"):
            # ФИЛЬТР ПО ДАТАМ
            safe_dates_rec = df_recruitment['Дата регистрации'].dropna()
            if not safe_dates_rec.empty:
                min_date_rec = safe_dates_rec.min().date()
                max_date_rec = safe_dates_rec.max().date()
            else:
                min_date_rec = pd.to_datetime('2020-01-01').date()
                max_date_rec = pd.to_datetime('today').date()
                
            date_range_rec = st.date_input("Период регистрации", [min_date_rec, max_date_rec], min_value=min_date_rec, max_value=max_date_rec)
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                vacancies = df_recruitment['Вакансия'].cat.categories if 'Вакансия' in df_recruitment.columns else []
                selected_vacancies = st.multiselect("Выберите вакансии", vacancies)
            with col2:
                types = df_recruitment['Типы'].cat.categories if 'Типы' in df_recruitment.columns else []
                selected_types = st.multiselect("Тип подразделения", types)
            with col3:
                sources = df_recruitment['Источник подбора'].cat.categories if 'Источник подбора' in df_recruitment.columns else []
                selected_sources = st.multiselect("Источник привлечения", sources)
            submit = st.form_submit_button("Применить фильтры ⚡")

        # ПРИМЕНЕНИЕ ФИЛЬТРОВ
        filtered_rec = df_recruitment.copy()
        
        if isinstance(date_range_rec, tuple) and len(date_range_rec) == 2:
            start_date, end_date = date_range_rec
            filtered_rec = filtered_rec[
                (filtered_rec['Дата регистрации'].dt.date >= start_date) & 
                (filtered_rec['Дата регистрации'].dt.date <= end_date)
            ]

        if selected_vacancies:
            filtered_rec = filtered_rec[filtered_rec['Вакансия'].isin(selected_vacancies)]
        if selected_types:
            filtered_rec = filtered_rec[filtered_rec['Типы'].isin(selected_types)]
        if selected_sources:
            filtered_rec = filtered_rec[filtered_rec['Источник подбора'].isin(selected_sources)]

        # --- Базовые Метрики ---
        total_candidates = len(filtered_rec)
        hired_mask = filtered_rec['Принят на работу'].isin(['Штат', 'Стажёр', 'Волонтёр']) if 'Принят на работу' in filtered_rec.columns else pd.Series(False, index=filtered_rec.index)
        hired = hired_mask.sum()
        conv_rate = (hired / total_candidates * 100) if total_candidates > 0 else 0
        avg_age = filtered_rec[hired_mask]['Возраст'].mean() if 'Возраст' in filtered_rec.columns else np.nan
        
        if 'Пол' in filtered_rec.columns:
            hired_gender = filtered_rec[hired_mask]['Пол'].astype(str).str.strip().str.upper()
            m_count = hired_gender.str.startswith('М').sum()
            f_count = hired_gender.str.startswith('Ж').sum()
            gender_text = f"М: {m_count} | Ж: {f_count}"
        else:
            gender_text = "Нет данных"

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Всего кандидатов", f"{total_candidates} чел.")
        m2.metric("Успешно трудоустроено", f"{hired} чел.")
        m3.metric("Конверсия в найм", f"{conv_rate:.1f}%")
        m4.metric("Ср. возраст (найм)", f"{avg_age:.1f} лет" if not pd.isna(avg_age) else "Нет данных")
        m5.metric("М / Ж (найм)", gender_text)

        st.divider()

        # --- Источники подбора: Объем и Конверсия ---
        f_col1, f_col2 = st.columns(2)
        
        with f_col1:
            st.subheader("Источники подбора (Общий объем трафика)")
            if 'Источник подбора' in filtered_rec.columns:
                src_traffic = filtered_rec['Источник подбора'].astype(str).str.strip()
                src_traffic = src_traffic.replace({'nan': 'Не указано', '': 'Не указано'})
                
                src_traffic = src_traffic.value_counts().reset_index()
                src_traffic.columns = ['Источник', 'Всего кандидатов']
                
                fig_source_vol = px.pie(src_traffic, names='Источник', values='Всего кандидатов',
                                        hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                # ИЗМЕНЕНИЕ: Теперь показываем "Название источника" и "Количество"
                fig_source_vol.update_traces(textposition='inside', textinfo='label+value')
                st.plotly_chart(fig_source_vol, use_container_width=True)

        with f_col2:
            st.subheader("Качество источников (Конверсия в найм)")
            if 'Источник подбора' in filtered_rec.columns:
                src_df = filtered_rec.copy()
                src_df['Источник подбора'] = src_df['Источник подбора'].astype(str).str.strip()
                src_df['Источник подбора'] = src_df['Источник подбора'].replace({'nan': 'Не указано', '': 'Не указано'})
                
                src_df['Is_Hired'] = hired_mask.astype(int)
                
                src_conv = src_df.groupby('Источник подбора').agg(
                    Total=('Is_Hired', 'count'), 
                    Hired=('Is_Hired', 'sum')
                ).reset_index()
                
                src_conv['Конверсия %'] = (src_conv['Hired'] / src_conv['Total'] * 100).round(1)
                src_conv = src_conv[src_conv['Total'] > 0].sort_values('Конверсия %', ascending=True)
                
                src_conv['Текст на графике'] = src_conv['Конверсия %'].astype(str) + '% (' + src_conv['Hired'].astype(str) + ' из ' + src_conv['Total'].astype(str) + ')'
                
                fig_src_conv = px.bar(src_conv, x='Конверсия %', y='Источник подбора', orientation='h',
                                      text='Текст на графике', hover_data=['Total', 'Hired'],
                                      color='Конверсия %', color_continuous_scale='Greens')
                fig_src_conv.update_traces(textposition='outside')
                max_x = src_conv['Конверсия %'].max()
                fig_src_conv.update_layout(xaxis_range=[0, max_x * 1.3 if max_x > 0 else 100])
                st.plotly_chart(fig_src_conv, use_container_width=True)

        st.divider()
        
        # --- Структура Кандидатов: Филиалы, ЦБО, HO ---
        st.subheader("Структура кандидатов: Филиалы, МХБ/ЦБО и HO")
        r1, r2, r3 = st.columns(3)

        if 'Типы' in filtered_rec.columns and 'Все подразделения' in filtered_rec.columns:
            filtered_rec_copy = filtered_rec.copy()
            
            filtered_rec_copy['Типы_str'] = filtered_rec_copy['Типы'].astype(str).str.strip()
            filtered_rec_copy['Типы_str'] = filtered_rec_copy['Типы_str'].replace({'nan': 'Не указано', '': 'Не указано'})
            
            filtered_rec_copy['Подразделение_str'] = filtered_rec_copy['Все подразделения'].astype(str).str.strip()
            filtered_rec_copy['Подразделение_str'] = filtered_rec_copy['Подразделение_str'].replace({'nan': 'Не указано', '': 'Не указано'})

            with r1:
                st.markdown("**1. Общее распределение (по Типам)**")
                types_data = filtered_rec_copy['Типы_str'].value_counts().reset_index()
                types_data.columns = ['Тип', 'Количество']
                
                if not types_data.empty:
                    fig_types = px.pie(types_data, names='Тип', values='Количество', hole=0.4)
                    fig_types.update_traces(textposition='inside', textinfo='value+percent') 
                    fig_types.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.2))
                    st.plotly_chart(fig_types, use_container_width=True)
                else:
                    st.info("Нет данных")

            with r2:
                st.markdown("**2. Внутри Филиалов**")
                branches_df = filtered_rec_copy[filtered_rec_copy['Типы_str'].str.contains('Филиал|филиал', na=False)]
                branch_data = branches_df['Подразделение_str'].value_counts().reset_index()
                branch_data.columns = ['Филиал', 'Количество']
                
                if not branch_data.empty:
                    fig_branch = px.pie(branch_data, names='Филиал', values='Количество', hole=0.4)
                    fig_branch.update_traces(textposition='inside', textinfo='value+percent')
                    fig_branch.update_layout(showlegend=False)
                    st.plotly_chart(fig_branch, use_container_width=True)
                else:
                    st.info("Нет данных по Филиалам")

            with r3:
                st.markdown("**3. Внутри МХБ / ЦБО**")
                mhb_df = filtered_rec_copy[filtered_rec_copy['Типы_str'].str.contains('МХБ|ЦБО|мхб|цбо', na=False)]
                mhb_data = mhb_df['Подразделение_str'].value_counts().reset_index()
                mhb_data.columns = ['МХБ/ЦБО', 'Количество']
                
                if not mhb_data.empty:
                    fig_mhb = px.pie(mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                    fig_mhb.update_traces(textposition='inside', textinfo='value+percent')
                    fig_mhb.update_layout(showlegend=False)
                    st.plotly_chart(fig_mhb, use_container_width=True)
                else:
                    st.info("Нет данных по МХБ/ЦБО")
                    
        st.divider()
        st.subheader("Динамика регистраций кандидатов")
        if 'Дата регистрации' in filtered_rec.columns:
            trend_data = filtered_rec.dropna(subset=['Дата регистрации']).copy()
            trend_data['Период'] = trend_data['Дата регистрации'].dt.to_period('M')
            trend_counts = trend_data['Период'].value_counts().reset_index()
            trend_counts.columns = ['Период', 'Кандидаты']
            trend_counts = trend_counts.sort_values('Период')
            trend_counts['Период'] = trend_counts['Период'].astype(str) 
            
            fig_trend = px.line(trend_counts, x='Период', y='Кандидаты', markers=True, color_discrete_sequence=['#ff7f0e'])
            st.plotly_chart(fig_trend, use_container_width=True)

    else:
        st.warning("Файл по найму не загружен. Добавьте его в боковой панели слева.")

# === ВКЛАДКА 2: ОБУЧЕНИЕ ===
with tab2:
    if df_training is not None:
        with st.form("training_filters"):
            # ФИЛЬТР ПО ДАТАМ (Обучение)
            safe_dates_tr = df_training['Дата прохождения'].dropna()
            if not safe_dates_tr.empty:
                min_date_tr = safe_dates_tr.min().date()
                max_date_tr = safe_dates_tr.max().date()
            else:
                min_date_tr = pd.to_datetime('2020-01-01').date()
                max_date_tr = pd.to_datetime('today').date()
                
            date_range_tr = st.date_input("🗓️ Период обучения", [min_date_tr, max_date_tr], min_value=min_date_tr, max_value=max_date_tr)
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                departments = df_training['Отдел'].cat.categories if 'Отдел' in df_training.columns else []
                selected_deps = st.multiselect("Выберите филиалы / отделы", departments)
            with col2:
                courses = df_training['Название курса'].cat.categories if 'Название курса' in df_training.columns else []
                selected_courses = st.multiselect("Выберите курсы", courses)
            submit_tr = st.form_submit_button("Применить фильтры ⚡")

        filtered_tr = df_training.copy()
        
        if isinstance(date_range_tr, tuple) and len(date_range_tr) == 2:
            start_date_tr, end_date_tr = date_range_tr
            filtered_tr = filtered_tr[
                (filtered_tr['Дата прохождения'].dt.date >= start_date_tr) & 
                (filtered_tr['Дата прохождения'].dt.date <= end_date_tr)
            ]

        if selected_deps:
            filtered_tr = filtered_tr[filtered_tr['Отдел'].isin(selected_deps)]
        if selected_courses:
            filtered_tr = filtered_tr[filtered_tr['Название курса'].isin(selected_courses)]

        if 'ФИО сотрудника' in filtered_tr.columns:
            unique_trained = filtered_tr['ФИО сотрудника'].nunique()
        else:
            unique_trained = len(filtered_tr)
            
        total_courses_completed = len(filtered_tr)
        total_hours = filtered_tr['Часы обучения'].sum() if 'Часы обучения' in filtered_tr.columns else 0
        avg_score = filtered_tr['Результат теста (%)'].mean() if 'Результат теста (%)' in filtered_tr.columns else np.nan

        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Уникальных сотрудников", f"{unique_trained} чел.")
        t2.metric("Пройдено курсов (кол-во)", f"{total_courses_completed}")
        t3.metric("Суммарно часов обучения", f"{int(total_hours)} ч.")
        t4.metric("Средний результат теста", f"{avg_score:.1f}%" if not pd.isna(avg_score) else "Нет тестов")

        st.divider()

        st.subheader("Структура обучения: Распределение по Отделам / Филиалам")
        if 'Отдел' in filtered_tr.columns:
            dept_data = filtered_tr['Отдел'].astype(str).str.strip()
            dept_data = dept_data.replace({'nan': 'Не указано', '': 'Не указано'})
            
            dept_data = dept_data.value_counts().reset_index()
            dept_data.columns = ['Отдел', 'Количество пройденных курсов']
            
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                st.markdown("**Доля обучения по отделам (Топ-15)**")
                fig_dept_pie = px.pie(dept_data.head(15), names='Отдел', values='Количество пройденных курсов', hole=0.4)
                fig_dept_pie.update_traces(textposition='inside', textinfo='percent+label')
                fig_dept_pie.update_layout(showlegend=False)
                st.plotly_chart(fig_dept_pie, use_container_width=True)
                
            with d_col2:
                st.markdown("**Объем обучения по отделам (Все)**")
                dept_data_sorted = dept_data.sort_values('Количество пройденных курсов', ascending=True).tail(15)
                fig_dept_bar = px.bar(dept_data_sorted, x='Количество пройденных курсов', y='Отдел', orientation='h',
                                      text='Количество пройденных курсов', color='Количество пройденных курсов', color_continuous_scale='Purples')
                fig_dept_bar.update_traces(textposition='outside')
                st.plotly_chart(fig_dept_bar, use_container_width=True)

        st.divider()

        t_g1, t_g2 = st.columns(2)
        with t_g1:
            st.subheader("Популярность образовательных курсов")
            if 'Название курса' in filtered_tr.columns:
                course_data = filtered_tr['Название курса'].astype(str).str.strip()
                course_data = course_data.replace({'nan': 'Не указано', '': 'Не указано'})
                
                course_data = course_data.value_counts().reset_index()
                course_data.columns = ['Курс', 'Обучений']
                course_data = course_data[course_data['Обучений'] > 0]
                fig_courses = px.treemap(course_data, path=['Курс'], values='Обучений',
                                         color='Обучений', color_continuous_scale='Teal')
                st.plotly_chart(fig_courses, use_container_width=True)

        with t_g2:
            st.subheader("Успеваемость: Результат теста по Должностям")
            if 'Должность' in filtered_tr.columns and 'Результат теста (%)' in filtered_tr.columns:
                pos_data = filtered_tr.dropna(subset=['Результат теста (%)']).copy()
                pos_data['Должность_str'] = pos_data['Должность'].astype(str).str.strip()
                pos_data['Должность_str'] = pos_data['Должность_str'].replace({'nan': 'Не указано', '': 'Не указано'})
                
                avg_pos = pos_data.groupby('Должность_str')['Результат теста (%)'].mean().reset_index()
                avg_pos = avg_pos.sort_values('Результат теста (%)', ascending=True).tail(10)
                fig_pos = px.bar(avg_pos, x='Результат теста (%)', y='Должность_str', orientation='h', 
                                 text_auto='.1f', color='Результат теста (%)', color_continuous_scale='Greens')
                st.plotly_chart(fig_pos, use_container_width=True)
                
        st.divider()
        st.subheader("Нагрузка по должностям (Затраченные Часы)")
        if 'Должность' in filtered_tr.columns and 'Часы обучения' in filtered_tr.columns:
            hours_data = filtered_tr.copy()
            hours_data['Должность_str'] = hours_data['Должность'].astype(str).str.strip()
            hours_data['Должность_str'] = hours_data['Должность_str'].replace({'nan': 'Не указано', '': 'Не указано'})
            
            total_hours_pos = hours_data.groupby('Должность_str')['Часы обучения'].sum().reset_index()
            total_hours_pos = total_hours_pos[total_hours_pos['Часы обучения'] > 0].sort_values('Часы обучения', ascending=True).tail(15)
            
            fig_hours = px.bar(total_hours_pos, x='Часы обучения', y='Должность_str', orientation='h',
                               text='Часы обучения', color='Часы обучения', color_continuous_scale='Oranges')
            fig_hours.update_traces(textposition='outside')
            st.plotly_chart(fig_hours, use_container_width=True)
            
    else:
        st.warning("Файл по обучению не загружен. Добавьте его в боковой панели слева.")
