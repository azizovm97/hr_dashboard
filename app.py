import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# --- Настройка страницы ---
st.set_page_config(page_title="HR Analytics Dashboard", layout="wide")
st.title("HR Дашборд")

# --- Боковая панель для загрузки ---
with st.sidebar:
    st.header("📂 Загрузка оригинальных файлов Excel")
    st.markdown("Перетащите сюда ваши файлы из Excel напрямую.")
    recruitment_file = st.file_uploader("1. Файл по найму (Приём)", type=['xlsx'])
    training_file = st.file_uploader("2. Файл по обучению", type=['xlsx'])

# --- Вспомогательные функции ---
def get_header_row_index(file_bytes, keywords):
    """Принимает bytes, возвращает индекс строки-заголовка."""
    df_preview = pd.read_excel(BytesIO(file_bytes), header=None, nrows=15)
    for idx, row in df_preview.iterrows():
        # Явное приведение каждой ячейки к str, None → ''
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
# Принимаем bytes, а не файловый объект — bytes сериализуются корректно
@st.cache_data(show_spinner=False)
def load_recruitment_data(file_bytes: bytes):
    header_idx = get_header_row_index(file_bytes, ['фио кандидата', 'вакансия'])
    df = pd.read_excel(BytesIO(file_bytes), header=header_idx)
    df.columns = df.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()

    if 'ФИО кандидата' in df.columns:
        df = df.dropna(subset=['ФИО кандидата'])

    cols_to_keep = ['Дата регистрации', 'Вакансия', 'Типы', 'Все подразделения',
                    'Тест', 'Отправлен за документами', 'Принят на работу',
                    'Источник подбора', 'Статус']
    existing_cols = [col for col in cols_to_keep if col in df.columns]
    df = df[existing_cols]

    cat_cols = ['Вакансия', 'Типы', 'Все подразделения', 'Отправлен за документами',
                'Принят на работу', 'Источник подбора', 'Статус']
    for col in cat_cols:
        if col in df.columns:
            df[col] = df[col].astype('category')

    if 'Дата регистрации' in df.columns:
        df['Дата регистрации'] = pd.to_datetime(df['Дата регистрации'], errors='coerce')

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

    cols_to_keep = ['Дата прохождения', 'Отдел', 'Должность',
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

# Читаем bytes один раз, передаём в кэшируемые функции
if recruitment_file is not None:
    recruitment_bytes = recruitment_file.read()
    df_recruitment = load_recruitment_data(recruitment_bytes)

if training_file is not None:
    training_bytes = training_file.read()
    df_training = load_training_data(training_bytes)

if recruitment_file is None and training_file is None:
    st.info("Загрузите файл Excel в боковой панели слева для начала работы.")
    st.stop()

# --- Вкладки ---
tab1, tab2 = st.tabs(["Аналитика найма", "Корпоративное обучение"])

# === ВКЛАДКА 1: НАЙМ ===
with tab1:
    st.header("Аналитика воронки подбора")
    if df_recruitment is not None:
        with st.form("recruitment_filters"):
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

        filtered_rec = df_recruitment.copy()
        if selected_vacancies:
            filtered_rec = filtered_rec[filtered_rec['Вакансия'].isin(selected_vacancies)]
        if selected_types:
            filtered_rec = filtered_rec[filtered_rec['Типы'].isin(selected_types)]
        if selected_sources:
            filtered_rec = filtered_rec[filtered_rec['Источник подбора'].isin(selected_sources)]

        total_candidates = len(filtered_rec)
        hired = len(filtered_rec[filtered_rec['Принят на работу'].isin(['Штат', 'Стажёр'])]) if 'Принят на работу' in filtered_rec.columns else 0
        conv_rate = (hired / total_candidates * 100) if total_candidates > 0 else 0

        m1, m2, m3 = st.columns(3)
        m1.metric("Всего кандидатов в воронке", f"{total_candidates} чел.")
        m2.metric("Успешно трудоустроено", f"{hired} чел.")
        m3.metric("Конверсия в найм", f"{conv_rate:.1f}%")

        st.divider()

        g1, g2 = st.columns(2)
        with g1:
            st.subheader("Источники подбора (Количество)")
            if 'Источник подбора' in filtered_rec.columns:
                clean_sources = filtered_rec[filtered_rec['Источник подбора'].notna()]['Источник подбора']
                clean_sources = clean_sources.astype(str).str.strip()
                source_data = clean_sources.value_counts().reset_index()
                source_data.columns = ['Источник', 'Количество кандидатов']
                source_data = source_data[source_data['Количество кандидатов'] > 0]
                fig_source = px.pie(source_data, names='Источник', values='Количество кандидатов',
                                    hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_source.update_traces(textposition='inside', textinfo='value+label')
                st.plotly_chart(fig_source, use_container_width=True)

        with g2:
            st.subheader("Успешность тестов по вакансиям")
            if 'Успешность_теста_%' in filtered_rec.columns and 'Вакансия' in filtered_rec.columns:
                score_data = filtered_rec.dropna(subset=['Успешность_теста_%'])
                score_data = score_data.copy()
                score_data['Вакансия'] = score_data['Вакансия'].astype(str)
                avg_scores = score_data.groupby('Вакансия')['Успешность_теста_%'].mean().reset_index()
                avg_scores = avg_scores.sort_values(by='Успешность_теста_%', ascending=True).tail(10)
                fig_scores = px.bar(avg_scores, x='Успешность_теста_%', y='Вакансия', orientation='h',
                                    color='Успешность_теста_%', color_continuous_scale='Blues')
                st.plotly_chart(fig_scores, use_container_width=True)
    else:
        st.warning("Файл по найму не загружен. Добавьте его в боковой панели слева.")

# === ВКЛАДКА 2: ОБУЧЕНИЕ ===
with tab2:
    st.header("Мониторинг корпоративного обучения")
    if df_training is not None:
        with st.form("training_filters"):
            col1, col2 = st.columns(2)
            with col1:
                departments = df_training['Отдел'].cat.categories if 'Отдел' in df_training.columns else []
                selected_deps = st.multiselect("Выберите филиалы / отделы", departments)
            with col2:
                courses = df_training['Название курса'].cat.categories if 'Название курса' in df_training.columns else []
                selected_courses = st.multiselect("Выберите курсы", courses)
            submit_tr = st.form_submit_button("Применить фильтры ⚡")

        filtered_tr = df_training.copy()
        if selected_deps:
            filtered_tr = filtered_tr[filtered_tr['Отдел'].isin(selected_deps)]
        if selected_courses:
            filtered_tr = filtered_tr[filtered_tr['Название курса'].isin(selected_courses)]

        total_trained = len(filtered_tr)
        total_hours = filtered_tr['Часы обучения'].sum() if 'Часы обучения' in filtered_tr.columns else 0
        avg_score = filtered_tr['Результат теста (%)'].mean() if 'Результат теста (%)' in filtered_tr.columns else np.nan

        t1, t2, t3 = st.columns(3)
        t1.metric("Всего обучено сотрудников", f"{total_trained} чел.")
        t2.metric("Суммарно часов обучения", f"{int(total_hours)} ч.")
        t3.metric("Средний результат теста", f"{avg_score:.1f}%" if not pd.isna(avg_score) else "Нет тестов")

        st.divider()

        t_g1, t_g2 = st.columns(2)
        with t_g1:
            st.subheader("Популярность образовательных курсов")
            if 'Название курса' in filtered_tr.columns:
                course_data = filtered_tr[filtered_tr['Название курса'].notna()]['Название курса'].astype(str).value_counts().reset_index()
                course_data.columns = ['Курс', 'Обучений']
                course_data = course_data[course_data['Обучений'] > 0]
                fig_courses = px.treemap(course_data, path=['Курс'], values='Обучений',
                                         color='Обучений', color_continuous_scale='Teal')
                st.plotly_chart(fig_courses, use_container_width=True)

        with t_g2:
            st.subheader("Нагрузка по отделам (часы)")
            if 'Отдел' in filtered_tr.columns and 'Часы обучения' in filtered_tr.columns:
                filtered_tr = filtered_tr.copy()
                filtered_tr['Отдел_str'] = filtered_tr['Отдел'].astype(str)
                dept_hours = filtered_tr.groupby('Отдел_str')['Часы обучения'].sum().reset_index()
                dept_hours = dept_hours[dept_hours['Часы обучения'] > 0].sort_values(by='Часы обучения', ascending=False)
                fig_hours = px.bar(dept_hours, x='Отдел_str', y='Часы обучения',
                                   color='Отдел_str', text_auto=True)
                fig_hours.update_layout(xaxis_title="Отдел", showlegend=False)
                st.plotly_chart(fig_hours, use_container_width=True)
    else:
        st.warning("Файл по обучению не загружен. Добавьте его в боковой панели слева.")
