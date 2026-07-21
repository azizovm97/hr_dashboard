import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
import os
import json

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

# --- Настройки файлов ---
DEFAULT_FILENAME = "HR_data.xlsx"
USERS_FILE = "users.json"

# --- Работа с базой данных пользователей (JSON) ---
def load_users():
    default_users = {
        "admin": {"password": "admin123", "role": "admin", "name": "Главный Администратор"},
        "viewer1": {"password": "viewer123", "role": "viewer", "name": "Сотрудник (Просмотр)"}
    }
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default_users
    else:
        save_users(default_users)
        return default_users

def save_users(users_dict):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users_dict, f, ensure_ascii=False, indent=4)

if 'users_db' not in st.session_state:
    st.session_state['users_db'] = load_users()

if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False
if 'username' not in st.session_state:
    st.session_state['username'] = ""
if 'role' not in st.session_state:
    st.session_state['role'] = ""
if 'name' not in st.session_state:
    st.session_state['name'] = ""

# --- Боковая панель: Авторизация и Админ-панель ---
with st.sidebar:
    st.header("Авторизация")
    
    if not st.session_state['authenticated']:
        with st.form("login_form"):
            username_input = st.text_input("Логин")
            password_input = st.text_input("Пароль", type="password")
            submit_login = st.form_submit_button("Войти")
            
            if submit_login:
                db = st.session_state['users_db']
                if username_input in db and db[username_input]["password"] == password_input:
                    st.session_state['authenticated'] = True
                    st.session_state['username'] = username_input
                    st.session_state['role'] = db[username_input]["role"]
                    st.session_state['name'] = db[username_input]["name"]
                    st.rerun()
                else:
                    st.error("Неверный логин или пароль ❌")
        st.stop()
    else:
        st.success(f"Вы вошли как:\n**{st.session_state['name']}**")
        if st.button("Выйти"):
            st.session_state['authenticated'] = False
            st.session_state['username'] = ""
            st.session_state['role'] = ""
            st.session_state['name'] = ""
            st.rerun()

    st.divider()

    # --- ПАНЕЛЬ УПРАВЛЕНИЯ ДЛЯ АДМИНА ---
    if st.session_state['role'] == "admin":
        with st.expander("Админ-панель", expanded=False):
            st.subheader("Изменить пароль")
            target_user = st.selectbox("Выберите пользователя", list(st.session_state['users_db'].keys()))
            new_pass = st.text_input("Новый пароль", type="password", key="new_pass_input")
            
            if st.button("Обновить пароль"):
                if new_pass.strip():
                    st.session_state['users_db'][target_user]["password"] = new_pass.strip()
                    save_users(st.session_state['users_db'])
                    st.success(f"Пароль для '{target_user}' изменен!")
                else:
                    st.warning("Введите новый пароль.")
            
            st.divider()
            st.subheader("Новый пользователь")
            new_login = st.text_input("Логин")
            new_name = st.text_input("Имя")
            new_password = st.text_input("Пароль", type="password", key="new_user_pass")
            new_role = st.selectbox("Права", ["viewer", "admin"], format_func=lambda x: "Администратор (загрузка файлов)" if x == "admin" else "Просмотр (только дашборд)")
            
            if st.button("Создать пользователя"):
                if new_login.strip() and new_password.strip():
                    if new_login.strip() in st.session_state['users_db']:
                        st.error("Логин уже занят!")
                    else:
                        st.session_state['users_db'][new_login.strip()] = {
                            "password": new_password.strip(),
                            "role": new_role,
                            "name": new_name.strip() if new_name.strip() else new_login.strip()
                        }
                        save_users(st.session_state['users_db'])
                        st.success(f"Пользователь '{new_login}' создан!")
                        st.rerun()
                else:
                    st.warning("Заполните логин и пароль.")

        st.divider()

    st.header("Управление данными")
    file_bytes = None

    if st.session_state['role'] == 'admin':
        st.markdown("Загрузить новую версию файла:")
        uploaded_file = st.file_uploader("Загрузить единый Excel-файл", type=['xlsx'])
        
        if uploaded_file is not None:
            file_bytes = uploaded_file.getvalue()
            with open(DEFAULT_FILENAME, "wb") as f:
                f.write(file_bytes)
            st.success("Файл успешно обновлен!")
    else:
        st.info("У вас права на просмотр. Загрузка файлов заблокирована.")

    if not file_bytes and os.path.exists(DEFAULT_FILENAME):
        with open(DEFAULT_FILENAME, "rb") as f:
            file_bytes = f.read()

if not file_bytes:
    st.warning(f"Файл с данными (`{DEFAULT_FILENAME}`) не найден. Администратор должен загрузить файл.")
    st.stop()


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

def find_specific_sheet(file_bytes, target_keywords):
    """Ищет лист по ключевым словам в названии"""
    xls = pd.ExcelFile(BytesIO(file_bytes))
    for sheet in xls.sheet_names:
        sheet_lower = sheet.lower().strip()
        for kw in target_keywords:
            if kw in sheet_lower:
                return sheet
    return None

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

def apply_side_legend(fig):
    fig.update_layout(
        width=750,  
        height=450, 
        showlegend=True,
        legend=dict(
            title="", 
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.02,
            font=dict(size=14)
        ),
        margin=dict(l=0, r=0, t=20, b=20) 
    )
    return fig

# --- Загрузка данных из единого файла ---
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

df_recruitment = load_recruitment_data(file_bytes)
df_training = load_training_data(file_bytes)
df_attestation = load_generic_data(file_bytes, ['фио', 'категория', 'аттестация', 'дарача', 'должность'])

# --- Вкладки ---
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Аналитика найма", "Аналитика стажёров", "Аттестация", "Корпоративное обучение", "Общий штат", "Gallup"])

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

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Общее кол. кандидатов:", f"{total_candidates} чел.")
        m2.metric("Успешно трудоустроено", f"{hired} чел.")
        m3.metric("М / Ж (по найму)", gender_text)
        m4.metric("Ср. возраст (по найму)", f"{avg_age:.1f} лет" if not pd.isna(avg_age) else "Нет данных")

        st.divider()

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Общее кол-во заявок</b></div><br>", unsafe_allow_html=True)
        if 'Источник подбора' in filtered_rec.columns:
            src_traffic = filtered_rec['Источник подбора'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
            src_traffic.columns = ['Источник', 'Всего кандидатов']
            fig_source_vol = px.pie(src_traffic, names='Источник', values='Всего кандидатов', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_source_vol.update_traces(textposition='auto', textinfo='value')
            fig_source_vol = apply_side_legend(fig_source_vol)
            st.plotly_chart(fig_source_vol, use_container_width=False, key="fig_source_vol_tab1")

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
            fig_src_conv.update_layout(xaxis_range=[0, src_conv['Конверсия %'].max() * 1.3 if src_conv['Конверсия %'].max() > 0 else 100], coloraxis_colorbar=dict(title=""))
            st.plotly_chart(fig_src_conv, use_container_width=True, key="fig_src_conv_tab1")

        st.divider()

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Общее количество кандидатов: Филиалы, МХБ и ГО</b></div><br>", unsafe_allow_html=True)

        if 'Типы' in filtered_rec.columns and 'Все подразделения' in filtered_rec.columns:
            filtered_rec_copy = filtered_rec.copy()
            filtered_rec_copy['Типы_str'] = filtered_rec_copy['Типы'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            filtered_rec_copy['Подразделение_str'] = filtered_rec_copy['Все подразделения'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})

            st.markdown("#### 1. Общее количество кандидатов по ГО/Филиалу/ЦБО")
            types_data = filtered_rec_copy['Типы_str'].value_counts(dropna=False).reset_index()
            types_data.columns = ['Тип', 'Количество']
            if not types_data.empty:
                fig_types = px.pie(types_data, names='Тип', values='Количество', hole=0.4)
                fig_types.update_traces(textposition='auto', textinfo='percent') 
                fig_types = apply_side_legend(fig_types)
                st.plotly_chart(fig_types, use_container_width=False, key="fig_types_tab1")

            st.markdown("#### 2. Общее количество кандидатов по Филиалам")
            branches_df = filtered_rec_copy[filtered_rec_copy['Типы_str'].str.contains('Филиал|филиал', na=False, case=False)]
            branch_data = branches_df['Подразделение_str'].value_counts(dropna=False).reset_index()
            branch_data.columns = ['Филиал', 'Количество']
            if not branch_data.empty:
                fig_branch = px.pie(branch_data, names='Филиал', values='Количество', hole=0.4)
                fig_branch.update_traces(textposition='auto', textinfo='value')
                fig_branch = apply_side_legend(fig_branch)
                st.plotly_chart(fig_branch, use_container_width=False, key="fig_branch_tab1")

            st.markdown("#### 3. Общее количество кандидатов по ЦБО")
            mhb_df = filtered_rec_copy[filtered_rec_copy['Типы_str'].str.contains('МХБ|ЦБО|мхб|цбо', na=False, case=False)]
            mhb_data = mhb_df['Подразделение_str'].value_counts(dropna=False).reset_index()
            mhb_data.columns = ['МХБ/ЦБО', 'Количество']
            if not mhb_data.empty:
                fig_mhb = px.pie(mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                fig_mhb.update_traces(textposition='auto', textinfo='value')
                fig_mhb = apply_side_legend(fig_mhb)
                st.plotly_chart(fig_mhb, use_container_width=False, key="fig_mhb_tab1")

        st.divider()

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Принятые кандидаты: Филиалы, МХБ и ГО</b></div><br>", unsafe_allow_html=True)
        
        hired_df = filtered_rec[hired_mask].copy()

        if 'Типы' in hired_df.columns and 'Все подразделения' in hired_df.columns:
            hired_df['Типы_str'] = hired_df['Типы'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            hired_df['Подразделение_str'] = hired_df['Все подразделения'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            hired_df['Регион_calc'] = hired_df['Типы_str'].apply(categorize_region)

            st.markdown("#### 1. Принятые кандидаты по ГО/Филиалу/ЦБО")
            hired_reg_data = hired_df['Регион_calc'].value_counts().reset_index()
            hired_reg_data.columns = ['Регион', 'Количество']
            hired_reg_data = hired_reg_data[hired_reg_data['Количество'] > 0]
            if not hired_reg_data.empty:
                fig_hired_reg = px.pie(hired_reg_data, names='Регион', values='Количество', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                fig_hired_reg.update_traces(textposition='auto', textinfo='percent') 
                fig_hired_reg = apply_side_legend(fig_hired_reg)
                st.plotly_chart(fig_hired_reg, use_container_width=False, key="fig_hired_reg_tab1")
            else:
                st.info("Нет данных о найме")

            st.markdown("#### 2. Принятые кандидаты по Филиалам")
            hired_branches = hired_df[hired_df['Регион_calc'] == 'Филиал']
            hired_branch_data = hired_branches['Подразделение_str'].value_counts().reset_index()
            hired_branch_data.columns = ['Филиал', 'Количество']
            hired_branch_data = hired_branch_data[hired_branch_data['Количество'] > 0]
            if not hired_branch_data.empty:
                fig_hired_branch = px.pie(hired_branch_data, names='Филиал', values='Количество', hole=0.4)
                fig_hired_branch.update_traces(textposition='auto', textinfo='value')
                fig_hired_branch = apply_side_legend(fig_hired_branch)
                st.plotly_chart(fig_hired_branch, use_container_width=False, key="fig_hired_branch_tab1")
            else:
                st.info("В филиалах нет найма за этот период")

            st.markdown("#### 3. Принятые кандидаты по ЦБО")
            hired_mhb = hired_df[hired_df['Регион_calc'] == 'МХБ / ЦБО']
            hired_mhb_data = hired_mhb['Подразделение_str'].value_counts().reset_index()
            hired_mhb_data.columns = ['МХБ/ЦБО', 'Количество']
            hired_mhb_data = hired_mhb_data[hired_mhb_data['Количество'] > 0]
            if not hired_mhb_data.empty:
                fig_hired_mhb = px.pie(hired_mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                fig_hired_mhb.update_traces(textposition='auto', textinfo='value')
                fig_hired_mhb = apply_side_legend(fig_hired_mhb)
                st.plotly_chart(fig_hired_mhb, use_container_width=False, key="fig_hired_mhb_tab1")
            else:
                st.info("В ЦБО нет найма за этот период")

        st.divider()

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Регистрация кандидатов (по месяцам)</b></div><br>", unsafe_allow_html=True)
        if 'Дата регистрации' in filtered_rec.columns:
            trend_data = filtered_rec.dropna(subset=['Дата регистрации']).copy()
            if not trend_data.empty:
                trend_data['Период'] = trend_data['Дата регистрации'].dt.strftime('%Y-%m')
                trend_counts = trend_data['Период'].value_counts().reset_index().sort_values('Период')
                trend_counts.columns = ['Период', 'Кандидаты']
                
                fig_trend = px.line(trend_counts, x='Период', y='Кандидаты', markers=True, color_discrete_sequence=['#ff7f0e'])
                fig_trend.update_xaxes(type='category')
                fig_trend.update_layout(legend_title_text='')
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
        
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Стажёры: Филиалы, ЦБО, ГО</b></div><br>", unsafe_allow_html=True)

        if 'Типы' in filtered_int.columns and 'Все подразделения' in filtered_int.columns:
            filtered_int['Типы_str'] = filtered_int['Типы'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            filtered_int['Подразделение_str'] = filtered_int['Все подразделения'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            filtered_int['Регион_calc'] = filtered_int['Типы_str'].apply(categorize_region)

            st.markdown("#### 1. Стажёров: Филиалы, ЦБО, ГО")
            int_reg_data = filtered_int['Регион_calc'].value_counts().reset_index()
            int_reg_data.columns = ['Регион', 'Количество']
            int_reg_data = int_reg_data[int_reg_data['Количество'] > 0]
            if not int_reg_data.empty:
                fig_int_reg = px.pie(int_reg_data, names='Регион', values='Количество', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                fig_int_reg.update_traces(textposition='auto', textinfo='percent')
                fig_int_reg = apply_side_legend(fig_int_reg)
                st.plotly_chart(fig_int_reg, use_container_width=False, key="fig_int_reg_tab2")
            else:
                st.info("Нет данных")

            st.markdown("#### 2. Данные по филиалам")
            int_branches = filtered_int[filtered_int['Регион_calc'] == 'Филиал']
            int_branch_data = int_branches['Подразделение_str'].value_counts().reset_index()
            int_branch_data.columns = ['Филиал', 'Количество']
            int_branch_data = int_branch_data[int_branch_data['Количество'] > 0]
            if not int_branch_data.empty:
                fig_int_branch = px.pie(int_branch_data, names='Филиал', values='Количество', hole=0.4)
                fig_int_branch.update_traces(textposition='auto', textinfo='value')
                fig_int_branch = apply_side_legend(fig_int_branch)
                st.plotly_chart(fig_int_branch, use_container_width=False, key="fig_int_branch_tab2")
            else:
                st.info("В филиалах нет стажёров")

            st.markdown("#### 3. Данные по ЦБО")
            int_mhb = filtered_int[filtered_int['Регион_calc'] == 'МХБ / ЦБО']
            int_mhb_data = int_mhb['Подразделение_str'].value_counts().reset_index()
            int_mhb_data.columns = ['МХБ/ЦБО', 'Количество']
            int_mhb_data = int_mhb_data[int_mhb_data['Количество'] > 0]
            if not int_mhb_data.empty:
                fig_int_mhb = px.pie(int_mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                fig_int_mhb.update_traces(textposition='auto', textinfo='value')
                fig_int_mhb = apply_side_legend(fig_int_mhb)
                st.plotly_chart(fig_int_mhb, use_container_width=False, key="fig_int_mhb_tab2")
            else:
                st.info("В ЦБО нет стажёров")

        st.divider()
        
        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Распределение стажёров по должностям</b></div><br>", unsafe_allow_html=True)
        
        if 'Вакансия' in filtered_int.columns:
            vac_counts = filtered_int['Вакансия'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'}).value_counts().reset_index()
            vac_counts.columns = ['Вакансия', 'Количество']
            fig_vac_int = px.bar(vac_counts.head(15), x='Количество', y='Вакансия', orientation='h', text='Количество', color='Количество', color_continuous_scale='Purples')
            fig_vac_int.update_traces(textposition='outside')
            fig_vac_int.update_layout(yaxis={'categoryorder':'total ascending'}, coloraxis_colorbar=dict(title=""))
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

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Разбивка по подразделениям: Филиалы/ЦБО/ГО</b></div><br>", unsafe_allow_html=True)
        
        if col_region and col_subdiv:
            df_att_filtered['Типы_str'] = df_att_filtered[col_region].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            df_att_filtered['Подразделение_str'] = df_att_filtered[col_subdiv].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            df_att_filtered['Регион_calc'] = df_att_filtered['Типы_str'].apply(categorize_region)

            st.markdown("#### 1. Аттестация сотрудников по подразделениям: Филиалы/ЦБО/ГО")
            att_reg_data = df_att_filtered['Регион_calc'].value_counts().reset_index()
            att_reg_data.columns = ['Регион', 'Количество']
            att_reg_data = att_reg_data[att_reg_data['Количество'] > 0]
            if not att_reg_data.empty:
                fig_att_reg = px.pie(att_reg_data, names='Регион', values='Количество', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set3)
                fig_att_reg.update_traces(textposition='auto', textinfo='value')
                fig_att_reg = apply_side_legend(fig_att_reg)
                st.plotly_chart(fig_att_reg, use_container_width=False, key="fig_att_reg_tab3")
            else:
                st.info("Нет данных")

            st.markdown("#### 2. Данные по Филиалам")
            att_branches = df_att_filtered[df_att_filtered['Регион_calc'] == 'Филиал']
            att_branch_data = att_branches['Подразделение_str'].value_counts().reset_index()
            att_branch_data.columns = ['Филиал', 'Количество']
            att_branch_data = att_branch_data[att_branch_data['Количество'] > 0]
            if not att_branch_data.empty:
                fig_att_branch = px.pie(att_branch_data, names='Филиал', values='Количество', hole=0.4)
                fig_att_branch.update_traces(textposition='auto', textinfo='value')
                fig_att_branch = apply_side_legend(fig_att_branch)
                st.plotly_chart(fig_att_branch, use_container_width=False, key="fig_att_branch_tab3")
            else:
                st.info("В филиалах нет данных")

            st.markdown("#### 3. Данные по ЦБО")
            att_mhb = df_att_filtered[df_att_filtered['Регион_calc'] == 'МХБ / ЦБО']
            att_mhb_data = att_mhb['Подразделение_str'].value_counts().reset_index()
            att_mhb_data.columns = ['МХБ/ЦБО', 'Количество']
            att_mhb_data = att_mhb_data[att_mhb_data['Количество'] > 0]
            if not att_mhb_data.empty:
                fig_att_mhb = px.pie(att_mhb_data, names='МХБ/ЦБО', values='Количество', hole=0.4)
                fig_att_mhb.update_traces(textposition='auto', textinfo='value')
                fig_att_mhb = apply_side_legend(fig_att_mhb)
                st.plotly_chart(fig_att_mhb, use_container_width=False, key="fig_att_mhb_tab3")
            else:
                st.info("В ЦБО нет данных")
        else:
            st.warning("В файле не найдены колонки для определения региона или подразделения.")
            
    else:
        st.warning("Файл по аттестации не загружен.")

# === ВКЛАДКА 4: ОБУЧЕНИЕ ===
with tab4:
    df_training = load_training_data(file_bytes)
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
            fig_pf.update_layout(yaxis={'categoryorder':'total ascending'}, legend_title_text='')
            st.plotly_chart(fig_pf, use_container_width=True, key="fig_pf_tab4")

    else:
        st.warning("Файл по обучению не загружен.")

# === ВКЛАДКА 5: ОБЩИЙ ШТАТ (ИЗ ЛИСТА 'Штат') ===
with tab5:
    st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Аналитика: общий штат сотрудников</b></div><br>", unsafe_allow_html=True)
    
    try:
        sheet_name_staff = find_specific_sheet(file_bytes, ['штат', 'staff', 'сотрудники']) or "Штат"
        df_sheet_staff = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name_staff, header=None)
        
        months_raw = df_sheet_staff.iloc[1, 3:15].values
        hired_vals = df_sheet_staff.iloc[3, 3:15].values
        resigned_vals = df_sheet_staff.iloc[4, 3:15].values
        end_vals = df_sheet_staff.iloc[5, 3:15].values
        
        months_list = []
        hired_list = []
        resigned_list = []
        end_list = []
        
        for i, m in enumerate(months_raw):
            if pd.notna(m):
                m_str = m.strftime('%Y-%m') if hasattr(m, 'strftime') else str(m)
                months_list.append(m_str)
                hired_list.append(pd.to_numeric(hired_vals[i], errors='coerce'))
                resigned_list.append(pd.to_numeric(resigned_vals[i], errors='coerce'))
                end_list.append(pd.to_numeric(end_vals[i], errors='coerce'))

        df_chart = pd.DataFrame({
            'Месяц': months_list,
            'Принят на работу': hired_list,
            'Уволился': resigned_list,
            'Конец периода': end_list
        }).dropna(subset=['Принят на работу', 'Уволился'], how='all')

        latest_total = df_chart['Конец периода'].dropna().iloc[-1] if not df_chart.empty and not df_chart['Конец периода'].dropna().empty else 0
        total_hired_sum = df_chart['Принят на работу'].sum()

        m1, m2 = st.columns(2)
        m1.metric("Текущая численность штата", f"{int(latest_total)} чел.")
        m2.metric("Всего принято за период", f"{int(total_hired_sum)} чел.")

        st.divider()

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Динамика приема и увольнения сотрудников по месяцам</b></div><br>", unsafe_allow_html=True)
        
        df_melted = df_chart.melt(id_vars=['Месяц'], value_vars=['Принят на работу', 'Уволился'], 
                                  var_name='Показатель', value_name='Количество')
        
        fig_staff_dynamics = px.bar(
            df_melted, 
            x='Месяц', 
            y='Количество', 
            color='Показатель', 
            barmode='group',
            text='Количество',
            color_discrete_map={
                'Принят на работу': '#2ca02c', # Зеленый цвет
                'Уволился': '#d62728'          # Красный цвет
            }
        )
        fig_staff_dynamics.update_traces(textposition='outside')
        fig_staff_dynamics = apply_side_legend(fig_staff_dynamics)
        st.plotly_chart(fig_staff_dynamics, use_container_width=True, key="fig_staff_dynamics_tab5")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### Подробная таблица штатного расписания по месяцам")
        st.dataframe(df_chart, use_container_width=True)
        
    except Exception as e:
        st.error(f"Не удалось прочитать данные штата: {e}")

# === ВКЛАДКА 6: GALLUP (С МЕТРИКАМИ eNPS И СРАВНЕНИЕМ) ===
with tab6:
    st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Аналитика вовлеченности</b></div><br>", unsafe_allow_html=True)
    
    try:
        sheet_name_gallup = find_specific_sheet(file_bytes, ['gallup', 'галлоп', 'вовлеченность'])
        if sheet_name_gallup:
            df_gallup = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name_gallup)
            df_gallup.columns = df_gallup.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()
            
            quarter_cols = [c for c in df_gallup.columns if any(w in str(c).lower() for w in ['квартал', 'quarter', 'q', 'период', 'год', 'дата'])]
            
            if quarter_cols:
                q_col = quarter_cols[0]
                
                st.markdown(f"**Фильтр (`{q_col}`):**")
                quarters_list = df_gallup[q_col].dropna().unique()
                
                selected_quarters = st.multiselect(
                    "Выберите кварталы для сравнения:", 
                    options=list(quarters_list), 
                    default=list(quarters_list)
                )
                
                if selected_quarters:
                    df_gallup_filtered = df_gallup[df_gallup[q_col].isin(selected_quarters)].copy()
                    numeric_cols = df_gallup_filtered.select_dtypes(include=[np.number]).columns.tolist()
                    
                    if q_col in numeric_cols:
                        numeric_cols.remove(q_col)
                    
                    if numeric_cols:
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown("### Ключевые показатели выбранных периодов")

                        # Ищем колонку eNPS
                        enps_col = next((c for c in numeric_cols if 'enps' in str(c).lower() or 'nps' in str(c).lower()), None)
                        
                        cols = st.columns(len(selected_quarters))
                        for i, q in enumerate(selected_quarters):
                            q_data = df_gallup_filtered[df_gallup_filtered[q_col] == q]
                            with cols[i]:
                                if enps_col:
                                    enps_val = q_data[enps_col].mean()
                                    st.metric(f"eNPS ({q})", f"{enps_val:.1f}" if pd.notna(enps_val) else "Нет данных")
                                else:
                                    # Если eNPS нет, показываем среднее по всем числовым показателям
                                    avg_all = q_data[numeric_cols].mean().mean()
                                    st.metric(f"Ср. оценка ({q})", f"{avg_all:.2f}" if pd.notna(avg_all) else "Нет данных")

                        st.divider()
                        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Сравнение показателей вовлеченности (Q1-Q12)</b></div><br>", unsafe_allow_html=True)
                        
                        # Если eNPS есть, убираем его из списка для столбчатого графика (чтобы не ломал масштаб, если eNPS = 70, а вопросы = 4.5)
                        plot_cols = numeric_cols.copy()
                        if enps_col and enps_col in plot_cols:
                            plot_cols.remove(enps_col)

                        if plot_cols:
                            df_melted_gallup = df_gallup_filtered.melt(
                                id_vars=[q_col], 
                                value_vars=plot_cols, 
                                var_name='Показатель', 
                                value_name='Оценка'
                            )
                            
                            fig_gallup_all = px.bar(
                                df_melted_gallup, 
                                x='Показатель', 
                                y='Оценка', 
                                color=q_col, 
                                barmode='group',
                                text='Оценка',
                                color_discrete_sequence=px.colors.qualitative.Set2
                            )
                            fig_gallup_all.update_traces(textposition='outside')
                            fig_gallup_all.update_layout(xaxis_tickangle=-45)
                            fig_gallup_all = apply_side_legend(fig_gallup_all)
                            st.plotly_chart(fig_gallup_all, use_container_width=True, key="fig_gallup_all_metrics")
                                                    
                    st.markdown("#### Подробные данные опроса Gallup")
                    st.dataframe(df_gallup_filtered, use_container_width=True)
                else:
                    st.warning("⚠️ Пожалуйста, выберите хотя бы один квартал из фильтра выше.")
            else:
                st.info("Колонка с кварталом не обнаружена автоматически. Ниже приведена полная таблица Gallup:")
                st.dataframe(df_gallup, use_container_width=True)
        else:
            st.warning("Лист с названием 'Gallup' не найден в загруженном файле HR_data.xlsx.")
    except Exception as e:
        st.error(f"Ошибка при чтении листа Gallup: {e}")
