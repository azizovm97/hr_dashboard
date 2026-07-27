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
    st.header("🔐 Авторизация")
    
    if not st.session_state['authenticated']:
        with st.form("login_form"):
            username_input = st.text_input("Логин")
            password_input = st.text_input("Пароль", type="password")
            submit_login = st.form_submit_button("Войти 🚀")
            
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
        st.success(f"Вы вошли как:\n**{st.session_state['name']}**\n*(Права: {st.session_state['role'].upper()})*")
        if st.button("Выйти (Logout) 🚪"):
            st.session_state['authenticated'] = False
            st.session_state['username'] = ""
            st.session_state['role'] = ""
            st.session_state['name'] = ""
            st.rerun()

    st.divider()

    # --- ПАНЕЛЬ УПРАВЛЕНИЯ ДЛЯ АДМИНА ---
    if st.session_state['role'] == "admin":
        with st.expander("🛠️ Админ-панель", expanded=False):
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

    st.header("📂 Управление данными")
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
        st.info("👁️ У вас права на просмотр. Загрузка файлов заблокирована.")

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
            row_str = [str(v).lower().strip() if pd.notna(v) else '' for v in row]
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
    if 'го' in val_str or 'ho' in val_str or 'головной' in val_str or 'центральный' in val_str or 'сарбонк' in val_str: return 'Головной офис'
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
def load_attestation_data(file_bytes: bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    target_sheet = None
    for sheet in xls.sheet_names:
        if 'аттестация' in sheet.lower() and 'справочник' not in sheet.lower():
            target_sheet = sheet
            break
            
    if not target_sheet: return None
    
    df_preview = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, header=None, nrows=15)
    header_idx = 0
    for idx, row in df_preview.iterrows():
        row_str = [str(v).lower().strip() if pd.notna(v) else '' for v in row]
        if 'результат' in row_str or any('фио' in v for v in row_str):
            if len([v for v in row_str if v]) > 3:
                header_idx = idx
                break
                
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, header=header_idx)
    df.columns = df.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()
    return df

@st.cache_data(show_spinner=False)
def load_daraja_data(file_bytes: bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    target_sheet = None
    for sheet in xls.sheet_names:
        if 'дараҷа' in sheet.lower() or 'дарача' in sheet.lower():
            target_sheet = sheet
            break
            
    if not target_sheet: return None
    
    df_preview = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, header=None, nrows=15)
    header_idx = 0
    for idx, row in df_preview.iterrows():
        row_str = [str(v).lower().strip() if pd.notna(v) else '' for v in row]
        if 'грейд до' in row_str or 'грейд после' in row_str or 'дата интервью' in row_str or any('фио' in v for v in row_str):
            if len([v for v in row_str if v]) > 3:
                header_idx = idx
                break
                
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, header=header_idx)
    df.columns = df.columns.str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True).str.strip()
    return df

df_recruitment = load_recruitment_data(file_bytes)
df_training = load_training_data(file_bytes)
df_attestation = load_attestation_data(file_bytes)
df_daraja = load_daraja_data(file_bytes)

# --- Вкладки ---
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Аналитика найма", "Аналитика стажёров", "Аттестация", "Корпоративное обучение", "Общий штат", "Gallup", "Дараҷа"
])

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

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Успешный найм кандидатов: Филиалы, МХБ и ГО (Только принятые)</b></div><br>", unsafe_allow_html=True)
        
        hired_df = filtered_rec[hired_mask].copy()

        if 'Типы' in hired_df.columns and 'Все подразделения' in hired_df.columns:
            hired_df['Типы_str'] = hired_df['Типы'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            hired_df['Подразделение_str'] = hired_df['Все подразделения'].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            hired_df['Регион_calc'] = hired_df['Типы_str'].apply(categorize_region)

            st.markdown("#### 1. Успешный найм кандидатов по ГО/Филиалу/ЦБО")
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

            st.markdown("#### 2. Успешный найм кандидатов по Филиалам")
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

            st.markdown("#### 3. Успешный найм кандидатов по ЦБО")
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
                st.info("В МХБ / ЦБО нет найма за этот период")

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

            st.markdown("#### 1. Распределение стажёров: Филиалы, ЦБО, ГО")
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
        col_region = cols_lower.get('типы') or cols_lower.get('регион') or cols_lower.get('тип')
        col_subdiv = cols_lower.get('департамент / подразделение') or cols_lower.get('все подразделения') or cols_lower.get('подразделение') or cols_lower.get('отдел')

        df_att_filtered = df_attestation.copy()

        # --- ПОДСЧЕТ И ФИЛЬТРАЦИЯ ПО РЕЗУЛЬТАТУ ---
        passed_count = 0
        failed_count = 0

        if 'Результат' in df_att_filtered.columns:
            df_att_filtered['Результат_clean'] = df_att_filtered['Результат'].astype(str).str.strip().str.lower()
            
            passed_count = len(df_att_filtered[df_att_filtered['Результат_clean'] == 'сдал'])
            failed_count = len(df_att_filtered[df_att_filtered['Результат_clean'] == 'не сдал'])

            valid_results = ['сдал', 'не сдал']
            mask = df_att_filtered['Результат_clean'].isin(valid_results)
            df_att_filtered = df_att_filtered[mask]

        total_attested = len(df_att_filtered)

        m1, m2, m3 = st.columns(3)
        m1.metric("Количество аттестованных", f"{total_attested} чел.")
        if 'Результат' in df_att_filtered.columns:
            m2.metric("Успешно (Сдал)", f"{passed_count} чел.")
            m3.metric("Не успешно (Не сдал)", f"{failed_count} чел.")
            
        st.divider()

        st.markdown("<div style='background-color: #D32F2F; color: white; padding: 10px; border-radius: 8px; text-align: center;'><b>Разбивка по подразделениям: Филиалы/ЦБО/ГО</b></div><br>", unsafe_allow_html=True)
        
        if col_region and col_subdiv:
            df_att_filtered['Типы_str'] = df_att_filtered[col_region].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            df_att_filtered['Подразделение_str'] = df_att_filtered[col_subdiv].astype(str).str.strip().replace({'nan': 'Не указано', '': 'Не указано'})
            df_att_filtered['Регион_calc'] = df_att_filtered['Типы_
