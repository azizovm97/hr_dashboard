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
                        st.error
