import streamlit as st
import json
import re
import os
import time
import shutil
import pandas as pd
from datetime import datetime
import openai
from groq import Groq
import google.generativeai as genai
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font, Alignment
from io import BytesIO
import requests
import ast
import tempfile
import subprocess

# ================= PAGE CONFIG =================
st.set_page_config(
    page_title="EOD Report Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================= HIDE FOOTER =================
st.markdown("""
<style>
    footer {visibility: hidden;}
    .stAppFooter {display: none;}
</style>
""", unsafe_allow_html=True)

# ================= THEME STATE =================
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

st.session_state.setdefault("selected_employee_name", "Omkar Patil")
st.session_state.setdefault("selected_employee_position", "Social Media & Digital Marketing Executive")

def toggle_theme():
    st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"

# ================= ROBUST DATE PARSER =================
def parse_date(date_str):
    if not date_str:
        return datetime.now().date()
    match = re.search(r'(\d{4}-\d{2}-\d{2})', date_str)
    if match:
        try:
            return datetime.strptime(match.group(1), "%Y-%m-%d").date()
        except ValueError:
            pass
    return datetime.now().date()

# ================= DYNAMIC TIME SLOTS =================
def get_time_slots(lunch_hour):
    slots = []
    for h in range(10, 18):
        start = h
        end = h + 1
        if h < 12:
            start_str = f"{h:02d}:00 am"
        elif h == 12:
            start_str = "12:00 pm"
        else:
            start_str = f"{h-12:02d}:00 pm"
        if end < 12:
            end_str = f"{end:02d}:00 am"
        elif end == 12:
            end_str = "12:00 pm"
        else:
            end_str = f"{end-12:02d}:00 pm"
        slots.append(f"{start_str} to {end_str}")
    return slots

def get_time_slots_short(lunch_hour):
    slots = []
    for h in range(10, 18):
        start = h
        end = h + 1
        if h < 12:
            start_str = f"{h:02d}:00"
        elif h == 12:
            start_str = "12:00"
        else:
            start_str = f"{h-12:02d}:00"
        if end < 12:
            end_str = f"{end:02d}:00"
        elif end == 12:
            end_str = "12:00"
        else:
            end_str = f"{end-12:02d}:00"
        slots.append(f"{start_str}-{end_str}")
    return slots

def create_short_to_full_map(lunch_hour):
    short = get_time_slots_short(lunch_hour)
    full = get_time_slots(lunch_hour)
    return dict(zip(short, full))

# ================= THEME-BASED STYLING =================
theme = st.session_state.theme
bg_primary = "#0f0f1a" if theme == "dark" else "#f0f2f6"
bg_secondary = "rgba(20,20,40,0.85)" if theme == "dark" else "rgba(255,255,255,0.85)"
text_color = "#ffffff" if theme == "dark" else "#000000"
card_bg = "rgba(30,30,50,0.6)" if theme == "dark" else "rgba(255,255,255,0.7)"
border_color = "rgba(255,255,255,0.15)" if theme == "dark" else "rgba(0,0,0,0.1)"
shadow_color = "rgba(0,0,0,0.3)" if theme == "dark" else "rgba(0,0,0,0.1)"
header_grad = "linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%)"
preview_bg = "rgba(255,255,255,0.08)" if theme == "dark" else "white"
preview_border = "#444" if theme == "dark" else "#ddd"
table_header_bg = "#2a2a3e" if theme == "dark" else "#f0f0f0"
table_text = "#ffffff" if theme == "dark" else "#000000"
table_border = "#555" if theme == "dark" else "#ddd"

st.markdown(f"""
<style>
    .block-container {{
        padding: 0.6rem 0.8rem 0.3rem 0.8rem !important;
        max-width: 100% !important;
    }}
    .stApp {{
        background: {bg_primary};
        color: {text_color};
        transition: background 0.3s ease, color 0.3s ease;
    }}
    .main-header {{
        background: {header_grad};
        background-size: 200% 200%;
        animation: gradientMove 6s ease infinite;
        padding: 0.5rem 0.6rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 0.4rem;
        margin-top: 0.4rem;
        box-shadow: 0 4px 20px rgba(102, 126, 234, 0.4);
    }}
    @keyframes gradientMove {{
        0% {{ background-position: 0% 50%; }}
        50% {{ background-position: 100% 50%; }}
        100% {{ background-position: 0% 50%; }}
    }}
    .main-header h1 {{
        font-size: 1.3rem;
        font-weight: 700;
        margin: 0;
        line-height: 1.3;
        text-shadow: 0 2px 10px rgba(0,0,0,0.3);
    }}
    .main-header p {{
        font-size: 0.75rem;
        opacity: 0.9;
        margin: 0;
        line-height: 1.3;
    }}
    .glass-card {{
        background: {card_bg};
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border-radius: 12px;
        padding: 0.4rem 0.6rem;
        margin-bottom: 0.4rem;
        box-shadow: 0 4px 20px {shadow_color};
        border: 1px solid {border_color};
        transition: all 0.3s ease;
        animation: fadeInUp 0.6s ease;
    }}
    .glass-card:hover {{
        transform: translateY(-3px);
        box-shadow: 0 8px 30px {shadow_color};
    }}
    @keyframes fadeInUp {{
        0% {{ opacity: 0; transform: translateY(20px); }}
        100% {{ opacity: 1; transform: translateY(0); }}
    }}
    .slot-card {{
        background: linear-gradient(135deg, rgba(102,126,234,0.15), rgba(118,75,162,0.15));
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
        border-radius: 10px;
        padding: 0.3rem 0.6rem;
        margin-bottom: 0.3rem;
        border: 1px solid {border_color};
        box-shadow: 0 4px 12px {shadow_color};
        transition: all 0.3s ease;
        color: {text_color};
    }}
    .slot-card:hover {{
        transform: scale(1.01);
        box-shadow: 0 6px 20px rgba(102,126,234,0.3);
    }}
    .slot-label {{
        font-weight: 600;
        font-size: 0.85rem;
        color: {text_color};
        display: inline-block;
        min-width: 100px;
    }}
    .lunch-card {{
        background: linear-gradient(135deg, rgba(255,193,7,0.2), rgba(255,152,0,0.2));
        border: 1px solid rgba(255,193,7,0.3);
        color: {text_color};
    }}
    .preview-card {{
        border: 1px solid {preview_border};
        border-radius: 10px;
        padding: 8px 12px;
        background: {preview_bg};
        box-shadow: 0 4px 15px {shadow_color};
        margin-bottom: 6px;
        backdrop-filter: blur(8px);
        transition: all 0.4s ease;
        animation: slideUp 0.8s ease;
        color: {text_color};
    }}
    @keyframes slideUp {{
        0% {{ opacity: 0; transform: translateY(40px); }}
        100% {{ opacity: 1; transform: translateY(0); }}
    }}
    .preview-title {{ font-size: 16px; font-weight: bold; text-align: center; margin-bottom: 4px; }}
    .preview-detail {{ display: flex; justify-content: space-between; padding: 2px 0; border-bottom: 1px solid {preview_border}; font-size: 13px; }}
    .preview-detail-label {{ font-weight: bold; width: 100px; }}
    .preview-table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 6px;
        font-size: 12px;
        color: {table_text};
    }}
    .preview-table th {{
        background-color: {table_header_bg};
        font-weight: bold;
        border: 1px solid {table_border};
        padding: 3px 6px;
        text-align: left;
        color: {text_color};
    }}
    .preview-table td {{
        border: 1px solid {table_border};
        padding: 3px 6px;
        text-align: left;
    }}
    .stTextInput, .stDateInput, .stSelectbox, .stTextArea {{ margin-bottom: 0.1rem !important; }}
    .stButton button {{
        padding: 0.3rem 0.8rem !important;
        font-size: 0.85rem !important;
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
        background: {header_grad} !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
        box-shadow: 0 2px 10px rgba(102, 126, 234, 0.4) !important;
        min-height: 44px !important;
        width: 100% !important;
    }}
    .stButton button:hover {{
        transform: scale(1.03);
        box-shadow: 0 4px 20px rgba(102, 126, 234, 0.6) !important;
    }}
    .stDownloadButton button {{
        padding: 0.3rem 0.8rem !important;
        font-size: 0.85rem !important;
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
        background: #28a745 !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
        box-shadow: 0 2px 10px rgba(40, 167, 69, 0.4) !important;
        min-height: 44px !important;
        width: 100% !important;
    }}
    .stDownloadButton button:hover {{
        transform: scale(1.03);
        box-shadow: 0 4px 20px rgba(40, 167, 69, 0.6) !important;
    }}
    .stDownloadButton button:nth-of-type(2) {{
        background: #dc3545 !important;
        box-shadow: 0 2px 10px rgba(220, 53, 69, 0.4) !important;
    }}
    .stDownloadButton button:nth-of-type(2):hover {{
        box-shadow: 0 4px 20px rgba(220, 53, 69, 0.6) !important;
    }}
    .css-1d391kg {{
        background: {bg_secondary} !important;
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        color: {text_color} !important;
        padding: 0.2rem 0.3rem !important;
        border-right: 1px solid {border_color} !important;
        box-shadow: 4px 0 30px {shadow_color} !important;
        transition: all 0.3s ease;
    }}
    .css-1d391kg * {{
        font-size: 0.8rem !important;
        color: {text_color} !important;
    }}
    .css-1d391kg .stSelectbox select,
    .css-1d391kg .stTextInput input,
    .css-1d391kg .stDateInput input {{
        background: rgba(60, 60, 90, 0.5) !important;
        color: {text_color} !important;
        border: 1px solid {border_color} !important;
        border-radius: 6px !important;
        font-size: 0.8rem !important;
        padding: 0.2rem 0.4rem !important;
        height: 2.2rem !important;
        backdrop-filter: blur(4px) !important;
    }}
    .css-1d391kg .stButton button {{
        background: rgba(120, 120, 200, 0.7) !important;
        color: {text_color} !important;
        font-size: 0.75rem !important;
        padding: 0.2rem 0.6rem !important;
        height: 2rem !important;
        border: 1px solid {border_color} !important;
        border-radius: 6px !important;
        backdrop-filter: blur(4px) !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2) !important;
        min-height: 38px !important;
    }}
    .css-1d391kg .stButton button:hover {{
        background: rgba(150, 150, 220, 0.9) !important;
        transform: scale(1.02);
    }}
    .css-1d391kg .stExpander {{
        border: 1px solid {border_color} !important;
        background: rgba(40, 40, 60, 0.4) !important;
        border-radius: 8px !important;
        padding: 0.1rem 0.2rem !important;
        backdrop-filter: blur(8px) !important;
        margin-bottom: 0.2rem !important;
    }}
    .css-1d391kg .stExpander .stExpanderHeader {{
        color: {text_color} !important;
        background: rgba(40, 40, 60, 0.2) !important;
        font-size: 0.8rem !important;
        padding: 0.2rem 0.4rem !important;
        border-radius: 6px !important;
    }}
    .css-1d391kg hr {{
        border-color: {border_color} !important;
        margin: 0.15rem 0 !important;
    }}
    .css-1d391kg .stImage {{
        margin-bottom: 0.1rem !important;
        filter: drop-shadow(0 0 8px rgba(102, 126, 234, 0.3)) !important;
    }}
    @media (max-width: 768px) {{
        .main-header h1 {{ font-size: 1.2rem !important; }}
        .main-header p {{ font-size: 0.7rem !important; }}
        .stButton button {{ font-size: 0.9rem !important; padding: 0.4rem 0.6rem !important; min-height: 48px !important; }}
        .stDownloadButton button {{ font-size: 0.9rem !important; padding: 0.4rem 0.6rem !important; min-height: 48px !important; }}
        .preview-table {{ font-size: 10px !important; }}
        .preview-table th, .preview-table td {{ padding: 2px 4px !important; }}
        .css-1d391kg {{ padding: 0.1rem 0.2rem !important; }}
        .css-1d391kg * {{ font-size: 0.7rem !important; }}
        .css-1d391kg .stButton button {{ font-size: 0.7rem !important; padding: 0.1rem 0.3rem !important; min-height: 34px !important; }}
    }}
</style>
""", unsafe_allow_html=True)

# ================= CONFIG =================
CONFIG_FILE = ".eod_config.json"
HISTORY_FILE = "history.json"
EMPLOYEES_FILE = "employees.json"

DEFAULT_EMPLOYEE = "Omkar Patil"
DEFAULT_POSITION = "Social Media & Digital Marketing Executive"

PROVIDERS = {
    "Groq (Fastest)": {
        "default_model": "llama-3.1-8b-instant",
        "api_key_required": True,
    },
    "OpenAI (ChatGPT)": {
        "default_model": "gpt-4o-mini",
        "api_key_required": True,
    },
    "Google Gemini": {
        "default_model": "gemini-1.5-flash",
        "api_key_required": True,
    },
    "Ollama (local)": {
        "default_model": "phi3",
        "api_key_required": False,
    }
}

# ============= TEMPLATE MANAGEMENT =============
TEMPLATE_DIR = "templates"
os.makedirs(TEMPLATE_DIR, exist_ok=True)

def get_template_list():
    files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.xlsx')]
    return sorted(files)

def load_template_bytes(filename):
    path = os.path.join(TEMPLATE_DIR, filename)
    if os.path.exists(path):
        with open(path, "rb") as f:
            return f.read()
    if os.path.exists(filename):
        with open(filename, "rb") as f:
            return f.read()
    return None

DEFAULT_TEMPLATE_BYTES = None
DEFAULT_TEMPLATE_FILE = "EOD_SAMPLE_FINAL.xlsx"
if os.path.exists(DEFAULT_TEMPLATE_FILE):
    with open(DEFAULT_TEMPLATE_FILE, "rb") as f:
        DEFAULT_TEMPLATE_BYTES = f.read()

# ============= HELPER =============
def generate_filename(date_obj, employee_name, extension):
    date_str = date_obj.strftime("%d-%m-%y")
    first_name = employee_name.split()[0].upper() if employee_name.strip() else "UNKNOWN"
    return f"EOD_{date_str}_{first_name}.{extension}"

# ============= PERSISTENCE HELPERS =============
def load_json_file(filepath, default=None):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r") as f:
                return json.load(f)
        except:
            return default if default is not None else {}
    return default if default is not None else {}

def save_json_file(filepath, data):
    with open(filepath, "w") as f:
        json.dump(data, f, indent=2)

def load_config():
    return load_json_file(CONFIG_FILE, {})

def save_config(provider, model, api_key=None):
    data = {"provider": provider, "model": model}
    if api_key is not None:
        data["api_key"] = api_key
    save_json_file(CONFIG_FILE, data)

def clear_config():
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)

# ============= EMPLOYEE & HISTORY =============
def load_employees():
    return load_json_file(EMPLOYEES_FILE, [])

def save_employees(employees):
    save_json_file(EMPLOYEES_FILE, employees)

def add_employee(name, position):
    employees = load_employees()
    if not any(e["name"] == name for e in employees):
        employees.append({"name": name, "position": position})
        save_employees(employees)
    return employees

def delete_employee(name):
    employees = load_employees()
    employees = [e for e in employees if e["name"] != name]
    save_employees(employees)
    return employees

def load_history():
    return load_json_file(HISTORY_FILE, [])

def save_history(history):
    save_json_file(HISTORY_FILE, history)

def add_history_entry(employee_name, position, date, schedule):
    history = load_history()
    history.append({
        "timestamp": datetime.now().isoformat(),
        "employee_name": employee_name,
        "position": position,
        "date": date,
        "schedule": schedule
    })
    save_history(history)

def clear_history():
    if os.path.exists(HISTORY_FILE):
        os.remove(HISTORY_FILE)

# ============= JSON PARSER =============
def extract_and_clean_json(raw_text):
    raw_text = re.sub(r'```json\s*', '', raw_text)
    raw_text = re.sub(r'```\s*', '', raw_text)
    start = raw_text.find('{')
    end = raw_text.rfind('}')
    if start != -1 and end != -1:
        json_str = raw_text[start:end+1]
    else:
        start = raw_text.find('[')
        end = raw_text.rfind(']')
        if start != -1 and end != -1:
            json_str = raw_text[start:end+1]
        else:
            raise ValueError("No JSON-like structure found")
    try:
        return json.loads(json_str)
    except:
        pass
    json_str = re.sub(r'([{,])\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*:', r'\1"\2":', json_str)
    json_str = re.sub(r',\s*([}\]])', r'\1', json_str)
    try:
        return json.loads(json_str)
    except:
        pass
    try:
        py_str = json_str.replace('null', 'None').replace('true', 'True').replace('false', 'False')
        return ast.literal_eval(py_str)
    except:
        pass
    raise ValueError("Could not parse JSON")

# ============= AI GENERATION =============
def generate_schedule(user_tasks, employee_name, position, report_date, provider, api_key, model_name, lunch_hour, progress_callback=None):
    if PROVIDERS[provider]["api_key_required"] and not api_key:
        raise ValueError(f"API key for {provider} is missing. Please enter it in the sidebar.")

    slot_labels = get_time_slots(lunch_hour)
    lunch_label = slot_labels[lunch_hour - 10]
    short_to_full = create_short_to_full_map(lunch_hour)

    force_dash_slots = set()
    if user_tasks:
        for line in user_tasks.split('\n'):
            line = line.strip()
            if not line:
                continue
            match = re.match(r'^(\d{2}:\d{2}-\d{2}:\d{2})\s*:\s*-\s*$', line)
            if match:
                short_time = match.group(1)
                if short_time in short_to_full:
                    full_slot = short_to_full[short_time]
                    force_dash_slots.add(full_slot)

    prompt = f"""
You are an assistant that fills an End‑of‑Day work report.

The report has these 8 hourly slots (lunch break is fixed at **{lunch_label}** – you MUST set activity="Lunch Break" and description="Lunch Break" for that slot):

The user has provided tasks for specific time slots in the following format:
{user_tasks}

Instructions:
- If the user wrote a task for a slot (e.g., "10:00-11:00: posted stories"), use that as the activity title and write a professional description (1‑2 sentences).
- If the user wrote "-" for a slot (e.g., "11:00-12:00: -"), set both activity and description to "-".
- If a slot is not mentioned, distribute the remaining tasks intelligently or fill with appropriate activities.
- For the lunch slot, always use activity="Lunch Break" and description="Lunch Break".

Return **only** a valid JSON object with:
- "employee_name"
- "position"
- "date"
- "schedule": an array of objects with keys "slot", "activity", "description". Include exactly the above 8 slots (lunch must be fixed).

Use double quotes for all keys and string values. No trailing commas. Do not include any text outside the JSON.

Employee: {employee_name}
Position: {position}
Date: {report_date}
"""
    timeout_seconds = 60 if provider == "Ollama (local)" else 30
    max_retries = 2
    last_error = None
    for attempt in range(max_retries):
        if progress_callback:
            progress_callback(30 + attempt * 20, f"Contacting AI ({provider})...")
        try:
            if provider == "Groq (Fastest)":
                client = Groq(api_key=api_key)
                model = model_name or PROVIDERS[provider]["default_model"]
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.2,
                    max_tokens=500,
                    timeout=timeout_seconds
                )
                raw = response.choices[0].message.content
            elif provider == "OpenAI (ChatGPT)":
                client = openai.OpenAI(api_key=api_key)
                model = model_name or PROVIDERS[provider]["default_model"]
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.2,
                    max_tokens=500,
                    timeout=timeout_seconds
                )
                raw = response.choices[0].message.content
            elif provider == "Google Gemini":
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_name or PROVIDERS[provider]["default_model"])
                response = model.generate_content(
                    prompt,
                    generation_config=genai.types.GenerationConfig(temperature=0.2, max_output_tokens=500),
                    request_options={"timeout": timeout_seconds}
                )
                raw = response.text
            elif provider == "Ollama (local)":
                try:
                    requests.get("http://localhost:11434", timeout=5)
                except requests.ConnectionError:
                    raise RuntimeError("Ollama is not running. Please start Ollama.")
                client = openai.OpenAI(base_url="http://localhost:11434/v1", api_key="dummy")
                model = model_name or PROVIDERS[provider]["default_model"]
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.2,
                    max_tokens=500,
                    timeout=timeout_seconds
                )
                raw = response.choices[0].message.content
            else:
                raise ValueError("Unsupported provider")

            try:
                data = extract_and_clean_json(raw)
                break
            except Exception as parse_error:
                last_error = f"JSON parsing error: {parse_error}"
                if attempt == max_retries - 1:
                    raise RuntimeError(f"Could not parse AI response: {raw[:300]}...\nError: {parse_error}")
                else:
                    time.sleep(1)
                    continue
        except Exception as e:
            last_error = str(e)
            if attempt == max_retries - 1:
                raise RuntimeError(f"AI request failed: {e}")
            else:
                time.sleep(2)
                continue
    else:
        raise RuntimeError(f"Failed after {max_retries} attempts. Last error: {last_error}")

    if "schedule" not in data or not isinstance(data["schedule"], list):
        data["schedule"] = []
    schedule_dict = {entry.get("slot", "").strip(): entry for entry in data["schedule"] if "slot" in entry}
    complete_schedule = []
    for slot in slot_labels:
        if slot == lunch_label:
            complete_schedule.append({"slot": slot, "activity": "Lunch Break", "description": "Lunch Break"})
        elif slot in schedule_dict:
            entry = schedule_dict[slot]
            if slot in force_dash_slots:
                entry["activity"] = "-"
                entry["description"] = "-"
            else:
                entry["activity"] = entry.get("activity", "No specific task")
                entry["description"] = entry.get("description", "No description provided.")
            complete_schedule.append(entry)
        else:
            complete_schedule.append({"slot": slot, "activity": "No specific task", "description": "No description provided."})
    data["schedule"] = complete_schedule
    data["employee_name"] = data.get("employee_name", employee_name)
    data["position"] = data.get("position", position)
    data["date"] = data.get("date", report_date)
    return data

# ============= EXCEL GENERATION =============
def create_excel_from_schedule(schedule_data, template_bytes=None, time_slots=None):
    if template_bytes is None:
        template_bytes = DEFAULT_TEMPLATE_BYTES
    if time_slots is None:
        time_slots = get_time_slots(13)

    try:
        if template_bytes:
            wb = load_workbook(BytesIO(template_bytes))
            ws = wb.active
        else:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws['B1'] = "EOD REPORT"
            ws.merge_cells('B1:C1')
            ws['B1'].font = Font(name='Calibri', size=16, bold=True)
            ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
            ws['B3'] = "Name Of Employee"
            ws['C3'] = schedule_data.get("employee_name", DEFAULT_EMPLOYEE)
            ws['B4'] = "Position"
            ws['C4'] = schedule_data.get("position", DEFAULT_POSITION)
            ws['B5'] = "Date"
            ws['C5'] = schedule_data.get("date", datetime.now().strftime("%Y-%m-%d"))
            for r in [3,4,5]:
                ws[f'B{r}'].font = Font(bold=True)
            ws['B7'] = "Time"
            ws['C7'] = "Activity"
            ws['D7'] = "Description"
            for col in ['B','C','D']:
                ws[f'{col}7'].font = Font(bold=True)
            for i, slot in enumerate(time_slots):
                ws[f'B{8+i}'] = slot
            ws.column_dimensions['B'].width = 18
            ws.column_dimensions['C'].width = 28
            ws.column_dimensions['D'].width = 35
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = False
            thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                          top=Side(style='thin'), bottom=Side(style='thin'))
            for row in ws.iter_rows(min_row=7, max_row=7+len(time_slots), min_col=2, max_col=4):
                for cell in row:
                    cell.border = thin
            for cell_addr in ['C3','C4','C5']:
                ws[cell_addr].border = thin
    except Exception as e:
        st.error(f"Error creating workbook: {e}")
        raise

    def safe_write(cell_address, value):
        try:
            cell = ws[cell_address]
            for merged_range in ws.merged_cells:
                if cell.coordinate in merged_range:
                    top_left = merged_range.start_cell
                    top_left.value = value
                    return
            cell.value = value
        except:
            try:
                ws[cell_address] = value
            except:
                pass

    safe_write('C3', schedule_data.get("employee_name", DEFAULT_EMPLOYEE))
    safe_write('C4', schedule_data.get("position", DEFAULT_POSITION))
    safe_write('C5', schedule_data.get("date", datetime.now().strftime("%Y-%m-%d")))

    schedule_list = schedule_data.get("schedule", [])
    start_row = 8
    for idx, slot in enumerate(time_slots):
        row = start_row + idx
        entry = next((item for item in schedule_list if item.get("slot", "").strip().lower() == slot.lower()), None)
        activity = entry.get("activity", "No specific task") if entry else "No specific task"
        description = entry.get("description", "No description provided.") if entry else "No description provided."
        safe_write(f'C{row}', activity)
        safe_write(f'D{row}', description)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============= PDF GENERATION =============
def create_pdf_from_schedule(schedule_data, template_bytes=None, time_slots=None):
    if template_bytes is None:
        template_bytes = DEFAULT_TEMPLATE_BYTES
    excel_bytes = create_excel_from_schedule(schedule_data, template_bytes, time_slots)
    excel_data = excel_bytes.getvalue()

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_xlsx:
        tmp_xlsx.write(excel_data)
        xlsx_path = tmp_xlsx.name

    pdf_path = xlsx_path.replace('.xlsx', '.pdf')

    try:
        subprocess.run([
            'soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', os.path.dirname(pdf_path),
            xlsx_path
        ], check=True, timeout=60)

        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()

        os.unlink(xlsx_path)
        os.unlink(pdf_path)

        return BytesIO(pdf_bytes)

    except (subprocess.CalledProcessError, FileNotFoundError, Exception) as e:
        if os.path.exists(xlsx_path):
            os.unlink(xlsx_path)
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)
        st.warning("LibreOffice not available. Using fallback PDF (formatting may differ).")
        return create_fallback_pdf(schedule_data, time_slots)

def create_fallback_pdf(schedule_data, time_slots=None):
    if time_slots is None:
        time_slots = get_time_slots(13)
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from io import BytesIO

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=15, leftMargin=15,
                            topMargin=15, bottomMargin=15)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Title'],
                                 alignment=1, fontSize=16, spaceAfter=8,
                                 fontName='Helvetica-Bold')
    elements.append(Paragraph("EOD REPORT", title_style))
    elements.append(Spacer(1, 4))

    detail_data = [
        ["Name Of Employee", schedule_data.get("employee_name", "N/A")],
        ["Position", schedule_data.get("position", "N/A")],
        ["Date", schedule_data.get("date", "N/A")]
    ]
    detail_table = Table(detail_data, colWidths=[2.0*inch, 5.5*inch])
    detail_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('BOX', (0,0), (-1,-1), 0, colors.white),
        ('INNERGRID', (0,0), (-1,-1), 0, colors.white),
    ]))
    elements.append(detail_table)
    elements.append(Spacer(1, 10))

    schedule_list = schedule_data.get("schedule", [])
    table_data = [["Time", "Activity", "Description"]]
    for entry in schedule_list:
        table_data.append([
            entry.get("slot", ""),
            entry.get("activity", ""),
            entry.get("description", "")
        ])

    col_widths = [1.2*inch, 2.2*inch, 3.9*inch]
    schedule_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    schedule_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#CCCCCC")),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('WORDWRAP', (0,0), (-1,-1), True),
    ]))
    elements.append(schedule_table)

    doc.build(elements)
    pdf_data = buffer.getvalue()
    buffer.close()
    return BytesIO(pdf_data)

# ============= CONFETTI =============
def confetti():
    st.components.v1.html("""
    <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1"></script>
    <script>
        confetti({
            particleCount: 150,
            spread: 70,
            origin: { y: 0.6 }
        });
        setTimeout(() => {
            confetti({
                particleCount: 100,
                spread: 50,
                origin: { y: 0.5 }
            });
        }, 500);
    </script>
    """, height=0)

# ============= STREAMLIT UI =============
st.markdown("""
<div class="main-header">
    <h1>📋 EOD Report Generator</h1>
    <p>AI‑powered End‑of‑Day reports – fast, accurate, and beautifully formatted</p>
</div>
""", unsafe_allow_html=True)

config = load_config()
saved_provider = config.get("provider", "Groq (Fastest)")
saved_model = config.get("model", "")
saved_api_key = config.get("api_key", "")

# ---- Sidebar ----
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/000000/google-forms.png", width=25)
    if st.button(f"🌓 Switch to {'Light' if st.session_state.theme == 'dark' else 'Dark'} Theme", use_container_width=True):
        toggle_theme()
        st.rerun()

    # ---- Lunch Break Slider ----
    st.markdown("## 🕒 Lunch Break")
    lunch_hour = st.slider(
        "Select lunch break start:",
        min_value=12,
        max_value=14,
        value=13,
        step=1,
        format="%d:00 PM"
    )

    with st.expander("⚙️ Config", expanded=False):
        provider = st.selectbox("AI Provider", options=list(PROVIDERS.keys()),
                                index=list(PROVIDERS.keys()).index(saved_provider) if saved_provider in PROVIDERS else 0)
        if PROVIDERS[provider]["api_key_required"]:
            if saved_api_key:
                st.success("✅ API key is set")
                if st.button("🔄 Change API Key"):
                    save_config(provider, model_name, "")
                    st.rerun()
            else:
                api_key = st.text_input("API Key", type="password")
                if st.button("💾 Save API Key"):
                    if api_key:
                        save_config(provider, model_name, api_key)
                        st.success("✅ API key saved")
                        st.rerun()
                    else:
                        st.warning("Please enter an API key")
                if "Groq" in provider:
                    st.caption("Get free key at [console.groq.com](https://console.groq.com)")
        else:
            api_key = None
            st.info("Ollama – no key needed.")
        model_name = st.text_input("Model (optional)",
                                   placeholder=PROVIDERS[provider]["default_model"],
                                   value=saved_model if saved_provider == provider else "")

        if 'prev_provider' not in st.session_state:
            st.session_state.prev_provider = provider
            st.session_state.prev_model = model_name

        if (provider != st.session_state.prev_provider or
            model_name != st.session_state.prev_model):
            current_key = saved_api_key if saved_api_key else ""
            save_config(provider, model_name, current_key)
            st.session_state.prev_provider = provider
            st.session_state.prev_model = model_name
            st.success("✅ Config saved")

        if st.button("🗑️ Clear all config (including API key)"):
            clear_config()
            st.success("Config cleared.")
            st.rerun()

    with st.expander("👥 Employees", expanded=False):
        employees = load_employees()
        employee_names = [e["name"] for e in employees]
        if employee_names:
            selected_emp = st.selectbox("Select employee", employee_names,
                                        index=employee_names.index(st.session_state.selected_employee_name) if st.session_state.selected_employee_name in employee_names else 0)
            emp_details = next((e for e in employees if e["name"] == selected_emp), None)
            if emp_details:
                st.session_state.selected_employee_name = emp_details["name"]
                st.session_state.selected_employee_position = emp_details["position"]
        else:
            selected_emp = None

        with st.expander("➕ Add"):
            new_name = st.text_input("Name")
            new_pos = st.text_input("Position")
            if st.button("Add"):
                if new_name and new_pos:
                    add_employee(new_name, new_pos)
                    st.success(f"Added {new_name}")
                    st.rerun()
                else:
                    st.warning("Fill both")

        if selected_emp and st.button("🗑️ Delete"):
            delete_employee(selected_emp)
            st.success(f"Deleted {selected_emp}")
            st.rerun()

    st.divider()
    st.markdown("## 📁 Template")

    template_files = get_template_list()
    if DEFAULT_TEMPLATE_BYTES is not None:
        default_name = DEFAULT_TEMPLATE_FILE
        if default_name not in template_files:
            template_files.insert(0, default_name)
    template_options = ["Built-in"] + template_files

    selected_template = st.selectbox("Select Template", template_options)
    if selected_template == "Built-in":
        template_bytes = None
        st.info("Using built‑in layout")
    else:
        template_bytes = load_template_bytes(selected_template)
        if template_bytes is None:
            st.warning(f"Template '{selected_template}' not found. Using built‑in.")
            template_bytes = None
        else:
            st.success(f"Loaded '{selected_template}'")

    uploaded_file = st.file_uploader("Upload .xlsx", type=["xlsx"])
    if uploaded_file is not None:
        save_path = os.path.join(TEMPLATE_DIR, uploaded_file.name)
        with open(save_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"Saved '{uploaded_file.name}' to templates folder.")
        st.rerun()

    with st.expander("📜 History", expanded=False):
        if st.button("🗑️ Clear History", use_container_width=True):
            clear_history()
            st.success("History cleared.")
            st.rerun()

        all_history = load_history()
        employee_list = sorted(set(entry["employee_name"] for entry in all_history))
        selected_filter = st.selectbox("Filter", ["All"] + employee_list)
        if selected_filter != "All":
            filtered_history = [h for h in all_history if h["employee_name"] == selected_filter]
        else:
            filtered_history = all_history

        if filtered_history:
            groups = {}
            for entry in reversed(filtered_history):
                emp = entry.get("employee_name", "Unknown")
                if emp not in groups:
                    groups[emp] = []
                groups[emp].append(entry)
            for emp, entries in groups.items():
                with st.expander(f"👤 {emp} ({len(entries)})"):
                    for idx, entry in enumerate(entries[:5]):
                        display_date = entry.get("date", "N/A")
                        timestamp = entry.get("timestamp", str(idx))
                        if st.button(f"📂 {display_date}", key=f"load_{emp}_{timestamp}"):
                            st.session_state.loaded_report = entry
                            st.rerun()
                        st.caption(f"Position: {entry.get('position', '')}")
        else:
            st.write("No reports yet.")

        if "loaded_report" in st.session_state and st.session_state.loaded_report is not None:
            if st.button("🗑️ Clear loaded", use_container_width=True):
                st.session_state.loaded_report = None
                st.rerun()

# ---- Session state ----
if "loaded_report" not in st.session_state:
    st.session_state.loaded_report = None
if "last_schedule" not in st.session_state:
    st.session_state.last_schedule = None
if "last_input" not in st.session_state:
    st.session_state.last_input = ""
if "last_config" not in st.session_state:
    st.session_state.last_config = {}
if "current_schedule" not in st.session_state:
    st.session_state.current_schedule = None

# ---- Main area ----
left_col, right_col = st.columns([0.4, 0.6], gap="small")

with left_col:
    st.markdown("### ✍️ Inputs")
    employees = load_employees()
    employee_names = [e["name"] for e in employees]
    if employee_names:
        current_name = st.session_state.selected_employee_name
        if current_name not in employee_names:
            current_name = employee_names[0]
        selected_emp_main = st.selectbox("👤 Employee", employee_names, index=employee_names.index(current_name))
        emp_details_main = next((e for e in employees if e["name"] == selected_emp_main), None)
        if emp_details_main:
            st.session_state.selected_employee_name = emp_details_main["name"]
            st.session_state.selected_employee_position = emp_details_main["position"]
    else:
        st.session_state.selected_employee_name = DEFAULT_EMPLOYEE
        st.session_state.selected_employee_position = DEFAULT_POSITION

    position = st.text_input("💼 Position", value=st.session_state.selected_employee_position)
    report_date = st.date_input("📅 Date", value=datetime.now())

    st.markdown("### 📝 Task Summary")
    st.caption("Describe your tasks for the day. You can use '10:00-11:00: task' to pin tasks to specific slots, or just list them.")

    if "free_task_input" not in st.session_state:
        st.session_state.free_task_input = ""

    user_tasks = st.text_area(
        "",
        value=st.session_state.free_task_input,
        height=200,
        placeholder="e.g., 10:00-11:00: posted stories\n11:00-12:00: -\n12:00-1:00: started reel editing\n2:00-3:00: created AI images\n...",
        key="free_task_input"
    )

    templates_quick = {
        "Feng Shui & Content": "10:00-11:00: Posted Feng Shui stories\n11:00-12:00: Created post on July animal signs\n12:00-13:00: Started editing Kedarnath reel\n14:00-15:00: Created 15 AI creatives for reel\n15:00-16:00: Continued editing reel\n16:00-17:00: Reviewed performance\n17:00-18:00: Planned next steps",
        "Meetings & Documentation": "10:00-11:00: Team sync meeting\n11:00-12:00: Wrote meeting notes\n12:00-13:00: Follow-up emails\n14:00-15:00: Project planning\n15:00-16:00: Client call\n16:00-17:00: Prepared status report\n17:00-18:00: Reviewed and finalized",
        "Development & Testing": "10:00-11:00: Fixed bugs\n11:00-12:00: Developed new feature\n12:00-13:00: Code review\n14:00-15:00: Wrote tests\n15:00-16:00: Deployed to staging\n16:00-17:00: Updated documentation\n17:00-18:00: Sprint planning",
        "Custom": ""
    }
    selected_template_quick = st.selectbox("📝 Quick template", list(templates_quick.keys()), key="quick_template")
    if selected_template_quick != "Custom" and templates_quick[selected_template_quick]:
        if st.button("📋 Load template", use_container_width=True):
            st.session_state.free_task_input = templates_quick[selected_template_quick]
            st.rerun()

    col_gen, col_reg = st.columns(2)
    with col_gen:
        generate_clicked = st.button("🚀 Generate", type="primary", use_container_width=True)
    with col_reg:
        regenerate_clicked = st.button("🔄 Regenerate", use_container_width=True,
                                       disabled=st.session_state.last_schedule is None)

with right_col:
    # Editable schedule using text_inputs
    if st.session_state.current_schedule is not None:
        st.markdown("### 📊 Edit Schedule")
        st.caption("You can edit Activity and Description directly below.")

        schedule_list = st.session_state.current_schedule
        edited_schedule = []

        for idx, entry in enumerate(schedule_list):
            cols = st.columns([2, 2, 3])  # Time Slot, Activity, Description
            with cols[0]:
                st.markdown(f"**{entry['slot']}**")
            with cols[1]:
                new_activity = st.text_input(
                    "Activity",
                    value=entry["activity"],
                    key=f"edit_act_{idx}",
                    label_visibility="collapsed"
                )
            with cols[2]:
                new_description = st.text_input(
                    "Description",
                    value=entry["description"],
                    key=f"edit_desc_{idx}",
                    label_visibility="collapsed"
                )
            edited_schedule.append({
                "slot": entry["slot"],
                "activity": new_activity,
                "description": new_description
            })

        # Update session state with edits
        st.session_state.current_schedule = edited_schedule

        # Build schedule data for download
        schedule_data = {
            "employee_name": st.session_state.selected_employee_name,
            "position": st.session_state.selected_employee_position,
            "date": report_date.strftime("%Y-%m-%d"),
            "schedule": edited_schedule
        }
        st.session_state.last_schedule = schedule_data

        # Download buttons
        time_slots = get_time_slots(lunch_hour)
        excel_data = create_excel_from_schedule(schedule_data, template_bytes, time_slots)
        pdf_data = create_pdf_from_schedule(schedule_data, template_bytes, time_slots)

        file_date = parse_date(schedule_data["date"])
        emp_name = schedule_data["employee_name"]
        excel_filename = generate_filename(file_date, emp_name, "xlsx")
        pdf_filename = generate_filename(file_date, emp_name, "pdf")

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="📥 Download Excel",
                data=excel_data,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col2:
            st.download_button(
                label="📄 Download PDF",
                data=pdf_data,
                file_name=pdf_filename,
                mime="application/pdf",
                use_container_width=True
            )
    else:
        st.info("👈 Type your tasks and click Generate to create a schedule.")

# ---- Generation logic ----
if generate_clicked or regenerate_clicked:
    if regenerate_clicked:
        tasks = st.session_state.last_input
        cfg = st.session_state.last_config
        provider_used = cfg["provider"]
        api_key_used = cfg["api_key"]
        model_used = cfg["model"]
        emp_used = cfg["employee"]
        pos_used = cfg["position"]
        date_used = cfg["date"]
        lunch_hour_used = cfg.get("lunch_hour", 13)
    else:
        tasks = st.session_state.free_task_input.strip()
        if not tasks:
            st.warning("Please describe your daily tasks.")
            st.stop()
        if PROVIDERS[provider]["api_key_required"]:
            saved_key = load_config().get("api_key", "")
            if saved_key:
                api_key = saved_key
            else:
                api_key = ""
            if not api_key:
                st.warning("Please enter your API key in the sidebar config.")
                st.stop()
        else:
            api_key = None
        emp_used = st.session_state.selected_employee_name
        pos_used = st.session_state.selected_employee_position
        date_used = report_date.strftime("%Y-%m-%d")
        lunch_hour_used = lunch_hour
        st.session_state.last_input = tasks
        st.session_state.last_config = {
            "provider": provider,
            "api_key": api_key,
            "model": model_name,
            "employee": emp_used,
            "position": pos_used,
            "date": date_used,
            "lunch_hour": lunch_hour_used
        }
        provider_used = provider
        api_key_used = api_key
        model_used = model_name

    progress_bar = st.progress(0)
    status_text = st.empty()
    def update_progress(progress, message):
        progress_bar.progress(progress)
        status_text.markdown(f"<span style='color:#667eea; font-weight:500;'>{message}</span>", unsafe_allow_html=True)

    try:
        update_progress(10, "🚀 Initializing...")
        data = generate_schedule(tasks, emp_used, pos_used, date_used, provider_used, api_key_used, model_used, lunch_hour_used, progress_callback=update_progress)
        update_progress(95, "✨ Finalizing...")

        st.session_state.last_schedule = data
        add_history_entry(emp_used, pos_used, date_used, data["schedule"])
        emp_list = load_employees()
        if not any(e["name"] == emp_used for e in emp_list):
            add_employee(emp_used, pos_used)

        # Store the schedule in session state for editing
        st.session_state.current_schedule = data["schedule"]

        st.balloons()
        confetti()
        st.success("✅ Report generated successfully! 🎉")
        status_text.empty()
        progress_bar.empty()
        st.rerun()
    except Exception as e:
        st.error(f"❌ {e}")
        progress_bar.empty()
        status_text.empty()
        st.session_state.last_schedule = None