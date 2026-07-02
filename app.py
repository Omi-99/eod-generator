st.markdown(f"""
<style>
    /* ---------- GLOBAL ---------- */
    * {{
        box-sizing: border-box;
        margin: 0;
        padding: 0;
    }}
    .block-container {{
        padding: 0.6rem 0.8rem 0.4rem 0.8rem !important;
        max-width: 1200px !important;
        margin: 0 auto !important;
    }}
    .stApp {{
        background: {bg_primary};
        color: {text_color};
        transition: background 0.5s ease, color 0.3s ease;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }}
    /* ---------- HEADER ---------- */
    .main-header {{
        background: {header_grad};
        background-size: 300% 300%;
        animation: gradientFlow 8s ease infinite;
        padding: 0.6rem 1.2rem;
        border-radius: 20px;
        color: white;
        text-align: center;
        margin-bottom: 0.6rem;
        margin-top: 0.4rem;
        box-shadow: 0 8px 40px rgba(102, 126, 234, 0.4);
        border: 1px solid rgba(255,255,255,0.15);
        backdrop-filter: blur(4px);
        position: relative;
        overflow: hidden;
    }}
    .main-header::after {{
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(circle at 30% 30%, rgba(255,255,255,0.1) 0%, transparent 60%);
        pointer-events: none;
        animation: shine 10s linear infinite;
    }}
    @keyframes shine {{
        0% {{ transform: translateX(-100%) rotate(20deg); }}
        100% {{ transform: translateX(100%) rotate(20deg); }}
    }}
    @keyframes gradientFlow {{
        0% {{ background-position: 0% 50%; }}
        50% {{ background-position: 100% 50%; }}
        100% {{ background-position: 0% 50%; }}
    }}
    .main-header h1 {{
        font-size: 1.6rem;
        font-weight: 700;
        letter-spacing: -0.5px;
        margin: 0;
        line-height: 1.2;
        text-shadow: 0 2px 20px rgba(0,0,0,0.3);
        position: relative;
        z-index: 2;
    }}
    .main-header p {{
        font-size: 0.85rem;
        opacity: 0.9;
        margin: 0.1rem 0 0 0;
        font-weight: 300;
        letter-spacing: 0.3px;
        position: relative;
        z-index: 2;
    }}
    /* ---------- CARDS ---------- */
    .glass-card {{
        background: {card_bg};
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
        border-radius: 16px;
        padding: 0.5rem 0.8rem;
        margin-bottom: 0.4rem;
        box-shadow: 0 8px 32px {shadow_color};
        border: 1px solid {border_color};
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }}
    .glass-card:hover {{
        transform: translateY(-3px);
        box-shadow: 0 12px 48px {shadow_color};
    }}
    /* ---------- SLOT CARDS ---------- */
    .slot-card {{
        background: linear-gradient(135deg, rgba(102,126,234,0.12), rgba(118,75,162,0.12));
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
        border-radius: 12px;
        padding: 0.2rem 0.5rem;
        margin-bottom: 0.2rem;
        border: 1px solid {border_color};
        box-shadow: 0 4px 16px {shadow_color};
        transition: all 0.3s ease;
        color: {text_color};
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 6px;
    }}
    .slot-card:hover {{
        transform: scale(1.01);
        box-shadow: 0 6px 24px rgba(102,126,234,0.25);
        border-color: rgba(102,126,234,0.4);
    }}
    .slot-label {{
        font-weight: 600;
        font-size: 0.85rem;
        color: {text_color};
        min-width: 85px;
        margin-right: 6px;
        letter-spacing: 0.2px;
    }}
    .lunch-card {{
        background: linear-gradient(135deg, rgba(255,193,7,0.15), rgba(255,152,0,0.15));
        border: 1px solid rgba(255,193,7,0.3);
        color: {text_color};
    }}
    /* ---------- BUTTONS ---------- */
    .stButton button {{
        padding: 0.5rem 0.8rem !important;
        font-size: 0.9rem !important;
        border-radius: 12px !important;
        transition: all 0.3s ease !important;
        background: {header_grad} !important;
        background-size: 200% 200% !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 20px rgba(102, 126, 234, 0.4) !important;
        min-height: 48px !important;
        width: 100% !important;
        touch-action: manipulation !important;
        letter-spacing: 0.3px;
        position: relative;
        overflow: hidden;
    }}
    .stButton button:hover {{
        transform: scale(1.03);
        box-shadow: 0 8px 32px rgba(102, 126, 234, 0.6) !important;
        background-position: 100% 50% !important;
    }}
    .stButton button:active {{
        transform: scale(0.97);
    }}
    /* Download buttons */
    .stDownloadButton button {{
        padding: 0.5rem 0.8rem !important;
        font-size: 0.9rem !important;
        border-radius: 12px !important;
        transition: all 0.3s ease !important;
        background: #28a745 !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 20px rgba(40, 167, 69, 0.4) !important;
        min-height: 48px !important;
        width: 100% !important;
        touch-action: manipulation !important;
        letter-spacing: 0.3px;
    }}
    .stDownloadButton button:hover {{
        transform: scale(1.03);
        box-shadow: 0 8px 32px rgba(40, 167, 69, 0.6) !important;
    }}
    .stDownloadButton button:nth-of-type(2) {{
        background: #dc3545 !important;
        box-shadow: 0 4px 20px rgba(220, 53, 69, 0.4) !important;
    }}
    .stDownloadButton button:nth-of-type(2):hover {{
        box-shadow: 0 8px 32px rgba(220, 53, 69, 0.6) !important;
    }}
    /* ---------- SIDEBAR ---------- */
    .css-1d391kg {{
        background: {bg_secondary} !important;
        backdrop-filter: blur(20px) !important;
        -webkit-backdrop-filter: blur(20px) !important;
        color: {text_color} !important;
        padding: 0.3rem 0.4rem !important;
        border-right: 1px solid {border_color} !important;
        box-shadow: 4px 0 40px {shadow_color} !important;
        transition: all 0.3s ease;
    }}
    .css-1d391kg * {{
        font-size: 0.85rem !important;
        color: {text_color} !important;
    }}
    .css-1d391kg .stSelectbox select,
    .css-1d391kg .stTextInput input,
    .css-1d391kg .stDateInput input {{
        background: rgba(60, 60, 90, 0.4) !important;
        color: {text_color} !important;
        border: 1px solid {border_color} !important;
        border-radius: 10px !important;
        font-size: 0.85rem !important;
        padding: 0.4rem 0.6rem !important;
        height: 44px !important;
        backdrop-filter: blur(4px) !important;
        width: 100% !important;
        transition: border 0.3s ease;
    }}
    .css-1d391kg .stSelectbox select:focus,
    .css-1d391kg .stTextInput input:focus,
    .css-1d391kg .stDateInput input:focus {{
        border-color: #6C63FF !important;
        box-shadow: 0 0 0 3px rgba(108, 99, 255, 0.2) !important;
    }}
    .css-1d391kg .stButton button {{
        background: rgba(108, 99, 255, 0.7) !important;
        color: white !important;
        font-size: 0.8rem !important;
        padding: 0.3rem 0.6rem !important;
        height: 40px !important;
        border: 1px solid {border_color} !important;
        border-radius: 10px !important;
        backdrop-filter: blur(4px) !important;
        box-shadow: 0 2px 12px rgba(0,0,0,0.2) !important;
        min-height: 40px !important;
        transition: all 0.3s ease;
    }}
    .css-1d391kg .stButton button:hover {{
        background: rgba(108, 99, 255, 0.9) !important;
        transform: scale(1.02);
    }}
    /* ---------- INPUTS ---------- */
    .stTextInput input, .stDateInput input, .stSelectbox select {{
        background: rgba(60, 60, 90, 0.3) !important;
        color: {text_color} !important;
        border: 1px solid {border_color} !important;
        border-radius: 12px !important;
        font-size: 0.9rem !important;
        padding: 0.5rem 0.7rem !important;
        height: 48px !important;
        backdrop-filter: blur(4px) !important;
        width: 100% !important;
        transition: border 0.3s ease, box-shadow 0.3s ease;
    }}
    .stTextInput input:focus, .stDateInput input:focus, .stSelectbox select:focus {{
        border-color: #6C63FF !important;
        box-shadow: 0 0 0 4px rgba(108, 99, 255, 0.15) !important;
    }}
    /* ---------- PREVIEW TABLE ---------- */
    .preview-table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 4px;
        font-size: 12px;
        color: {table_text};
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 16px {shadow_color};
    }}
    .preview-table th {{
        background-color: {table_header_bg};
        font-weight: 600;
        border: 1px solid {table_border};
        padding: 6px 10px;
        text-align: left;
        color: {text_color};
        letter-spacing: 0.3px;
    }}
    .preview-table td {{
        border: 1px solid {table_border};
        padding: 6px 10px;
        text-align: left;
        background: {preview_bg};
    }}
    .preview-table tr:hover td {{
        background: rgba(108, 99, 255, 0.05);
    }}
    /* ---------- MOBILE RESPONSIVE ---------- */
    @media (max-width: 768px) {{
        .block-container {{
            padding: 0.3rem 0.4rem 0.2rem 0.4rem !important;
        }}
        .main-header h1 {{ font-size: 1.3rem !important; }}
        .main-header p {{ font-size: 0.7rem !important; }}
        .stButton button, .stDownloadButton button {{
            font-size: 0.9rem !important;
            padding: 0.5rem 0.4rem !important;
            min-height: 52px !important;
        }}
        .slot-label {{
            font-size: 0.75rem !important;
            min-width: 70px !important;
        }}
        .slot-card {{
            padding: 0.2rem 0.4rem !important;
            margin-bottom: 0.15rem !important;
        }}
        .stTextInput input, .stDateInput input, .stSelectbox select {{
            font-size: 16px !important;
            height: 48px !important;
            padding: 0.4rem 0.5rem !important;
        }}
        .css-1d391kg {{
            padding: 0.1rem 0.2rem !important;
            min-width: 260px !important;
        }}
        .css-1d391kg * {{
            font-size: 0.75rem !important;
        }}
        .stColumns {{
            flex-direction: column !important;
        }}
        .preview-table {{
            font-size: 10px !important;
        }}
        .preview-table th, .preview-table td {{
            padding: 4px 6px !important;
        }}
    }}
    @media (max-width: 480px) {{
        .main-header h1 {{ font-size: 1.1rem !important; }}
        .main-header p {{ font-size: 0.65rem !important; }}
        .slot-label {{
            font-size: 0.7rem !important;
            min-width: 60px !important;
        }}
        .preview-table {{
            font-size: 9px !important;
        }}
        .preview-table th, .preview-table td {{
            padding: 3px 4px !important;
        }}
    }}
</style>
""", unsafe_allow_html=True)