import streamlit as st

def inject_premium_styles():
    """Injects global 'Editorial Ledger' CSS and Typography (Manrope/Inter)."""
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@200;400;700;800&family=Inter:wght@300;400;600&display=swap');

        /* Typography: Apply to content but NOT to system icons/expander-arrows */
        .stMarkdown p, .stMarkdown li, .stMarkdown label, .stTable, .stDataFrame {
            font-family: 'Inter', sans-serif !important;
            color: #1b1c1c !important;
        }

        /* Headers with Manrope */
        h1, h2, h3, h4, h5, h6 {
            font-family: 'Manrope', sans-serif !important;
            font-weight: 700 !important;
            color: #050e39 !important;
            letter-spacing: -0.02em !important;
        }

        /* The Editorial Card Surface */
        .premium-card {
            background-color: #ffffff;
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 4px 24px rgba(0, 0, 0, 0.04);
            border: 1px solid rgba(198, 197, 208, 0.2);
        }

        /* Action Hub Theme */
        .action-hub-error { background: linear-gradient(145deg, #fff5f5, #ffffff); border-left: 5px solid #ba1a1a; border-radius: 8px; padding: 20px; margin-bottom: 16px; }
        .action-hub-warning { background: linear-gradient(145deg, #fffaf3, #ffffff); border-left: 5px solid #e8881f; border-radius: 8px; padding: 20px; margin-bottom: 16px; }

        /* Glossy Primary Button - High Contrast Fix */
        div.stButton > button {
            background: linear-gradient(135deg, #050e39 0%, #1c244e 100%) !important;
            border: none !important;
            border-radius: 8px !important;
            padding: 12px 32px !important;
            transition: all 0.3s ease !important;
            box-shadow: 0 4px 12px rgba(5, 14, 57, 0.2) !important;
        }
        
        /* Force White Text on ALL button children (p, span, labels) */
        div.stButton > button * {
            color: #ffffff !important;
            font-weight: 700 !important;
            font-family: 'Manrope', sans-serif !important;
            text-shadow: 0 1px 2px rgba(0,0,0,0.2) !important;
        }

        /* Refined Expander (FIX: Don't break the SVG icons) */
        .stExpander {
            border: none !important;
            background-color: #f8f9fa !important;
            border-radius: 12px !important;
            margin-bottom: 16px !important;
            border: 1px solid rgba(0,0,0,0.05) !important;
        }
        
        /* Targeted Title text - Leave the summary arrow alone */
        div[data-testid="stExpanderSummary"] > div:last-child p {
            font-family: 'Manrope', sans-serif !important;
            font-weight: 700 !important;
            font-size: 1.05rem !important;
            color: #050e39 !important;
            margin: 0 !important;
        }

        /* Pill Badges */
        .pill-error {
            background-color: #ffdad6;
            color: #ba1a1a;
            padding: 2px 10px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }
        .pill-warning {
            background-color: #ffdcc1;
            color: #6c3a00;
            padding: 2px 10px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }
        </style>
    """, unsafe_allow_html=True)

def render_premium_header(title, subtitle=None):
    """Renders a styled header for content sections."""
    st.markdown(f"### {title}")
    if subtitle:
        st.markdown(f"<p style='color: #46464f; margin-top: -10px; margin-bottom: 20px;'>{subtitle}</p>", unsafe_allow_html=True)

def action_hub_container(type='error'):
    """Context manager for rendering audit findings in styled containers."""
    css_class = 'action-hub-error' if type == 'error' else 'action-hub-warning'
    return st.container() # In Streamlit, styling full containers requires CSS injection based on nesting or IDs, but for this MVP we'll use markdown blocks inside.

def render_finding_card(title, data_dict, type='error'):
    """Renders a high-fidelity card for audit findings (minified for Streamlit compatibility)."""
    bg_color = "#fff5f5" if type == 'error' else "#fffaf3"
    border_color = "#ba1a1a" if type == 'error' else "#e8881f"
    text_color = "#93000a" if type == 'error' else "#673700"
    
    # We build a single-line minified HTML string to avoid Streamlit markdown parsing issues
    html = f'<div style="background: {bg_color}; border-left: 5px solid {border_color}; border-radius: 8px; padding: 16px; margin-bottom: 16px;">'
    html += f'<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">'
    html += f'<h5 style="color: {text_color}; border: none; padding: 0; margin: 0;">{title}</h5>'
    html += '</div>'
    html += '<div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-top: 12px;">'
    
    for label, value in data_dict.items():
        html += '<div>'
        html += f'<p style="font-size: 0.75rem; color: #46464f; margin: 0;">{label}</p>'
        html += f'<p style="font-size: 1rem; font-weight: 600; color: {text_color}; margin: 0;">{value}</p>'
        html += '</div>'
        
    html += '</div></div>'
    st.markdown(html, unsafe_allow_html=True)
