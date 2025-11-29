import os
import tempfile
import logging
from dotenv import load_dotenv

load_dotenv()
logging.basicConfig(level=logging.INFO)

from src.document_parser import parse_document
from src.ai_analyzer import analyze_document, GEMINI_AVAILABLE
from src.template_manager import get_template, TEMPLATES
from src.ppt_generator import create_pptx

# Store last analysis result for template switching
_last_analysis = None
_last_filename = None


def generate_pptx_from_file(path: str, forced_theme: str = None) -> tuple[str, str]:
    """Generate a PPTX file from the given document path.
    Returns (output_path, detected_theme).
    """
    global _last_analysis, _last_filename
    
    structured = parse_document(path)
    analysis = analyze_document(structured)
    _last_analysis = analysis
    _last_filename = os.path.splitext(os.path.basename(path))[0]
    
    theme = forced_theme or analysis.get('theme', '簡約')
    template = get_template(theme)
    out_name = f'{_last_filename}.pptx'
    output_path = os.path.join(tempfile.gettempdir(), out_name)
    create_pptx(analysis['slides'], template, output_path)
    return output_path, theme


def regenerate_with_theme(theme: str) -> tuple[str, str]:
    """Regenerate PPT with a different template using cached analysis."""
    global _last_analysis, _last_filename
    
    if _last_analysis is None:
        raise ValueError('請先上傳並生成 PPT')
    
    template = get_template(theme)
    out_name = f'{_last_filename}.pptx'
    output_path = os.path.join(tempfile.gettempdir(), out_name)
    create_pptx(_last_analysis['slides'], template, output_path)
    return output_path, theme


def main():
    """Launch Gradio web UI."""
    if not GEMINI_AVAILABLE:
        print('⚠️  GEMINI_API_KEY not set: using heuristic summarizer (basic mode)')
    else:
        print('✅ GEMINI_API_KEY detected: using AI-powered analysis')
    
    import gradio as gr

    custom_css = """
    @import url('https://fonts.googleapis.com/css2?family=Bangers&family=Fredoka:wght@400;600&family=Space+Mono&display=swap');
    
    /* ========== CANDY ARCADE THEME ========== */
    
    /* Animated Gradient Background */
    body, .gradio-container {
        background: linear-gradient(-45deg, #ff6b6b, #feca57, #48dbfb, #ff9ff3, #54a0ff);
        background-size: 400% 400%;
        animation: gradientShift 15s ease infinite;
        min-height: 120vh;
    }
    
    @keyframes gradientShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Main Container */
    .gradio-container > .main {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 30px;
        margin: 20px auto;
        padding: 30px;
        box-shadow: 
            0 20px 60px rgba(0, 0, 0, 0.15),
            0 0 0 5px rgba(255, 255, 255, 0.8),
            inset 0 0 80px rgba(255, 255, 255, 0.5);
        max-width: 1200px;
        overflow-x: hidden;
    }

    /* Prevent horizontal scroll on body */
    body, html {
        overflow-x: hidden !important;
    }

    .gradio-container {
        overflow-x: hidden !important;
    }

    /* ========== TITLE SECTION ========== */
    /* Wrapper to contain the wiggle animation and prevent layout shift */
    .prose:first-of-type {
        overflow: hidden;
        padding: 20px;
        text-align: center;
    }

    h1 {
        font-family: 'Bangers', cursive !important;
        font-size: 3em !important;
        color: #2d3436 !important;
        text-align: center;
        letter-spacing: 4px;
        text-shadow: 
            3px 3px 0px #ff6b6b,
            6px 6px 0px #feca57,
            9px 9px 0px #48dbfb;
        margin: 10px 10px 20px 10px !important;
        padding: 25px 30px !important;
        background: linear-gradient(135deg, #fff9c4 0%, #ffe082 100%);
        border-radius: 20px;
        border: 4px solid #2d3436;
        box-shadow: 8px 8px 0px #2d3436;
        transform-origin: center center;
        animation: wiggle 3s ease-in-out infinite;
        /* Center the title */
        display: block;
        width: calc(100% - 20px);
        max-width: calc(100% - 20px);
    }
    
    @keyframes wiggle {
        0%, 100% { transform: rotate(-1deg); }
        50% { transform: rotate(1deg); }
    }
    
    /* Subtitle */
    .prose p {
        font-family: 'Fredoka', sans-serif !important;
        color: #636e72 !important;
        text-align: center;
        font-size: 1.3em !important;
        font-weight: 600;
        background: #dfe6e9;
        padding: 12px 24px;
        border-radius: 50px;
        display: inline-block;
        margin: 15px auto;
        border: 3px solid #b2bec3;
    }

    /* ========== CARD BLOCKS ========== */
    .block {
        background: #ffffff !important;
        border-radius: 20px !important;
        border: 4px solid #2d3436 !important;
        box-shadow: 6px 6px 0px #2d3436 !important;
        padding: 20px !important;
        transition: box-shadow 0.2s ease;
        overflow: visible !important;
    }
    
    .block:hover {
        box-shadow: 8px 8px 0px #2d3436 !important;
    }

    /* ========== LABELS ========== */
    label span {
        font-family: 'Fredoka', sans-serif !important;
        font-size: 18px !important;
        font-weight: 600 !important;
        color: #2d3436 !important;
        background: linear-gradient(135deg, #fd79a8 0%, #e84393 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    /* ========== MEGA BUTTON ========== */
    button.primary {
        font-family: 'Bangers', cursive !important;
        font-size: 24px !important;
        letter-spacing: 3px;
        background: linear-gradient(135deg, #00b894 0%, #00cec9 50%, #0984e3 100%) !important;
        border: 4px solid #2d3436 !important;
        box-shadow: 6px 6px 0px #2d3436 !important;
        color: #ffffff !important;
        padding: 20px 40px !important;
        border-radius: 16px !important;
        text-shadow: 2px 2px 0px rgba(0,0,0,0.3);
        transition: all 0.15s ease;
        position: relative;
        overflow: hidden;
    }
    
    button.primary::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: linear-gradient(
            45deg,
            transparent 40%,
            rgba(255,255,255,0.3) 50%,
            transparent 60%
        );
        animation: shine 3s infinite;
    }
    
    @keyframes shine {
        0% { transform: translateX(-100%) rotate(45deg); }
        100% { transform: translateX(100%) rotate(45deg); }
    }
    
    button.primary:hover {
        transform: translate(-3px, -3px) scale(1.02);
        box-shadow: 9px 9px 0px #2d3436 !important;
    }
    
    button.primary:active {
        transform: translate(3px, 3px);
        box-shadow: 3px 3px 0px #2d3436 !important;
    }

    /* ========== GAME LOG (Terminal Style) ========== */
    textarea {
        font-family: 'Space Mono', monospace !important;
        font-size: 15px !important;
        color: #e17055 !important;
        background: #F0F8FF !important;
        border: 3px solid #e17055 !important;
        border-radius: 16px !important;
        padding: 16px !important;
        box-shadow: 
            inset 0 0 30px rgba(0, 255, 136, 0.1),
            6px 6px 0px #2d3436;
        line-height: 1.6;
    }

    /* ========== ITEM NAME INPUT (Dark background with light text) ========== */
    input[type="text"],
    .textbox input {
        font-family: 'Space Mono', monospace !important;
        font-size: 14px !important;
        background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%) !important;
        border: 3px solid #e17055 !important;
        border-radius: 12px !important;
        color: #00ff88 !important;
        font-weight: 600;
        padding: 12px 16px !important;
        box-shadow: inset 0 0 15px rgba(0, 255, 136, 0.1);
    }

    input[type="text"]::placeholder,
    .textbox input::placeholder {
        color: #636e72 !important;
    }

    /* ========== DROPDOWN ========== */
    /* Make sure dropdown is not hidden */
    div[data-testid="dropdown"],
    div[data-testid="dropdown"] > div,
    .wrap {
        overflow: visible !important;
    }

    div[data-testid="dropdown"] input {
        font-family: 'Fredoka', sans-serif !important;
        font-size: 16px !important;
        font-weight: 600;
        color: #2d3436 !important;
        background: linear-gradient(135deg, #ffeaa7 0%, #fdcb6e 100%) !important;
        border: 3px solid #2d3436 !important;
        border-radius: 12px !important;
        padding: 12px 16px !important;
        box-shadow: 4px 4px 0px #2d3436;
        cursor: pointer;
    }

    /* Dropdown list - use multiple selectors for compatibility */
    ul[role="listbox"],
    div[data-testid="dropdown"] ul,
    .dropdown-menu,
    [class*="dropdown"] ul {
        z-index: 99999 !important;
        background: #ffeaa7 !important;
        border: 3px solid #2d3436 !important;
        border-radius: 12px !important;
        box-shadow: 6px 6px 0px #2d3436 !important;
    }

    ul[role="listbox"] li,
    div[data-testid="dropdown"] ul li {
        font-family: 'Fredoka', sans-serif !important;
        font-size: 16px !important;
        font-weight: 600;
        color: #2d3436 !important;
        padding: 12px 16px !important;
        border-bottom: 2px dashed #f39c12;
        transition: all 0.2s ease;
        cursor: pointer;
    }

    ul[role="listbox"] li:hover,
    div[data-testid="dropdown"] ul li:hover {
        background: #f39c12 !important;
        color: #ffffff !important;
    }
    
    ul[role="listbox"] li:last-child,
    div[data-testid="dropdown"] ul li:last-child {
        border-bottom: none;
    }

    /* Override any overflow:hidden that might hide dropdown */
    .block {
        overflow: visible !important;
    }

    /* ========== FILE UPLOAD ========== */
    div[data-testid="file"] {
        background: linear-gradient(135deg, #a29bfe 0%, #6c5ce7 100%) !important;
        border: 4px dashed #2d3436 !important;
        border-radius: 20px !important;
        padding: 30px !important;
        transition: all 0.3s ease;
    }
    
    div[data-testid="file"]:hover {
        transform: scale(1.02);
        border-style: solid;
    }

    div[data-testid="file"] span {
        font-family: 'Fredoka', sans-serif !important;
        font-size: 16px !important;
        font-weight: 600;
        color: #ffffff !important;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
    }

    div[data-testid="file"] a {
        font-family: 'Fredoka', sans-serif !important;
        font-size: 16px !important;
        font-weight: 600;
        color: #2d3436 !important;
        background: #ffeaa7 !important;
        padding: 8px 16px !important;
        border-radius: 8px !important;
        border: 3px solid #2d3436;
        box-shadow: 3px 3px 0px #2d3436;
        text-decoration: none !important;
        transition: all 0.2s ease;
    }

    div[data-testid="file"] a:hover {
        transform: translate(-2px, -2px);
        box-shadow: 5px 5px 0px #2d3436;
        background: #fdcb6e !important;
    }

    /* ========== SECTION HEADERS ========== */
    h4, h3 {
        font-family: 'Bangers', cursive !important;
        font-size: 20px !important;
        color: #e17055 !important;
        letter-spacing: 2px;
        text-shadow: 2px 2px 0px #fab1a0;
        margin: 20px 0 10px 0 !important;
    }

    /* ========== STATUS BADGE ========== */
    .prose strong {
        font-family: 'Fredoka', sans-serif !important;
        font-size: 18px !important;
        font-weight: 600;
        color: #ffffff !important;
        background: linear-gradient(135deg, #00b894 0%, #55a3ff 100%);
        padding: 8px 20px;
        border-radius: 50px;
        border: 3px solid #2d3436;
        box-shadow: 4px 4px 0px #2d3436;
        display: inline-block;
    }

    /* ========== ITEM NAME BOX (Light text on dark background) ========== */
    .item-name-box input {
        color: #e17055 !important;
        font-family: 'Space Mono', monospace !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%) !important;
        border: 3px solid #00ff88 !important;
    }

    /* ========== DOWNLOAD SECTION ========== */
    .download-file {
        background: linear-gradient(135deg, #55efc4 0%, #00b894 100%) !important;
        border: 4px solid #2d3436 !important;
        border-radius: 20px !important;
        box-shadow: 6px 6px 0px #2d3436 !important;
        animation: pulse 2s ease-in-out infinite;
    }
    
    @keyframes pulse {
        0%, 100% { box-shadow: 6px 6px 0px #2d3436; }
        50% { box-shadow: 6px 6px 20px rgba(0, 184, 148, 0.5), 6px 6px 0px #2d3436; }
    }

    .download-file span {
        color: #2d3436 !important;
    }

    .download-file a {
        background: #fff !important;
        color: #00b894 !important;
        font-weight: 700 !important;
        border: 3px solid #2d3436 !important;
    }

    /* ========== FUN EXTRAS ========== */
    /* Emoji bounce on labels */
    label span::first-letter {
        display: inline-block;
        animation: bounce 1s ease infinite;
    }
    
    @keyframes bounce {
        0%, 100% { transform: translateY(0); }
        50% { transform: translateY(-5px); }
    }

    /* Row spacing */
    .row {
        gap: 24px !important;
    }

    /* Column styling */
    .column {
        gap: 16px !important;
    }

    /* Hide footer for cleaner look */
    footer {
        display: none !important;
    }
    """

    def on_generate(uploaded_file, choice_theme):
        if uploaded_file is None:
            return '❌ 請先上傳文件', None, None, 'Auto'
        try:
            src_path = uploaded_file.name
            theme = choice_theme if choice_theme != 'Auto' else None
            out, detected_theme = generate_pptx_from_file(src_path, forced_theme=theme)
            filename = os.path.basename(out)
            status = f'✅ MISSION SUCCESS\n------------------\n[ITEM ACQUIRED]: {filename}\n[SKIN EQUIPPED]: {detected_theme}'
            return status, out, filename, detected_theme
        except Exception as e:
            return f'❌ SYSTEM FAILURE: {str(e)}', None, None, 'Auto'

    with gr.Blocks(css=custom_css, title='🎮 Auto PPT Generator') as demo:
        gr.Markdown('# 🕹️ AUTO PPT GENERATOR')
        gr.Markdown('> INSERT COIN TO START... OR JUST UPLOAD A FILE! <')
        
        with gr.Row():
            with gr.Column():
                file_input = gr.File(label='📂 LOAD CARTRIDGE', file_types=['.txt', '.md', '.docx'])
                theme_dropdown = gr.Dropdown(
                    choices=['Auto', '專業技術', '學術研究', '簡約'],
                    value='Auto',
                    label='🎨 SELECT CLASS'
                )
                generate_btn = gr.Button('🚀 START GAME', variant='primary')
            
            with gr.Column():
                output_text = gr.Textbox(label='📟 GAME LOG', lines=6)
                gr.Markdown('#### 💾 YOUR LOOT')
                output_filename = gr.Textbox(label='📄 ITEM NAME', interactive=False, elem_classes=['item-name-box'])
                output_file = gr.File(label='⬇️ GRAB IT!', elem_classes=['download-file'])
        
        # Show API status
        api_status = '🟢 ONLINE' if GEMINI_AVAILABLE else '🔴 OFFLINE (Insert Token)'
        gr.Markdown(f'**SERVER STATUS**: {api_status}')

        generate_btn.click(
            on_generate,
            inputs=[file_input, theme_dropdown],
            outputs=[output_text, output_file, output_filename, theme_dropdown]
        )

    demo.launch()


if __name__ == '__main__':
    main()
