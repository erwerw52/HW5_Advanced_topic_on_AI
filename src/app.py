import streamlit as st
import os
import tempfile
from process_ppt import create_from_template
from pathlib import Path

# 設置頁面配置
st.set_page_config(
    page_title="PPT 風格轉換器",
    page_icon="🎮",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義 CSS 樣式 - 遊戲風格
st.markdown("""
    <style>
    /* 背景漸變動畫 */
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #8b5cf6 100%);
        animation: gradient 15s ease infinite;
        background-size: 200% 200%;
    }
    
    @keyframes gradient {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    /* 按鈕樣式 */
    .stButton>button {
        background: linear-gradient(90deg, #a855f7 0%, #ec4899 50%, #f59e0b 100%);
        color: white !important;
        font-size: 22px;
        font-weight: bold;
        padding: 20px 40px;
        border-radius: 20px;
        border: 4px solid #fbbf24;
        box-shadow: 0 10px 25px rgba(168, 85, 247, 0.5), 0 0 30px rgba(236, 72, 153, 0.3);
        transition: all 0.3s;
        text-transform: uppercase;
        letter-spacing: 2px;
    }
    
    .stButton>button:hover {
        transform: translateY(-5px) scale(1.05);
        box-shadow: 0 15px 35px rgba(168, 85, 247, 0.7), 0 0 50px rgba(236, 72, 153, 0.5);
        border-color: #fff;
    }
    
    /* 遊戲卡片 */
    .game-card {
        background: linear-gradient(145deg, rgba(139, 92, 246, 0.95), rgba(124, 58, 237, 0.95));
        border-radius: 20px;
        padding: 25px;
        box-shadow: 0 15px 35px rgba(0,0,0,0.3), inset 0 0 20px rgba(255,255,255,0.1);
        margin: 15px 0;
        border: 4px solid #a855f7;
        position: relative;
        overflow: hidden;
    }
    
    .game-card::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: linear-gradient(45deg, transparent, rgba(255,255,255,0.1), transparent);
        transform: rotate(45deg);
        animation: shine 3s infinite;
    }
    
    @keyframes shine {
        0% { transform: translateX(-100%) translateY(-100%) rotate(45deg); }
        100% { transform: translateX(100%) translateY(100%) rotate(45deg); }
    }
    
    /* 標題文字 */
    .title-text {
        color: #fbbf24;
        text-shadow: 0 0 10px #a855f7, 0 0 20px #ec4899, 0 0 30px #f59e0b,
                     4px 4px 0px #7c3aed, 6px 6px 0px #6b21a8;
        font-size: 3.5em;
        font-weight: bold;
        text-align: center;
        margin-bottom: 20px;
        animation: glow 2s ease-in-out infinite alternate;
    }
    
    @keyframes glow {
        from { text-shadow: 0 0 10px #a855f7, 0 0 20px #ec4899, 0 0 30px #f59e0b,
                            4px 4px 0px #7c3aed, 6px 6px 0px #6b21a8; }
        to { text-shadow: 0 0 20px #a855f7, 0 0 30px #ec4899, 0 0 40px #f59e0b,
                          4px 4px 0px #7c3aed, 6px 6px 0px #6b21a8; }
    }
    
    /* 副標題 */
    .subtitle-text {
        color: #fcd34d;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
        font-size: 1.3em;
        text-align: center;
        animation: pulse 2s ease-in-out infinite;
    }
    
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.7; }
    }
    
    /* 進度文字 */
    .progress-text {
        color: #fbbf24;
        font-size: 1.4em;
        font-weight: bold;
        text-align: center;
        text-shadow: 0 0 10px #a855f7, 2px 2px 4px rgba(0,0,0,0.5);
    }
    
    /* 檔案上傳器 */
    .stFileUploader {
        background: rgba(139, 92, 246, 0.3);
        border: 3px dashed #fbbf24;
        border-radius: 15px;
        padding: 20px;
    }
    
    /* 側邊欄 */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #7c3aed 0%, #5b21b6 100%);
    }
    
    /* 標題樣式 */
    h3 {
        color: #fbbf24 !important;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
    }
    
    /* 文字顏色 */
    .stMarkdown p, .stMarkdown li {
        color: white !important;
    }
    
    /* 遊戲標題區 */
    .game-header {
        background: linear-gradient(90deg, #a855f7, #ec4899);
        padding: 10px 20px;
        border-radius: 15px;
        margin-bottom: 20px;
        border: 3px solid #fbbf24;
        box-shadow: 0 5px 15px rgba(168, 85, 247, 0.5);
    }
    
    .game-header h2, .game-header h3 {
        margin: 5px 0 !important;
    }
    
    /* 下載卡片 */
    .download-card {
        background: linear-gradient(145deg, #a855f7, #ec4899);
        border-radius: 15px;
        padding: 20px;
        margin: 10px 0;
        border: 3px solid #fbbf24;
        box-shadow: 0 8px 20px rgba(236, 72, 153, 0.5);
        text-align: center;
    }
    
    /* 下載按鈕 */
    .stDownloadButton>button {
        background: linear-gradient(90deg, #fbbf24, #f59e0b) !important;
        color: #7c3aed !important;
        font-weight: bold;
        font-size: 18px;
        border: 3px solid #7c3aed !important;
        border-radius: 10px;
        padding: 12px 24px;
        box-shadow: 0 5px 15px rgba(251, 191, 36, 0.5);
    }
    
    .stDownloadButton>button:hover {
        transform: scale(1.05);
        box-shadow: 0 8px 20px rgba(251, 191, 36, 0.7);
    }
    
    /* 成功/錯誤/警告/資訊訊息 */
    .stSuccess, .stError, .stWarning, .stInfo {
        background: rgba(139, 92, 246, 0.2) !important;
        border-radius: 10px;
        padding: 10px;
        border-left: 5px solid #fbbf24 !important;
    }
    </style>
""", unsafe_allow_html=True)

# 標題
st.markdown('<h1 class="title-text">🎮 PPT 魔法轉換器 🎮</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle-text">✨ 上傳簡報，立即獲得兩種酷炫風格！✨</p>', unsafe_allow_html=True)

# 初始化 session state
if 'conversions' not in st.session_state:
    st.session_state.conversions = 0
if 'output_files' not in st.session_state:
    st.session_state.output_files = []

# 側邊欄
with st.sidebar:
    st.markdown('<div class="game-header">', unsafe_allow_html=True)
    st.markdown("## 🎯 遊戲規則")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("""
    ### 🕹️ 如何開始遊戲：
    
    **步驟 1️⃣** 📤 上傳你的 PPT 檔案
    
    **步驟 2️⃣** 🚀 點擊「開始魔法轉換」
    
    **步驟 3️⃣** 💾 下載兩種風格的簡報
    
    ---
    
    ### 🎨 自動生成風格：
    - 🌟 **Maeve 風格** - 現代科技感
    - 🎨 **水彩有機形狀** - 藝術水彩風
    
    ---
    
    ### ⚡ 遊戲設定：
    - 📁 格式：.pptx
    - 📦 大小：< 50MB
    - ⏱️ 時間：視檔案而定
    """)
    
    st.markdown("---")
    st.markdown('<div class="game-header">', unsafe_allow_html=True)
    st.markdown("### 🏆 遊戲統計")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.metric("🎯 完成關卡數", st.session_state.conversions)
    
    if st.session_state.conversions > 0:
        st.success(f"⭐ 已生成 {st.session_state.conversions * 2} 個風格檔案！")
    
    st.markdown("---")
    st.markdown("### 🎮 成就系統")
    if st.session_state.conversions >= 10:
        st.markdown("🏆 **轉換大師** - 完成10次轉換！")
    elif st.session_state.conversions >= 5:
        st.markdown("⭐ **風格玩家** - 完成5次轉換！")
    elif st.session_state.conversions >= 1:
        st.markdown("✨ **新手上路** - 完成首次轉換！")
    else:
        st.markdown("🎯 開始你的第一次轉換吧！")

# 主要內容區域
st.markdown('<div class="game-card">', unsafe_allow_html=True)
st.markdown("### 🎯 關卡任務：上傳你的簡報")

uploaded_file = st.file_uploader(
    "🎮 拖放檔案到這裡或點擊選擇",
    type=['pptx'],
    help="只支援 .pptx 格式的 PowerPoint 檔案"
)

if uploaded_file:
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.success(f"✅ **檔案名稱**\n{uploaded_file.name}")
    with col_info2:
        st.info(f"📦 **檔案大小**\n{uploaded_file.size / 1024:.2f} KB")
    with col_info3:
        st.warning(f"🎨 **將生成風格**\n2 種")
else:
    st.info("💡 上傳你的 PPT 檔案，系統將自動生成 Maeve 和水彩有機形狀兩種風格！")

st.markdown('</div>', unsafe_allow_html=True)

# 獲取模板資料夾中的模板 - 使用絕對路徑更可靠
script_dir = Path(__file__).parent.absolute()
project_root = script_dir.parent
template_dir = project_root / 'ppt' / 'template'

# 確保目錄存在並取得模板檔案
template_files = []
if template_dir.exists():
    template_files = list(template_dir.glob('*.pptx'))
else:
    st.error(f"❌ 找不到模板資料夾！")
    st.code(f"尋找路徑: {template_dir}")

# 動態建立風格列表（根據實際找到的檔案）
selected_styles = []
style_display_names = {
    "Maeve.pptx": "🌟 Maeve 風格",
    "WatercolorOrganicShapes.pptx": "🎨 水彩有機形狀風格"
}

for template_file in template_files:
    file_name = template_file.name
    display_name = style_display_names.get(file_name, f"🎨 {template_file.stem}")
    selected_styles.append((display_name, file_name, template_file))

# 顯示風格預覽（只在找到模板時顯示）
if len(selected_styles) >= 2:
    st.markdown('<div class="game-card">', unsafe_allow_html=True)
    st.markdown("### 🎨 即將生成的風格預覽")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #667eea, #764ba2); padding: 20px; border-radius: 15px; text-align: center; border: 3px solid #fbbf24;">
            <h3 style="color: #fbbf24; margin: 0;">🌟 Maeve 風格</h3>
            <p style="color: white; margin: 10px 0;">現代科技感設計</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_s2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #a855f7, #ec4899); padding: 20px; border-radius: 15px; text-align: center; border: 3px solid #fbbf24;">
            <h3 style="color: #fbbf24; margin: 0;">🎨 水彩有機形狀</h3>
            <p style="color: white; margin: 10px 0;">藝術水彩風格</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
elif len(selected_styles) == 0:
    st.error(f"❌ 找不到模板檔案！")
    st.info(f"📁 查找路徑: {template_dir}")
    st.info("💡 請確認以下檔案存在：\n- Maeve.pptx\n- WatercolorOrganicShapes.pptx")
elif len(selected_styles) == 1:
    st.warning(f"⚠️ 只找到 1 個模板檔案")
    st.info(f"已找到: {selected_styles[0][1]}")
    st.info("建議: 至少需要 2 個模板才能體驗完整功能")

# 轉換按鈕
st.markdown("<br>", unsafe_allow_html=True)
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    convert_button = st.button("🚀 開始魔法轉換！生成兩種風格 🎨", use_container_width=True)

# 處理轉換
if convert_button:
    if not uploaded_file:
        st.error("❌ 請先上傳一個 PPT 檔案！")
    elif len(selected_styles) == 0:
        st.error("❌ 找不到任何模板檔案！")
    else:
        # 創建臨時目錄
        with tempfile.TemporaryDirectory() as temp_dir:
            # 保存上傳的檔案到臨時目錄
            input_path = os.path.join(temp_dir, uploaded_file.name)
            with open(input_path, 'wb') as f:
                f.write(uploaded_file.read())
            
            st.markdown('<p class="progress-text">⚡ 轉換魔法啟動中...</p>', unsafe_allow_html=True)
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 處理每個風格
            output_files = []
            total_styles = len(selected_styles)
            
            for idx, (display_name, file_name, template_path) in enumerate(selected_styles):
                status_text.markdown(f'<p class="progress-text">✨ 魔法進行中... {display_name} ✨</p>', unsafe_allow_html=True)
                progress_bar.progress((idx + 0.5) / total_styles)
                
                # 生成輸出檔名: 原檔名_模板名.pptx
                base_name = Path(uploaded_file.name).stem
                template_name = Path(file_name).stem
                output_filename = f"{base_name}_{template_name}.pptx"
                output_path = os.path.join(temp_dir, output_filename)
                
                try:
                    # 執行轉換 (input_path, template_path, output_path)
                    create_from_template(input_path, str(template_path), output_path)
                    
                    # 讀取生成的檔案到記憶體
                    with open(output_path, 'rb') as f:
                        output_data = f.read()
                    
                    output_files.append({
                        'name': output_filename,
                        'data': output_data,
                        'style': display_name
                    })
                    
                    progress_bar.progress((idx + 1) / total_styles)
                    
                except Exception as e:
                    st.error(f"❌ {display_name} 轉換失敗: {str(e)}")
                    import traceback
                    with st.expander("查看錯誤詳情"):
                        st.code(traceback.format_exc())
                    continue
            
            # 清理進度顯示
            status_text.empty()
            progress_bar.empty()
            
            # 儲存結果到 session state 以便重複下載
            st.session_state.output_files = output_files
            
            # 顯示結果
            if output_files:
                st.success(f"🎉 關卡完成！成功生成 {len(output_files)} 種風格！")
                st.balloons()
                
                # 更新統計
                st.session_state.conversions += 1
                
                # 顯示獎勵訊息
                st.markdown("""
                    <div style="text-align: center; margin: 30px 0;">
                        <h2 style="color: #fbbf24; text-shadow: 0 0 20px #a855f7;">
                            🏆 任務完成！獲得獎勵 🏆
                        </h2>
                        <p style="color: #fcd34d; font-size: 1.3em; text-shadow: 2px 2px 4px rgba(0,0,0,0.5);">
                            ✨ 成功生成酷炫風格簡報！✨
                        </p>
                    </div>
                """, unsafe_allow_html=True)
                
                # 成就解鎖提示
                if st.session_state.conversions == 1:
                    st.info("🎯 成就解鎖：【新手上路】完成首次轉換！")
                elif st.session_state.conversions == 5:
                    st.warning("⭐ 成就解鎖：【風格玩家】完成 5 次轉換！")
                elif st.session_state.conversions == 10:
                    st.error("🏆 成就解鎖：【轉換大師】完成 10 次轉換！")
            else:
                st.error("😢 所有模板轉換都失敗了，請檢查錯誤訊息。")

# 顯示已生成的檔案（即使不在轉換按鈕區塊內也能下載）
if st.session_state.output_files:
    st.markdown('<div class="game-card">', unsafe_allow_html=True)
    st.markdown("### 💎 已生成的檔案 - 隨時可下載")
    
    cols = st.columns(len(st.session_state.output_files))
    for idx, output_file in enumerate(st.session_state.output_files):
        with cols[idx]:
            st.markdown(f"""
            <div class="download-card">
                <h3 style="color: #fbbf24; margin: 10px 0;">{output_file['style']}</h3>
                <p style="color: white; margin: 5px 0;">📄 {output_file['name']}</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.download_button(
                label=f"📥 下載 {output_file['style']}",
                data=output_file['data'],
                file_name=output_file['name'],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                key=f"persistent_download_{idx}"
            )
    
    st.markdown('</div>', unsafe_allow_html=True)

# 頁尾
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("""
    <div style="text-align: center; padding: 30px;">
        <div style="background: linear-gradient(90deg, #a855f7, #ec4899); padding: 20px; border-radius: 15px; border: 3px solid #fbbf24; box-shadow: 0 10px 30px rgba(168, 85, 247, 0.5);">
            <p style="color: #fbbf24; font-size: 1.3em; font-weight: bold; margin: 0; text-shadow: 2px 2px 4px rgba(0,0,0,0.5);">
                🎮 Made with 💜 by PPT Master 🎮
            </p>
            <p style="color: white; margin: 10px 0; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">
                ⭐ Keep Playing, Keep Creating! ⭐
            </p>
        </div>
    </div>
""", unsafe_allow_html=True)
