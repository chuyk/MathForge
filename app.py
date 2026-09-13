import os
import sys
import json
import re
import time
import io
import zipfile
import concurrent.futures
import tempfile
import traceback
import streamlit as st
import docx
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import numpy as np

# 確保模組搜尋路徑包含當前目錄與 _tools
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from exam_builder import build_student_exam, build_teacher_exam
from prompts.adapter_prompt import SYSTEM_PROMPT

# =======================================================
# 頁面配置與現代化極致視覺
# =======================================================
st.set_page_config(
    page_title="阿凱老師的數學考卷改題排版神器",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自訂優雅 CSS 樣式
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=Outfit:wght@500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Noto Sans TC', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 2.2rem 2.5rem;
        border-radius: 16px;
        color: white;
        box-shadow: 0 10px 25px -5px rgba(30, 58, 138, 0.25);
        margin-bottom: 2rem;
    }
    .main-header h1 {
        color: white;
        font-family: 'Outfit', sans-serif;
        font-size: 2.3rem;
        font-weight: 700;
        margin: 0 0 0.5rem 0;
        letter-spacing: -0.5px;
    }
    .main-header p {
        color: #e0e7ff;
        font-size: 1.1rem;
        margin: 0;
        opacity: 0.95;
    }
    .badge {
        display: inline-block;
        background-color: rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(8px);
        padding: 0.3rem 0.8rem;
        border-radius: 9999px;
        font-size: 0.85rem;
        font-weight: 600;
        margin-bottom: 0.8rem;
    }
    .feature-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.2rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #ecfdf5;
        border: 1px solid #a7f3d0;
        border-left: 6px solid #10b981;
        padding: 1.2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

# =======================================================
# 憑證與設定本地持久化存取
# =======================================================
SETTINGS_FILE = os.path.join(CURRENT_DIR, ".user_settings.json")

def load_local_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_local_settings(code, key):
    try:
        data = {"activation_code": code, "api_key": key}
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass

# =======================================================
# 側邊欄：啟動碼、API 金鑰與模型選擇
# =======================================================
saved_cfg = load_local_settings()
init_code = saved_cfg.get("activation_code", "")
init_key = saved_cfg.get("api_key", os.environ.get("GEMINI_API_KEY", ""))

with st.sidebar:
    st.image("https://img.icons8.com/isometric/100/compass--v1.png", width=70)
    st.markdown("### ⚙️ 核心設定 (Settings)")
    
    # 啟動碼驗證欄位（支援持久化暫存）
    activation_code = st.text_input(
        "🔐 系統啟動碼 (Activation Code)",
        type="password",
        value=init_code,
        placeholder="請輸入啟動碼",
        help="必須輸入正確啟動碼（kai）才能啟用改題排版功能。"
    )
    if activation_code.strip() == "kai":
        st.success("✅ 啟動碼驗證通過", icon="🟢")
    elif activation_code.strip():
        st.error("❌ 啟動碼錯誤", icon="🔴")
    else:
        st.info("ℹ️ 請輸入啟動碼以解鎖功能")

    st.divider()

    # API 金鑰欄位（支援持久化暫存）
    api_key = st.text_input(
        "🔑 Gemini API Key",
        type="password",
        value=init_key,
        help="使用者自備金鑰（BYOK），金鑰自動暫存於本機/瀏覽器，不用每次重複輸入。"
    )
    
    # 自動保存至本地快取
    if activation_code or api_key:
        save_local_settings(activation_code, api_key)
    
    st.markdown("""
    <small>
        👉 <a href="https://aistudio.google.com/app/apikey" target="_blank" style="color: #2563eb; text-decoration: none; font-weight: 500;">
        免費取得 Google AI Studio API Key ↗
        </a>
    </small>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    st.markdown("#### 🤖 AI 模型選擇")
    model_options = [
        "gemini-3.8-flash",
        "gemini-3.7-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.5-flash-lite",
    ]
    model_choice = st.selectbox(
        "選擇推理模型",
        options=model_options,
        index=0,
        help="Google 官方最新前瞻推理模型，預設推薦使用速度與推理兼備的 gemini-3.8-flash。"
    )
    
    st.divider()

    st.markdown("#### 🎨 試卷附圖格式")
    img_format_choice = st.radio(
        "選擇附圖嵌入規格",
        options=[
            "SVG 向量圖 (在 Word 點「轉換為圖形」可自由改字與線條)",
            "PNG 高解析圖 (300 DPI 印刷點陣圖，穩定度高)"
        ],
        index=0,
        help="• SVG 向量圖：微軟 Office 官方標準向量格式。在 Word 中點選圖片後，可於工具列點選「轉換為圖形 (Convert to Shape)」，即可將圖形拆解成獨立文字與線條，自由修改頂點英文字母與線段！\n• PNG：傳統 300 DPI 點陣圖。"
    )
    use_svg = "SVG" in img_format_choice
    
    st.divider()
    
    st.markdown("#### 📏 大考排版標準 (已鎖定)")
    st.caption("📄 **紙張規格**：JIS B4（257 × 364 mm）")
    st.caption("🖋️ **字型大小**：中文標楷體 13pt，英數 Times New Roman 13pt")
    st.caption("🔴 **詳解顏色**：深紅色 RGB(180, 0, 0)")
    st.caption("📐 **選項機制**：三層式無表格純段落智慧調配")
    st.caption("🔢 **方程式引擎**：Word 原生 OMML 可點選編輯公式")
    st.caption("⚡ **效能防護**：智慧算式脫殼分流（Word 全選複製極速不卡死）")

# 雙向同步至瀏覽器 localStorage（無 iframe 原生執行）
sync_js = f"""
<script>
(function() {{
    const STORAGE_KEY_CODE = "mathforge_activation_code";
    const STORAGE_KEY_API = "mathforge_gemini_api_key";
    
    const pyCode = "{activation_code}";
    const pyKey = "{api_key}";
    if (pyCode) localStorage.setItem(STORAGE_KEY_CODE, pyCode);
    if (pyKey) localStorage.setItem(STORAGE_KEY_API, pyKey);
    
    function syncLocalStorage() {{
        const inputs = document.querySelectorAll('input[type="password"]');
        if (inputs.length >= 2) {{
            const codeInput = inputs[0];
            const keyInput = inputs[1];
            
            codeInput.addEventListener("input", (e) => {{
                localStorage.setItem(STORAGE_KEY_CODE, e.target.value);
            }});
            keyInput.addEventListener("input", (e) => {{
                localStorage.setItem(STORAGE_KEY_API, e.target.value);
            }});
            
            const localCode = localStorage.getItem(STORAGE_KEY_CODE);
            const localKey = localStorage.getItem(STORAGE_KEY_API);
            const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
            
            if (localCode && !codeInput.value) {{
                setter.call(codeInput, localCode);
                codeInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                codeInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
            if (localKey && !keyInput.value) {{
                setter.call(keyInput, localKey);
                keyInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                keyInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
        }}
    }}
    setTimeout(syncLocalStorage, 300);
    setTimeout(syncLocalStorage, 800);
}})();
</script>
"""
st.html(sync_js, unsafe_allow_javascript=True)

# =======================================================
# 主畫面：頂部 Banner
# =======================================================
st.markdown("""
<div class="main-header">
    <div class="badge">✦ 宜蘭縣中華國中阿凱老師製作 ✦</div>
    <h1>阿凱老師的數學考卷改題排版神器</h1>
    <p>一鍵上傳試卷（DOCX / PDF），全自動素養改題、程式化精確繪圖，產出學校大考最高規格雙版本 Word 檔！</p>
</div>
""", unsafe_allow_html=True)

# 檔案上傳區
col_upload, col_info = st.columns([1.6, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "📂 上傳原始試卷 (.docx 或 .pdf)",
        type=["docx", "pdf"],
        help="支援標準 Word 考卷或 PDF 掃描檔/排版檔"
    )

with col_info:
    st.markdown("""
    <div class="feature-card">
        <h4 style="margin: 0 0 0.5rem 0; color: #1e3a8a;">✨ 系統自動交付成果</h4>
        <ul style="margin: 0; padding-left: 1.2rem; font-size: 0.93rem; color: #475569; line-height: 1.5;">
            <li><b>檔案 1：全卷試題與後附詳解(B4_13pt版)</b><br><small>題目在先，分頁後為簡答表與詳解</small></li>
            <li><b>檔案 2：逐題詳解教師備課卷(B4_13pt版)</b><br><small>題題對照，詳解文字與公式全紅字</small></li>
            <li><b>全卷公式與文字統一 13pt</b><br><small>Word 原生 OMML 公式注入 13pt，杜絕公式縮小落差</small></li>
            <li><b>現代 Office 模式（無相容模式）</b><br><small>原生解鎖向量圖形與現代方程式工具</small></li>
            <li><b>300 DPI 向量級精準附圖</b><br><small>坐標無壓線、雙箭頭標準標示</small></li>
        </ul>
    </div>
    <div class="feature-card" style="border-left: 4px solid #2563eb; background: #f8fafc;">
        <h4 style="margin: 0 0 0.5rem 0; color: #1e3a8a;">💡 老師改題必讀小秘訣</h4>
        <ul style="margin: 0; padding-left: 1.2rem; font-size: 0.90rem; color: #334155; line-height: 1.5;">
            <li><b>🎯 單次改題以不超過 20 題最佳！若超過 20 題建議分兩次上傳，AI 輸出最完整、不中斷且速度最快。</li>
            <li><b>🌟 上傳「詳解卷/解答卷」效果最佳</b>：原卷若附帶答案或解法，AI 能 100% 洞悉測驗重點，改寫出的生活素養情境最貼切、推導零失誤！</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

# 輔助函式：提取上傳檔案純文字
def extract_file_content(file_obj):
    filename = file_obj.name.lower()
    if filename.endswith(".docx"):
        doc = docx.Document(file_obj)
        full_text = []
        for p in doc.paragraphs:
            if p.text.strip():
                full_text.append(p.text)
        for table in doc.tables:
            for row in table.rows:
                row_txt = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
                if row_txt:
                    full_text.append(row_txt)
        return "\n".join(full_text)
    elif filename.endswith(".pdf"):
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(stream=file_obj.read(), filetype="pdf")
            full_text = []
            for page in doc:
                full_text.append(page.get_text())
            return "\n".join(full_text)
        except ImportError:
            import pypdf
            reader = pypdf.PdfReader(file_obj)
            full_text = [page.extract_text() for page in reader.pages if page.extract_text()]
            return "\n".join(full_text)
    return ""

# 單次調用 AI 模型輔助函式
def call_single_model_attempt(api_key: str, model_name: str, system_prompt: str, user_prompt: str) -> str:
    full_prompt = system_prompt + "\n\n" + user_prompt
    resp_text = ""
    
    # 優先調用 Google 官方最新 SDK (from google import genai)
    try:
        from google import genai
        from google.genai import types
        
        client = genai.Client(api_key=api_key)
        
        # 1. 優先調用 models.generate_content (官方生產標準模式，速度最快、穩定度最高)
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=user_prompt,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    response_mime_type="application/json"
                )
            )
            if response.text:
                resp_text = response.text.strip()
        except Exception:
            pass

        # 2. 次要嘗試 Interactions API
        if not resp_text:
            try:
                interaction = client.interactions.create(
                    model=model_name,
                    input=full_prompt,
                    response_format={
                        "type": "text",
                        "mime_type": "application/json"
                    }
                )
                if hasattr(interaction, "output_text") and interaction.output_text:
                    resp_text = interaction.output_text
                elif hasattr(interaction, "outputs") and interaction.outputs:
                    extracted = []
                    for out in interaction.outputs:
                        if hasattr(out, "text") and out.text:
                            extracted.append(out.text)
                        elif hasattr(out, "content") and out.content:
                            extracted.append(str(out.content))
                    resp_text = "\n".join(extracted)
            except Exception:
                pass
            
    except Exception:
        # 3. 兼容 legacy google.generativeai
        import google.generativeai as legacy_genai
        legacy_genai.configure(api_key=api_key)
        legacy_model = legacy_genai.GenerativeModel(
            model_name=model_name,
            generation_config={"response_mime_type": "application/json"}
        )
        response = legacy_model.generate_content([
            {"role": "user", "parts": [full_prompt]}
        ])
        resp_text = response.text.strip()
        
    return resp_text

# =======================================================
# 執行改題流程
# =======================================================
if uploaded_file is not None:
    st.markdown("---")
    btn_start = st.button("🚀 開始智能改題與 B4 考卷排版", type="primary", use_container_width=True)
    
    if btn_start:
        if activation_code.strip() != "kai":
            st.error("🔒 請先在左側欄輸入正確的「系統啟動碼」才能執行！")
            st.stop()

        if not api_key:
            st.error("❌ 請先在左側欄輸入您的 Google Gemini API Key！")
            st.stop()
            
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # 1. 讀取並解析原始考卷
            status_text.info("【1/4】📄 正在讀取並解析原始試卷文字與題型架構...")
            progress_bar.progress(15)
            exam_raw_text = extract_file_content(uploaded_file)
            if not exam_raw_text.strip():
                st.error("無法從上傳檔案中提取有效文字，請確認檔案格式是否正確。")
                st.stop()

            # 2. 測試 API 連線並呼叫 Gemini 進行改題
            status_text.info(f"【2/4】🌐 正在與 Google 伺服器握手連線，驗證 API Key 與模型「{model_choice}」...")
            progress_bar.progress(20)
            
            conn_ok = False
            try:
                from google import genai
                test_client = genai.Client(api_key=api_key)
                try:
                    test_client.models.get(model=model_choice)
                    conn_ok = True
                except Exception:
                    test_client.models.generate_content(
                        model=model_choice,
                        contents="ping",
                        config=genai.types.GenerateContentConfig(max_output_tokens=2)
                    )
                    conn_ok = True
            except Exception:
                try:
                    import google.generativeai as legacy_genai
                    legacy_genai.configure(api_key=api_key)
                    legacy_genai.get_model(f"models/{model_choice}")
                    conn_ok = True
                except Exception:
                    pass

            if conn_ok:
                status_text.success(f"【2/4】🌐 ✅ 已成功連線至 Google 伺服器！模型「{model_choice}」握手成功，即刻展開數學命題...")
                time.sleep(1.2)
            else:
                status_text.info(f"【2/4】🌐 已向 Google 伺服器送出連線請求，即刻展開數學命題...")
                time.sleep(0.8)

            candidate_models = [model_choice] + [m for m in model_options if m != model_choice]
            exam_data = None
            successful_model = None
            
            user_prompt = f"""請針對以下原始考卷內容進行「全卷改題」，嚴格依據系統規範產出純 JSON 格式：\n\n{exam_raw_text}"""
            
            for m_idx, curr_model in enumerate(candidate_models):
                progress_bar.progress(20 + int(30 * (m_idx / len(candidate_models))))
                start_t = time.time()
                error_msg = ""
                raw_text = ""
                
                with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
                    future = executor.submit(call_single_model_attempt, api_key, curr_model, SYSTEM_PROMPT, user_prompt)
                    
                    while True:
                        try:
                            raw_text = future.result(timeout=0.6)
                            break
                        except concurrent.futures.TimeoutError:
                            elapsed = int(time.time() - start_t)
                            status_text.info(f"【2/4】🤖 正由「{curr_model}」深入演算全卷試題與詳解...（已耗時 {elapsed} 秒，大考題目深度生成中，請耐心稍候）")
                        except Exception as req_err:
                            error_msg = str(req_err)
                            break
                
                if error_msg:
                    # 辨識是否為 429 額度耗盡或速率限制
                    is_429 = ("429" in error_msg) or ("quota" in error_msg.lower()) or ("rate" in error_msg.lower())
                    next_name = candidate_models[m_idx + 1] if m_idx + 1 < len(candidate_models) else "無"
                    if is_429:
                        st.warning(f"⚠️ 模型 **{curr_model}** 觸發 Google 免費額度上限 (429 Rate Limit)，正自動為您切換至 `{next_name}` 繼續嘗試！")
                    else:
                        st.warning(f"⚠️ 模型 **{curr_model}** 無法連線或呼叫異常（{error_msg[:80]}），自動切換至 `{next_name}` 繼續...")
                    time.sleep(1)
                    continue
                    
                if raw_text:
                    try:
                        clean_json = re.sub(r'^```json\s*', '', raw_text.strip())
                        clean_json = re.sub(r'\s*```$', '', clean_json)
                        parsed = json.loads(clean_json)
                        if parsed.get("questions"):
                            exam_data = parsed
                            successful_model = curr_model
                            total_sec = round(time.time() - start_t, 1)
                            status_text.success(f"✅ 模型「{curr_model}」改題成功！共耗時 {total_sec} 秒。")
                            time.sleep(1)
                            break
                        else:
                            st.warning(f"⚠️ 模型 **{curr_model}** 回傳題目為空，嘗試下一個模型...")
                    except Exception:
                        st.warning(f"⚠️ 模型 **{curr_model}** 回傳格式需校正，嘗試下一個模型...")
                        time.sleep(1)
                        
            if not exam_data:
                st.error("❌ 所有 AI 模型皆嘗試完畢，可能因 API Key 免費額度暫時用盡或連線問題。建議稍等 1 分鐘讓額度恢復後再試！")
                st.stop()
                
            unit_title = exam_data.get("unit_title", "國中數學段考試卷（改題版）")
            subtitle = exam_data.get("subtitle", "範圍：數學科段考  ｜  班級：______ 座號：___ 姓名：__________")
            questions = exam_data.get("questions", [])
            
            # 3. 執行程式化繪圖（即時顯示每題繪圖進度）
            progress_bar.progress(55)
            work_dir = tempfile.mkdtemp(prefix="mathforge_")
            images_dir = os.path.join(work_dir, "images")
            os.makedirs(images_dir, exist_ok=True)
            
            plots_needed = [q for q in questions if q.get("needs_plot") and q.get("plot_code")]
            svg_files_dict = {}
            if plots_needed:
                for p_idx, q in enumerate(plots_needed, 1):
                    mode_label = "SVG 向量圖 (含 300 DPI 備用圖)" if use_svg else "300 DPI 印刷高解析圖"
                    status_text.info(f"【3/4】📐 正在使用 Python Matplotlib 繪製第 {q['num']} 題 {mode_label} ({p_idx}/{len(plots_needed)})...")
                    progress_bar.progress(55 + int(20 * p_idx / len(plots_needed)))
                    
                    base_name = f"Q{q['num']}_adapted"
                    img_path = os.path.join(images_dir, f"{base_name}.png")
                    svg_path = os.path.join(images_dir, f"{base_name}.svg")
                    
                    orig_fig_savefig = Figure.savefig
                    orig_plt_savefig = plt.savefig
                    
                    def custom_fig_savefig(self, fname, *args, **kwargs):
                        res = orig_fig_savefig(self, fname, *args, **kwargs)
                        try:
                            kwargs_svg = kwargs.copy()
                            kwargs_svg.pop('dpi', None)
                            orig_fig_savefig(self, svg_path, format='svg', *args, **kwargs_svg)
                        except Exception as svg_err:
                            print(f"SVG generation warning for Q{q['num']}: {svg_err}")
                        return res
                        
                    local_scope = {"save_path": img_path, "plt": plt, "np": np}
                    try:
                        Figure.savefig = custom_fig_savefig
                        plt.savefig = lambda fname, *args, **kwargs: custom_fig_savefig(plt.gcf(), fname, *args, **kwargs)
                        exec(q["plot_code"], {}, local_scope)
                        # 防護：若代碼忘記調用 savefig，但目前有開啟之圖表
                        if plt.get_fignums():
                            if not os.path.exists(img_path):
                                orig_fig_savefig(plt.gcf(), img_path, dpi=300, bbox_inches='tight', pad_inches=0.15)
                            if not os.path.exists(svg_path):
                                orig_fig_savefig(plt.gcf(), svg_path, format='svg', bbox_inches='tight', pad_inches=0.15)
                    except Exception as plot_err:
                        print(f"Plotting error for Q{q['num']}: {plot_err}")
                    finally:
                        Figure.savefig = orig_fig_savefig
                        plt.savefig = orig_plt_savefig
                        plt.close('all')

                    if os.path.exists(img_path):
                        q["image_path"] = img_path
                        q["img_width_inch"] = 3.2
                    if os.path.exists(svg_path):
                        with open(svg_path, "rb") as f_svg:
                            svg_bytes = f_svg.read()
                        q["svg_bytes"] = svg_bytes
                        svg_files_dict[f"第{q['num']}題_附圖.svg"] = svg_bytes
                        if use_svg:
                            q["svg_path"] = svg_path
                        else:
                            q["svg_path"] = None
            else:
                status_text.info("【3/4】📐 本卷題目無須額外繪圖，直接進入排版階段...")
                progress_bar.progress(75)

            # 4. 生成雙版本 Word 考卷（細緻拆分雙卷生成狀態）
            status_text.info("【4/4】📄 正在套用 JIS B4 與 OMML 公式引擎：生成「學生/分離卷」...")
            progress_bar.progress(82)
            
            file1_name = f"{unit_title}_全卷試題與後附詳解(B4_13pt版).docx"
            file2_name = f"{unit_title}_逐題詳解教師備課卷(B4_13pt版).docx"
            
            path1 = os.path.join(work_dir, file1_name)
            path2 = os.path.join(work_dir, file2_name)
            
            build_student_exam(questions, unit_title, subtitle, path1)
            
            status_text.info("【4/4】📄 正在套用 JIS B4 與 OMML 公式引擎：生成「教師/全紅字詳解備課卷」...")
            progress_bar.progress(93)
            build_teacher_exam(questions, unit_title, subtitle, path2)
            
            progress_bar.progress(100)
            status_text.empty()
            
            # 讀取生成之 Word 檔至記憶體，存入 session_state 持久化
            with open(path1, "rb") as f1:
                f1_bytes = f1.read()
            with open(path2, "rb") as f2:
                f2_bytes = f2.read()
                
            svg_zip_bytes = None
            svg_zip_name = None
            if svg_files_dict:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for s_fname, s_data in svg_files_dict.items():
                        zf.writestr(s_fname, s_data)
                svg_zip_bytes = zip_buffer.getvalue()
                svg_zip_name = f"{unit_title}_全卷附圖SVG向量包.zip"

            st.session_state["download_data"] = {
                "file1_name": file1_name,
                "file1_bytes": f1_bytes,
                "file2_name": file2_name,
                "file2_bytes": f2_bytes,
                "svg_zip_name": svg_zip_name,
                "svg_zip_bytes": svg_zip_bytes,
                "questions": questions,
                "use_svg": use_svg,
            }
            
        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"❌ 執行過程中發生錯誤：{str(e)}")
            with st.expander("查看詳細錯誤追蹤 (Error Traceback)"):
                st.code(traceback.format_exc())

# =======================================================
# 成果展示與持久化下載專區（保證下載真實 .docx 檔案）
# =======================================================
if "download_data" in st.session_state:
    data_dict = st.session_state["download_data"]
    
    st.markdown("---")
    st.markdown("""
    <div class="success-box">
        <h3 style="margin: 0 0 0.4rem 0; color: #065f46;">🎉 恭喜！改題與 B4 排版已全部圓滿完成！</h3>
        <p style="margin: 0; color: #047857;">已為您產出臺灣學校大考最高規格之雙版本 Word 檔（內含 Word 原生可點選編輯公式與三層式選項調配）。</p>
    </div>
    """, unsafe_allow_html=True)
    
    # 下載按鈕專區（以標準記憶體二進位位元組流下載，100% 確保檔名與 .docx 格式）
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            label="📥 下載：全卷試題與後附詳解 (學生/分離卷 .docx)",
            data=data_dict["file1_bytes"],
            file_name=data_dict["file1_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    with c2:
        st.download_button(
            label="📥 下載：逐題詳解教師備課卷 (教師/對照卷 .docx)",
            data=data_dict["file2_bytes"],
            file_name=data_dict["file2_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
        
    if data_dict.get("svg_zip_bytes"):
        st.download_button(
            label="📦 下載：全卷題目附圖 SVG 向量圖包 (.zip)",
            data=data_dict["svg_zip_bytes"],
            file_name=data_dict["svg_zip_name"],
            mime="application/zip",
            help="包含本份試卷所有幾何附圖的獨立 SVG 原始向量檔案，可用 Inkscape、Illustrator 或 PowerPoint 自由編輯！",
            use_container_width=True
        )
        
    if data_dict.get("use_svg"):
        st.info("💡 **小撇步：如何在 Word 中直接修改幾何附圖？**\n\n"
                "1. 在 Word 中滑鼠點選題目附圖，上方功能區會出現 **【圖形格式 (Graphics Format)】**。\n"
                "2. 點選 **【轉換為圖形 (Convert to Shape)】**（或在圖上按右鍵點選此選項）。\n"
                "3. 整張圖形立即解構為微軟原生向量圖形與文字！\n"
                "4. 您可以任意點選頂點字母（如 A, B, C）直接改字、拖移坐標位置、或調整線條粗細與色彩！")
    
    # 預覽產生之試題
    st.markdown("### 🔍 改題成果即時預覽 (Preview)")
    for q in data_dict["questions"]:
        with st.expander(f"題號 {q['num']}：{q['stem'][:40]}..."):
            st.markdown(f"**題幹**：{q['stem']}")
            st.markdown("**選項**：")
            for opt in q["options"]:
                st.markdown(f"- {opt}")
            st.markdown(f"**答案**：`({q['ans']})`")
            st.markdown("**詳解**：")
            for step in q["explanation"]:
                st.markdown(f"> {step}")
            if q.get("image_path") and os.path.exists(q["image_path"]):
                st.image(q["image_path"], caption=f"第 {q['num']} 題附圖", width=350)
                if q.get("svg_bytes"):
                    st.download_button(
                        label=f"📥 下載第 {q['num']} 題 SVG 向量檔",
                        data=q["svg_bytes"],
                        file_name=f"第{q['num']}題_附圖.svg",
                        mime="image/svg+xml",
                        key=f"dl_svg_{q['num']}"
                    )
