"""
外贸跟单工单智能审核系统 - 主入口
基于 AI 大模型的外贸单据智能比对与审核工具

部署目标：Streamlit Community Cloud

v3.2 更新：
- 新增自定义审核规则功能（侧边栏可编辑，优先级高于内置规则）
- 前端界面全面美化（现代SaaS风格）
- 支持集团关联公司识别、COA匹配规则
- 内置规则透明展示（侧边栏只读查看）
"""

import streamlit as st
from utils.config_manager import (
    init_session_state,
    is_disclaimer_accepted,
    accept_disclaimer,
    is_disclaimer_skip,
    get_disclaimer_step,
    set_disclaimer_step,
    reset_disclaimer,
    get_selected_model,
    get_api_key,
    get_active_api_key,
    get_zhipu_ocr_api_key,
    is_deep_think_enabled,
    is_audit_cancelled,
    set_cancel_audit,
    get_token_warning,
    set_token_warning,
    get_custom_rules,
    KEY_CUSTOM_RULES,
    MODEL_OPTIONS,
    KEY_SELECTED_MODEL,
    KEY_API_KEY,
    KEY_DEEPSEEK_API_KEY,
    KEY_ZHIPU_API_KEY,
    KEY_ZHIPU_OCR_API_KEY,
    KEY_DEEP_THINK,
    KEY_CANCEL_AUDIT,
)
from utils.file_parser import parse_file, get_image_thumbnail
from utils.llm_client import (
    LLMError,
    call_llm_with_image,
    test_connection,
    IMAGE_OCR_PROMPT,
)
from utils.audit_orchestrator import run_full_audit
from utils.report_generator import (
    generate_marked_report,
    generate_detail_report,
    generate_zip,
)
from utils.history_manager import (
    add_history_record,
    get_history_records,
    get_history_count,
)
from utils.rules_config import BUILTIN_RULES_DISPLAY

# 支持的文件类型列表（统一管理）
ALLOWED_DOC_TYPES = ["pdf", "doc", "docx", "xlsx", "xls"]

# ============================================================
# 页面基础配置
# ============================================================
st.set_page_config(
    page_title="外贸跟单工单智能审核系统",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="auto",
)

# 初始化 session_state
init_session_state()


# ============================================================
# 自定义样式（全面升级视觉标准）
# ============================================================
st.markdown(
    """
    <style>
    /* ---- 全局 ---- */
    .stApp {
        background: #FAFBFC;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif;
    }

    /* ---- 免责声明 ---- */
    .disclaimer-box {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 16px;
        padding: 40px 48px;
        margin: 40px auto;
        max-width: 640px;
        box-shadow: 0 8px 30px rgba(0,0,0,0.06);
    }
    .disclaimer-box h3 {
        text-align: center;
        color: #111827;
        margin-bottom: 24px;
        font-size: 24px;
        font-weight: 700;
        letter-spacing: 0;
    }
    .disclaimer-box p {
        color: #4B5563;
        line-height: 1.85;
        font-size: 15px;
    }

    /* ---- 二级确认 ---- */
    .confirm-box {
        background: #FEF2F2;
        border: 1px solid #FECACA;
        border-radius: 16px;
        padding: 32px 40px;
        margin: 40px auto;
        max-width: 600px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.04);
    }
    .confirm-box h3 {
        text-align: center;
        color: #DC2626;
        font-size: 20px;
        font-weight: 700;
    }
    .confirm-box p {
        color: #4B5563;
        line-height: 1.85;
        font-size: 15px;
    }

    /* ---- 页面标题 ---- */
    .main-title {
        font-size: 28px;
        font-weight: 700;
        color: #111827;
        margin-bottom: 4px;
        letter-spacing: -0.3px;
    }
    .main-subtitle {
        color: #9CA3AF;
        font-size: 14px;
        margin-top: -4px;
        margin-bottom: 32px;
        font-weight: 400;
    }

    /* ---- 上传区栏标题 ---- */
    .upload-col-title {
        font-size: 15px;
        font-weight: 600;
        margin-bottom: 12px;
        padding: 10px 16px;
        border-radius: 10px;
        color: #fff;
        letter-spacing: 0.3px;
    }
    .upload-col-title-left {
        background: #4F46E5;
    }
    .upload-col-title-right {
        background: #10B981;
    }

    /* ---- 提示文字 ---- */
    .hint-text {
        color: #9CA3AF;
        font-size: 13px;
        line-height: 1.6;
    }

    /* ---- 空状态 ---- */
    .empty-state {
        text-align: center;
        padding: 48px 24px;
        color: #9CA3AF;
        font-size: 15px;
        border: 2px dashed #D1D5DB;
        border-radius: 12px;
        margin: 24px 0;
        background: #FFFFFF;
    }
    .empty-state .icon {
        font-size: 40px;
        display: block;
        margin-bottom: 12px;
    }

    /* ---- 结果区标题 ---- */
    .result-section-title {
        font-size: 18px;
        font-weight: 700;
        color: #111827;
        margin-top: 16px;
        margin-bottom: 8px;
        padding-bottom: 8px;
        border-bottom: 2px solid #4F46E5;
        display: inline-block;
    }

    /* ---- 下载区域 ---- */
    .download-section {
        margin-top: 16px;
        padding-top: 16px;
        border-top: 1px dashed #E5E7EB;
    }
    .download-area {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 12px;
        padding: 24px 28px;
        margin-top: 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }

    /* ---- 统计卡片 ---- */
    .stat-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 12px;
        padding: 20px 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        transition: border-color 0.2s ease;
    }
    .stat-card:hover {
        border-color: #C7D2FE;
    }
    .stat-card .label {
        font-size: 13px;
        color: #9CA3AF;
        margin-bottom: 6px;
        font-weight: 500;
    }
    .stat-card .value {
        font-size: 32px;
        font-weight: 700;
        letter-spacing: -1px;
    }

    /* ---- 问题卡片 ---- */
    .issue-card {
        border-radius: 10px;
        padding: 14px 18px;
        margin: 10px 0;
        transition: box-shadow 0.15s ease;
    }
    .issue-card:hover {
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }

    /* ---- 历史记录卡片 ---- */
    .history-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 12px 16px;
        margin: 8px 0;
        transition: border-color 0.2s ease;
    }
    .history-card:hover {
        border-color: #4F46E5;
    }

    /* ---- 隐藏默认元素 ---- */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* ---- 按钮 ---- */
    .stButton > button {
        border-radius: 10px;
        font-weight: 600;
        padding: 8px 20px;
        transition: all 0.2s ease;
        border: none;
        letter-spacing: 0.3px;
    }
    .stButton > button:hover {
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 600;
        padding: 8px 20px;
        transition: all 0.2s ease;
        letter-spacing: 0.3px;
    }
    .stDownloadButton > button:hover {
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }

    /* ---- 侧边栏 ---- */
    section[data-testid="stSidebar"] {
        background: #1E1E2E;
    }
    section[data-testid="stSidebar"] * {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #F1F5F9 !important;
        font-size: 15px;
        font-weight: 600;
        letter-spacing: 0.3px;
    }
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] .hint-text {
        color: #94A3B8 !important;
    }
    section[data-testid="stSidebar"] .stDivider,
    section[data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,0.08) !important;
    }

    /* 侧边栏输入框 */
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] .stTextInput input,
    section[data-testid="stSidebar"] .stTextInput > div,
    section[data-testid="stSidebar"] .stTextInput > div > div,
    section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] .stTextArea textarea,
    section[data-testid="stSidebar"] .stTextArea > div {
        background-color: rgba(255,255,255,0.06) !important;
        color: #F1F5F9 !important;
        border-color: rgba(255,255,255,0.12) !important;
        caret-color: #E2E8F0 !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] .stTextInput input:focus,
    section[data-testid="stSidebar"] .stTextInput > div:focus-within,
    section[data-testid="stSidebar"] .stTextArea textarea:focus,
    section[data-testid="stSidebar"] .stTextArea > div:focus-within {
        border-color: #4F46E5 !important;
        box-shadow: 0 0 0 1px rgba(79,70,229,0.4) !important;
    }

    /* 侧边栏密码眼睛按钮 */
    section[data-testid="stSidebar"] .stTextInput button {
        background: transparent !important;
        color: #94A3B8 !important;
        border: none !important;
    }
    section[data-testid="stSidebar"] .stTextInput button:hover {
        color: #E2E8F0 !important;
        background: rgba(255,255,255,0.08) !important;
    }

    /* 侧边栏下拉框 */
    section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"],
    section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] > div {
        background-color: rgba(255,255,255,0.06) !important;
        color: #F1F5F9 !important;
        border-color: rgba(255,255,255,0.12) !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] * {
        color: #F1F5F9 !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="popover"],
    section[data-testid="stSidebar"] [data-baseweb="menu"],
    section[data-testid="stSidebar"] [role="listbox"],
    section[data-testid="stSidebar"] [role="option"] {
        background-color: #272741 !important;
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] [role="option"]:hover {
        background-color: rgba(79,70,229,0.25) !important;
    }

    /* 侧边栏 placeholder */
    section[data-testid="stSidebar"] input::placeholder,
    section[data-testid="stSidebar"] textarea::placeholder {
        color: #64748B !important;
    }

    /* 侧边栏按钮 */
    section[data-testid="stSidebar"] .stButton > button {
        background: rgba(79,70,229,0.15) !important;
        color: #A5B4FC !important;
        border: 1px solid rgba(79,70,229,0.3) !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background: rgba(79,70,229,0.3) !important;
        color: #C7D2FE !important;
    }

    /* 侧边栏 toggle / checkbox / expander */
    section[data-testid="stSidebar"] .stToggle label span,
    section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] .streamlit-expanderHeader,
    section[data-testid="stSidebar"] details summary span {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] details > div {
        background: rgba(255,255,255,0.03) !important;
        border-radius: 8px;
    }
    section[data-testid="stSidebar"] .stCheckbox label span {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] .stAlert * {
        color: inherit !important;
    }

    /* ---- Expander ---- */
    .streamlit-expanderHeader {
        font-weight: 600;
        font-size: 15px;
    }

    /* ---- 文件上传 ---- */
    .stFileUploader > div {
        border-radius: 10px;
    }

    /* ---- 分割线 ---- */
    hr {
        border: none;
        border-top: 1px solid #E5E7EB;
        margin: 20px 0;
    }

    /* ---- 版本标签 ---- */
    .version-tag {
        display: inline-block;
        background: #4F46E5;
        color: white;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 免责声明弹窗
# ============================================================
def show_disclaimer() -> None:
    """显示免责声明页面（阻断式，通过前不渲染主界面）。"""

    step = get_disclaimer_step()

    # ---------- 阶段：二级确认 ----------
    if step == "confirming":
        st.markdown("&nbsp;", unsafe_allow_html=True)
        st.markdown(
            """
            <div class="confirm-box">
            <h3 style="text-align:center; color:#cf1322;">二次确认</h3>
            <p>
            您选择了本次会话不再显示免责声明。请再次确认：您已充分理解
            本程序的审核结果仅供参考，使用后仍需进行人工复核。确定不再
            提示吗？
            </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ★ 改动：用 columns 居中
        _, center_col, _ = st.columns([1, 2, 1])
        with center_col:
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("确认，不再提示", use_container_width=True, type="primary"):
                    accept_disclaimer()
                    st.rerun()
            with col_b:
                if st.button("取消，保留提示", use_container_width=True):
                    set_disclaimer_step("initial")
                    st.rerun()

        st.stop()

    # ---------- 阶段：初始免责声明 ----------
    st.markdown("&nbsp;", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="disclaimer-box">
        <h3>免责声明</h3>
        <p>
        本程序基于AI大模型进行辅助审核，旨在帮助您发现单据中可能存在
        的数据不一致或疏漏之处。但AI审核存在局限性，无法保证100%准确，
        审核结果仅供参考，不能替代人工复核。
        </p>
        <p>
        请您在使用本程序后，务必对审核报告中的标记项逐一确认，并结合
        实际业务情况做出最终判断。因未进行人工复核而导致的任何错误或
        损失，本程序不承担相关责任。
        </p>
        <p style="font-weight:600;">
        使用本程序即表示您已知悉并同意以上内容。
        </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ★ 改动：用 columns 居中
    _, center_col, _ = st.columns([1, 2, 1])
    with center_col:
        skip = st.checkbox("本次会话不再提示", key="_disclaimer_skip_cb")

        if st.button("我已知晓，进入系统", type="primary", use_container_width=True):
            if skip:
                set_disclaimer_step("confirming")
                st.rerun()
            else:
                accept_disclaimer()
                st.rerun()

    st.stop()


# ============================================================
# 审核规则弹窗
# ============================================================
@st.dialog("📖 系统内置审核规则", width="large")
def show_builtin_rules_dialog():
    """以弹窗形式展示内置审核规则（只读）。"""
    st.markdown(BUILTIN_RULES_DISPLAY)
    st.markdown("---")
    st.markdown(
        '<p class="hint-text">'
        '以上为系统内置规则摘要。AI 实际审核时会使用更详细的执行指令，'
        '核心逻辑与上述规则一致。如需调整审核行为，请使用「编辑自定义规则」功能。'
        '</p>',
        unsafe_allow_html=True,
    )
    if st.button("关闭", use_container_width=True):
        st.rerun()


@st.dialog("✏️ 自定义补充审核规则", width="large")
def show_custom_rules_dialog():
    """以弹窗形式编辑自定义审核规则，有明确的保存/取消按钮。"""
    st.markdown(
        "**自定义规则的优先级高于系统内置规则**，当两者冲突时以自定义规则为准。\n\n"
        "修改后请点击下方「保存规则」按钮使其生效。留空则不添加额外规则。"
    )
    st.markdown("---")

    # 读取当前已保存的自定义规则作为默认值
    current_rules = get_custom_rules() or ""
    new_rules = st.text_area(
        "请输入自定义审核规则",
        value=current_rules,
        height=400,
        placeholder=(
            "例如：\n"
            "- 所有金额必须精确到小数点后两位\n"
            "- 忽略发货地址差异\n"
            "- 本票货物为危险品，需要特别检查UN编号\n"
            "- 客户要求所有单据上必须显示HS Code"
        ),
        key="_temp_custom_rules_editor",
    )

    st.markdown("---")
    col_save, col_cancel = st.columns(2)
    with col_save:
        if st.button("💾 保存规则", type="primary", use_container_width=True):
            st.session_state[KEY_CUSTOM_RULES] = new_rules
            st.toast("✅ 自定义规则已保存并立即生效！")
            import time
            time.sleep(0.8)
            st.rerun()
    with col_cancel:
        if st.button("取消", use_container_width=True):
            st.rerun()


# ============================================================
# 侧边栏
# ============================================================
def render_sidebar() -> None:
    """渲染侧边栏设置区域。"""
    with st.sidebar:
        st.markdown("### ⚙️ 系统设置")
        st.markdown(
            '<p class="hint-text">提示：密钥仅在当前会话有效，刷新页面后需重新输入</p>',
            unsafe_allow_html=True,
        )

        st.divider()

        # ----- 大模型选择 -----
        current_idx = (
            MODEL_OPTIONS.index(get_selected_model())
            if get_selected_model() in MODEL_OPTIONS
            else 0
        )
        st.selectbox(
            "🤖 大模型选择",
            options=MODEL_OPTIONS,
            index=current_idx,
            key=KEY_SELECTED_MODEL,
        )

        # ----- API 密钥输入（每个模型独立存储，互不覆盖）-----
        selected_model = get_selected_model()

        if selected_model == "DeepSeek":
            st.text_input(
                "🔑 DeepSeek API 密钥",
                type="password",
                placeholder="请输入 DeepSeek API 密钥",
                key=KEY_DEEPSEEK_API_KEY,
            )
            st.text_input(
                "🖼️ 智谱OCR密钥（扫描件识别用，选填）",
                type="password",
                placeholder="如需OCR扫描件，请输入智谱API密钥",
                key=KEY_ZHIPU_OCR_API_KEY,
            )
            st.markdown(
                '<p class="hint-text">💡 扫描件PDF需要智谱GLM进行OCR识别。如不需要OCR可留空。</p>',
                unsafe_allow_html=True,
            )
        elif selected_model == "智谱GLM":
            st.text_input(
                "🔑 智谱GLM API 密钥",
                type="password",
                placeholder="请输入智谱 API 密钥",
                key=KEY_ZHIPU_API_KEY,
            )

        # ----- DeepSeek 深度思考模式开关（仅 DeepSeek 时显示）-----
        if get_selected_model() == "DeepSeek":
            st.divider()
            st.toggle(
                "🧠 深度思考模式",
                key=KEY_DEEP_THINK,
                help="开启后使用 DeepSeek Reasoner 模型，审核更准确但速度较慢且消耗更多 token",
            )
            if is_deep_think_enabled():
                st.markdown(
                    '<p class="hint-text">⚡ 深度思考模式已开启：审核更准确，但速度较慢且消耗更多 token。'
                    '超时时间已自动延长至 300 秒。</p>',
                    unsafe_allow_html=True,
                )

        # ----- 测试连接按钮 -----
        st.divider()
        if st.button("🔗 测试连接", use_container_width=True):
            _handle_test_connection()

        st.divider()

        # ----- 审核规则区域 -----
        st.markdown("### 📋 审核规则")

        if st.button("📖 查看系统内置审核规则", use_container_width=True):
            show_builtin_rules_dialog()

        if st.button("✏️ 编辑自定义规则", use_container_width=True):
            show_custom_rules_dialog()

        # 显示当前自定义规则的状态
        current_custom = get_custom_rules()
        if current_custom and current_custom.strip():
            preview_text = current_custom.strip()[:50]
            if len(current_custom.strip()) > 50:
                preview_text += "..."
            st.success("✅ 已配置自定义规则")
            st.markdown(
                f'<p class="hint-text">当前规则预览：{preview_text}</p>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<p class="hint-text">未设置自定义规则，将使用系统默认规则</p>',
                unsafe_allow_html=True,
            )

        st.divider()

        # ----- 审核历史记录 -----
        _render_sidebar_history()

        st.divider()

        # ----- 重新显示免责声明 -----
        if st.button("📜 重新显示免责声明", use_container_width=True):
            reset_disclaimer()
            st.rerun()

        st.divider()

        st.markdown(
            '<p class="hint-text" style="text-align:center;">'
            '<span class="version-tag">v3.2</span><br/><br/>'
            'AI辅助审核，仅供参考<br/>'
            '支持 PDF / Word / Excel / 图片输入<br/>'
            '统一输出 Excel 格式报告</p>',
            unsafe_allow_html=True,
        )


# ============================================================
# 侧边栏：审核历史记录
# ============================================================
def _render_sidebar_history() -> None:
    """在侧边栏渲染审核历史记录区域。"""
    history_count = get_history_count()
    st.markdown(f"### 📜 审核历史 ({history_count})")
    st.markdown(
        '<p class="hint-text">历史记录仅在当前会话有效，刷新页面后将清空</p>',
        unsafe_allow_html=True,
    )

    if history_count == 0:
        st.markdown(
            '<p class="hint-text" style="text-align:center;">暂无审核记录</p>',
            unsafe_allow_html=True,
        )
        return

    records = get_history_records()
    for record in records:
        record_id = record["id"]
        timestamp = record["timestamp"]
        file_names = record["file_names"]
        total_red = record["total_red"]
        total_yellow = record["total_yellow"]
        total_blue = record["total_blue"]

        files_display = "、".join(file_names[:3])
        if len(file_names) > 3:
            files_display += f" 等{len(file_names)}份"

        badge = ""
        if total_red > 0:
            badge = f"🔴{total_red}"
        if total_yellow > 0:
            badge += f" 🟡{total_yellow}"
        if total_blue > 0:
            badge += f" 🔵{total_blue}"
        if not badge:
            badge = "✅ 无问题"

        with st.expander(f"#{record_id} {timestamp}", expanded=False):
            st.markdown(f"**文件：**{files_display}")
            st.markdown(f"**标记：**{badge.strip()}")
            if st.button(
                "📋 查看详情",
                key=f"history_view_{record_id}",
                use_container_width=True,
            ):
                st.session_state["audit_result"] = record["audit_result"]
                st.session_state["viewing_history"] = record_id
                st.rerun()


# ============================================================
# 测试连接处理
# ============================================================
def _handle_test_connection() -> None:
    """处理侧边栏的测试连接按钮点击。"""
    api_key = get_active_api_key()
    provider = get_selected_model()

    if not api_key or not api_key.strip():
        st.sidebar.error("请先输入API密钥")
        return

    with st.sidebar:
        with st.spinner("正在测试连接..."):
            try:
                reply = test_connection(provider, api_key)
                st.success("✅ 连接成功")
            except LLMError as e:
                st.error(f"连接失败：{e.message}")
            except Exception:
                st.error("连接失败：网络异常或服务不可用，请稍后重试")


# ============================================================
# 文件解析预览辅助
# ============================================================
@st.cache_data(show_spinner=False)
def _cached_parse_file(_file_content: bytes, filename: str) -> dict:
    """缓存文件解析结果，避免每次 rerun 重复解析。"""
    import io
    buf = io.BytesIO(_file_content)
    buf.name = filename
    return parse_file(buf)


def _render_single_preview(uploaded_file) -> None:
    """为单个上传文件渲染解析预览 expander。"""
    if uploaded_file is None:
        return

    try:
        file_bytes = uploaded_file.getvalue()
        result = _cached_parse_file(file_bytes, uploaded_file.name)
    except Exception:
        st.warning(f"文件 {getattr(uploaded_file, 'name', '未知')} 解析异常，请检查文件是否完整")
        return

    # ★ 新增：扫描件PDF特殊预览
    if result.get("is_scanned_pdf"):
        filename = result["filename"]
        page_images = result.get("pdf_page_images", [])
        page_count = len(page_images)
        with st.expander(f"📕 {filename} 解析预览", expanded=False):
            st.info(
                f"📷 检测到该PDF为扫描件（共 {page_count} 页），"
                f"文字内容将在点击「开始审核」后通过 AI-OCR 自动识别。\n\n"
                f"请确保已在左侧边栏填写「智谱OCR密钥」（如使用DeepSeek模型）"
                f"或选择「智谱GLM」模型。"
            )
            # 显示第一页缩略图让用户确认文件上传正确
            if page_images:
                import base64 as _b64
                try:
                    first_page_bytes = _b64.b64decode(page_images[0])
                    st.image(
                        first_page_bytes,
                        caption=f"第 1 页预览（共 {page_count} 页）",
                        use_container_width=True,
                    )
                except Exception:
                    st.caption("（缩略图生成失败，不影响审核功能）")
        return  # 提前返回，不走后面的普通预览逻辑

    filename = result["filename"]
    is_image = result["is_image"]
    success = result["success"]
    content = result["content"]
    file_type = result["type"]

    if is_image and success:
        with st.expander(f"🖼️ {filename} 解析预览"):
            thumb = get_image_thumbnail(uploaded_file)
            if thumb is not None:
                st.image(thumb, caption=filename, use_container_width=False)
            st.info("图片内容将在审核时由AI识别")
            if result["image_base64"]:
                preview = result["image_base64"][:80] + "..."
                st.markdown(
                    f'<p class="hint-text">Base64 编码预览: {preview}</p>',
                    unsafe_allow_html=True,
                )
    elif success:
        icon = "📄"
        if file_type == "pdf":
            icon = "📕"
        elif file_type == "docx":
            icon = "📘"
        elif file_type == "xlsx":
            icon = "📗"

        with st.expander(f"{icon} {filename} 解析预览"):
            if len(content) > 5000:
                st.text(content[:5000])
                st.info(f"文件内容较长（共 {len(content)} 字符），此处仅显示前 5000 字符")
            else:
                st.text(content)
    else:
        with st.expander(f"⚠️ {filename} 解析预览"):
            if "为空" in content:
                st.warning("该文件内容为空，无法提取有效信息")
            elif "损坏" in content or "解析失败" in content:
                st.error("该文件可能已损坏或格式不正确，请更换文件后重新上传")
            elif "旧版" in content:
                st.warning("该文件为旧版 .doc 格式，建议另存为 .docx 后重新上传")
            else:
                st.error("该文件解析失败，请检查文件是否损坏或格式是否正确")
            if content:
                st.text(content)


def _render_file_preview(uploaded, *, multi: bool) -> None:
    """为上传控件渲染解析预览。"""
    if uploaded is None:
        return
    if multi:
        if not uploaded:
            return
        for f in uploaded:
            _render_single_preview(f)
    else:
        _render_single_preview(uploaded)


# ============================================================
# 输入校验（统一中文友好提示）
# ============================================================
def _validate_audit_inputs(po_file, audit_files, result_placeholder) -> bool:
    """校验审核启动的所有前置条件，失败返回 False。"""
    api_key = get_active_api_key()

    if not api_key or not api_key.strip():
        result_placeholder.error("请先在左侧边栏中配置API密钥，然后再开始审核")
        return False

    if po_file is None:
        result_placeholder.error("请先上传PO文件（审核核心依据）")
        return False

    if not audit_files:
        result_placeholder.warning("请在右侧上传至少一份待审核文件（如CI、PL等）")
        return False

    po_size = getattr(po_file, 'size', 0)
    if po_size > 50 * 1024 * 1024:
        result_placeholder.error("PO文件过大（超过50MB），请压缩后重新上传")
        return False

    for f in audit_files:
        fsize = getattr(f, 'size', 0)
        if fsize > 50 * 1024 * 1024:
            result_placeholder.error(
                f"文件 {getattr(f, 'name', '未知')} 过大（超过50MB），请压缩后重新上传"
            )
            return False

    return True


def _validate_po_data(po_data: dict, result_placeholder) -> bool:
    """校验PO解析结果。"""
    if not po_data.get("success"):
        po_content = po_data.get("content", "")
        if "为空" in po_content:
            result_placeholder.error("PO文件内容为空，请检查文件是否正确或重新上传")
        elif "旧版" in po_content:
            result_placeholder.error("PO文件为旧版 .doc 格式，请用Word或WPS另存为 .docx 后重新上传")
        elif "解析失败" in po_content:
            result_placeholder.error("PO文件解析失败，文件可能已损坏，请更换文件后重新上传")
        else:
            result_placeholder.error("PO文件解析失败，请检查文件是否完整且格式正确")
        return False
    return True


# ============================================================
# 审核启动处理（完整审核流程）
# ============================================================
def _handle_audit_start(
    po_file,
    template_file,
    prev_files,
    ref_images,
    audit_files,
    result_placeholder,
) -> None:
    """点击开始审核后，执行完整审核流程。"""
    api_key = get_active_api_key()
    provider = get_selected_model()
    deep_think = is_deep_think_enabled() and provider == "DeepSeek"

    # 重置取消标志
    set_cancel_audit(False)

    # --- 前置校验 ---
    if not _validate_audit_inputs(po_file, audit_files, result_placeholder):
        return

    # --- 解析所有上传文件 ---
    try:
        po_data = parse_file(po_file)
    except Exception:
        result_placeholder.error("PO文件读取异常，文件可能已损坏，请重新上传")
        return

    if not _validate_po_data(po_data, result_placeholder):
        return

    template_data = None
    if template_file:
        try:
            template_data = parse_file(template_file)
        except Exception:
            result_placeholder.warning("模板文件解析异常，将跳过模板比对继续审核")

    last_ticket_data = None
    if prev_files:
        last_ticket_data = []
        for f in prev_files:
            try:
                last_ticket_data.append(parse_file(f))
            except Exception:
                result_placeholder.warning(
                    f"上一票文件 {getattr(f, 'name', '未知')} 解析异常，已跳过"
                )

    other_refs_data = None
    if ref_images:
        other_refs_data = []
        for f in ref_images:
            try:
                other_refs_data.append(parse_file(f))
            except Exception:
                result_placeholder.warning(
                    f"参考图片 {getattr(f, 'name', '未知')} 读取异常，已跳过"
                )

    target_files_data = []
    for f in audit_files:
        try:
            target_files_data.append(parse_file(f))
        except Exception:
            result_placeholder.warning(
                f"待审核文件 {getattr(f, 'name', '未知')} 解析异常，已跳过"
            )

    if not target_files_data:
        result_placeholder.error("所有待审核文件解析失败，请检查文件后重新上传")
        return

    # --- 执行审核 ---
    with result_placeholder.container():
        mode_label = "（深度思考模式）" if deep_think else ""
        status_container = st.status(
            f"正在准备审核...{mode_label}",
            expanded=True,
        )

        # 取消审核按钮
        cancel_col1, cancel_col2, cancel_col3 = st.columns([2, 1, 2])
        with cancel_col2:
            cancel_btn = st.button(
                "⏹️ 取消审核",
                use_container_width=True,
                key="cancel_audit_btn",
            )
            if cancel_btn:
                set_cancel_audit(True)

        step_counter = [0]

        def on_progress(msg: str):
            step_counter[0] += 1
            status_container.update(label=f"🔍 {msg}", state="running")
            status_container.write(f"✅ 步骤 {step_counter[0]}：{msg}")

        def check_cancel() -> bool:
            return is_audit_cancelled()

        try:
            audit_result = run_full_audit(
                provider=provider,
                api_key=api_key,
                po_data=po_data,
                target_files_data=target_files_data,
                last_ticket_data=last_ticket_data,
                template_data=template_data,
                other_refs_data=other_refs_data,
                progress_callback=on_progress,
                cancel_check=check_cancel,
                deep_think=deep_think,
                zhipu_ocr_api_key=get_zhipu_ocr_api_key(),
                custom_rules=get_custom_rules(),
            )
        except LLMError as e:
            status_container.update(label="❌ 审核出错", state="error")
            status_container.write(f"❌ 错误：{e.message}")
            st.error(f"审核过程中遇到问题：{e.message}")
            return
        except Exception:
            status_container.update(label="❌ 审核出错", state="error")
            status_container.write("❌ 审核过程中发生未知错误")
            st.error("审核过程中发生未知错误，请重试。如反复出现请检查网络连接或联系技术支持")
            return

        # 检查是否被取消
        if audit_result.get("cancelled"):
            status_container.update(label="⏹️ 审核已取消", state="error")
            status_container.write("审核已被用户取消")
            st.warning("审核已取消，已完成的部分结果仍可查看")
            set_cancel_audit(False)
            # 如果有部分结果，仍然显示
            if audit_result.get("per_file_results"):
                st.session_state["audit_result"] = audit_result
                _render_audit_results(audit_result)
            return

        status_container.update(label="✅ 审核完成！", state="complete")

        # 显示 Token 警告
        token_warning = audit_result.get("token_warning", "")
        if token_warning:
            st.warning(token_warning)

        # 保存到 session_state
        st.session_state["audit_result"] = audit_result

        # 保存到历史记录
        file_names = [t.get("filename", "未知") for t in target_files_data]
        add_history_record(audit_result, file_names)

        # --- 显示审核结果 ---
        _render_audit_results(audit_result)


# ============================================================
# 审核结果渲染
# ============================================================
def _render_audit_results(audit_result: dict) -> None:
    """渲染审核结果总览和详情。"""
    per_file = audit_result.get("per_file_results", {})
    cross_check = audit_result.get("cross_check_result")
    errors = audit_result.get("errors", [])

    # --- 总览统计 ---
    total_red = 0
    total_yellow = 0
    total_blue = 0

    for fname, res in per_file.items():
        summary = res.get("summary", {})
        total_red += summary.get("red", 0)
        total_yellow += summary.get("yellow", 0)
        total_blue += summary.get("blue", 0)

    if cross_check:
        cs = cross_check.get("summary", {})
        total_red += cs.get("red", 0)
        total_yellow += cs.get("yellow", 0)
        total_blue += cs.get("blue", 0)

    total_all = total_red + total_yellow + total_blue

    st.markdown("---")
    st.markdown('<p class="result-section-title">📊 审核结果总览</p>', unsafe_allow_html=True)

    # 显示历史记录查看提示
    if st.session_state.get("viewing_history"):
        st.info(f"📜 当前查看的是历史审核记录 #{st.session_state['viewing_history']}")

    if total_all == 0 and not errors:
        st.success("🎉 恭喜！本次审核未发现明显问题，但仍建议进行人工复核")
    else:
        cols = st.columns(4)
        with cols[0]:
            st.markdown(
                f'<div class="stat-card"><div class="label">总标记数</div>'
                f'<div class="value" style="color:#1a1a2e;">{total_all}</div></div>',
                unsafe_allow_html=True,
            )
        with cols[1]:
            st.markdown(
                f'<div class="stat-card"><div class="label">🔴 高风险</div>'
                f'<div class="value" style="color:#cf1322;">{total_red}</div></div>',
                unsafe_allow_html=True,
            )
        with cols[2]:
            st.markdown(
                f'<div class="stat-card"><div class="label">🟡 需注意</div>'
                f'<div class="value" style="color:#d48806;">{total_yellow}</div></div>',
                unsafe_allow_html=True,
            )
        with cols[3]:
            st.markdown(
                f'<div class="stat-card"><div class="label">🔵 格式提醒</div>'
                f'<div class="value" style="color:#1677ff;">{total_blue}</div></div>',
                unsafe_allow_html=True,
            )

    # --- 错误警告 ---
    if errors:
        st.warning("以下文件在审核过程中遇到问题：")
        for err in errors:
            st.markdown(f"- {err}")

    # --- 逐份文件详情 + 下载按钮 ---
    st.markdown("---")
    st.markdown('<p class="result-section-title">📄 逐份审核详情</p>', unsafe_allow_html=True)
    all_reports: list = []

    for fname, res in per_file.items():
        issues = res.get("issues", [])
        summary = res.get("summary", {})
        issue_count = summary.get("total", len(issues))

        badge = ""
        if summary.get("red", 0) > 0:
            badge = "🔴"
        elif summary.get("yellow", 0) > 0:
            badge = "🟡"
        elif summary.get("blue", 0) > 0:
            badge = "🔵"
        else:
            badge = "✅"

        with st.expander(
            f"{badge} {fname}（{issue_count} 处标记）",
            expanded=(summary.get("red", 0) > 0),
        ):
            if not issues:
                st.success("未发现需要标记的问题")
            else:
                _render_issues_table(issues)

            st.markdown('<div class="download-section"></div>', unsafe_allow_html=True)
            _render_download_buttons(fname, res, all_reports)

    # --- 交叉比对结果 ---
    if cross_check:
        cross_issues = cross_check.get("issues", [])
        cross_summary = cross_check.get("summary", {})
        cross_count = cross_summary.get("total", len(cross_issues))

        st.markdown("---")
        st.markdown(
            '<p class="result-section-title">🔄 单据间交叉比对</p>',
            unsafe_allow_html=True,
        )

        with st.expander(
            f"交叉比对结果（{cross_count} 处标记）",
            expanded=(cross_summary.get("red", 0) > 0),
        ):
            if not cross_issues:
                st.success("所有单据之间数据一致")
            else:
                _render_issues_table(cross_issues)

            if cross_issues:
                st.markdown('<div class="download-section"></div>', unsafe_allow_html=True)
                _render_cross_check_download(cross_check, all_reports)

    # --- 一键打包下载 ---
    if all_reports:
        _render_bulk_download(all_reports)


def _render_issues_table(issues: list) -> None:
    """渲染单份审核结果的 issues 列表。"""
    for issue in issues:
        level = issue.get("level", "YELLOW")
        issue_id = issue.get("id", "?")

        if level == "RED":
            color = "#DC2626"
            bg = "#FEF2F2"
            border_color = "#EF4444"
            icon = "🔴"
        elif level == "YELLOW":
            color = "#D97706"
            bg = "#FFFBEB"
            border_color = "#F59E0B"
            icon = "🟡"
        else:
            color = "#2563EB"
            bg = "#EFF6FF"
            border_color = "#3B82F6"
            icon = "🔵"

        field_name = issue.get("field_name", "")
        your_value = issue.get("your_value", "")
        source_value = issue.get("source_value", "")
        source = issue.get("source", "")
        suggestion = issue.get("suggestion", "")
        location = issue.get("field_location", "")

        st.markdown(
            f"""<div class="issue-card" style="background:{bg}; border-left:4px solid {border_color};
            padding:14px 18px; margin:10px 0; border-radius:10px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.04);">
            <strong style="font-size:14px; color:{color};">{icon} [{issue_id}] {field_name}</strong>
            {f'<span style="color:#9ca3af; font-size:12px; margin-left:8px;">{location}</span>' if location else ''}
            <br/>
            <div style="margin-top:8px; font-size:13px; color:#4b5563;">
            <span style="color:#6b7280;">单据上的值：</span><code style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:4px;">{your_value}</code><br/>
            <span style="color:#6b7280;">PO原始值：</span><code style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:4px;">{source_value}</code>
            {f'<br/><span style="color:#9ca3af; font-size:12px;">数据来源：{source}</span>' if source else ''}
            </div>
            <div style="margin-top:8px; color:{color}; font-size:13px; font-weight:500;">💡 {suggestion}</div>
            </div>""",
            unsafe_allow_html=True,
        )


# ============================================================
# 报告下载辅助函数
# ============================================================
def _get_contract_no_from_result(audit_result: dict) -> str:
    """从审核结果中提取合同号。"""
    per_file = audit_result.get("per_file_results", {})
    for fname, res in per_file.items():
        issues = res.get("issues", [])
        for issue in issues:
            field_name = issue.get("field_name", "")
            if any(
                kw in field_name
                for kw in ("合同", "Invoice", "invoice", "Contract", "contract")
            ):
                val = issue.get("source_value", "")
                if val:
                    return str(val).strip()
    return "未知合同号"


def _guess_doc_type_from_filename(fname: str) -> str:
    """从文件名猜测单据类型。"""
    name_lower = fname.lower()
    hints = {
        "ci": "商业发票CI",
        "invoice": "商业发票CI",
        "发票": "商业发票CI",
        "pl": "装箱单PL",
        "packing": "装箱单PL",
        "装箱": "装箱单PL",
        "booking": "托书Booking",
        "托书": "托书Booking",
        "生产通知": "生产通知单",
        "production": "生产通知单",
        "发货申请": "发货申请单",
        "shipping": "发货申请单",
    }
    for keyword, doc_type in hints.items():
        if keyword in name_lower:
            return doc_type
    return "待审核单据"


def _render_download_buttons(fname: str, file_result: dict, all_reports: list) -> None:
    """为单个文件渲染标记版和详情版下载按钮。"""
    issues = file_result.get("issues", [])
    original_text = file_result.get("original_text", "")
    contract_no = _get_contract_no_from_result(
        st.session_state.get("audit_result", {})
    )
    doc_type = _guess_doc_type_from_filename(fname)

    dl_col1, dl_col2 = st.columns(2)

    with dl_col1:
        try:
            marked_buf, marked_name = generate_marked_report(
                original_text=original_text,
                audit_issues=issues,
                doc_type=doc_type,
                contract_no=contract_no,
            )
            st.download_button(
                label="📝 下载标记版 (Excel)",
                data=marked_buf,
                file_name=marked_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_marked_{fname}",
            )
            all_reports.append((marked_name, marked_buf))
        except Exception:
            st.error("标记版报告生成失败，请重试")

    with dl_col2:
        try:
            detail_buf, detail_name = generate_detail_report(
                audit_issues=issues,
                doc_type=doc_type,
                contract_no=contract_no,
            )
            st.download_button(
                label="📋 下载详情版 (Excel)",
                data=detail_buf,
                file_name=detail_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_detail_{fname}",
            )
            all_reports.append((detail_name, detail_buf))
        except Exception:
            st.error("详情版报告生成失败，请重试")


def _render_cross_check_download(cross_check_result: dict, all_reports: list) -> None:
    """渲染交叉比对结果的下载按钮。"""
    cross_issues = cross_check_result.get("issues", [])
    contract_no = _get_contract_no_from_result(
        st.session_state.get("audit_result", {})
    )

    try:
        detail_buf, detail_name = generate_detail_report(
            audit_issues=cross_issues,
            doc_type="交叉比对",
            contract_no=contract_no,
        )
        st.download_button(
            label="📋 下载交叉比对详情 (Excel)",
            data=detail_buf,
            file_name=detail_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_cross_check",
        )
        all_reports.append((detail_name, detail_buf))
    except Exception:
        st.error("交叉比对报告生成失败，请重试")


def _render_bulk_download(all_reports: list) -> None:
    """渲染一键打包下载按钮。"""
    st.markdown("---")
    st.markdown(
        '<div class="download-area">', unsafe_allow_html=True
    )
    st.markdown(
        '<p class="result-section-title">📦 报告打包下载</p>',
        unsafe_allow_html=True,
    )

    contract_no = _get_contract_no_from_result(
        st.session_state.get("audit_result", {})
    )

    btn_col1, btn_col2, btn_col3 = st.columns([1, 2, 1])
    with btn_col2:
        try:
            zip_buf, zip_name = generate_zip(all_reports, contract_no=contract_no)
            st.download_button(
                label="📦 一键打包下载全部报告",
                data=zip_buf,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True,
                type="primary",
                key="dl_bulk_zip",
            )
            st.markdown(
                f'<p class="hint-text" style="text-align:center; margin-top:8px;">'
                f'包含 {len(all_reports)} 份 Excel 报告文件</p>',
                unsafe_allow_html=True,
            )
        except Exception:
            st.error("打包下载失败，请重试")

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# 主界面
# ============================================================
def render_main_page() -> None:
    """渲染主界面内容。"""

    # ============================================================
    # 主页面标题
    # ============================================================
    st.markdown(
        '<p class="main-title">📋 外贸跟单工单智能审核系统</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="main-subtitle">'
        '基于AI大模型的外贸单据智能比对与审核工具 — 帮您兜住每一个问题 | '
        '支持 PDF / Word / Excel / 图片输入，统一输出 Excel 格式报告</p>',
        unsafe_allow_html=True,
    )

    # ============================================================
    # 第一行：核心文件上传（PO + 待审核文件），等宽两栏
    # ============================================================
    col_po, col_audit = st.columns(2, gap="large")

    with col_po:
        st.markdown(
            '<div class="upload-col-title upload-col-title-left">'
            '📁 参照数据源（审核基准）</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<p class="hint-text">PO文件（必须上传）</p>',
            unsafe_allow_html=True,
        )
        po_file = st.file_uploader(
            "上传PO文件",
            type=ALLOWED_DOC_TYPES,
            label_visibility="collapsed",
            key="po_uploader",
        )
        _render_file_preview(po_file, multi=False)

    with col_audit:
        st.markdown(
            '<div class="upload-col-title upload-col-title-right">'
            '📄 待审核文件（本票单据）</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<p class="hint-text">上传本票需要审核的单据</p>',
            unsafe_allow_html=True,
        )
        audit_files = st.file_uploader(
            "上传待审核文件",
            type=ALLOWED_DOC_TYPES + ["jpg", "png", "jpeg"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key="audit_uploader",
        )
        st.markdown(
            '<p class="hint-text">'
            '支持 PDF / Word (.doc/.docx) / Excel (.xlsx/.xls) 格式<br>'
            '可上传 CI、PL、托书、生产通知单、发货申请单等</p>',
            unsafe_allow_html=True,
        )
        _render_file_preview(audit_files, multi=True)

    # ============================================================
    # 第二行：辅助参考文件（可选），三等分
    # ============================================================
    st.markdown("---")
    st.markdown("**📎 辅助参考文件（可选，提供后可提升审核准确性）**")

    aux_col1, aux_col2, aux_col3 = st.columns(3, gap="medium")

    with aux_col1:
        template_file = st.file_uploader(
            "📋 标准模板文件（建议提供）",
            type=ALLOWED_DOC_TYPES,
            key="template_uploader",
            help="上传标准模板文件，系统将对比待审核文件与模板的格式和内容差异",
        )
        if template_file:
            _render_file_preview(template_file, multi=False)

    with aux_col2:
        prev_files = st.file_uploader(
            "📂 上一票对应文件（建议提供）",
            type=ALLOWED_DOC_TYPES,
            accept_multiple_files=True,
            key="prev_uploader",
            help="上传上一票的同类单据，系统将进行前后票据对比",
        )
        if prev_files:
            _render_file_preview(prev_files, multi=True)

    with aux_col3:
        ref_images = st.file_uploader(
            "🖼️ 其他参考截图（可选）",
            type=["jpg", "png", "jpeg"],
            accept_multiple_files=True,
            key="ref_uploader",
            help="如有聊天记录截图、邮件截图、客户要求截图等参考信息可在此上传",
        )
        if ref_images:
            _render_file_preview(ref_images, multi=True)

    # ============================================================
    # 审核按钮 — 居中
    # ============================================================
    st.markdown("---")
    _, btn_col, _ = st.columns([2, 1, 2])
    with btn_col:
        start_audit = st.button(
            "🚀 开始审核",
            type="primary",
            use_container_width=True,
        )

    # ============================================================
    # 审核结果区域
    # ============================================================
    result_placeholder = st.container()

    if start_audit:
        # 清除历史查看状态
        st.session_state.pop("viewing_history", None)
        _handle_audit_start(
            po_file, template_file, prev_files, ref_images,
            audit_files, result_placeholder,
        )
    elif "audit_result" in st.session_state and st.session_state["audit_result"]:
        with result_placeholder:
            _render_audit_results(st.session_state["audit_result"])


# ============================================================
# 主流程
# ============================================================
def main() -> None:
    # 1. 免责声明拦截
    if not is_disclaimer_accepted():
        show_disclaimer()

    # 2. 侧边栏
    render_sidebar()

    # 3. 主界面
    render_main_page()


if __name__ == "__main__":
    main()
