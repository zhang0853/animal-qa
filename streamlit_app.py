"""
动物检疫知识库智能问答系统 - Streamlit Web版
功能：流式输出、历史记录、追问、Word导出
安全：管理员密码保护
"""

import streamlit as st
import time
import os
from datetime import datetime
from knowledge_base import KnowledgeBase, export_single_qa, export_batch_qa

# ========== 页面配置 ==========
st.set_page_config(
    page_title="动物检疫知识库",
    page_icon="🐄",
    layout="wide"
)

# 创建导出目录
EXPORT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "导出问答")
os.makedirs(EXPORT_DIR, exist_ok=True)

# ========== 密码配置（你可以修改这个密码）==========
ADMIN_PASSWORD = "admin123"  # 管理员密码，可以改成你想要的

# ========== 初始化 ==========
if "messages" not in st.session_state:
    st.session_state.messages = []
if "history_items" not in st.session_state:
    st.session_state.history_items = []
if "kb" not in st.session_state:
    with st.spinner("正在加载知识库..."):
        st.session_state.kb = KnowledgeBase()
        success, msg = st.session_state.kb.load_index()
        if not success:
            st.error(f"❌ 加载失败: {msg}")
            st.info("请点击左侧「重建索引」")
if "current_answer" not in st.session_state:
    st.session_state.current_answer = ""
if "current_sources" not in st.session_state:
    st.session_state.current_sources = []
if "selected_history" not in st.session_state:
    st.session_state.selected_history = set()
if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False
if "password_input" not in st.session_state:
    st.session_state.password_input = ""


# ========== 辅助函数 ==========
def add_to_history(question, answer, sources):
    st.session_state.history_items.append({
        "question": question,
        "answer": answer,
        "sources": sources,
        "timestamp": datetime.now().strftime("%m-%d %H:%M")
    })


def clear_history():
    st.session_state.history_items = []
    st.session_state.messages = []
    st.session_state.selected_history = set()
    st.rerun()


def delete_selected():
    for idx in sorted(list(st.session_state.selected_history), reverse=True):
        if idx < len(st.session_state.history_items):
            del st.session_state.history_items[idx]
    st.session_state.selected_history = set()
    st.rerun()


def toggle_select_all():
    if len(st.session_state.selected_history) == len(st.session_state.history_items):
        st.session_state.selected_history = set()
    else:
        st.session_state.selected_history = set(range(len(st.session_state.history_items)))
    st.rerun()


def check_admin_password(password):
    """验证管理员密码"""
    return password == ADMIN_PASSWORD


# ========== 侧边栏 ==========
with st.sidebar:
    st.markdown("# 🐄 动物检疫知识库")
    st.caption("智能问答系统")

    st.divider()

    # ========== 管理员区域（需要密码）==========
    st.markdown("### 🔐 管理员区域")

    if not st.session_state.admin_mode:
        # 未登录状态，显示密码输入框
        pwd = st.text_input("管理员密码", type="password", key="admin_pwd_input")
        if st.button("🔓 进入管理模式", use_container_width=True):
            if check_admin_password(pwd):
                st.session_state.admin_mode = True
                st.success("✅ 管理员模式已启用")
                st.rerun()
            else:
                st.error("❌ 密码错误")
        st.info("普通用户无需密码，直接使用问答功能")
    else:
        # 已登录状态，显示管理功能
        st.success("👑 管理员模式已启用")
        if st.button("🔒 退出管理模式", use_container_width=True):
            st.session_state.admin_mode = False
            st.rerun()

        st.divider()

        # 管理员专属功能
        st.markdown("### 📚 知识库管理")

        # 显示当前状态
        if st.session_state.kb.vectorstore:
            st.info(f"当前索引: {st.session_state.kb.vectorstore.index.ntotal} 个向量")
        else:
            st.warning("索引未加载")

        # 重建索引按钮
        if st.button("🔄 重建索引", use_container_width=True, type="primary"):
            with st.spinner("正在重建索引..."):
                success, msg = st.session_state.kb.rebuild_index()
                if success:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

        # 显示知识库目录
        with st.expander("📁 查看知识库文件"):
            kb_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "检疫知识库")
            if os.path.exists(kb_dir):
                files = [f for f in os.listdir(kb_dir) if f.endswith(('.pdf', '.docx', '.txt'))]
                if files:
                    for f in files:
                        st.write(f"📄 {f}")
                else:
                    st.write("暂无文档")
            else:
                st.write("目录不存在")

    st.divider()

    # ========== 普通用户功能（所有人可见）==========
    st.markdown("### 🌐 联网搜索")
    enable_web = st.checkbox("启用联网搜索", value=st.session_state.kb.config.get("enable_web_search", True))
    if enable_web != st.session_state.kb.config.get("enable_web_search"):
        st.session_state.kb.config["enable_web_search"] = enable_web

    web_count = st.number_input("搜索条数", 1, 10, st.session_state.kb.config.get("web_search_count", 3))
    if web_count != st.session_state.kb.config.get("web_search_count"):
        st.session_state.kb.config["web_search_count"] = web_count

    st.divider()

    st.markdown("### 🗑️ 历史管理")
    if st.button("清空所有历史", use_container_width=True):
        clear_history()

    st.divider()

    # ========== 统计信息 ==========
    st.markdown("### 📊 统计")
    vector_count = st.session_state.kb.vectorstore.index.ntotal if st.session_state.kb.vectorstore else 0
    st.metric("向量数", vector_count)
    st.metric("对话轮数", len(st.session_state.messages) // 2)
    st.metric("历史记录", len(st.session_state.history_items))

    # 显示当前模式
    if st.session_state.admin_mode:
        st.markdown("---")
        st.markdown("🔴 **管理员模式**")

# ========== 主界面 ==========
st.title("🐄 动物检疫知识库智能问答系统")
st.caption("基于本地知识库 + 大模型 | 专业动物检疫咨询服务")

# 创建两列
chat_col, history_col = st.columns([3, 1])

with chat_col:
    st.markdown("### 💬 对话")

    # 显示对话
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # 输入框
    if prompt := st.chat_input("输入您的问题..."):
        # 用户消息
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # AI回复
        with st.chat_message("assistant"):
            msg_placeholder = st.empty()
            full_response = ""

            history = [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages[:-1]]

            try:
                for chunk in st.session_state.kb.ask_stream(prompt, history):
                    if chunk["type"] == "chunk":
                        full_response += chunk["content"]
                        msg_placeholder.markdown(full_response + "▌")
                    elif chunk["type"] == "complete":
                        msg_placeholder.markdown(full_response)
                        st.session_state.messages.append({"role": "assistant", "content": full_response})
                        st.session_state.current_answer = full_response
                        st.session_state.current_sources = chunk.get("sources", [])
                        add_to_history(prompt, full_response, chunk.get("sources", []))

                        # 显示来源
                        if chunk.get("sources"):
                            with st.expander("📚 信息来源"):
                                for src in chunk["sources"][:5]:
                                    if src.get("type") == "local":
                                        st.write(f"📄 {src['source']}")
                                    else:
                                        st.write(f"🌐 {src.get('title', '网络信息')}")
                    elif chunk["type"] == "error":
                        st.error(f"错误: {chunk['content']}")
            except Exception as e:
                st.error(f"请求失败: {e}")

with history_col:
    st.markdown("### 📋 历史记录")

    # 操作按钮
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("全选", use_container_width=True):
            toggle_select_all()
    with col2:
        if st.button("删除", use_container_width=True):
            delete_selected()
    with col3:
        if st.button("清空", use_container_width=True):
            clear_history()

    st.markdown("---")

    # 导出按钮
    if st.button("💾 导出当前问答", use_container_width=True):
        if st.session_state.current_answer:
            current_q = ""
            for msg in reversed(st.session_state.messages):
                if msg["role"] == "user":
                    current_q = msg["content"]
                    break
            if current_q:
                filename = f"问答_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                filepath = os.path.join(EXPORT_DIR, filename)
                export_single_qa(current_q, st.session_state.current_answer, st.session_state.current_sources, filepath)
                with open(filepath, "rb") as f:
                    st.download_button("📥 下载", f, filename, use_container_width=True)
        else:
            st.warning("暂无内容")

    if st.button("📚 批量导出", use_container_width=True):
        if st.session_state.selected_history:
            items = [st.session_state.history_items[i] for i in st.session_state.selected_history]
            filename = f"批量导出_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            filepath = os.path.join(EXPORT_DIR, filename)
            export_batch_qa(items, filepath)
            with open(filepath, "rb") as f:
                st.download_button("📥 下载", f, filename, use_container_width=True)
        else:
            st.warning("请先选择记录")

    st.markdown("---")

    # 历史列表
    if st.session_state.history_items:
        for idx, item in enumerate(st.session_state.history_items):
            is_selected = idx in st.session_state.selected_history
            col_a, col_b = st.columns([0.1, 0.9])
            with col_a:
                if st.checkbox("", value=is_selected, key=f"chk_{idx}"):
                    if not is_selected:
                        st.session_state.selected_history.add(idx)
                    else:
                        st.session_state.selected_history.discard(idx)
            with col_b:
                preview = item["question"][:30] + "..." if len(item["question"]) > 30 else item["question"]
                st.write(f"**{item['timestamp']}**")
                st.write(preview)
                if st.button("追问", key=f"ask_{idx}", use_container_width=True):
                    # 直接添加到输入框
                    st.session_state.messages.append({"role": "user", "content": item["question"]})
                    st.rerun()
                st.markdown("---")
    else:
        st.info("暂无历史记录")

# 提示
if not st.session_state.kb.vectorstore:
    st.warning("⚠️ 知识库未加载，请管理员在左侧输入密码后重建索引")