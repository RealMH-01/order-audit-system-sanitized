"""
审核历史记录管理模块
在 Streamlit session_state 中管理审核历史记录。
历史记录仅在当前会话有效，刷新页面后将清空。
"""

import copy
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional

import streamlit as st

logger = logging.getLogger(__name__)

KEY_AUDIT_HISTORY = "audit_history"


def _ensure_history_initialized() -> None:
    """确保历史记录列表已初始化。"""
    if KEY_AUDIT_HISTORY not in st.session_state:
        st.session_state[KEY_AUDIT_HISTORY] = []


def add_history_record(
    audit_result: Dict[str, Any],
    file_names: List[str],
) -> None:
    """将一次审核结果保存到历史记录。

    Args:
        audit_result: 完整的审核结果字典。
        file_names: 参与审核的文件名列表。
    """
    _ensure_history_initialized()

    # 存入历史时去掉大文本字段，节省内存
    cleaned_result = copy.deepcopy(audit_result)
    for fname, res in cleaned_result.get("per_file_results", {}).items():
        res.pop("original_text", None)

    per_file = cleaned_result.get("per_file_results", {})
    cross_check = cleaned_result.get("cross_check_result")

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

    record = {
        "id": len(st.session_state[KEY_AUDIT_HISTORY]) + 1,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_names": file_names,
        "total_red": total_red,
        "total_yellow": total_yellow,
        "total_blue": total_blue,
        "total_issues": total_red + total_yellow + total_blue,
        "audit_result": cleaned_result,
    }

    st.session_state[KEY_AUDIT_HISTORY].append(record)
    logger.info("审核历史已保存，共 %d 条记录", len(st.session_state[KEY_AUDIT_HISTORY]))


def get_history_records() -> List[Dict[str, Any]]:
    """获取所有历史记录（按时间倒序）。"""
    _ensure_history_initialized()
    return list(reversed(st.session_state[KEY_AUDIT_HISTORY]))


def get_history_record_by_id(record_id: int) -> Optional[Dict[str, Any]]:
    """根据 ID 获取单条历史记录。"""
    _ensure_history_initialized()
    for record in st.session_state[KEY_AUDIT_HISTORY]:
        if record["id"] == record_id:
            return record
    return None


def get_history_count() -> int:
    """获取历史记录总数。"""
    _ensure_history_initialized()
    return len(st.session_state[KEY_AUDIT_HISTORY])


def clear_history() -> None:
    """清空所有历史记录。"""
    st.session_state[KEY_AUDIT_HISTORY] = []
