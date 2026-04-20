"""
Token 长度估算与内容分段处理工具模块
用于在构建 Prompt 前预估 token 数量，并在超限时进行智能截断/分段。
"""

import logging
import re
from typing import List, Tuple

logger = logging.getLogger(__name__)

# ============================================================
# 模型 Token 限制配置
# ============================================================
MODEL_TOKEN_LIMITS = {
    "deepseek": 64000,       # DeepSeek 约 64K tokens
    "deepseek-reasoner": 64000,
    "zhipu": 128000,         # 智谱 GLM 约 128K tokens
}

# 安全阈值比例：使用模型上限的 80%
SAFETY_RATIO = 0.80

# 系统提示词和输出格式预留 token（约 3000 tokens）
RESERVED_TOKENS = 3000


# ============================================================
# Token 估算函数
# ============================================================
def estimate_tokens(text: str) -> int:
    """估算文本的 token 数量。

    使用混合估算方法：
    - 中文字符：约 2 tokens/字
    - 英文单词：约 1.5 tokens/词
    - 数字和标点：约 1 token/字符

    先尝试使用 tiktoken 进行精确计算，失败则回退到字符估算。

    Args:
        text: 待估算的文本。

    Returns:
        预估的 token 数量。
    """
    if not text:
        return 0

    # 尝试使用 tiktoken 进行精确计算
    try:
        import tiktoken
        enc = tiktoken.get_encoding("cl100k_base")
        return len(enc.encode(text))
    except Exception:
        pass

    # 回退到字符估算方法
    return _estimate_tokens_by_char(text)


def _estimate_tokens_by_char(text: str) -> int:
    """基于字符的 token 估算方法。"""
    if not text:
        return 0

    chinese_chars = len(re.findall(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]', text))
    remaining_text = re.sub(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]', '', text)
    english_words = len(re.findall(r'[a-zA-Z]+', remaining_text))
    numbers = len(re.findall(r'\d+', remaining_text))
    punctuation = len(re.findall(r'[^\w\s]', remaining_text))

    estimated = int(
        chinese_chars * 2
        + english_words * 1.5
        + numbers * 1
        + punctuation * 1
    )

    return max(estimated, 1)


def get_model_token_limit(provider: str) -> int:
    """获取指定模型的 token 上限。

    Args:
        provider: 模型标识（如 "deepseek", "zhipu"）。

    Returns:
        模型的 token 上限。
    """
    provider_lower = provider.lower().strip()
    for key, limit in MODEL_TOKEN_LIMITS.items():
        if key in provider_lower:
            return limit
    # 默认返回较保守的限制
    return 64000


def get_safe_token_limit(provider: str) -> int:
    """获取安全 token 限制（模型上限的 80%）。

    Args:
        provider: 模型标识。

    Returns:
        安全的 token 上限。
    """
    return int(get_model_token_limit(provider) * SAFETY_RATIO) - RESERVED_TOKENS


# ============================================================
# 内容截断 / 摘要处理
# ============================================================
def truncate_text(text: str, max_tokens: int) -> str:
    """将文本截断到指定的 token 数以内。

    按行截断，保留尽可能多的完整行。

    Args:
        text: 原始文本。
        max_tokens: 最大 token 数。

    Returns:
        截断后的文本。
    """
    if estimate_tokens(text) <= max_tokens:
        return text

    lines = text.split('\n')
    result_lines = []
    current_tokens = 0

    for line in lines:
        line_tokens = estimate_tokens(line)
        if current_tokens + line_tokens > max_tokens:
            break
        result_lines.append(line)
        current_tokens += line_tokens

    truncated = '\n'.join(result_lines)
    truncated += "\n\n[注意：内容已被截断以适应模型上下文窗口限制，部分末尾内容可能未包含在审核范围内]"
    return truncated


def smart_split_content(
    po_text: str,
    target_text: str,
    other_texts: List[str],
    provider: str,
) -> Tuple[str, str, List[str], bool]:
    """智能分段处理内容，确保不超过模型 token 限制。

    优先策略：
    1. PO 全文必须完整保留（PO 是审核基准不能截断）
    2. 待审核文件内容进行摘要压缩
    3. 辅助文本（上一票、模板等）在必要时进行截断

    Args:
        po_text: PO 原文。
        target_text: 待审核文件原文。
        other_texts: 其他辅助文本列表（模板、上一票等）。
        provider: 模型提供商标识。

    Returns:
        (处理后的po_text, 处理后的target_text, 处理后的other_texts, 是否进行了截断)
    """
    safe_limit = get_safe_token_limit(provider)
    was_truncated = False

    # 计算各部分 token 数
    po_tokens = estimate_tokens(po_text)
    target_tokens = estimate_tokens(target_text)
    other_tokens = [estimate_tokens(t) for t in other_texts]
    total_other_tokens = sum(other_tokens)
    total_tokens = po_tokens + target_tokens + total_other_tokens

    logger.info(
        "Token 估算 - PO: %d, 目标: %d, 辅助: %d, 总计: %d, 安全限制: %d",
        po_tokens, target_tokens, total_other_tokens, total_tokens, safe_limit
    )

    if total_tokens <= safe_limit:
        return po_text, target_text, other_texts, False

    # 需要截断处理
    was_truncated = True

    # 步骤1: PO 保留全文，计算剩余可用 token
    remaining = safe_limit - po_tokens

    if remaining <= 0:
        # 极端情况：PO 本身已超限，截断 PO
        po_text = truncate_text(po_text, safe_limit * 7 // 10)
        remaining = safe_limit - estimate_tokens(po_text)
        was_truncated = True

    # 步骤2: 为待审核文件分配至少 60% 的剩余空间
    target_budget = int(remaining * 0.6)
    other_budget = remaining - target_budget

    # 步骤3: 截断待审核文件内容（如果需要）
    if target_tokens > target_budget:
        target_text = truncate_text(target_text, target_budget)

    # 步骤4: 按比例截断辅助文本
    if total_other_tokens > other_budget and other_texts:
        processed_others = []
        per_other_budget = other_budget // max(len(other_texts), 1)
        for text in other_texts:
            if estimate_tokens(text) > per_other_budget:
                processed_others.append(truncate_text(text, per_other_budget))
            else:
                processed_others.append(text)
        other_texts = processed_others

    return po_text, target_text, other_texts, was_truncated
