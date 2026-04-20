"""
大模型 API 对接模块
支持 DeepSeek（OpenAI 兼容）和智谱 GLM（zhipuai SDK）。
支持 DeepSeek 深度思考模式（deepseek-reasoner）。
所有错误统一转为中文友好提示，通过 LLMError 异常抛出。
"""

import logging
from typing import List, Dict, Optional

logger = logging.getLogger(__name__)

# ============================================================
# 超时配置
# ============================================================
TIMEOUT_NORMAL = 120
TIMEOUT_DEEP_THINK = 300

# ============================================================
# 自定义异常
# ============================================================
class LLMError(Exception):
    """大模型调用异常，message 始终为面向用户的中文提示。"""
    def __init__(self, message: str):
        self.message = message
        super().__init__(message)


class LLMTimeoutError(LLMError):
    """大模型调用超时异常。"""
    def __init__(self, timeout_seconds: int):
        message = (
            f"AI 模型响应超时（等待超过 {timeout_seconds} 秒），"
            "可能是网络问题或文件内容过长，请稍后重试或减少上传文件数量"
        )
        super().__init__(message)


# ============================================================
# 内部：错误码 - 中文提示映射
# ============================================================
_HTTP_ERROR_MAP: Dict[int, str] = {
    401: "API密钥无效，请检查后重新输入",
    403: "API密钥无效或无权访问，请检查后重新输入",
    402: "API余额不足，请充值后重试",
    429: "请求过于频繁，请稍后再试",
    500: "AI服务内部错误，请稍后重试",
    502: "AI服务暂时不可用，请稍后重试",
    503: "AI服务暂时不可用，请稍后重试",
}


def _friendly_error(exc: Exception) -> str:
    """将各种异常转为中文友好提示字符串。"""
    exc_str = str(exc).lower()

    try:
        from openai import (
            AuthenticationError,
            RateLimitError,
            APITimeoutError,
            APIConnectionError,
            APIStatusError,
        )
        if isinstance(exc, AuthenticationError):
            return "API密钥无效，请检查后重新输入"
        if isinstance(exc, RateLimitError):
            if "insufficient" in exc_str or "balance" in exc_str:
                return "API余额不足，请充值后重试"
            return "请求过于频繁，请稍后再试"
        if isinstance(exc, APITimeoutError):
            return "AI 模型响应超时，可能是网络问题或文件内容过长，请稍后重试或减少上传文件数量"
        if isinstance(exc, APIConnectionError):
            return "网络连接失败，请检查网络后重试"
        if isinstance(exc, APIStatusError):
            status = getattr(exc, "status_code", None)
            if status and status in _HTTP_ERROR_MAP:
                return _HTTP_ERROR_MAP[status]
    except ImportError:
        pass

    try:
        import zhipuai as _zhipuai
        if isinstance(exc, _zhipuai.APIAuthenticationError):
            return "API密钥无效，请检查后重新输入"
        if isinstance(exc, _zhipuai.APIReachLimitError):
            return "请求过于频繁，请稍后再试"
        if isinstance(exc, _zhipuai.APITimeoutError):
            return "AI 模型响应超时，可能是网络问题或文件内容过长，请稍后重试或减少上传文件数量"
        if isinstance(exc, _zhipuai.APIStatusError):
            status = getattr(exc, "status_code", None)
            if status and status in _HTTP_ERROR_MAP:
                return _HTTP_ERROR_MAP[status]
    except (ImportError, AttributeError):
        pass

    if "timeout" in exc_str or "timed out" in exc_str:
        return "AI 模型响应超时，可能是网络问题或文件内容过长，请稍后重试或减少上传文件数量"
    if "connection" in exc_str or "network" in exc_str or "resolve" in exc_str:
        return "网络连接失败，请检查网络后重试"
    if "401" in exc_str or "authentication" in exc_str or "unauthorized" in exc_str:
        return "API密钥无效，请检查后重新输入"
    if "429" in exc_str or "rate" in exc_str:
        return "请求过于频繁，请稍后再试"
    if "insufficient" in exc_str or "balance" in exc_str or "402" in exc_str:
        return "API余额不足，请充值后重试"

    return "发生错误：请求AI服务时出现异常，请重试"


# ============================================================
# DeepSeek（OpenAI 兼容）
# ============================================================
_DEEPSEEK_BASE_URL = "https://api.deepseek.com"
_DEEPSEEK_TEXT_MODEL = "deepseek-chat"
_DEEPSEEK_REASONER_MODEL = "deepseek-reasoner"
_DEEPSEEK_VISION_MODEL = "deepseek-chat"


def _call_deepseek(
    api_key: str,
    messages: List[Dict],
    temperature: float = 0.1,
    deep_think: bool = False,
    timeout: Optional[int] = None,
) -> str:
    """调用 DeepSeek API（OpenAI 兼容格式）。"""
    from openai import OpenAI

    if timeout is None:
        timeout = TIMEOUT_DEEP_THINK if deep_think else TIMEOUT_NORMAL

    model = _DEEPSEEK_REASONER_MODEL if deep_think else _DEEPSEEK_TEXT_MODEL

    client = OpenAI(
        api_key=api_key,
        base_url=_DEEPSEEK_BASE_URL,
        timeout=timeout,
    )

    kwargs = {"model": model, "messages": messages}
    if not deep_think:
        kwargs["temperature"] = temperature

    # 深度思考模式下，将 system 消息转为 user 消息前缀
    if deep_think:
        processed_messages = []
        for msg in messages:
            if msg.get("role") == "system":
                processed_messages.append({
                    "role": "user",
                    "content": "[系统指令]\n" + msg["content"]
                })
            else:
                processed_messages.append(msg)
        kwargs["messages"] = processed_messages

    response = client.chat.completions.create(**kwargs)
    content = response.choices[0].message.content if response.choices else ""
    if not content or not content.strip():
        raise LLMError("AI未返回有效内容，请重试")
    return content.strip()


def _call_deepseek_vision(
    api_key: str,
    prompt: str,
    image_base64: str,
    timeout: Optional[int] = None,
) -> str:
    """调用 DeepSeek 多模态接口识别图片。"""
    from openai import OpenAI

    if timeout is None:
        timeout = TIMEOUT_NORMAL

    client = OpenAI(
        api_key=api_key,
        base_url=_DEEPSEEK_BASE_URL,
        timeout=timeout,
    )
    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {
                    "type": "image_url",
                    "image_url": {
                        "url": "data:image/png;base64," + image_base64
                    },
                },
            ],
        }
    ]
    response = client.chat.completions.create(
        model=_DEEPSEEK_VISION_MODEL,
        messages=messages,
        temperature=0.1,
    )
    content = response.choices[0].message.content if response.choices else ""
    if not content or not content.strip():
        raise LLMError("AI未返回有效内容，请重试")
    return content.strip()


# ============================================================
# 智谱 GLM（zhipuai SDK）
# ============================================================
_ZHIPU_TEXT_MODEL = "glm-4-flash"
_ZHIPU_VISION_MODEL = "glm-4.6v"


def _call_zhipu(
    api_key: str,
    messages: List[Dict],
    temperature: float = 0.1,
    timeout: Optional[int] = None,
) -> str:
    """调用智谱 GLM API。"""
    from zhipuai import ZhipuAI

    if timeout is None:
        timeout = TIMEOUT_NORMAL

    client = ZhipuAI(api_key=api_key)
    response = client.chat.completions.create(
        model=_ZHIPU_TEXT_MODEL,
        messages=messages,
        temperature=temperature,
        timeout=timeout,
    )
    content = response.choices[0].message.content if response.choices else ""
    if not content or not content.strip():
        raise LLMError("AI未返回有效内容，请重试")
    return content.strip()


def _call_zhipu_vision(
    api_key: str,
    prompt: str,
    image_base64: str,
    timeout: Optional[int] = None,
) -> str:
    """调用智谱 GLM 多模态接口识别图片。"""
    from zhipuai import ZhipuAI

    if timeout is None:
        timeout = TIMEOUT_NORMAL

    client = ZhipuAI(api_key=api_key)
    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {
                    "type": "image_url",
                    "image_url": {
                        "url": "data:image/png;base64," + image_base64
                    },
                },
            ],
        }
    ]
    response = client.chat.completions.create(
        model=_ZHIPU_VISION_MODEL,
        messages=messages,
        temperature=0.1,
        timeout=timeout,
    )
    content = response.choices[0].message.content if response.choices else ""
    if not content or not content.strip():
        raise LLMError("AI未返回有效内容，请重试")
    return content.strip()


# ============================================================
# 统一对外接口
# ============================================================
def _resolve_provider(provider: str) -> str:
    """将用户界面中的模型名称映射为内部 provider 标识。"""
    mapping = {
        "deepseek": "deepseek",
        "智谱glm": "zhipu",
        "zhipu": "zhipu",
    }
    return mapping.get(provider.lower().strip(), provider.lower().strip())


def call_llm(
    provider: str,
    api_key: str,
    messages: List[Dict],
    temperature: float = 0.1,
    deep_think: bool = False,
    timeout: Optional[int] = None,
) -> str:
    """统一文本调用接口。

    Args:
        provider: "DeepSeek" 或 "智谱GLM"
        api_key: 用户的 API 密钥
        messages: 对话消息列表
        temperature: 采样温度，默认 0.1
        deep_think: 是否使用深度思考模式（仅 DeepSeek 支持）
        timeout: 超时时间（秒），None 则使用默认值

    Returns:
        模型回复的文本字符串

    Raises:
        LLMError: 包含中文友好错误提示
    """
    if not api_key or not api_key.strip():
        raise LLMError("请先输入API密钥")

    prov = _resolve_provider(provider)

    if timeout is None:
        timeout = TIMEOUT_DEEP_THINK if (deep_think and prov == "deepseek") else TIMEOUT_NORMAL

    try:
        if prov == "deepseek":
            return _call_deepseek(
                api_key, messages, temperature,
                deep_think=deep_think, timeout=timeout,
            )
        elif prov == "zhipu":
            return _call_zhipu(api_key, messages, temperature, timeout=timeout)
        else:
            raise LLMError("不支持的模型提供商: " + provider)
    except LLMError:
        raise
    except Exception as exc:
        logger.error("LLM 调用失败 [%s]: %s", prov, exc)
        raise LLMError(_friendly_error(exc)) from exc


def call_llm_with_image(
    provider: str,
    api_key: str,
    prompt: str,
    image_base64: str,
    timeout: Optional[int] = None,
) -> str:
    """多模态调用接口（图片 OCR）。"""
    if not api_key or not api_key.strip():
        raise LLMError("请先输入API密钥")
    if not image_base64:
        raise LLMError("图片数据为空，无法识别")

    prov = _resolve_provider(provider)

    if timeout is None:
        timeout = TIMEOUT_NORMAL

    try:
        if prov == "deepseek":
            return _call_deepseek_vision(api_key, prompt, image_base64, timeout=timeout)
        elif prov == "zhipu":
            return _call_zhipu_vision(api_key, prompt, image_base64, timeout=timeout)
        else:
            raise LLMError("不支持的模型提供商: " + provider)
    except LLMError:
        raise
    except Exception as exc:
        logger.error("LLM 视觉调用失败 [%s]: %s", prov, exc)
        raise LLMError(_friendly_error(exc)) from exc


# ============================================================
# 便捷函数
# ============================================================
IMAGE_OCR_PROMPT = (
    "请识别这张图片中的所有文字内容，"
    "按原始排版格式输出，不要遗漏任何信息。"
)


def test_connection(provider: str, api_key: str) -> str:
    """测试连接：发送简单消息验证密钥是否有效。"""
    messages = [
        {"role": "user", "content": "请回复'连接成功'四个字"}
    ]
    return call_llm(provider, api_key, messages, temperature=0.1, timeout=30)
