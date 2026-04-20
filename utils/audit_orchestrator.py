"""
审核流程总调度器
协调文件解析、图片 OCR、单据审核、交叉比对的完整流程。
新增：Token 长度检测与分段处理、超时控制、细粒度进度反馈、取消审核支持。
"""

import json as _json
import logging
import time
from typing import Any, Callable, Dict, List, Optional

from utils.audit_engine import (
    build_audit_prompt,
    build_cross_check_prompt,
    build_custom_rules_review_prompt,
    parse_audit_result,
)
from utils.file_parser import parse_file
from utils.llm_client import (
    LLMError,
    call_llm,
    call_llm_with_image,
    IMAGE_OCR_PROMPT,
    TIMEOUT_NORMAL,
    TIMEOUT_DEEP_THINK,
)
from utils.token_utils import (
    estimate_tokens,
    smart_split_content,
    get_safe_token_limit,
)

logger = logging.getLogger(__name__)

# 从文件名推断单据类型的映射表
_DOC_TYPE_HINTS = {
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
    "报关": "报关单",
    "customs": "报关单",
    "coa": "COA质检证书",
    "质检": "COA质检证书",
    "certificate": "COA质检证书",
}


def _guess_doc_type(filename: str) -> str:
    """根据文件名猜测单据类型。"""
    name_lower = filename.lower()
    for keyword, doc_type in _DOC_TYPE_HINTS.items():
        if keyword in name_lower:
            return doc_type
    return "待审核单据"


# ============================================================
# ★ 代码层面兜底：强制修正"说了没问题却标RED"的矛盾issue
# ============================================================
_POSITIVE_KEYWORDS = [
    "符合", "正常", "没有问题", "无需处理", "属于正常",
    "这一点是符合的", "完全符合", "不构成错误",
    "合理的", "可以接受", "无误", "无需修改",
    "正确的", "没有错误", "属正常", "不影响",
    "无需更改", "合规", "一致的", "匹配",
    "集团内部正常分工", "正常的集团", "内部分工",
    "是为了确认", "已被正确执行",
]

# ★ 只有这些字段相关的RED才允许被降级为YELLOW
# 其余所有字段的RED一律不降级
_DOWNGRADE_ALLOWED_FIELDS = [
    "卖方", "卖家", "seller", "买方", "买家", "buyer",
    "发货人", "shipper", "收货人", "consignee",
    "地址", "address", "签章", "公司名", "抬头",
    "通知方", "notify",
]


def _post_process_force_downgrade(audit_result: dict) -> dict:
    """对审核结果做最终兜底修正。

    只处理一种情况：level=RED 且字段属于公司主体/地址类，
    同时 suggestion 中明确表示"符合""正常""没有问题"等肯定性表述。

    交易核心字段（合同号、金额、数量、产品名等）的RED绝不降级。
    """
    if not audit_result or "issues" not in audit_result:
        return audit_result

    changed = False
    for issue in audit_result["issues"]:
        if issue.get("level") != "RED":
            continue
        # 先检查字段是否属于允许降级的范围（公司主体/地址类）
        field = issue.get("field_name", "").lower()
        if not any(kw in field for kw in _DOWNGRADE_ALLOWED_FIELDS):
            continue
        # 只有允许降级的字段，才进一步检查suggestion
        sugg = issue.get("suggestion", "")
        if any(kw in sugg for kw in _POSITIVE_KEYWORDS):
            issue["level"] = "YELLOW"
            changed = True
            logger.info(
                "兜底降级: issue '%s' 从RED降为YELLOW (suggestion含肯定表述: %s)",
                issue.get("field_name", "?"),
                sugg[:80],
            )

    if changed:
        # 重新编号和计数
        reds = [i for i in audit_result["issues"] if i["level"] == "RED"]
        yellows = [i for i in audit_result["issues"] if i["level"] == "YELLOW"]
        blues = [i for i in audit_result["issues"] if i["level"] == "BLUE"]

        for idx, issue in enumerate(reds, 1):
            issue["id"] = f"R-{idx:02d}"
        for idx, issue in enumerate(yellows, 1):
            issue["id"] = f"Y-{idx:02d}"
        for idx, issue in enumerate(blues, 1):
            issue["id"] = f"B-{idx:02d}"

        audit_result["issues"] = reds + yellows + blues
        audit_result["summary"] = {
            "total": len(audit_result["issues"]),
            "red": len(reds),
            "yellow": len(yellows),
            "blue": len(blues),
        }

    return audit_result


def run_full_audit(
    provider: str,
    api_key: str,
    po_data: Dict[str, Any],
    target_files_data: List[Dict[str, Any]],
    last_ticket_data: Optional[List[Dict[str, Any]]] = None,
    template_data: Optional[Dict[str, Any]] = None,
    other_refs_data: Optional[List[Dict[str, Any]]] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
    deep_think: bool = False,
    zhipu_ocr_api_key: Optional[str] = None,
    custom_rules: Optional[str] = None,
) -> Dict[str, Any]:
    """执行完整审核流程。

    Args:
        provider: 模型提供商名称。
        api_key: API 密钥。
        po_data: PO 文件的解析结果 dict。
        target_files_data: 待审核文件的解析结果 dict 列表。
        last_ticket_data: 上一票文件的解析结果 dict 列表（可选）。
        template_data: 标准模板的解析结果 dict（可选）。
        other_refs_data: 其他参考文件的解析结果 dict 列表（可选）。
        progress_callback: 进度回调函数，接收字符串参数。
        cancel_check: 取消检查函数，返回 True 表示用户要求取消。
        deep_think: 是否启用深度思考模式。

    Returns:
        {
            "per_file_results": {
                "文件名1": {审核结果字典},
                "文件名2": {审核结果字典},
            },
            "cross_check_result": {交叉比对结果字典} 或 None,
            "errors": ["无法审核的文件及原因列表"],
            "token_warning": "",  # Token 长度警告信息
            "cancelled": False,   # 是否被用户取消
        }
    """

    def _progress(msg: str) -> None:
        if progress_callback:
            progress_callback(msg)
        logger.info(msg)

    def _is_cancelled() -> bool:
        if cancel_check and cancel_check():
            return True
        return False

    result: Dict[str, Any] = {
        "per_file_results": {},
        "cross_check_result": None,
        "errors": [],
        "token_warning": "",
        "cancelled": False,
    }

    # ==========================================================
    # 步骤 1：处理图片 OCR
    # ==========================================================
    _progress("正在准备审核数据...")

    if _is_cancelled():
        result["cancelled"] = True
        return result

    # ==========================================================
    # 判断用于 OCR 的 provider 和 api_key（支持混合模式）
    # ==========================================================
    has_scanned_pdf = (
        (po_data.get("is_scanned_pdf") and po_data.get("pdf_page_images"))
        or any(t.get("is_scanned_pdf") and t.get("pdf_page_images") for t in target_files_data)
    )
    has_ref_images = (
        other_refs_data
        and any(d.get("is_image") and d.get("image_base64") for d in other_refs_data)
    )
    needs_ocr = has_scanned_pdf or has_ref_images

    if provider.lower().strip() in ("deepseek",):
        if needs_ocr:
            if zhipu_ocr_api_key and zhipu_ocr_api_key.strip():
                ocr_provider = "智谱GLM"
                ocr_api_key = zhipu_ocr_api_key
                _progress("检测到需要图片识别，将使用智谱GLM进行OCR，文字审核仍使用DeepSeek...")
            else:
                if has_scanned_pdf:
                    error_msg = (
                        "检测到上传的PDF为扫描件（图片型PDF），需要AI图片识别（OCR）来提取文字。"
                        "DeepSeek不支持图片识别。请在左侧边栏填写「智谱OCR密钥」后重新审核，"
                        "或将大模型切换为「智谱GLM」。"
                    )
                    _progress("❌ " + error_msg)
                    result["errors"].append(error_msg)
                    return result
                else:
                    ocr_provider = provider
                    ocr_api_key = api_key
                    _progress("⚠️ 参考图片需要OCR但未提供智谱OCR密钥，将跳过图片识别")
        else:
            ocr_provider = provider
            ocr_api_key = api_key
    else:
        ocr_provider = provider
        ocr_api_key = api_key

    # 处理参考图片的 OCR
    other_refs_texts: list[str] = []
    if other_refs_data:
        images_to_ocr = [
            d for d in other_refs_data if d.get("is_image") and d.get("image_base64")
        ]
        skip_image_ocr = (
            provider.lower().strip() in ("deepseek",)
            and not (zhipu_ocr_api_key and zhipu_ocr_api_key.strip())
        )
        if images_to_ocr and not skip_image_ocr:
            _progress(f"正在识别截图内容...（共 {len(images_to_ocr)} 张图片）")
            for img in images_to_ocr:
                if _is_cancelled():
                    result["cancelled"] = True
                    return result

                fname = img.get("filename", "未知图片")
                _progress(f"正在识别: {fname}")
                try:
                    ocr_text = call_llm_with_image(
                        ocr_provider, ocr_api_key, IMAGE_OCR_PROMPT, img["image_base64"]
                    )
                    img["content"] = ocr_text
                    other_refs_texts.append(f"[{fname}]\n{ocr_text}")
                except LLMError as e:
                    err_msg = f"图片 {fname} 识别失败: {e.message}"
                    result["errors"].append(err_msg)
                    _progress(f"⚠️ {err_msg}")

        for d in other_refs_data:
            if not d.get("is_image") and d.get("content") and d.get("success"):
                other_refs_texts.append(
                    f"[{d.get('filename', '参考文件')}]\n{d['content']}"
                )

    # ==========================================================
    # 步骤 1.2：扫描件PDF自动OCR
    # ==========================================================

    # --- 处理 PO 扫描件 ---
    if po_data.get("is_scanned_pdf") and po_data.get("pdf_page_images"):
        page_images = po_data["pdf_page_images"]
        total_pages = len(page_images)
        _progress(f"检测到PO为扫描件PDF，正在进行AI-OCR识别...（共 {total_pages} 页）")
        ocr_parts: list[str] = []
        for pg_idx, page_b64 in enumerate(page_images, start=1):
            if _is_cancelled():
                result["cancelled"] = True
                _progress("⚠️ 审核已被用户取消")
                return result

            _progress(f"正在识别PO扫描件第 {pg_idx}/{total_pages} 页...")
            try:
                ocr_text = call_llm_with_image(
                    ocr_provider, ocr_api_key, IMAGE_OCR_PROMPT, page_b64
                )
                ocr_parts.append(f"{'='*20} 第 {pg_idx} 页 {'='*20}\n{ocr_text}")
            except LLMError as e:
                err_msg = f"PO扫描件第 {pg_idx} 页识别失败: {e.message}"
                result["errors"].append(err_msg)
                _progress(f"⚠️ {err_msg}")
                ocr_parts.append(f"{'='*20} 第 {pg_idx} 页 {'='*20}\n[识别失败]")

        po_data["content"] = "\n\n".join(ocr_parts)
        _progress("✅ PO扫描件OCR识别完成")

    # --- 处理待审核文件中的扫描件 ---
    for target in target_files_data:
        if _is_cancelled():
            result["cancelled"] = True
            _progress("⚠️ 审核已被用户取消")
            return result

        if not target.get("is_scanned_pdf") or not target.get("pdf_page_images"):
            continue

        t_fname = target.get("filename", "未知文件")
        t_page_images = target["pdf_page_images"]
        t_total_pages = len(t_page_images)
        _progress(f"检测到 {t_fname} 为扫描件PDF，正在进行AI-OCR识别...（共 {t_total_pages} 页）")
        t_ocr_parts: list[str] = []
        all_failed = True
        for pg_idx, page_b64 in enumerate(t_page_images, start=1):
            if _is_cancelled():
                result["cancelled"] = True
                _progress("⚠️ 审核已被用户取消")
                return result

            _progress(f"正在识别 {t_fname} 第 {pg_idx}/{t_total_pages} 页...")
            try:
                ocr_text = call_llm_with_image(
                    ocr_provider, ocr_api_key, IMAGE_OCR_PROMPT, page_b64
                )
                t_ocr_parts.append(f"{'='*20} 第 {pg_idx} 页 {'='*20}\n{ocr_text}")
                all_failed = False
            except LLMError as e:
                err_msg = f"{t_fname} 第 {pg_idx} 页识别失败: {e.message}"
                result["errors"].append(err_msg)
                _progress(f"⚠️ {err_msg}")
                t_ocr_parts.append(f"{'='*20} 第 {pg_idx} 页 {'='*20}\n[识别失败]")

        target["content"] = "\n\n".join(t_ocr_parts)
        if all_failed:
            target["success"] = False
        else:
            target["success"] = True
        _progress(f"✅ {t_fname} 扫描件OCR识别完成")

    # 准备各文本
    po_text = po_data.get("content", "")
    template_text = (
        template_data.get("content", "") if template_data and template_data.get("success") else None
    )
    # 按单据类型索引上一票文件，避免把所有上一票文件拼成一个大文本
    last_ticket_by_type: Dict[str, str] = {}
    if last_ticket_data:
        for d in last_ticket_data:
            if d.get("content") and d.get("success"):
                doc_type = _guess_doc_type(d.get("filename", ""))
                last_ticket_by_type[doc_type] = d["content"]

    # ==========================================================
    # 步骤 1.5：Token 长度检测与智能分段处理
    # ==========================================================
    _progress("正在检测内容长度...")

    token_warning_issued = False
    for target in target_files_data:
        target_content = target.get("content", "")
        if not target_content or not target.get("success"):
            continue

        t_fname = target.get("filename", "")
        t_type = _guess_doc_type(t_fname)
        matched_last = last_ticket_by_type.get(t_type)

        pre_aux_texts: list[str] = []
        if matched_last:
            pre_aux_texts.append(matched_last)
        if template_text:
            pre_aux_texts.append(template_text)
        pre_aux_texts.extend(other_refs_texts)

        po_proc, target_proc, aux_proc, was_truncated = smart_split_content(
            po_text=po_text,
            target_text=target_content,
            other_texts=pre_aux_texts,
            provider=provider,
        )

        if was_truncated and not token_warning_issued:
            token_warning_issued = True
            warning_msg = (
                "⚠️ 文件内容较长，已自动优化处理，审核结果可能不如短文件精确。"
                "建议减少单次上传文件数量或拆分较长的文件。"
            )
            result["token_warning"] = warning_msg
            _progress(warning_msg)

    # ==========================================================
    # 步骤 2：逐份审核每个待审核文件
    # ==========================================================
    successful_targets: list[dict] = []
    total_files = len(target_files_data)

    for idx, target in enumerate(target_files_data, 1):
        if _is_cancelled():
            result["cancelled"] = True
            _progress("⚠️ 审核已被用户取消")
            return result

        fname = target.get("filename", f"文件{idx}")
        target_content = target.get("content", "")
        target_type = _guess_doc_type(fname)

        if not target_content or not target.get("success"):
            err_msg = f"{fname}: 文件解析失败，无法审核"
            result["errors"].append(err_msg)
            continue

        start_time = time.time()
        _progress(f"正在审核第 {idx}/{total_files} 份文件：{fname}...")

        matched_last_ticket = last_ticket_by_type.get(target_type)

        auxiliary_texts: list[str] = []
        if matched_last_ticket:
            auxiliary_texts.append(matched_last_ticket)
        if template_text:
            auxiliary_texts.append(template_text)
        auxiliary_texts.extend(other_refs_texts)

        po_processed, target_processed, aux_processed, _ = smart_split_content(
            po_text=po_text,
            target_text=target_content,
            other_texts=auxiliary_texts,
            provider=provider,
        )

        last_ticket_processed = None
        template_processed = None
        other_refs_processed = []
        aux_idx = 0
        if matched_last_ticket and aux_idx < len(aux_processed):
            last_ticket_processed = aux_processed[aux_idx]
            aux_idx += 1
        if template_text and aux_idx < len(aux_processed):
            template_processed = aux_processed[aux_idx]
            aux_idx += 1
        if aux_idx < len(aux_processed):
            other_refs_processed = aux_processed[aux_idx:]

        messages = build_audit_prompt(
            po_text=po_processed,
            target_text=target_processed,
            target_type=target_type,
            last_ticket_text=last_ticket_processed,
            template_text=template_processed,
            other_refs=other_refs_processed if other_refs_processed else None,
            deep_think=deep_think,
            custom_rules=custom_rules,
        )

        audit_result = _call_and_parse(
            provider, api_key, messages, fname, result["errors"],
            deep_think=deep_think,
            progress_callback=lambda msg, f=fname, i=idx, t=total_files, st=start_time: _progress(
                f"正在审核第 {i}/{t} 份文件：{f}（已耗时 {int(time.time() - st)} 秒）— {msg}"
            ),
        )

        # ★ 第二轮：自定义规则修正
        if audit_result is not None and custom_rules and custom_rules.strip():
            _progress(f"正在根据自定义规则修正 {fname} 的审核结果...")
            try:
                original_json_str = _json.dumps(audit_result, ensure_ascii=False, indent=2)

                review_messages = build_custom_rules_review_prompt(
                    original_result_json=original_json_str,
                    custom_rules=custom_rules,
                    target_filename=fname,
                )

                review_result = _call_and_parse(
                    provider, api_key, review_messages,
                    f"{fname}(自定义规则修正)", result["errors"],
                    deep_think=False,
                )

                if review_result is not None:
                    audit_result = review_result
                    _progress(f"✅ {fname} 自定义规则修正完成")
                else:
                    _progress(f"⚠️ {fname} 自定义规则修正失败，使用原始审核结果")
            except Exception as e:
                logger.warning("自定义规则修正异常: %s", e)
                _progress(f"⚠️ {fname} 自定义规则修正异常，使用原始审核结果")

        # ★ 第三步：代码兜底——强制修正自相矛盾的标记
        if audit_result is not None:
            audit_result = _post_process_force_downgrade(audit_result)

        elapsed_final = time.time() - start_time
        if audit_result is not None:
            audit_result["original_text"] = target_content
            result["per_file_results"][fname] = audit_result
            successful_targets.append(
                {"type": target_type, "content": target_content}
            )
            _progress(
                f"✅ {fname} 审核完成（耗时 {int(elapsed_final)} 秒）"
            )
        else:
            _progress(
                f"❌ {fname} 审核失败（耗时 {int(elapsed_final)} 秒）"
            )

    # ==========================================================
    # 步骤 3：交叉比对（仅当有多份待审核文件成功时）
    # ==========================================================
    if _is_cancelled():
        result["cancelled"] = True
        _progress("⚠️ 审核已被用户取消")
        return result

    if len(successful_targets) >= 2:
        _progress("正在进行单据间交叉比对...")
        cross_start = time.time()
        cross_messages = build_cross_check_prompt(successful_targets, custom_rules=custom_rules)
        cross_result = _call_and_parse(
            provider, api_key, cross_messages, "交叉比对", result["errors"],
            deep_think=deep_think,
        )
        cross_elapsed = time.time() - cross_start
        _progress(f"✅ 交叉比对完成（耗时 {int(cross_elapsed)} 秒）")

        # ★ 交叉比对的自定义规则修正（第二轮）
        if cross_result is not None and custom_rules and custom_rules.strip():
            _progress("正在根据自定义规则修正交叉比对结果...")
            try:
                cross_json_str = _json.dumps(cross_result, ensure_ascii=False, indent=2)

                review_messages = build_custom_rules_review_prompt(
                    original_result_json=cross_json_str,
                    custom_rules=custom_rules,
                    target_filename="交叉比对",
                )

                review_cross = _call_and_parse(
                    provider, api_key, review_messages,
                    "交叉比对(自定义规则修正)", result["errors"],
                    deep_think=False,
                )

                if review_cross is not None:
                    cross_result = review_cross
                    _progress("✅ 交叉比对自定义规则修正完成")
                else:
                    _progress("⚠️ 交叉比对自定义规则修正失败，使用原始结果")
            except Exception as e:
                logger.warning("交叉比对自定义规则修正异常: %s", e)
                _progress("⚠️ 交叉比对自定义规则修正异常，使用原始结果")

        # ★ 交叉比对也做代码兜底
        if cross_result is not None:
            cross_result = _post_process_force_downgrade(cross_result)

        result["cross_check_result"] = cross_result

    _progress("审核完成！")
    return result


def _call_and_parse(
    provider: str,
    api_key: str,
    messages: List[Dict],
    file_label: str,
    errors: list,
    max_retries: int = 2,
    deep_think: bool = False,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Optional[Dict[str, Any]]:
    """调用大模型并解析 JSON 结果，失败时自动重试。"""
    for attempt in range(1, max_retries + 1):
        try:
            if progress_callback:
                if attempt == 1:
                    progress_callback("正在等待AI响应...")
                else:
                    progress_callback(f"第 {attempt} 次重试...")

            llm_response = call_llm(
                provider, api_key, messages,
                temperature=0.1,
                deep_think=deep_think,
            )
            parsed = parse_audit_result(llm_response)
            if parsed is not None:
                return parsed

            if attempt < max_retries:
                logger.warning(
                    "[%s] 第%d次尝试: JSON解析失败，准备重试", file_label, attempt
                )
                truncated_response = (
                    llm_response[:500] + "..." if len(llm_response) > 500 else llm_response
                )
                messages = messages + [
                    {"role": "assistant", "content": truncated_response},
                    {
                        "role": "user",
                        "content": (
                            "你的回复无法被解析为JSON格式。"
                            "请严格按照要求的JSON格式重新输出结果，"
                            "不要包含任何JSON以外的文字。"
                        ),
                    },
                ]
                continue
            else:
                err_msg = f"{file_label}: AI返回结果格式异常，无法解析"
                errors.append(err_msg)
                logger.error("[%s] JSON解析最终失败: %s", file_label, llm_response[:300])
                return None

        except LLMError as e:
            if attempt < max_retries:
                logger.warning(
                    "[%s] 第%d次尝试失败: %s，准备重试",
                    file_label,
                    attempt,
                    e.message,
                )
                continue
            err_msg = f"{file_label}: {e.message}"
            errors.append(err_msg)
            return None
