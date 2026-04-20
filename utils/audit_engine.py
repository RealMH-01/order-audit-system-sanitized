# -*- coding: utf-8 -*-
"""
审核引擎模块
负责构造审核 prompt、交叉比对 prompt，以及解析大模型返回的 JSON 结果。
"""

import json
import logging
import re
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


# ============================================================
# 系统提示词（所有审核共用）
# ============================================================
_SYSTEM_PROMPT = """你是一个专业的外贸跟单单据审核助手。你的工作是帮助跟单员检查单据中的数据是否准确、是否与PO一致、是否与其他相关单据一致。你不是来批评跟单员的，你是来帮他们兜住问题的。

【审核核心原则】一切以PO为准。

【风险等级定义标准——必须严格遵守】

RED（红色/高风险）——与PO不一致的字段默认标RED：
  总规则：单据上的任何字段值与PO对应字段不一致，都标RED。这是默认规则，无需逐一列举。
  包括但不限于：合同号/Invoice No./订单号、PO号（当客户提供了独立PO号时）、
  产品名称/描述、数量、单价、总金额、币种、贸易术语（实质性变更如FOB→CIF）、
  收货人/买方名称及信息、目的港、HS编码、唛头、UN编号/危险品等级/包装组别、
  银行信息等。不允许以"可能是笔误""差异不大""可能是关联编号"等理由自行降级。
  特别强调：
  1. 合同号/Contract No./协议号/商业发票号(Invoice No.)/订单号是同一个编号，
     必须与PO上的合同号逐字符完全一致，哪怕只差一个字符也必须标RED，没有任何例外。
  2. 金额计算实际错误（单价×数量≠总金额等）必须标RED。
  3. 数量明确不符必须标RED。
  4. 币种不同必须标RED。
  5. 客户方（买方/申请人）的所有信息必须与PO严格一致，任何不一致标RED。
  6. 贸易术语实质性变更（如FOB→CIF，影响运费、保险费承担方和交货地点）必须标RED。
  7. 数字格式存在欧洲格式与英美格式的歧义（如1.234可能是一千二百三十四或一点二三四）必须标RED。

  唯一例外——以下情况不标RED，改标YELLOW：
  卖方/发货人的公司名称和地址——由于集团内部分工、子公司代签等行业惯例，
  若单据上的卖方/发货人公司名称或地址与PO不完全一致，但属于同一集团体系内的关联公司，
  标YELLOW提醒即可，不强制标RED。除此之外，所有与PO不一致的字段一律标RED。

YELLOW（黄色/需注意）——信息存在差异但可合理解释：
  1. 卖方/发货人/供应商/受益人的公司名称或地址与PO不同（集团内部分工、多地址等行业惯例）
  2. 计量单位不同但换算后金额一致（如KG与TON，1 TON = 1000 KG）
  3. 贸易术语本身相同、只是书写格式不同（如FOB SHANGHAI vs FOB Shanghai Port，术语都是FOB）
  4. PO号的比对规则：PO号与合同号/Invoice No.是不同的概念。
     如果PO文件上没有出现客户自己的独立PO号，单据上用合同号代替PO号是正常操作，标YELLOW提醒即可。
     如果PO文件上有客户自己的PO号，单据上的PO号应与之一致，不一致则标RED。
  5. 某个字段在PO中找不到对应信息

BLUE（蓝色/格式提醒）——纯格式或排版建议，不影响实际业务：
  如日期格式差异、大小写不统一、多余空格等"""


# ============================================================
# 审核 prompt 构造
# ============================================================
def build_audit_prompt(
    po_text: str,
    target_text: str,
    target_type: str,
    last_ticket_text: Optional[str] = None,
    template_text: Optional[str] = None,
    other_refs: Optional[List[str]] = None,
    deep_think: bool = False,
    custom_rules: Optional[str] = None,
) -> List[Dict[str, str]]:
    """构造单份单据审核的完整 messages 列表。

    Args:
        po_text: PO 的解析文本（必须有）。
        target_text: 待审核单据的解析文本。
        target_type: 单据类型，如 "商业发票CI"、"装箱单PL"、"托书Booking"、
                     "生产通知单"、"发货申请单"。
        last_ticket_text: 上一票对应文件的解析文本（可选）。
        template_text: 标准模板的解析文本（可选）。
        other_refs: 其他参考信息文本列表（可选）。
        deep_think: 是否启用深度思考模式。
        custom_rules: 【已废弃】自定义规则改为第二轮单独处理，此处不再注入。
                      参数保留仅为兼容历史调用方，不会被使用。

    Returns:
        可直接传给 call_llm 的 messages 列表。
    """
    # --- 构造 user 消息 ---
    parts: list[str] = []

    parts.append(f"【审核任务】\n请对以下{target_type}进行全面审核。")

    parts.append(f"\n【PO原文（审核核心依据）】\n{po_text}")

    parts.append(f"\n【待审核的{target_type}原文】\n{target_text}")

    if last_ticket_text:
        parts.append(
            f"\n【上一票对应文件原文（用于变更比对）】\n{last_ticket_text}"
        )

    if template_text:
        parts.append(
            f"\n【公司标准模板原文（用于格式比对）】\n{template_text}"
        )

    if other_refs:
        parts.append("\n【其他参考信息】")
        for i, ref in enumerate(other_refs, 1):
            parts.append(f"参考{i}：\n{ref}")

    # --- 审核规则 ---
    rules = _build_audit_rules(
        target_type=target_type,
        has_last_ticket=last_ticket_text is not None,
        has_template=template_text is not None,
    )
    parts.append(rules)

    user_content = "\n".join(parts)

    return [
        {"role": "system", "content": _SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]


def _build_audit_rules(
    target_type: str,
    has_last_ticket: bool,
    has_template: bool,
) -> str:
    """构造审核规则文本块。"""

    rules = f"""
【审核规则——请严格遵守】

【审核总则——最高优先级】
默认规则：单据字段与PO不一致 = RED。不需要逐个字段判断是否"实质性"，只要值不同就标RED，由人工决定是否接受。
唯一允许标YELLOW而非RED的情况：卖方/发货人公司名称和地址的集团内部差异。
其余所有字段的不一致，无论差异大小，一律RED。

一、源头比对（最高优先级）：
将待审核单据上的每一个数据字段与PO逐一比对，
包括但不限于：
合同号/Invoice No.、PO号、客户名称及地址、品名（中英文）、
数量、单价、总金额、包装方式及件数、毛重/净重、
贸易术语（FOB/CIF/FCA等）、
收货人/通知人信息（姓名、地址、电话、邮箱、TAX ID）、
卸货港/目的地、HS CODE、
UN编号/危险品等级/包装组别、银行信息、币种。
发现任何不一致都必须标记。

特别注意以下比对规则：
- 合同号/Invoice No./订单号比对（最高优先级，无例外）：
  重要前提：合同号(Contract No.)、商业发票号(Invoice No.)、订单号是同一个编号，只是在不同单据上叫法不同。
  单据上的这个编号必须与PO上的合同号完全一致，逐字符比对。
  只要有任何一个字符不同，无论差异多么微小，都必须标RED。
  绝对不允许以任何理由（包括但不限于"集团内部关联编号""子公司编号""可能是笔误"）将其降级为YELLOW或BLUE。
  这是硬性规则，没有例外。

- PO号比对（注意：PO号与合同号是不同的概念）：
  PO号是客户方的采购订单编号，与卖方的合同号/Invoice No.不是同一个号。
  如果PO文件上有客户自己的独立PO号，则单据上的PO号字段应与客户PO号一致，不一致标RED。
  如果PO文件上没有出现客户自己的独立PO号，单据上用合同号填入PO号字段是正常的替代做法，只标YELLOW提醒。
  PO号与合同号不一致本身不是错误（因为它们本来就是两个不同的编号体系），不要因此标RED。

- 贸易术语比对：必须区分"术语实质性变更"和"书写格式差异"。
  实质性变更（必须标RED）：如PO写FOB Shanghai Port，CI上出现CIF Shanghai Port
  或CIF Ningbo Port，这是贸易条款实质性改变（影响运费、保险费承担方和交货地点）。
  书写格式差异（只标YELLOW）：如FOB SHANGHAI vs FOB Shanghai Port vs FOB shanghai，
  术语都是FOB，只是大小写或写法不同。
- 客户方（买方/申请人）信息：必须与PO严格一致，任何不一致标RED。
- 我方（卖方/供应商/受益人）公司名称和地址：与PO不同属于正常现象（集团内部分工、
  多地址等行业惯例），只标YELLOW不标RED。这是唯一允许不标RED的例外情况。
- 计量单位比对：如KG与TON，1 TON = 1000 KG，换算后金额一致的不得标RED，只标YELLOW。

二、数值逻辑校验：
- 单价 × 数量 是否等于总金额
- 大写金额与数字金额是否一致
- 净重 + 包装皮重 是否约等于毛重
  （参考：IBC每个皮重约55kg，桶的皮重根据实际情况判断）
- 件数 × 每件净重 是否等于总净重

二（补充）、数字格式歧义检测（必须严格执行）：
当检测到数字格式可能存在欧洲格式（逗号作小数点、句点作千分位）与英美格式（句点作小数点、逗号作千分位）的歧义时，必须标为 RED 高风险。
例如：
- "1.234" 可能表示 1234（欧洲格式，句点是千分位）或 1.234（英美格式，句点是小数点）
- "1,234" 可能表示 1234（英美格式，逗号是千分位）或 1.234（欧洲格式，逗号是小数点）
- "12.345,67" 这种明确的欧洲格式与 "12,345.67" 英美格式之间的差异
判断标准：当一个数字中同时包含逗号和句点、或者在上下文中无法明确判断数字采用的是哪种计数格式、或者该数字按不同格式理解时数值相差巨大的，必须标RED，并在批注中明确说明：
"该数字存在格式歧义（可能为欧洲计数格式或英美计数格式），两种理解方式下数值相差巨大，请务必与客户确认具体数值，避免因数字格式误解导致损失。"
不允许AI自行判断采用哪种格式，必须标红交由人工确认。"""

    if has_last_ticket:
        rules += """

三、与上一票比对（已提供上一票文件）：
找出本票与上一票之间所有变化的字段，逐一列出。
客户的购买行为一般不变，所以任何变化都值得注意。"""

    if has_template:
        rules += """

四、与标准模板比对（已提供标准模板）：
检查抬头信息、公司地址、联系方式等固定内容是否符合标准模板。"""

    rules += """

五、特殊情况标记：
- PO中没有PO号，单据上用合同号代替的 → 黄色标记
- 发货申请单中发货要求写"无" → 黄色标记提醒确认
- 某个字段在PO中找不到对应信息 → 黄色标记并注明

【严重程度分级】
请严格按照系统提示中定义的 RED/YELLOW/BLUE 三级标准执行，不可混淆。
重申总规则：与PO不一致 = RED，唯一例外是卖方/发货人公司名称和地址。
特别提醒：贸易术语实质性变更（如FOB→CIF）必须标RED；仅大小写/写法差异只标YELLOW。
客户方（买方）信息不一致标RED。

【语气要求】
- 友好、专业，以"帮你兜住问题"的姿态
- 正确示例："此处金额与PO不一致，可能是笔误，建议核实后修改"
- 错误示例："错误：金额不正确"
- 不确定时："此处与上一票不同，请确认是否为客户新要求"

【输出格式——严格按以下JSON格式输出，不要有任何其他文字】
{
  "summary": {
    "total": 总标记数,
    "red": 红色数量,
    "yellow": 黄色数量,
    "blue": 蓝色数量
  },
  "issues": [
    {
      "id": "R-01",
      "level": "RED",
      "field_name": "字段中文名称",
      "field_location": "该字段在单据上大致位置的描述",
      "your_value": "待审核单据上填写的值",
      "source_value": "数据源中的原始值",
      "source": "数据来源说明，如PO第1页",
      "suggestion": "友好的中文建议"
    }
  ]
}

issues按字段在单据上从上到下出现的位置排列。
同一颜色内按位置顺序编号：R-01,R-02... Y-01,Y-02... B-01,B-02...
如果没有发现任何问题，返回 issues 为空列表，summary 各项为 0。"""

    return rules


# ============================================================
# 交叉比对 prompt 构造
# ============================================================
def build_cross_check_prompt(
    all_parsed_targets: List[Dict[str, str]],
    custom_rules: Optional[str] = None,
) -> List[Dict[str, str]]:
    """构造多单据交叉比对的完整 messages 列表。

    Args:
        all_parsed_targets: 列表，每个元素是
            {"type": 单据类型, "content": 解析文本}

    Args (续):
        custom_rules: 【已废弃】自定义规则改为第二轮单独处理，此处不再注入。
                      参数保留仅为兼容历史调用方，不会被使用。

    Returns:
        可直接传给 call_llm 的 messages 列表。
    """
    parts: list[str] = []

    parts.append("【交叉比对任务】")
    parts.append(
        "请检查以下多份单据之间，相同字段的数据是否一致。"
        "例如CI和PL的数量是否一致、CI和托书的合同号是否一致等。"
    )

    for i, target in enumerate(all_parsed_targets, 1):
        doc_type = target.get("type", f"单据{i}")
        content = target.get("content", "")
        parts.append(f"\n【单据{i}：{doc_type}】\n{content}")

    parts.append("""
【交叉比对规则】
1. 对比所有单据中出现的相同字段：
   合同号、PO号、客户名称、品名、数量、单价、总金额、
   包装方式、件数、毛重/净重、贸易术语、
   收货人/通知人、卸货港/目的地等。
2. 任何单据之间的数据不一致都需要标记。
3. 以第一份单据（通常是CI）为基准进行比对。
4. 注意数字格式歧义：当检测到数字格式可能存在欧洲格式与英美格式的歧义时，
   必须标为 RED 高风险，并在批注中明确说明格式歧义风险，
   不允许AI自行判断采用哪种格式，必须标红交由人工确认。

【严重程度分级】
- RED：不同单据之间同一字段数据不一致，可能导致业务错误；数字格式存在歧义
- YELLOW：数据存在差异但可能是合理的（如不同单据使用不同格式）

【语气要求】
- 友好、专业，以"帮你兜住问题"的姿态
- 示例："CI上的数量与PL上的数量不一致，建议核实"

【输出格式——严格按以下JSON格式输出，不要有任何其他文字】
{
  "summary": {
    "total": 总标记数,
    "red": 红色数量,
    "yellow": 黄色数量,
    "blue": 0
  },
  "issues": [
    {
      "id": "R-01",
      "level": "RED",
      "field_name": "字段中文名称",
      "field_location": "涉及的单据名称",
      "your_value": "单据A上的值",
      "source_value": "单据B上的值",
      "source": "涉及的具体单据说明",
      "suggestion": "友好的中文建议"
    }
  ]
}

如果所有单据之间数据完全一致，返回 issues 为空列表，summary 各项为 0。""")

    # 注：自定义规则不再在此处注入（同 build_audit_prompt），
    # 改为第二轮由 build_custom_rules_review_prompt 单独处理。

    user_content = "\n".join(parts)

    return [
        {"role": "system", "content": _SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]


# ============================================================
# 第二轮：自定义规则修正 prompt 构造
# ============================================================
_CUSTOM_RULES_REVIEW_SYSTEM_PROMPT = """你是一个审核结果修正助手。你的唯一任务是：根据用户的自定义规则，修正一份已有的审核结果。

你会收到：
1. 一份由系统内置规则生成的原始审核结果（JSON格式）
2. 用户设置的自定义补充规则

【你的工作流程】
逐条检查原始结果中的每个issue，判断它是否与自定义规则相关：
- 如果无关：保持不变，原样输出
- 如果相关且需要降级：修改level（如RED→YELLOW），修改suggestion
- 如果相关且需要升级：修改level（如YELLOW→RED），修改suggestion
- 如果相关且应删除：从issues列表中移除
- 如果自定义规则要求检查的内容不在原始结果中：新增issue

【绝对优先权规则——这是你必须遵守的最高原则】
自定义规则的权威性高于一切。当自定义规则说某种情况"正常""允许""必须如此"时，
无论原始结果中该issue的level是什么，你都必须将其降级为YELLOW或删除。
你没有权力保留RED，即使你认为这个问题"值得注意"或"需要确认"。

同样，当自定义规则要求某种情况必须标RED时，无论原始结果中该issue的level是什么，
你都必须升级为RED。

【不可降级保护规则】
除"卖方/发货人公司名称和地址的集团内部差异"外，
所有与PO不一致而被标RED的字段，一律不允许降级为YELLOW或删除。
自定义规则不能覆盖这条硬性规则，除非自定义规则中明确写了"忽略XX字段差异"这样的精确指令。
具体而言，以下类型的RED绝对不允许降级：
1. 合同号/Invoice No./订单号与PO上的合同号不一致（合同号、Invoice No.、订单号是同一个编号）
2. 金额计算错误
3. 数量不符
4. 币种不一致
5. 产品名称/描述不一致
6. 收货人/买方信息不一致
7. 目的港不一致
8. HS编码不一致
这些是单据的核心数据，一旦出错后果严重。
"集团内部分工""关联公司"等规则仅适用于卖方/发货人的公司名称和地址字段，
不适用于合同号、金额、数量、产品名等交易核心字段。
注意：PO号与合同号是不同的概念，不在此保护名单内，PO号的升降级遵循正常的自定义规则流程。

【具体示例——你必须严格参照执行】

示例1（降级场景——集团内部分工）：
  背景：某集团下有多家公司（如华瑞国际贸易、华瑞精细化工、华瑞新材料等），
  生产由A公司负责，对外销售统一由B公司负责。
  因此PO上的卖方是B公司（如华瑞国际贸易），但COA（质检证书）由实际生产方A公司（如华瑞精细化工）开具，
  这是正常的集团内部分工。
  自定义规则："戊二醛产品的COA必须由华瑞精细化工开具"
  原始结果中有一条：level=RED，field_name=卖方/供应商名称，your_value=华瑞精细化工，source_value=华瑞国际贸易
  正确做法：level改为YELLOW，suggestion写"COA由生产方华瑞精细化工开具，PO卖方为集团销售公司华瑞国际贸易，属于集团内部正常分工，无需处理。"

示例2（降级场景——COA上的关联字段差异）：
  当COA的开具方（如华瑞精细化工）与PO卖方（如华瑞国际贸易）属于同一集团但不同公司时，
  COA上出现的公司地址、联系方式、电话、传真等信息自然也会与PO上的不同（因为是不同的公司）。
  这些差异都是集团内部分工的正常结果，不是错误。
  正确做法：这类由开具方不同而导致的所有关联字段差异，都应降级为YELLOW，绝对不能标RED。
  suggestion示例："该地址为COA开具方华瑞精细化工的地址，与PO卖方华瑞国际贸易地址不同属正常，两者为同一集团不同公司。"

示例3（升级场景）：
  自定义规则："所有单据必须包含HS Code，缺失必须标RED"
  原始结果中有一条：level=YELLOW，field_name=HS Code，suggestion=PO中未找到对应信息
  正确做法：level改为RED，suggestion写"根据审核要求，HS Code为必填项，当前单据中缺失，请补充。"

示例4（删除场景）：
  自定义规则："忽略所有日期格式差异"
  原始结果中有一条：level=BLUE，field_name=日期格式
  正确做法：从issues列表中删除此条。

【输出要求】
- 输出修正后的完整JSON，格式与原始结果完全一致
- summary中的数字必须与修正后的issues列表一致（重新计数）
- id编号根据修正后的level重新编排（R-01,R-02...Y-01,Y-02...B-01,B-02...）
- suggestion中不要出现"降级""升级""原始结果"等元描述词汇
- 直接给出修正后的业务建议，语气友好专业"""


def build_custom_rules_review_prompt(
    original_result_json: str,
    custom_rules: str,
    target_filename: str,
) -> List[Dict[str, str]]:
    """构造第二轮自定义规则修正的 prompt。

    Args:
        original_result_json: 第一轮审核输出的完整JSON字符串。
        custom_rules: 用户设置的自定义审核规则文本。
        target_filename: 被审核的文件名（用于上下文说明）。

    Returns:
        可直接传给 call_llm 的 messages 列表。
    """
    user_content = f"""【被审核文件】{target_filename}

【原始审核结果（由系统内置规则生成）】
{original_result_json}

【用户自定义补充规则】
{custom_rules}

请根据自定义规则对上述审核结果进行修正。
如果自定义规则与某些issue无关，保持原样即可。
如果所有issue都不需要修正且不需要新增，直接原样输出即可。

输出格式必须与原始结果完全一致：
{{
  "summary": {{
    "total": 总标记数,
    "red": 红色数量,
    "yellow": 黄色数量,
    "blue": 蓝色数量
  }},
  "issues": [
    {{
      "id": "编号",
      "level": "RED/YELLOW/BLUE",
      "field_name": "字段名",
      "field_location": "位置",
      "your_value": "单据值",
      "source_value": "源值",
      "source": "数据来源",
      "suggestion": "修正后的建议"
    }}
  ]
}}

严格输出JSON，不要有任何其他文字。"""

    return [
        {"role": "system", "content": _CUSTOM_RULES_REVIEW_SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]


# ============================================================
# 审核结果 JSON 解析
# ============================================================
def parse_audit_result(llm_response: str) -> Optional[Dict[str, Any]]:
    """从大模型回复中提取并解析 JSON 结果。

    处理可能的情况：
    - 大模型在 JSON 前后加了多余文字或 markdown 代码块标记
    - JSON 格式有小问题（尝试修复）

    Args:
        llm_response: 大模型的原始回复字符串。

    Returns:
        解析后的 Python 字典，失败时返回 None。
    """
    if not llm_response or not llm_response.strip():
        logger.warning("大模型回复为空")
        return None

    text = llm_response.strip()

    # 尝试策略1：直接解析
    result = _try_parse_json(text)
    if result is not None:
        return _validate_audit_result(result)

    # 尝试策略2：提取 ```json ... ``` 代码块
    json_block = _extract_json_from_codeblock(text)
    if json_block:
        result = _try_parse_json(json_block)
        if result is not None:
            return _validate_audit_result(result)

    # 尝试策略3：提取第一个 { ... } 块
    json_obj = _extract_first_json_object(text)
    if json_obj:
        result = _try_parse_json(json_obj)
        if result is not None:
            return _validate_audit_result(result)

    # 尝试策略4：修复常见 JSON 问题后重试
    fixed = _fix_common_json_issues(text)
    if fixed:
        json_obj_fixed = _extract_first_json_object(fixed)
        if json_obj_fixed:
            result = _try_parse_json(json_obj_fixed)
            if result is not None:
                return _validate_audit_result(result)

    logger.warning("无法从大模型回复中解析JSON: %s", text[:200])
    return None


def _try_parse_json(text: str) -> Optional[Dict]:
    """尝试将文本解析为 JSON 字典。"""
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return data
    except (json.JSONDecodeError, ValueError):
        pass
    return None


def _extract_json_from_codeblock(text: str) -> Optional[str]:
    """从 markdown 代码块中提取 JSON。"""
    pattern = r"```(?:json)?\s*\n?(.*?)\n?\s*```"
    match = re.search(pattern, text, re.DOTALL)
    if match:
        return match.group(1).strip()
    return None


def _extract_first_json_object(text: str) -> Optional[str]:
    """提取文本中第一个完整的 JSON 对象 {...}。"""
    start = text.find("{")
    if start == -1:
        return None

    depth = 0
    in_string = False
    escape_next = False

    for i in range(start, len(text)):
        ch = text[i]

        if escape_next:
            escape_next = False
            continue

        if ch == "\\":
            if in_string:
                escape_next = True
            continue

        if ch == '"' and not escape_next:
            in_string = not in_string
            continue

        if in_string:
            continue

        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start : i + 1]

    return None


def _fix_common_json_issues(text: str) -> Optional[str]:
    """修复 JSON 中的常见格式问题。"""
    fixed = text

    # 移除尾部多余逗号 (如 ",}" 或 ",]")
    fixed = re.sub(r",\s*([}\]])", r"\1", fixed)

    # 替换中文引号为英文引号
    fixed = fixed.replace("\u201c", '"').replace("\u201d", '"')
    fixed = fixed.replace("\u2018", "'").replace("\u2019", "'")

    # 替换中文冒号为英文冒号
    fixed = fixed.replace("：", ":")

    return fixed


def _validate_audit_result(data: Dict) -> Optional[Dict]:
    """验证解析结果结构是否符合预期，并补全缺失字段。"""
    # 确保有 summary
    if "summary" not in data:
        data["summary"] = {"total": 0, "red": 0, "yellow": 0, "blue": 0}
    else:
        summary = data["summary"]
        for key in ("total", "red", "yellow", "blue"):
            if key not in summary:
                summary[key] = 0
            try:
                summary[key] = int(summary[key])
            except (ValueError, TypeError):
                summary[key] = 0

    # 确保有 issues 列表
    if "issues" not in data:
        data["issues"] = []
    elif not isinstance(data["issues"], list):
        data["issues"] = []

    # 校验每个 issue 的字段
    valid_issues = []
    for issue in data["issues"]:
        if not isinstance(issue, dict):
            continue
        issue.setdefault("id", "?-??")
        issue.setdefault("level", "YELLOW")
        issue.setdefault("field_name", "未知字段")
        issue.setdefault("field_location", "")
        issue.setdefault("your_value", "")
        issue.setdefault("source_value", "")
        issue.setdefault("source", "")
        issue.setdefault("suggestion", "")
        # 标准化 level
        issue["level"] = issue["level"].upper()
        if issue["level"] not in ("RED", "YELLOW", "BLUE"):
            issue["level"] = "YELLOW"

        valid_issues.append(issue)

    data["issues"] = valid_issues

    # 重新计算 summary
    red_count = sum(1 for i in valid_issues if i["level"] == "RED")
    yellow_count = sum(1 for i in valid_issues if i["level"] == "YELLOW")
    blue_count = sum(1 for i in valid_issues if i["level"] == "BLUE")
    data["summary"] = {
        "total": len(valid_issues),
        "red": red_count,
        "yellow": yellow_count,
        "blue": blue_count,
    }

    return data
