# 订单审核系统 - AI-OCR 扫描件PDF识别功能

## 新增功能说明

本次修改为项目添加了 **AI-OCR 自动识别扫描件PDF** 的能力。当用户上传扫描件PDF（图片型PDF）时，系统会自动检测并将每页转为图片，然后调用项目已有的AI大模型接口进行OCR文字识别。

## 修改/新建文件清单（共4个文件）

### 1. `packages.txt` ✨ 新建
- 内容：`poppler-utils`
- 用途：Streamlit Community Cloud 安装系统级依赖，`pdf2image` 需要 poppler

### 2. `requirements.txt` 📝 修改
- 末尾新增 `pdf2image` 依赖

### 3. `utils/file_parser.py` 📝 修改（4处改动）
1. **import 区域**：新增 `from pdf2image import convert_from_bytes`
2. **`parse_pdf()` 函数**：新增 `all_pages_empty` 变量追踪是否所有页面都没有文字，如果全空则返回 `[扫描件]` 前缀
3. **`parse_file()` 函数**：PDF 分支增加扫描件检测逻辑，设置 `is_scanned_pdf` 和 `pdf_page_images` 字段
4. **新增 `_pdf_to_images_base64()` 函数**：将PDF每页转为图片并返回base64编码列表

### 4. `utils/audit_orchestrator.py` 📝 修改
- 在 `run_full_audit()` 中新增"步骤 1.2：扫描件PDF自动OCR"
- 处理PO扫描件：逐页调用 `call_llm_with_image` 进行OCR，结果写回 `po_data["content"]`
- 处理待审核文件中的扫描件：同样逻辑，OCR结果写回 `target["content"]`
- 支持进度回调、取消检查、单页失败不中断

## 数据流说明

```
扫描件PDF上传
    → parse_file() → parse_pdf() 检测到全空 → 返回 "[扫描件] ..."
    → is_scanned_pdf = True
    → _pdf_to_images_base64() 转图片
    → pdf_page_images 存入 result dict
    
用户点击审核
    → audit_orchestrator 步骤1.2 检测 is_scanned_pdf
    → 逐页调用 call_llm_with_image 做OCR
    → OCR结果写回 content
    → 后续正常审核流程
```

## parse_file() 返回值新增字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `is_scanned_pdf` | `bool` | 是否为扫描件PDF（仅PDF类型有值） |
| `pdf_page_images` | `List[str]` | 每页图片的base64编码（仅扫描件有值） |

## 注意事项

- 现有功能完全不受影响（正常文本型PDF走原有流程）
- 新增字段使用 `.get()` 访问，向后兼容
- `pdf2image` 在 Streamlit Community Cloud 上需要 `poppler-utils` 系统依赖
- 最多处理PDF前10页（`max_pages=10`），防止超大PDF消耗过多资源
