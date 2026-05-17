---
name: thesis-converter
description: >
  北航本科毕设 LaTeX→Word 论文转换技能。将 LaTeX 论文源码（.tex）按北航2026教务部模板规范转换为可直接提交的 .docx 文件。涵盖模板 XML 分析、OOXML 生成、格式修复、页眉/页码/文本框/VML处理全流程。触发：论文转换、格式修复、LaTeX转Word、毕设排版、北航模板。
metadata:
  short-description: LaTeX→北航Word论文转换(含模板XML/VML/页码/OOXML)
  triggers: 论文转换, LaTeX转Word, 毕设排版, 北航模板, thesis, docx conversion
  project: BUAA Undergraduate Thesis
  tools: latex_parser.py, gen_round1.py, ooxml_generator.py, omml_generator.py
---

# 北航毕设 LaTeX→Word 论文转换技能

将 LaTeX 论文源文件转换为符合北航2026教务部模板规范的 Word 文档，完整保留封面、摘要、正文、参考文献、图表公式。

## 核心管线

```
LaTeX .tex 文件
    │ latex_parser.py
    ▼ thesis_content.json (结构化内容)
    │ gen_round1.py
    ▼ thesis_r{N}.docx (最终输出)
    ├─ ooxml_generator.py  (段落/表格/公式/图片)
    ├─ omml_generator.py   (LaTeX→OMML 公式引擎)
    └─ 北航模板 DOCX        (Flat OPC XML 基座)
```

## 模板 XML 分析要诀

### 模板结构
- DOCX = ZIP 包，核心文件: `word/document.xml`, `word/styles.xml`, `word/header*.xml`
- PowerShell 提取: `Add-Type -AssemblyName System.IO.Compression.FileSystem; $zip = [System.IO.Compression.ZipFile]::OpenRead(...)`
- rId 映射在 `word/_rels/document.xml.rels`

### VML + DrawingML 双重表示 ⚠️
- 页眉文本框有 VML (`<v:shape style="height:28pt">`) + DrawingML (`<a:ext cy="355600">`) 两套
- VML 的 `o:gfxdata` 含 base64 ZIP，修改尺寸会触发 Word "无法读取内容"
- **原则**: 不要修改模板 header 的尺寸属性，用模板已有 header
- **正文 0.88cm 方案**: header20.xml (rId33, VML 18.9pt) 天然接近模板正文尺寸，仅注入 PAGE 域即可
- 前置用 header8.xml (rId17, VML 28pt/0.99cm)，正文用 header20.xml (rId33)

### 页眉 rId 映射速查
| rId | Header | VML 高度 | 用途 |
|-----|--------|---------|------|
| rId17 | header8.xml | 28pt(0.99cm) | 前置(摘要/TOC) |
| rId33 | header20.xml | 18.9pt(~0.88cm) | 正文全部 |
| rId16 | header8.xml | 同rId17 | 模板原名 |
| rId26 | header14.xml | 18.9pt | 模板正文(需PAGE域注入) |

### sectPr 节属性要点
- `pgNumType fmt="upperRoman"` → 罗马数字页码
- `pgNumType fmt="decimal" w:start="1"` → 阿拉伯数字从1开始
- `pgMar w:header="567" w:footer="851"` → 页眉/页脚距(模板标准)
- `headerReference r:id="rIdXX"` → 指定本节使用的页眉文件
- 封面节 `start="9"` 不应改为 upperRoman，需跳过

### 关键 XML 速查
| 功能 | XML 路径 | 示例 |
|------|---------|------|
| 文本框尺寸 | `wps:spPr > a:xfrm > a:ext` | `cx="5756910" cy="355600"` |
| 页眉横线 | `pBdr > bottom single` | `sz="6" space="0"` |
| 页码格式 | `sectPr > pgNumType` | `fmt="upperRoman"` / `fmt="decimal"` |
| 行距 | `spacing` | `line="360" lineRule="auto"` (1.5x) |
| 页码域 | `fldChar` + `instrText PAGE` | 动态页码 |

## 格式对照表

| 元素 | 字体 | 字号 | 样式 |
|------|------|------|------|
| H1 章标题 | 黑体 | 三号(sz=32) | 不加粗,居中, before=500 after=500 line=480 |
| H2/H3/H4 | 黑体 | 小四(sz=24) | 不加粗 |
| 正文 | 宋体 | 小四(sz=24) | 首行缩进480, line=360 |
| 图题 | 宋体 | 五号(sz=21) | 不加粗,居中 |
| 表题 | 黑体 | 五号(sz=21) | 居中 |
| 表头 | 宋体 | 五号(sz=21) | 不加粗 |
| 页眉 | 校名黑体 sz=28 | 页码宋体 sz=21 | |
| 引用 | [N] 上标 | | |
| 页码 | 前置 upperRoman | 正文 decimal start=1 | |

## 常见陷阱

1. **pStyle 粗体继承**: Style 2/4 自带 `<w:b/>`，必须 `<w:b w:val="0"/>` 显式关闭
2. **re.DOTALL 跨边界**: 匹配 `<wps:txbx>...</wps:txbx>` 时 `.*?` 可能跨过 `</wps:txbx>` 边界。先提取块再操作
3. **ET.fromstring 命名空间**: `xml.etree.ElementTree` 解析裸 `<w:r>` 静默丢内容，用 `_make_element` helper
4. **中文间换行顺序**: `[\u4e00-\u9fa5]\n[\u4e00-\u9fa5]` 必须在全局 `\n→空格` 之前执行
5. **VML base64 冲突**: 修改 header 尺寸须同时改 VML + DrawingML，否则被覆盖；同时改触发 "无法读取"
6. **分类号/单位代码**: 模板已有文本，只需替换占位值 (TN953→V224)，不要插入新 run
7. **小数点点被吃**: `clean_inline_latex` 的 `^\d+[.)、]` 中 `.` 会匹配十进制点，移除它
8. **numId 冲突**: Word 编号系统全局扁平，手动 `[N]` 前缀比自动编号更安全
9. **文本框行距**: `line=360 lineRule=auto` = 1.5x，Word 对浮动文本框行距计算与正文不同
10. **中文编码陷阱**: Python 源文件中文可能与模板 XML 中文不匹配，用 `chr(0x4f4d)+chr(0x4ee3)` 构造搜索串
11. **LaTeX 残留**: `\begin{enumerate\n}` (换行在`}`前) 需 `\\begin\{[^}]*\}?` 匹配部分缺失`}`的情况
12. **附录格式**: 附录章题用 `back_matter=True` (无编号)，每项标为"附录A XXX""附录B XXX"
13. **附录表格**: 表格在 chapter 级 (`appendix.tables`) 和 section 级都要遍历

## R75 最新基线
- 附录格式: 附录A/附录B (back_matter), tables distributed to sections
- **公式段落拆分**: `_replace_eq_with_ref` 用 `\n\n` 包裹 REF 标记，保留公式前后段落边界
- **附录解析**: `parse_thesis()` 显式解析 `data/appendix.tex`，存入 `result["appendix"]`
- **app_label 重置**: gen_round1.py 附录内容生成前必须 `app_label = ord('A')`（共享变量被 ref_map 段消费）
- REF 域已废弃: `\ref{}` → 静态号段（`ref_map['num']`），消除格式继承问题
- 浮体插入: 公式段前(原始行为) + 表格/图片段后(自然阅读序)
- 正文页眉: header20.xml(rId33) VML 18.9pt ~0.88cm
- 前置页眉: header8.xml(rId17) VML 28pt 0.99cm

## 关键教训追加 (R59-R74)
14. **公式段落边界**: `_replace_eq_with_ref` 必须用 `\n\n` 包裹 `[REF:eq:xxx]` 标记，否则公式前后文本被合并为一个段落。gen_round1.py 将 REF-only 段落清洗为空后跳过，形成 引导文本段→公式段→后续文本段 的正确结构。
15. **附录数据源**: `parse_thesis()` 必须显式读取 `data/appendix.tex` 并调用 `parse_chapter()` 解析。仅 gen_round1.py 有附录生成逻辑但 JSON 无 `appendix` 键 → 附录静默丢失。
16. **app_label 共享变量陷阱**: gen_round1.py 中 R67 ref_map 构建段会 `app_label += 1` 遍历附录节，R66 附录内容生成段复用同一变量但未重置。**修复**: 在 R66 段第一行加 `app_label = ord('A')`。
17. **REF 域格式继承**: Word REF 域继承书签源格式(sz=21/粗体)。`\* MERGEFORMAT`/`\* Charformat` 均不稳定。**终极方案**: 用 `ref_map['num']` 静态替换 `\ref{}`
18. **浮体插入顺序**: 公式→段前, 表格/图片→段后。用 `_pending_floats` defer 非公式浮体
19. **附录表格分布**: parser 将所有表放 chapter 级，需手动分配到 sections 避免全挤最后
14. **页眉页码间距**: header20 的 PAGE 域注入后不要在 _new20 中追加 space run（会与"页"字前的空格叠加成4空格）。间距统一由直接修改"页" run 实现（`<w:t xml:space="preserve">  页</w:t>` = 2空格）
15. **文本框宽度溢出**: header20 的"第"后仅保留模板原始 1 个空格（模板为 `第 48 `），多加空格会导致双数字页码（第2章+）超出 WPS 文本框 426.1pt 宽度触发换行
18. **upperRoman 不默认起始 I**: OOXML 中 `fmt="upperRoman"` 无 `start` 属性时从上一节累计计数继续，非重置为 I。前端页首节需显设 `start="1"`

## 页眉文本框 0.88cm 专项（R50-R61 教训）
- 模板正文 header(14-22) VML 18.9pt，Word 显示~0.88cm
- **失败尝试**: 直接改 header8 VML (28pt→24.95pt) → Word "无法读取内容"；改 DrawingML 不同步 VML → 被覆盖
- **失败尝试**: 新建 header_body.xml → Word 修复删除
- **失败尝试**: 修改 header14.xml 注入 PAGE 域 → 页码失效(原为静态"1")
- **成功方案(R61)**: 用模板已有 header20.xml(rId33)，仅注入动态 PAGE 域 + 补"页"字前导空格
- 正文节断 headerReference 全部指向 rId33，前置保留 rId17

## 分页符嵌入技巧
- 参考文献前需分页：在标题段落首 run 中嵌入 `<w:r><w:br w:type="page"/></w:r>`
- 体区替换后单独追加 `<w:p>` 可能被 lxml 丢弃，嵌入到 make_paragraph 返回的段落中更可靠

## 反馈单模板

```markdown
# 转化反馈单 — Round {N}
> **基线**: R{N-1} | **输出**: thesis_r{N}.docx

## 一、本轮处理的问题
### 1. 问题描述 ✅
**根因**: ... | **修复**: ... | **验证**: ...

## 二、人工反馈区
### 反馈 1: 问题
- [ ] ✅ ... - [ ] ❌ 需修改：___

## 三、代办更新（反馈单中永远保留此模块！）
| 编号 | 优先级 | 任务 | 状态 |
```

## 验证方法

```powershell
# 提取 XML 并计数
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead("thesis.docx")
$doc = (New-Object System.IO.StreamReader($zip.GetEntry("word/document.xml").Open())).ReadToEnd()

# 验证项
([regex]::Matches($doc, 'bottom="single"')).Count  # 页眉横线
([regex]::Matches($doc, 'fmt="upperRoman"')).Count  # 罗马页码
([regex]::Matches($doc, 'fmt="decimal"')).Count     # 阿拉伯页码
([regex]::Matches($doc, 'w:header="567"')).Count    # 页眉间距
```

## 轮次记录 (R1→R65)

| 里程碑 | 轮次 | 关键修复 |
|--------|------|---------|
| 核心管线 | R1-R17 | 引用映射, TOC, OMML, PAGE域 |
| 格式修复 | R18-R40 | 去粗, 表样式, 小数点, 分类号, LaTeX残留 |
| 页码页码 | R34-R37 | footer5清理, 封面节跳过, 正文start=1 |
| 页眉横线 | R31-R33 | text box线保留, 镜像段落中性和 |
| 文本框 | R50-R59 | 0.88cm尝试(失败), header20替代方案 |
| 稳定基线 | R61 | header20(正文0.88cm) + header8(前置0.99cm) |
| 页眉间距 | R65 | header20 页码-页字间距修正(4空格→2空格) |
| 章2换行 | R66 | header20 "第"后空格还原(2→1) |
| 模板还原 | R67 | header20 恢复模板 R3/R4 run 结构: trailing-space run + 保留原始页 run(含kern=0) |
| 宽度补偿 | R68 | 2+2空格自适应: 垫片R2 9→7空格回收~24pt |
| 宽度补偿2 | R69 | 修正西文空格宽度补偿(R1 4→2, R2 9→4), 正则规避XML属性顺序 |
| 分节修复 | R70 | _SECT_BREAK_XML 移除冗余 w:br page, 改用 w:type=nextPage (WPS兼容) |
| 加宽文本框 | R71 | header20 VML宽度 426.1→442.8pt (15.62cm), 移除o:gfxdata避校验冲突 |
| 禁用自适应 | R72 | mso-fit-shape-to-text:t→f, 防WPS自动缩窄文本框 |
| 切换页眉 | R73 | 正文节从header20(rId33)切换到header8(rId16), 宽度453.3pt×28pt |
| 摘要页码 | R74 | 前端页upperRoman跳过条件细化: start=9跳过, start=1(摘要)转为Roman |
| 漏洞修复 | R75 | _BODY_START_SECTPR fallback封装w:p防cleanup误删 |
| 罗马起始 | R76 | upperRoman不默认从I开始, 需显式start=1; 仅首节设start, 后续延续 |
