#  📜 遗嘱 / HANDOFF — Stage R1-R58

> **写给**: 下一轮接手此管线的人类或 Agent  
> **作者**: Sisyphus (R1-R58, 2026-05-17)  
> **状态**: R58 稳定基线  
> **上一份遗嘱**: `HANDOFF_遗嘱_必读.md` (2026-05-14, 已过时——本文件覆盖)

---

## 一、我完成了什么（R1→R31，31轮迭代）

### 核心管线
| 成果 | 文件 | 说明 |
|------|------|------|
| ✅ 主转化脚本 | `output1/gen_round1.py` (62KB) | 单文件运行产出有效 DOCX |
| ✅ OOXML 生成器 | `output1/ooxml_generator.py` (55KB) | 段落/标题/表格/公式/题注 |
| ✅ OMML 公式引擎 | `output1/omml_generator.py` (27KB) | LaTeX→OMML，支持根号/积分/下标/分数 |
| ✅ LaTeX 解析器 | `output1/latex_parser.py` (27KB) | .tex → thesis_content.json |
| ✅ 内容数据 | `output1/thesis_content.json` (98KB) | 4章, 13图, 9表, 13公式, 34参考文献 |

### 格式修复里程碑
| 轮次 | 关键修复 |
|------|---------|
| R2 | 引用 key→[N] 映射 + TOC 动态域代码 |
| R3 | `~` 残留、标题 numPr、公式下标修复 |
| R10 | OMML 全面修复 + LaTeX 数学保护(save→clean→restore) |
| R15 | `\sqrt` 根号 OMML 渲染 |
| R17 | 页眉阿拉伯 PAGE 域 + 标题等级 parser bug |
| R19 | Gemini 中文间换行零空格 |
| R20 | 标题去粗 w:b w:val="0"+bCs |
| R24 | 表样式管线 BUAA_Thesis_Table (tblStylePr) |
| R25 | 表格安全重构 (ET.fromstring→_make_element) |
| R29 | 表格表头→宋体、表题→黑体、表格居中 |
| R31 | 空白页根治 |
| R32 | LaTeX内容同步(16图/16公式/37参考文献) |
| R33 | 页眉横线正确修复(保留text box线、删镜像段落线) |
| R34 | 页码修复(移除footer5 PAGE域) + header=720(后回退) |
| R35 | header回退至模板标准567(以模板为准) |
| R36 | 页码起始值修复(封面sectPr跳过) |
| R37 | 正文start=1 + 参考文献分页 + 学号确认 |
| R38 | 参考文献编号格式(回退手动[N])  |
| R39 | 表头去粗 + numId冲突修复 |
| R40 | 参考文献手动编号根除冲突 + LaTeX残留清理 |
| R41 | 表内小数点恢复 + 参考文献章题H1 + 公式粗体(omml) |
| R42 | 公式渲染修复(\mathbf/β/×) |
| R43 | 标题去粗(style2继承b) |
| R44 | 章节间距区分(正文/后置) + 分类号JSON提取 |
| R45 | 章节间距统一template style2 |
| R46 | 分类号注入封面(+回退) |
| R47 | 分类号OOXML run修复(+回退) |
| R48 | 分类号TN953→V224替换 |
| R49 | 学号对齐(jc=right移除) + logo位置确认 |
| R58 | 回退至R49稳定基线(header8全页,rId16) |

### R50-R57 实验记录(已回退)
| 轮次 | 尝试 | 结果 |
|------|------|------|
| R50 | 页眉行距 line=240 | 用户要求1.5x,回退 |
| R51 | 裁剪标记 ftr=964 | 部分修复 |
| R52 | header_body.xml (cy=0.88) | Word修复删除 |
| R53 | header14.xml | 页眉丢失 |
| R54 | spAutoFit移除+VML修改 | 0.88cm被Word覆盖 |
| R55 | rId17全页+VML 24.95pt | 无法读取内容 |
| R56 | header14 PAGE域注入 | 无法读取内容,页码失效 |
| R57 | rId17全页+headerRef注入 | 前置页码失效 |

### 格式标准（当前）
| 元素 | 字体 | 字号 | 样式 |
|------|------|------|------|
| H1 章标题 | 黑体 | 三号(sz=32) | 不加粗, 居中 |
| H2/H3/H4 | 黑体 | 小四(sz=24) | 不加粗 |
| 正文 | 宋体 | 小四(sz=24) | 首行缩进 480 |
| 图题 | 宋体 | 五号(sz=21) | 不加粗, 居中 |
| 表题 | 黑体 | 五号(sz=21) | 加粗, 居中 |
| 表头 | 宋体 | 五号(sz=21) | 不加粗 |
| 表体 | 宋体 | 五号(sz=21) | — |
| 页眉 | 校名黑体四号(sz=28), 页码宋体五号(sz=21) | | |
| 引用 | [N] 上标格式 | | |
| 页码 | 前端 upperRoman, 正文 decimal start=1 | | |

---

## 二、当前已知问题（R58 遗留）

| 优先级 | 问题 | 状态 |
|--------|------|------|
| 🟡 | 正文页眉文本框 0.99cm (模板为~0.88cm) | 待修(需处理VML+base64) |
| 🟡 | 参考文献章节格式向模板看齐 | 待修 |
| 🟢 | 公式粗体、表格OMML | P2 |

### VML/DrawingML 双重表示教训
模板页眉文本框有 VML (`<v:shape style="height:28pt">`) + DrawingML (`<a:ext cy="355600">`) 两套表示。VML 的 `o:gfxdata` 含 base64 ZIP 内嵌原始尺寸。仅改 DrawingML 会被 VML 覆盖；同时改会触发"无法读取内容"(base64 校验失败)。**安全方案**：不修改模板 header 的尺寸属性。

---

## 三、最重要的技术教训

### 1. 样式继承陷阱
模板 pStyle 自带粗体/字体定义，会**自动继承**到段落。仅"不写 `<w:b/>`"不够——必须 `<w:b w:val="0"/>` 显式关闭。

### 2. ET.fromstring 命名空间黑洞
`xml.etree.ElementTree` 解析裸 `<w:r>` 无 xmlns → 静默吞错，表格内容全丢。安全方案：用 `_make_element`/`_make_sub_element` helper。

### 3. 中文间换行
Gemini 正则 `([\u4e00-\u9fa5])\s*\n\s*([\u4e00-\u9fa5]) → \1\2` 必须在 `clean_inline_latex` 全局 `\n→空格` **之前**执行。

### 4. 页眉横线来源
北航模板页眉使用 **WPS text box** (`wps:txbx`) 承载校名和页码。横线来自两个来源：
- Header 段落 `pBdr bottom=single`（**应保留**的正确横线）
- Text box 段落也可能有 `pBdr`（需移除的冗余横线）
R31 方向反了——保留了 text box 线，删了 header 段线。

### 5. 空白页根因
两个 section break 相邻 → 空页。解决：将 body sectPr **合并进第一段 pPr**，而非单独成 `<w:p>`。

### 6. Agent 派遣原则
- team mode 在 Windows 因 EPERM fsync 不可用，用 `task(category="deep", run_in_background=true)` 替代
- Agent 结果会因 session 清理丢失——用 `background_output` 需在收到通知后立即采集
- 复杂修复（涉及多文件联动）**亲自做**，简单/独立修复 delegate

---

## 四、关键文件架构

```
E:\thesis\毕设开题\
├── HANDOFF_遗嘱_必读.md          ← 上一份遗嘱（2026-05-14，已过时）
├── .handoff.md                    ← 本文件（交接上下文）
├── 北航本科生论文模板-教务部2026年发.docx ← 原始模板(勿修改!)
│
├── output1/                       ← 主力工作目录
│   ├── gen_round1.py              ← ★ 主转换脚本
│   ├── ooxml_generator.py         ← OOXML 片段生成
│   ├── omml_generator.py          ← OMML 公式引擎
│   ├── latex_parser.py            ← LaTeX→JSON 解析
│   ├── thesis_content.json        ← 结构化内容
│   ├── thesis_r31.docx            ← 最新输出
│   ├── thesis_r{29,30}.docx       ← 近两版备份
│   ├── flatopc_to_docx.py         ← Flat OPC 打包(已不推荐)
│   └── split_xml_for_review.py    ← 大 XML 分块
│
├── tools_archive/                 ← 归档工具
│   ├── shared/                    ← 核心脚本副本
│   ├── v3_round/                  ← 当前轮工具备份
│   └── docs/                      ← 经验文档
│
├── 反馈单_R{N}.md                 ← R2-R31 各轮反馈
├── 版本更新记录.md                ← C-R1 到 C-R31
└── latex模板/                     ← LaTeX 源文件
```

---

## 五、接手步骤

```
1. 读本文件 + R31 反馈单（了解最新问题）
2. 读 版本更新记录.md（了解每轮修了什么）
3. 读 tools_archive/docs/经验记录_R11_R17.md（技术细节）
4. 运行 gen_round1.py → 确认能产出 DOCX
5. 从 页眉横线方向修正 开始（R31 遗留的最高优先级问题）
6. 每次修改后：lxml 验证 + Word 打开测试
```

---

## 六、绝对不要做的事

1. ❌ 不要修改原始模板 `北航本科生论文模板-教务部2026年发.docx`
2. ❌ 不要用 ET.fromstring 解析裸命名空间 XML——用 `_make_element` helper
3. ❌ 不要在多次 agent 派遣后不亲自验证——agent 可能未执行或执行方向反
4. ❌ 不要跳过 Word 实际打开测试
5. ❌ 不要把 `w:header` 统一设为 567 以下——会导致表格与页眉干涉

---

*遗嘱由 Sisyphus 于 2026-05-16 撰写。31 轮迭代，数百次 agent 派遣。*
