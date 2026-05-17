# 当前已知问题（懒得改，大家有兴趣可以调试）

页眉页码可能会换行-->进入页眉，横向拉长文本框即可

摘要第一页不是大写罗马数字-->手动调下

没有做多图排版、附录的适配

表格序号在交叉引用中正文里可能会乱

部分复杂公式难渲染

积分上下会留小框-->手打空格

“本人声明”四个字会丢失

。。。

后面就交给你了蟹不肉

# BUAA Thesis Converter Skill

北航本科毕设 LaTeX → Word 论文转换技能。将 LaTeX 论文源码按北航2026教务部模板规范转换为可直接提交的 .docx 文件。

## 管线

```
LaTeX .tex → latex_parser.py → thesis_content.json → gen_round1.py → thesis.docx
                                  ↑                        ↑
                           omml_generator.py        ooxml_generator.py
```

## 脚本

| 文件 | 用途 |
|------|------|
| `scripts/latex_parser.py` | LaTeX → JSON 结构化解析（含附录） |
| `scripts/gen_round1.py` | 主生成脚本：JSON → OOXML → DOCX |
| `scripts/ooxml_generator.py` | OOXML 段落/表格/图片/公式生成器 |
| `scripts/omml_generator.py` | LaTeX → OMML 公式引擎 |
| `scripts/flatopc_to_docx.py` | Flat OPC XML → DOCX 打包 |

## 最新基线 (R75)

- 附录解析：parse_thesis() 显式解析 data/appendix.tex
- 公式段落边界：_replace_eq_with_ref 用 \n\n 保留段落分隔
- 模板占位清理：自动移除模板自带附录 A-C 占位
- 附录标签 A-F，表 A.1-D.1
