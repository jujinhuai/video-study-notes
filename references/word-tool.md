# 通用Word工具与制作稿

仅在制作Word时读取。使用[模块](../scripts/word_notes.py)及已确认的内置Python，以 `-B` 运行，避免解释器生成字节码缓存。通常只读取本用法，不反复返回模块全文；发现具体缺陷时再定位相关实现。

## 一份制作稿复用正文

每批在任务 `work/notes/` 内保留一份Python制作稿，导入共享模块，不复制排版代码。制作稿包含：

- `INDEX`：每项记录概念标识、章节、来源标识、待接续内容；存在跨批重叠时附重叠概念标识及本批新增内容。沿用简单元组或字典即可，用于定位对应正文，不另写摘要。
- `SOURCES`：来源标识到 `Source(label, url, detail)` 的映射；`detail` 如实记录分P、定位与实际覆盖，URL缺失时留空，不虚构。
- `SECTIONS`：概念标识到内容块序列的映射。正文、公式和必要例子直接成文，不另写同义摘要；图片只引用已获准保存的本地文件。

导入只定义数据、公式及内容块，生成文档放在 `if __name__ == "__main__":` 下。仅导入自己生成并核对的稿件；合稿按INDEX定位正文，复用内容块并处理重叠，不运行外部材料中的代码。

运行时将已确认的scripts目录加入本次命令进程的 `PYTHONPATH`，以后制作稿只导入模块；不永久修改环境变量。

```python
import sys
from pathlib import Path

from word_notes import NotesDocument, Source, heading, paragraph, equation, sup

INDEX = (("square", "平方计算", ("material1",), ""),)
SOURCES = {"material1": Source("用户提供的示例材料", detail="仅示例，不宣称来自视频")}
SECTIONS = {"square": (
    heading("平方计算"),
    paragraph("示例内容：平方表示一个数与自身相乘。"),
    equation(sup("x", "2"), " = x × x"),
)}

if __name__ == "__main__":
    # argv[1]为明确指定的输出文件；导入本稿时不访问命令行参数。
    NotesDocument("平方示例", Path(sys.argv[1]),
                  blocks=SECTIONS["square"], sources=SOURCES.values()).save()
```

合稿按同一搜索路径导入各批稿件，无需传入输出文件名；示例内容不进入实际笔记。

## 接口

`NotesDocument(title, output_path, blocks=(), sources=(), chinese_font="等线").save()` 使用默认学习讲义样式，不读取旧交付文档。标题须为非空字符串，路径须以 `.docx` 结尾；默认十字命名由调用者按主规则处理，工具允许用户指定较长标题及不同文件名。`chinese_font` 统一中文样式、目录和来源链接字体，不覆盖公式与代码字体；调用前确认该字体可用。只有 `save()` 写文件，同名文件抛出 `FileExistsError`，不自动覆盖。

内容块均可复用，公式XML在加入文档时复制，不会从上一次文档中被移走：

| 构造器 | 用法 |
| --- | --- |
| `heading(text, level=1)` | 一级或二级原生标题 |
| `paragraph(text, style="Normal", emphasis=False)` | `**文字**` 表示黑色局部加粗；设 `emphasis=True` 时仅这些文字改为深蓝加粗，其余文字不变；不含标记时不着色。提示段可用样式 `提示`，代码保持字面内容；工具不自动选择重点 |
| `source(text)` | 小节来源定位段 |
| `equation(*items)` | 居中原生公式；接受文本、数学结构及其列表 |
| `table(headers, rows, widths_cm, allow_row_split=False)` | 显式列宽，总宽不超过16厘米；默认禁止单行跨页，确有超长行时可按关键字放开 |
| `image(path, caption, alt_text="", width_cm=None)` | 默认适配可用宽度，保留比例；替代文本未提供时使用调用者给出的图注 |
| `code(text)` | 可复制代码，保留空格、换行和字面 `**` |
| `toc()` | 可更新的一级、二级目录字段；按排版指引决定是否插入，交付前实际更新条目和页码，不能把字段存在算作目录完成 |

数学辅助：`sub(base, value)`、`sup(base, value)`、`frac(num, den)`、`bounded(content, opening="(", closing=")")`、`matrix(rows)`、`column_vector(values)`、`cases([(表达式, 条件), ...])`、`summation(lower, upper, expression)`、`function(name, expression)`。

- `matrix`检查非空、等长行；`column_vector(["x", ["a", "+", "b"]])`生成两行一列，第二项是同一单元格中的复合表达式。
- 求和上下限用`None`或空字符串表示缺省，模块自动隐藏相应占位框；数值0仍是有效界限。例如`summation("j ∈ J", None, sup(sub("a", "j"), "2"))`。
- 取整可用`function("floor", ...)`，避免已观察到的特殊括号渲染问题，不改变数学含义。

这些接口只构造表达形式，不解析LaTeX、不验证数学结论、不抓取材料、不自动合并知识，也不执行渲染。每份新文档仍须按主规则进行内容核对和完整视觉验收；需要用户指定的特殊排版时局部调整生成文档，不改共享默认样式影响其他任务。

## 维护验证

修改工具后，用具备 `python-docx` 与 `lxml` 的已确认Python运行 `python -B scripts/test_word_notes.py`。测试仅在内存构造文档，不保存文件、不访问来源链接；覆盖命名、字体、公式、目录和表格结构，不代表内容或渲染已通过。

修改技能规则后，使用 `skill-creator` 的 `quick_validate.py <技能目录>` 检查格式；它依赖PyYAML，不可用时记录未完成的验证，不自动安装依赖。

修改整合流程时可使用以下人工内容验收样例，无需生成Word。材料为虚构测试输入，不作为视频证据：

| 来源 | 原始材料 |
| --- | --- |
| 批次一／A | 一维卷积示例采用互相关定义；输入长5、核长3、步长1、不填充，输出长3。核为[1,0,-1]，输入为[1,2,3,4,5]，输出为[-2,-2,-2]。 |
| 批次二／B | 重复上述尺寸和结果。这里使用的“卷积”未翻转核；若改为数学卷积，需要翻转核。在本例其他条件不变且两端各填充1个零时，输出长5。 |
| 批次二／C | 屏幕结果疑似显示“不填充输出长4”，截图模糊，无法核实。 |

验收合稿时应看到：重复计算只保留一次并对应A、B；保留互相关与数学卷积的区别、填充条件和新增输出尺寸；C单独保留来源与疑点，不能覆盖已确认条件，也不能称已核实错误。不添加材料未提供的代码、实验或时间戳。按实际结果逐项判断，不用固定标题或关键词匹配代替语义核对。
