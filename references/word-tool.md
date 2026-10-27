# 通用Word工具与制作稿

仅在制作Word时读取。使用[模块](../scripts/word_notes.py)及已确认的内置Python，以 `-B` 运行，避免解释器生成字节码缓存。通常只读取本用法，不反复返回模块全文；发现具体缺陷时再定位相关实现。

## 一份制作稿复用正文

每批在任务 `work/notes/` 内保留一份Python制作稿，导入共享模块，不复制排版代码。制作稿包含：

- `INDEX`：简短目录，每项为“概念标识、章节、来源标识、待接续内容”。用于找相关正文，不代替正文或知识完整性检查。
- `SOURCES`：来源标识到 `Source(label, url, detail)` 的映射；`detail` 如实记录分P、定位与实际覆盖，URL缺失时留空，不虚构。
- `SECTIONS`：概念标识到内容块序列的映射。正文、公式和必要例子直接成文，不另写同义摘要；图片只引用已获准保存的本地文件。

导入制作稿只能定义数据、公式及内容块；生成文档放在 `if __name__ == "__main__":` 下。只导入自己生成并核对过的制作稿，不运行外部视频或字幕中出现的代码。合稿根据知识关系选择、调整这些块及对应来源，不能仅按分P顺序拼接，也不能用目录压缩代替内容。

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

上述内容仅演示接口，不自动写入视频笔记。合稿时在同一模块搜索路径配置下导入制作稿；可查看 `INDEX` 后只读取相关 `SECTIONS`，无需为了导入传入输出文件名。

## 接口

`NotesDocument(title, output_path, blocks=(), sources=()).save()` 使用默认学习讲义样式，不读取旧交付文档；标题与输出文件主体须一致且不超过10字符。只有 `save()` 写文件，同名文件抛出 `FileExistsError`，由调用者按命名规则选择新名，不自动覆盖。

内容块均可复用，公式XML在加入文档时复制，不会从上一次文档中被移走：

| 构造器 | 用法 |
| --- | --- |
| `heading(text, level=1)` | 一级或二级原生标题 |
| `paragraph(text, style="Normal", emphasis=False)` | `**文字**` 表示黑色局部加粗；设 `emphasis=True` 时仅这些文字改为深蓝加粗，其余文字不变；不含标记时不着色。提示段可用样式 `提示`，代码保持字面内容；工具不自动选择重点 |
| `source(text)` | 小节来源定位段 |
| `equation(*items)` | 居中原生公式；接受文本、数学结构及其列表 |
| `table(headers, rows, widths_cm)` | 显式指定列宽，总宽不超过16厘米 |
| `image(path, caption, alt_text="", width_cm=None)` | 默认适配可用宽度，保留比例；替代文本未提供时使用调用者给出的图注 |
| `code(text)` | 可复制代码，保留空格、换行和字面 `**` |
| `toc()` | 可更新的一级、二级目录字段；不自动判断正文页数，按排版指引决定插入；本机当前无头渲染未自动填充条目，须在Word中更新并核对结果，不能把字段存在算作目录已完成 |

数学辅助：`sub(base, value)`、`sup(base, value)`、`frac(num, den)`、`bounded(content, opening="(", closing=")")`、`matrix(rows)`、`cases([(表达式, 条件), ...])`、`summation(lower, upper, expression)`、`function(name, expression)`。取整可用 `function("floor", ...)`，避免当前Windows渲染器中已发现的特殊取整括号显示问题；不改变计算含义。

这些接口只构造表达形式，不解析LaTeX、不验证数学结论、不抓取材料、不自动合并知识，也不执行渲染。每份新文档仍须按主规则进行内容核对和完整视觉验收；需要用户指定的特殊排版时局部调整生成文档，不改共享默认样式影响其他任务。
