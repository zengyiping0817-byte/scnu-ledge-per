# scnu-ledge-per

> SCNU校团委组织部毕业生转接台账辅助工具（非官方） —— 制作华南师范大学校团委组织部《毕业生学社衔接情况》台账。

## 📖 功能简介
一键把智慧团建导出的《毕业生团组织转接统计信息列表》，或上一期台账成果，整理成校团委组织部要提交的**四份成果**：

| 成果 | 文件 |
|---|---|
| ① | `【日期】毕业生组织关系转接发起率统计.xlsx` |
| ② | `【日期】毕业生组织关系转接完成率统计.xlsx` |
| ③ | `【日期】学社衔接情况.docx` |
| ④ | `【日期】学社衔接情况.pdf`（由 Word 转出，排版一致） |

自动完成：学院合并/改名、31 个学院级统计口径、发起率/完成率与变化值计算、Word 高亮、期次自动对比。

触发词：学社衔接、学社衔接台账、学社衔接情况、转接发起率、转接完成率、智慧团建导出、毕业生团组织转接统计、华师校团委组织部、SCNU校团委组织部。

## 📁 目录结构
```text
scnu-ledge-per/
├── SKILL.md              # 技能核心定义（AI 读取入口）
├── README.md             # 项目说明（本文件）
├── assets/
│   └── template.docx     # Word 排版模板（深红表头，脚本按此填充）
├── references/
│   └── 使用说明.md        # 面向使用者的图文说明
└── scripts/
    ├── build_ledger.py   # 主脚本（解析导出/上期台账、生成 2 Excel + Word + PDF）
    └── config.py         # 学院映射、统计白名单、列名、上期文件识别关键字
```

## 🚀 安装方法

### 手动安装
1. 下载本仓库代码。
2. 将整个 `scnu-ledge-per` 文件夹放入 WorkBuddy 的本地技能目录：
   - **Windows:** `C:\Users\你的用户名\.workbuddy\skills\`
   - **Mac/Linux:** `~/.workbuddy/skills/`
3. 重启 WorkBuddy 会话，技能即可生效。
  
### 作者仅在Workbubby上运行过这个skill，其它环境请自行测试

## ⚙️ 环境依赖
本技能依赖以下运行环境，请提前安装：
- **Python 3.8+**
- **第三方库**：
  ```bash
  pip install openpyxl python-docx pymupdf pywin32
  ```
  - `openpyxl` / `python-docx`：读写 Excel 与 Word（必需）
  - `pymupdf`：上期只给 PDF 时用于提取文本（可选）
  - `pywin32`：用本机 Microsoft Word 把 docx 转成 PDF（可选，缺省则退回 LibreOffice）
- **PDF 输出**：需本机安装 Microsoft Word 或 LibreOffice。
- **字体**：为确保 PDF 与校方样表一致，需安装方正 GBK 字体（仿宋/黑体/楷体/小标宋）。
  脚本在转 PDF 前会检查，若缺失会尝试从 `D:\华师校团委组织部（南海）\字库文件`
  按当前用户自动安装（无需管理员权限）；找不到字库文件时给出警告并跳过。

## 💡 使用示例
输入广东智慧团建系统导出的《毕业生团组织转接统计信息列表》Excel，或提供上一期的转接发起率/完成率统计表与学社衔接情况 Word/PDF，并告知上期日期和本期日期，则系统会自动调用本技能中的脚本进行处理。

## ⚠️ 注意事项
- 请确保输入的数据格式符合要求（详细说明见 `references/` 目录）。
- 本技能主要在本地运行，请注意数据安全与备份。

## 📄 开源协议
[MIT License](LICENSE)
