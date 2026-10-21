# 论文格式锁定智能体（Paper Format Lock）

把**论文范稿**投进来 → 自动**锁定格式** → 投入**需要更改的论文** → 输出完全符合范稿格式的新论文。

一条命令完成，全程可核对（111 项格式自动校验）。

---

## 一、这是什么

一个本地命令行工具 + 一个可复用的 WorkBuddy 技能（`paper-format-lock`）：

| 能力 | 说明 |
| --- | --- |
| **锁定格式** | 从范稿 .docx 提取完整格式档案：页面设置（纸张/边距）、默认字体（docDefaults）、全部样式定义、段落角色映射、自动编号、表格边框，存为 `profile.json` |
| **应用格式** | 对待排版论文逐段识别角色（标题/一至三级标题/正文/摘要/关键词/参考文献/题注/表格文字），套用范稿样式，清除混乱的直接格式 |
| **格式核对** | 从输出反提档案与锁定档案逐项比对，111 项全过才算合格 |

## 二、快速开始

环境：Python + python-docx（本机已装在
`C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe`，
若缺依赖先 `...python.exe -m pip install python-docx`）。

```bash
# 一步到位：锁定范稿格式 → 应用到待排版论文 → 输出 + 报告
python paperfmt.py run 课程论文模板.docx 待排版论文.docx \
    -o 输出_已排版论文.docx --report 排版报告.md

# 或分两步（格式档案可复用、可分享）
python paperfmt.py lock 课程论文模板.docx -o profile.json
python paperfmt.py apply profile.json 待排版论文.docx -o 输出.docx --report 报告.md

# 核对输出是否符合锁定格式
python paperfmt.py verify profile.json 输出.docx --report 核对报告.md
```

演示样本在 `samples/` 目录：`课程论文模板.docx`（范稿）、`待排版论文.docx`（格式混乱稿）、
`输出_已排版论文.docx`（结果）、`排版报告.md`、`核对报告.md`。

## 三、锁定了哪些格式

| 类别 | 内容 |
| --- | --- |
| 页面 | 纸张大小、方向、上/下/左/右页边距、页眉页脚距离（应用到所有节） |
| 默认 | docDefaults 默认字体/字号/段落格式 |
| 样式 | 范稿全部样式定义整体复制（含字体、字号、加粗、对齐、行距、缩进、大纲级别、段前段后） |
| 角色 | 标题、一/二/三级标题、正文、摘要标题、摘要正文、关键词、参考文献标题、参考文献条目、图表题注、表格文字 |
| 编号 | 范稿标题样式的自动编号定义（导入目标文档，并剥离手动编号避免重复） |
| 表格 | 表格边框（三线表等）、单元格文字格式 |
| 细节 | “关键词：”标签加粗而后续不加粗的段内混合格式，会拆分 run 还原 |

## 四、角色是怎么识别的

按优先级依次判断：**关键词段 → 摘要标题 → 参考文献标题 → 首个像标题的段落 → 样式名含
heading/标题或大纲级别 → 文本编号（一、／（一）／1.）→ 参考文献区 → 摘要区 → 图表题注 → 正文**。
范稿中靠手动排版（直接格式）的段落会合成为命名样式（如“论文-标题”）存入档案，应用时自动创建。

## 五、使用建议与局限

- **建议**：先用 `run` 跑一遍看报告，再按需微调；格式档案 `profile.json` 可保存复用，
  同一门课程后续论文直接 `apply` 即可。
- 页眉/页脚内容不锁定（各论文字内容不同，如需统一可告知）。
- 旧版 `.doc` 需先另存为 `.docx`。
- 图片本身不做处理，图片段落保留原对齐/缩进。
- 表格单元格的加粗会被保留（保护表头强调）。

## 六、文件说明

```
paper-format-agent/
├── paperfmt.py        # CLI 入口（lock / apply / run / verify）
├── paperfmt_core.py   # 核心库（提取、应用、核对、报告）
├── make_samples.py    # 生成演示样本（范稿 + 混乱稿）
├── samples/           # 演示样本与输出
└── dist/              # 技能分发包 paper-format-lock.zip
```

技能本体已安装到 `C:\Users\Administrator\.workbuddy\skills\paper-format-lock\`：
后续对话中直接说“按这个模板给我的论文排版”即可触发，无需记住命令。
