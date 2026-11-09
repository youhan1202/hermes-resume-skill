# 锐仕方达简历报告规范（跨平台 AI Agent 通用）

一套**招聘顾问简历报告（.docx）制作规范**，与具体 AI 平台无关。任何能读写文件或能执行命令的 Agent —— Hermes / Codex / Claude Code / Cursor / Windsurf / Gemini CLI / Coze / Dify / 自建工作流 —— 都能直接用。

## 📦 获取

- 仓库：https://github.com/youhan1202/hermes-resume-skill
- ZIP：https://github.com/youhan1202/hermes-resume-skill/archive/refs/heads/main.zip
- Git：`git clone --depth 1 https://github.com/youhan1202/hermes-resume-skill.git`

## 🚀 按平台的接入方式

### Hermes Agent（一条命令，不需要先加 tap）
```bash
hermes skills install youhan1202/hermes-resume-skill/risfond-resume-report
```
已装过要更新，重跑同一条；若提示已存在加 `--force`。

> ⚠️ 不要用 `hermes skills tap add` 那条老命令：tap 默认只扫仓库的 `skills/` 目录，本仓库技能在根下 `risfond-resume-report/`，加了也装不上。

### Codex / Cursor / Windsurf / Gemini CLI / Jules（读 AGENTS.md 的 Agent）
在本仓库根新建 `AGENTS.md`（内容可直接复制本文件第「核心规范」以下各节，或把 `risfond-resume-report/SKILL.md` 全文放进去），Agent 启动时会自动读取。
- Codex 全局生效：放到 `~/.codex/AGENTS.md`
- Claude Code：复制为项目根 `CLAUDE.md`

### Coze / Dify / 其他工作流平台
1. 下载 ZIP 解压
2. 把 `risfond-resume-report/SKILL.md` 全文放进知识库或系统提示词
3. `references/` 下文档按需上传（读简历方法、样本分析、公司介绍模板）

### 纯聊天窗口（无文件系统）
至少把 `risfond-resume-report/SKILL.md` 全文贴进去。要生成 .docx 需要平台支持代码执行（python-docx）。

### 一句话指令（可直接粘给任意 Agent）
```
请按「锐仕方达猎头简历报告」规范为我生成候选人简历报告。
规范来源：https://github.com/youhan1202/hermes-resume-skill
（zip: https://github.com/youhan1202/hermes-resume-skill/archive/refs/heads/main.zip）

步骤：
1) 下载并解压 zip，通读 risfond-resume-report/SKILL.md 全文，再读 references/candidate-file-reading.md 和 references/eight-sample-analysis.md；
2) 分析我提供的候选人简历，先输出【核心优势】草稿给我确认，不要直接生成；
3) 我确认后，再输出【综合评估】全文预览，我确认后再生成 docx；
4) 排版严格照 SKILL.md（宋体 10.5pt / 行距 1.5 倍 / 工作经历标题行 Bold #2E75B5 / 【】标签不加粗且与内容同段 / 工作年限精确到月）；
5) 生成后用 python-docx 读回逐段打印验证模块完整性，并告诉我本次哪些字段是简历原文、哪些是补写待核实。
```

## 📁 文件说明

```
risfond-resume-report/
├── SKILL.md                        # 核心规范（必读，约 370 行）
├── assets/                         # 页眉 logo、顶部标题图、4 张模块标题图
├── references/
│   ├── candidate-file-reading.md   # 读取候选人简历（PDF/DOC/DOCX）
│   ├── reading-old-doc-files.md    # 旧版 .doc 转换的坑
│   ├── eight-sample-analysis.md    # 8 份真实样本版式特征
│   ├── subagent-prompt-template.md # 多模型协作提示词（可选）
│   ├── it-industry-company-intros.md # IT/算力行业公司介绍模板
│   ├── windows-quirks.md           # Windows 生成脚本注意事项
│   └── github-api-push.md          # 用 API 推送更新（git 直连超时时）
└── scripts/
    ├── gen_resume_report.py        # python-docx 生成脚本模板
    └── docx_reverse_engineer.py    # 反编译 .docx 逐属性对标（格式不对时必跑）
```

## ✨ 功能特性

- 分析候选人简历（PDF/DOC/DOCX），产出符合锐仕方达版式的 .docx 报告
- 核心优势提炼 + 两步人工确认（草稿 → 全文预览）
- 排版严格对标真实样本（尺寸、字体、颜色、页眉 XML 位置）
- 离职原因中性化表述、销售岗/技术岗分类写法
- 支持多模型协作（平台支持子代理时）

## 📋 核心规范速览（细节全在 SKILL.md）

1. **排版结构**：页眉 logo → 顶部标题图 → 标题信息（推荐职位/顾问/日期）→ 基本信息表 → 综合评估 → 教育背景 → 工作经历 → 项目经历 → 尾部声明
2. **字体**：中文宋体、西文 Times New Roman，正文五号 10.5pt，**行距 1.5 倍**，段前段后 0（每个 run 必须显式设颜色）
3. **颜色**：工作经历/项目经历标题行 = Bold + `#2E75B5`；`汇报对象：`/`下属人数：` = Bold + `#000000`；其余标签加粗黑色、内容不加粗
4. **模块标题**：用 `assets/` 里的标题图（8.00 × 0.66 cm）替代文字标题；页眉 logo 6.00 × 0.83 cm 必须在真正的 Word 页眉里
5. **核心优势**：3–4 条，学历（仅统招）→ 年限+行业+核心能力 → 工具/体系 → 语言/管理；0-1 建设经验、产值规模前置
6. **【】标签**：一律不加粗，且与内容同段
7. **薪资**：`XXk*X薪=XXw税前  综合年薪：XXw税前`
8. **工作年限**：精确到月（`17年6个月`），禁止小数和「约」

## 🎯 使用流程

1. 提供候选人简历（PDF/DOC/DOCX）
2. Agent 分析后输出【核心优势】草稿 → 用户确认（常迭代 2–4 轮）
3. 输出【综合评估】全文预览 → 用户确认
4. 生成 .docx，自动打开
5. 交付时附「待核实字段清单」

## 📝 注意事项

- 公司介绍必须 50–100 字，搜索无结果时基于简历原文提炼，不可省略
- 非统招本科不写入核心优势（看教育经历时间线判断）
- 年营收/产值是平台指标，不可跨公司合计；个人业绩（销售额/降本额）可合计
- 教育背景：五号、1.5 倍行距、无段前距、无开头空格；基本信息字段之间不空行
- 已离职写「离职原因」，在职写「看机会原因」，且只写最近一段
- 不添加锐仕方达企业 logo（页眉已有）和装饰性分隔线
- 标题图容易漏：脚本复制到别处后 `assets/` 相对路径失效，必须用绝对路径
- 覆盖已打开的 docx 会 `PermissionError`：先存 .tmp 再 os.replace 重试

## 🔗 相关链接

- GitHub 仓库: https://github.com/youhan1202/hermes-resume-skill
- Hermes Agent 文档: https://hermes-agent.nousresearch.com/docs
