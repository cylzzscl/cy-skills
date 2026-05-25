# cy-skills

> 个人自制的 WorkBuddy / Claude Code 技能（Skill）合集

本仓库收录了 [cy](https://github.com/cylzzscl) 自己设计、编写和维护的 AI 技能，用于增强 WorkBuddy / Claude Code 的领域能力。每个子文件夹是一个独立的 skill，可以直接安装使用。

---

## Skill 清单

| Skill | 版本 | 说明 |
| --- | --- | --- |
| [career-plan-generator](./career-plan-generator%20v2.0-PUB/) | v2.0 (PUB) | 为中国高校大学生生成深度定制化《职业规划成长计划书》——公开版，无特定机构推广内容 |
| [career-plan-generator](./career-plan-generator%20v1.5-GOK/) | v1.5 (GOK) | 同上——国科学员定制版，含国科科技专项摸底与素拓建议 |

### 版本差异说明

- **v2.0-PUB（公开版）**：移除了国科科技相关的条件触发段落和推广内容，适合所有用户通用
- **v1.5-GOK（国科版）**：面向国科学员的定制版，第八章末自动追加「国科科技技术摸底建议」（IT 方向学生），第九章末追加「国科科技素拓活动建议」（活动经历薄弱的学生），封面含国科声明

两个版本的核心工作流、问询流程、方法论嵌入完全一致，差异仅在国科体系相关段落。

---

## 如何安装

### 前提条件

- 已安装 [WorkBuddy](https://www.codebuddy.cn/) 或 [Claude Code](https://docs.anthropic.com/en/docs/claude-code)

### 安装步骤

1. **克隆本仓库**

   ```bash
   git clone https://github.com/cylzzscl/cy-skills.git
   cd cy-skills
   ```

2. **选择你想安装的 skill 子文件夹**，将其复制到你的技能目录下

3. **技能目录位置**（任选其一）：

   - **用户级**（推荐，跨项目通用）：
     ```
     ~/.workbuddy/skills/<skill-name>/
     ```
     即：
     - Windows: `C:\Users\<你的用户名>\.workbuddy\skills\<skill-name>\`
     - macOS/Linux: `~/.workbuddy/skills/<skill-name>/`

   - **项目级**（仅当前项目生效）：
     ```
     <项目根目录>/.workbuddy/skills/<skill-name>/
     ```

4. **确保文件夹名称与 SKILL.md 中的 `name` 字段一致**

   例如安装 `career-plan-generator v2.0-PUB`：
   ```bash
   # 复制到用户级技能目录，并重命名为 SKILL.md 中定义的 name
   cp -r "cy-skills/career-plan-generator v2.0-PUB" ~/.workbuddy/skills/career-plan-generator
   ```

   Windows PowerShell：
   ```powershell
   Copy-Item -Recurse "cy-skills\career-plan-generator v2.0-PUB" "$env:USERPROFILE\.workbuddy\skills\career-plan-generator"
   ```

5. **重启 WorkBuddy / Claude Code**，技能将自动加载

### 安装验证

启动后输入 `/skills` 查看已加载的技能列表，确认 `career-plan-generator` 已出现。

---

## Skill 详细说明

### career-plan-generator

为中国高校大学生（本科 / 大专 / 硕士在读）生成深度定制化的《职业规划成长计划书》。

**核心特性**：
- 🎯 **强制分阶段问询**：36 轮单题交互，学生体验是"被深度咨询"而非"填表"
- 🔍 **四维真实网络调研**：行业 / 岗位 / 城市 / 组织，使用 WebSearch 获取真实信息
- 🧠 **霍兰德 + MBTI 双重匹配**：兴趣代码与性格类型交叉诊断
- 📐 **三套方法论嵌入**：OKR 目标拆解 + 时间管理矩阵 + SCQA 沟通模型
- 💪 **可迁移能力六大类 22 子项**：岗位映射 + 自评矩阵
- 🏖 **寒暑假 OKR 专项行动**：最重要的整段可支配时间单独规划
- ✉ **反重复温度信**：五维随机化机制，每次生成的结尾都不同
- 📄 **跨环境必出 Word**：Node → Python → Pandoc → Markdown 四级 fallback

**触发关键词**：
- 「写一份职业规划书」「帮我做生涯规划」「生成成长计划书」
- 「我是大X学生想做规划」「career plan」「职业生涯规划」
- 「不要那种笼统的、要能落地的规划书」

**目录结构**（以 v2.0-PUB 为例）：
```
career-plan-generator v2.0-PUB/
├── SKILL.md                  主流程定义（AI 读取的核心文件）
├── README.md                 技能说明
├── references/
│   ├── interview.md          5 阶段问询完整问题库（36 轮）
│   ├── research.md           四维调研维度与搜索关键词模板
│   ├── methodologies.md      方法论卡片库（OKR / 时间矩阵 / SCQA / 可迁移能力）
│   └── plan_template.md      正文章节模板 + 自检清单
└── scripts/
    ├── build_docx.js         docx 生成脚本（Node 版）
    └── build_docx.py         docx 生成脚本（Python 版·跨环境兜底）
```

**依赖**：
- 网络搜索工具（WebSearch / WebFetch）——四维调研需要
- Node.js + [docx](https://www.npmjs.com/package/docx) 包 或 Python + [python-docx](https://pypi.org/project/python-docx/)——生成 .docx

---

## 如何开发自己的 Skill

每个 skill 就是一个包含 `SKILL.md` 的文件夹。`SKILL.md` 的 frontmatter 定义了技能名称和触发条件：

```yaml
---
name: my-skill
description: 技能描述，AI 会根据这段文字判断是否自动启用
---

# 技能的详细指令...
```

更多参考资料：
- [WorkBuddy Skill 开发指南](https://www.codebuddy.cn/docs/workbuddy/skills)
- [Claude Code Custom Agents](https://docs.anthropic.com/en/docs/claude-code)

---

## License

MIT
