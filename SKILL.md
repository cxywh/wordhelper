---
name: wordhelper
version:1.0.0
description: >
  根据用户主题自动生成 Word 文档。
  Use when: 用户说"帮我写一份word文档"、"生成word报告"、"写一份关于...的文档"。
  NOT for: 简单的文本回答、Excel表格生成、PPT制作、PDF生成。
author:cxywh
author_url:https://github.com/cxywh
license:MIT
---



# Wordhelper

帮助用户生成 Word 文档的技能。

## Workflow

### Step 1:解析用户需求

- 从用户输入中提取:

- 文档主题(必填)

- 文档类型(调研报告/技术文档/会议纪要/方案策划等)

- 目标受众(领导/同事/客户/公开)

- 字数要求(如有)

- 语言风格(正式/中性/轻松)

如果没有明确,使用默认值:调研报告、中性风格、1500字左右。

### Step 2: 网络搜索获取资料

#### 首选方法：使用 `web_search` 工具
首先尝试使用 `web_search` 工具搜索主题相关的高质量信息：
- 搜索 5-10 个相关关键词变体
- 优先收录：权威来源、最新数据、典型案例
- 搜索结果保存在临时变量中

搜索关键词建议：
- 核心主题 + "最新进展"
- 核心主题 + "趋势分析"
- 核心主题 + "案例研究"

#### 备选方法：`cn-web-search` 技能（当 `web_search` 失败时）
如果 `web_search` 工具调用失败（例如 API 限制、网络问题），则使用 `cn-web-search` 技能进行网络搜索：

1. **多引擎并行搜索**：使用 `web_fetch` 工具调用多个中文搜索引擎，实现多源交叉验证
2. **推荐引擎组合**（根据 cn-web-search 技能建议）：
   - 百度（综合覆盖最全）
   - 360搜索（结果丰富）
   - 搜狗（中文内容强）
3. **执行示例**：
   ```
   # 中文通用搜索 - 多源并行
   web_fetch(url="https://www.baidu.com/s?wd={主题关键词}", extractMode="text", maxChars=8000)
   web_fetch(url="https://m.so.com/s?q={主题关键词}", extractMode="text", maxChars=8000)
   web_fetch(url="https://www.sogou.com/web?query={主题关键词}", extractMode="text", maxChars=8000)
   ```

#### 搜索策略优化
- **关键词扩展**：除了核心主题，尝试添加以下后缀：
  - "最新进展"、"趋势分析"、"案例研究"、"市场报告"、"行业分析"
- **多源交叉验证**：比较不同搜索引擎的结果，提取一致的关键信息
- **结果整合**：将各引擎的搜索结果合并，去重后保存到临时变量中

#### 错误处理机制
- **web_search 失败检测**：如果 `web_search` 返回错误或空结果，自动切换到 cn-web-search
- **fallback 流程**：
  1. 记录 web_search 失败原因（如有）
  2. 调用 cn-web-search 的推荐引擎组合
  3. 如果中文搜索失败，可尝试英文搜索引擎（如 Brave Search、DuckDuckGo Lite）
  4. 如果所有网络搜索都失败，提示用户并询问是否基于已有知识生成文档

### Step 3: 整理内容结构
根据文档类型生成大纲:

**调研报告结构:**
1. 引言(背景、目的、方法)
2. 核心发现(数据、趋势、关键洞察)
3. 案例分析(2-3个典型案例)
4. 挑战与机遇
5. 结论与建议

**技术文档结构:**
1. 概述
2. 环境准备/前置条件
3. 详细步骤
4. 常见问题
5. 参考资料

**会议纪要结构:**
1. 会议基本信息(时间、地点、参会人)
2. 讨论要点
3. 决议事项
4. 待办事项(责任人+截止时间)

### Step 4: 生成 Markdown 内容
按照选定的大纲填充内容:
- 用搜索到的信息支撑每个章节
- 保持逻辑连贯,避免堆砌
- 添加必要的图表描述(用文字描述,如"如下表所示...")
- 控制总字数在要求范围内

### Step 5: 转换为 Word 文档
使用包装脚本执行转换（自动处理虚拟环境）:

```bash
./scripts/doc.sh --title "文档标题" --input content.md --output "文档名称.docx"
```

或直接使用虚拟环境中的 Python:
```bash
/home/zhangwei/.openclaw/workspace/venv-wordhelper/bin/python scripts/doc.py --title "文档标题" --input content.md --output "文档名称.docx"
```

**注意**: 确保虚拟环境 `venv-wordhelper` 已创建且安装了 `python-docx` 库。

### Step 6: 保存文件并返回
文件命名规则:{主题}-{日期}.docx(例如:AI发展趋势-20260401.docx)

保存路径:当前工作目录或用户指定的目录

返回消息:文档已生成,文件路径为 xxx

## Prerequisites

### 1. 创建并配置虚拟环境
```bash
# 在工作空间创建虚拟环境
cd /home/zhangwei/.openclaw/workspace
python3 -m venv venv-wordhelper

# 激活虚拟环境并安装依赖
source venv-wordhelper/bin/activate
pip install python-docx
```

### 2. 验证安装
```bash
# 验证 python-docx 库是否可用
/home/zhangwei/.openclaw/workspace/venv-wordhelper/bin/python -c "import docx; print('python-docx可用')"
```

**注意**: 如果创建虚拟环境失败，请先安装 `python3-venv` 包:
```bash
sudo apt install python3.12-venv
```

## Output Format
生成成功后返回:
Word 文档已生成!
文件:{主题}-{日期}.docx
路径:{完整路径}
字数:{字数}
如果网络搜索失败:首先尝试使用 cn-web-search 技能作为备选方案(详见 Step 2);如果所有搜索方法都失败,提示用户检查网络,或询问是否基于已有知识生成
如果 Python 环境缺失:提示安装 python3 和必要的库(特别是 `python-docx`)
如果生成失败:保留 Markdown 中间文件,告知用户可以用 Word 打开并另存为 .docx

## 相关技能
- **cn-web-search**：中文网络搜索，获取实时资料

安装相关技能：`clawhub install <slug>`



