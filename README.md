# wordhelper

根据用户主题自动生成 Word 文档的技能。

**何时使用**：当用户说“帮我写一份word文档”、“生成word报告”、“写一份关于...的文档”时。

## 快速开始

1. **确保虚拟环境已配置**（只需一次）：
   ```bash
   cd /home/username/.openclaw/workspace
   python3 -m venv venv-wordhelper
   source venv-wordhelper/bin/activate
   pip install python-docx
   ```

2. **使用技能**：
   - 告诉助手需要生成什么主题的文档
   - 助手会搜索资料、整理结构、生成内容
   - 最后调用转换脚本生成 `.docx` 文件

## 工作流程

1. **解析需求**：提取文档主题、类型、受众、字数、风格
2. **网络搜索**：使用 `web_search` 或 `cn-web-search` 获取相关资料
3. **整理结构**：根据文档类型（调研报告/技术文档/会议纪要）生成大纲
4. **生成内容**：填充Markdown格式的内容，保持逻辑连贯
5. **转换文档**：调用 `doc.sh` 脚本将 Markdown 转换为 Word
6. **返回结果**：提供生成的 `.docx` 文件路径

## 命令示例

```bash
# 手动调用转换脚本
./scripts/doc.sh --title "文档标题" --input content.md --output "文档名称.docx"

# 或直接使用虚拟环境中的 Python
/home/username/.openclaw/workspace/venv-wordhelper/bin/python scripts/doc.py --title "文档标题" --input content.md --output "文档名称.docx"
```

## 注意事项

- **依赖**：需要 `python-docx` 库，通过虚拟环境安装
- **网络**：搜索失败时会自动切换备用搜索引擎
- **输出**：文件命名格式为 `{主题}-{日期}.docx`
- **错误处理**：生成失败时会保留 Markdown 中间文件

## 相关技能

- **cn-web-search**：中文网络搜索，获取实时资料
- 作者：joansongjr
- 作者的github主页：https://github.com/joansongjr
- cn-web-search技能的仓库：https://github.com/joansongjr/cn-web-search
- 请为cn-web-search技能点个star吧

---

简洁、直接、有效。用户只需说出需求，剩下的交给 wordhelper。