#!/bin/bash
# Word文档生成包装脚本
# 自动使用虚拟环境中的Python执行doc.py

set -e

# 虚拟环境路径
VENV_PATH="/home/zhangwei/.openclaw/workspace/venv-wordhelper"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# 检查虚拟环境是否存在
if [ ! -f "$VENV_PATH/bin/python" ]; then
    echo "错误: 虚拟环境不存在或Python解释器缺失"
    echo "请先创建虚拟环境:"
    echo "  cd /home/zhangwei/.openclaw/workspace"
    echo "  python3 -m venv venv-wordhelper"
    echo "  source venv-wordhelper/bin/activate"
    echo "  pip install python-docx"
    exit 1
fi

# 检查python-docx是否可用
if ! "$VENV_PATH/bin/python" -c "import docx" 2>/dev/null; then
    echo "错误: python-docx库未安装在虚拟环境中"
    echo "请安装: $VENV_PATH/bin/pip install python-docx"
    exit 1
fi

# 执行doc.py脚本
"$VENV_PATH/bin/python" "$SCRIPT_DIR/doc.py" "$@"