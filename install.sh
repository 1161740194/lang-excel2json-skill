#!/bin/bash

################################################################################
# lang-excel2json 自动安装脚本 (macOS)
# 版本: 1.0
# 用途: 自动配置 Excel 到 i18n JSON 转换工具
################################################################################

set -e  # 遇到错误立即退出

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 项目信息
PROJECT_NAME="lang-excel2json"
SCRIPT_SOURCE="/Users/hanyue/.claude/skills/lang-excel2json/scripts/excel_to_i18n_json.py"
CURRENT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$CURRENT_DIR/venv"
LOCAL_SCRIPT="$CURRENT_DIR/excel_to_i18n_json.py"

# 打印带颜色的消息
print_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

print_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

print_header() {
    echo ""
    echo -e "${GREEN}========================================${NC}"
    echo -e "${GREEN}  $1${NC}"
    echo -e "${GREEN}========================================${NC}"
    echo ""
}

# 检查命令是否存在
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# 开始安装
print_header "lang-excel2json 自动安装程序"

print_info "项目目录: $CURRENT_DIR"
echo ""

# 步骤 1: 检查 Python
print_header "步骤 1/6: 检查 Python 环境"

if ! command_exists python3; then
    print_error "未找到 Python 3"
    echo ""
    echo "请先安装 Python 3:"
    echo "  brew install python@3.14"
    echo ""
    echo "或访问: https://www.python.org/downloads/macos/"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
PYTHON_MAJOR=$(echo $PYTHON_VERSION | cut -d. -f1)
PYTHON_MINOR=$(echo $PYTHON_VERSION | cut -d. -f2)

print_success "找到 Python $PYTHON_VERSION"

if [ "$PYTHON_MAJOR" -lt 3 ] || ([ "$PYTHON_MAJOR" -eq 3 ] && [ "$PYTHON_MINOR" -lt 8 ]); then
    print_error "Python 版本过低 (需要 3.8+)"
    exit 1
fi

print_success "Python 版本检查通过 ✓"
echo ""

# 步骤 2: 检查源脚本
print_header "步骤 2/6: 检查源脚本文件"

if [ ! -f "$SCRIPT_SOURCE" ]; then
    print_error "源脚本文件不存在: $SCRIPT_SOURCE"
    echo ""
    echo "请确保 lang-excel2json skill 已正确安装"
    exit 1
fi

print_success "源脚本文件存在 ✓"
echo ""

# 步骤 3: 创建虚拟环境
print_header "步骤 3/6: 创建虚拟环境"

if [ -d "$VENV_DIR" ]; then
    print_warning "虚拟环境已存在，跳过创建"
else
    print_info "创建虚拟环境: $VENV_DIR"
    python3 -m venv "$VENV_DIR"
    print_success "虚拟环境创建成功 ✓"
fi

# 验证虚拟环境
if [ -f "$VENV_DIR/bin/python3" ]; then
    VENV_PYTHON=$("$VENV_DIR/bin/python3" --version 2>&1 | awk '{print $2}')
    print_success "虚拟环境 Python 版本: $VENV_PYTHON ✓"
else
    print_error "虚拟环境创建失败"
    exit 1
fi
echo ""

# 步骤 4: 复制脚本到本地
print_header "步骤 4/6: 复制脚本到本地"

if [ -f "$LOCAL_SCRIPT" ]; then
    print_warning "本地脚本已存在，备份为 excel_to_i18n_json.py.bak"
    cp "$LOCAL_SCRIPT" "$LOCAL_SCRIPT.bak"
fi

print_info "复制脚本文件..."
cp "$SCRIPT_SOURCE" "$LOCAL_SCRIPT"
chmod +x "$LOCAL_SCRIPT"

if [ -f "$LOCAL_SCRIPT" ]; then
    print_success "脚本复制成功 ✓"
    print_info "脚本位置: $LOCAL_SCRIPT"
else
    print_error "脚本复制失败"
    exit 1
fi
echo ""

# 步骤 5: 创建便捷命令脚本
print_header "步骤 5/6: 创建便捷命令"

WRAPPER_SCRIPT="$CURRENT_DIR/excel2json"

cat > "$WRAPPER_SCRIPT" << 'WRAPPER_EOF'
#!/bin/bash
# excel2json 便捷命令包装脚本

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_PYTHON="$SCRIPT_DIR/venv/bin/python3"
CONVERTER_SCRIPT="$SCRIPT_DIR/excel_to_i18n_json.py"

if [ ! -f "$VENV_PYTHON" ]; then
    echo "错误: 虚拟环境未找到"
    echo "请先运行: bash install.sh"
    exit 1
fi

if [ ! -f "$CONVERTER_SCRIPT" ]; then
    echo "错误: 转换脚本未找到"
    echo "请先运行: bash install.sh"
    exit 1
fi

# 使用虚拟环境的 Python 运行转换脚本
exec "$VENV_PYTHON" "$CONVERTER_SCRIPT" "$@"
WRAPPER_EOF

chmod +x "$WRAPPER_SCRIPT"

if [ -f "$WRAPPER_SCRIPT" ]; then
    print_success "便捷命令创建成功 ✓"
    print_info "可以使用命令: ./excel2json"
else
    print_error "便捷命令创建失败"
    exit 1
fi
echo ""

# 步骤 6: 创建 shell 别名配置
print_header "步骤 6/6: 配置 shell 别名（可选）"

SHELL_RC=""
if [ -n "$ZSH_VERSION" ]; then
    SHELL_RC="$HOME/.zshrc"
elif [ -n "$BASH_VERSION" ]; then
    SHELL_RC="$HOME/.bash_profile"
fi

ALIAS_LINE="alias excel2json='$WRAPPER_SCRIPT'"

if [ -n "$SHELL_RC" ] && [ -f "$SHELL_RC" ]; then
    if grep -q "alias excel2json=" "$SHELL_RC" 2>/dev/null; then
        print_warning "shell 别名已存在，跳过添加"
    else
        echo ""
        echo "是否要添加 'excel2json' 别名到 $SHELL_RC？"
        echo "这样可以在任何目录直接使用 excel2json 命令"
        echo ""
        read -p "添加别名? [y/N]: " -n 1 -r
        echo ""

        if [[ $REPLY =~ ^[Yy]$ ]]; then
            echo "" >> "$SHELL_RC"
            echo "# lang-excel2json 别名" >> "$SHELL_RC"
            echo "$ALIAS_LINE" >> "$SHELL_RC"
            print_success "别名已添加到 $SHELL_RC ✓"
            print_info "请运行以下命令使其生效:"
            echo ""
            echo "  source $SHELL_RC"
            echo ""
        else
            print_info "跳过添加别名"
            print_info "您仍然可以使用: $WRAPPER_SCRIPT"
        fi
    fi
else
    print_warning "未检测到 shell 配置文件，跳过别名配置"
fi
echo ""

# 测试安装
print_header "测试安装"

print_info "测试转换脚本..."

TEST_OUTPUT=$("$WRAPPER_SCRIPT" --help 2>&1)
TEST_EXIT_CODE=$?

if [ $TEST_EXIT_CODE -eq 0 ]; then
    print_success "转换脚本测试成功 ✓"
else
    print_error "转换脚本测试失败"
    echo "$TEST_OUTPUT"
    exit 1
fi
echo ""

# 安装完成
print_header "安装完成！"

echo -e "${GREEN}✓ Python 环境检查完成${NC}"
echo -e "${GREEN}✓ 虚拟环境创建完成${NC}"
echo -e "${GREEN}✓ 转换脚本复制完成${NC}"
echo -e "${GREEN}✓ 便捷命令创建完成${NC}"
echo ""

print_info "快速开始指南:"
echo ""
echo "1. 直接使用便捷命令:"
echo -e "   ${BLUE}./excel2json language.xlsx output.json --sheet buff --start 380 --end 415${NC}"
echo ""
echo "2. 或激活虚拟环境后使用:"
echo -e "   ${BLUE}source venv/bin/activate${NC}"
echo -e "   ${BLUE}python3 excel_to_i18n_json.py language.xlsx output.json${NC}"
echo -e "   ${BLUE}deactivate${NC}"
echo ""
echo "3. 查看帮助信息:"
echo -e "   ${BLUE}./excel2json --help${NC}"
echo ""
echo "4. 查看详细文档:"
echo -e "   ${BLUE}cat lang-excel2json使用文档.md${NC}"
echo ""

print_success "所有配置已完成！开始使用吧 🎉"
echo ""
