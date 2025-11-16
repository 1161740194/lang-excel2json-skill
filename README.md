# Excel to i18n JSON Converter

一键将 Excel 多语言文件转换为国际化 JSON 格式的工具

---

## 🚀 快速开始

### 操作流程

```bash
mkdir ~/.claude/skills/lang-excel2json

cp -r /Users/admin/Desktop/code/lang-excel2json/skill/* ~/.claude/skills/lang-excel2json

```

### 一键安装

```bash
bash install.sh
```

安装脚本会自动完成以下配置：
- ✅ 检查 Python 环境
- ✅ 创建虚拟环境
- ✅ 复制转换脚本
- ✅ 创建便捷命令
- ✅ 配置 shell 别名（可选）

### 立即使用

```bash
# 转换 Excel 文件
./excel2json language.xlsx output.json --sheet test --start 380 --end 415
```

---

## Claude Code用法

```bash
  # 创建虚拟环境
  python3 -m venv venv

  # 激活虚拟环境
  source venv/bin/activate

  claude

```

```
/lang-excel2json

```

---

## 💡 基本用法

### 方法一：使用便捷命令（推荐）

```bash
# 转换整个工作表
./excel2json language.xlsx output.json --sheet translations

# 转换指定行范围
./excel2json language.xlsx output.json --sheet test --start 100 --end 500

# 查看帮助
./excel2json --help
```

### 方法二：使用虚拟环境

```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate

# 使用转换脚本
python3 excel_to_i18n_json.py language.xlsx output.json

# 退出虚拟环境
deactivate
```

---

## 📁 输出格式

转换后生成标准 i18n JSON 格式：

```json
{
  "en": {
    "greeting": "Hello",
    "farewell": "Goodbye"
  },
  "zh-CN": {
    "greeting": "你好",
    "farewell": "再见"
  },
  "zh-HK": {
    "greeting": "你好",
    "farewell": "再見"
  }
}
```

可直接用于：
- React + react-i18next
- Vue 3 + vue-i18n
- Next.js + next-i18next

---

## 🌍 支持的语言

自动检测并转换 25+ 种语言：

英语 (en) | 简体中文 (zh-CN) | 繁体中文 (zh-HK/zh-TW) | 日语 (ja) | 韩语 (ko) | 阿拉伯语 (ar) | 德语 (de) | 法语 (fr) | 西班牙语 (es) | 意大利语 (it) | 葡萄牙语 (pt) | 俄语 (ru) | 印地语 (hi) | 孟加拉语 (bn) | 越南语 (vi) | 泰语 (th) | 印尼语 (id) | 马来语 (ms) | 菲律宾语 (fil) | 波兰语 (pl) | 荷兰语 (nl) | 土耳其语 (tr) | 波斯语 (fa) | 乌尔都语 (ur)

---

## 🔧 系统要求

- **操作系统**: macOS 10.15+ (Catalina 或更高)
- **Python**: 3.8+ (推荐 3.10+)
- **芯片**: 支持 Apple Silicon (M1/M2/M3) 和 Intel
- **依赖**: 无（仅使用 Python 标准库）

---

## 📞 获取帮助

- 查看命令帮助: `./excel2json --help`
- 重新安装: `bash install.sh`

---

## 🎯 常用命令

```bash
# 查看 Excel 工作表列表
python3 << 'EOF'
import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile('language.xlsx', 'r') as zf:
    with zf.open('xl/workbook.xml') as f:
        tree = ET.parse(f)
        root = tree.getroot()
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        sheets = root.findall('.//main:sheet', ns)
        print("\n可用的工作表:\n")
        for i, sheet in enumerate(sheets, 1):
            print(f"  {i}. {sheet.get('name')}")
        print()
EOF

# 验证 JSON 格式
python3 -m json.tool output.json > /dev/null && echo "✓ JSON 格式正确"

# 统计翻译数量
python3 << 'EOF'
import json
with open('output.json', 'r', encoding='utf-8') as f:
    data = json.load(f)
for lang, texts in sorted(data.items()):
    print(f"{lang:8s}: {len(texts):4d} 条翻译")
EOF
```

---

**版本**: 1.0
**更新日期**: 2025-11-16
**兼容性**: macOS 10.15+ | Python 3.8-3.14+ | Apple Silicon & Intel
