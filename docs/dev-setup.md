# TransVBA 开发环境搭建

## 前置条件

- Python 3.11+
- Windows（Word COM 功能需要 Microsoft Word）
- 可选：Microsoft Word（用于运行 `word_com` 标记的测试和 COM 列表解析）

## 安装

```powershell
# 克隆后，进入项目目录
cd TransVBA_git

# 安装项目及开发依赖
python -m pip install -e ".[dev]"
```

## 运行测试

```powershell
# 运行所有非 Word COM 测试（无需安装 Word）
python -m pytest -m "not word_com"

# 运行 Word COM 测试（需要安装 Microsoft Word）
python -m pytest -m word_com

# 运行全部测试
python -m pytest
```

## 运行应用

```powershell
python transvba.py
```

## 依赖说明

| 依赖 | 用途 |
|------|------|
| `python-docx` | docx 文件读写 |
| `lxml` | OOXML 直接操作 |
| `pywin32` | Word COM 自动化（可选，但推荐） |
| `pytest` | 测试框架（仅开发） |

## 常见问题

### 测试被跳过

如果看到 `skipped` 的 `word_com` 测试，这是正常的——说明当前环境没有安装 Word 或 COM 不可用。

### ModuleNotFoundError: No module named 'pytest'

运行 `python -m pip install -e ".[dev]"` 安装开发依赖。
