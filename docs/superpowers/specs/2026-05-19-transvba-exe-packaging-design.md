# TransVBA .exe 打包设计

## 目标

将 TransVBA 项目打包为单个 .exe 文件，用户双击即可运行，保留全部功能不变。

## 方案选择

| 项目 | 选择 | 理由 |
|---|---|---|
| 打包工具 | PyInstaller | 已在 dev 依赖中，成熟稳定 |
| 输出形式 | `--onefile` 单文件 | 用户要求单个 .exe |
| 窗口模式 | `--windowed` / `--noconsole` | 纯 GUI，无命令行黑窗 |
| 入口脚本 | `tvba.py` | 项目 main 入口 |

## 关键配置

### 隐藏导入

- `tkinter` — PyInstaller 不自动检测 tkinter 的隐藏导入，需显式声明
- `lxml` — 可能被遗漏

### 数据文件

- `templates/dapeng_internal.json` → 运行时通过 `sys._MEIPASS` 定位
- `templates/general_spec.json` → 同上
- `tvba_templates.py` 中 `TEMPLATES_DIR` 需适配 PyInstaller 的临时解压路径

### pywin32 COM

- 保留 win32com 打包，启动时惰性检测 Word 可用性
- WPS 用户自动降级，不影响使用

## 构建命令

```
pyinstaller --onefile --windowed --name "TransVBA-Pro" tvba.py
```

## 注意事项

- 首次启动比源码运行慢 2-3 秒（解压依赖）
- .exe 体积约 50-80MB（含 Python 运行时 + 所有依赖）
- 错误处理：程序崩溃时用 `messagebox.showerror` 弹窗显示错误
