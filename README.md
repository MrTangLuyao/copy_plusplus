# Copy++

A lightweight tool that strips formatting from copied text in real time.

一个实时清除复制内容格式的小工具。

## 功能

- 自动监控剪贴板，将复制的**富文本**（带字体、颜色、超链接等格式）转换为**纯文本**
- **图片不处理**，照常复制粘贴
- **纯文本不处理**，不影响正常使用
- 系统托盘常驻，关闭窗口自动最小化到托盘
- 一键启动 / 停止

## 使用场景

- 从网页、PDF、Word 复制内容到 Markdown / Notion / 邮件，不带乱七八糟的样式
- 写论文引用文献时，避免格式污染
- 整理资料时统一文本样式

## 安装与运行

### 方式一：直接运行 exe（推荐）

从 [Releases](https://github.com/MrTangLuyao/copy_plusplus/releases/) 下载 `Copy++.exe`，双击运行即可。无需安装 Python 或其他依赖。

### 方式二：从源码运行

需要 Python 3.7+：

```bash
pip install PyQt5
python copy_plus_plus_qt.py
```

## 使用说明

1. 启动程序，窗口出现，系统托盘有一个 `C+` 图标
2. 点击「启动」按钮，开始监控剪贴板
3. 正常复制粘贴即可，富文本会自动变成纯文本
4. 关闭窗口 = 最小化到托盘（程序继续运行）
5. 右键托盘图标 →「退出」才真正关闭

## 自行打包 exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "Copy++" copy_plus_plus_qt.py
```

生成的 exe 在 `dist/Copy++.exe`。

## 支持平台

Windows 7 / 10 / 11（x64）

## License

MIT
