# 🚀 NanoTools - All-in-One Utility | 迷你工具箱

A lightweight, high-performance desktop tool for AI background removal, PDF to Word conversion, and Video to GIF processing.  
一款轻量、高效的桌面工具，支持 AI 抠图、PDF 转 Word 以及视频转 GIF。

---

## ⚠️ Important Tips | 使用必读 (非常重要)

### 1. Fix for Random Crashes | 闪退避让指南
Due to a mysterious interaction with Windows File Explorer, **please drag the NanoTools window OUTSIDE of the folder/file preview area** before closing a converted file (e.g., closing a Word doc or Image viewer). Failure to do so may cause the program to crash.  
由于与系统资源管理器的神秘冲突，**在关闭转换后的输出文件（如 Word 或图片查看器）之前，请务必将本程序窗口挪动到该文件界面之外的位置**，否则程序可能会闪退。

### 2. Output Location | 输出位置
Automatic popup is disabled for stability. All converted files are saved in the **SAME DIRECTORY** as your input files.  
为了稳定性，程序不会自动弹出文件夹。所有转换后的文件均保存在与**输入文件相同的目录下**。

### 3. AI Model Loading | AI 模型加载
The AI Background Remover requires loading a model (~44MB) during its first run. It may appear unresponsive for a few seconds; please be patient.  
AI 抠图功能在首次运行时需要加载模型，可能会有数秒的卡顿，请耐心等待。

---

## ✨ Features | 功能特点

* **AI BG Remover**: One-click professional background removal. | **AI 抠图**: 一键去除图片背景。
* **PDF to Word**: Batch-compatible high-fidelity conversion. | **PDF 转 Word**: 高保真还原转换。
* **Video to GIF**: Hard-coded FFmpeg engine for maximum compatibility. | **视频转 GIF**: 内置 FFmpeg 引擎，极速转换。
* **System Monitor**: Real-time CPU/MEM/Network tracking. | **系统监控**: 实时掌控 CPU、内存及网络状态。

---
## 🛠️ Build Requirements | 开发环境

- Python 3.9+
- PyQt6, pyqtgraph, psutil
- pdf2docx, rembg, imageio
- **Required**: Place `ffmpeg.exe` in `/bin` and models in `/models`.

---  
## 🎁 Support the Author | 请作者吃辣条

If this tool helps you, feel free to check the `assets` folder! If you'd like to buy the author a pack of spicy strips (Latiao), your support is much appreciated!  
如果这个工具帮到了你，欢迎查看 `assets` 文件夹。如果你愿意请作者吃包辣条，感激不尽！


