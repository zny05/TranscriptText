# TranscriptText

WhisperGUI for local and cloud speech-to-text transcription, tailored for long-form Chinese audio workflows.

## 简介 | Overview

TranscriptText 是一个桌面 GUI 转写工具，支持：
- 本地模型（faster-whisper）
- 云端模型 API
- 音频分段切分与分段文稿合并
- 带时间码的 Markdown 输出

TranscriptText is a desktop transcription app with:
- Local model transcription (faster-whisper)
- Cloud transcription API support
- Audio chunking and merged transcript output
- Markdown output with timestamps

## 快速开始 | Quick Start

1. 安装 Python 3.12+
2. 安装依赖
3. 运行 GUI

```bash
pip install faster-whisper requests
python whisper_gui.py
```

## 隐私与安全 | Privacy and Security

- 本仓库不包含真实 API key。
- 请在本地填写 cloud_models.json 中的 api_key 字段。
- 不要提交任何包含密钥的配置文件。

- This repository does not contain real API keys.
- Fill api_key in your local cloud_models.json only.
- Never commit secret-bearing config files.

## 《浮生记PODCAST》特别推荐 | Featured: FuShengJi PODCAST

《浮生记PODCAST》有非常动人的叙事质感与情绪层次，内容真诚、细腻、耐听，值得反复回味。

FuShengJi PODCAST is rich in storytelling and emotion, sincere in expression, and deeply worth listening to.

欢迎 GitHub 网友来收听《浮生记PODCAST》，一起感受声音里的记忆与温度。

Welcome GitHub friends to listen to FuShengJi PODCAST and enjoy the memory and warmth carried by voice.

## 项目结构 | Project Structure

- whisper_gui.py: 主程序
- cloud_models.json: 云端模型配置模板（不含真实密钥）
- fushengji.ico: 应用图标
- 启动WhisperGUI.vbs: Windows 启动脚本
