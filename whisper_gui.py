#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
WhisperGUI — 音频转文字稿工具
基于 faster-whisper large-v3 本地模型
"""

import os
import json
import re
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import wave
from contextlib import closing
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ─── 默认模型路径（可在界面里修改）──────────────────────────────────────────


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


DEFAULT_MODEL_PATH = "model/whisper"
DEFAULT_CLOUD_URL = "https://www.dmxapi.cn/v1/audio/transcriptions"
DEFAULT_CLOUD_MODEL = "gpt-4o-transcribe"
DEFAULT_CLOUD_CONFIG = {
    "profiles": [
        {
            "name": "DMXAPI-gpt-4o-transcribe",
            "api_url": DEFAULT_CLOUD_URL,
            "model": DEFAULT_CLOUD_MODEL,
            "api_key": "",
            "response_format": "verbose_json",
        }
    ]
}


# ─── 工具函数 ───────────────────────────────────────────────────────────────

def fmt_ts(seconds: float) -> str:
    total = max(0, int(round(seconds)))
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def split_with_time(text: str, start: float, end: float):
    """按标点和可读性规则拆分文本，并按字符数比例估算各句结束时间。"""
    text = (text or "").strip()
    if not text:
        return []
    text = re.sub(r"\s+", "", text)

    # 先按常见中文标点切句/分句；若标点极少，再按长度强制切分提升可读性。
    parts = [p.strip() for p in re.split(r"(?<=[。！？!?；;…，,：:])", text) if p.strip()]
    punc_count = len(re.findall(r"[。！？!?；;…，,：:]", text))
    almost_no_punc = punc_count <= 1 and len(text) > 36
    if almost_no_punc and parts:
        forced = []
        for p in parts:
            if len(p) <= 24:
                forced.append(p)
                continue
            for i in range(0, len(p), 24):
                seg = p[i:i + 24]
                if seg and i + 24 < len(p) and not re.search(r"[。！？!?；;…，,：:]$", seg):
                    seg += "。"
                forced.append(seg)
        parts = [x for x in forced if x]

    if len(parts) <= 1:
        return [(text, end)]
    total_chars = sum(len(p) for p in parts)
    if total_chars == 0 or end <= start:
        return [(p, end) for p in parts]
    out = []
    cur = start
    for idx, p in enumerate(parts):
        ratio = len(p) / total_chars
        nxt = end if idx == len(parts) - 1 else (cur + (end - start) * ratio)
        out.append((p, nxt))
        cur = nxt
    return out


def fix_vocabulary(model_dir: str) -> None:
    """
    faster-whisper 加载时需要 vocabulary.txt，
    若目录里只有 vocabulary.json 则自动复制一份 vocabulary.txt。
    """
    p = Path(model_dir)
    txt = p / "vocabulary.txt"
    jsn = p / "vocabulary.json"
    if not txt.exists() and jsn.exists():
        shutil.copy(jsn, txt)


def normalize_cloud_url(url: str) -> str:
    u = (url or "").strip().rstrip("/")
    if not u:
        return ""
    if u.endswith("/audio/transcriptions"):
        return u
    if u.endswith("/v1"):
        return u + "/audio/transcriptions"
    return u + "/audio/transcriptions"


def get_audio_duration_seconds(audio_file: Path) -> float:
    try:
        with closing(wave.open(str(audio_file), "rb")) as wf:
            frames = wf.getnframes()
            rate = wf.getframerate()
        return (frames / rate) if rate else 0.0
    except Exception:
        pass

    ffprobe = shutil.which("ffprobe")
    if not ffprobe:
        return 0.0

    cmd = [
        ffprobe,
        "-v",
        "error",
        "-show_entries",
        "format=duration",
        "-of",
        "default=noprint_wrappers=1:nokey=1",
        str(audio_file),
    ]
    try:
        out = subprocess.check_output(cmd, text=True).strip()
        return float(out)
    except Exception:
        return 0.0


def parse_chunk_index(name: str):
    m = re.search(r"chunk_(\d+)", name, re.IGNORECASE)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


# ─── 主窗口 ─────────────────────────────────────────────────────────────────

class TranscribeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("浮生记Podcast语音转文字")
        self.root.geometry("740x980")
        self.root.minsize(720, 860)

        self.audio_files: list[str] = []
        self.cloud_profiles: list[dict] = []
        self.cloud_format_cache: dict[str, str] = {}
        self.is_running = False
        self._stop_flag = False

        self._apply_style()
        self._build_ui()
        self._reload_cloud_profiles()

    # ── 外观 ────────────────────────────────────────────────────────────────

    def _apply_style(self):
        style = ttk.Style()
        try:
            self.root.tk.call("source", "")
        except Exception:
            pass
        style.theme_use("clam")
        style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("Start.TButton", font=("Microsoft YaHei UI", 10, "bold"),
                        foreground="white", background="#2d7d46")
        style.map("Start.TButton", background=[("active", "#246838")])

    # ── 界面布局 ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── 顶部：文件列表 ──
        file_frame = ttk.LabelFrame(self.root, text="  音频文件列表（支持批量）", padding=8)
        file_frame.pack(fill=tk.X, padx=12, pady=(10, 4))

        btn_row = ttk.Frame(file_frame)
        btn_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(btn_row, text="➕ 添加文件…", command=self._select_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="🗑 移除选中", command=self._remove_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="清空列表", command=self._clear_files).pack(side=tk.LEFT)
        self.count_label = ttk.Label(btn_row, text="共 0 个文件")
        self.count_label.pack(side=tk.RIGHT, padx=8)

        list_wrap = ttk.Frame(file_frame)
        list_wrap.pack(fill=tk.BOTH, expand=False)
        self.file_listbox = tk.Listbox(list_wrap, height=5, selectmode=tk.EXTENDED,
                                       font=("Consolas", 9), activestyle="dotbox", width=110)
        sb = ttk.Scrollbar(list_wrap, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.config(yscrollcommand=sb.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # ── 选项区 ──
        opt_frame = ttk.LabelFrame(self.root, text="  设置", padding=8)
        opt_frame.pack(fill=tk.X, padx=12, pady=4)
        opt_frame.columnconfigure(1, weight=1)

        # 模型路径
        ttk.Label(opt_frame, text="模型目录：").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.model_var = tk.StringVar(value=DEFAULT_MODEL_PATH)
        ttk.Entry(opt_frame, textvariable=self.model_var).grid(row=0, column=1, sticky=tk.EW, padx=6)
        ttk.Button(opt_frame, text="浏览…", command=self._browse_model).grid(row=0, column=2)

        # 语言
        ttk.Label(opt_frame, text="识别语言：").grid(row=1, column=0, sticky=tk.W, pady=3)
        lang_wrap = ttk.Frame(opt_frame)
        lang_wrap.grid(row=1, column=1, sticky=tk.W, padx=6)
        self.lang_var = tk.StringVar(value="zh")
        for val, lbl in [("zh", "中文"), ("en", "英文"), ("ja", "日文"), ("auto", "自动检测")]:
            ttk.Radiobutton(lang_wrap, text=lbl, variable=self.lang_var, value=val).pack(side=tk.LEFT, padx=4)

        # 输出目录
        ttk.Label(opt_frame, text="输出目录：").grid(row=2, column=0, sticky=tk.W, pady=3)
        out_wrap = ttk.Frame(opt_frame)
        out_wrap.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=6)
        out_wrap.columnconfigure(0, weight=1)
        self.output_var = tk.StringVar(value="")
        ttk.Entry(out_wrap, textvariable=self.output_var).grid(row=0, column=0, sticky=tk.EW)
        ttk.Button(out_wrap, text="浏览…", command=self._browse_output).grid(row=0, column=1, padx=(4, 0))
        ttk.Label(out_wrap, text="（留空 = 与音频文件同目录）", foreground="gray").grid(row=0, column=2, padx=4)

        # 模式
        ttk.Label(opt_frame, text="转写模式：").grid(row=3, column=0, sticky=tk.W, pady=3)
        mode_wrap = ttk.Frame(opt_frame)
        mode_wrap.grid(row=3, column=1, sticky=tk.W, padx=6)
        self.engine_var = tk.StringVar(value="local")
        ttk.Radiobutton(mode_wrap, text="本地模型", variable=self.engine_var, value="local").pack(side=tk.LEFT, padx=4)
        ttk.Radiobutton(mode_wrap, text="云端API", variable=self.engine_var, value="cloud").pack(side=tk.LEFT, padx=4)

        # 分段
        ttk.Label(opt_frame, text="分段秒数：").grid(row=4, column=0, sticky=tk.W, pady=3)
        chunk_wrap = ttk.Frame(opt_frame)
        chunk_wrap.grid(row=4, column=1, columnspan=2, sticky=tk.W, padx=6)
        self.chunk_enable_var = tk.BooleanVar(value=False)
        self.chunk_sec_var = tk.StringVar(value="300")
        ttk.Checkbutton(chunk_wrap, text="启用分段切分", variable=self.chunk_enable_var).pack(side=tk.LEFT)
        ttk.Entry(chunk_wrap, textvariable=self.chunk_sec_var, width=8).pack(side=tk.LEFT, padx=(8, 4))
        ttk.Label(chunk_wrap, text="秒（建议 180~600）", foreground="gray").pack(side=tk.LEFT)

        # 云端 API
        cloud_frame = ttk.LabelFrame(opt_frame, text="云端模型参数", padding=6)
        cloud_frame.grid(row=5, column=0, columnspan=3, sticky=tk.EW, pady=(6, 0))
        cloud_frame.columnconfigure(1, weight=1)
        ttk.Label(cloud_frame, text="云端模型配置：").grid(row=0, column=0, sticky=tk.W, pady=2)
        profile_wrap = ttk.Frame(cloud_frame)
        profile_wrap.grid(row=0, column=1, sticky=tk.EW, padx=6)
        profile_wrap.columnconfigure(0, weight=1)
        self.cloud_profile_var = tk.StringVar(value="")
        self.cloud_profile_combo = ttk.Combobox(
            profile_wrap,
            textvariable=self.cloud_profile_var,
            state="readonly",
        )
        self.cloud_profile_combo.grid(row=0, column=0, sticky=tk.EW)
        self.cloud_profile_combo.bind("<<ComboboxSelected>>", self._on_cloud_profile_selected)
        ttk.Button(profile_wrap, text="重载", command=self._reload_cloud_profiles).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(profile_wrap, text="保存配置", command=self._save_cloud_profile).grid(row=0, column=2, padx=(4, 0))

        ttk.Label(cloud_frame, text="云端模型请求地址：").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.cloud_url_var = tk.StringVar(value=DEFAULT_CLOUD_URL)
        ttk.Entry(cloud_frame, textvariable=self.cloud_url_var).grid(row=1, column=1, sticky=tk.EW, padx=6)

        ttk.Label(cloud_frame, text="模型名称：").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.cloud_model_var = tk.StringVar(value=DEFAULT_CLOUD_MODEL)
        ttk.Entry(cloud_frame, textvariable=self.cloud_model_var).grid(row=2, column=1, sticky=tk.EW, padx=6)

        ttk.Label(cloud_frame, text="API Key：").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.cloud_key_var = tk.StringVar(value="")
        ttk.Entry(cloud_frame, textvariable=self.cloud_key_var, show="*").grid(row=3, column=1, sticky=tk.EW, padx=6)

        ttk.Label(
            cloud_frame,
            text="文档参考：https://doc.dmxapi.cn/gpt-4o-transcribe.html",
            foreground="gray",
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(2, 0))

        # ── 进度区 ──
        prog_frame = ttk.LabelFrame(self.root, text="  进度", padding=8)
        prog_frame.pack(fill=tk.X, padx=12, pady=4)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.progress_var,
                                            maximum=100, mode="determinate")
        self.progress_bar.pack(fill=tk.X)

        self.status_var = tk.StringVar(value="就绪，请添加音频文件后点击「开始转写」。")
        ttk.Label(prog_frame, textvariable=self.status_var, foreground="#2d7d46").pack(
            anchor=tk.W, pady=(4, 0))

        # ── 底部按钮（必须在 log_frame 之前 pack，否则会被 expand 挤出视图）──
        bot_frame = ttk.Frame(self.root)
        bot_frame.pack(fill=tk.X, padx=12, pady=(4, 10), side=tk.BOTTOM)

        self.start_btn = ttk.Button(bot_frame, text="▶  开始转写",
                                    command=self._start, style="Start.TButton", width=16)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.split_btn = ttk.Button(bot_frame, text="✂  仅分段切分", command=self._split_only, width=13)
        self.split_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ttk.Button(bot_frame, text="■  停止", command=self._stop,
                                   state=tk.DISABLED, width=10)
        self.stop_btn.pack(side=tk.LEFT)

        ttk.Button(bot_frame, text="🧩 合并分段文稿", command=self._merge_chunk_markdowns).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Button(bot_frame, text="📂 打开输出目录", command=self._open_output).pack(side=tk.RIGHT)

        # ── 日志区（expand=True 必须在底部按钮之后 pack）──
        log_frame = ttk.LabelFrame(self.root, text="  运行日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, font=("Consolas", 9),
            background="#1e1e1e", foreground="#d4d4d4", insertbackground="white")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        # 颜色标签
        self.log_text.tag_config("ok", foreground="#4ec9b0")
        self.log_text.tag_config("err", foreground="#f44747")
        self.log_text.tag_config("info", foreground="#9cdcfe")

    # ── 文件操作 ─────────────────────────────────────────────────────────────

    def _select_files(self):
        files = filedialog.askopenfilenames(
            title="选择音频文件",
            filetypes=[
                ("音频文件", "*.wav *.mp3 *.mp4 *.m4a *.flac *.ogg *.aac *.wma *.WAV *.MP3"),
                ("所有文件", "*.*"),
            ],
        )
        for f in files:
            if f not in self.audio_files:
                self.audio_files.append(f)
                duration = get_audio_duration_seconds(Path(f))
                if duration > 0:
                    item = f"{Path(f).name}    [{fmt_ts(duration)} | {int(round(duration))} 秒]"
                else:
                    item = f"{Path(f).name}    [--:--:-- | -- 秒]"
                self.file_listbox.insert(tk.END, item)
        self._update_count()

    def _remove_selected(self):
        selected = list(self.file_listbox.curselection())
        for idx in reversed(selected):
            self.file_listbox.delete(idx)
            self.audio_files.pop(idx)
        self._update_count()

    def _clear_files(self):
        self.audio_files.clear()
        self.file_listbox.delete(0, tk.END)
        self._update_count()

    def _update_count(self):
        self.count_label.config(text=f"共 {len(self.audio_files)} 个文件")

    def _browse_model(self):
        d = filedialog.askdirectory(title="选择模型目录（包含 model.bin 的文件夹）")
        if d:
            self.model_var.set(d)

    def _browse_output(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.output_var.set(d)

    def _open_output(self):
        d = self.output_var.get().strip()
        if d and Path(d).exists():
            os.startfile(d)
        elif self.audio_files:
            os.startfile(str(Path(self.audio_files[0]).parent))

    def _cloud_config_path(self) -> Path:
        return get_app_dir() / "cloud_models.json"

    def _ensure_cloud_config(self):
        p = self._cloud_config_path()
        if p.exists():
            return
        p.write_text(
            json.dumps(DEFAULT_CLOUD_CONFIG, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _reload_cloud_profiles(self):
        try:
            self._ensure_cloud_config()
            data = json.loads(self._cloud_config_path().read_text(encoding="utf-8"))
            profiles = data.get("profiles") or []
            cleaned = []
            for p in profiles:
                name = str(p.get("name") or "").strip()
                api_url = str(p.get("api_url") or "").strip()
                model = str(p.get("model") or "").strip()
                api_key = str(p.get("api_key") or "").strip()
                response_format = str(p.get("response_format") or "").strip()
                if name and api_url and model:
                    cleaned.append({
                        "name": name,
                        "api_url": api_url,
                        "model": model,
                        "api_key": api_key,
                        "response_format": response_format,
                    })
            if not cleaned:
                cleaned = DEFAULT_CLOUD_CONFIG["profiles"]
            self.cloud_profiles = cleaned
            self.cloud_format_cache.clear()
            for p in self.cloud_profiles:
                fmt = str(p.get("response_format") or "").strip()
                if fmt in {"verbose_json", "json", "text"}:
                    key = f"{p.get('api_url', '').strip()}|{p.get('model', '').strip()}"
                    self.cloud_format_cache[key] = fmt
            names = [p["name"] for p in cleaned]
            self.cloud_profile_combo["values"] = names
            self.cloud_profile_combo.current(0)
            self._apply_cloud_profile(0)
        except Exception as e:
            self._log(f"✗ 读取云端模型配置失败：{e}", "err")

    def _apply_cloud_profile(self, idx: int):
        if idx < 0 or idx >= len(self.cloud_profiles):
            return
        p = self.cloud_profiles[idx]
        self.cloud_profile_var.set(p.get("name", ""))
        self.cloud_url_var.set(p.get("api_url", ""))
        self.cloud_model_var.set(p.get("model", ""))
        fmt = str(p.get("response_format") or "").strip()
        if fmt in {"verbose_json", "json", "text"}:
            key = f"{p.get('api_url', '').strip()}|{p.get('model', '').strip()}"
            self.cloud_format_cache[key] = fmt
        # 仅在配置里有非空 key 时覆盖，避免误清空用户临时输入。
        cfg_key = p.get("api_key", "")
        if cfg_key:
            self.cloud_key_var.set(cfg_key)

    def _persist_cloud_profile_format(self, api_url: str, model: str, used_format: str) -> bool:
        if used_format not in {"verbose_json", "json", "text"}:
            return False

        target_name = self.cloud_profile_var.get().strip()
        changed = False

        for p in self.cloud_profiles:
            p_name = str(p.get("name") or "").strip()
            p_url = str(p.get("api_url") or "").strip()
            p_model = str(p.get("model") or "").strip()
            if (target_name and p_name == target_name) or (p_url == api_url and p_model == model):
                old_fmt = str(p.get("response_format") or "").strip()
                if old_fmt != used_format:
                    p["response_format"] = used_format
                    changed = True
                break

        if not changed:
            return False

        payload = {"profiles": self.cloud_profiles}
        self._cloud_config_path().write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return True

    def _save_cloud_profile(self):
        """将界面上填写的云端配置保存（更新或新增）到 cloud_models.json。"""
        name = self.cloud_profile_var.get().strip()
        api_url = self.cloud_url_var.get().strip()
        model = self.cloud_model_var.get().strip()
        api_key = self.cloud_key_var.get().strip()

        if not name:
            messagebox.showwarning("保存配置", "请先在下拉框中选择或输入一个配置名称。")
            return
        if not api_url or not model:
            messagebox.showwarning("保存配置", "请填写云端模型请求地址和模型名称。")
            return

        try:
            self._ensure_cloud_config()
            data = json.loads(self._cloud_config_path().read_text(encoding="utf-8"))
            profiles = data.get("profiles") or []

            # 查找同名 profile 更新，否则追加
            updated = False
            for p in profiles:
                if str(p.get("name") or "").strip() == name:
                    p["api_url"] = api_url
                    p["model"] = model
                    if api_key:
                        p["api_key"] = api_key
                    updated = True
                    break
            if not updated:
                entry = {"name": name, "api_url": api_url, "model": model, "api_key": api_key}
                profiles.append(entry)

            data["profiles"] = profiles
            self._cloud_config_path().write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            # 同步内存缓存
            self._reload_cloud_profiles()
            self._log(f"✓ 云端配置 [{name}] 已保存到 cloud_models.json", "ok")
        except Exception as e:
            messagebox.showerror("保存配置", f"保存失败：{e}")

    def _on_cloud_profile_selected(self, _evt=None):
        name = self.cloud_profile_var.get().strip()
        for i, p in enumerate(self.cloud_profiles):
            if p.get("name") == name:
                self._apply_cloud_profile(i)
                return

    def _merge_chunk_markdowns(self):
        initial_dir = self.output_var.get().strip()
        if not initial_dir:
            initial_dir = str(Path(self.audio_files[0]).parent) if self.audio_files else str(get_app_dir())

        target_dir = filedialog.askdirectory(title="选择分段文稿目录（包含 chunk_*.md）", initialdir=initial_dir)
        if not target_dir:
            return

        p = Path(target_dir)
        files = sorted(p.glob("chunk_*.md"), key=lambda x: x.name)
        if not files:
            messagebox.showwarning("提示", "未找到 chunk_*.md 文件")
            return

        ts_pattern = re.compile(r"[（(]\d{2}:\d{2}:\d{2}[）)]")
        merged_lines = [
            f"# {p.name} 合并文稿",
            "",
        ]

        kept = 0
        for f in files:
            lines = f.read_text(encoding="utf-8").splitlines()
            body = []
            for line in lines:
                t = line.strip()
                if not t:
                    continue
                if ts_pattern.search(t):
                    body.append(t)
            if body:
                kept += 1
                merged_lines.extend(body)
                merged_lines.append("")

        out_file = p / f"{p.name}_merged_publish.md"
        out_file.write_text("\n".join(merged_lines).strip() + "\n", encoding="utf-8")
        self._log(f"✓ 文稿合并完成：{out_file}  [源文件 {len(files)} 个，保留 {kept} 个]", "ok")
        self._set_status("文稿合并完成 ✓")
        messagebox.showinfo("完成", f"合并完成：\n{out_file}")

    # ── 日志输出 ─────────────────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = ""):
        def _do():
            self.log_text.insert(tk.END, msg + "\n", tag)
            self.log_text.see(tk.END)
        self.root.after(0, _do)

    def _set_status(self, msg: str):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _set_progress(self, pct: float):
        self.root.after(0, lambda: self.progress_var.set(pct))

    # ── 转写逻辑 ─────────────────────────────────────────────────────────────

    def _start(self):
        if not self.audio_files:
            messagebox.showwarning("提示", "请先添加音频文件！")
            return
        if self.is_running:
            return
        self.is_running = True
        self._stop_flag = False
        self.start_btn.config(state=tk.DISABLED)
        self.split_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        threading.Thread(target=self._worker, daemon=True).start()

    def _stop(self):
        self._stop_flag = True
        self._set_status("正在停止，等待当前文件处理完毕…")

    def _done(self):
        self.is_running = False
        self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.split_btn.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))

    def _get_chunk_seconds(self) -> int:
        try:
            sec = int(self.chunk_sec_var.get().strip())
        except Exception:
            raise ValueError("分段秒数必须是整数")
        if sec < 30:
            raise ValueError("分段秒数太小，建议至少 30 秒")
        return sec

    def _ensure_ffmpeg(self):
        if shutil.which("ffmpeg"):
            return
        raise RuntimeError("未找到 ffmpeg。请先安装 ffmpeg 并加入 PATH。")

    def _split_audio(self, audio_path: Path, chunk_sec: int) -> list[tuple[Path, float]]:
        self._ensure_ffmpeg()
        chunk_dir = audio_path.parent / f"{audio_path.stem}_chunks"
        chunk_dir.mkdir(parents=True, exist_ok=True)
        out_pattern = chunk_dir / "chunk_%03d.wav"

        cmd = [
            "ffmpeg",
            "-hide_banner",
            "-loglevel",
            "error",
            "-y",
            "-i",
            str(audio_path),
            "-f",
            "segment",
            "-segment_time",
            str(chunk_sec),
            "-ar",
            "16000",
            "-ac",
            "1",
            "-c:a",
            "pcm_s16le",
            str(out_pattern),
        ]
        subprocess.run(cmd, check=True)

        chunks = sorted(chunk_dir.glob("chunk_*.wav"))
        if not chunks:
            raise RuntimeError("分段失败：未生成任何分段文件")

        out: list[tuple[Path, float]] = []
        offset = 0.0
        for c in chunks:
            out.append((c, offset))
            offset += get_audio_duration_seconds(c)
        return out

    def _cloud_transcribe_chunk(self, audio_path: Path, lang: str | None):
        try:
            import requests
        except ImportError:
            raise RuntimeError("未安装 requests，请执行：pip install requests")

        api_url = normalize_cloud_url(self.cloud_url_var.get())
        api_model = self.cloud_model_var.get().strip()
        api_key = self.cloud_key_var.get().strip()

        if not api_url:
            raise RuntimeError("云端模型请求地址不能为空")
        if not api_model:
            raise RuntimeError("模型名称不能为空")
        if not api_key:
            raise RuntimeError("API Key 不能为空")

        headers = {
            "Authorization": f"Bearer {api_key}",
        }

        cache_key = f"{api_url}|{api_model}"
        # 兼容不同云模型：优先使用缓存的已知可用格式，避免每次先报不兼容再降级。
        cached_fmt = self.cloud_format_cache.get(cache_key)
        if cached_fmt in {"verbose_json", "json", "text"}:
            format_candidates = [cached_fmt] + [x for x in ["verbose_json", "json", "text"] if x != cached_fmt]
        else:
            format_candidates = ["verbose_json", "json", "text"]
        last_error = ""

        for fmt in format_candidates:
            data = {
                "model": api_model,
                "response_format": fmt,
            }
            if lang:
                data["language"] = lang

            with audio_path.open("rb") as f:
                files = {
                    "file": (audio_path.name, f, "audio/wav"),
                }
                resp = requests.post(api_url, headers=headers, data=data, files=files, timeout=900)

            if resp.status_code < 400:
                self.cloud_format_cache[cache_key] = fmt
                config_updated = self._persist_cloud_profile_format(api_url, api_model, fmt)
                if fmt == "text":
                    text = (resp.text or "").strip()
                    return [], text, (lang or "unknown"), fmt, config_updated

                try:
                    payload = resp.json()
                except Exception:
                    payload = {}
                segments = payload.get("segments") or []
                text = payload.get("text") or ""
                lang_info = payload.get("language") or (lang or "unknown")
                return segments, text, lang_info, fmt, config_updated

            body = (resp.text or "")[:500]
            last_error = f"HTTP {resp.status_code}: {body}"

            # 只有 response_format 不兼容时才继续降级，其他错误立即抛出。
            unsupported_fmt = (
                "response_format" in body and
                ("not compatible" in body or "unsupported_value" in body)
            )
            if not unsupported_fmt:
                raise RuntimeError(f"云端接口错误 {last_error}")
            # 无感降级：不输出错误感叹提示，继续尝试下一个兼容格式。

        raise RuntimeError(f"云端接口错误 {last_error}")

    def _split_only(self):
        if not self.audio_files:
            messagebox.showwarning("提示", "请先添加音频文件！")
            return
        if self.is_running:
            return
        try:
            chunk_sec = self._get_chunk_seconds()
        except Exception as e:
            messagebox.showerror("参数错误", str(e))
            return
        self.is_running = True
        self._stop_flag = False
        self.start_btn.config(state=tk.DISABLED)
        self.split_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        threading.Thread(target=self._split_worker, args=(chunk_sec,), daemon=True).start()

    def _split_worker(self, chunk_sec: int):
        total = len(self.audio_files)
        self._set_status("开始分段切分…")
        for idx, audio_file in enumerate(self.audio_files):
            if self._stop_flag:
                self._log("⛔ 用户已停止。", "err")
                break
            audio_path = Path(audio_file)
            self._log(f"[{idx+1}/{total}]  切分 {audio_path.name}", "info")
            self._set_progress(idx / total * 100)
            try:
                chunks = self._split_audio(audio_path, chunk_sec)
                chunk_dir = audio_path.parent / f"{audio_path.stem}_chunks"
                self._log(f"  ✓ 已切分：{chunk_dir}  [{len(chunks)} 段]", "ok")
            except Exception as e:
                self._log(f"  ✗ 切分错误：{e}", "err")
        self._set_progress(100)
        if not self._stop_flag:
            self._set_status("分段切分完成 ✓")
            self._log("\n✅ 分段切分完成。", "ok")
        self._done()

    def _worker(self):
        lang_sel = self.lang_var.get().strip()
        lang = None if lang_sel == "auto" else lang_sel
        use_cloud = self.engine_var.get().strip() == "cloud"
        model = None

        if not use_cloud:
            try:
                from faster_whisper import WhisperModel
            except ImportError:
                self._log("✗ faster-whisper 未安装，请在终端执行：", "err")
                self._log("  pip install faster-whisper", "info")
                self._done()
                return

            model_dir = self.model_var.get().strip()
            # 相对路径转换为绝对路径（相对于 exe/脚本所在目录）
            model_path = Path(model_dir)
            if not model_path.is_absolute():
                model_path = get_app_dir() / model_path
            model_dir = str(model_path)
            if not model_path.exists():
                self._log(f"✗ 模型目录不存在：{model_dir}", "err")
                self._done()
                return

            fix_vocabulary(model_dir)

            self._log(f"⏳ 加载模型：{model_dir}", "info")
            self._set_status("加载模型中…")
            try:
                model = WhisperModel(
                    model_dir,
                    device="cpu",
                    compute_type="int8",
                    local_files_only=True,
                )
            except Exception as e:
                self._log(f"✗ 模型加载失败：{e}", "err")
                self._done()
                return

            self._log("✓ 模型加载完成\n", "ok")
        else:
            self._log("⏳ 使用云端模型转写（按你填写的 API 参数）", "info")
        cloud_fmt_logged = False

        chunk_enable = bool(self.chunk_enable_var.get())
        try:
            chunk_sec = self._get_chunk_seconds()
        except Exception as e:
            self._log(f"✗ 参数错误：{e}", "err")
            self._done()
            return

        # 若用户直接导入 chunk_000.wav ~ chunk_N.wav，按文件顺序累计全局偏移，
        # 让每个文件写入的时间码都对齐到原始长音频时间轴。
        file_base_offsets = {}
        chunk_like = []
        same_parent = True
        parent_ref = None
        for f in self.audio_files:
            p = Path(f)
            idx = parse_chunk_index(p.stem)
            if idx is None:
                chunk_like = []
                break
            chunk_like.append((idx, p))
            if parent_ref is None:
                parent_ref = p.parent
            elif p.parent != parent_ref:
                same_parent = False

        if chunk_like and same_parent:
            offset = 0.0
            for _, p in sorted(chunk_like, key=lambda x: x[0]):
                file_base_offsets[str(p)] = offset
                offset += get_audio_duration_seconds(p)
            self._log("✓ 检测到分段文件序列，已启用跨文件累计时间轴", "ok")
        else:
            for f in self.audio_files:
                file_base_offsets[str(Path(f))] = 0.0

        total = len(self.audio_files)

        for idx, audio_file in enumerate(self.audio_files):
            if self._stop_flag:
                self._log("⛔ 用户已停止。", "err")
                break

            audio_path = Path(audio_file)
            out_dir = self.output_var.get().strip()
            if out_dir:
                Path(out_dir).mkdir(parents=True, exist_ok=True)
                md_path = Path(out_dir) / audio_path.with_suffix(".md").name
            else:
                md_path = audio_path.with_suffix(".md")

            self._log(f"[{idx+1}/{total}]  {audio_path.name}", "info")
            self._set_status(f"转写中 {idx+1}/{total}：{audio_path.name}")
            self._set_progress(idx / total * 100)

            try:
                file_base_offset = file_base_offsets.get(str(audio_path), 0.0)
                if chunk_enable:
                    chunk_meta = self._split_audio(audio_path, chunk_sec)
                    chunk_meta = [(cp, file_base_offset + co) for cp, co in chunk_meta]
                    self._log(f"  ✓ 分段完成：{len(chunk_meta)} 段（每段约 {chunk_sec} 秒）", "ok")
                else:
                    chunk_meta = [(audio_path, file_base_offset)]

                chunk_md_dir = md_path.parent / f"{audio_path.stem}_chunks_md"
                if chunk_enable:
                    chunk_md_dir.mkdir(parents=True, exist_ok=True)

                header = [
                    f"# {audio_path.stem} 语音文字稿",
                    "",
                    f"音频文件：{audio_path.name}",
                    f"识别语言：{lang or 'auto'}",
                    f"模式：{'云端API' if use_cloud else '本地模型'}",
                    f"分段：{'启用' if chunk_enable else '关闭'}（{chunk_sec} 秒）",
                    "",
                ]
                body_lines = []
                sentence_count = 0
                detected_lang = lang or "unknown"
                processed_chunks = 0

                for cidx, (chunk_path, chunk_offset) in enumerate(chunk_meta):
                    if self._stop_flag:
                        break

                    self._log(f"  [{cidx+1}/{len(chunk_meta)}] {chunk_path.name}", "info")
                    chunk_lines = []

                    try:
                        if use_cloud:
                            segments, full_text, cloud_lang, used_fmt, cfg_updated = self._cloud_transcribe_chunk(chunk_path, lang)
                            detected_lang = cloud_lang or detected_lang
                            if not cloud_fmt_logged:
                                self._log(f"✓ 云端格式：{used_fmt}", "ok")
                                cloud_fmt_logged = True
                            if cfg_updated:
                                self._log("✓ 云端模型配置已更新：已写入可用 response_format", "ok")

                            if segments:
                                for seg in segments:
                                    seg_text = (seg.get("text") or "").strip()
                                    if not seg_text:
                                        continue
                                    seg_start = float(seg.get("start") or 0.0) + chunk_offset
                                    seg_end = float(seg.get("end") or seg.get("start") or 0.0) + chunk_offset
                                    for sentence, t_end in split_with_time(seg_text, seg_start, seg_end):
                                        sentence_count += 1
                                        chunk_lines.append(f"{sentence}（{fmt_ts(t_end)}）")
                            else:
                                text = (full_text or "").strip()
                                if text:
                                    chunk_duration = get_audio_duration_seconds(chunk_path)
                                    for sentence, t_end in split_with_time(text, chunk_offset, chunk_offset + chunk_duration):
                                        sentence_count += 1
                                        chunk_lines.append(f"{sentence}（{fmt_ts(t_end)}）")
                        else:
                            segments, info = model.transcribe(
                                str(chunk_path),
                                language=lang,
                                vad_filter=True,
                                beam_size=5,
                                word_timestamps=False,
                            )
                            detected_lang = info.language or detected_lang

                            for seg in segments:
                                if self._stop_flag:
                                    break
                                seg_text = (seg.text or "").strip()
                                if not seg_text:
                                    continue
                                seg_start = seg.start + chunk_offset
                                seg_end = seg.end + chunk_offset
                                for sentence, t_end in split_with_time(seg_text, seg_start, seg_end):
                                    sentence_count += 1
                                    chunk_lines.append(f"{sentence}（{fmt_ts(t_end)}）")
                    except Exception as ce:
                        self._log(f"    ✗ 分段错误：{ce}", "err")

                    if not chunk_lines:
                        chunk_lines = ["（本分段无可用文字结果）"]

                    processed_chunks += 1
                    body_lines.extend(chunk_lines)

                    if chunk_enable:
                        chunk_duration = get_audio_duration_seconds(chunk_path)
                        chunk_start_ts = fmt_ts(chunk_offset)
                        chunk_end_ts = fmt_ts(chunk_offset + chunk_duration)
                        chunk_header = [
                            f"# {audio_path.stem} - {chunk_path.stem}",
                            "",
                            f"原始音频：{audio_path.name}",
                            f"分段文件：{chunk_path.name}",
                            f"时间范围：{chunk_start_ts} ~ {chunk_end_ts}（原始长音频时间）",
                            "",
                        ]
                        chunk_footer = ["", f"累计总句数（到本段）：{sentence_count} 句。"]
                        chunk_md_path = chunk_md_dir / f"{chunk_path.stem}.md"
                        chunk_md_path.write_text("\n".join(chunk_header + chunk_lines + chunk_footer), encoding="utf-8")
                        self._log(f"    ✓ 分段文档已保存：{chunk_md_path}", "ok")

                    header[3] = f"识别语言：{detected_lang}"
                    footer = [
                        "",
                        f"已处理分段：{processed_chunks}/{len(chunk_meta)}",
                        f"当前累计：{sentence_count} 句。",
                    ]
                    md_path.write_text("\n".join(header + body_lines + footer), encoding="utf-8")
                    self._log(f"    ✓ 总稿已增量保存：{md_path}", "ok")

                if not body_lines:
                    body_lines.append("（无可用文字结果）")

                header[3] = f"识别语言：{detected_lang}"
                final_footer = ["", f"共 {sentence_count} 句。"]
                md_path.write_text("\n".join(header + body_lines + final_footer), encoding="utf-8")
                self._log(f"  ✓ 已保存最终稿：{md_path}  [{sentence_count} 句]", "ok")

            except Exception as e:
                self._log(f"  ✗ 错误：{e}", "err")

        self._set_progress(100)
        if not self._stop_flag:
            self._set_status("全部完成 ✓")
            self._log("\n✅ 全部转写完成。", "ok")
        self._done()


# ─── 启动入口 ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app = TranscribeApp(root)
    root.mainloop()
