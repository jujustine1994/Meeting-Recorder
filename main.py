"""
Meeting Recorder
錄製電腦系統音訊（WASAPI Loopback）、麥克風或兩者混音，儲存為 MP3

模式說明：
  system  — WASAPI Loopback，捕捉系統所有輸出音訊
  mic     — 系統預設麥克風
  both    — 兩者同時錄製，存檔前混音成單一 MP3
"""

import os
import array
import math
import struct
import threading
import datetime
import time
import queue
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

import pyaudiowpatch as pyaudio
import lameenc


# ---- 可自行調整的進階參數（直接編輯此區段） ----
_DEFAULT_BIT_RATE       = 128    # kbps：MP3 位元率預設值，可在進階設定中更改
_MIX_WEIGHT_SYSTEM      = 0.6    # 系統音軌混音權重（兩軌混音時）
_MIX_WEIGHT_MIC         = 0.6    # 麥克風混音權重
_SILENCE_RMS_THRESHOLD  = 100    # 靜音判斷 RMS 閾值（Int16 0~32767）
_SILENCE_WARNING_SECS   = 10     # 連續靜音幾秒後顯示警告
_VU_DISPLAY_DIVISOR     = 80     # VU meter 顯示除數：正常說話 RMS 換算後約落在 20~50%
_CLIP_PEAK_THRESHOLD    = int(32767 * 0.9)   # 峰值達此值視為接近爆音（約 29491）
_CLIP_WARNING_HOLD_SECS = 1.5    # 爆音警示保持顯示的秒數，避免瞬間峰值一閃即逝


# ---- CTH Banner（終端機用，launcher 視窗可見）----
def show_cth_banner():
    b = "\033[90m"   # 邊框：深灰
    c = "\033[96m"   # CTH 字母：亮青
    y = "\033[93m"   # 署名：金黃
    r = "\033[0m"    # reset

    print(f"{b}/*  ================================  *\\{r}")
    print(f"{b} *                                    *{r}")
    print(f"{b} *    {c}██████╗████████╗██╗  ██╗{b}        *{r}")
    print(f"{b} *   {c}██╔════╝   ██║   ██║  ██║{b}        *{r}")
    print(f"{b} *   {c}██║        ██║   ███████║{b}        *{r}")
    print(f"{b} *   {c}██║        ██║   ██╔══██║{b}        *{r}")
    print(f"{b} *   {c}╚██████╗   ██║   ██║  ██║{b}        *{r}")
    print(f"{b} *    {c}╚═════╝   ╚═╝   ╚═╝  ╚═╝{b}        *{r}")
    print(f"{b} *                                    *{r}")
    print(f"{b} *          {y}created by CTH{b}            *{r}")
    print(f"{b}\\*  ================================  */{r}")
    print()


# ---- 音訊裝置 ----
def get_loopback_device(p: pyaudio.PyAudio, preferred_output_name: str = None) -> dict:
    """
    取得 WASAPI Loopback 裝置。

    preferred_output_name 不為 None 時，優先找名稱包含該字串的 loopback 裝置
    （對應使用者在裝置測試中選擇的輸出裝置）。
    找不到或未指定時，退回系統預設輸出裝置對應的 loopback。
    """
    try:
        wasapi_info = p.get_host_api_info_by_type(pyaudio.paWASAPI)
    except OSError:
        raise RuntimeError("找不到 WASAPI 音訊裝置，請確認音效卡驅動正常。")

    if preferred_output_name:
        for loopback in p.get_loopback_device_info_generator():
            if preferred_output_name in loopback["name"]:
                return loopback

    default_speakers = p.get_device_info_by_index(wasapi_info["defaultOutputDevice"])

    if not default_speakers.get("isLoopbackDevice", False):
        for loopback in p.get_loopback_device_info_generator():
            if default_speakers["name"] in loopback["name"]:
                return loopback

    return default_speakers


# ---- 靜音偵測輔助 ----
def _compute_rms(data: bytes) -> float:
    """
    計算 PCM Int16 音訊資料的 RMS（均方根）音量。
    回傳範圍 0.0 ~ 32767.0，靜音接近 0。
    """
    num_samples = len(data) // 2
    if num_samples == 0:
        return 0.0
    samples = struct.unpack(f"{num_samples}h", data)
    return math.sqrt(sum(s * s for s in samples) / num_samples)


def _rms_to_display_pct(rms: float) -> float:
    """把 RMS（0~32767）換算成 VU meter 顯示用的 0~100 百分比。"""
    return max(0.0, min(100.0, rms / _VU_DISPLAY_DIVISOR))


def _compute_peak(data: bytes) -> int:
    """
    計算 PCM Int16 音訊資料的峰值（樣本最大絕對值）。
    回傳範圍 0 ~ 32767，用於偵測爆音，與 RMS（平均音量）分開判斷。
    """
    num_samples = len(data) // 2
    if num_samples == 0:
        return 0
    samples = struct.unpack(f"{num_samples}h", data)
    return max(abs(s) for s in samples)


def _is_clipping(peak: int) -> bool:
    """峰值是否達到爆音警示閾值。"""
    return peak >= _CLIP_PEAK_THRESHOLD


# ---- 主視窗 ----
class MeetingRecorderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Meeting Recorder")
        self.root.resizable(False, False)

        # 錄音狀態
        self.is_recording = False
        self.record_frames: list[bytes] = []   # loopback 音訊暫存
        self.mic_frames:    list[bytes] = []   # 麥克風音訊暫存
        self.record_channels    = 2            # loopback 聲道數（由裝置決定，最多 2）
        self.record_sample_rate = 44100        # loopback 取樣率（由裝置決定）
        self.record_mic_channels    = 1        # 麥克風固定 mono
        self.record_mic_rate        = 44100    # 麥克風實際使用的取樣率
        self.start_time: float = 0.0
        self.is_paused: bool = False
        self._elapsed_before_pause: float = 0.0  # 暫停前已累計的秒數
        self.msg_queue: queue.Queue = queue.Queue()
        self._progress_line_active: bool = False  # log 最後一行是否為可覆寫的進度行
        self._record_thread: threading.Thread | None = None
        self._mic_thread:    threading.Thread | None = None
        self._record_stream = None   # 供外部呼叫 stop_stream() 解除 read() 阻塞
        self._mic_stream    = None
        self._save_mode: str = "system"        # 儲存時使用的模式，在停止時鎖定避免 race condition
        self._save_output_mode: str = "merge"      # 輸出方式快照（停止時從主執行緒鎖定）
        self._save_equalize: bool = False           # 等化開關快照
        self._save_gain_cap: float = 4.0            # 增益上限快照
        self._save_filter_silence: bool = True      # 靜音過濾快照
        self._save_bit_rate: int = _DEFAULT_BIT_RATE   # 停止時從主執行緒鎖定

        # 裝置選擇（None = 系統預設）
        self.selected_input_idx:    int | None = None   # 麥克風裝置 index
        self.selected_output_name:  str | None = None   # 輸出裝置名稱（用於比對 loopback）

        # 儲存設定
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.save_folder = desktop
        self.save_folder_var = tk.StringVar(value=desktop)
        self.filename_var    = tk.StringVar()

        # VU Meter
        self.vu_system_var = tk.DoubleVar(value=0)
        self.vu_mic_var    = tk.DoubleVar(value=0)

        # 進階設定
        self.output_mode       = tk.StringVar(value="merge")   # merge / separate / both
        self.equalize_enabled  = tk.BooleanVar(value=False)
        self.eq_gain_cap       = tk.IntVar(value=4)            # 倍數上限 1～16
        self.eq_filter_silence = tk.BooleanVar(value=True)
        self.bit_rate          = tk.IntVar(value=_DEFAULT_BIT_RATE)

        # 錄音模式（UI 用，tk.StringVar 只在主執行緒存取）
        self.record_mode = tk.StringVar(value="system")

        self._build_ui()
        self._poll_queue()

    # ---- UI 建置 ----
    def _build_ui(self):
        pad = {"padx": 14, "pady": 6}

        # row=0  儲存位置
        frame_folder = ttk.LabelFrame(self.root, text=" 儲存位置 ", padding=8)
        frame_folder.grid(row=0, column=0, sticky="ew", **pad)
        frame_folder.columnconfigure(0, weight=1)

        ttk.Entry(
            frame_folder, textvariable=self.save_folder_var,
            state="readonly", width=44
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            frame_folder, text="變更", command=self._change_folder
        ).grid(row=0, column=1)

        # row=1  錄音模式
        frame_mode = ttk.LabelFrame(self.root, text=" 錄音模式 ", padding=8)
        frame_mode.grid(row=1, column=0, sticky="ew", **pad)

        modes = [
            ("電腦聲音",      "system"),
            ("麥克風",        "mic"),
            ("系統 + 麥克風", "both"),
        ]
        self._mode_radios = []
        for text, value in modes:
            rb = ttk.Radiobutton(
                frame_mode, text=text,
                variable=self.record_mode, value=value,
            )
            rb.pack(side="left", padx=(0, 20))
            self._mode_radios.append(rb)

        ttk.Button(
            frame_mode, text=" ? ", width=3,
            command=self._show_mode_help,
        ).pack(side="right")

        # 輸出格式：非「系統 + 麥克風」模式時 gray out
        self._frame_output_inline = tk.Frame(frame_mode)
        self._frame_output_inline.pack(fill="x", pady=(8, 0))
        ttk.Label(
            self._frame_output_inline, text="輸出格式：",
            font=("", 9), foreground="#555555",
        ).pack(anchor="w")
        self._output_radios = []
        for text, val in [("合併一軌", "merge"), ("獨立兩軌", "separate"), ("兩個都要", "both")]:
            rb = ttk.Radiobutton(
                self._frame_output_inline, text=text,
                variable=self.output_mode, value=val,
            )
            rb.pack(anchor="w", padx=(12, 0))
            self._output_radios.append(rb)

        def _on_mode_change(*_):
            state = "normal" if self.record_mode.get() == "both" else "disabled"
            for rb in self._output_radios:
                rb.config(state=state)

        self.record_mode.trace_add("write", _on_mode_change)
        _on_mode_change()  # 設定初始 disabled 狀態

        # row=2  檔案名稱
        frame_name = ttk.LabelFrame(self.root, text=" 檔案名稱 ", padding=8)
        frame_name.grid(row=2, column=0, sticky="ew", **pad)
        frame_name.columnconfigure(0, weight=1)

        self.filename_entry = ttk.Entry(
            frame_name, textvariable=self.filename_var, width=44, font=("", 11)
        )
        self.filename_entry.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            frame_name, text=".mp3", foreground="gray"
        ).grid(row=0, column=1, padx=(4, 0))
        ttk.Label(
            frame_name, text="存檔前填好名稱，不填則自動用時間戳記命名",
            foreground="gray", font=("", 8)
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # row=3  錄音按鈕區
        frame_btn = tk.Frame(self.root)
        frame_btn.grid(row=3, column=0, pady=20)

        self.btn_record = ttk.Button(
            frame_btn, text="⏺  開始錄音",
            command=self._toggle_record, width=22
        )
        self.btn_record.pack(ipady=8)

        # 第二列按鈕：錄音中才顯示
        self._secondary_row = tk.Frame(frame_btn)

        self.btn_pause = ttk.Button(
            self._secondary_row, text="⏸  暫停",
            command=self._pause_recording, width=12
        )
        self.btn_pause.pack(side="left", ipady=4)

        self.btn_discard = ttk.Button(
            self._secondary_row, text="🗑  停止不儲存",
            command=self._discard_recording, width=14
        )
        self.btn_discard.pack(side="left", ipady=4, padx=(8, 0))

        self.timer_label = ttk.Label(
            frame_btn, text="00:00",
            font=("Consolas", 28, "bold"), foreground="gray"
        )
        self.timer_label.pack(pady=(10, 0))

        self.status_label = ttk.Label(
            frame_btn, text="等待開始錄音...", foreground="gray"
        )
        self.status_label.pack(pady=(4, 0))

        # VU Meter
        vu_frame = tk.Frame(frame_btn)
        vu_frame.pack(pady=(12, 0))

        ttk.Style().configure("Clip.Horizontal.TProgressbar", background="red")

        tk.Label(vu_frame, text="系統音訊", width=8, anchor="w",
                 font=("", 9)).grid(row=0, column=0, padx=(0, 6))
        self.vu_system_pb = ttk.Progressbar(vu_frame, variable=self.vu_system_var,
                        maximum=100, length=200)
        self.vu_system_pb.grid(row=0, column=1)
        self.vu_system_pct = ttk.Label(vu_frame, text="  0%", width=5,
                                        foreground="gray", font=("Consolas", 9))
        self.vu_system_pct.grid(row=0, column=2, padx=(4, 0))

        tk.Label(vu_frame, text="麥克風", width=8, anchor="w",
                 font=("", 9)).grid(row=1, column=0, padx=(0, 6), pady=(4, 0))
        self.vu_mic_pb = ttk.Progressbar(vu_frame, variable=self.vu_mic_var,
                        maximum=100, length=200)
        self.vu_mic_pb.grid(row=1, column=1, pady=(4, 0))
        self.vu_mic_pct = ttk.Label(vu_frame, text="  0%", width=5,
                                     foreground="gray", font=("Consolas", 9))
        self.vu_mic_pct.grid(row=1, column=2, padx=(4, 0), pady=(4, 0))
        self.mic_offline_label = ttk.Label(vu_frame, text="", foreground="red", font=("", 9))
        self.mic_offline_label.grid(row=1, column=3, padx=(6, 0), pady=(4, 0))

        self._sys_clip_until  = 0.0
        self._mic_clip_until  = 0.0
        self._sys_clip_active = False
        self._mic_clip_active = False

        # 裝置設定 + 進階設定 雙按鈕列
        btn_row = tk.Frame(frame_btn)
        btn_row.pack(pady=(12, 0))
        ttk.Button(btn_row, text="🔧 裝置設定與測試",
                   command=self._show_device_test).pack(side="left", padx=(0, 8))
        ttk.Button(btn_row, text="⚙ 進階設定",
                   command=self._show_advanced_settings).pack(side="left")

        # row=4  靜音警告橫幅（預設隱藏，偵測到連續靜音才顯示）
        self.silence_banner = tk.Frame(self.root, background="#FFA500", padx=12, pady=8)
        tk.Label(
            self.silence_banner,
            text="⚠  偵測到超過 10 秒沒有聲音，請確認：\n"
                 "系統是否靜音？播放裝置是否正確？",
            background="#FFA500", foreground="white",
            font=("", 10, "bold"), justify="left"
        ).pack(anchor="w")

        # row=5  錄音記錄
        frame_log = ttk.LabelFrame(self.root, text=" 錄音記錄 ", padding=8)
        frame_log.grid(row=5, column=0, sticky="ew", padx=14, pady=(0, 14))

        self.log_text = scrolledtext.ScrolledText(
            frame_log, width=52, height=6,
            state="disabled", font=("Consolas", 9)
        )
        self.log_text.pack(fill="x")

        self.root.columnconfigure(0, weight=1)
        self._log("請確認儲存位置與錄音模式，然後按「開始錄音」。")

    # ---- UI 互動 ----
    def _show_mode_help(self):
        """錄音模式說明彈窗"""
        win = tk.Toplevel(self.root)
        win.title("錄音模式說明")
        win.resizable(False, False)
        win.grab_set()  # modal，關閉前不能操作主視窗

        # 取樣率顯示：錄音後為實際偵測值，錄音前為預設值
        sys_rate_text  = f"{self.record_sample_rate} Hz  /  {'立體聲' if self.record_channels == 2 else '單聲道'}"
        mic_rate_text  = f"{self.record_mic_rate} Hz  /  單聲道"

        modes_info = [
            (
                "🖥  電腦聲音",
                "錄製所有從電腦播放的聲音。\n使用 WASAPI Loopback 技術，靜音狀態下依音效卡而定仍可錄音。",
                "Teams、Zoom、YouTube、任何會議軟體",
                sys_rate_text,
            ),
            (
                "🎙  麥克風",
                "只錄你說話的聲音，不含電腦播放的內容。",
                "只需要記錄自己發言的場合",
                mic_rate_text,
            ),
            (
                "🔀  系統 + 麥克風",
                "同時錄製電腦聲音與麥克風。\n可選擇合併一軌、獨立兩軌或兩個都存（在主畫面「輸出格式」選擇）。\n注意：若兩者取樣率不同，麥克風聲音速度可能略有偏差。",
                "想同時保留會議音訊與自己的旁白",
                f"系統 {sys_rate_text}  ／  麥克風 {mic_rate_text}",
            ),
        ]

        for i, (title, body, use_case, rate) in enumerate(modes_info):
            lf = ttk.LabelFrame(win, text=f"  {title}  ", padding=10)
            lf.grid(row=i, column=0, sticky="ew", padx=16, pady=(12 if i == 0 else 4, 4))

            # 適用場合：藍色粗體，讓使用者一眼找到選擇依據
            tk.Label(
                lf, text=f"✦ 適用：{use_case}",
                foreground="#0078D4", font=("", 10, "bold"),
                justify="left",
            ).grid(row=0, column=0, sticky="w")

            ttk.Label(
                lf, text=body, wraplength=320, justify="left",
                foreground="#444444",
            ).grid(row=1, column=0, sticky="w", pady=(4, 0))

            ttk.Label(
                lf, text=f"取樣率：{rate}",
                foreground="gray", font=("", 8)
            ).grid(row=2, column=0, sticky="w", pady=(6, 0))

        ttk.Label(
            win,
            text="* 取樣率於首次錄音後更新為實際裝置數值",
            foreground="gray", font=("", 8)
        ).grid(row=len(modes_info), column=0, padx=16, sticky="w")

        ttk.Button(win, text="關閉", command=win.destroy).grid(
            row=len(modes_info) + 1, column=0, pady=12
        )
        win.columnconfigure(0, weight=1)

    def _show_device_test(self):
        """裝置設定與測試對話框：錄音前確認麥克風與系統音訊是否有訊號"""
        win = tk.Toplevel(self.root)
        win.title("裝置設定與測試")
        win.resizable(False, False)
        win.grab_set()
        pad = {"padx": 14, "pady": 6}

        # --- 列舉裝置（僅 WASAPI，與 Windows 設定顯示一致）---
        # PortAudio 會對同一個實體裝置透過 MME / DirectSound / WASAPI 各列一次，
        # 造成下拉選單出現大量重複項目。過濾只保留 WASAPI host API 的裝置即可。
        p_enum = pyaudio.PyAudio()
        input_devices  = [("系統預設", None)]   # (顯示名稱, device_index)
        output_devices = [("系統預設", None)]   # (顯示名稱, device_name_for_loopback)
        try:
            wasapi_idx = p_enum.get_host_api_info_by_type(pyaudio.paWASAPI)["index"]
        except Exception:
            wasapi_idx = None  # 取不到時不過濾，保留全部
        for i in range(p_enum.get_device_count()):
            try:
                info = p_enum.get_device_info_by_index(i)
                if wasapi_idx is not None and info["hostApi"] != wasapi_idx:
                    continue  # 略過非 WASAPI 的裝置
                if info["maxInputChannels"] > 0 and not info.get("isLoopbackDevice", False):
                    input_devices.append((info["name"], i))
                if info["maxOutputChannels"] > 0 and not info.get("isLoopbackDevice", False):
                    output_devices.append((info["name"], info["name"]))
            except Exception:
                pass
        p_enum.terminate()

        # 測試狀態旗標：用單元素 list 包裝，讓巢狀 closure 可以修改其值
        # （Python 3 closure 可讀取外層變數，但無法直接對其重新賦值；list 可繞過此限制）
        mic_running  = [False]
        sys_running  = [False]

        # ===================== 麥克風區塊 =====================
        frame_mic = ttk.LabelFrame(win, text=" 🎙  輸入裝置（麥克風） ", padding=10)
        frame_mic.grid(row=0, column=0, sticky="ew", **pad)
        frame_mic.columnconfigure(0, weight=1)

        in_var = tk.StringVar()
        in_var.set(next((d[0] for d in input_devices if d[1] == self.selected_input_idx),
                        "系統預設"))
        ttk.Combobox(frame_mic, textvariable=in_var,
                     values=[d[0] for d in input_devices],
                     state="readonly", width=42).grid(row=0, column=0, columnspan=2,
                                                       sticky="ew", pady=(0, 8))

        mic_level = tk.DoubleVar(value=0)
        ttk.Progressbar(frame_mic, variable=mic_level,
                        maximum=100, length=340).grid(row=1, column=0, columnspan=2,
                                                       sticky="ew", pady=(0, 4))
        mic_status = ttk.Label(frame_mic, text="請對著麥克風說話，確認音量指示條有所反應",
                               foreground="gray")
        mic_status.grid(row=2, column=0, columnspan=2, sticky="w")
        btn_mic = ttk.Button(frame_mic, text="▶ 開始測試")
        btn_mic.grid(row=3, column=0, sticky="w", pady=(8, 0))
        btn_mic_stop = ttk.Button(frame_mic, text="■ 停止", state="disabled")
        btn_mic_stop.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(8, 0))

        def mic_worker(device_idx):
            """
            背景執行緒：持續讀取麥克風音訊並更新音量指示條。

            RMS 範圍 0~32767（Int16 最大值），除以 327.67 換算為 0~100 的百分比。
            mic_level.set() 直接從背景執行緒呼叫：tkinter DoubleVar 的 set 在 CPython
            下是執行緒安全的，Progressbar 會在下次主迴圈繪製時反映新值。

            finally 中的 win.after(0, ...) 用來將 UI 還原操作排回主執行緒執行，
            外層 try/except 防止視窗已關閉時 after() 拋出 TclError。

            使用 MME 開啟麥克風（同 _record_mic_worker），確保測試結果與實際錄音一致。
            """
            p = pyaudio.PyAudio()
            try:
                # 用 MME 繞過 Discord 等在 WASAPI 層套用的音訊增強（見 PITFALLS.md Pitfall 2）
                mme_idx = self._find_mme_mic_device(p, device_idx)
                dev_info = (p.get_device_info_by_index(mme_idx)
                            if mme_idx is not None
                            else p.get_default_input_device_info())
                sample_rate = int(dev_info["defaultSampleRate"])
                stream = p.open(format=pyaudio.paInt16, channels=1, rate=sample_rate,
                                frames_per_buffer=512, input=True,
                                input_device_index=mme_idx)
                while mic_running[0]:
                    data = stream.read(512, exception_on_overflow=False)
                    rms  = _compute_rms(data)
                    try:
                        mic_level.set(min(100, rms / 327.67))
                    except Exception:
                        break  # 視窗已關閉，DoubleVar 失效，結束迴圈
                stream.stop_stream()
                stream.close()
            except Exception as e:
                try:
                    win.after(0, lambda: mic_status.config(
                        text=f"錯誤：{e}", foreground="red"))
                except Exception:
                    pass
            finally:
                p.terminate()
                mic_running[0] = False
                try:
                    win.after(0, lambda: (
                        btn_mic.config(state="normal"),
                        btn_mic_stop.config(state="disabled"),
                        mic_level.set(0),
                    ))
                except Exception:
                    pass

        def start_mic_test():
            if mic_running[0]:
                return
            idx = next((d[1] for d in input_devices if d[0] == in_var.get()), None)
            mic_running[0] = True
            mic_status.config(text="測試中，請對著麥克風說話，確認音量指示條有所反應", foreground="#0078D4")
            btn_mic.config(state="disabled")
            btn_mic_stop.config(state="normal")
            threading.Thread(target=mic_worker, args=(idx,), daemon=True).start()

        def stop_mic_test():
            mic_running[0] = False
            mic_status.config(text="已停止", foreground="gray")

        btn_mic.config(command=start_mic_test)
        btn_mic_stop.config(command=stop_mic_test)

        # ===================== 電腦聲音區塊 =====================
        frame_sys = ttk.LabelFrame(win, text=" 🖥  電腦聲音（WASAPI Loopback） ", padding=10)
        frame_sys.grid(row=1, column=0, sticky="ew", **pad)
        frame_sys.columnconfigure(0, weight=1)

        out_var = tk.StringVar()
        out_var.set(next((d[0] for d in output_devices if d[1] == self.selected_output_name),
                         "系統預設"))
        ttk.Combobox(frame_sys, textvariable=out_var,
                     values=[d[0] for d in output_devices],
                     state="readonly", width=42).grid(row=0, column=0, columnspan=2,
                                                       sticky="ew", pady=(0, 8))

        sys_level = tk.DoubleVar(value=0)
        ttk.Progressbar(frame_sys, variable=sys_level,
                        maximum=100, length=340).grid(row=1, column=0, columnspan=2,
                                                       sticky="ew", pady=(0, 4))
        sys_status = ttk.Label(frame_sys,
                               text="請先播放任意音訊（音樂、影片等），再按下開始測試",
                               foreground="gray")
        sys_status.grid(row=2, column=0, columnspan=2, sticky="w")
        btn_sys = ttk.Button(frame_sys, text="▶ 開始測試")
        btn_sys.grid(row=3, column=0, sticky="w", pady=(8, 0))
        btn_sys_stop = ttk.Button(frame_sys, text="■ 停止", state="disabled")
        btn_sys_stop.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(8, 0))

        def sys_worker(output_name):
            """
            背景執行緒：透過 WASAPI Loopback 持續讀取系統音訊並更新音量指示條。

            output_name 為「系統預設」時傳 None 給 get_loopback_device，
            由該函式自動比對系統預設輸出裝置對應的 loopback。
            其餘邏輯同 mic_worker（RMS 換算、執行緒安全、視窗關閉防護）。
            """
            p = pyaudio.PyAudio()
            try:
                device = get_loopback_device(p, output_name if output_name != "系統預設" else None)
                ch = min(device["maxInputChannels"] or 2, 2)
                sr = int(device["defaultSampleRate"])
                stream = p.open(format=pyaudio.paInt16, channels=ch, rate=sr,
                                frames_per_buffer=512, input=True,
                                input_device_index=device["index"])
                while sys_running[0]:
                    data = stream.read(512, exception_on_overflow=False)
                    rms  = _compute_rms(data)
                    try:
                        sys_level.set(min(100, rms / 327.67))
                    except Exception:
                        break  # 視窗已關閉，結束迴圈
                stream.stop_stream()
                stream.close()
            except Exception as e:
                try:
                    win.after(0, lambda: sys_status.config(
                        text=f"錯誤：{e}", foreground="red"))
                except Exception:
                    pass
            finally:
                p.terminate()
                sys_running[0] = False
                try:
                    win.after(0, lambda: (
                        btn_sys.config(state="normal"),
                        btn_sys_stop.config(state="disabled"),
                        sys_level.set(0),
                    ))
                except Exception:
                    pass

        def start_sys_test():
            if sys_running[0]:
                return
            out_name = out_var.get()
            sys_running[0] = True
            sys_status.config(
                text="測試中，音量指示條有反應表示系統音訊可正常錄製",
                foreground="#0078D4")
            btn_sys.config(state="disabled")
            btn_sys_stop.config(state="normal")
            threading.Thread(target=sys_worker, args=(out_name,), daemon=True).start()

        def stop_sys_test():
            sys_running[0] = False
            sys_status.config(text="已停止", foreground="gray")

        btn_sys.config(command=start_sys_test)
        btn_sys_stop.config(command=stop_sys_test)

        # ===================== 確認 / 取消 =====================
        frame_btns = tk.Frame(win)
        frame_btns.grid(row=2, column=0, pady=12)

        def confirm():
            mic_running[0] = False
            sys_running[0] = False
            sel_in  = in_var.get()
            sel_out = out_var.get()
            self.selected_input_idx   = next((d[1] for d in input_devices  if d[0] == sel_in),  None)
            self.selected_output_name = next((d[1] for d in output_devices if d[0] == sel_out), None)
            win.destroy()

        def on_close():
            mic_running[0] = False
            sys_running[0] = False
            win.destroy()

        ttk.Button(frame_btns, text="確認選擇", command=confirm).pack(side="left", padx=8)
        ttk.Button(frame_btns, text="取消",     command=on_close).pack(side="left")
        win.protocol("WM_DELETE_WINDOW", on_close)
        win.columnconfigure(0, weight=1)

    def _show_advanced_settings(self):
        win = tk.Toplevel(self.root)
        win.title("進階設定")
        win.resizable(False, False)
        win.grab_set()
        pad = {"padx": 14, "pady": 6}

        # 本地變數：pre-populate 目前設定，Cancel 不影響 app 狀態
        br_var     = tk.IntVar(value=self.bit_rate.get())
        eq_var     = tk.BooleanVar(value=self.equalize_enabled.get())
        cap_var    = tk.IntVar(value=self.eq_gain_cap.get())
        filter_var = tk.BooleanVar(value=self.eq_filter_silence.get())

        # ---- 音質 ----
        frame_quality = ttk.LabelFrame(win, text=" 音質 ", padding=10)
        frame_quality.grid(row=0, column=0, sticky="ew", **pad)
        for label, val in [("128 kbps（標準，一般會議適用）", 128),
                            ("192 kbps（較好音質）",           192),
                            ("320 kbps（最高，後製 / Podcast）", 320)]:
            ttk.Radiobutton(frame_quality, text=label,
                            variable=br_var, value=val).pack(anchor="w")

        # ---- 混音設定 ----
        frame_mix = ttk.LabelFrame(
            win, text=" 混音設定（僅「系統＋麥克風」有效） ", padding=10)
        frame_mix.grid(row=1, column=0, sticky="ew", **pad)

        ttk.Checkbutton(
            frame_mix,
            text="自動等化音量（讓兩軌響度接近）",
            variable=eq_var,
        ).pack(anchor="w")
        ttk.Label(
            frame_mix,
            text="僅影響合併輸出；獨立音軌保留原始音量",
            foreground="gray", font=("", 8),
        ).pack(anchor="w", pady=(2, 8))

        # 等化子設定（eq 關閉時 gray out）
        frame_eq_sub = tk.Frame(frame_mix)
        frame_eq_sub.pack(fill="x")
        frame_eq_sub.columnconfigure(2, weight=1)

        lbl_cap = ttk.Label(frame_eq_sub, text="增益上限：")
        lbl_cap.grid(row=0, column=0, sticky="w")
        spn_cap = ttk.Spinbox(frame_eq_sub, from_=1, to=16,
                              textvariable=cap_var, width=5)
        spn_cap.grid(row=0, column=1, sticky="w", padx=(4, 0))
        lbl_x = ttk.Label(frame_eq_sub, text="x  (1～16)", foreground="gray")
        lbl_x.grid(row=0, column=2, sticky="w", padx=(4, 0))
        chk_filter = ttk.Checkbutton(
            frame_eq_sub,
            text="靜音過濾：排除靜音段再計算 RMS",
            variable=filter_var,
        )
        chk_filter.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))

        _eq_sub_widgets = [lbl_cap, spn_cap, lbl_x, chk_filter]

        def _update_eq_state(*_):
            state = "normal" if eq_var.get() else "disabled"
            for w in _eq_sub_widgets:
                w.config(state=state)

        eq_var.trace_add("write", _update_eq_state)
        _update_eq_state()  # 設定初始 gray-out 狀態

        # ---- 確認 / 取消 ----
        frame_btns = tk.Frame(win)
        frame_btns.grid(row=2, column=0, pady=12)

        def confirm():
            self.bit_rate.set(br_var.get())
            self.equalize_enabled.set(eq_var.get())
            self.eq_gain_cap.set(max(1, min(16, cap_var.get())))
            self.eq_filter_silence.set(filter_var.get())
            win.destroy()

        ttk.Button(frame_btns, text="確認", command=confirm).pack(side="left", padx=8)
        ttk.Button(frame_btns, text="取消", command=win.destroy).pack(side="left")
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.columnconfigure(0, weight=1)

    def _change_folder(self):
        folder = filedialog.askdirectory(
            title="選擇錄音檔儲存位置",
            initialdir=self.save_folder,
            parent=self.root,
        )
        if folder:
            self.save_folder = folder
            self.save_folder_var.set(folder)

    def _toggle_record(self):
        if self.is_recording or self.is_paused:
            self._stop_recording()
        else:
            self._start_recording()

    def _pause_recording(self):
        self.is_recording = False
        self.is_paused = True
        self._elapsed_before_pause += time.time() - self.start_time

        self.btn_record.config(state="disabled")
        self.btn_pause.config(state="disabled")
        self.btn_discard.config(state="disabled")
        self.status_label.config(text="暫停中...", foreground="gray")
        self.timer_label.config(foreground="#FF8C00")

        threading.Thread(target=self._wait_for_pause, daemon=True).start()

    def _wait_for_pause(self):
        """背景：等錄音執行緒結束，釋放 PyAudio，再通知 UI 進入已暫停狀態"""
        self._force_stop_streams()
        if self._record_thread:
            self._record_thread.join(timeout=5)
        if self._mic_thread:
            self._mic_thread.join(timeout=5)
        if not (self._record_thread and self._record_thread.is_alive()) and \
           not (self._mic_thread    and self._mic_thread.is_alive()):
            try:
                self._pa.terminate()
            except Exception:
                pass
        self.msg_queue.put(("paused", None))

    def _resume_recording(self):
        self.is_paused = False
        self.is_recording = True
        self.start_time = time.time()
        mode = self.record_mode.get()

        self._pa = pyaudio.PyAudio()

        self.btn_record.config(text="⏹  停止並儲存", state="normal")
        self.btn_pause.config(text="⏸  暫停", command=self._pause_recording, state="normal")
        self.btn_discard.config(state="normal")
        self.status_label.config(text="錄音中...", foreground="red")
        self.timer_label.config(foreground="red")

        self._update_timer()

        if mode in ("system", "both"):
            self._record_thread = threading.Thread(
                target=self._record_worker, args=(self._pa,), daemon=True)
            self._record_thread.start()

        if mode in ("mic", "both"):
            self._mic_thread = threading.Thread(
                target=self._record_mic_worker,
                args=(self._pa, mode == "mic"),
                daemon=True,
            )
            self._mic_thread.start()

    def _discard_recording(self):
        confirmed = messagebox.askyesno(
            "確認停止不儲存",
            "確定要停止錄音並捨棄所有資料？\n此操作無法復原。",
            icon="warning",
            parent=self.root,
        )
        if not confirmed:
            return
        self.is_recording = False
        self.is_paused = False
        self.btn_record.config(state="disabled", text="停止中...")
        self.btn_pause.config(state="disabled")
        self.btn_discard.config(state="disabled")
        self.status_label.config(text="停止中...", foreground="gray")
        self.timer_label.config(foreground="gray")
        self.filename_entry.config(state="normal")
        threading.Thread(target=self._cleanup_after_discard, daemon=True).start()

    def _cleanup_after_discard(self):
        self._force_stop_streams()
        if self._record_thread:
            self._record_thread.join(timeout=5)
        if self._mic_thread:
            self._mic_thread.join(timeout=5)
        if not (self._record_thread and self._record_thread.is_alive()) and \
           not (self._mic_thread    and self._mic_thread.is_alive()):
            try:
                self._pa.terminate()
            except Exception:
                pass
        self.record_frames = []
        self.mic_frames = []
        self.msg_queue.put(("discarded", None))

    def _set_mode_radios_state(self, state: str):
        for rb in self._mode_radios:
            rb.config(state=state)

    def _start_recording(self):
        self.is_recording  = True
        self.is_paused     = False
        self.record_frames = []
        self.mic_frames    = []
        self.start_time    = time.time()
        self._elapsed_before_pause = 0.0
        mode = self.record_mode.get()
        self._save_mode = mode  # 供背景執行緒在錄音途中判斷模式（stop 前 _save_mode 才鎖定完整快照）

        # 共用一個 PyAudio 實例，避免兩個執行緒各自 Pa_Initialize() 造成 C 層 assert crash
        self._pa = pyaudio.PyAudio()

        self.btn_record.config(text="⏹  停止並儲存")
        self.btn_pause.config(text="⏸  暫停", command=self._pause_recording, state="normal")
        self.btn_discard.config(state="normal")
        self._secondary_row.pack(pady=(8, 0))
        self.status_label.config(text="錄音中...", foreground="red")
        self.timer_label.config(foreground="red")
        self.filename_entry.config(state="disabled")
        self._set_mode_radios_state("disabled")  # 錄音中不允許切換模式

        self._update_timer()

        if mode in ("system", "both"):
            self._record_thread = threading.Thread(
                target=self._record_worker, args=(self._pa,), daemon=True)
            self._record_thread.start()

        if mode in ("mic", "both"):
            # check_silence：Mode "mic" 才由麥克風 worker 負責靜音偵測；
            # Mode "both" 的靜音偵測由 loopback worker 負責
            self._mic_thread = threading.Thread(
                target=self._record_mic_worker,
                args=(self._pa, mode == "mic"),
                daemon=True,
            )
            self._mic_thread.start()

    def _update_timer(self):
        """每秒更新計時器（root.after 確保在主執行緒執行）"""
        if self.is_recording:
            elapsed = int(self._elapsed_before_pause + time.time() - self.start_time)
            mins = elapsed // 60
            secs = elapsed % 60
            self.timer_label.config(text=f"{mins:02d}:{secs:02d}")
            self.root.after(1000, self._update_timer)

    def _stop_recording(self):
        self.is_recording = False

        # 在主執行緒鎖定模式，避免背景 _save_after_stop 從 tkinter StringVar 讀取
        self._save_mode = self.record_mode.get()
        self._save_output_mode    = self.output_mode.get()
        self._save_equalize       = self.equalize_enabled.get()
        self._save_gain_cap       = float(self.eq_gain_cap.get())
        self._save_filter_silence = self.eq_filter_silence.get()
        self._save_bit_rate       = self.bit_rate.get()

        self.btn_record.config(state="disabled", text="儲存中...")
        self.status_label.config(text="轉換為 MP3 中...", foreground="gray")
        self.timer_label.config(foreground="gray")
        self.filename_entry.config(state="normal")

        t = threading.Thread(target=self._save_after_stop, daemon=True)
        t.start()

    # ---- 錄音執行緒：Loopback ----
    def _record_worker(self, p: pyaudio.PyAudio):
        """
        錄製系統音訊（WASAPI Loopback）並做靜音偵測。

        p：由 _start_recording 建立的共用 PyAudio 實例，
           不在此處 terminate（由 _save_after_stop 統一管理）。
        open_stream 設計為 closure 以便在裝置切換後重新開啟。
        """
        try:
            chunk = 512

            def open_stream(channels=None, sample_rate=None):
                """取得當前預設輸出的 loopback stream，沿用指定格式確保 PCM 連續性"""
                device = get_loopback_device(p, self.selected_output_name)
                ch = channels or min(device["maxInputChannels"] or 2, 2)
                sr = sample_rate or int(device["defaultSampleRate"])
                s = p.open(
                    format=pyaudio.paInt16,
                    channels=ch,
                    rate=sr,
                    frames_per_buffer=chunk,
                    input=True,
                    input_device_index=device["index"],
                )
                return s, ch, sr

            stream, self.record_channels, self.record_sample_rate = open_stream()
            self._record_stream = stream

            # 靜音閾值：Int16 RMS < 100（範圍 0~32767）
            # 對應約 0.3% 最大音量，足以區分真實靜音與極低背景雜訊
            SILENCE_RMS_THRESHOLD = _SILENCE_RMS_THRESHOLD
            SILENCE_WARNING_SECS  = _SILENCE_WARNING_SECS
            silence_start  = None
            silence_warned = False

            while self.is_recording:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    self.record_frames.append(data)

                    # ---- 靜音偵測 ----
                    rms = _compute_rms(data)
                    self.msg_queue.put(("vu_system", min(100.0, rms / 327.67)))
                    if rms < SILENCE_RMS_THRESHOLD:
                        if silence_start is None:
                            silence_start = time.time()
                        elif not silence_warned and (time.time() - silence_start) >= SILENCE_WARNING_SECS:
                            self.msg_queue.put(("silence_warning", True))
                            silence_warned = True
                    else:
                        silence_start = None
                        if silence_warned:
                            self.msg_queue.put(("silence_warning", False))
                            silence_warned = False

                except OSError:
                    # 插拔耳機 / 切換播放裝置 / 睡眠喚醒後音訊 session 重置導致 stream 失效
                    try:
                        stream.stop_stream()
                        stream.close()
                    except Exception:
                        pass
                    self._record_stream = None
                    self.msg_queue.put(("vu_system", 0.0))
                    self.msg_queue.put(("warning", "系統音訊裝置中斷，嘗試重新連線（最多等 30 秒）..."))

                    recovered = False
                    for _ in range(30):
                        time.sleep(1)
                        if not self.is_recording:
                            break
                        try:
                            stream, _, _ = open_stream(
                                channels=self.record_channels,
                                sample_rate=self.record_sample_rate,
                            )
                            self._record_stream = stream
                            recovered = True
                            self.msg_queue.put(("warning", "系統音訊裝置已重新連線，錄音繼續"))
                            break
                        except Exception:
                            continue

                    if not recovered:
                        if self.is_recording:
                            self.msg_queue.put(("device_failed",
                                "系統音訊裝置無法恢復（逾時 30 秒），自動儲存已錄部分。"))
                        break

            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
            self._record_stream = None

        except Exception as e:
            self.msg_queue.put(("error", str(e)))

    # ---- MME 裝置查找 ----
    def _find_mme_mic_device(self, p: pyaudio.PyAudio,
                             wasapi_idx: int | None = None) -> int | None:
        """
        回傳 MME host API 下對應的麥克風裝置 index。

        使用 MME 而非 WASAPI 開啟麥克風，可繞過 Windows 在 WASAPI 層套用的
        音訊增強（AGC、降噪、回音消除）。Discord 等通話軟體會在 WASAPI 層
        啟用這些增強，導致同時錄音時訊號失真。MME 直接存取原始硬體音訊，
        不受影響。

        wasapi_idx：要比對的 WASAPI 裝置 index；省略時使用 self.selected_input_idx。
        比對邏輯：MME 裝置名稱通常是 WASAPI 名稱截斷至前 31 個字元，
        以 WASAPI 名稱開頭比對 MME 名稱。
        找不到對應裝置時 fallback 到 MME 預設輸入。
        """
        if wasapi_idx is None:
            wasapi_idx = self.selected_input_idx

        try:
            mme_info    = p.get_host_api_info_by_type(pyaudio.paMME)
            mme_api_idx = mme_info["index"]
        except Exception:
            return wasapi_idx  # MME 不可用，沿用原裝置

        if wasapi_idx is None:
            default_idx = mme_info.get("defaultInputDevice", -1)
            return default_idx if default_idx >= 0 else None

        try:
            wasapi_name = p.get_device_info_by_index(wasapi_idx)["name"]
        except Exception:
            default_idx = mme_info.get("defaultInputDevice", -1)
            return default_idx if default_idx >= 0 else None

        for i in range(p.get_device_count()):
            try:
                dev = p.get_device_info_by_index(i)
                if (dev["hostApi"] == mme_api_idx
                        and dev["maxInputChannels"] > 0
                        and wasapi_name.startswith(dev["name"])):
                    return i
            except Exception:
                continue

        # 找不到對應 MME 裝置，fallback 到 MME 預設輸入並警告使用者
        self.msg_queue.put(("warning",
            f"找不到麥克風「{wasapi_name}」對應的 MME 裝置，"
            f"改用 MME 預設輸入。實際錄音裝置可能與所選不符，請至裝置設定確認。"))
        default_idx = mme_info.get("defaultInputDevice", -1)
        return default_idx if default_idx >= 0 else None

    # ---- 錄音執行緒：麥克風 ----
    def _record_mic_worker(self, p: pyaudio.PyAudio, check_silence: bool = False):
        """
        錄製麥克風音訊。

        p：由 _start_recording 建立的共用 PyAudio 實例，不在此處 terminate。
        check_silence=True：Mode mic 時由本 worker 負責靜音偵測。
        Mode both 時為 False，靜音偵測交由 loopback worker 處理。

        取樣率盡量對齊 self.record_sample_rate（loopback 的 rate），
        方便後續混音。若麥克風不支援，fallback 到麥克風原生 rate 並記錄，
        混音時會顯示警告（兩個 rate 不一致會造成輕微音速偏差）。

        裝置選擇使用 MME 而非 WASAPI，見 _find_mme_mic_device。
        """
        try:
            chunk      = 512
            target_rate = self.record_sample_rate  # 對齊 loopback rate
            mme_idx    = self._find_mme_mic_device(p)

            try:
                stream = p.open(
                    format=pyaudio.paInt16,
                    channels=1,
                    rate=target_rate,
                    frames_per_buffer=chunk,
                    input=True,
                    input_device_index=mme_idx,
                )
                self.record_mic_channels = 1
                self.record_mic_rate     = target_rate
            except Exception:
                # 麥克風不支援目標 rate，退回麥克風原生 rate
                dev_info = (p.get_device_info_by_index(mme_idx)
                            if mme_idx is not None
                            else p.get_default_input_device_info())
                fallback = int(dev_info["defaultSampleRate"])
                stream = p.open(
                    format=pyaudio.paInt16,
                    channels=1,
                    rate=fallback,
                    frames_per_buffer=chunk,
                    input=True,
                    input_device_index=mme_idx,
                )
                self.record_mic_channels = 1
                self.record_mic_rate     = fallback

            self._mic_stream = stream

            # 靜音閾值說明同 _record_worker
            SILENCE_RMS_THRESHOLD = _SILENCE_RMS_THRESHOLD
            SILENCE_WARNING_SECS  = _SILENCE_WARNING_SECS
            silence_start  = None
            silence_warned = False

            while self.is_recording:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    self.mic_frames.append(data)

                    rms = _compute_rms(data)
                    self.msg_queue.put(("vu_mic", min(100.0, rms / 327.67)))

                    if check_silence:
                        if rms < SILENCE_RMS_THRESHOLD:
                            if silence_start is None:
                                silence_start = time.time()
                            elif not silence_warned and (time.time() - silence_start) >= SILENCE_WARNING_SECS:
                                self.msg_queue.put(("silence_warning", True))
                                silence_warned = True
                        else:
                            silence_start = None
                            if silence_warned:
                                self.msg_queue.put(("silence_warning", False))
                                silence_warned = False

                except OSError:
                    try:
                        stream.stop_stream()
                        stream.close()
                    except Exception:
                        pass
                    self._mic_stream = None
                    self.msg_queue.put(("vu_mic", 0.0))
                    self.msg_queue.put(("mic_offline", True))
                    self.msg_queue.put(("warning", "麥克風裝置中斷，嘗試重新連線（最多等 30 秒）..."))

                    recovered = False
                    for _ in range(30):
                        time.sleep(1)
                        if not self.is_recording:
                            break
                        try:
                            new_stream = p.open(
                                format=pyaudio.paInt16,
                                channels=1,
                                rate=self.record_mic_rate,
                                frames_per_buffer=chunk,
                                input=True,
                                input_device_index=mme_idx,
                            )
                            stream = new_stream
                            self._mic_stream = stream
                            recovered = True
                            self.msg_queue.put(("mic_offline", False))
                            self.msg_queue.put(("warning", "麥克風裝置已重新連線，錄音繼續"))
                            break
                        except Exception:
                            continue

                    if not recovered:
                        if self.is_recording:
                            level = "warning" if self._save_mode == "both" else "error"
                            self.msg_queue.put((level,
                                "麥克風裝置無法恢復（逾時 30 秒），麥克風錄音已停止"))
                        break

            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
            self._mic_stream = None

        except Exception as e:
            # Mode "both"：麥克風失敗不中止程式，但通知使用者
            # Mode "mic"：視為致命錯誤
            level = "warning" if self._save_mode == "both" else "error"
            self.msg_queue.put((level, f"麥克風錯誤：{e}"))

    # ---- 混音 ----
    def _mix_pcm(self, sys_data: bytes, sys_ch: int,
                 mic_data: bytes, mic_ch: int) -> bytes:
        """
        混合 Loopback（通常 stereo）和麥克風（mono）的 PCM Int16。
        輸出聲道數與 sys_ch 相同。

        混音權重各 0.6：
          - 0.5 理論上不會爆音，但混出來音量偏小
          - 0.6 讓整體音量更接近原始，偶爾超出 Int16 範圍時由 clamp 截斷
          - > 0.7 爆音風險明顯增加

        已知限制：
          若 sys 和 mic 的 sample_rate 不同（見 record_sample_rate vs record_mic_rate），
          兩個 array 長度比例會不一致，truncate 後 mic 音訊速度會輕微偏差。
          正確做法是 resample（需 numpy 或 soxr），目前接受此限制。
        """
        sys_arr = array.array('h')
        sys_arr.frombytes(sys_data)

        mic_arr = array.array('h')
        mic_arr.frombytes(mic_data)

        # Mic upmix：mono → stereo（L/R 複製相同 sample）
        if mic_ch == 1 and sys_ch == 2:
            stereo = array.array('h')
            for s in mic_arr:
                stereo.append(s)
                stereo.append(s)
            mic_arr = stereo

        # 以長的那軌為準，短的補零（靜音），避免 mic 中途斷線時合併音訊被截短
        length = max(len(sys_arr), len(mic_arr))
        sys_len = len(sys_arr)
        mic_len = len(mic_arr)

        mixed = array.array('h', [
            max(-32768, min(32767, int(
                (sys_arr[i] if i < sys_len else 0) * _MIX_WEIGHT_SYSTEM +
                (mic_arr[i] if i < mic_len else 0) * _MIX_WEIGHT_MIC
            )))
            for i in range(length)
        ])
        return mixed.tobytes()

    @staticmethod
    def _compute_equalize_gain(
        sys_frames: list,
        mic_frames: list,
        filter_silence: bool = True,
        gain_cap: float = 4.0,
    ) -> tuple:
        """
        回傳 (sys_gain, mic_gain)。只有較小的那軌 gain > 1.0。
        filter_silence=True 時排除 RMS < 100 的 chunk，避免靜音段影響平均。
        任一軌全靜音時回傳 (1.0, 1.0) 不等化。
        """
        THRESHOLD = _SILENCE_RMS_THRESHOLD

        def active_rms(frames):
            values = [_compute_rms(c) for c in frames]
            if filter_silence:
                values = [v for v in values if v >= THRESHOLD]
            return (sum(values) / len(values)) if values else 0.0

        sys_rms = active_rms(sys_frames)
        mic_rms = active_rms(mic_frames)

        if sys_rms == 0 or mic_rms == 0:
            return 1.0, 1.0

        if sys_rms >= mic_rms:
            return 1.0, min(sys_rms / mic_rms, gain_cap)
        else:
            return min(mic_rms / sys_rms, gain_cap), 1.0

    @staticmethod
    def _apply_gain_to_pcm(data: bytes, gain: float) -> bytes:
        """對 Int16 PCM 套用增益倍數，結果 clamp 至 [-32768, 32767]。"""
        if gain == 1.0:
            return data
        arr = array.array('h')
        arr.frombytes(data)
        scaled = array.array('h', [
            max(-32768, min(32767, int(s * gain))) for s in arr
        ])
        return scaled.tobytes()

    # ---- 儲存輔助方法 ----
    def _encode_to_mp3(self, pcm_data: bytes, channels: int, sample_rate: int,
                       bit_rate: int = _DEFAULT_BIT_RATE,
                       progress_cb=None) -> bytes:
        encoder = lameenc.Encoder()
        encoder.set_bit_rate(bit_rate)
        encoder.set_in_sample_rate(sample_rate)
        encoder.set_channels(channels)
        encoder.set_quality(2)

        if progress_cb is None or len(pcm_data) == 0:
            return encoder.encode(pcm_data) + encoder.flush()

        # 切 20 段，對齊 frame 邊界（channels × 2 bytes）
        frame_size = channels * 2
        total = len(pcm_data)
        chunk_size = max((total // 20 // frame_size) * frame_size, frame_size)

        result = b""
        offset = 0
        while offset < total:
            chunk = pcm_data[offset:offset + chunk_size]
            result += encoder.encode(chunk)
            offset += len(chunk)
            progress_cb(min(100, int(offset / total * 100)))

        return result + encoder.flush()

    def _save_file(self, mp3_data: bytes, base_name: str, suffix: str = "") -> str:
        name = base_name + suffix
        filepath = os.path.join(self.save_folder, f"{name}.mp3")
        counter = 2
        while os.path.exists(filepath):
            filepath = os.path.join(self.save_folder, f"{name} ({counter}).mp3")
            counter += 1
        with open(filepath, "wb") as f:
            f.write(mp3_data)
        return filepath

    def _force_stop_streams(self):
        """
        從外部強制停止活躍的音訊串流。

        stream.read() 在 WASAPI 裝置停止派發 chunk 時會無限阻塞，
        導致 join(timeout) 超時後 pa.terminate() 在 thread 仍活著時被呼叫，
        引發 C 層 crash（閃退）。呼叫 stop_stream() 可喚醒阻塞中的 read()
        讓 thread 收到 OSError 後自行退出。
        """
        for attr in ("_record_stream", "_mic_stream"):
            s = getattr(self, attr, None)
            if s is not None:
                try:
                    s.stop_stream()
                except Exception:
                    pass

    # ---- 儲存執行緒 ----
    def _save_after_stop(self):
        """
        等待所有錄音執行緒結束後，轉換並儲存 MP3。
        先 force-stop streams 讓 read() 阻塞解除，確保 join 能在 timeout 內完成。

        Code path 依 mode × output_mode 分為五條：
          system                        → 單一 encode → 存檔
          mic                           → 單一 encode → 存檔
          both，mic 無資料               → fallback 到 system 模式，單一 encode → 存檔
          both，output_mode=separate    → encode sys + encode mic → 存兩檔
          both，output_mode=merge       → 混音 → encode merged → 存一檔
          both，output_mode=both        → encode sys + encode mic + 混音 → encode merged → 存三檔
        """
        self._force_stop_streams()
        if self._record_thread:
            self._record_thread.join(timeout=5)
        if self._mic_thread:
            self._mic_thread.join(timeout=5)

        # 只在 thread 確實結束後才 terminate，避免殘留 thread 持有 handle
        record_alive = self._record_thread and self._record_thread.is_alive()
        mic_alive    = self._mic_thread    and self._mic_thread.is_alive()
        if not record_alive and not mic_alive:
            try:
                self._pa.terminate()
            except Exception:
                pass
        else:
            self.msg_queue.put(("warning", "錄音執行緒未正常結束，PyAudio 資源暫時保留（重啟程式可釋放）"))

        try:
            mode = self._save_mode  # 已在主執行緒鎖定，不從 tkinter StringVar 讀取

            custom_name = self.filename_var.get().strip()
            base_name   = custom_name if custom_name else (
                "meeting_" + datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            )

            if mode == "system":
                if not self.record_frames:
                    self.msg_queue.put(("error", "沒有錄到任何音訊"))
                    return
                pcm_data    = b"".join(self.record_frames)
                channels    = self.record_channels
                sample_rate = self.record_sample_rate

            elif mode == "mic":
                if not self.mic_frames:
                    self.msg_queue.put(("error", "沒有錄到任何音訊（麥克風）"))
                    return
                pcm_data    = b"".join(self.mic_frames)
                channels    = self.record_mic_channels
                sample_rate = self.record_mic_rate

            else:  # "both"
                if not self.record_frames:
                    self.msg_queue.put(("error", "沒有錄到任何系統音訊"))
                    return

                if not self.mic_frames:
                    # 麥克風無資料，退回純系統音訊
                    self.msg_queue.put(("warning", "麥克風無資料，改以「電腦聲音」模式儲存"))
                    mp3_data = self._encode_to_mp3(
                        b"".join(self.record_frames),
                        self.record_channels, self.record_sample_rate,
                        bit_rate=self._save_bit_rate,
                        progress_cb=self._make_progress_cb("音訊", 1, 1))
                    self.msg_queue.put(("progress_done", "✓ 編碼完成"))
                    filepath = self._save_file(mp3_data, base_name)
                    self.msg_queue.put(("saved", [filepath]))
                    return

                output_mode = self._save_output_mode
                total_enc = (2 if output_mode in ("separate", "both") else 0) + \
                            (1 if output_mode in ("merge", "both") else 0)
                enc_idx = 0
                saved_paths = []
                # snapshot 複製，避免 encode 途中主執行緒的 discard 操作清空 list
                sys_frames_snap = list(self.record_frames)
                mic_frames_snap = list(self.mic_frames)

                # ---- 獨立音軌 ----
                if output_mode in ("separate", "both"):
                    enc_idx += 1
                    sys_mp3 = self._encode_to_mp3(
                        b"".join(sys_frames_snap),
                        self.record_channels, self.record_sample_rate,
                        bit_rate=self._save_bit_rate,
                        progress_cb=self._make_progress_cb("system", enc_idx, total_enc))
                    self.msg_queue.put(("progress_done", "✓ system 編碼完成"))
                    enc_idx += 1
                    mic_mp3 = self._encode_to_mp3(
                        b"".join(mic_frames_snap),
                        self.record_mic_channels, self.record_mic_rate,
                        bit_rate=self._save_bit_rate,
                        progress_cb=self._make_progress_cb("mic", enc_idx, total_enc))
                    self.msg_queue.put(("progress_done", "✓ mic 編碼完成"))
                    saved_paths.append(self._save_file(sys_mp3, base_name, "_system"))
                    saved_paths.append(self._save_file(mic_mp3, base_name, "_mic"))

                # ---- 合併音軌 ----
                if output_mode in ("merge", "both"):
                    if self.record_mic_rate != self.record_sample_rate:
                        self.msg_queue.put(("warning",
                            f"麥克風取樣率（{self.record_mic_rate} Hz）與系統音訊"
                            f"（{self.record_sample_rate} Hz）不一致，麥克風聲音可能略有偏差"))

                    self.msg_queue.put(("status", "混音中..."))
                    sys_pcm = b"".join(sys_frames_snap)
                    mic_pcm = b"".join(mic_frames_snap)

                    if self._save_equalize:
                        sys_gain, mic_gain = self._compute_equalize_gain(
                            sys_frames_snap, mic_frames_snap,
                            filter_silence=self._save_filter_silence,
                            gain_cap=self._save_gain_cap,
                        )
                        sys_pcm = self._apply_gain_to_pcm(sys_pcm, sys_gain)
                        mic_pcm = self._apply_gain_to_pcm(mic_pcm, mic_gain)

                    mixed_pcm = self._mix_pcm(
                        sys_pcm, self.record_channels,
                        mic_pcm, self.record_mic_channels)
                    enc_idx += 1
                    merged_mp3 = self._encode_to_mp3(
                        mixed_pcm, self.record_channels, self.record_sample_rate,
                        bit_rate=self._save_bit_rate,
                        progress_cb=self._make_progress_cb("合併音訊", enc_idx, total_enc))
                    self.msg_queue.put(("progress_done", "✓ 合併音訊編碼完成"))
                    saved_paths.append(self._save_file(merged_mp3, base_name))

                if not saved_paths:
                    self.msg_queue.put(("error", f"未知的輸出模式：{output_mode}"))
                    return
                self.msg_queue.put(("saved", saved_paths))
                return

            mp3_data = self._encode_to_mp3(pcm_data, channels, sample_rate,
                                           bit_rate=self._save_bit_rate,
                                           progress_cb=self._make_progress_cb("音訊", 1, 1))
            self.msg_queue.put(("progress_done", "✓ 編碼完成"))
            filepath = self._save_file(mp3_data, base_name)
            self.msg_queue.put(("saved", [filepath]))

        except Exception as e:
            self.msg_queue.put(("error", str(e)))

    # ---- 執行緒安全 UI 更新 ----
    def _log(self, msg: str):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _log_progress(self, text: str, done: bool = False):
        """覆寫 log 最後一行（進度更新），done=True 時清除旗標讓該行永久保留。"""
        self.log_text.config(state="normal")
        if self._progress_line_active:
            self.log_text.delete("end-1c linestart", "end-1c")
            self.log_text.insert("end-1c", text)
        else:
            self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self._progress_line_active = not done

    def _make_progress_cb(self, label: str, file_idx: int, total: int):
        """回傳一個 callback，每次呼叫時把進度推進 msg_queue。"""
        def cb(pct: int):
            self.msg_queue.put(("progress",
                f"⏳ 編碼 {label} ({file_idx}/{total}) {pct}%"))
        return cb

    def _reset_ui_after_stop(self):
        """錄音與儲存流程完全結束後，還原所有 UI 元件狀態"""
        self.is_paused = False
        self._elapsed_before_pause = 0.0
        self.btn_record.config(state="normal", text="⏺  開始錄音")
        self.btn_pause.config(text="⏸  暫停", command=self._pause_recording, state="normal")
        self.btn_discard.config(state="normal")
        self._secondary_row.pack_forget()
        self.timer_label.config(text="00:00", foreground="gray")
        self.silence_banner.grid_remove()
        self._set_mode_radios_state("normal")
        self.vu_system_var.set(0)
        self.vu_mic_var.set(0)
        self.vu_system_pct.config(text=f"{0:3d}%")
        self.vu_mic_pct.config(text=f"{0:3d}%")
        self.mic_offline_label.config(text="")

    def _poll_queue(self):
        """
        每 100ms 從 msg_queue 拉訊息更新 UI。
        所有 UI 操作都在此（主執行緒）執行，背景執行緒只放訊息進 queue。

        訊息類型：
          saved           — 儲存成功，data = list[filepath]
          error           — 致命錯誤，data = 錯誤訊息
          warning         — 非致命警告，data = 警告訊息（顯示在 log，不中止流程）
          status          — 狀態文字更新，data = 狀態字串
          silence_warning — 靜音偵測，data = True（顯示）/ False（隱藏）
          vu_system       — 系統音訊音量，data = float 0~100
          vu_mic          — 麥克風音量，data = float 0~100
          paused          — 暫停完成，data = None
          discarded       — 錄音已捨棄，data = None
          progress        — MP3 編碼進度，data = 字串（覆寫 log 最後一行）
          progress_done   — 編碼階段完成，data = 字串（覆寫最後一行後永久保留）
          device_failed   — 系統音訊裝置不可恢復，自動觸發儲存流程
          mic_offline     — 麥克風離線警示，data = True（顯示）/ False（清除）
        """
        try:
            while True:
                msg_type, data = self.msg_queue.get_nowait()

                if msg_type not in ("progress", "progress_done"):
                    self._progress_line_active = False

                if msg_type == "progress":
                    self._log_progress(data, done=False)
                    continue

                if msg_type == "progress_done":
                    self._log_progress(data, done=True)
                    continue

                if msg_type == "saved":
                    for fp in data:
                        self._log(f"✓  {os.path.basename(fp)}")
                    self.status_label.config(
                        text=f"已儲存：{os.path.basename(data[-1])}", foreground="green")
                    self._reset_ui_after_stop()

                elif msg_type == "error":
                    self._log(f"[ERROR] {data}")
                    self._reset_ui_after_stop()
                    self.status_label.config(text="發生錯誤，請查看記錄", foreground="red")
                    self.is_recording = False

                elif msg_type == "warning":
                    # 非致命：顯示在 log 但不中斷流程
                    self._log(f"[WARNING] {data}")

                elif msg_type == "vu_system":
                    val = max(0.0, min(100.0, float(data)))
                    self.vu_system_var.set(val)
                    self.vu_system_pct.config(text=f"{int(val):3d}%")

                elif msg_type == "vu_mic":
                    val = max(0.0, min(100.0, float(data)))
                    self.vu_mic_var.set(val)
                    self.vu_mic_pct.config(text=f"{int(val):3d}%")

                elif msg_type == "mic_offline":
                    self.mic_offline_label.config(text="⚠ 麥克風斷線" if data else "")

                elif msg_type == "paused":
                    self.mic_offline_label.config(text="")
                    self.btn_record.config(state="normal")
                    self.btn_pause.config(text="▶  繼續錄音",
                                          command=self._resume_recording, state="normal")
                    self.btn_discard.config(state="normal")
                    self.status_label.config(text="已暫停", foreground="#FF8C00")

                elif msg_type == "device_failed":
                    self._log(f"[WARNING] {data}")
                    if self.is_recording:  # 防止用戶已按停止時重複觸發
                        self._stop_recording()

                elif msg_type == "discarded":
                    self._log("✗  錄音已捨棄（未儲存）")
                    self._reset_ui_after_stop()
                    self.status_label.config(text="錄音已捨棄", foreground="gray")

                elif msg_type == "status":
                    self.status_label.config(text=data, foreground="gray")

                elif msg_type == "silence_warning":
                    if data:
                        self.silence_banner.grid(row=4, column=0, sticky="ew", padx=14, pady=(0, 6))
                    else:
                        self.silence_banner.grid_remove()

        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)


# ---- 入口 ----
def main():
    show_cth_banner()
    root = tk.Tk()
    MeetingRecorderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
