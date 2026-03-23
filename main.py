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
from tkinter import ttk, filedialog, scrolledtext

import pyaudiowpatch as pyaudio
import lameenc


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
        self.msg_queue: queue.Queue = queue.Queue()
        self._record_thread: threading.Thread | None = None
        self._mic_thread:    threading.Thread | None = None
        self._save_mode: str = "system"        # 儲存時使用的模式，在停止時鎖定避免 race condition

        # 裝置選擇（None = 系統預設）
        self.selected_input_idx:    int | None = None   # 麥克風裝置 index
        self.selected_output_name:  str | None = None   # 輸出裝置名稱（用於比對 loopback）

        # 儲存設定
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.save_folder = desktop
        self.save_folder_var = tk.StringVar(value=desktop)
        self.filename_var    = tk.StringVar()

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

        self.timer_label = ttk.Label(
            frame_btn, text="00:00",
            font=("Consolas", 28, "bold"), foreground="gray"
        )
        self.timer_label.pack(pady=(10, 0))

        self.status_label = ttk.Label(
            frame_btn, text="等待開始錄音...", foreground="gray"
        )
        self.status_label.pack(pady=(4, 0))

        ttk.Button(
            frame_btn, text="🔧 裝置設定與測試",
            command=self._show_device_test,
        ).pack(pady=(12, 0))

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
                "同時錄製電腦聲音與麥克風，存檔前混成一軌。\n注意：若兩者取樣率不同，麥克風聲音速度可能略有偏差。",
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

        # --- 列舉裝置 ---
        p_enum = pyaudio.PyAudio()
        input_devices  = [("系統預設", None)]   # (顯示名稱, device_index)
        output_devices = [("系統預設", None)]   # (顯示名稱, device_name_for_loopback)
        for i in range(p_enum.get_device_count()):
            try:
                info = p_enum.get_device_info_by_index(i)
                if info["maxInputChannels"] > 0 and not info.get("isLoopbackDevice", False):
                    input_devices.append((info["name"], i))
                if info["maxOutputChannels"] > 0 and not info.get("isLoopbackDevice", False):
                    output_devices.append((info["name"], info["name"]))
            except Exception:
                pass
        p_enum.terminate()

        # --- 測試狀態旗標（用 list 讓 closure 可修改）---
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

        def mic_poll():
            if mic_running[0]:
                try:
                    win.after(80, mic_poll)
                except Exception:
                    pass

        def mic_worker(device_idx):
            p = pyaudio.PyAudio()
            try:
                stream = p.open(format=pyaudio.paInt16, channels=1, rate=44100,
                                frames_per_buffer=512, input=True,
                                input_device_index=device_idx)
                while mic_running[0]:
                    data = stream.read(512, exception_on_overflow=False)
                    rms  = _compute_rms(data)
                    try:
                        mic_level.set(min(100, rms / 327.67))
                    except Exception:
                        break
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
                        break
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
        if not self.is_recording:
            self._start_recording()
        else:
            self._stop_recording()

    def _set_mode_radios_state(self, state: str):
        for rb in self._mode_radios:
            rb.config(state=state)

    def _start_recording(self):
        self.is_recording  = True
        self.record_frames = []
        self.mic_frames    = []
        self.start_time    = time.time()
        mode = self.record_mode.get()

        # 共用一個 PyAudio 實例，避免兩個執行緒各自 Pa_Initialize() 造成 C 層 assert crash
        self._pa = pyaudio.PyAudio()

        self.btn_record.config(text="⏹  停止並儲存")
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
            elapsed = int(time.time() - self.start_time)
            mins = elapsed // 60
            secs = elapsed % 60
            self.timer_label.config(text=f"{mins:02d}:{secs:02d}")
            self.root.after(1000, self._update_timer)

    def _stop_recording(self):
        self.is_recording = False

        # 在主執行緒鎖定模式，避免背景 _save_after_stop 從 tkinter StringVar 讀取
        self._save_mode = self.record_mode.get()

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

            # 靜音閾值：Int16 RMS < 100（範圍 0~32767）
            # 對應約 0.3% 最大音量，足以區分真實靜音與極低背景雜訊
            SILENCE_RMS_THRESHOLD = 100
            SILENCE_WARNING_SECS  = 10
            silence_start  = None
            silence_warned = False

            while self.is_recording:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    self.record_frames.append(data)

                    # ---- 靜音偵測 ----
                    rms = _compute_rms(data)
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
                    # 插拔耳機 / 切換播放裝置導致 stream 失效，重新取得新裝置
                    try:
                        stream.stop_stream()
                        stream.close()
                    except Exception:
                        pass
                    time.sleep(0.5)  # 等 Windows 完成裝置切換
                    if not self.is_recording:
                        break
                    try:
                        stream, _, _ = open_stream(
                            channels=self.record_channels,
                            sample_rate=self.record_sample_rate,
                        )
                    except Exception:
                        time.sleep(1)  # 裝置仍不可用，避免 busy-wait 佔滿 CPU
                        continue

            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass

        except Exception as e:
            self.msg_queue.put(("error", str(e)))

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
        """
        try:
            chunk = 512
            target_rate = self.record_sample_rate  # 對齊 loopback rate

            try:
                stream = p.open(
                    format=pyaudio.paInt16,
                    channels=1,
                    rate=target_rate,
                    frames_per_buffer=chunk,
                    input=True,
                    input_device_index=self.selected_input_idx,
                )
                self.record_mic_channels = 1
                self.record_mic_rate     = target_rate
            except Exception:
                # 麥克風不支援目標 rate，退回麥克風原生 rate
                dev_info = (p.get_device_info_by_index(self.selected_input_idx)
                            if self.selected_input_idx is not None
                            else p.get_default_input_device_info())
                fallback = int(dev_info["defaultSampleRate"])
                stream = p.open(
                    format=pyaudio.paInt16,
                    channels=1,
                    rate=fallback,
                    frames_per_buffer=chunk,
                    input=True,
                    input_device_index=self.selected_input_idx,
                )
                self.record_mic_channels = 1
                self.record_mic_rate     = fallback

            # 靜音閾值說明同 _record_worker
            SILENCE_RMS_THRESHOLD = 100
            SILENCE_WARNING_SECS  = 10
            silence_start  = None
            silence_warned = False

            while self.is_recording:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    self.mic_frames.append(data)

                    if check_silence:
                        rms = _compute_rms(data)
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
                    break  # 麥克風失效，結束錄音

            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass

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

        # 兩個執行緒可能有微小長度差異，截到短的那個
        length = min(len(sys_arr), len(mic_arr))

        mixed = array.array('h', [
            max(-32768, min(32767, int(sys_arr[i] * 0.6 + mic_arr[i] * 0.6)))
            for i in range(length)
        ])
        return mixed.tobytes()

    # ---- 儲存執行緒 ----
    def _save_after_stop(self):
        """
        等待所有錄音執行緒結束後，轉換並儲存 MP3。
        join timeout=3s：足以涵蓋最壞情況（OSError recovery 的 0.5s sleep + 重開 stream）。
        """
        if self._record_thread:
            self._record_thread.join(timeout=3)
        if self._mic_thread:
            self._mic_thread.join(timeout=3)

        # 所有 stream 已關閉，統一釋放共用 PyAudio 實例
        try:
            self._pa.terminate()
        except Exception:
            pass

        try:
            mode = self._save_mode  # 已在主執行緒鎖定，不從 tkinter StringVar 讀取

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
                    # 麥克風無資料（開啟失敗或立即斷線），退回純系統音訊並警告
                    self.msg_queue.put(("warning", "麥克風無資料，改以「電腦聲音」模式儲存"))
                    pcm_data    = b"".join(self.record_frames)
                    channels    = self.record_channels
                    sample_rate = self.record_sample_rate
                else:
                    if self.record_mic_rate != self.record_sample_rate:
                        # 取樣率不一致，混音仍繼續但聲速會有輕微偏差
                        self.msg_queue.put(("warning",
                            f"麥克風取樣率（{self.record_mic_rate} Hz）與系統音訊"
                            f"（{self.record_sample_rate} Hz）不一致，麥克風聲音可能略有偏差"))

                    self.msg_queue.put(("status", "混音中..."))
                    pcm_data = self._mix_pcm(
                        b"".join(self.record_frames), self.record_channels,
                        b"".join(self.mic_frames),    self.record_mic_channels,
                    )
                    channels    = self.record_channels
                    sample_rate = self.record_sample_rate

            encoder = lameenc.Encoder()
            encoder.set_bit_rate(128)
            encoder.set_in_sample_rate(sample_rate)
            encoder.set_channels(channels)
            encoder.set_quality(2)  # 2=高品質，7=快速低品質

            mp3_data = encoder.encode(pcm_data) + encoder.flush()

            custom_name = self.filename_var.get().strip()
            base_name   = custom_name if custom_name else (
                "meeting_" + datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            )

            # 同名已存在時自動加流水號，避免覆蓋
            filepath = os.path.join(self.save_folder, f"{base_name}.mp3")
            counter = 2
            while os.path.exists(filepath):
                filepath = os.path.join(self.save_folder, f"{base_name} ({counter}).mp3")
                counter += 1

            with open(filepath, "wb") as f:
                f.write(mp3_data)

            self.msg_queue.put(("saved", filepath))

        except Exception as e:
            self.msg_queue.put(("error", str(e)))

    # ---- 執行緒安全 UI 更新 ----
    def _log(self, msg: str):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _reset_ui_after_stop(self):
        """錄音與儲存流程完全結束後，還原所有 UI 元件狀態"""
        self.btn_record.config(state="normal", text="⏺  開始錄音")
        self.timer_label.config(text="00:00", foreground="gray")
        self.silence_banner.grid_remove()
        self._set_mode_radios_state("normal")

    def _poll_queue(self):
        """
        每 100ms 從 msg_queue 拉訊息更新 UI。
        所有 UI 操作都在此（主執行緒）執行，背景執行緒只放訊息進 queue。

        訊息類型：
          saved           — 儲存成功，data = filepath
          error           — 致命錯誤，data = 錯誤訊息
          warning         — 非致命警告，data = 警告訊息（顯示在 log，不中止流程）
          status          — 狀態文字更新，data = 狀態字串
          silence_warning — 靜音偵測，data = True（顯示）/ False（隱藏）
        """
        try:
            while True:
                msg_type, data = self.msg_queue.get_nowait()

                if msg_type == "saved":
                    filename = os.path.basename(data)
                    self._log(f"✓  {filename}")
                    self._reset_ui_after_stop()
                    self.status_label.config(text=f"已儲存：{filename}", foreground="green")

                elif msg_type == "error":
                    self._log(f"[ERROR] {data}")
                    self._reset_ui_after_stop()
                    self.status_label.config(text="發生錯誤，請查看記錄", foreground="red")
                    self.is_recording = False

                elif msg_type == "warning":
                    # 非致命：顯示在 log 但不中斷流程
                    self._log(f"[WARNING] {data}")

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
