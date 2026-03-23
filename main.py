"""
Meeting Recorder
錄製電腦系統音訊（WASAPI Loopback），儲存為 MP3
"""

import os
import sys
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
def get_loopback_device(p: pyaudio.PyAudio) -> dict:
    """取得系統預設輸出的 WASAPI Loopback 裝置"""
    try:
        wasapi_info = p.get_host_api_info_by_type(pyaudio.paWASAPI)
    except OSError:
        raise RuntimeError("找不到 WASAPI 音訊裝置，請確認音效卡驅動正常。")

    default_speakers = p.get_device_info_by_index(wasapi_info["defaultOutputDevice"])

    if not default_speakers.get("isLoopbackDevice", False):
        for loopback in p.get_loopback_device_info_generator():
            if default_speakers["name"] in loopback["name"]:
                return loopback

    return default_speakers


# ---- 主視窗 ----
class MeetingRecorderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Meeting Recorder")
        self.root.resizable(False, False)

        self.is_recording = False
        self.record_frames: list[bytes] = []
        self.record_channels = 2
        self.record_sample_rate = 44100
        self.start_time: float = 0.0
        self.msg_queue: queue.Queue = queue.Queue()
        self._record_thread: threading.Thread | None = None

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.save_folder = desktop
        self.save_folder_var = tk.StringVar(value=desktop)
        self.filename_var = tk.StringVar()

        self._build_ui()
        self._poll_queue()

    # ---- UI 建置 ----
    def _build_ui(self):
        pad = {"padx": 14, "pady": 6}

        # 儲存位置
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

        # 檔案名稱
        frame_name = ttk.LabelFrame(self.root, text=" 檔案名稱 ", padding=8)
        frame_name.grid(row=1, column=0, sticky="ew", **pad)
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

        # 錄音按鈕區
        frame_btn = tk.Frame(self.root)
        frame_btn.grid(row=2, column=0, pady=20)

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

        # 靜音警告橫幅（預設隱藏，偵測到靜音才顯示）
        self.silence_banner = tk.Frame(self.root, background="#FFA500", padx=12, pady=8)
        tk.Label(
            self.silence_banner,
            text="⚠  偵測到超過 10 秒沒有聲音，請確認：\n"
                 "系統是否靜音？播放裝置是否正確？",
            background="#FFA500", foreground="white",
            font=("", 10, "bold"), justify="left"
        ).pack(anchor="w")

        # 錄音記錄
        frame_log = ttk.LabelFrame(self.root, text=" 錄音記錄 ", padding=8)
        frame_log.grid(row=4, column=0, sticky="ew", padx=14, pady=(0, 14))

        self.log_text = scrolledtext.ScrolledText(
            frame_log, width=52, height=6,
            state="disabled", font=("Consolas", 9)
        )
        self.log_text.pack(fill="x")

        self.root.columnconfigure(0, weight=1)
        self._log("請確認儲存位置，然後按「開始錄音」。")

    # ---- UI 互動 ----
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

    def _start_recording(self):
        self.is_recording = True
        self.record_frames = []
        self.start_time = time.time()

        self.btn_record.config(text="⏹  停止並儲存")
        self.status_label.config(text="錄音中...", foreground="red")
        self.timer_label.config(foreground="red")
        self.filename_entry.config(state="disabled")  # 錄音中鎖定，避免誤改

        self._update_timer()

        self._record_thread = threading.Thread(target=self._record_worker, daemon=True)
        self._record_thread.start()

    def _update_timer(self):
        """每秒更新計時器（主執行緒安全）"""
        if self.is_recording:
            elapsed = int(time.time() - self.start_time)
            mins = elapsed // 60
            secs = elapsed % 60
            self.timer_label.config(text=f"{mins:02d}:{secs:02d}")
            self.root.after(1000, self._update_timer)

    def _stop_recording(self):
        self.is_recording = False
        self.btn_record.config(state="disabled", text="儲存中...")
        self.status_label.config(text="轉換為 MP3 中...", foreground="gray")
        self.timer_label.config(foreground="gray")
        self.filename_entry.config(state="normal")

        # 等待錄音執行緒結束後再儲存，避免 race condition
        t = threading.Thread(target=self._save_after_stop, daemon=True)
        t.start()

    # ---- 錄音執行緒 ----
    def _record_worker(self):
        p = pyaudio.PyAudio()
        try:
            chunk = 512

            def open_stream(channels=None, sample_rate=None):
                """取得當前預設播放裝置的 loopback stream，沿用指定的格式參數"""
                device = get_loopback_device(p)
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

            SILENCE_RMS_THRESHOLD = 100   # Int16 最大值 32767，低於此視為靜音
            SILENCE_WARNING_SECS  = 10
            silence_start  = None
            silence_warned = False

            while self.is_recording:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    self.record_frames.append(data)

                    # ---- 靜音偵測（計算 RMS）----
                    num_samples = len(data) // 2
                    if num_samples > 0:
                        samples = struct.unpack(f"{num_samples}h", data)
                        rms = math.sqrt(sum(s * s for s in samples) / num_samples)

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
                    # 裝置失效（插拔耳機 / 切換播放裝置），重新取得新裝置
                    # 沿用原格式（channels/sample_rate）確保 PCM 資料前後一致
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
                        continue

            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass

        except Exception as e:
            self.msg_queue.put(("error", str(e)))
        finally:
            p.terminate()

    # ---- 儲存執行緒 ----
    def _save_after_stop(self):
        """等待錄音執行緒完全結束，再執行 MP3 轉換儲存"""
        if self._record_thread:
            self._record_thread.join(timeout=3)

        try:
            if not self.record_frames:
                self.msg_queue.put(("error", "沒有錄到任何音訊"))
                return

            pcm_data = b"".join(self.record_frames)

            encoder = lameenc.Encoder()
            encoder.set_bit_rate(128)
            encoder.set_in_sample_rate(self.record_sample_rate)
            encoder.set_channels(self.record_channels)
            encoder.set_quality(2)  # 2=高品質

            mp3_data = encoder.encode(pcm_data) + encoder.flush()

            custom_name = self.filename_var.get().strip()
            if custom_name:
                base_name = custom_name
            else:
                base_name = "meeting_" + datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

            # 同名檔案已存在時自動加 (2), (3)...
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

    def _poll_queue(self):
        """每 100ms 從 queue 拉訊息更新 UI"""
        try:
            while True:
                msg_type, data = self.msg_queue.get_nowait()
                if msg_type == "saved":
                    filepath = data
                    filename = os.path.basename(filepath)
                    self._log(f"✓  {filename}")
                    self.btn_record.config(state="normal", text="⏺  開始錄音")
                    self.status_label.config(
                        text=f"已儲存：{filename}", foreground="green"
                    )
                    self.timer_label.config(text="00:00", foreground="gray")
                    self.silence_banner.grid_remove()
                elif msg_type == "error":
                    self._log(f"[ERROR] {data}")
                    self.btn_record.config(state="normal", text="⏺  開始錄音")
                    self.status_label.config(text="發生錯誤，請查看記錄", foreground="red")
                    self.timer_label.config(text="00:00", foreground="gray")
                    self.silence_banner.grid_remove()
                    self.is_recording = False
                elif msg_type == "silence_warning":
                    if data:  # True = 顯示警告
                        self.silence_banner.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 6))
                    else:     # False = 聲音恢復，隱藏警告
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
