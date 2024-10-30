import ctypes
import os
import re
import subprocess
import threading
import time
import tkinter as tk
import wave
from tkinter import ttk

import numpy as np
import pyaudio
import win32gui

ctypes.windll.shcore.SetProcessDpiAwareness(1)

# 采样格式
FORMATS = {
    "8位": pyaudio.paInt8,
    "16位": pyaudio.paInt16,
    "24位": pyaudio.paInt24,
    "32位": pyaudio.paInt32,
    "32位浮点": pyaudio.paFloat32
}

WAVEFORM_SIZE = 100
WAVEFORM_SCALE = 4

RECORD_DIR = "recordings"
SONG_DIR = "songs"


# 根据时间生成文件名
def generate_filename():
    return time.strftime("%Y%m%d%H%M%S")


def parse_title(title):
    # 按照 - 分割歌曲名和歌手
    parts = title.split(" - ")
    song = parts[0].strip()
    artist = parts[1].strip() if len(parts) > 1 else ""
    # 按照 / 歌手
    parts = artist.split(" / ")
    return song, parts


# 获取桌面窗口列表，返回(窗口句柄，窗口标题)列表
def get_window_list():
    window_list = []

    def enum_windows(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                window_list.append((hwnd, title))
        return True

    win32gui.EnumWindows(enum_windows, None)
    return window_list


def convert_wav_to_flac(wav_file, flac_file, metadata):
    command = [
        ".\\ffmpeg.exe",
        '-loglevel', 'warning',
        "-i", wav_file,
        "-metadata", f"title={metadata['title']}",
        "-metadata", f"artist={'; '.join(metadata['artist'])}",
        "-y",
        flac_file
    ]
    print(f"正在转换为flac... {command}")
    subprocess.run(command)


def is_silent(data, threshold=500):
    """ 判断音频数据是否为静音 """
    return np.max(np.abs(data)) < threshold


def remove_silence(input_file, output_file, threshold=500):
    with wave.open(input_file, 'rb') as wf:
        # 读取音频参数
        n_channels = wf.getnchannels()
        sampwidth = wf.getsampwidth()
        framerate = wf.getframerate()
        n_frames = wf.getnframes()

        # 读取音频数据
        audio_data = wf.readframes(n_frames)

    # 将字节数据转换为 NumPy 数组
    audio_array = np.frombuffer(audio_data, dtype=np.int16)

    # 找到开始和结束的非静音部分
    start_index = 0
    end_index = n_frames

    # 查找前静音
    for i in range(len(audio_array)):
        if not is_silent(audio_array[i:i + 1], threshold):
            start_index = i
            break

    # 查找后静音
    for i in range(len(audio_array) - 1, -1, -1):
        if not is_silent(audio_array[i:i + 1], threshold):
            end_index = i + 1
            break

    # 切割出非静音部分
    trimmed_audio = audio_array[start_index:end_index]

    # 保存为新文件
    with wave.open(output_file, 'wb') as wf_out:
        wf_out.setnchannels(n_channels)
        wf_out.setsampwidth(sampwidth)
        wf_out.setframerate(framerate)
        wf_out.writeframes(trimmed_audio.tobytes())

    print(f"去除静音成功，已保存: {output_file}")


class AudioRecorder:
    def __init__(self):
        # 初始化 PyAudio
        self.p = pyaudio.PyAudio()
        self.stream = None

        # 全局变量
        self.is_recording = False
        self.waveform = None
        self.waveform_buffer = None
        self.waveform_init()
        self.wavefile = None
        self.wavefile_name = None
        # 无声时间
        self.silence_time = 0
        self.silence_watch_enabled = False
        self.is_need_split = False
        self.song_name = None
        self.song_metadata = {}

        self.root = None
        self.device_combobox = None
        self.channels_combobox = None
        self.rate_combobox = None
        self.chunk_combobox = None
        self.format_combobox = None
        self.status_label = None
        self.recording_dot = None
        self.waveform_canvas = None
        self.filename_entry = None
        self.recording_time_label = None
        self.start_time = None
        self.start_button = None
        self.stop_button = None
        self.auto_split_checkbutton = None
        self.auto_split_var = None
        self.auto_rename_checkbutton = None
        self.window_combobox = None
        self.auto_rename_var = None
        self.convert_flac_var = None
        self.convert_flac_checkbutton = None
        # 自动化按钮
        self.automatic_button = None
        self.setup_gui()

    def waveform_init(self):
        self.waveform = np.zeros(WAVEFORM_SIZE)
        self.waveform_buffer = np.zeros(0)

    # 列出可用设备
    def list_devices(self):
        device_list = []
        for i in range(self.p.get_device_count()):
            device_info = self.p.get_device_info_by_index(i)
            if device_info['maxInputChannels'] > 0:
                device_list.append((device_info['index'], device_info['name']))
        return device_list

    def process_wav_file(self, old_filename, song_name, metadata):
        old_filename = os.path.join(RECORD_DIR, old_filename)

        # 去除前后静音
        remove_silence(old_filename, old_filename)

        new_filename = os.path.join(SONG_DIR, song_name + ".flac")
        i = 1
        while os.path.exists(new_filename):
            new_filename = os.path.join(SONG_DIR, f"{song_name}({i}).flac")
            i += 1
        if self.convert_flac_var.get():
            convert_wav_to_flac(old_filename, new_filename, metadata)
        else:
            os.rename(old_filename, new_filename)
            print(f"文件已重命名为：{new_filename}")

    def start_recording(self):
        self.is_recording = True
        self.start_time = time.time()
        self.waveform_init()
        self.status_label.config(text="正在录音...")
        self.recording_dot.config(fg='red')
        self.blink_dot()
        self.filename_entry.config(state=tk.DISABLED)
        self.device_combobox.config(state=tk.DISABLED)
        self.channels_combobox.config(state=tk.DISABLED)
        self.rate_combobox.config(state=tk.DISABLED)
        self.chunk_combobox.config(state=tk.DISABLED)
        self.format_combobox.config(state=tk.DISABLED)
        self.recording_time_label.config(text="00:00:00")
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        threading.Thread(target=self.record_audio, daemon=True).start()

    def update_recording_time(self):
        if self.is_recording:
            recording_time = time.time() - self.start_time
            self.recording_time_label.config(text=time.strftime("%H:%M:%S", time.gmtime(recording_time)))

    def blink_dot(self):
        if self.is_recording:
            current_color = self.recording_dot.cget("fg")
            new_color = "red" if current_color != "red" else "#f0f0f0"
            self.recording_dot.config(fg=new_color)
            self.status_label.after(500, self.blink_dot)
        else:
            self.recording_dot.config(fg="black")
        self.update_recording_time()

    def update_waveform(self, format_, channels, rate, data):
        dt = None
        match format_:
            case pyaudio.paInt8:
                dt = np.int8
            case pyaudio.paInt16:
                dt = np.int16
            case pyaudio.paInt24:
                dt = np.int32
            case pyaudio.paInt32:
                dt = np.int32
            case pyaudio.paFloat32:
                dt = np.float32
        data = np.frombuffer(data, dtype=dt)
        data = data.reshape(-1, channels)
        data = data.mean(axis=1)

        self.waveform_buffer = np.append(self.waveform_buffer, data)
        window_size = rate // 100 * WAVEFORM_SCALE  # 每 10ms 取一个样本
        window_time = window_size / rate
        while len(self.waveform_buffer) >= window_size:
            start = 0
            end = window_size
            sample = self.waveform_buffer[start:end]
            amplitude = np.abs(np.max(sample) - np.min(sample))
            if dt != np.float32:
                amplitude /= np.iinfo(dt).max * 2
            self.waveform = np.append(self.waveform, amplitude)
            self.waveform_buffer = self.waveform_buffer[window_size:]

            # 计算无声时间，单位为秒
            if amplitude < 0.01:
                self.silence_time += window_time
            else:
                self.silence_time = 0
                self.silence_watch_enabled = True
            if self.silence_time > 0.5 and self.silence_watch_enabled and time.time() - self.start_time > 10:
                if self.auto_split_var.get():
                    self.is_need_split = True
                self.silence_watch_enabled = False

        if len(self.waveform) > WAVEFORM_SIZE:
            self.waveform = self.waveform[-WAVEFORM_SIZE:]

    def draw_waveform(self):
        self.waveform_canvas.delete("all")
        for i in range(WAVEFORM_SIZE):
            amplitude = self.waveform[i]
            x0 = i * WAVEFORM_SCALE
            y0 = 25 - int(amplitude * 25)
            x1 = (i + 1) * WAVEFORM_SCALE
            y1 = 27 + int(amplitude * 25)
            self.waveform_canvas.create_rectangle(x0, y0, x1, y1, fill="green")
        self.root.after(50, self.draw_waveform)

    def new_wavefile(self, filename, channels, rate, samp_width):
        if self.wavefile:
            self.wavefile.close()
        self.wavefile = wave.open(os.path.join(RECORD_DIR, filename), 'wb')
        self.wavefile.setnchannels(channels)
        self.wavefile.setsampwidth(samp_width)
        self.wavefile.setframerate(rate)
        self.wavefile_name = filename

        return self.wavefile

    def record(self, input_device_index, format_, channels, rate, chunk):
        wf = self.new_wavefile(self.get_filename(), channels, rate, self.p.get_sample_size(format_))
        data_written = 0
        self.stream = self.p.open(format=format_,
                                  channels=channels,
                                  rate=rate,
                                  input=True,
                                  input_device_index=input_device_index,
                                  frames_per_buffer=chunk)

        print(
            f"开始录音... 设备：{input_device_index}，通道数：{channels}，采样率：{rate}，块大小：{chunk}，格式：{format_}，文件名：{self.get_filename()}")
        while self.is_recording:
            data = self.stream.read(chunk)
            self.update_waveform(format_, channels, rate, data)

            data_written += len(data)
            if data_written > 4294967295 or self.is_need_split:
                old_filename = self.get_filename()
                self.set_filename(generate_filename())
                self.is_need_split = False
                self.start_time = time.time()
                print(f"录音已分割，新文件名：{self.get_filename()}")
                wf = self.new_wavefile(self.get_filename(), channels, rate, self.p.get_sample_size(format_))
                data_written = len(data)

                if self.auto_rename_var.get() and (song_name := self.song_name.get()):
                    # 重命名文件
                    threading.Thread(target=self.process_wav_file, args=(old_filename, song_name, self.song_metadata), daemon=True).start()

            wf.writeframes(data)
        print("录音结束。")
        self.status_label.config(text="录音已停止。")
        self.recording_dot.config(fg='black')

        wf.close()
        self.stream.stop_stream()
        self.stream.close()

    def record_audio(self):
        # 音频参数
        input_device_index = int(self.device_combobox.get().split(":")[0])
        format_ = self.get_format()
        channels = int(self.channels_combobox.get())
        rate = int(self.rate_combobox.get())
        chunk = int(self.chunk_combobox.get())

        try:
            self.record(input_device_index, format_, channels, rate, chunk)
        except Exception as e:
            print(e)
            self.status_label.config(text="录音失败！")
            self.recording_dot.config(fg='black')

    def stop_recording(self):
        self.is_recording = False
        self.filename_entry.config(state=tk.NORMAL)
        self.device_combobox.config(state=tk.NORMAL)
        self.channels_combobox.config(state=tk.NORMAL)
        self.rate_combobox.config(state=tk.NORMAL)
        self.chunk_combobox.config(state=tk.NORMAL)
        self.format_combobox.config(state=tk.NORMAL)
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def get_format(self):
        format_str = self.format_combobox.get()
        return FORMATS[format_str]

    def get_filename(self):
        return self.filename_entry.get() + ".wav"

    def set_filename(self, filename):
        self.filename_entry.config(state=tk.NORMAL)
        self.filename_entry.delete(0, tk.END)
        self.filename_entry.insert(0, filename)
        self.filename_entry.config(state=tk.DISABLED)

    def open_automatic(self):
        automatic_window = tk.Toplevel(self.root)
        automatic_window.title("自动化")
        automatic_window.resizable(False, False)
        frame = ttk.Frame(automatic_window)
        frame.pack(padx=8, pady=8)

        self.auto_split_checkbutton = ttk.Checkbutton(frame, text="无声时自动分割", variable=self.auto_split_var)
        self.auto_split_checkbutton.grid(row=0, sticky="w", column=1)

        # 进程ID-窗口标题列表
        window_list = get_window_list()
        select_list = [f"{hwnd}: {title}" for hwnd, title in window_list]

        ttk.Label(frame, text="选择窗口：").grid(row=1, column=0, sticky="e")
        self.window_combobox = ttk.Combobox(frame, values=select_list, width=40)
        self.window_combobox.grid(row=1, column=1, sticky="w")
        self.window_combobox.current(0)

        # 获取窗口标题作为文件名
        self.auto_rename_checkbutton = ttk.Checkbutton(frame, text="使用窗口标题作为歌曲名",
                                                       variable=self.auto_rename_var, command=self.auto_rename)
        self.auto_rename_checkbutton.grid(row=2, column=1, sticky="w")

        # 歌曲名
        ttk.Label(frame, text="歌曲名：").grid(row=3, column=0, sticky="e")
        ttk.Label(frame, textvariable=self.song_name).grid(row=3, column=1, sticky="w")

        # 是否转换为flac
        self.convert_flac_checkbutton = ttk.Checkbutton(frame, text="转换为flac", variable=self.convert_flac_var)
        self.convert_flac_checkbutton.grid(row=4, column=1, sticky="w")

    def auto_rename(self):
        if not self.auto_rename_var.get():
            self.window_combobox.config(state=tk.NORMAL)
            return
        self.window_combobox.config(state=tk.DISABLED)

        hwnd = int(self.window_combobox.get().split(":")[0])
        old_title = ""
        duration = 0

        def watch_window_title():
            nonlocal old_title, duration
            if not self.auto_rename_var.get():
                return
            title = str(win32gui.GetWindowText(hwnd)).strip()
            song_title, artist = parse_title(title)
            title = f"{','.join(artist)}-{song_title}"
            if not artist:
                title = song_title
            if title and title != old_title:
                if duration > 5:
                    old_title = title
                    self.song_name.set(title)
                    self.song_metadata = {
                        "title": song_title,
                        "artist": artist
                    }
                    duration = 0
                duration += 5

            self.root.after(5000, watch_window_title)

        watch_window_title()

    def setup_gui(self):
        # 创建录音目录
        if not os.path.exists(RECORD_DIR):
            os.makedirs(RECORD_DIR)
        if not os.path.exists(SONG_DIR):
            os.makedirs(SONG_DIR)
        # 设置 Tkinter 界面
        self.root = tk.Tk()
        self.root.resizable(False, False)
        self.root.title("音乐录制器")

        self.auto_split_var = tk.IntVar(value=0)
        self.auto_rename_var = tk.IntVar(value=0)
        self.song_name = tk.StringVar()
        self.convert_flac_var = tk.IntVar(value=0)

        frame = ttk.Frame(self.root)
        frame.pack(padx=8, pady=8)

        conf_frame = ttk.Frame(frame)
        conf_frame.grid(row=0, sticky="w")
        conf_frame.grid_rowconfigure(0, pad=4)
        conf_frame.grid_rowconfigure(1, pad=4)
        conf_frame.grid_rowconfigure(2, pad=4)
        conf_frame.grid_rowconfigure(3, pad=4)
        conf_frame.grid_rowconfigure(4, pad=4)

        # 设备选择
        ttk.Label(conf_frame, text="选择设备：").grid(row=0, column=0, sticky="e")
        devices = self.list_devices()
        self.device_combobox = ttk.Combobox(conf_frame, values=[f"{index}: {name}" for index, name in devices])
        self.device_combobox.grid(row=0, column=1, columnspan=3, sticky="ew")
        default_device_index = self.p.get_default_input_device_info()['index']
        self.device_combobox.current([index for index, name in devices].index(default_device_index))

        # 音频参数
        ttk.Label(conf_frame, text="通道数：").grid(row=1, column=0, sticky="e")
        self.channels_combobox = ttk.Combobox(conf_frame, values=[str(i) for i in range(1, 9)], width=10)
        self.channels_combobox.grid(row=1, column=1, sticky="w")
        self.channels_combobox.current(1)

        ttk.Label(conf_frame, text="采样率：").grid(row=1, column=2, sticky="e")
        self.rate_combobox = ttk.Combobox(conf_frame,
                                          values=["8000", "11025", "16000", "22050", "32000", "44100", "48000", "64000",
                                                  "88200", "96000", "176400", "192000"], width=10)
        self.rate_combobox.grid(row=1, column=3, sticky="w")
        self.rate_combobox.current(5)

        ttk.Label(conf_frame, text="块大小：").grid(row=3, column=0, sticky="e")
        self.chunk_combobox = ttk.Combobox(conf_frame,
                                           values=["128", "256", "512", "1024", "2048", "4096", "5120", "6144", "7168",
                                                   "8192", "12288", "16384", "32768"], width=10)
        self.chunk_combobox.grid(row=3, column=1, sticky="w")
        self.chunk_combobox.current(8)

        ttk.Label(conf_frame, text="采样格式：").grid(row=3, column=2, sticky="e", padx=(40, 0))
        self.format_combobox = ttk.Combobox(conf_frame, values=[format_str for format_str in FORMATS.keys()], width=10)
        self.format_combobox.grid(row=3, column=3, sticky="w")
        self.format_combobox.current(1)

        # 文件名
        ttk.Label(conf_frame, text="文件名：").grid(row=4, column=0, sticky="e")
        self.filename_entry = ttk.Entry(conf_frame, width=40)
        self.filename_entry.grid(row=4, column=1, columnspan=3, sticky="ew")
        self.filename_entry.insert(0, generate_filename())

        operation_frame = ttk.Frame(frame)
        operation_frame.grid(row=1, pady=(10, 0), sticky="e")

        # 按钮
        self.start_button = ttk.Button(operation_frame, text="开始录音", command=self.start_recording)
        self.start_button.grid(row=0, column=0)

        # 停止按钮, 默认禁用
        self.stop_button = ttk.Button(operation_frame, text="停止录音", command=self.stop_recording, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=(4, 0))

        # 自动化按钮
        self.automatic_button = ttk.Button(operation_frame, text="自动化", command=self.open_automatic)
        self.automatic_button.grid(row=0, column=2, padx=(4, 0))

        status_frame = ttk.Frame(frame)
        status_frame.grid(row=2, pady=(10, 0), sticky="w")

        # 状态标签和闪烁的红点
        ttk.Label(status_frame, text="REC: ", font="system").grid(row=0, column=0, sticky="e")
        self.recording_dot = tk.Label(status_frame, text="", font="segoemdl2", fg="black")
        self.recording_dot.grid(row=0, column=1, sticky="w", padx=(4, 0))
        self.recording_time_label = ttk.Label(status_frame, text="00:00:00", font="system")
        self.recording_time_label.grid(row=0, column=2, sticky="e", padx=(10, 0))

        ttk.Label(status_frame, text="状态：").grid(row=0, column=3, sticky="e", padx=(40, 0))
        self.status_label = ttk.Label(status_frame, text="准备录音...")
        self.status_label.grid(row=0, column=4, sticky="w")

        self.waveform_canvas = tk.Canvas(frame, width=WAVEFORM_SIZE * WAVEFORM_SCALE, height=50, bg="black")
        self.waveform_canvas.grid(row=3, pady=(10, 0))

        self.draw_waveform()
        # 启动 Tkinter 事件循环
        self.root.mainloop()


if __name__ == "__main__":
    AudioRecorder()
