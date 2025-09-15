import tkinter as tk
from tkinter import Label, Button, Entry, IntVar, Frame, Checkbutton
from PIL import Image, ImageTk
import cv2
import openpyxl
import subprocess
import os
import signal
import threading
import queue
import atexit
import traceback
import numpy as np

def load_students_info(excel_path, sheet_index=0):
    print(f"Loading students info from sheet index: {sheet_index}")
    workbook = openpyxl.load_workbook(excel_path, read_only=True)
    sheets = workbook.sheetnames
    if 0 <= sheet_index < len(sheets):
        sheet = workbook[sheets[sheet_index]]
    else:
        print(f"Invalid sheet index: {sheet_index}. Using the first sheet.")
        sheet = workbook.active
    students_info = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        exam_id, name = row[0], row[1]
        if exam_id and name:
            students_info.append((exam_id, name))
    print(f"Loaded {len(students_info)} students")
    workbook.close()
    return students_info

class CameraApp:
    def __init__(self, master, excel_path):
        self.master = master
        self.excel_path = excel_path
        self.students_info = []
        self.current_student_index = 0
        self.ffmpeg_process = None
        self.master.title("学生录像系统")
        self.master.geometry("1000x800")

        self.vid = cv2.VideoCapture(0)
        self.vid.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
        self.vid.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)

        # 创建一个主框架来包含所有元素
        self.main_frame = Frame(master)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建画布框架，并使其能够扩展
        self.canvas_frame = Frame(self.main_frame)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.canvas_frame)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.label = Label(self.main_frame, text="加载中...", font=("Arial", 12))
        self.label.pack(pady=5)

        self.recording_status = Label(self.main_frame, text="就绪", font=("Arial", 12), fg="green")
        self.recording_status.pack(pady=5)

        button_frame = Frame(self.main_frame)
        button_frame.pack(pady=10)

        self.btn_recording = Button(button_frame, text="开始录像", command=self.toggle_recording, width=15, height=2)
        self.btn_recording.pack(side=tk.LEFT, padx=5)

        self.btn_snapshot = Button(button_frame, text="拍照", command=self.take_snapshot, width=15, height=2)
        self.btn_snapshot.pack(side=tk.LEFT, padx=5)

        self.btn_previous = Button(button_frame, text="上一个学生", command=self.previous_student, width=15, height=2)
        self.btn_previous.pack(side=tk.LEFT, padx=5)

        self.btn_next = Button(button_frame, text="下一个学生", command=self.next_student, width=15, height=2)
        self.btn_next.pack(side=tk.LEFT, padx=5)

        self.rotate_var = IntVar(value=0)
        self.chk_rotate = Checkbutton(button_frame, text="旋转180度", variable=self.rotate_var, command=self.toggle_rotation, width=10, height=2)
        self.chk_rotate.pack(side=tk.LEFT, padx=5)

        sheet_frame = Frame(self.main_frame)
        sheet_frame.pack(pady=10)
        
        Label(sheet_frame, text="Sheet索引：").pack(side=tk.LEFT)
        self.sheet_entry = Entry(sheet_frame, width=5)
        self.sheet_entry.insert(0, "0")
        self.sheet_entry.pack(side=tk.LEFT, padx=5)

        self.btn_load_sheet = Button(sheet_frame, text="加载Sheet", command=self.load_sheet, width=10, height=1)
        self.btn_load_sheet.pack(side=tk.LEFT, padx=5)

        self.queue = queue.Queue()
        atexit.register(self.cleanup)

        self.master.after(100, self.load_excel_data)
        self.master.after(10, self.update)
        self.master.after(100, self.process_queue)

        # 绑定窗口大小变化事件
        self.master.bind("<Configure>", self.on_resize)

    def on_resize(self, event):
        # 当窗口大小改变时，调整画布大小
        self.update_canvas_size()

    def update_canvas_size(self):
        # 获取主窗口的当前大小
        window_width = self.master.winfo_width()
        window_height = self.master.winfo_height()

        # 计算画布应该的大小（减去按钮和标签的高度）
        canvas_height = window_height - 150  # 假设按钮和标签总高度约为150像素
        canvas_width = window_width - 20  # 留一些边距

        # 设置画布大小
        self.canvas.config(width=canvas_width, height=canvas_height)

    def load_excel_data(self):
        print("Starting to load Excel data")
        threading.Thread(target=self._load_excel_data_thread, daemon=True).start()

    def _load_excel_data_thread(self):
        try:
            students_info = load_students_info(self.excel_path, 0)
            self.queue.put(("update_students", students_info))
        except Exception as e:
            print(f"Error in _load_excel_data_thread: {e}")
            traceback.print_exc()
            self.queue.put(("error", str(e)))
        finally:
            self.queue.put(("done", None))

    def process_queue(self):
        try:
            while True:
                message, data = self.queue.get_nowait()
                if message == "update_students":
                    print(f"Updating students: {len(data)} students loaded")
                    self.students_info = data
                    self.update_student_info()
                elif message == "error":
                    print(f"An error occurred: {data}")
                elif message == "done":
                    print("Finished processing queue")
                    break
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self.process_queue)

    def load_sheet(self):
        try:
            sheet_index = int(self.sheet_entry.get())
            print(f"Loading sheet index: {sheet_index}")
            self.label.config(text="正在加载...")
            threading.Thread(target=self._load_new_sheet, args=(sheet_index,), daemon=True).start()
        except ValueError:
            print("Invalid sheet index")
            self.label.config(text="无效的Sheet索引")

    def _load_new_sheet(self, sheet_index):
        try:
            students_info = load_students_info(self.excel_path, sheet_index)
            self.queue.put(("update_students", students_info))
        except Exception as e:
            print(f"Error in _load_new_sheet: {e}")
            traceback.print_exc()
            self.queue.put(("error", str(e)))
        finally:
            self.queue.put(("done", None))

    def toggle_rotation(self):
        print(f"Rotation toggled: {'开启' if self.rotate_var.get() == 1 else '关闭'}")

    def toggle_recording(self):
        if hasattr(self, 'is_recording') and self.is_recording:
            self.stop_recording()
            self.next_student()
        else:
            self.start_recording()

    def start_recording(self):
        exam_id, name = self.students_info[self.current_student_index]
        video_name = f"{exam_id}_{name}.mp4"
        
        # 创建一个管道来传输视频帧
        self.frame_queue = queue.Queue(maxsize=100)
        
        # 启动一个新线程来处理视频帧
        self.recording_thread = threading.Thread(target=self.process_frames, args=(video_name,))
        self.recording_thread.start()
        
        self.is_recording = True
        self.recording_status.config(text="正在录像", fg="red")
        self.btn_recording.config(text="结束录像")
        print(f"Recording started for {name} ({exam_id}).")

    def process_frames(self, video_name):
        fourcc = cv2.VideoWriter_fourcc(*'mp4v')
        out = cv2.VideoWriter(video_name, fourcc, 30.0, (1920, 1080))
        
        while self.is_recording:
            try:
                frame = self.frame_queue.get(timeout=1)
                out.write(frame)
            except queue.Empty:
                continue
        
        out.release()

    def stop_recording(self):
        if hasattr(self, 'is_recording') and self.is_recording:
            self.is_recording = False
            self.recording_thread.join()
            exam_id, name = self.students_info[self.current_student_index]
            print(f"Recording stopped and saved for {name} ({exam_id}).")
            self.recording_status.config(text="就绪", fg="green")
            self.btn_recording.config(text="开始录像")

    def take_snapshot(self):
        ret, frame = self.vid.read()
        if ret:
            if self.rotate_var.get() == 1:
                frame = cv2.rotate(frame, cv2.ROTATE_180)
            exam_id, name = self.students_info[self.current_student_index]
            photo_name = f"{exam_id}_{name}.png"
            cv2.imwrite(photo_name, frame)
            print(f"Photo saved as {photo_name}")
            # 如果正在录像，显示拍照提示
            if hasattr(self, 'is_recording') and self.is_recording:
                print(f"Photo taken during recording for {name} ({exam_id})")

    def next_student(self):
        if self.current_student_index < len(self.students_info) - 1:
            if hasattr(self, 'is_recording') and self.is_recording:
                self.stop_recording()
            self.current_student_index += 1
            self.update_student_info()

    def previous_student(self):
        if self.current_student_index > 0:
            if hasattr(self, 'is_recording') and self.is_recording:
                self.stop_recording()
            self.current_student_index -= 1
            self.update_student_info()

    def update_student_info(self):
        if self.students_info:
            exam_id, name = self.students_info[self.current_student_index]
            self.label.config(text=f"当前学生：{name} ({exam_id})")
        else:
            self.label.config(text="没有学生信息")

    def update(self):
        ret, frame = self.vid.read()
        if ret:
            if self.rotate_var.get() == 1:
                frame = cv2.rotate(frame, cv2.ROTATE_180)
            
            # 如果正在录像，将帧添加到队列中
            if hasattr(self, 'is_recording') and self.is_recording:
                if self.frame_queue.qsize() < 100:  # 限制队列大小以防内存溢出
                    self.frame_queue.put(cv2.resize(frame, (1920, 1080)))
            
            # 获取当前画布大小
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()

            # 调整frame大小以适应画布，保持宽高比
            frame_height, frame_width = frame.shape[:2]
            aspect_ratio = frame_width / frame_height
            
            if canvas_width / canvas_height > aspect_ratio:
                new_height = canvas_height
                new_width = int(new_height * aspect_ratio)
            else:
                new_width = canvas_width
                new_height = int(new_width / aspect_ratio)

            frame_resized = cv2.resize(frame, (new_width, new_height))
            self.photo = ImageTk.PhotoImage(image=Image.fromarray(cv2.cvtColor(frame_resized, cv2.COLOR_BGR2RGB)))
            
            # 清除之前的图像并创建新的图像
            self.canvas.delete("all")
            self.canvas.create_image(canvas_width//2, canvas_height//2, image=self.photo, anchor=tk.CENTER)

        self.master.after(10, self.update)

    def cleanup(self):
        print("Cleaning up resources...")
        if hasattr(self, 'is_recording') and self.is_recording:
            self.stop_recording()
        if self.vid.isOpened():
            self.vid.release()
        cv2.destroyAllWindows()
        print("Cleanup completed.")

    def on_closing(self):
        print("Window is closing. Cleaning up...")
        self.cleanup()
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    excel_path = "mt.xlsx"
    app = CameraApp(root, excel_path)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    try:
        root.mainloop()
    except Exception as e:
        print(f"An error occurred in main loop: {e}")
        traceback.print_exc()
    finally:
        print("Program is exiting. Performing final cleanup...")
        app.cleanup()