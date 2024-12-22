import win32gui
import win32process
import psutil
import pythoncom
from win32com.client import Dispatch
import tkinter as tk
from collections import deque
import os
import subprocess
import threading
import time

class PathQueue:
    def __init__(self, maxsize=10):
        self.queue = deque(maxlen=maxsize)
        self.lock = threading.Lock()
        self.callbacks = []
        self.root = None
        self.last_path = None

    def add_observer(self, callback):
        self.callbacks.append(callback)

    def notify_observers(self):
        if self.callbacks and self.root:
            self.root.after(0, self.callbacks[0])

    def add_path(self, path):
        if path != self.last_path:
            with self.lock:
                self.last_path = path
                if path in self.queue:
                    self.queue.remove(path)
                self.queue.append(path)
                self.notify_observers()

    def get_paths(self):
        with self.lock:
            return list(self.queue)

class ExplorerTracker:
    def __init__(self):
        self.path_queue = PathQueue(maxsize=10)
        self.running = False
        self.last_active_window = None

    def decode_mixed_encoding(self, path):
        if path.startswith('file:///'):
            path = path[8:]
        
        replacements = {
            '%DC': 'Ü',
            '%D6': 'Ö',
            '%20': ' '
        }
        
        for enc, char in replacements.items():
            path = path.replace(enc, char)
        
        return path.replace('/', '\\')

    def get_explorer_path(self, hwnd):
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            
            if process.name().lower() == 'explorer.exe':
                pythoncom.CoInitialize()
                shell = Dispatch("Shell.Application")
                windows = shell.Windows()
                
                for window in windows:
                    try:
                        if window.HWND == hwnd:
                            path = window.LocationURL
                            if path.startswith('file:///'):
                                return self.decode_mixed_encoding(path)
                    except:
                        continue
                        
                pythoncom.CoUninitialize()
        except:
            pass
        return None

    def check_active_window(self):
        hwnd = win32gui.GetForegroundWindow()
        if hwnd != self.last_active_window:
            self.last_active_window = hwnd
            path = self.get_explorer_path(hwnd)
            if path:
                self.path_queue.add_path(path)

    def start(self):
        self.running = True
        def track_loop():
            while self.running:
                self.check_active_window()
                time.sleep(0.5)
        self.tracker_thread = threading.Thread(target=track_loop, daemon=True)
        self.tracker_thread.start()

    def stop(self):
        self.running = False
        if hasattr(self, 'tracker_thread'):
            self.tracker_thread.join(timeout=1.0)

class GUI:
    def __init__(self, path_queue):
        self.root = tk.Tk()
        self.root.title("Recent Path History")
        self.path_queue = path_queue
        self.path_queue.root = self.root
        
        self.path_queue.add_observer(self.update_list)
        
        window_width = 400
        window_height = 400
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')

        title_label = tk.Label(
            self.root,
            text="Son ziyaret edilen klasörler",
            font=('Segoe UI', 10, 'bold')
        )
        title_label.pack(pady=5)

        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.path_rows = []
        self.update_list()

    def create_path_row(self, path):
        row_frame = tk.Frame(self.scrollable_frame)
        row_frame.pack(fill=tk.X, padx=5, pady=2)

        display_path = path.replace(' ', '_')
        if len(display_path) > 50:
            display_path = display_path[:47] + "..."

        open_button = tk.Button(
            row_frame,
            text="Aç",
            command=lambda p=path: self.open_folder(p),
            font=('Segoe UI', 9),
            width=6
        )
        open_button.pack(side=tk.LEFT, padx=(0, 5))

        try:
            path_label = tk.Label(
                row_frame,
                text=display_path,
                font=('Segoe UI', 9),
                anchor='w',
                padx=5
            )
            path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            def show_tooltip(event):
                tooltip = tk.Toplevel()
                tooltip.wm_overrideredirect(True)
                tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
                
                frame = tk.Frame(
                    tooltip,
                    borderwidth=1,
                    relief='solid',
                    background='#ffffe0'
                )
                frame.pack(fill='both', expand=True)
                
                tip_label = tk.Label(
                    frame, 
                    text=path,
                    justify=tk.LEFT,
                    background="#ffffe0",
                    font=('Segoe UI', 9),
                    padx=5,
                    pady=2,
                    wraplength=400
                )
                tip_label.pack()
                
                def hide_tooltip(event):
                    tooltip.destroy()
                
                path_label.bind('<Leave>', hide_tooltip)
                tooltip.bind('<Leave>', hide_tooltip)
            
            if len(path) > 50:
                path_label.bind('<Enter>', show_tooltip)

        except:
            pass
            
        return row_frame

    def open_folder(self, path):
        try:
            os.startfile(path)
        except:
            try:
                subprocess.run(f'explorer "{path}"', shell=True)
            except:
                pass

    def update_list(self):
        for row in self.path_rows:
            row.destroy()
        self.path_rows.clear()

        paths = self.path_queue.get_paths()
        for path in reversed(paths):
            row = self.create_path_row(path)
            self.path_rows.append(row)

def main():
    tracker = ExplorerTracker()
    tracker.start()
    
    gui = GUI(tracker.path_queue)
    
    try:
        gui.root.mainloop()
    finally:
        tracker.stop()

if __name__ == "__main__":
    main()