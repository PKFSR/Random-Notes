import threading
import time
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import *
from ttkbootstrap import Style
import jwt
import loguru
import requests
import ttkbootstrap as ttk
import psutil
import win32api
import win32con
import win32gui
import win32process


class WinGUI(Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.tk_label_lya5vl8j = self.__tk_label_lya5vl8j(self)
        self.tk_input_lya5w9xh = self.__tk_input_lya5w9xh(self)
        self.tk_select_box_lya5weto = self.__tk_select_box_lya5weto(self)
        self.tk_label_lya5wr9g = self.__tk_label_lya5wr9g(self)
        self.tk_button_lya5x9r0 = self.__tk_button_lya5x9r0(self)
        self.tk_button_lya6ck06 = self.__tk_button_lya6ck06(self)

        self.procedure_dict = dict()
        self.get_window_procedure()

    def __win(self):
        self.title("窗口自动排版工具1.0")
        # 设置窗口大小、居中
        width = 497
        height = 250
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)

        self.resizable(width=False, height=False)

    def scrollbar_autohide(self, vbar, hbar, widget):
        """自动隐藏滚动条"""

        def show():
            if vbar: vbar.lift(widget)
            if hbar: hbar.lift(widget)

        def hide():
            if vbar: vbar.lower(widget)
            if hbar: hbar.lower(widget)

        hide()
        widget.bind("<Enter>", lambda e: show())
        if vbar: vbar.bind("<Enter>", lambda e: show())
        if vbar: vbar.bind("<Leave>", lambda e: hide())
        if hbar: hbar.bind("<Enter>", lambda e: show())
        if hbar: hbar.bind("<Leave>", lambda e: hide())
        widget.bind("<Leave>", lambda e: hide())

    def v_scrollbar(self, vbar, widget, x, y, w, h, pw, ph):
        widget.configure(yscrollcommand=vbar.set)
        vbar.config(command=widget.yview)
        vbar.place(relx=(w + x) / pw, rely=y / ph, relheight=h / ph, anchor='ne')

    def h_scrollbar(self, hbar, widget, x, y, w, h, pw, ph):
        widget.configure(xscrollcommand=hbar.set)
        hbar.config(command=widget.xview)
        hbar.place(relx=x / pw, rely=(y + h) / ph, relwidth=w / pw, anchor='sw')

    def create_bar(self, master, widget, is_vbar, is_hbar, x, y, w, h, pw, ph):
        vbar, hbar = None, None
        if is_vbar:
            vbar = Scrollbar(master)
            self.v_scrollbar(vbar, widget, x, y, w, h, pw, ph)
        if is_hbar:
            hbar = Scrollbar(master, orient="horizontal")
            self.h_scrollbar(hbar, widget, x, y, w, h, pw, ph)
        self.scrollbar_autohide(vbar, hbar, widget)

    def __tk_label_lya5vl8j(self, parent):
        label = Label(parent, text="排版布局", anchor="center", )
        label.place(x=20, y=20, width=62, height=30)
        return label

    def __tk_input_lya5w9xh(self, parent):
        ipt = Entry(parent, )
        ipt.place(x=100, y=20, width=56, height=30)
        return ipt

    def __tk_select_box_lya5weto(self, parent):
        cb = Combobox(parent, state="readonly", )
        cb.place(x=245, y=20, width=150, height=30)
        return cb

    def __tk_label_lya5wr9g(self, parent):
        label = Label(parent, text="排版程序：", anchor="center", )
        label.place(x=170, y=20, width=67, height=30)
        return label

    def __tk_button_lya5x9r0(self, parent):
        btn = Button(parent, text="一键排版", takefocus=False, command=self.start_typesetting)
        btn.place(x=171, y=110, width=155, height=69)
        return btn

    def __tk_button_lya6ck06(self, parent):
        btn = Button(parent, text="刷新程序列表", takefocus=False, command=self.reget_window_procedure)
        btn.place(x=405, y=20, width=80, height=30)
        return btn

    @staticmethod
    def get_all_window_handles():
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                hwnds.append(hwnd)
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds

    @staticmethod
    def get_window_process_info(hwnd):
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            process = psutil.Process(pid)
            process_name = process.name()
            exe_path = process.exe()
            return process_name, exe_path
        except psutil.NoSuchProcess:
            return None, None

    @staticmethod
    def get_screen_size():
        width = win32api.GetSystemMetrics(0)
        height = win32api.GetSystemMetrics(1)
        return width, height

    def get_window_procedure(self):
        hwnds = self.get_all_window_handles()
        for hwnd in hwnds:
            process_name, exe_path = self.get_window_process_info(hwnd)
            self.procedure_dict[hwnd] = process_name
        self.tk_select_box_lya5weto['values'] = tuple(list(set(self.procedure_dict.values())))

    def reget_window_procedure(self):
        self.procedure_dict.clear()
        self.get_window_procedure()
        self.tk_select_box_lya5weto['values'] = tuple(list(set(self.procedure_dict.values())))

    def auto_typesetting(self, win_count, xy_list):
        while 1:
            hw_idx = 0
            self.reget_window_procedure()
            for hwnd, process_name in self.procedure_dict.items():
                if process_name == self.tk_select_box_lya5weto.get():
                    if hw_idx >= win_count:
                        continue
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # 恢复窗口（如果最小化）
                    win32gui.SetForegroundWindow(hwnd)  # 设置为前台窗口
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, xy_list[hw_idx][0], xy_list[hw_idx][1],
                                          xy_list[hw_idx][2], xy_list[hw_idx][3], win32con.SWP_SHOWWINDOW)  # 移动窗口
                    hw_idx += 1
            time.sleep(3)

    def start_typesetting(self):
        try:
            layout_type = self.tk_input_lya5w9xh.get()
            w_count, h_count = layout_type.split("x")
            w_count, h_count = int(w_count), int(h_count)
        except:
            messagebox.showerror("错误", "布局设置错误，格式应为：宽x高")
            return
        if self.tk_button_lya5x9r0.cget("text") == "布局刷新中":
            messagebox.showerror("错误", "布局刷新中，请勿重复操作")
            return
        v = self.tk_select_box_lya5weto.get()
        if not v:
            messagebox.showerror("错误", "请选择排版程序")
            return
        w, h = self.get_screen_size()
        w_unit = w / w_count
        h_unit = h / h_count
        # 计算每个单元格的坐标和宽高
        one_w = int(w / w_count)
        one_h = int(h / h_count)
        xy_list = list()
        for i in range(w_count):
            for j in range(h_count):
                x = int(i * w_unit)
                y = int(j * h_unit)
                xy_list.append((x, y, one_w, one_h))
        win_count = w_count * h_count
        self.tk_button_lya5x9r0.config(text="布局刷新中")
        run_info_thread = threading.Thread(target=self.auto_typesetting, args=(win_count, xy_list), daemon=True)
        run_info_thread.start()


if __name__ == "__main__":
    win = WinGUI()
    win.mainloop()
