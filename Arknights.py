import win32con as con
import win32gui as gui
import pyautogui as agui
import time
import xlrd
import tkinter as tki
import threading
from tkinter import messagebox
import tkinter.font as tf
import os
import sys


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# 定义常量
excel_file = resource_path("arknights_pos.xlsx")
program_name = "明日方舟"


def thread_it(func, *args):
    '''将函数放入线程中执行'''
    # 创建线程
    t = threading.Thread(target=func, args=args)
    # 守护线程
    t.setDaemon(True)
    # 启动线程
    t.start()


class Arknights:
    def __init__(self, offset=35):
        self.offset = offset  # 左上角位置偏移值
        self.pos_ul = [-1, -1]  # 窗口左上角位置
        self.resolution = [1280, 720]  # 窗口大小
        self.hwnd = -1  # 程序句柄
        self.app_list = set()
        self.view = None
        self.updateProgramState()
        self.curr_interface = 'main_interface'
        self.xlsx = xlrd.open_workbook(excel_file)

    # 回调函数
    def foo(self, hwnd, mouse):
        if gui.IsWindow(hwnd) and gui.IsWindowEnabled(hwnd) and gui.IsWindowVisible(hwnd):
            self.app_list.add(gui.GetWindowText(hwnd))

    # 更新程序状态
    def updateProgramState(self):
        # 检查程序是否启动
        self.app_list = set()
        self.hwnd = -1
        gui.EnumWindows(self.foo, 0)
        lt = [t for t in self.app_list if t]
        for t in lt:
            if t.find(program_name) >= 0:
                self.hwnd = gui.FindWindow(None, t)
                break
        # 更新位置信息
        if self.hwnd != -1:
            size = gui.GetWindowRect(self.hwnd)
            self.pos_ul[0], self.pos_ul[1] = size[0], size[1] + self.offset
            self.resolution[0], self.resolution[1] = size[2] - size[0], (size[2] - size[0]) / 16 * 9

    # 从表格中获取点的相对位置和指向界面
    def get_pos(self, sheet_name, point_name):
        sheet = self.xlsx.sheet_by_name(sheet_name)
        col0 = sheet.col(0)
        col0 = [cell.value for cell in col0]
        row_p = sheet.row(col0.index(point_name))
        pos_x, pos_y = row_p[3].value, row_p[4].value
        pos_p = (pos_x, pos_y, row_p[5].value)
        return pos_p

    # 判断当前点击是否合法
    def judge_manipulate(self, target_interface):
        # 导航栏只在主界面按不到
        if target_interface == 'navigation':
            return self.curr_interface != 'main_interface'
        # 除了导航栏，其他界面都必须严格对应
        else:
            return self.curr_interface == target_interface

    # 点击
    def click(self, interface, point, wait=1, stay_after=False):
        gui.ShowWindow(self.hwnd, con.SW_SHOWDEFAULT)
        gui.SetWindowPos(self.hwnd, con.HWND_TOPMOST, 0, 0, 0, 0, con.SWP_NOMOVE | con.SWP_NOSIZE)
        self.updateProgramState()
        if self.hwnd == -1:
            if self.view:
                self.view.curr_state.set('当前状态：程序未运行')
                return
        if self.judge_manipulate(interface):
            relative_pos = self.get_pos(interface, point)
            mouse_x, mouse_y = agui.position()
            click_x = self.resolution[0] * relative_pos[0] + self.pos_ul[0]
            click_y = self.resolution[1] * relative_pos[1] + self.pos_ul[1]
            agui.click(click_x, click_y)
            if not stay_after:
                gui.SetWindowPos(self.hwnd, con.HWND_BOTTOM, 0, 0, 0, 0, con.SWP_NOMOVE | con.SWP_NOSIZE)
                agui.moveTo(mouse_x, mouse_y)
            time.sleep(wait)
            self.curr_interface = relative_pos[2]
        else:
            if self.view:
                self.view.curr_state.set('当前状态：脚本出错，请重启程序')
            return

    def click_pos(self, relative_x, relative_y, wait=1, stay_after=False):
        if self.hwnd == -1:
            if self.view:
                self.view.curr_state.set('当前状态：程序未运行')
                return
        gui.ShowWindow(self.hwnd, con.SW_SHOWDEFAULT)
        gui.SetWindowPos(self.hwnd, con.HWND_TOPMOST, 0, 0, 0, 0, con.SWP_NOMOVE | con.SWP_NOSIZE)
        self.updateProgramState()
        mouse_x, mouse_y = agui.position()
        click_x = self.resolution[0] * relative_x + self.pos_ul[0]
        click_y = self.resolution[1] * relative_y + self.pos_ul[1]
        agui.click(click_x, click_y)
        if not stay_after:
            gui.SetWindowPos(self.hwnd, con.HWND_BOTTOM, 0, 0, 0, 0, con.SWP_NOMOVE | con.SWP_NOSIZE)
            agui.moveTo(mouse_x, mouse_y)
        time.sleep(wait)

    # 登陆
    def login_manipulation(self):
        # 进入登陆界面
        self.click_pos(0.5, 0.5)
        time.sleep(10)
        # 按下登陆按钮
        self.click_pos(0.5, 0.7)
        time.sleep(12)
        self.curr_interface = 'main_interface'
        self.click('main_interface', '签到关闭', wait=5)
        self.click('main_interface', '签到关闭', wait=3)
        self.click('main_interface', '签到关闭', wait=2)
        self.click('main_interface', '签到关闭', wait=2)
        self.click('main_interface', '签到关闭', wait=2)
        self.click('main_interface', '签到关闭', wait=2)

    # 基建收菜
    def construct_manipulation(self):
        # 进入基建界面
        if self.curr_interface == 'main_interface':
            self.click('main_interface', '基建')
        else:
            self.click(self.curr_interface, '导航')
            self.click(self.curr_interface, '基建')
        time.sleep(3)
        # 收菜
        gui.ShowWindow(self.hwnd, con.SW_SHOWDEFAULT)
        gui.SetWindowPos(self.hwnd, con.HWND_TOPMOST, 0, 0, 0, 0, con.SWP_NOMOVE | con.SWP_NOSIZE)
        color_noti_grey = (222, 222, 222)
        color_noti_red = (203, 77, 84)
        color_noti_blue = (47, 168, 223)
        color_grb = [color_noti_grey, color_noti_red, color_noti_blue]
        pos_upper = self.get_pos('construction', 'EMERGENCY')
        click_x = self.resolution[0] * pos_upper[0] + self.pos_ul[0]
        click_y = self.resolution[1] * pos_upper[1] + self.pos_ul[1]
        color_upper = agui.screenshot().getpixel((click_x, click_y))
        color_differ = []
        for color in color_grb:
            diff = 0
            for i in range(3):
                diff += abs(color[i] - color_upper[i])
            color_differ.append(diff)
        pos_noti = 'EMERGENCY'
        if color_differ.index(min(color_differ)) == 1:  # red
            pos_noti = 'NOTIFICATION'
        elif color_differ.index(min(color_differ)) == 0:  # grey
            pos_noti = None
        # 判断主页是否有菜可收
        if pos_noti:  # 有菜可以收
            self.click('construction', pos_noti)
            for i in range(3):
                self.click('construction', "待办事项_1", stay_after=True)
        # 会客室收菜
        self.click('construction', '会客室', wait=2)
        self.click('saloon', '会客2_入口')
        self.click('saloon', '会客2_自产线索')
        self.click('saloon', '会客2_自产领取', wait=2)
        self.click('saloon', '会客2_自产后退')
        self.click('saloon', '会客2_接受线索')
        self.click('saloon', '会客2_接受领取')
        self.click('saloon', '会客2_接受后退')
        self.click('saloon', '会客2_传递线索', wait=2)
        for i in range(5):
            self.click('saloon', '会客2_传递选中')
            self.click('saloon', '会客2_传递赠送')
        self.click('saloon', '会客2_传递退出')
        self.click('saloon', '导航')
        self.click('navigation', '首页', wait=3)

    # 游戏循环
    def game_cycle(self, round, time_per_round=140):
        for i in range(round):
            if self.view:
                self.view.curr_state.set('当前状态：游戏循环（第%d轮）' % (i + 1))
            time.sleep(2)
            self.curr_interface = 'after_select'
            time.sleep(1)
            self.click('after_select', '开始')
            time.sleep(1)
            self.click('after_select', '选人界面开始')
            gui.ShowWindow(self.hwnd, con.SW_MINIMIZE)
            time.sleep(time_per_round)
            self.click_pos(0.86, 0.4)
            time.sleep(3)
            self.click_pos(0.86, 0.4)
            time.sleep(4)
        if self.view:
            self.view.curr_state.set('当前状态：正在清空任务...')
        self.clear_task()
        if self.view:
            self.view.curr_state.set('当前状态：待命')

    # 购物
    def store_manipulation(self):
        # 进入商店界面
        if self.curr_interface == 'main_interface':
            self.click('main_interface', '采购', wait=2)
        else:
            self.click(self.curr_interface, '导航')
            self.click(self.curr_interface, '采购', wait=2)
        # 买东西
        self.click('store', '信用交易所')
        self.click('store', '领取信用')
        self.click('store', '购买返回')
        for i in range(6):
            self.click('store', '购买物品' + str(i + 1))
            self.click('store', '购买确认')
            self.click('store', '购买返回')
            self.click('store', '购买返回')
        self.click('store', '导航')
        self.click('navigation', '首页', wait=3)

    def clear_task(self):
        # 进入任务界面
        if self.curr_interface == 'main_interface':
            self.click('main_interface', '任务', wait=2)
        else:
            self.click(self.curr_interface, '导航')
            self.click(self.curr_interface, '任务', wait=2)
        for i in range(25):
            self.click('task', '领取1')
        self.click('task', '周常')
        for i in range(25):
            self.click('task', '领取1')
        self.click('task', '导航')
        self.click('navigation', '首页', wait=2)

    # 一键收菜
    def one_step(self):
        if self.view:
            self.view.curr_state.set('当前状态：正在初始化...')
        time.sleep(2)
        if self.view:
            self.view.curr_state.set('当前状态：正在登录...')
        self.login_manipulation()
        if self.view:
            self.view.curr_state.set('当前状态：正在基建收菜...')
        self.construct_manipulation()
        if self.view:
            self.view.curr_state.set('当前状态：正在采购...')
        self.store_manipulation()
        if self.view:
            self.view.curr_state.set('当前状态：待命')


class ArkController:
    def __init__(self, ark: 'Arknights', view: 'ArkView'):
        self.ark = ark
        self.view = view


class ArkView:

    def __init__(self, ark: 'Arknights'):
        self.ark = ark
        self.ark.view = self
        self.root = tki.Tk()
        self.root.title("ArknightsHelper")
        self.curr_state = tki.StringVar()
        self.curr_state.set('当前状态：待命')
        self.ark.updateProgramState()
        ww = 400
        wh = 300
        x = (self.root.winfo_screenwidth() - ww) / 2
        y = (self.root.winfo_screenheight() - wh) / 2
        self.root.geometry("%dx%d+%d+%d" % (ww, wh, x, y))
        self.controller = None
        label_welcome = tki.Label(self.root, text="明日方舟小助手", font=tf.Font(family="微软雅黑", size=25))
        label_welcome.pack()
        button_game_cycle = tki.Button(self.root, text="游戏循环", command=self.game_cycle_action)
        button_game_cycle.pack()
        button_one_step = tki.Button(self.root, text="OneStep", command=self.one_step_action)
        button_one_step.pack()
        label_curr_state = tki.Label(self.root, textvariable=self.curr_state)
        label_curr_state.pack()

    def game_cycle_action(self):
        pop_window = tki.Toplevel()
        pop_window.title("输入信息")
        ww = 300
        wh = 200
        x = (pop_window.winfo_screenwidth() - ww) / 2
        y = (pop_window.winfo_screenheight() - wh) / 2
        pop_window.geometry("%dx%d+%d+%d" % (ww, wh, x, y))
        pop_window.wm_attributes("-topmost", 1)
        tki.Label(pop_window, text='刷几轮：').pack()
        txt1 = tki.Entry(pop_window)
        txt1.pack()
        tki.Label(pop_window, text='一轮多久：').pack()
        txt2 = tki.Entry(pop_window)
        txt2.pack()

        def conf():
            t1 = txt1.get()
            t2 = txt2.get()
            if self.curr_state.get() == '当前状态：待命':
                if t1 == "" or t2 == "":
                    tki.messagebox.showwarning("提示", "输入不合法")
                else:
                    thread_it(self.ark.game_cycle, int(t1), int(t2))
                    pop_window.destroy()
            else:
                tki.messagebox.showwarning('提示', "脚本正在运行")

        tki.Button(pop_window, text="开始刷", command=conf).pack()
        pop_window.mainloop()

    def one_step_action(self):
        if self.curr_state.get() == '当前状态：待命':
            thread_it(self.ark.one_step)
        else:
            tki.messagebox.showwarning('提示', "脚本正在运行")

    def root_run(self):
        self.root.mainloop()


if __name__ == '__main__':
    ark_back = Arknights()
    ark_front = ArkView(ark_back)
    ark_front.root_run()
