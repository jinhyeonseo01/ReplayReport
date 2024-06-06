import os
import asyncio
import logging
import traceback
import time
import math
import threading

import json
import PIL
import tkinter
from tkinter import filedialog
import customtkinter
import pynput.keyboard

import win32con
import win32api
import win32gui

import mimetypes

from PIL import ImageGrab, ImageTk, Image

import win
import smtplib
import email
from email.message import EmailMessage


from pynput import keyboard, mouse

from queue import Queue, Empty


#pyautogui
#customtkinter
#pillow
#pynput
#pywin32


logging.basicConfig(
    level=logging.DEBUG
)

targetFrame = 60
targetDeltaFrame = 1000/targetFrame
uiDataQueue = Queue()

nowTime = time.perf_counter_ns()
prevTime = nowTime
deltaTime = targetDeltaFrame

defaultEvent = {'type':"None"}

lock = threading.Lock()




def main():
    global nowTime
    window = None
    try:
        window = Window()

        window.frame()
        window.mainloop()
    except Exception as e:
        logging.error(f"비정상종료")
        logging.error(traceback.format_exc())
        window.onClose()
        window.frame()
    pass


processTitle = ""
processImage : Image = None
processImageTK : ImageTk = None
processImage_hOffset = 450
processUpdateDeltaTime = 0.3
processInputList = []
processOffset = (0,0)
processOffsetTime = time.perf_counter_ns()

recordData = {
    'processTitle' : processTitle,
    'processInputList' : []
}

playStartTime = time.perf_counter_ns()
playList = []

def processImageGet():
    global processImage, lock, processImageTK

    with lock:
        if(processImage == None):
            return None
        processImageTK = ImageTk.PhotoImage(processImage)
        return processImageTK
    return None


def processImageSet():
    global processImage, lock, processImageTK, processImage_hOffset

    with lock:
        hwnd = win.find_hwnd_by_title(processTitle)
        if(hwnd != 0):
            left, top, right, bottom = win.get_window_rect(hwnd)
            left += 2
            right -= 6
            bottom -= 6
            if processImage != None:
                processImage.close()
            processImage = ImageGrab.grab(bbox=(left, top, right, bottom))
            w = right - left
            h = bottom - top
            processImage_hOffset = 450
            processImage = processImage.resize((int(processImage_hOffset*(w/h)), int(processImage_hOffset)))


def remove_tk(widget):

    for child in widget.winfo_children():
        remove_tk(child)

    widget.grid_forget()
    widget.pack_forget()
    widget.place_forget()
    widget.destroy()

def forgetAll(widget):
    for child in widget.winfo_children():
        forgetAll(child)
    forget(widget)

def forget(widget):
    widget.grid_forget()
    widget.pack_forget()
    widget.place_forget()

class Window(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.protocol("WM_DELETE_WINDOW", self.onClose)
        self.bind('<Escape>', lambda e: self.onClose())
        self.title("Replay Report")
        self.defualtSize = (1280, 720)
        self.geometry(
            f"{self.defualtSize[0]}x{self.defualtSize[1]}+{int(self.winfo_screenwidth() / 2 - self.defualtSize[0] / 2)}+{int(self.winfo_screenheight() / 2 - self.defualtSize[1] / 2)}(pixel)")

        self.bind("<KeyPress>",
                  lambda event: uiDataQueue.put({**defaultEvent, **{"type":"Input", 'press':"Down", "event":event}}))
        self.bind("<KeyRelease>",
                  lambda event: uiDataQueue.put({**defaultEvent, **{"type":"Input", 'press':"Up", "event":event}}))


        uiDataQueue.put({**defaultEvent, **{"type": "Init"}})
        uiDataQueue.put({**defaultEvent, **{"type": "SetPage_Main"}})

        self.pageDefualt = {
            'name':"None",
            'first':True,
            'root':None,
            'init':lambda window, page:{logging.debug(f"{page['name']} 초기화 되지 않은 초기화 함수")},
            'disable': lambda window, page, next: {logging.debug(f"{page['name']} 초기화 되지 않은 비활성화 함수")}
        }

        self.pageList = []
        self.pageIndex = -1

        self.timeList = []

        def ImageHook():
            while not self.processImageCapture.stop:
                global processUpdateDeltaTime
                processImageSet()
                time.sleep(processUpdateDeltaTime)

        self.processImageCapture = threading.Thread(target=ImageHook, daemon=True)
        self.processImageCapture.stop = False

    def onClose(self):
        uiDataQueue.put({**defaultEvent, **{"type": "Close"}})
        pass

    def update(self, data):
        if data['type'] == "Init":
            logging.debug("Window Init")

            customtkinter.set_appearance_mode("Light")
            customtkinter.set_default_color_theme("green")

            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=0)
            self.grid_rowconfigure(1, weight=1)

            self.menu = customtkinter.CTkFrame(self,
                                               corner_radius=0,
                                               fg_color="transparent")
            self.menu.grid(row=0, column=0, sticky="nsew")
            
            self.menu_file_buttom = customtkinter.CTkButton(self.menu,
                                                            corner_radius=0,
                                                            text='새로 만들기',
                                                            width=80,
                                                            text_color="black",
                                                            fg_color="transparent",
                                                            command=lambda :self.NewRecord())
            self.menu_file_buttom.grid(row=0, column=0, sticky="ns")
            self.menu_view_buttom = customtkinter.CTkButton(self.menu,
                                                            corner_radius=0,
                                                            text='불러오기',
                                                            width=80,
                                                            text_color="black",
                                                            fg_color="transparent",
                                                            command=lambda :self.LoadRecord())
            self.menu_view_buttom.grid(row=0, column=1, sticky="ns")
            self.menu_option_buttom = customtkinter.CTkButton(self.menu,
                                                            corner_radius=0,
                                                            text='저장',
                                                            width=80,
                                                            text_color="black",
                                                            fg_color="transparent",
                                                            command=lambda :self.SaveRecord())
            self.menu_option_buttom.grid(row=0, column=2, sticky="ns")

            self.menu_option_buttom = customtkinter.CTkButton(self.menu,
                                                              corner_radius=0,
                                                              text='전송',
                                                              width=80,
                                                              text_color="black",
                                                              fg_color="transparent",
                                                              command=lambda :uiDataQueue.put({**defaultEvent, **{"type": "SetPage_Send"}}))
            self.menu_option_buttom.grid(row=0, column=3, sticky="ns")
            self.menu_help_buttom = customtkinter.CTkButton(self.menu,
                                                            corner_radius=0,
                                                            text='도움말',
                                                            width=80,
                                                            text_color="black",
                                                            fg_color="transparent")
            self.menu_help_buttom.grid(row=0, column=4, sticky="ns")

            #options = [i[1] for i in win.getProcessInfoList()]
            #option_menu = customtkinter.CTkOptionMenu(self, values=options)
            #option_menu.pack(padx=20, pady=20)

            #print(win.getProcessInfoList())
            #print(win.find_hwnd_by_title("계산기"))
            #left, top, right, bottom = win.get_window_rect(win.find_hwnd_by_title("계산기"))

            #self.canvas = customtkinter.CTkCanvas(width=right-left,height=bottom-top)
            #self.canvas.pack()
            #customtkinter.CTkLabel(image=ImageTk.Image())
            #win.send_key_event(win.find_window_handle('*asd - 메모장'), 65)

            #hid = win.find_window_handle('C Test - Microsoft Visual Studio')
            #win32gui.ShowWindow(hid, win32con.SW_RESTORE)
            #win32gui.SetForegroundWindow(hid)
            #time.sleep(0.1)
            #win.send_key_event(hid,  win32con.VK_RETURN)
            #self.img_tk = None

            def InitMain(window, page):
                global processTitle

                root : customtkinter.CTkFrame = page['root']
                print('initMain')

                self.isPlay = False
                self.isRecording = False
                self.isPause = False


                root.grid_columnconfigure(0, weight=1)
                root.grid_rowconfigure(0, weight=0)
                root.grid_rowconfigure(1, weight=0)

                root.topFrame = customtkinter.CTkFrame(root,
                                                       corner_radius=0,
                                                       fg_color='transparent')
                root.topFrame.grid(row=0, column=0, sticky="nsew")
                root.topFrame.grid_columnconfigure((0),weight=1)


                root.processFrame = customtkinter.CTkFrame(root.topFrame,
                                                        corner_radius=0,
                                                        fg_color='transparent')
                root.processFrame.grid(row=0, column=0, pady=(10,10), sticky="nsew")


                processLabel = customtkinter.CTkLabel(root.processFrame,
                                                      text="프로세스 선택",
                                                      font=customtkinter.CTkFont(size=14))
                processLabel.pack(side=customtkinter.LEFT, padx = (15,10), expand=False)

                options = [i[1] for i in win.getProcessInfoList()]
                options.sort()
                self.option_var = customtkinter.StringVar(root, options[0])
                processTitle = options[0]
                def optionUpdate(e,v,s):
                    global processTitle
                    processTitle = self.option_var.get()
                    print(processTitle)
                self.option_var.trace_add("write", optionUpdate)
                process_UI = customtkinter.CTkOptionMenu(root.processFrame,
                                                        variable=self.option_var,
                                                        width=400,
                                                        values=options)
                process_UI.pack(side=customtkinter.LEFT, expand=False)

                reload = customtkinter.CTkButton(root.processFrame,
                                               text="",
                                               image=customtkinter.CTkImage(
                                               light_image=Image.open('./Images/reload.png')),
                                               width=50,
                                               command=lambda: process_UI.configure(values=sorted([i[1] for i in win.getProcessInfoList()])))

                reload.pack(side=customtkinter.LEFT, padx=(10))

                tempText = customtkinter.CTkLabel(root.processFrame,
                                                              text="레코드 : ",
                                                              font=customtkinter.CTkFont(size=14))
                tempText.pack(side=customtkinter.LEFT, padx=(20,0))

                self.recordStateText = customtkinter.CTkLabel(root.processFrame,
                                                          text="정지",
                                                          font=customtkinter.CTkFont(size=14))
                self.recordStateText.pack(side=customtkinter.LEFT, padx=(0,0))

                #------------------------------------------------------------------------------------------------

                root.recordFrame = customtkinter.CTkFrame(root.topFrame,
                                                        corner_radius=0,
                                                        fg_color='transparent')
                root.recordFrame.grid(row=1, column=0, sticky="nsew")
                root.recordFrame2 = customtkinter.CTkFrame(root.recordFrame,
                                                        corner_radius=0,
                                                        fg_color='transparent')
                root.recordFrame2.pack(side=customtkinter.TOP)


                stop = customtkinter.CTkButton(root.recordFrame2,
                                               text="",
                                               image=customtkinter.CTkImage(
                                                   light_image=Image.open('./Images/stop.png')),
                                               width=50,
                                               command=lambda: self.Stop())

                stop.pack(side=customtkinter.LEFT, padx=(2,2))


                record = customtkinter.CTkButton(root.recordFrame2,
                                                text='',
                                                image=customtkinter.CTkImage(light_image=Image.open('./Images/record.png')),
                                                width=50,
                                                command=lambda: self.Record())

                record.pack(side=customtkinter.LEFT, padx=(2,2))


                play = customtkinter.CTkButton(root.recordFrame2,
                                                text='',
                                                image=customtkinter.CTkImage(light_image=Image.open('./Images/play.png')),
                                                width=50,
                                                command=lambda: self.Play())

                play.pack(side=customtkinter.LEFT, padx=(2,2))



                pause = customtkinter.CTkButton(root.recordFrame2,
                                                text='',
                                                image=customtkinter.CTkImage(light_image=Image.open('./Images/pause.png')),
                                                width=50,
                                                command=lambda: self.Pause())

                pause.pack(side=customtkinter.LEFT, padx=(2,2))

                #----------------------------------------------------------------

                root.processImage = customtkinter.CTkCanvas(root,
                                                            width=600,
                                                            height=400)
                root.processImage.grid(row=1, column=0,pady=(5,5))

                # ----------------------------------------------------------------

                self.timelineText = customtkinter.CTkLabel(root,
                                                        text="타임라인",
                                                        anchor='w',
                                                        font=customtkinter.CTkFont(size=18))
                self.timelineText.grid(row=2, column=0, padx=(20,20),pady=(5, 5), sticky="ew")

                self.timeline = customtkinter.CTkScrollableFrame(root,
                                                            corner_radius=0,
                                                            height=80,
                                                            orientation="horizontal")
                self.timeline.grid(row=3,column=0,pady=(5,5), sticky="ew")

                self.testTimeLine = customtkinter.CTkFrame(self.timeline,
                                                           width=5,
                                                           corner_radius=0,
                                                           fg_color='black')
                self.testTimeLine.pack(padx=(0,0),side=customtkinter.LEFT)
                self.timelineTKList = []

            def InitSend(root, page):
                root = page['root']
                print('initSecond')
                root.grid_columnconfigure(0, weight=1)
                root.grid_rowconfigure(0, weight=1)
                self.sendFrame = customtkinter.CTkFrame(root,
                                                        corner_radius=10,
                                                        width=500,
                                                        height=400)
                self.sendFrame.grid(row=0, column=0, pady=(100,100),sticky="ns")

                self.emailFrame = customtkinter.CTkFrame(self.sendFrame,
                                                        corner_radius=10,
                                                         fg_color='transparent',
                                                        width=400)
                self.emailFrame.grid(row=0, column=0,pady=(50,10),padx=(10,10),sticky="ew")

                self.emailText = customtkinter.StringVar(value="xxx@gmail.com")

                self.emailLabal = customtkinter.CTkLabel(self.emailFrame,
                                                         text="이메일 : ",
                                                         fg_color='transparent',
                                                         width= 40,
                                                         anchor='e',
                                                         font=customtkinter.CTkFont(size=14))
                self.emailLabal.grid(row=0, column=0, sticky="ew")
                self.emailEntry = customtkinter.CTkEntry(self.emailFrame,
                                                         textvariable=self.emailText,
                                                         width=300)
                self.emailEntry.grid(row=0,column=1,sticky="ew")

                self.ttFrame = customtkinter.CTkFrame(self.sendFrame,
                                                         corner_radius=10,
                                                         fg_color='transparent',
                                                         width=400)
                self.ttFrame.grid(row=1, column=0, pady=(20, 10), padx=(10, 10), sticky="ew")

                self.ttText = customtkinter.StringVar(value="문제사항")

                self.ttLabal = customtkinter.CTkLabel(self.ttFrame,
                                                         text="제목 : ",
                                                         fg_color='transparent',
                                                         width=40,
                                                         anchor='e',
                                                         font=customtkinter.CTkFont(size=14))
                self.ttLabal.grid(row=0, column=0, sticky="ew")
                self.ttEntry = customtkinter.CTkEntry(self.ttFrame,
                                                         textvariable=self.ttText,
                                                         width=300)
                self.ttEntry.grid(row=0, column=1, sticky="ew")

                self.contextText = customtkinter.StringVar(value="사유")
                self.contextFrame = customtkinter.CTkFrame(self.sendFrame,
                                                         corner_radius=10,
                                                        fg_color='transparent',
                                                         width=400)
                self.contextFrame.grid(row=2, column=0, pady=(10, 10), padx=(10, 10), sticky="ew")
                self.contextLabal = customtkinter.CTkLabel(self.contextFrame,
                                                          text="내용 : ",
                                                          fg_color='transparent',
                                                          width=40,
                                                          anchor='e',
                                                          font=customtkinter.CTkFont(size=14))
                self.contextLabal.grid(row=0, column=0, sticky="ew")
                self.contextTextBox = customtkinter.CTkTextbox(self.contextFrame,
                                                               width=300)
                self.contextTextBox.grid(row=0, column=1, sticky="ew")

                buttom = customtkinter.CTkButton(self.sendFrame,
                                 text="전송하기",
                                 command=lambda :self.SendEmail(self.emailText.get(), self.ttText.get(), self.contextTextBox.get("1.0", customtkinter.END)))
                buttom.grid(row=3,pady=(10,10))
                buttom2 = customtkinter.CTkButton(self.sendFrame,
                                                 text="돌아가기",
                                                 command=lambda :uiDataQueue.put({**defaultEvent, **{"type": "SetPage_Main"}}))
                buttom2.grid(row=4, pady=(10, 10))


            mainPage = self.pageDefualt.copy()
            mainPage['name'] = 'Main'
            mainPage['init'] = InitMain
            self.pageList.append(mainPage)

            sendPage = self.pageDefualt.copy()
            sendPage['name'] = 'Send'
            sendPage['init'] = InitSend
            self.pageList.append(sendPage)


            for i in self.pageList:
                frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
                i['root'] = frame
            for i in self.pageList:
                i['init'](self, i)

            pass
        if data['type'] == "Update":
            global processTitle
            nowTime = time.perf_counter_ns()
            if(self.pageIndex == self.FindPageIndexByName('Main')):
                if ((nowTime - self.mainImageUpdateTime) / (10**9)) >= processUpdateDeltaTime:
                    #self.processImageCapture
                    #customtkinter.CTkCanvas().configure(image=)
                    image = processImageGet()
                    if(image != None):
                        canvas:customtkinter.CTkCanvas = self.FindPageByName('Main')['root'].processImage
                        canvas.configure(width=image.width(), height=image.height())
                        canvas.image = image

                        canvas.create_image(0, 0, anchor=customtkinter.NW,image=image)
                    self.mainImageUpdateTime = nowTime

                if(self.isPlay and (not self.isPause)): #재생
                    nowTime = time.perf_counter_ns()
                    deltaTime = nowTime - playStartTime
                    #print(deltaTime/(10**9))
                    if(len(playList) == 0):
                        self.Stop()
                    else:
                        if(len(playList) != 0):
                            print(playList[0]['time']/(10**9))
                        while(len(playList) != 0 and deltaTime >= playList[0]['time']): # 이부분 while로
                            print(playList[0])
                            if (playList[0]['type'] == 'Mouse' and playList[0]['event'] == 'Move'):
                                dataList = playList[0]['data']
                                if(len(dataList) != 0):
                                    dataTime = dataList[0][2]
                                    if (deltaTime >= dataTime):
                                        newData = playList[0].copy()
                                        newData['data'] = dataList[0]
                                        win.send_input(newData, processOffset)

                                        dataList.remove(dataList[0])
                                    else:
                                        break
                                else:
                                    playList.remove(playList[0])
                                pass
                            else:
                                print("실행")
                                newData = playList[0]
                                win.send_input(newData, processOffset)
                                playList.remove(playList[0])


                if self.isRecording and (not self.isPause):
                    for i in processInputList:
                        if (i['type'] == 'Keyboard' and i['pressed'] == "Down"):
                            if i['data'] == 'f12':
                                self.Stop()
                                break
                        if(not any(j[0] == i for j in self.timelineTKList)):
                            text = "None"
                            pady = (5, 5)
                            if(i['type'] == 'Keyboard'):
                                text = f"Key\n{i['pressed']}"
                            elif(i['type'] == 'Mouse'):
                                if (i['event'] == 'Click'):
                                    text = f"M\n{i['pressed']}"
                                if (i['event'] == 'Move'):
                                    text = f"M\nMove"
                                    pady = (15,15)
                                if (i['event'] == 'Scroll'):
                                    text = f"M\nScroll"

                            temp = customtkinter.CTkButton(self.timeline,
                                                            text=text,
                                                            width=20,
                                                            corner_radius=3,
                                                            )

                            temp.pack(padx=(1, 1), pady=pady, fill='y', side=customtkinter.LEFT)
                            self.timelineTKList.append((i, temp))

                    for i in self.timelineTKList:
                        if type(i[0]['data']) == list:
                            w = 20+(3*(len(i[0]['data'])//30))
                            if i[1].cget("width") != w:
                                i[1].configure(width=w)
                    pass
                else:
                    processInputList.clear()
            if (self.pageIndex == self.FindPageIndexByName('Send')):
                pass

        if data['type'] == 'Input':
            logging.debug(f'특수문자 {data['press']} : {data['event'].keysym}')

        if data['type'] == 'Close':
            hookMouseThread.stop()
            hookKeyboardThread.stop()
            if self.processImageCapture.is_alive():
                self.processImageCapture.stop = True
            logging.debug("Window Close")
            self.destroy()
        if data['type'] == "SetPage_Main":
            self.pageIndex = self.FindPageIndexByName('Main')
            page = self.FindPageByIndex(self.pageIndex)
            self.ChangePage(page)
            if not self.processImageCapture.is_alive():
                self.processImageCapture.stop = False
                self.processImageCapture.start()
            self.mainImageUpdateTime = time.perf_counter_ns()
            if page['first']:
                #page.
                page['first'] = False

        if data['type'] == "SetPage_Temp":
            self.pageIndex = self.FindPageIndexByName('Temp')
            page = self.FindPageByIndex(self.pageIndex)
            self.ChangePage(page)
            if page['first']:
                page['first'] = False

        if data['type'] == "SetPage_Send":
            self.pageIndex = self.FindPageIndexByName('Send')
            page = self.FindPageByIndex(self.pageIndex)
            self.ChangePage(page)
            if page['first']:
                page['first'] = False

    def Stop(self):
        if (self.isPlay or self.isRecording):
            if (self.isPlay):
                # 재생중이었을때
                pass
            if (self.isRecording):
                print(processInputList)
                while(processInputList[0]['type'] == 'Mouse' and processInputList[0]['event'] == 'Click'):
                    processInputList.remove(processInputList[0])
                recordData['processTitle'] = processTitle
                recordData['processInputList'] = processInputList.copy()
                pass
            pass
        self.recordStateText.configure(text="정지", text_color='black')
        self.isPlay = False
        self.isRecording = False
        self.isPause = False

    def Record(self):
        global processOffset, processOffsetTime
        if (((not self.isPlay) and (not self.isRecording)) or (self.isRecording and self.isPause)):
            self.recordStateText.configure(text="녹화", text_color='red')
            self.isRecording = True
            self.isPause = False

            for i in self.timelineTKList:
                remove_tk(i[1])
            self.timelineTKList.clear()

            processOffsetTime = time.perf_counter_ns()
            processInputList.clear()
            hwnd = win.find_hwnd_by_title(processTitle)
            if (hwnd != 0):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                rect = win.get_window_rect(hwnd)
                processOffset = (rect[0], rect[1])

        pass

    def Play(self):
        global processOffset, processOffsetTime, playStartTime, playList
        if (((not self.isPlay) and (not self.isRecording)) or (self.isPlay and self.isPause)):
            self.recordStateText.configure(text="재생", text_color='green')
            self.isPlay = True
            self.isPause = False
            playStartTime = time.perf_counter_ns()
            playList = json.loads(json.dumps(recordData['processInputList']))

            processOffsetTime = time.perf_counter_ns()
            hwnd = win.find_hwnd_by_title(processTitle)
            if (hwnd != 0):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                rect = win.get_window_rect(hwnd)
                processOffset = (rect[0], rect[1])

            # hid =  win.find_hwnd_by_title('Sannabi')
            # win32gui.ShowWindow(hid, win32con.SW_RESTORE)
            # win32gui.SetForegroundWindow(hid)
            # time.sleep(0.1)
            # win.send_key_event(hid, win32con.VK_SPACE)
        pass

    def Pause(self):
        if (self.isPlay or self.isRecording):
            self.recordStateText.configure(text="일시정지", text_color='yellow')
            self.isPause = True



    def frame(self):
        global nowTime, prevTime, deltaTime
        nowTime = time.perf_counter_ns()
        deltaTime = (nowTime - prevTime) / (10 ** 9)
        try:
            while True:
                result = uiDataQueue.get_nowait()
                self.update(result)
        except Empty:
            pass
        uiDataQueue.put({**defaultEvent, **{"type":"Update"}})
        self.after(max(0, int(targetDeltaFrame)), self.frame)
        prevTime = nowTime

    def ChangePage(self, page):
        for index, i in enumerate(self.pageList):
            if i['name'] == page['name']:
                i['root'].grid(row=1, column=0, sticky="nsew")
                i['root'].focus_set()
            else:
                forget(i['root'])
        return -1

    def FindPageByName(self, name):
        for i in self.pageList:
            if i['name'] == name:
                return i
        return None
    def FindPageByIndex(self, index):
        if(len(self.pageList) <= index):
            return None
        return self.pageList[self.pageIndex]
    def FindPageIndexByName(self, name):
        for index,i in enumerate(self.pageList):
            if i['name'] == name:
                return index
        return -1
    def FindPageIndexByPage(self, page):
        for index,i in enumerate(self.pageList):
            if i['name'] == page['name']:
                return index
        return -1


    def NewRecord(self):
        self.Stop()
        for i in self.timelineTKList:
            remove_tk(i[1])
        self.timelineTKList.clear()
        processInputList.clear()


    def LoadRecord(self):
        global processTitle, recordData
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")]
        )
        try:
            self.Stop()
            with open(file_path, 'r', encoding='utf-8') as file:
                recordData = json.load(file)
                processTitle = recordData['processTitle']
                self.option_var.set(processTitle)



                for i in self.timelineTKList:
                    remove_tk(i[1])
                self.timelineTKList.clear()
                processInputList = recordData['processInputList']

                for i in processInputList:
                    if (i['type'] == 'Keyboard' and i['pressed'] == "Down"):
                        if i['data'] == 'f12':
                            self.Stop()
                            break
                    if(not any(j[0] == i for j in self.timelineTKList)):
                        text = "None"
                        pady = (5, 5)
                        if(i['type'] == 'Keyboard'):
                            text = f"Key\n{i['pressed']}"
                        elif(i['type'] == 'Mouse'):
                            if (i['event'] == 'Click'):
                                text = f"M\n{i['pressed']}"
                            if (i['event'] == 'Move'):
                                text = f"M\nMove"
                                pady = (15,15)
                            if (i['event'] == 'Scroll'):
                                text = f"M\nScroll"

                        temp = customtkinter.CTkButton(self.timeline,
                                                        text=text,
                                                        width=20,
                                                        corner_radius=3,
                                                        )

                        temp.pack(padx=(1, 1), pady=pady, fill='y', side=customtkinter.LEFT)
                        self.timelineTKList.append((i, temp))
        except Exception as e:
            print(f"Error loading JSON file: {e}")
    def GetRecordJson(self):
        return json.dump(recordData)
        pass

    def SaveRecord(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    json.dump(recordData, file, indent=4)
                    print(f"JSON data saved successfully to {file_path}")
            except Exception as e:
                print(f"Error saving JSON file: {e}")
        pass

    def get_mime_type(self, fname):
        ctype, _ = mimetypes.guess_type(self, fname)
        if not ctype:
            ctype = 'appliciation/octect-stream'
        maintype, subtype = ctype.split('/', 1)
        return maintype, subtype

    def SendEmail(self,email, title, context):
        id = 'niclrain01@gmail.com'
        passwd = 'cyxo snsd zxcq xysk'  # this is app password
        mail_server = smtplib.SMTP('smtp.gmail.com', 587)
        mail_server.ehlo()
        mail_server.starttls()
        mail_server.login(id, passwd)
        msg = EmailMessage()
        msg['Subject'] = title
        msg['From'] = 'Record Replay'
        msg['To'] = email
        msg.set_content(context)

        fname = 'recordData.json'
        with open(fname, 'w', encoding='utf-8') as file:
            json.dump(recordData, file, indent=4)
        #maintype, subtype = self.get_mime_type(fname)
        with open(fname,'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='json', filename=fname)

        mail_server.send_message(msg)
        mail_server.quit()


def on_key_press(key:pynput.keyboard.Key):
    # 키를 누를 때 실행되는 함수
    keyString = win.ConvertKeyKeyboardToAutoGUI(key)
    processInputList.append({
        'type':"Keyboard",
        'pressed':"Down",
        'data':keyString,
        'time': time.perf_counter_ns() - processOffsetTime
    })
    #print(f'Key pressed: {key}')

def on_key_release(key):
    # 키를 누를 때 실행되는 함수
    keyString = win.ConvertKeyKeyboardToAutoGUI(key)
    print((time.perf_counter_ns() - processOffsetTime) / (10 ** 9))
    processInputList.append({
        'type': "Keyboard",
        'pressed': "Up",
        'data': keyString,
        'time': int(time.perf_counter_ns() - processOffsetTime)
    })
    #print(f'Key released: {key}')
    #keyboard.KeyCode.from_char()

def on_mouse(x, y, button, pressed):
    # 키를 누를 때 실행되는 함수
    keyString = win.ConvertKeyMouseToAutoGUI(button)
    processInputList.append({
        'type': "Mouse",
        'pressed': "Down" if pressed else "Up",
        'data': (x-processOffset[0], y-processOffset[1], keyString),
        'time':int(time.perf_counter_ns() - processOffsetTime),
        'event': "Click"
    })
    #print(f'Mouse: {button} {pressed}')
def on_move(x, y):
    if(len(processInputList) != 0 and processInputList[-1]['type'] == "Mouse" and processInputList[-1]['event']) == "Move":
        processInputList[-1]['data'].append((x-processOffset[0], y-processOffset[1], int(time.perf_counter_ns() - processOffsetTime)))
    else:
        processInputList.append({
            'type': "Mouse",
            'pressed': "Down",
            'data': [(x-processOffset[0], y-processOffset[1], int(time.perf_counter_ns() - processOffsetTime))],
            'time': int(time.perf_counter_ns() - processOffsetTime),
            'event': "Move"
        })
    pass

def on_scroll(x, y, dx, dy):
    processInputList.append({
        'type': "Mouse",
        'pressed': "Down" if dy < 0 else "Up",
        'data': (x-processOffset[0], y-processOffset[1]),
        'time':int(time.perf_counter_ns() - processOffsetTime),
        'event': "Scroll"
    })
    #print(f'Scrolled {"down" if dy < 0 else "up"} at ({x}, {y})')

def HookKeyboardEvent():
    with hookKeyboardThread as listener:
        listener.join()
    logging.debug("Hook Keyboard Close")
def HookMouseEvent():
    with hookMouseThread as listener:
        listener.join()
    logging.debug("Hook Mouse Close")


hookMouseThread = mouse.Listener(on_click=on_mouse, on_move=on_move, on_scroll=on_scroll)
hookKeyboardThread = keyboard.Listener(on_press=on_key_press, on_release=on_key_release)
threading.Thread(target=HookMouseEvent, daemon=True).start()
threading.Thread(target=HookKeyboardEvent, daemon=True).start()

main()