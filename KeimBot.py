#!/usr/bin/python
# -*- coding: UTF-8 -*-

import pyperclip
import pyautogui
import win32gui
import win32con
import win32api
from win32con import WM_INPUTLANGCHANGEREQUEST
from PIL import ImageGrab
import cv2
import time
import sys
import numpy as np
from aip import AipOcr


class Bot(object):
    def __init__(self,baiduAPK:dict):
        self.Window_Size=[]
        self.sleeptime=0.5
        self.hwnd=None
        self.APP_ID=''
        self.API_KEY=''
        self.SECRET_KEY=''
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        self.AppName=''
        self.AppClassName=''
        """
        初始化
        """
        if(baiduAPK!=None):
            if(baiduAPK.__contains__('AppName')):
                self.AppName=baiduAPK.get('AppName')
            else:
                raise Exception('AppName is Error')

            if(baiduAPK.__contains__('APIID')):
                self.APP_ID=baiduAPK['APIID']
            else:
                raise Exception('APIID is Error')

            if(baiduAPK.__contains__('AppClassName')):
                self.AppClassName=baiduAPK.get('AppClassName')
            else:
                raise Exception("AppClassName is Error")

            if baiduAPK.__contains__('APIKEY'):
                self.API_KEY=baiduAPK['APIKEY']
            else:
                raise Exception('APIKEY is Error')

            if baiduAPK.__contains__('SECRETKEY'):
                self.SECRET_KEY=baiduAPK['SECRETKEY']
            else:
                raise Exception('SECRETKEY is Error')

        pass
        
    pass

    def Exec(self,commands:list):
        """
        执行bot脚本 {COM:K,VALUE}
        """
        if sys.platform!='win32':
            raise Exception('Error System OS')
        pass
        self.__Initialization()
        if(self.hwnd!=None):
            for value in commands:
                if(value["COM"]=='K'):
                    self.__KeyboardOption(value["VALUE"])
                elif value["COM"]=='M':
                    self.__MouseOption(value["VALUE"])
                elif value["COM"]=='W':
                    self.__sleepOption(value["VALUE"])
                elif value['COM']=='C':
                    img=self.__CutImage(value["VALUE"])
                    msg=self.__VerificatioImage(img)
                    return msg
                else:
                    print('NOT FIND COMMAND')
            pass
        pass
    pass

    def __VerificatioImage(self,img):
        """
        验证微信图片
        """
        msg={'ST':False,"VS":[]}
        try:
            picfile='tmp.jpg'
            vimg=np.array(img)
            p0=cv2.cvtColor(vimg,cv2.COLOR_RGB2GRAY)
            cv2.imwrite(picfile, p0, [int(cv2.IMWRITE_JPEG_QUALITY),95])
            i = open(picfile, 'rb')
            img = i.read()
            aipOcr = AipOcr(self.APP_ID, self.API_KEY, self.SECRET_KEY)  # 初始化AipFace对象
            message = aipOcr.basicAccurate(img)
            for text in message.get('words_result'):  # 识别的内容
                msg['VS'].append(text.get('words'))
            pass
            msg['ST']=True
            return msg
        except Exception as ex:
            print(f"VerificatioImage:{ex}")
            return msg
    pass

    def __CutImage(self,margin:list=[0,0,0,0]):
        """
        截图 margin 左上右下
        """
        vimg=ImageGrab.grab((self.Window_Size[0]+margin[0],
        self.Window_Size[1]+margin[1],
        self.Window_Size[2]+margin[2],
        self.Window_Size[3]+margin[3]))
        vimg= vimg.convert("RGB")
        return vimg
    pass

    def __Initialization(self):
        """
        docstring
        """
        try:
            self.hwnd = win32gui.FindWindow(self.AppClassName, self.AppName)
            if(self.hwnd!=None):
               win32gui.ShowWindow(self.hwnd, win32con.SW_SHOWNORMAL)
               win32gui.SetForegroundWindow(self.hwnd)
               self.Window_Size=win32gui.GetWindowRect(self.hwnd)
            pass
        except Exception as ex:
            self.hwnd=None
            pass
    pass

    def __sleepOption(self, ky):
        """
        停止
        """
        T=self.sleeptime
        if(ky.__contains__('T')):
            T=ky['T']

        time.sleep(T)
    pass

    def __KeyboardOption(self,ky:dict):
        """
        键盘操作方法
        dict: T:0输入 1:键 2:快捷键 M:内容 {T:0,M:'abc'}} 
        """
        T=0
        M=''
        status=True
        if(ky.__contains__('T')):
            T=ky['T']
        else:
            status=False

        if(ky.__contains__('M')):
            M=ky['M']
        else:
            status=False
        
        if status==False:
            raise Exception('参数错误')
        pass
        
        #结构{T:0`1,V:['1','2','3']}
        #0为英文
        #1为中文
        if T==0:
            pyperclip.copy(M)
            pyautogui.hotkey("ctrl",'v')
        elif T==2:
            if(len(M)==1):
                pyautogui.hotkey(M[0])
            elif(len(M)==2):
                pyautogui.hotkey(M[0],M[1])
            elif(len(M)==3):
                pyautogui.hotkey(M[0],M[1],M[2])
            else:
                pass
        else:
            pyautogui.press(M)
        pass
    pass
   
    def __MouseOption(self,mu:dict):
        """
        鼠标操作方法
        dict: T:0移动 1为操作 x:坐标 y:坐标 E:左击1 中击2 右击3 C:点击次数 {T:1,E:0,X:1,Y:2,C:1}
        """
        T=0
        X=0
        Y=0
        E=0
        C=0
        but=''
        status=True

        if(mu.__contains__('X')):
            X=mu['X']
        else:
            status=False

        if(mu.__contains__('Y')):
            Y=mu['Y']
        else:
            status=False

        if(mu.__contains__('T')):
            T=mu['T']
        else:
            status=False

        if(status==True and T!=0): 
            if(mu.__contains__('E')):
                E=mu['E']
            else:
                status=False

            if(mu.__contains__('C')):
                C=mu['C']
            else:
                status=False
        pass

        if(status==False):
            raise Exception('参数错误')
        pass

        if E==0:
            but='left'
        elif E==1:
            but='middle'
        else:
            but='right'
        pass

        if(T==0):
            pyautogui.moveTo(x=self.Window_Size[0]+X,y=self.Window_Size[1]+Y)
        elif(T==1):
            if(X==None and Y==None):
                pyautogui.click(x=None,y=None,clicks=C,button=but,
            duration=0.0, tween=pyautogui.linear)
            else:
                pyautogui.click(x=self.Window_Size[0]+X,y=self.Window_Size[1]+Y,clicks=C,button=but,
            duration=0.0, tween=pyautogui.linear)
        elif T==2:
            pyautogui.dragTo(x=X, y=Y, duration=3,button=but)
        pass
    pass
pass



'''
命令结构定义
[
	{'COM':'K','VALUE':{'T':0,'M':'测试输入值'}},             #键盘输入abc
    {'COM':'K','VALUE':{'T':1,'M':'enter'}},            #键盘指定键enter
    {'COM':'M','VALUE':{'T':0,'X':300,'Y':300}},             #鼠标移动到x1 y2位置
    {'COM':'M','VALUE':{'T':1,'E':0,'X':1,'Y':2,'C':1}},     #鼠标在x1 y2位置左击1次
    {'COM':'M','VALUE':{'T':1,'E':2,'X':1,'Y':2,'C':1}},     #鼠标在x1 y2位置中击1次
    {'COM':'M','VALUE':{'T':1,'E':3,'X':1,'Y':2,'C':1}},     #鼠标在x1 y2位置右击1次
    {'COM':'M','VALUE':{'T':3,'E':3,'X':1,'Y':2,'C':1}}     #鼠标拖拽到x1 y2位置
]
'''
