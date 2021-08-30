import comtypes.client
import win32con
import win32com.client
import commctrl
import time
import threading
import win32gui
import os
import tkinter as tk
from tkinter import ttk
import glob
import logging
import ctypes


user32 = comtypes.windll.user32
flag = False

dir_path = os.path.dirname(os.path.realpath('__file__'))


class ProjectConstants:

    password = '063'
    timeout_second = 100
    fail_sleep_duration_second = 0.1
    

class WaitException(Exception):
    pass


def raw_str(string):
    return comtypes.c_char_p(bytes(string, 'utf-8'))
    

def sleep():
    time.sleep(ProjectConstants.fail_sleep_duration_second)
    

def unlock_vba_project(application):
    id_password = 0x155e
    id_ok = 1

    password_window = user32.FindWindowA(None, raw_str("VBAProject Password"))
    if password_window == 0:
        raise WaitException("Fail to Find Password Window")
    
    print("Found Password Window")
    user32.SendMessageA(password_window, commctrl.TCM_SETCURFOCUS, 1, 0)
    
    text_box = user32.GetDlgItem(passowrd_window, id_password)
    ok_button = user32.GetDlgItem(password_window, id_ok)
    if text_box == 0 and ok_button == 0:
        raise WaitException("Fail to Find Textbox and OK Button in Password Window")
    
    user32.SetFocus(text_box)
    user32.SendMessageA(text_box, win32com.WM_SETTEXT, None, raw_str(ProjectConstants.password))
    
    Length = user32.SendMessageA(text_box, win32com.WM_GETTEXTLENGTH)
    if length != len(ProjectConstants.password):
        raise WaitException("Fail to Verify Password Length")
    
    user32.SetFocus(ok_button)
    user32.SendMessageA(ok_button, win32con.BM_CLICK, 0, 0)
    return True


def close_vba_project_window(application):
    











