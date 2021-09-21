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
    id_ok = 1
    password_window = user32.FindWindowA(None, raw_str("VBAProject - Project Properties"))
    if password_window == 0:
        raise WaitException("Fail to Find Project Properties Window to Close")
    
    print("Found Project Properties Window to Close")
    user32.SendMessageA(password_window, commctrl.TCM_SETCURFOCUS, 1, 0)
    
    ok_button = user32.GetDlgItem(password_window, id_ok)
    if ok_button == 0:
        raise WaitExceptiion("Fail to find ok button in project properties window")
    
    user32.SetFocus(ok_button)
    user32.SendMessageA(ok_button, win32con.BM_ClICK, 0)


    
def lock_vba_project(application):
    id_ok = 1
    id_tabcontrol = 0x3020
    id_subdialog = 0x8002
    id_checkbox_lock = 0x1557
    id_textbox_pass1 = 0x1555
    id_textbox_pass2 = 0x1556
    
    password_window = user32.FindWindowA(None, raw_str("VBA Project - Project Properties"))
    if password_window == 0:
        raise WaitException("Fail to find project properties window")
    
    print("Found project properties window")
    tabcontrol = user32.GetDlgItem(password_window, id_tabcontrol)
    user32.SendMessageA(tabcontrol, commctrl.TCM_SETCURFOCUS, 1, 0)
    if user32.SendMessageA(tabcontrol, commctrl.TCM_GETCURFOCUS) != 1:
        raise WaitException("Fail to change tab control")
    
    subdialog = user32.FindWindowExA(password_window, 0, id_subdialog, None)
    if subdialog == 0:
        raise WaitException("Fail to find subdialog")
    
    checkbox_lock = user32.GetDlgItem(subdialog, id_checkbox_lock)
    if checkbox_lock == 0:
        raise WaitException("Fail to find checkbox")
    
    user32.SetFocus(checkbox_lock)
    user32.SendMessageA(checkbox_lock, win32con.BM_SETCHECK, win32con.BST_CHECKED, 0)
    
    checkbox_state = user32.SendMessageA(checkbox_lock, win32con.BM_GETCHECK)
    if checkbox_state != win32con.BST_CHECKED:
        raise WaitException("Fail to activate checkbox")
        
    textbox_pass1 = user32.GetDlgItem(subdialog, id_textbox_pass1)
    if textbox_pass1 == 0:
        raise WaitException("Fail to find password box 1")
    
    user32.SetFocus(textbox_pass1)
    user32.SendMessageA(textbox_pass1, win32con.WM_SETTEXT, None, raw_str(ProjectConstants.password))
    length = user32SendMessage(textbox_pass1, win32con.WM_GETTEXTLENGTH)
    
    if length != len("063"):
        raise WaitException("Fail to complete password box 1")
    
    
    textbox_pass2 = user32.GetDlgItem(subdialog, id_textbox_pass2)
    user32.SetFocus(textbox_pass2)
    if textbox_pass2 == 0:
        raise WaitException("Fail to find password box 2")
    
    user32.SetFocus(textbox_pass2)
    user32.SendMessageA(textbox_pass2, win32con.WM_SETTEXT, None, raw_str(ProjectConstants.password))
    length = user32SendMessage(textbox_pass2, win32con.WM_GETTEXTLENGTH)
    
    if length != len("063"):
        raise WaitException("Fail to complete password box 2")
    
    
    ok_button = user32.GetDlgItem(password_window, id_ok)
    if ok_button == 0:
        raise WaitException("Fail to find OK button")
    
    user32.SetFocus(ok_button)
    user32.SendMessageA(ok_button, win32con.BM_CLICK, 0)
    return True
        
    


def extract_lookup(col_index, row_range, ws):
    return [data.Value for data in [ws.Range(loc)
            for loc in [str(col_index) + str(ii)
            for ii in row_range]]]


def wait_loop(timeout_sec, application, func):
    timeout = time.time() + timeout_sec
    while time.time() < timeout:
        try:
            done_run = func(application)
            if done_run:
                break
        
        except WaitException as e:
            print(str(e))
            sleep()

            
            
def change_property_data(wb_, new_p_version_):
    property_ws = wb_.Worksheets("Property Data")
    cell = property_ws.Range("B32")
    cell.Value = new_p_version_

    
    
    
def change_reference_tables(wb_):
    ref_key_col = 'AI'
    ref_val_col = 'AJ'
    ref_start_row = 3
    ref_end_row = 24
    
    to_replace = {'Undoubted: 8.0
                  'Unrated > 5 years': 3.0,
                  'Large pool': 6.0,
                  'Small pool': 3.0}
    
    reference_ws = wb_.Worksheets("Reference Tables")
    row_range = range(ref_start_row, ref_end_row + 1)
    lookup_keys = extract_lookup(ref_key_col, row_range, reference_ws)
    lookup_values = extract_lookup(ref_val_col, row_range, reference_ws)
    original_values = dict(zip(lookup_keys, lookup_values))
    
    new_values = original_values.copy()
    
    for k, v in to_replace.items():
        new_values[k] = v
    
    for i, k in zip(row_range, lookup_keys):
        reference_ws.Range(str(ref_val_col) + str(i)).Value = new_values[k]





def change_debt_Formula(wb_):
    formula_ws = wb_.Worksheets("Debt")
    
    for i in range(2, 6):
        
        cell1 = formula_ws.Range("O{0}".format(i))
        cell2 = formula_ws.Range("P{0}".format(i))
        
        Formula1 = "=IFERROR(-PMT(XXXXXXXX)".format(i)
        Formula2 = same as above
        
        cell1.Value = Formula1
        cell2.Value = Formula2
    
    time.sleep(3)

    
    
    
    
def change_vba_prologue(app_, timeout_second_):
    app_.CommandBars.ExecuteMso("ViewCode")
    wait_loop(timeout_second_, app_, unlock_vba_project)
    wait_loop(timeout_second_, app_, close_vba_project_window)

    
    
    
    
def change_vba(wb_, old_c_version_, new_c_version_):
    match = old_c_version_
    replacement = new_c_version_
    
    code_base = wb_.VBAProject.VBComponents("Complete").CodeModule
    startrow = 0
    
    while True:
        success, startrow, startcol, endrow, endcol = code_base.Find(match, startrow +1, 1, -1, -1)
        
        if not success:
            break
        
        old_line = code_base.Lines(startrow, 1)
        new_line = old_line[:startcol - 1] + replacement + old_line[endcol - 1:]
        code_base.ReplaceLine(startrow, new_line)

        
        
        
        
def change_vba_formula(wb_, old_intersect1, old_intersect2, old_intersect3):
    match1 = old_intersect1
    match2 = old_intersect2
    match3 = old_intersect3
    
    code_base = wb_.VBAProject.VBAComponents("Sheet03").CodeModule
    startrow = 0
    
    while True:
        success, startrow, startcol, endrow, endcol = code_base.Find(match1, startrow +1, 1, -1, -1)
        
        if not success:
            break
            
        new_line = "'" + match1
        code_base.ReplaceLine(startrow, new_line)
    
    startrow = 0
    while True:
        success, startrow, startcol, endrow, endcol = code_base.Find(match2, startrow +1, 1, -1, -1)
        
        if not success:
            break
            
        new_line = "'" + match2
        code_base.ReplaceLine(startrow, new_line)
        
    new_line = "'" + match3
    code_base.ReplaceLine(startrow, new_line)
    
    time.sleep(2)
    

    
    
    
    
def change_back_vba_formula(wb_, new_intersect1, new_intersect2, old_intersect1, old_intersect2, old_intersect3):
    match1 = new_intersect1
    match2 = new_intersect2
    
    code_base = wb_.VBProject.VBComponents("Sheet03").CodeModule
    startrow = 0
    
    while True:
        success, startrow, startcol, endrow, endcol = code_base.Find(match1, startrow +1, 1, -1, -1)
        
        if not success:
            break
            
        new_line = old_intersect1
        code_base.ReplaceLine(startrow, new_line)
        
    
    startrow = 0
    
    while True:
        success, startrow, startcol, endrow, endcol = code_base.Find(match2, startrow +1, 1, -1, -1)
        
        if not success:
            break
            
        new_line = old_intersect2
        code_base.ReplaceLine(startrow, new_line)
        
    
    new_line = old_intersect3
    code_base.ReplaceLine(startrow, new_line)
    
    time.sleep(2)
    

    
    
    
def change_vba_epilogue(app_, timeout_second_):
    id_project_properties = 2578
    app_.VBE.CommandBars.FindControl(Id=id_project_properties).Execute()
    wait_loop(timeout_second_, app_, lock_vba_project)

    
    
def terminate():
    global flag
    while (1):
        hwnd = win32gui.FindWindow(None, 'VBAProject Password')
        if hwnd != 0:
            print("\n")
            print("\n")
            print("Found Password Window")
            id_password = 0x155e
            id_ok = 1
            
            text_box = user32.GetDlgItem(hwnd, id_password)
            ok_button = user32.GetDlgItem(hwnd, id_ok)
            if text_box == 0 and ok_button == 0:
                raise WaitException("Fail to find textbox and okbutton in password window")
                
            user32.SetFocus(text_box)
            user32.SendMessageA(text_box, win32con.WM_SETTEXT, None, raw_str(ProjectConstants.password))
            
            user32.SetFocus(ok_button)
            user32.SendMessageA(ok_button, win32con.BM_CLICK, 0, 0)
            break
            
        if flag == True:
            break
            

    


    
    
    
class run_main:
    
    def __init__(self, entry, root):
        self.entry = entry
        self.root = root
        self.sf   = ' ' 
        self.of   = ' '
        
    def main(self):
        my_progress['value'] = 10
        self.root.update_idletasks()
        
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        
        Path      = self.entry['Folder_Path'].get()
        logfile   = dir_path + '\\Excel_Editor_Automation.log'
        new_p_version = 'Version: 1.3'
        old_c_version = 'Completed-V2'
        new_c_version = 'Completed-V3'
        new_e_version = '1.3'
        
        match1  = 'If Not Application.Intersect(ActiveCell, Range("B2:X45")) Is Nothing Then'
        match2  = 'Sheet3.Range("A4").ClearContents'
        match3  = 'End If'
        
        new_match1 = "'" + match1
        new_match2 = "'" + match2
        
        LOG_FORMAT = "%(levelname)s:%(asctime)s:%(message)s"
        
        try:
            logging.basicConfig(filename = logfile, level = logging.DEBUG, format = LOG_FORMAT, filemode = 'w')
            logger = logging.getLogger()
        except Exception as e:
            ctypes.windll.user32.MessageBox(0, 'Issue for creating log file', 'Warning', 1)
            logger.warning('Issue for creating log file')
            logger.warning(e)
            quit(self.root)
            
            
        my_progreww['value'] = 20
        self.root.upadte_idletasks()
        
        i = 0
        logger.info("Pre-Varibles are set up and will start the 'for' loop")
        
        for f in os.listdir(Path):
            
            if f.endswith(".xlsm"):
                inp   = Path + '\\' + f
                outp  = Path + '\\Output' + "(Converted_NewVersion_{0})_",format(new_e_version) + f
                xlsmCounter = len(glob.glob1(Path, "*.xlsm"))
                increment   = (90-20)/xlsmCounter
                
                logger.info("Forloop-Variables are set up")
                
                my_progress['value'] = my_progress['value'] + increment
                self.root.update_idletasks()
                
                
                
                try:
                    wb = app.Workbooks.Open(inp)
                    time.sleep(5)
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0} is not found or opened".format(f), "warning", 1)
                    logger.warning("{0} is not found or opened".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                try:
                    wb.Unprotect(ProjectConstants.password)
                    time.sleep(1)
                    logger.info("{0} has been unprotected".format(f))
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0} can not be unprotected".format(f), "warning", 1)
                    logger.warning("{0} can not be unprotected".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                
                
                try:
                    change_property_data(wb, new_p_version)
                    time.sleep(1)
                    logger.info("{0}'s Property Data Tab has been updated".format(f))
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0}'s Property Data Tab can not be updated".format(f), "warning", 1)
                    logger.warning("{0}'s Property Data Tab can not be updated".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                   
                
                try:
                    change_reference_tables(wb)
                    time.sleep(1)
                    logger.info("{0}'s Reference Data Tab has been updated".format(f))
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0}'s Reference Data Tab can not be updated".format(f), "warning", 1)
                    logger.warning("{0}'s Reference Data Tab can not be updated".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                    
                    
                t = threading.Thread(target = terminate)
                t.start()
                
                
                try:
                    app.CommandBars.ExecuteMso("ViewCode")
                except:
                    t.join()
                    None
                if t.is_alive():
                    t.join()
                    time.sleep(1)
                    
                    
                    
                
                try:
                    change_vba(wb, old_c_version, new_c_version)
                    time.sleep(1)
                    logger.info("{0}'s VBA has been updated".format(f))
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0}'s VBA can not be updated".format(f), "warning", 1)
                    logger.warning("{0}'s VBA can not be updated".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                    
                    
                try:
                    change_vba_formula(wb, match1, match2, match3)
                    time.sleep(1)
                    logger.info("{0}'s VBA Debt Formula has been Commented".format(f))
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0}'s VBA Debt Formula can not be commented".format(f), "warning", 1)
                    logger.warning("{0}'s VBA Debt Formula can not be commented".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                    
                    
                try:
                    change_debt_formula(wb)
                    time.sleep(1)
                    logger.info("{0}'s Debt Tab Formula has been updated".format(f))
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0}'s Debt Tab Formula can not be updated".format(f), "warning", 1)
                    logger.warning("{0}'s Debt Tab Formula can not be updated".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                    
                    
                try:
                    change_back_vba_formula(wb, new_match1, new_match2, match1, match2, match3)
                    time.sleep(1)
                    logger.info("{0}'s VBA Debt Formula has been changed back".format(f))
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0}'s VBA Debt Formula can not be changed back".format(f), "warning", 1)
                    logger.warning("{0}'s VBA Debt Formula can not be changed back".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                    
                    
                wb.Protect(ProjectConstants.password)
                time.sleep(1)
                logger.info("{0} has been re-protected".format(f))
                app.DisplayAlerts = False
                
                
                
                try:
                    wb.SaveAs(Filename = outp)
                    time.sleep(1)
                    logger.info("{0} has been saved".format(f))
                except Exception as e:
                    ctypes.windll.user32.MessageBoxW(0, "{0} can not be saved as".format(f), "warning", 1)
                    logger.warning("{0} can not be saved as".format(f))
                    logger.warning(e)
                    wb.Close(False)
                    app.Quit()
                    continue
                    
                    
                    
                
                app.DisplayAlert = True
                wb.Close()
                
                app.Quit()
                
                i += 1
                logger.info("{0}'s updates are done".format(f))
                
                
            if i == 0:
                ctypes.windll.user32.MessageBoxW(0, "No file is updated or No xlsm file has been found in the folder", "info", 1)
                logger.warning("No file starting with 'EMV' has been found in the folder. Check your Path again")
                quit(self.root)
                
            my_progress['value'] = 100
            self.root.update_idletasks()
            
            original_files = len(glob.glob1(path, "*.xlsm"))
            
            logger.info("The Whole Process Completed")
            print("\n")
            print("\n")
            print("-"*50)
            print("\n")
            print("The work is done")
            print("\n")
            print("-"*50)
            
            self.sf = i
            self.0f = original_files
            
            time.sleep(2)
            
            
            
            
    def message(self):
        self.main()
        ctypes.windll.user32.MessageBoxW(0, "Process completed! You've successfully converted {0} out of {1} files".format(self.sf, self.of), "info", 1)
        
        
            
          

        
        
        
        
def makeform(root):
    entries = {}
    
    lab = tk.Label(root, width=18, text='Folder Path:', font=(None, 10, 'bold'), anchor='w', fg='White', bg='black')
    lab.place(x=30, y=30)
    
    lab = tk.Label(root, width=18, text='Progress Bar:', font=(None, 10, 'bold'), anchor='w', fg='White', bg='black')
    lab.place(x=30, y=130)
    
    folder_path_text  = tk.StringVar()
    folder_path_entry = tk.Entry(root, textvariable=folder_path_text)
    
    folder_path_entry.place(x=150, y=30, width=700, height=25)
    entries['Folder_Path'] = folder_path_entry
    
    return entries



def quit(root):
    root.destory()

    
    


root = tk.Tk()
w = 900
h = 180
ws = root.winfo_screenwidth()
hs = root.winfo_screenheight()

x = (ws/2) - (w/2)
y = (hs/2) - (h/2)

root.geometry("%dx%d+%d+%d" % (w, h, x, y))
root.configure(background='black')
root.attributes('-alpha', 0.90)
file_location = makeform(root)

my_progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=300, mode="determinate")
my_progress.place(x=150, y=130)

Quick = tk.Button(root, text="Quit", command=lambda root=root: quit(root), height=1, width=10, bg='White', fg='Black', font=(None, 10, 'bold'))
Quick.place(x=260, y=80)


Submit = tk.Button(root, text="Run", command=lambda e=file_location, root=root: [run_main(e, root).message()], height=1, width=10, bg="White", fg="Black", fond=(None, 10, 'bold'))
Submit.place(x=150, y=80)


root.mainloop()
                
                
                
        
        
        
        
        
    
    
    
    
    
    
