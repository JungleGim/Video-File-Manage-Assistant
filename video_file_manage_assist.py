"""
Script provides a GUI for helping to manage MKV files when ripping/updating/storing/pre-processing media rips.

Functions include:
1) File length export
2) File bulk re-name + organize

Required (non-standard) dependencies:
-cv2
-pandas
-openpyxl (pandas open excel required)
"""

#-----------------------------imports
import datetime
import tkinter as tk
import os
from tkinter import filedialog, font, ttk, messagebox, Scrollbar
import cv2
import csv
import shutil
import pandas as pd
import numpy as np
from pathlib import Path

#-----------------------------supporting methods, classes, constants
#--class for media file properties
class media_file_props:
    def __init__(self, kwargs):
        """class containing media file properties to import/export"""
        self.file_name = kwargs.get('name')         #file name
        self.full_path = kwargs.get('path')         #full file path
        self.est_time_raw = kwargs.get('runtime')   #estimated playtime in seconds
        self.good_estimate = kwargs.get('est_ok')   #able to successfully get estimate
        self.est_time_str = None                    #estimated file length, formatted string in HH:MM:SS
        self.calc_timestr()                         #when instancing, create output formatted string
    
    def calc_timestr(self):
        """function generates the output estimated time string in a human readable format of HH:MM:SS"""
        if self.good_estimate == True: self.est_time_str = str(datetime.timedelta(seconds=int(self.est_time_raw)))
        else: self.est_time_str = ''

#--class for asking user file/folder to parse
class user_prompt_open_type(tk.Toplevel):
    def __init__(self, master):
        """toplevel child window for prompting users how/what type of file/dir to open"""
        super().__init__(master)
        self.grab_set()                 #force focus
        self.title("Select Open Type")  #title bar
        self.resizable(False,False)     #fixed size
        self.master_ref = master
        self.result = None              #result choice

        self.protocol("WM_DELETE_WINDOW", self.on_close)    #handle window close button
        self.init_main_window()                             #initialize window elements
        self.wait_window()                                  #wait in this window until destroyed

    def init_main_window(self):
        """function initiates the various user window elements"""
        question = tk.Label(self, font=sys_fnt_txt, text='Would you like to parse all files\nin a directory, or a single file?')
        question.grid(row=0,column=0,columnspan=2, padx=10, pady=10)
        btn_file = tk.Button(self, text='Parse Directory', font=sys_fnt_BTN, command=lambda:self.set_restult(CONST_openType['dir'])) 
        btn_file.grid(row=1, column=0, padx=10, pady=10)
        btn_file = tk.Button(self, text='Single File', font=sys_fnt_BTN, command=lambda:self.set_restult(CONST_openType['file'])) 
        btn_file.grid(row=1, column=1, padx=10, pady=10)
        btn_close = tk.Button(self, text='Cancel', font=sys_fnt_BTN, command=self.on_close) 
        btn_close.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

    def set_restult(self,type):
        """function sets the result after users make their selection
        
        :param type: type of open action user selected
        :type type: `CONST_openType` entry
        """
        self.result = type
        self.on_close()     #close window
    
    def on_close(self): #make no changes
        """function is called when the window is closed"""
        self.destroy()

#--class for asking user to update files
class user_prompt_update(tk.Toplevel):
    def __init__(self, master):
        """toplevel child window for prompting users to select input template for updates"""
        super().__init__(master)
        self.grab_set()                 #force focus
        self.title("Select Open Type")  #title bar
        self.resizable(False,False)     #fixed size
        self.master_ref = master
        self.tmplt_var = tk.StringVar() #temp result path
        self.del_var = tk.BooleanVar()  #temp delete variable
        self.template_path = None       #result choice
        self.del_old = False            #delete old files - default false

        self.protocol("WM_DELETE_WINDOW", self.on_close)    #handle window close button
        self.init_main_window()                             #initialize window elements
        self.wait_window()                                  #wait in this window until destroyed

    def init_main_window(self):
        """function initiates the various user window elements"""
        self.tmplt_var.set('<select file>')
        question = tk.Label(self, font=sys_fnt_txt, text='Please select update template and\nselect options to update/move files')
        question.grid(row=0,column=0,columnspan=2, padx=10, pady=10)
        btn_sel_tmplt = tk.Button(self, text='Select Template File', font=sys_fnt_BTN, command=self.browse_template_file)
        btn_sel_tmplt.grid(row=1, column=0, padx=10, pady=10)
        file_lbl = tk.Label(self,font=sys_fnt_BTN, textvariable=self.tmplt_var, state='disabled')
        file_lbl.grid(row=1, column=1, padx=10, pady=10)
        del_old_ckbx = tk.Checkbutton(self, text='Delete old files?', font=sys_fnt_txt, variable=self.del_var)
        del_old_ckbx.grid(row=2, column=0, padx=10, pady=10)
        ctl_frame = tk.Frame(self)
        ctl_frame.grid(row=3, column=0, padx=10, pady=(10,20), columnspan=2, sticky=tk.EW)
        ctl_frame.grid_columnconfigure(0,weight=1); ctl_frame.grid_columnconfigure(1,weight=1)
        btn_upd = tk.Button(ctl_frame, text='Update', font=sys_fnt_BTN, command=self.set_restult)
        btn_upd.grid(row=0, column=0, padx=10, pady=10)
        btn_close = tk.Button(ctl_frame, text='Cancel', font=sys_fnt_BTN, command=self.on_close)
        btn_close.grid(row=0, column=1, padx=10, pady=10)
    
    def browse_template_file(self):
        """function displays a file browser to select the input template
        
        :returns: file path of template
        :rtype: `os` path
        """
        filetypes=[]                                                            #temp list for accepted file types
        for k,v in CONST_template_formats.items(): filetypes.append((k,v))      #build list for file picker
        dialog_opts = {'filetypes':filetypes,
                       'title':'Update Template'}  
        self.tmplt_var.set(open_file_dialog(dialog_opts))  #set temp result

    def check_results(self):
        """function error checks the selected results before closing.
        
        :returns: results are valid/ok
        :rval: `boolean` - results are OK
        """
        rval = False                        #default return
        crnt_path = self.tmplt_var.get()    #current selected file path
        
        if check_file_exists(crnt_path)==True:                          #if the selected file exists
            file_ext = get_file_ext(os.path.basename(crnt_path))        #get its extension
            if file_ext in list(CONST_template_formats.values()):       #if its a valid extension
                if file_ext == '.csv': data_pd = pd.read_csv(crnt_path, nrows=1) #open if CSV
                else: data_pd = pd.read_excel(crnt_path, nrows=1)                #open if excel type
                file_hdrs = list(data_pd.columns)                   #put headers in a list
                if sys_tmplt_oldFile_hdrName in file_hdrs and sys_tmplt_newFile_hdrName in file_hdrs:
                    rval = True #was able to find a valid file that contains both required headers, OK to import

        return rval

    def set_restult(self):
        """function sets the result after users make their selection
        """
        if self.check_results() == True:
            self.template_path = self.tmplt_var.get()   #update outputs
            self.del_old = self.del_var.get()
            self.on_close()     #close window
        else:
            messagebox.showerror("Error", "File selection not valid or cannot find required headers.\nTemplate must have colunms for \"File_Path\" and \"New_File_Path\"")
    
    def on_close(self): #make no changes
        """function is called when the window is closed"""
        self.destroy()

#--class for error message on upload check
class wndw_notify(tk.Toplevel):
    def __init__(self, parent, kwargs):
        '''custom notification window class. Fixed size window that wraps text and can handle longer messages.
        Based on passed kwargs can be one of several types. Meant to cover instances where the built-in tk 
        notification window types are not sufficient (like in the case of long messages)'''
        super().__init__(parent)                #init as a sub-window of the parent
        #---core window options
        self.grab_set()                         #force focus on this window
        self.resizable(False,False)             #not resizable

        #---local vars for window elements
        self.txt_title = kwargs.get('title','Message')      #title bar text
        self.txt_msg = kwargs.get('message',None)           #message text
        self.err_list = kwargs.get('err_list',None)         #list of error tuples
        self.continue_upd = False                           #result of the user selection

        self.protocol("WM_DELETE_WINDOW", self.on_close)    #handle window close button
        self.wndw_init()    #initialize window elements
        self.upd_listbox()  #update listbox with errors
        self.bell()         #popup/wanring notification sound
        self.wait_window()  #wait in this window until destroyed
    
    def wndw_init(self):
        """function initiates/builds various window elements"""
        #---frames for grouping widgets
        self.frm_text = tk.Frame(self); self.frm_text.grid(row=0, column=0, padx=10, pady=(20,10))      #frame space for the notification text
        self.frm_errs = tk.Frame(self); self.frm_errs.grid(row=1, column=0, padx=10, pady=10, sticky=tk.NSEW)           #frame space for error list
        self.frm_ctl = tk.Frame(self); self.frm_ctl.grid(row=2, column=0, columnspan=2, pady=(10,20))   #frame space for control buttons
        self.frm_text.grid_columnconfigure(1, weight=1)     #assign the extra weight to column 1 (message text space)
        self.frm_errs.grid_columnconfigure(0, weight=1)     #fill the full width

        #---common elements
        #-display icon
        ico = tk.Label(self.frm_text, image="::tk::icons::information")
        ico.grid(row=0, column=0, sticky=tk.W)
        #-dispaly text
        message_text = tk.Label(self.frm_text, font=sys_fnt_txt, text=self.txt_msg, wraplength=sys_wrap_len)
        message_text.grid(row=0,column=1)
        #-title bar
        self.title(self.txt_title)

        #-error listbox
        self.err_listbox = tk.Listbox(self.frm_errs, font=sys_fnt_txt, width=50, height=10) #create listbox to display errors
        self.err_listbox.grid(row=0, column=0, sticky=tk.NSEW)
        err_scrl_v = Scrollbar(self.frm_errs); err_scrl_v.grid(row=0, column=1, sticky=tk.NS)   #vert scrollbar
        self.err_listbox.config(yscrollcommand = err_scrl_v.set)    #attach listbox yscroll
        err_scrl_v.config(command = self.err_listbox.yview)         #binding scroll to listbox property
        err_scrl_h = Scrollbar(self.frm_errs, orient='horizontal'); err_scrl_h.grid(row=1, column=0, sticky=tk.EW)   #horz scrollbar
        self.err_listbox.config(xscrollcommand = err_scrl_h.set)    #attach listbox yscroll
        err_scrl_h.config(command = self.err_listbox.xview)         #binding scroll to listbox property

        #-check the contained errors for building the remaining part of the frame
        err_types = []                                          #temp list for new files
        for tup in self.err_list:
            if tup[0] not in err_types: err_types.append(tup[0])

        if CONST_err_types['err'] in err_types:
            ctl_txt = tk.Label(self.frm_ctl, font=sys_fnt_txt, text='Unable to update, please resolve errors before continuing.', wraplength=sys_wrap_len)
            ctl_txt.grid(row=0,column=0)
            btn_ok = tk.Button(self.frm_ctl, text='Ok', font=sys_fnt_BTN, command=lambda:self.set_restult(False))
            btn_ok.grid(row=1, column=0, padx=10, pady=10)
        else:
            ctl_txt = tk.Label(self.frm_ctl, font=sys_fnt_txt, text='Warnings Detected. Would you like to continue with the update?', wraplength=sys_wrap_len)
            ctl_txt.grid(row=0,column=0, columnspan=2)
            btn_yes = tk.Button(self.frm_ctl, text='Yes', font=sys_fnt_BTN, command=lambda:self.set_restult(True))
            btn_yes.grid(row=1, column=0, padx=10, pady=10)
            btn_no = tk.Button(self.frm_ctl, text='No', font=sys_fnt_BTN, command=lambda:self.set_restult(False))
            btn_no.grid(row=1, column=1, padx=10, pady=10)
    
    def upd_listbox(self):
        """function updates the listbox values"""
        self.err_listbox.delete(0, tk.END)          #clear listbox
        for tup in self.err_list:                   #cycle through entries in error dict
            string = f"{tup[0]}: {tup[1]}"          #make the display string
            self.err_listbox.insert(tk.END, string) #insert display string
    
    def set_restult(self,result):
        """function sets the result after users make their selection
        
        :param result: result of the selection
        :type result: `bool` - True OK to continue
        """
        self.continue_upd = result
        self.on_close()     #close window
    
    def on_close(self): #make no changes
        """function is called when the window is closed"""
        self.destroy()

def open_file_dialog(file_dialog_kwargs):
    """function opens the file picker dialog to prompt the user for a file to
    open. dialog is configurable based on the passed kwargs. Function takes a dict 
    of kwargs for the tkinter `ask open` dialog
    
    :param file_dialog_kwargs: dict of kwargs for oepn file dialog
    :type file_dialog_kwargs: `kwargs` for `filedialog.askopenfilename`
    :returns: file path
    :rtype: `os` path
    """
    rpath = filedialog.askopenfilename(**file_dialog_kwargs)    #prompt to get the location and name
    if rpath == '': rpath = None
    return rpath

def open_dir_dialog(file_dialog_kwargs):
    """function opens the file picker dialog to prompt the user for a folder to
    open. dialog is configurable based on the passed kwargs. Function takes a dict 
    of kwargs for the tkinter `ask open` dialog
    
    :param file_dialog_kwargs: dict of kwargs for oepn file dialog
    :type file_dialog_kwargs: `kwargs` for `filedialog.askopenfilename`
    :returns: file path
    :rtype: `os` path
    """
    rpath = filedialog.askdirectory(**file_dialog_kwargs)    #prompt to get the location and name
    if rpath == '': rpath = None
    return rpath

def check_dir_exists(dir_path):
    """function checks if the directory exists - ONLY for directories, not files
    
    :param dir_path: absolute file path to the directory to find
    :type dir_path: string
    :returns: true if directory exists
    :rtype: bool
    """
    rval = False                                #temp return value - default false if not found
    if os.path.isdir(dir_path): rval = True     #if found, update to true
    return rval

def check_file_exists(file_path):
    """function checks if the file exists - ONLY for files, not directories
    
    :param file_path: absolute file path to the file to find
    :type file_path: string
    :returns: true if file exists
    :rtype: bool
    """
    rval = False                                #temp return value - default false if not found
    if os.path.isfile(file_path): rval = True   #if found, update to true
    return rval

def get_day_seconds():
    """function returns the total seconds since midnight of the current day
    :returns: seconds since midnight
    :rtype: `int`
    """
    now = datetime.datetime.now()
    midnight = now.replace(hour=0, minute=0, second=0, microsecond=0)
    return (now - midnight).seconds

def get_num_files_in_dir(path):
    """function gets total number of files in the passed directory and its sub-directories
    
    :param dir_path: absolute file path to the directory to find
    :type dir_path: string
    :returns: number of files
    :rtype: `int`
    """
    num_files = 0
    for dirpath, dirnames, filenames in os.walk(path):      #loop through all files in root dir and sub dirs
        for filename in filenames:                              #for each file
            num_files+=1                                        #count number of files
    return num_files

def get_file_ext(file_name):
    """function returns the extension of the passed filename
    """
    return os.path.splitext(file_name)[-1].lower()

#--constnant dict for open type
CONST_openType = {'file':1,
                  'dir':2}

#--constant dict for accepted media property formats
CONST_media_formats = {'MKV': '.mkv'}

#--constant dict for accepted input template formats
CONST_template_formats = {'XLSX': '.xlsx',
                          'XLS': '.xls',
                          'CSV': '.csv'}

#--constant dict for the output CSV file used when parsing files
CONST_CSVexport_hdrs = {'file_name':'File_Name',    #format for dict is {'media_file_props.attr_name':'CSV output File Header'}
                        'full_path':'File_Path',
                        'good_estimate':'Good_Estimate',
                        'est_time_raw':'Runtime_Seconds',
                        'est_time_str':'Runtime_HH:MM:SS'}

#--constant dict for error types
CONST_err_types = {'err':'error',
                   'warn':'warning'}

#--Misc theme constants for labels and buttons
sys_fnt_HDR1 = ('Arial', '20', 'normal', 'roman')
sys_fnt_txt = ('Arial', '14', 'normal', 'roman')
sys_fnt_BTN = ('Arial', '14', 'normal', 'roman')
sys_fnt_small = ('Arial', '8', 'normal', 'roman')

#--file import key columns
sys_tmplt_oldFile_hdrName = 'File_Path'
sys_tmplt_newFile_hdrName = 'New_File_Path'

#--misc constants
sys_wrap_len = 550
sys_err_wndw_width = 600


#-----------------------------main window

class wndw_Main(tk.Tk):
    def __init__(self):
        """Primary application window"""
        tk.Tk.__init__(self)
        self.protocol("WM_DELETE_WINDOW", self.on_close)    #handle window close button
        self.init_main_window()                             #initialize application

    def init_main_window(self):
        """function initiates the various user window elements"""
        self.title("Media Manager Helper")      #title bar
        main_label = tk.Label(self, text='Media Manager Helper', font=sys_fnt_HDR1)
        main_label.grid(row=0,column=0, padx=10, pady=10)
        btn_exp_props = tk.Button(self, text='Export Properties', font=sys_fnt_BTN, command=self.userCMD_parse_files) 
        btn_exp_props.grid(row=1, column=0, padx=10, pady=10)
        btn_upd_files = tk.Button(self, text='Update Files', font=sys_fnt_BTN, command=self.userCMD_update_files) 
        btn_upd_files.grid(row=2, column=0, padx=10, pady=10)
        btn_close = tk.Button(self, text='Close', font=sys_fnt_BTN, command=self.on_close) 
        btn_close.grid(row=6, column=0, padx=10, pady=(10,20))
        
        #--main frame to hide progress bar shenanigans
        self.pb_frame = tk.Frame(self); self.pb_frame.grid(row=3, column=0, padx=10, pady=(20,10))
        self.main_pb_var = tk.IntVar()
        self.main_pb = ttk.Progressbar(self.pb_frame, orient='horizontal', mode='determinate', length=200, maximum=100, variable=self.main_pb_var)
        self.main_pb.grid(row=0, column=0, padx=10, pady=(0,5))
        self.pb_label_var = tk.StringVar()
        self.pb_label = tk.Label(self.pb_frame, text='', font=sys_fnt_small, textvariable=self.pb_label_var)
        self.pb_label.grid(row=1, column=0, padx=10)

        #--secondary frame to "hide" the progress bar
        self.pb_hide_frame = tk.Frame(self); self.pb_hide_frame.grid(row=3, column=0, sticky=tk.NSEW)
        self.progress_bar_enable(False) #hide progress bar and label

    def on_close(self): #make no changes
        """function is called when the window is closed"""
        self.destroy()

    def progress_bar_update(self, kwargs):
        """function updates the status/state of the progress bar and its label"""
        lbl_txt = kwargs.pop('label_text',None)     #pop off label text
        if lbl_txt is not None:                     #if value is present
             self.pb_label_var.set(lbl_txt)         #then update label
        
        val = kwargs.pop('value',None)              #pop off value
        if val is not None:                         #if value set
            self.main_pb_var.set(val)               #then update
        
        self.update_idletasks()                     #last update window
    
    def progress_bar_enable(self, en):
        """function enables/disables the progress bar, effectively hiding it and its label"""
        if en==True: self.pb_hide_frame.grid_remove()   #remove the hide frame
        else: self.pb_hide_frame.grid()                 #re-place the hide frame
        self.update_idletasks()                         #last update window
    
    def userCMD_parse_files(self):
        """function called when user wants to parse media file(s)"""
        opn_type, input_dirpath = self.userCMD_parse_open()             #have user select file or directory
        if opn_type is not None and input_dirpath is not None:
            dialog_opts = {'title':'Output File Directory'} #directory picker options
            output_dirpath=open_dir_dialog(dialog_opts)                 #have user select location to save the output via browser prompt
            if output_dirpath is not None:
                files = self.parse_dir_mediafiles(input_dirpath, opn_type)  #parse file(s) at selected path
                self.create_mediafiles_csv(files, output_dirpath)           #after looping through all files, create a CSV of the output
                messagebox.showinfo("Success", "File(s) successfully parsed and output file created!")
            else: messagebox.showerror("Error", "Output file location not chosen, stopping operation")
        #else user canceled in previous step so no action was taken
    
    def userCMD_update_files(self):
        """function called when user wants to update media file names based on a selected template"""
        prompt_user_update = user_prompt_update(self)                   #prompt user for inputs in new top-level
        files_to_update = self.get_files_to_update(prompt_user_update.template_path)    #get list of files to update
        del_old = prompt_user_update.del_old                            #assign if should delete old
        if len(files_to_update) > 0:                                    #if input template has valid updates
            cont_upd = self.update_files_error_check(files_to_update)   #then do an error check                
            if cont_upd == True:                                        #if user chose to continue or no errors
                self.update_media_files(files_to_update,del_old)        #then perform the move
            else:
                messagebox.showwarning("No Action", "User canceled, no move/update will be performed.")
        else:
            messagebox.showerror("Error","Template not selected or template has no entries")

    def userCMD_parse_open(self):
        """function asks the user if they'd like to parse a single file or directory of files. After getting input
        from the user, 
        
        :param self: Description
        :returns: file or directory choice, selected path
        :rtype: `CONST_openType` entry, `os` path
        """
        opn_type=None; rpath = None                     #temp vars for open type and object path

        promt_user_open = user_prompt_open_type(self)   #prompt user what they'd like to parse (single file or folder)
        opn_type = promt_user_open.result               #get result of prompt

        #--get path depending on the type of open (file or dir)
        if opn_type == CONST_openType['file']:            
            filetypes=[]                                                #temp list for accepted file types
            for k,v in CONST_media_formats.items(): filetypes.append((k,v))   #build list for file picker
            dialog_opts = { 'filetypes':filetypes,
                              'title':'File to Parse'}  
            rpath = open_file_dialog(dialog_opts)
        elif opn_type == CONST_openType['dir']:
            dialog_opts = {'title':'Directory to Parse'} #directory picker options
            rpath = open_dir_dialog(dialog_opts)

        return opn_type, rpath

    def parse_dir_mediafiles(self, path, open_type):
        """function prases the file at the passed path or all files in the passed directory and its sub directories
        
        :param path: file or directory path to process
        :type path: `os` path
        :param open_type: file or directory choice
        :param open_type: `CONST_openType` entry
        :returns: list of file property objects to output
        :rtype: `media_file_props` instance
        """
        tmp_parsed_files = []   #temp list of files found to return information on

        if open_type == CONST_openType['file']:
            filename = os.path.basename(path)
            file_ext = get_file_ext(filename)                       #get the extention
            if file_ext in CONST_media_formats.values():            #if its a valid file type, then process
                dur, est_success = self.get_mediafile_rawdir(path)  #get the file duration
                #--append result to parsed files list
                tmp_parsed_files.append(media_file_props({'name':filename,
                                                          'path':path,
                                                          'runtime':dur,
                                                          'est_ok':est_success}))
            else:
                messagebox.showerror("Error","Selected file type is not supported.")
        elif open_type == CONST_openType['dir']:
            self.progress_bar_enable(True)                              #show progress bar
            self.progress_bar_update({'label_text':'Parsing Files'})        #set label
            file_count = 0; num_files = get_num_files_in_dir(path)          #setup vars for status
            for dirpath, dirnames, filenames in os.walk(path):      #loop through all files in root dir and sub dirs
                for filename in filenames:                              #for each file
                    file_count+=1                                       #update file counter
                    pb_num = round(file_count/num_files*100)            #update file count
                    self.progress_bar_update({'value':pb_num})          #update progress bar

                    file_ext = get_file_ext(filename)                   #get the extention
                    if file_ext in CONST_media_formats.values():              #if its a valid file type, then process
                        file_path = os.path.join(dirpath, filename)         #make the full path for the file
                        dur, est_success = self.get_mediafile_rawdir(file_path) #get the file duration
                        #--append result to parsed files list
                        tmp_parsed_files.append(media_file_props({'name':filename,
                                                                  'path':file_path,
                                                                  'runtime':dur,
                                                                  'est_ok':est_success}))          
            self.progress_bar_enable(False)                              #all done, so hide progress bar
        else:
            messagebox.showerror("Error", "Valid path/object was not selected. Please try again.")

        return tmp_parsed_files

    def get_mediafile_rawdir(self,path):
        """function gets the duration of the file at the passed path
        :param path: file path to process
        :type path: `os` path
        """
        duration_seconds = None #temp return for duration
        est_success = False     #temp return if estimation was successful

        #---first try the primary method of capture
        cap = cv2.VideoCapture(path)                            #open capture path
        frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))    #get the total number of frames
        fps = cap.get(cv2.CAP_PROP_FPS)                         #and the frames per second
        
        if frame_count > 0 and fps > 0:         #if able to get frames, then primary method is OK
            cap.release()                               #release the capture path - not needed anymore
            duration_seconds=round(frame_count/fps,0)   #calculate the duration
            est_success = True                          #and set success
        else:                                   #else use the fallback method
            #sometimes with some codecs/containers fps or frames is unreliable, so as an alternative
            cap.set(cv2.CAP_PROP_POS_AVI_RATIO, 1)              #seek to the end of the file
            duration_msec = cap.get(cv2.CAP_PROP_POS_MSEC)      #and read the milliseconds at the end
            cap.release()                                       #release the capture path
            if duration_msec > 0:
                duration_seconds = round(duration_msec/1000.0,0)    #alternate calculation for duration
                est_success = True                                  #and set success
        #if both fail, reading was not accurate/successful so a None value will be returned

        return duration_seconds, est_success

    def create_mediafiles_csv(self, files, savedir):
        """function creates the output CSV file that contains the parsed file(s) properties
        
        :param files: file properties to generate CSV file
        :type files: list of `media_file_props` instance(s)
        :param savedir: directory path to output created CSV file at
        :type savedir: `os` directory path
        """
        date_str = datetime.datetime.today().strftime('%Y%m%d')         #get current date
        sec_str = str(get_day_seconds())                                #get current day seconds
        outfilename = date_str+'_'+sec_str+'_'+'MediaProperties.csv'    #build output file name
        out_filepath = savedir +'/' + outfilename                       #build output file path

        #--create output CSV
        self.progress_bar_enable(True)                              #show progress bar
        self.progress_bar_update({'label_text':'Creating CSV'})         #set label
        file_count = 0; num_files = len(files)                          #setup vars for status
        with open(out_filepath,'w',newline='') as out_csv:      #make a CSV file
            file_writer = csv.writer(out_csv,delimiter=',')     #and create the writer
            file_header = []    #output header list
            for v in CONST_CSVexport_hdrs.values():
                file_header.append(v)                           #build header output
            file_writer.writerow(file_header)                   #write header
            for file in files:                                  #loop through all file(s)
                file_count+=1                                       #update file counter
                pb_num = round(file_count/num_files*100)        #update file count
                self.progress_bar_update({'value':pb_num})      #update progress bar

                line_out = []
                for k in CONST_CSVexport_hdrs.keys():
                    line_out.append(getattr(file,k))            #make output list
                file_writer.writerow(line_out)                  #and write line
        self.progress_bar_enable(False)                              #all done, so hide progress bar
    
    def update_files_error_check(self, file_tup_list):
        """function checks for errors in the files to update
        
        :param file_tup_list: list of files to update/move and the new distination/name
        :type file_tup_list: `list` of `string``tuples` formatted as [(old_file_path1,new_file_path1),(old_file_2....)...(n)]
        :returns: true/false if user wants to continue updating files (despite any errors
        :rtype: `bool`
        """
        cont_upd = False    #temp continue update status result
        errors = []         #temp error list (msg_type,msg_text)

        self.progress_bar_enable(True)                              #show progress bar
        self.progress_bar_update({'label_text':'Checking For Errors'})    #set label
        file_count = 0; num_files = len(file_tup_list)              #setup vars for status

        new_files = []                                          #temp list for new files
        for tup in file_tup_list: new_files.append(tup[1])      #convert list of JUST new files
 
        for tup in file_tup_list:
            file_count+=1                                       #update file counter
            pb_num = round(file_count/num_files*100)            #update file count
            self.progress_bar_update({'value':pb_num})          #update progress bar
            
            if check_file_exists(tup[0]) == False:
                errors.append((CONST_err_types['err'],'row:'+str(file_count+1)+' | Cannot find input file: '+tup[0]))
            if tup[1] == '':
                errors.append((CONST_err_types['warn'],'row:'+str(file_count+1)+' | Output file is blank, conversion/move will be skipped'))
            if check_file_exists(tup[1]) == True:
                errors.append((CONST_err_types['err'],'row:'+str(file_count+1)+' | Output file already exists: '+tup[1]))
            if tup[1] != '' and new_files.count(tup[1]) > 1:
                errors.append((CONST_err_types['err'],'row:'+str(file_count+1)+' | Output file exists more than once: '+tup[1]))     

        self.progress_bar_enable(False)                              #hide progress bar

        if len(errors) > 0:
            notify_opts = {'title':'Issues in Template',
                           'message':'Potential issues were found in the update template, listed below',
                           'err_list':errors}
            err_msg = wndw_notify(self,notify_opts)
            cont_upd = err_msg.continue_upd     #update with user's response
        else: cont_upd = True   #otherwise no errorsm, OK to update

        return cont_upd
    
    def get_files_to_update(self, tmplt_path):
        """function opens the passed template file and builds the required new/old file list to update
        
        :param tmplt_path: path to the template file for updates
        :type tmplt_path: `os` path
        :returns: list of tuples for new/old files
        :rtype: [(old_filepath_1,new_filepath_1),...,(old_filepath_n,new_filepath_n)]
        """
        rfiles=[] #return list files to update
        if tmplt_path is not None:
            file_ext = os.path.splitext(os.path.basename(tmplt_path))       #get its extension
            if file_ext[1] in list(CONST_template_formats.values()):        #if its a valid extension
                if file_ext == '.csv': data_pd = pd.read_csv(tmplt_path)    #open if CSV
                else: data_pd = pd.read_excel(tmplt_path)                   #open if excel type

                data_pd = data_pd.replace(np.nan, '')                       #convert "NaN" values to None
                for row in data_pd.itertuples(index=False): #iterate through all pands row and make list of tuples to update
                    rfiles.append((getattr(row,sys_tmplt_oldFile_hdrName),getattr(row,sys_tmplt_newFile_hdrName)))          
            
        return rfiles

    def update_media_files(self, file_tup_list, del_old):
        """function updates/moves the files based on the passed list
        
        :param file_tup_list: list of files to update/move and the new distination/name
        :type file_tup_list: `list` of `string``tuples` formatted as [(old_file_path1,new_file_path1),(old_file_2....)...(n)]
        :param del_old: should the old files be deleted?
        :type del_old: `bool` - True to delete old files
        """
        self.progress_bar_enable(True)                                  #show progress bar
        self.progress_bar_update({'label_text':'Updating/Moving Files'})  #set label
        file_count = 0; num_files = len(file_tup_list)                  #setup vars for status

        for tup in file_tup_list:
            file_count+=1                                       #update file counter
            pb_num = round(file_count/num_files*100)            #update file count
            self.progress_bar_update({'value':pb_num})          #update progress bar

            old_file = tup[0]; new_file = tup[1]        #temp new/old file names
            if new_file != '':
                new_file_dir = os.path.dirname(new_file)    #get directory of "new_file"
                if not check_dir_exists(new_file_dir):      #if destination directory doesn't exist
                    #os.mkdir(new_file_dir)                      
                    Path(new_file_dir).mkdir(parents=True, exist_ok=True)   #then make it
                shutil.copy(old_file, new_file) #copy file from "File_Path" to "New_File_Path"
        self.progress_bar_enable(False)                                 #hide progress bar

        #--wait to remove "old" files until done in case they were copied to multiple places or multiple times
        if del_old == True:                             #if set to delete old files
            self.progress_bar_enable(True)                                  #show progress bar
            self.progress_bar_update({'label_text':'Deleting Old Files'})   #set label
            file_count = 0;                                                 #setup vars for status
            for tup in file_tup_list:                   #loop through list
                file_count+=1                                   #update file counter
                pb_num = round(file_count/num_files*100)        #update file count
                self.progress_bar_update({'value':pb_num})      #update progress bar

                old_file = tup[0]; os.remove(old_file)          #and remove old file
            self.progress_bar_enable(False)                                 #hide progress bar

#-----------------------------main loop
if __name__ == "__main__":
    app = wndw_Main()
    app.mainloop()
