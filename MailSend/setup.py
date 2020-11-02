# -*- coding: utf-8 -*-
"""
Created on Mon Feb 10 13:51:28 2020

@author: lenovoz
"""

#cx_freeze ile
from cx_Freeze import  setup,Executable
import sys

build_exe_options = {"packages": ["os","tkinter"],"includes":["tkinter"]}

base = None
if (sys.platform == "win32"):
    base = "Win32GUI"

setup(  name = "MailSendT",
        version = "0.1",
        description = "Mail Gönderme",
        options = {"build_exe": build_exe_options},
        executables = [Executable("MailSendT.py", base=base)])