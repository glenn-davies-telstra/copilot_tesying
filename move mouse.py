# -*- coding: utf-8 -*-
"""
Created on Thu Mar 26 13:31:19 2020

@author: d284876
"""



import ctypes


ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)
input('{Press enter to exit}')
ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
