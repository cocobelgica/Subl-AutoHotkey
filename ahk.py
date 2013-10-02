import sublime, sublime_plugin
import subprocess
import os
import re
from ctypes import *

class ahk(sublime_plugin.TextCommand):

	def build(self, ahk_exe):
		filename = self.view.file_name()
		if filename:
			pid = self.run_script(filename, ahk_exe)
			print_msg = "[BUILD] <" + (filename if pid else "Cancelled=Not an AHK file") + ">"

		else:
			re.IGNORECASE
			if re.search("(AutoHotkey|AHK|Plain text)", self.view.settings().get('syntax')):
				code = self.view.substr(sublime.Region(0, self.view.size()))
				
				# in case current view is the sublime console
				# usually happens when command is launched via
				# key binding while sublime console has focus
				if self.view != sublime.active_window().active_view():
					# code = re.sub("\\\\n", "\n", code)
					code = code.replace("\\n", "\n")
				
				pid = self.run_code(code, ahk_exe)
				print_msg = "[BUILD] <PID=" + str(pid) + ">"
			
			else: print_msg = "[BUILD] <Cancelled=Not an AHK code>"
		
		print(print_msg)

	
	def run_script(self, script, ahk_exe):
		if os.path.splitext(script)[1] != ".ahk":
			return False
		
		sp = subprocess.Popen([ahk_exe, script], cwd=os.path.dirname(script))
		return sp.pid

	
	def run_code(self, code, ahk_exe):
		PIPE_ACCESS_OUTBOUND = 0x00000002
		PIPE_UNLIMITED_INSTANCES = 255
		INVALID_HANDLE_VALUE = -1

		pipename = "AHK_" + str(windll.kernel32.GetTickCount())
		pipe = "\\\\.\\pipe\\" + pipename

		__PIPE_GA_ = windll.kernel32.CreateNamedPipeW(c_wchar_p(pipe),
		                                              PIPE_ACCESS_OUTBOUND,
		                                              0,
		                                              PIPE_UNLIMITED_INSTANCES,
		                                              0,
		                                              0,
		                                              0,
		                                              None)

		__PIPE_ = windll.kernel32.CreateNamedPipeW(c_wchar_p(pipe),
		                                           PIPE_ACCESS_OUTBOUND,
		                                           0,
		                                           PIPE_UNLIMITED_INSTANCES,
		                                           0,
		                                           0,
		                                           0,
		                                           None)

		if (__PIPE_ == INVALID_HANDLE_VALUE or __PIPE_GA_ == INVALID_HANDLE_VALUE):
			print("Failed to create named pipe.")
			return False

		sp = subprocess.Popen([ahk_exe, pipe], cwd=os.path.expanduser("~"))
		pid = sp.pid
		if not pid:
			print('Could not open file: "' + pipe + '"')
			return False

		windll.kernel32.ConnectNamedPipe(__PIPE_GA_, None)
		windll.kernel32.CloseHandle(__PIPE_GA_)
		windll.kernel32.ConnectNamedPipe(__PIPE_, None)
		
		script = chr(0xfeff) + code
		written = c_ulong(0)
		
		fSuccess = windll.kernel32.WriteFile(__PIPE_,
		                                     script,
		                                     (len(script)+1)*2,
		                                     byref(written),
		                                     None)
		if not fSuccess:
			return False

		windll.kernel32.CloseHandle(__PIPE_)
		return pid


class ahkCommand(ahk):
	
	def run(self, edit, param=None, ahk_exe='C:\\Program Files\\AutoHotkey\\AutoHotkey.exe'):
		if param is None:
			self.build(ahk_exe)
		
		elif param == "$help":
			subprocess.Popen(["C:\\Windows\\hh.exe", "C:\\Program Files\\AutoHotkey\\AutoHotkey.chm"])

		elif param == "$win_spy":
			subprocess.Popen(["C:\\Program Files\\AutoHotkey\\AU3_Spy.exe"])

		elif os.path.isfile(param):
			self.run_script(param, ahk_exe)

		else:
			self.run_code(param.replace("\\n", "\n"), ahk_exe)