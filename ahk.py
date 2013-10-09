import sublime, sublime_plugin
import os
import subprocess
import threading
import functools
import re
import time
from ctypes import *

class AhkProcListener(object):
    
    def on_data(self, proc, data):
        pass

    def on_finished(self, proc):
        pass


class AhkAsyncProcess(object):
	
	def __init__(self, target, is_file, listener, ahk_exe):
		self.ahk_exe = ahk_exe
		self.listener = listener
		
		if is_file:
			self.run_file(target)
		else:
			self.run_pipe(target)

	def run_file(self, script):
		self.start_time = time.time()

		self.proc = subprocess.Popen([self.ahk_exe, "/ErrorStdOut", script],
		                             cwd=os.path.dirname(script),
		                             stdout=subprocess.PIPE,
		                             stderr=subprocess.PIPE,
		                             shell=True,
		                             universal_newlines=True)

		if self.proc.stdout:
			threading.Thread(target=self.read_stdout).start()

		if self.proc.stderr:
			threading.Thread(target=self.read_stderr).start()

	def run_pipe(self, code):
		PIPE_ACCESS_OUTBOUND = 0x00000002
		PIPE_UNLIMITED_INSTANCES = 255
		INVALID_HANDLE_VALUE = -1

		self.start_time = time.time()
		
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

		self.proc = subprocess.Popen([self.ahk_exe, "/ErrorStdOut", pipe],
		                             cwd=os.path.expanduser("~"),
		                             stdout=subprocess.PIPE,
		                             stderr=subprocess.PIPE,
		                             shell=True,
		                             universal_newlines=True)
		
		if not self.proc.pid:
			print('Could not open file: "' + pipe + '"')

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
		
		if self.proc.stdout:
			threading.Thread(target=self.read_stdout).start()

		if self.proc.stderr:
			threading.Thread(target=self.read_stderr).start()

	def exit_code(self):
		return self.proc.poll()

	def read_stdout(self):
		while True:
			data = os.read(self.proc.stdout.fileno(), 2**15)

			if len(data) > 0:
				self.listener.on_data(self, data)
			else:
				self.proc.stdout.close()
				self.listener.on_finished(self)
				break

	def read_stderr(self):
		while True:
			data = os.read(self.proc.stderr.fileno(), 2**15)

			if len(data) > 0:
				self.listener.on_data(self, data)
			else:
				self.proc.stderr.close()
				break


class ahkCommand(sublime_plugin.TextCommand, AhkProcListener):
	
	def run(self, edit, param=None, ahk_exe="C:/Program Files/AutoHotkey/AutoHotkey.exe"):
		if param is None:
			file_name = self.view.file_name()
			if file_name:
				param = file_name
				is_file = True
			else:
				is_file = False

			self.build(param, is_file, ahk_exe)

		elif param == "$help":
			subprocess.Popen(["C:/Windows/hh.exe", "C:/Program Files/AutoHotkey/AutoHotkey.chm"])

		elif param == "$win_spy":
			subprocess.Popen(["C:/Program Files/AutoHotkey/AU3_Spy.exe"])

		elif os.path.isfile(param):
			self.build(param, True, ahk_exe)

		else:
			self.build(param, False, ahk_exe)

	def build(self, target, is_file, ahk_exe):
		if is_file:
			if os.path.splitext(target)[1] != ".ahk":
				print("[Finished - Not an AHK script]")
				return False
		
		else:
			if target is None:
				re.IGNORECASE
				if not re.search("(AutoHotkey|Plain text)", self.view.settings().get('syntax')):
					print("[Finished - Not an AHK code]")
					return False
				
				target = self.view.substr(sublime.Region(0, self.view.size()))
				if self.view != sublime.active_window().active_view():
					target = target.replace("\\n", "\n")
		
		try:
			self.proc = AhkAsyncProcess(target, is_file, self, ahk_exe)
		except Exception as e:
			self.print_data(None, str(e))

	def print_data(self, proc, data):
		output = data.decode("UTF-8")
		output = output.replace("\r\n", "\n").replace("\r", "\n")
		print(output, end="")

	def finish(self, proc):
		elapsed = time.time() - proc.start_time
		exit_code = proc.exit_code()
		print("[Finished in %.1fs with exit code %d]" % (elapsed, exit_code))

	def on_data(self, proc, data):
		sublime.set_timeout(functools.partial(self.print_data, proc, data), 0)

	def on_finished(self, proc):
		sublime.set_timeout(functools.partial(self.finish, proc), 0)