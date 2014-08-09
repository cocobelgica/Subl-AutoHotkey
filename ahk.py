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
	
	def __init__(self, target, is_file, listener, ahk_exe, start_dir=False, args="", include_print=False):
		self.ahk_exe = os.path.normpath(ahk_exe)
		self.listener = listener
		
		if is_file:
			self.run_file(target, start_dir, args)
		else:
			# append 'print()' function to AHK code
			# function writes string to stdout which
			# is captured by Sublime Text - useful for debugging
			if include_print:
				ahk_print = ('print(str) {\n'
				             '\tif !DllCall("GetStdHandle", "Int", -11, "Ptr")\n'
				             '\t\treturn false\n'
				             '\tFileAppend, % str . "`n", *\n}')
				target += "\n\n{}".format(ahk_print)
			self.run_pipe(target, start_dir, args)

	def run_file(self, script, start_dir, args=""):
		cmd = [self.ahk_exe, "/ErrorStdOut", script]
		if not start_dir: start_dir = os.path.dirname(script)
		if args:
			cmd.extend(args)
		print("Running " + " ".join(cmd[:3]) +
		      "".join(" {!r}".format(x) for x in args))
		self.start_time = time.time()
		
		self.proc = subprocess.Popen(args=cmd,
		                             cwd=start_dir,
		                             stdout=subprocess.PIPE,
		                             stderr=subprocess.PIPE,
		                             shell=True,
		                             universal_newlines=True)

		if self.proc.stdout:
			threading.Thread(target=self.read_stdout).start()

		if self.proc.stderr:
			threading.Thread(target=self.read_stderr).start()

	def run_pipe(self, code, start_dir, args=""):
		PIPE_ACCESS_OUTBOUND = 0x00000002
		PIPE_UNLIMITED_INSTANCES = 255
		INVALID_HANDLE_VALUE = -1
		
		script_name = "AHK_" + str(windll.kernel32.GetTickCount())
		pipe_name = "\\\\.\\pipe\\" + script_name

		__PIPE_GA_ = windll.kernel32.CreateNamedPipeW(c_wchar_p(pipe_name),
		                                              PIPE_ACCESS_OUTBOUND,
		                                              0,
		                                              PIPE_UNLIMITED_INSTANCES,
		                                              0,
		                                              0,
		                                              0,
		                                              None)

		__PIPE_ = windll.kernel32.CreateNamedPipeW(c_wchar_p(pipe_name),
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

		cmd = [self.ahk_exe, "/ErrorStdOut", pipe_name]
		if not start_dir: start_dir = os.path.expanduser("~")
		if args:
			cmd.extend(args)
		print("Running " + " ".join(cmd[:3]) +
		      "".join(" {!r}".format(x) for x in args))
		self.start_time = time.time()
		
		self.proc = subprocess.Popen(args=cmd,
		                             cwd=start_dir,
		                             stdout=subprocess.PIPE,
		                             stderr=subprocess.PIPE,
		                             shell=True,
		                             universal_newlines=True)
		
		if not self.proc.pid:
			print('Could not open file: "' + pipe_name + '"')

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
	
	def run(self, edit, cmd=None, ahk_exe="C:/Program Files/AutoHotkey/AutoHotkey.exe", **kwargs):
		if cmd is None:
			file_name = self.view.file_name()
			if file_name:
				if self.view.is_dirty():
					cmd = None
					is_file = False
					kwargs['start_dir'] = os.path.dirname(file_name)
				else:
					cmd = file_name
					is_file = True
			else:
				is_file = False

			self.build(cmd, is_file, ahk_exe, **kwargs)

		elif cmd == "$quick_run":
			self.build(None, False, ahk_exe, **kwargs)

		elif cmd == "$help":
			subprocess.Popen(["C:/Windows/hh.exe", "C:/Program Files/AutoHotkey/AutoHotkey.chm"])

		elif cmd == "$win_spy":
			subprocess.Popen(["C:/Program Files/AutoHotkey/AU3_Spy.exe"])

		elif os.path.isfile(cmd):
			self.build(cmd, True, ahk_exe, **kwargs)

		else:
			self.build(cmd, False, ahk_exe, **kwargs)

	def build(self, target, is_file, ahk_exe, **kwargs):
		if is_file:
			if os.path.splitext(target)[1] != ".ahk":
				print("[Finished - Not an AHK script]")
				return False
		
		else:
			if target is None:
				view = self.view
				if view != sublime.active_window().active_view():
					view = sublime.active_window().active_view()

				re.IGNORECASE
				if not re.search("(AutoHotkey|Plain text)", view.settings().get("syntax")):
					print("[Finished - Not an AHK code]")
					return False
				
				target = view.substr(sublime.Region(0, view.size()))

		try:
			self.proc = AhkAsyncProcess(target, is_file, self, ahk_exe, **kwargs)
		except Exception as e:
			self.print_data(None, str(e).encode('utf-8'))

	def print_data(self, proc, data):
		output = data.decode('utf-8')
		output = output.replace("\r\n", "\n").replace("\r", "\n")
		print(output, end="")

	def finish(self, proc):
		elapsed = time.time() - proc.start_time
		exit_code = proc.exit_code()
		while exit_code is None:
			exit_code = proc.exit_code()
		print("[Finished in {:.1f}s with exit code {:d}]".format(elapsed, exit_code))

	def on_data(self, proc, data):
		sublime.set_timeout(functools.partial(self.print_data, proc, data), 0)

	def on_finished(self, proc):
		sublime.set_timeout(functools.partial(self.finish, proc), 0)