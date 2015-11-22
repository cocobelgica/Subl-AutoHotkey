import sublime, sublime_plugin
import os, sys
import subprocess
import threading
import functools
import time
import collections

class ProcessListener(object):
    def on_data(self, proc, data):
        pass

    def on_finished(self, proc):
        pass


class AhkAsyncProcess(object):
    def __init__(self, cmd, listener, working_dir='', write=b'', **kwargs):
        self.listener = listener
        self.start_time = time.time()
        
        self.proc = subprocess.Popen(
            args = cmd,
            cwd = working_dir,
            stdin = subprocess.PIPE,
            stdout = subprocess.PIPE,
            stderr = subprocess.PIPE,
            universal_newlines = False
        )

        if len(write) > 0:
            self.proc.stdin.write(write)
            self.proc.stdin.close()

        if self.proc.stdout:
            threading.Thread(target=self.read_stdout).start()

        if self.proc.stderr:
            threading.Thread(target=self.read_stderr).start()

    def kill(self):
        self.proc.terminate()
        self.listener = None

    def exit_code(self):
        return self.proc.poll()

    def read_stdout(self):
        while True:
            data = os.read(self.proc.stdout.fileno(), 2**15)

            if len(data) > 0:
                if self.listener:
                    self.listener.on_data(self, data)
            else:
                self.proc.stdout.close()
                if self.listener:
                    self.listener.on_finished(self)
                break

    def read_stderr(self):
        while True:
            data = os.read(self.proc.stderr.fileno(), 2**15)

            if len(data) > 0:
                if self.listener:
                    self.listener.on_data(self, data)
            else:
                self.proc.stderr.close()
                break


class AhkExecCommand(sublime_plugin.WindowCommand, ProcessListener):
    BLOCK_SIZE = 2**14
    text_queue = collections.deque()
    text_queue_proc = None
    text_queue_lock = threading.Lock()

    proc = None

    def run(self, ahk_exe=None, ahk_script=None, working_dir='', script_args='', codepage=65001, **kwargs):
        # clear the text_queue
        self.text_queue_lock.acquire()
        try:
            self.text_queue.clear()
            self.text_queue_proc = None
        finally:
            self.text_queue_lock.release()
        
        if not hasattr(self, 'output_view'):
            self.output_view = self.window.create_output_panel('ahk_exec')
        
        self.output_view.settings().set('line_numbers', False)
        self.output_view.settings().set('gutter', False)
        self.output_view.settings().set('scroll_past_end', False)
        self.output_view.assign_syntax('Packages/Text/Plain text.tmLanguage')

        # From Default.exec.py
        # Call create_output_panel a second time after assigning the above
        # settings, so that it'll be picked up as a result buffer
        self.window.create_output_panel('ahk_exec')

        if ahk_exe is None:
            ahk_exe = 'C:/Program Files/AutoHotkey/AutoHotkey.exe'
        ahk_exe = os.path.normpath(ahk_exe)
        
        active_view = self.window.active_view()
        if ahk_script is None:
            ahk_script = active_view.file_name() or '*'

        code = ''
        if ahk_script == '*':
            code = active_view.substr(sublime.Region(0, active_view.size()))
        
        self.encoding = 'utf-8'
        if codepage == 0:
            self.encoding = 'mbcs' # Python-specific, Windows only - ANSI codepage(CP_ACP)
        elif codepage == 65001:
            self.encoding = 'utf-8' # alt='cp65001' - new in version 3.3, Windows UTF-8
        elif codepage == 1200:
            self.encoding = 'utf-16-le'
        elif codepage == 1252:
            self.encoding = 'windows-1252'
        
        if working_dir == '':
            if ahk_script == '*':
                working_dir = os.path.expanduser('~')
            else:
                working_dir = os.path.dirname(ahk_script)

        cmd = [ahk_exe, '/CP{:d}'.format(codepage), '/ErrorStdOut', ahk_script]
        if len(script_args) > 0:
            cmd.extend(script_args)

        self.proc = None
        print('Running ' + ' '.join(cmd[:4]) + ''.join(' {!r}'.format(i) for i in script_args))

        self.window.run_command('show_panel', {'panel': 'output.ahk_exec'})

        try:
            self.proc = AhkAsyncProcess(cmd, self, working_dir, code.encode(self.encoding), **kwargs)

            self.text_queue_lock.acquire()
            try:
                self.text_queue_proc = self.proc
            finally:
                self.text_queue_lock.release()

        except Exception as e:
            self.append_string(None, str(e) + '\n')
            self.append_string(None, '[Finished]')

    def append_string(self, proc, str):
        self.text_queue_lock.acquire()

        was_empty = False
        try:
            if proc != self.text_queue_proc:
                if proc:
                    proc.kill()
                return

            if len(self.text_queue) == 0:
                was_empty = True
                self.text_queue.append('')

            available = self.BLOCK_SIZE - len(self.text_queue[-1])

            if len(str) < available:
                cur = self.text_queue.pop()
                self.text_queue.append(cur + str)
            else:
                self.text_queue.append(str)

        finally:
            self.text_queue_lock.release()

        if was_empty:
            sublime.set_timeout(self.service_text_queue, 0)

    def service_text_queue(self):
        self.text_queue_lock.acquire()

        is_empty = False
        try:
            if len(self.text_queue) == 0:
                return

            str = self.text_queue.popleft()
            is_empty = (len(self.text_queue) == 0)
        finally:
            self.text_queue_lock.release()

        self.output_view.run_command('append', {'characters': str, 'force': True, 'scroll_to_end': True})

        if not is_empty:
            sublime.set_timeout(self.service_text_queue, 1)

    def finish(self, proc):
        elapsed = time.time() - proc.start_time
        exit_code = proc.exit_code()
        if exit_code == 0 or exit_code == None:
            self.append_string(proc, '[Finished in {:.1f}s]'.format(elapsed))
        else:
            self.append_string(proc, '[Finished in {:.1f}s with exit code {:d}]'.format(elapsed, exit_code))

        if proc != self.proc:
            return

    def on_data(self, proc, data):
        try:
            str = data.decode(self.encoding)
        except:
            str = '[Decode error - output not ' + self.encoding + ']\n'
            proc = None

        str = str.replace('\r\n', '\n').replace('\r', '\n')

        self.append_string(proc, str)

    def on_finished(self, proc):
        sublime.set_timeout(functools.partial(self.finish, proc), 0)