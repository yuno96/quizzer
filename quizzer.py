#!/usr/bin/env python3
# -*- coding: UTF-8 -*-

import os
import sys
import errno
import logging
from tkinter import *
#import portalocker
#from tempfile import gettempdir
import openpyxl
from tkinter.messagebox import showinfo
from tkinter import filedialog
import random

class Quizzer:
	def __init__(self, root=None):
		self.root = root
		self.logging = logging
		self.DBPATH = os.path.join(os.getcwd(), 'data')
		self.DBNAME = 'cache'
		self.ICONPATH = os.path.join(os.getcwd(), 'icons')
		#self.LOCK_FILE = os.path.join(gettempdir(), 'quizzer-flock')
		self.HEADLIST_MAX = 64
		self.root.title('quizzer')
		self.q_str = StringVar()
		self.fd_excel = None
		self.qna_pool = {}
		self.qna_idx = 0
		self.qna_sum = 0
		self.right_ans = ''
		self.qna_toggle = False

		self.logging.basicConfig(level=logging.DEBUG,
			format='%(levelname)s:%(filename)s:%(funcName)s:%(lineno)d:%(message)s')

		# MENU
		## File
		self.menubar = Menu(root)
		menu_file = Menu(self.menubar, tearoff=0)
		menu_file.add_command(label='Open DB', command=self.open_db)
		menu_file.add_command(label='Exit', command=self.root.quit)
		self.menubar.add_cascade(label='File', menu=menu_file)
		## Help
		about_menu = Menu(self.menubar, tearoff=0)
		about_menu.add_command(label='About', command=self.show_about)
		self.menubar.add_cascade(label='Help',  menu=about_menu)
		root.config(menu=self.menubar)

		# BODY
		self.q_str.set('Question 00/00')
		tmp = Label(root, textvariable=self.q_str)
		tmp.grid(row=0, column=0)
		sbar = Scrollbar(root)
		sbar.grid(row=0,column=2, sticky=N+S)
		self.text = Text(root, height=3, bg='lightgray', yscrollcommand=sbar.set)
		self.text.grid(row=0, column=1)
		sbar.config(command=self.text.yview)

		tmp = Label(root, text='Answer:')
		tmp.grid(row=1, column=0)
		self.myans = Entry(root)
		self.myans.grid(row=1, column=1, sticky=W+E)
		self.myans.bind('<Return>', self.eval_myans)

		self.ans = Entry(root, bg='lightgray')
		self.ans.grid(row=2, column=1, sticky=W+E)

		# CheckButton
		self.is_random = IntVar()
		self.cb = Checkbutton(root, text='Random', variable=self.is_random)
		self.cb.grid(row=3, column=1)


	def show_about(self):
		about_msg = '''
		\rYou can get the latest version here:
		\rhttps://github.com/yuno96/quizzer.git
		'''
		showinfo('About', about_msg)

	def put_color(self, obj, color):
		obj.config({'background':color})

	def clear_widget(self):
		self.text.delete('1.0', END)
		self.ans.delete(0, END)
		self.myans.delete(0, END)
		self.put_color(self.myans, 'white')

	def clear_all(self):
		self.clear_widget()
		self.right_ans = ''
		self.qna_toggle = False
		self.qna_idx = 0
		self.qna_sum = 0

	def put_qna(self):
		self.clear_widget()
		'''
		if self.qna_idx >= self.qna_sum-1:
			showinfo('info', 'You have done the Quiz.')
			return
		'''
		try:
			alist = self.qna_pool[self.qna_idx]
			self.text.insert(INSERT, alist[0])
			self.q_str.set('Question %02d/%02d:' % 
					(self.qna_idx+1, self.qna_sum))
			self.myans.focus_set()
		except:
			showinfo('info', 'You have done the Quiz.')
			return ''

		return alist[1]

	def open_db(self):
		self.logging.debug('yep')
		fname = filedialog.askopenfilename(initialdir=self.DBPATH, title='Select db file')
		self.logging.debug(fname)
		#fname = os.path.join(self.DBPATH, 'sampledb.xlsx')
		if not os.path.exists(fname):
			self.logging.warning('There is nofile')
			return

		if self.fd_excel and 'close' in dir(self.fd_excel):
			self.fd_excel.close()

		self.clear_all()
		self.fd_excel = openpyxl.load_workbook(fname)
		self.logging.debug(type(self.fd_excel))
		# Get active sheet
		ws = self.fd_excel.active
		#ws = self.fd_excel.get_sheet_by_name('Sheeet')
		self.qna_sum = 0
		tmp_pool = {}
		for r in ws.rows:
			if not r[0].value:
				continue
			tmp_pool[self.qna_sum] = [r[0].value, r[1].value]
			self.qna_sum += 1
			#print ('%s %s' % (r[0].value, r[1].value))
		if 'close' in dir(self.fd_excel):
			self.fd_excel.close()

		if self.is_random.get():
			vals = list(tmp_pool.values())
			random.shuffle(vals)
			self.qna_pool = dict(zip(tmp_pool.keys(), vals))
		else:
			self.qna_pool = tmp_pool

		self.right_ans = self.put_qna().strip()

	def the_answer_is(self, my_ans, right_ans):
		mycolor = 'red'
		if my_ans == right_ans:
			mycolor = 'lightblue'
		self.put_color(self.myans, mycolor) 
		self.ans.insert(INSERT, right_ans)

	def eval_myans(self, event):
		self.logging.debug('toggle=%s idx=%d', self.qna_toggle, self.qna_idx)
		if not self.qna_toggle:
			self.the_answer_is(self.myans.get().strip(), 
					self.right_ans)
		else:
			self.qna_idx += 1
			self.right_ans = self.put_qna().strip()

		self.qna_toggle = not self.qna_toggle

	def run(self):
		self.logging.debug('run')


if __name__ == '__main__':

	root = Tk()
	root.tk.call('encoding', 'system', 'utf-8')
	#root.option_add( "*font", "lucida 9" )
	
	quizzer = Quizzer(root)
	#quizzer = Quizzer(None)
	#quizzer.run()

	root.mainloop()
