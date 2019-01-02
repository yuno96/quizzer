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
		self.qna_pool = {}
		self.q_str = StringVar()

		self.logging.basicConfig(level=logging.DEBUG,
			format='%(levelname)s:%(filename)s:%(funcName)s:%(lineno)d:%(message)s')

		# MENU
		## File
		self.menubar = Menu(root)
		menu_file = Menu(self.menubar)
		menu_file.add_command(label='Open DB', command=self.open_db)
		menu_file.add_command(label='Open DB', command=self.root_quit)
		self.menubar.add_cascade(label='File', menu=menu_file)
		## Help
		about_menu = Menu(self.menubar)
		about_menu.add_command(label='About', command=self.show_about)
		self.menubar.add_cascade(label='Help',  menu=about_menu)
		root.config(menu=self.menubar)

		# BODY
		self.q_str.set('Q0:Q0')
		tmp = Label(root, textvariable=self.q_str)
		tmp.pack(side=LEFT)
		sbar = Scrollbar(root)
		sbar.pack(side=Right, fill=Y)
		self.text = Text(root, yscrollcommand=scrollbar.set)
		self.text.pack(side=LEFT, fill=BOTH)
		sbar.config(command=self.text.yview)

		tmp = Label(root, Text='Answer:')
		tmp.pack(side=LEFT)
		self.myans = Entry(root)
		self.myans.pack(side=LEFT, fill=BOTH)
		self.myans.bind('<Enter>', self.eval_myans)

		self.ans = Entry(root)
		self.ans.pack(side=LEFT, fill=BOTH)

        def show_about(self):
		about_msg = '''
		\rYou can get the latest version here:
		\rhttps://github.com/yuno96/quizzer.git
		'''
		showinfo('About', about_msg)

	def open_db(self):
		self.logging.debug('yep')
		fname = os.path.join(self.DBPATH, 'sampledb.xlsx')
		if not os.path.exists(fname):
			self.logging.warning('There is not file')

		if not self.fd_excel:
			self.fd_excel.close()

		wb = openpyxl.load_workbook(fname)
		# Get active sheet
		ws = wb.active
		#ws = wb.get_sheet_by_name('Sheeet')
		cnt = 0
		for r in ws.rows:
			if not r[0].value:
				continue
			self.qna_pool.append({cnt:[r[0].value, r[1].value]})
			cnt += 1
			print ('%s %s' % (r[0].value, r[1].value))
		wb.close()

		self.q_str.set('Question 0/%d:'%cnt)

	def eval_myans(self):
		self.logging.debug('yep')


	def run(self):
		self.logging.debug('run')


if __name__ == '__main__':

	root = Tk()
	root.tk.call('encoding', 'system', 'utf-8')
	#root.option_add( "*font", "lucida 9" )
	
	#quizzer = Quizzer(root)
	quizzer = Quizzer(None)
	#quizzer.run()

	root.mainloop()
