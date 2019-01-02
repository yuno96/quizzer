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

class Quizzer:
	def __init__(self, root=None):
		self.root = root
		self.logging = logging
		self.DBPATH = os.path.join(os.getcwd(), 'data')
		self.DBNAME = 'cache'
		self.ICONPATH = os.path.join(os.getcwd(), 'icons')
		#self.LOCK_FILE = os.path.join(gettempdir(), 'quizzer-flock')
		self.HEADLIST_MAX = 64
		#self.root.title('quizzer')

		self.logging.basicConfig(level=logging.DEBUG,
			format='%(levelname)s:%(filename)s:%(funcName)s:%(lineno)d:%(message)s')

	def run(self):
		self.logging.debug('run')
		wb = openpyxl.load_workbook('sampledb.xlsx')
		# Get active sheet
		ws = wb.active
		#ws = wb.get_sheet_by_name('Sheeet')

		for r in ws.rows:
			if not r[0].value:
				continue
			print ('%s %s' % (r[0].value, r[1].value))


if __name__ == '__main__':

	#root = Tk()
	#root.tk.call('encoding', 'system', 'utf-8')
	#root.option_add( "*font", "lucida 9" )
	
	#quizzer = Quizzer(root)
	quizzer = Quizzer(None)
	quizzer.run()

	#root.mainloop()
