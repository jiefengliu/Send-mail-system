#encoding: gbk

from distutils.core import setup 
import py2exe
'''
在命令行执行  python setup.py py2exe
'''

setup( windows=["EmailTest.py", "MainStart.py", "MainWindow.py","smtplib.py","ReadWriteFile.py",],
      data_files=[("configuration.ini")],options={"py2exe":{"dll_excludes":["MSVCP90.dll",]}})