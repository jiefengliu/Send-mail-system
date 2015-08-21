#encoding: gbk


from distutils.core import setup
import py2exe

setup(windows=["EmailTest.py", "MainStart.py", "MainWindow.py","smtplib.py"],
      data_files=[("config",["configuration.ini"])])