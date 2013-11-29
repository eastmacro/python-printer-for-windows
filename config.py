# -*- coding: utf-8 -*-'
import configparser 

config = configparser.RawConfigParser()  

con = open('./data.ini',encoding='GBK').read()

config.read_string(con)

liney_1 = config.getint("line","liney_1")
liney_2 = config.getint("line","liney_2")
liney_3 = config.getint("line","liney_3")
liney_4 = config.getint("line","liney_4")


graduation_type      = config.get("main","graduation_type")
graduation_time_year = config.get("main","graduation_time_year")
graduation_time_month = config.get("main","graduation_time_month")
graduation_time_dat = config.get("main","graduation_time_dat")

