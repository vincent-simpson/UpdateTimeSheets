@echo off
REM This batch file is to create a database named "users" if it doesn't already exist.

REM this will kill any currently running mysql processes on localhost
TASKKILL /F /IM mysqld.exe

set root=c:\xampp\
%root%mysql\bin\mysqld.exe --defaults-file=%root%mysql\bin\my.ini


mysql.exe -uroot -e "DROP DATABASE IF EXISTS users;"
mysql.exe -uroot -e "CREATE DATABASE users;"
mysql.exe -uroot -e "USE users;"

PAUSE
CLS
EXIT