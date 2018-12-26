@echo off
REM This batch file is to create a database named "users" if it doesn't already exist.

set root=c:\xampp\
%root%mysql\bin\mysqld.exe --defaults-file=%root%mysql\bin\my.ini


mysql.exe -uroot -ptoor -e "DROP DATABASE IF EXISTS users;"
mysql.exe -uroot -ptoor -e "CREATE DATABASE users;"
mysql.exe -uroot -ptoor -e "USE users;"

PAUSE
CLS
EXIT