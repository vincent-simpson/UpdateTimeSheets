@echo off
mysql.exe -uroot -e "DROP DATABASE IF EXISTS users;"
mysql.exe -uroot -e "CREATE DATABASE users;"
mysql.exe -uroot -e "USE users;"

PAUSE
CLS
EXIT