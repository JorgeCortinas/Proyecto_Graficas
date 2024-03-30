@echo off

rem Cierra todas las instancias de Excel
wmic process where name="EXCEL.EXE" terminate

