@echo off
::查找目录下的所有ui文件

::当前目录
set root=%cd%
call:uicFile %root%

::子目录
for /f %%a in ('dir /ad/b') do (
    call:uicFile %root%\%%a\
)
pause

::查找给定目录下的文件
:uicFile
for %%a in (%~1*.ui) do (
    ::使用pyuic5处理ui文件为py文件,格式按照Qt的默认格式ui_XXX.py
    pyuic5.exe -o %~1ui_%%~na.py %%a
    echo build: %%a 
    echo   ==^>^> %~1ui_%%~na.py 
    echo.
)
goto:eof
