@echo off
for /r %%i in (*.png) do (
    ren "%%i" "%%~ni.jpg"
)