@echo off
echo 正在開始批量清理目錄...

:: 刪除當前目錄下所有子目錄內的 .git 資料夾
for /d /r %%i in (.git) do (
    if exist "%%i" (
        echo 正在刪除: %%i
        rd /s /q "%%i"
    )
)

:: 刪除當前目錄下所有子目錄內的 .idea 資料夾
for /d /r %%i in (.idea) do (
    if exist "%%i" (
        echo 正在刪除: %%i
        rd /s /q "%%i"
    )
)

:: 刪除當前目錄下所有子目錄內的 .venv 資料夾
for /d /r %%i in (.venv) do (
    if exist "%%i" (
        echo 正在刪除: %%i
        rd /s /q "%%i"
    )
)

echo 清理完成！
pause