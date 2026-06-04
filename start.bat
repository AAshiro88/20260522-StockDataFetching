@echo off
:: 強制指定使用 UTF-8 全域編碼
chcp 65001 >nul

title 即時報價監控系統 

:: 使用最安全的英文顯示進度，100% 避開編碼衝突
echo [System] Activating base env...

:: 使用絕對路徑呼叫 conda 初始化批次檔
call "C:\ProgramData\anaconda3\Scripts\activate.bat" base 2>nul
if errorlevel 1 (
    :: 備用方案：若上述路徑失敗，嘗試 base 的 activate 檔案
    call "C:\ProgramData\anaconda3\Scripts\activate.bat" base 2>nul
)
if errorlevel 1 (
    :: 第二備用方案：嘗試 miniconda 的 activate 檔案
    call "C:\ProgramData\anaconda3\Scripts\activate.bat" base 2>nul
)

echo %CONDA_DEFAULT_ENV% | findstr /i "base" >nul
if errorlevel 1 (
    echo [System] Failed to activate base, trying direct run.
) else (
    echo [System] Environment base activated successfully.
)

echo --------------------------------------------------

if exist "main.py" (
    :: 若環境啟用成功則使用 python，失敗則強制使用 py312 的絕對路徑執行
    echo %CONDA_DEFAULT_ENV% | findstr /i "base" >nul
    if errorlevel 1 (
        "C:\ProgramData\anaconda3\python.exe" main.py
    ) else (
        python main.py
    )
) else (
    echo [Error] Cannot find main.py. Please check your filename.
)

echo --------------------------------------------------
echo Program finished.
pause