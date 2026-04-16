@echo off
chcp 65001 >nul
echo ========================================
echo 自动化计算工具 - C++ Win32版 编译脚本
echo ========================================
echo.

where cl >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在查找 Visual Studio...
    
    if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" (
        call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
    ) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvars64.bat" (
        call "C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvars64.bat"
    ) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvars64.bat" (
        call "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvars64.bat"
    ) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvars64.bat" (
        call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvars64.bat"
    ) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\VC\Auxiliary\Build\vcvars64.bat" (
        call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\VC\Auxiliary\Build\vcvars64.bat"
    ) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\VC\Auxiliary\Build\vcvars64.bat" (
        call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\VC\Auxiliary\Build\vcvars64.bat"
    ) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\VC\Auxiliary\Build\vcvars64.bat" (
        call "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\VC\Auxiliary\Build\vcvars64.bat"
    ) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\VC\Auxiliary\Build\vcvars64.bat" (
        call "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\VC\Auxiliary\Build\vcvars64.bat"
    ) else (
        echo 错误: 未找到 Visual Studio，请安装 Visual Studio 2017 或更高版本
        pause
        exit /b 1
    )
)

echo.
echo 正在编译...
echo.

cl /nologo /W3 /O2 /EHsc /DNDEBUG ^
   /D "UNICODE" /D "_UNICODE" ^
   /D "WIN32" /D "_WINDOWS" ^
   main.cpp ^
   /Fe:自动化计算工具.exe ^
   /link user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib ^
   /SUBSYSTEM:WINDOWS ^
   /ENTRY:wWinMainCRTStartup

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo 编译成功！
    echo 输出文件: 自动化计算工具.exe
    echo ========================================
    
    for %%A in (自动化计算工具.exe) do echo 文件大小: %%~zA 字节 (约 %%~zAKB)
    
    del /q *.obj >nul 2>&1
) else (
    echo.
    echo ========================================
    echo 编译失败！
    echo ========================================
)

echo.
pause
