@echo off
chcp 65001 >nul
echo ========================================
echo 自动化计算工具 - 最小体积编译
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
echo 正在编译（最小体积优化）...
echo.

cl /nologo /W0 /O1 /Os /Oy /GF /GL /Gy /EHsc /DNDEBUG ^
   /D "UNICODE" /D "_UNICODE" ^
   /D "WIN32" /D "_WINDOWS" ^
   main.cpp ^
   /Fe:自动化计算工具_mini.exe ^
   /link user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib ^
   /SUBSYSTEM:WINDOWS ^
   /ENTRY:wWinMainCRTStartup ^
   /LTCG /OPT:REF /OPT:ICF /MERGE:.rdata=.text /MERGE:.pdata=.rdata /SECTION:.text,ER

if %errorlevel% equ 0 (
    echo.
    echo 正在使用 UPX 压缩...
    where upx >nul 2>&1
    if %errorlevel% equ 0 (
        upx --best --ultra-brute 自动化计算工具_mini.exe
    ) else (
        echo 未找到 UPX，跳过压缩步骤
        echo 可从 https://upx.github.io 下载 UPX 进行进一步压缩
    )
    
    echo.
    echo ========================================
    echo 编译成功！
    echo 输出文件: 自动化计算工具_mini.exe
    echo ========================================
    
    for %%A in (自动化计算工具_mini.exe) do echo 文件大小: %%~zA 字节
    
    del /q *.obj >nul 2>&1
) else (
    echo.
    echo ========================================
    echo 编译失败！
    echo ========================================
)

echo.
pause
