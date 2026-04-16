# 自动化计算工具 - PowerShell 编译脚本
# 需要安装 MinGW-w64 或 Visual Studio

$ErrorActionPreference = "Stop"
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "自动化计算工具 - C++ Win32版 编译脚本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$compiler = $null
$compileCmd = $null

# 检查 MinGW-w64 (g++)
if (Get-Command g++ -ErrorAction SilentlyContinue) {
    $compiler = "g++ (MinGW-w64)"
    $compileCmd = {
        param($output)
        g++ -std=c++17 -O2 -s -mwindows -DNDEBUG `
            -DUNICODE -D_UNICODE `
            main.cpp -o $output `
            -luser32 -lgdi32 -lcomctl32 -lcomdlg32 -lshell32 -lole32
    }
}
# 检查 clang++
elseif (Get-Command clang++ -ErrorAction SilentlyContinue) {
    $compiler = "clang++"
    $compileCmd = {
        param($output)
        clang++ -std=c++17 -O2 -s -mwindows -DNDEBUG `
            -DUNICODE -D_UNICODE `
            main.cpp -o $output `
            -luser32 -lgdi32 -lcomctl32 -lcomdlg32 -lshell32 -lole32
    }
}
# 检查 Visual Studio (cl)
elseif (Get-Command cl -ErrorAction SilentlyContinue) {
    $compiler = "cl (Visual Studio)"
    $compileCmd = {
        param($output)
        cl /nologo /W3 /O2 /EHsc /DNDEBUG `
            /D "UNICODE" /D "_UNICODE" `
            /D "WIN32" /D "_WINDOWS" `
            main.cpp `
            /Fe:$output `
            /link user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib `
            /SUBSYSTEM:WINDOWS
    }
}

if (-not $compiler) {
    Write-Host "错误: 未找到C++编译器！" -ForegroundColor Red
    Write-Host ""
    Write-Host "请安装以下任一编译器：" -ForegroundColor Yellow
    Write-Host "  1. MinGW-w64: https://www.mingw-w64.org/" -ForegroundColor Yellow
    Write-Host "  2. Visual Studio: https://visualstudio.microsoft.com/" -ForegroundColor Yellow
    Write-Host "  3. LLVM/Clang: https://llvm.org/" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "按回车键退出"
    exit 1
}

Write-Host "使用编译器: $compiler" -ForegroundColor Green
Write-Host ""
Write-Host "正在编译..." -ForegroundColor Yellow
Write-Host ""

$outputFile = "自动化计算工具.exe"

try {
    & $compileCmd $outputFile
    
    if (Test-Path $outputFile) {
        $fileSize = (Get-Item $outputFile).Length
        $sizeKB = [math]::Round($fileSize / 1KB, 2)
        $sizeMB = [math]::Round($fileSize / 1MB, 2)
        
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "编译成功！" -ForegroundColor Green
        Write-Host "输出文件: $outputFile" -ForegroundColor Green
        Write-Host "文件大小: $fileSize 字节 ($sizeKB KB, $sizeMB MB)" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        
        # 清理中间文件
        Remove-Item *.obj -ErrorAction SilentlyContinue
    } else {
        throw "编译失败，未生成输出文件"
    }
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "编译失败！" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
}

Write-Host ""
Read-Host "按回车键退出"
