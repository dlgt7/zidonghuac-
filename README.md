# 自动化计算工具 - C++ Win32版

## 简介

这是自动化计算工具的C++ Win32 API原生实现版本，相比Python版本具有以下优势：

- **体积小**：编译后仅约 **100-300 KB**（Python打包约15-25 MB）
- **启动快**：原生代码，无需解释器
- **无依赖**：单文件可执行，无需安装任何运行时
- **跨平台编译**：支持MSVC、MinGW、Clang

## 功能

- 打包数据查询
- 状态字查询
- 内存映象网计算
- 模拟量计算（西门子模拟量转换）
- 速度转换计算

## 编译方法

### 方法一：使用 PowerShell 脚本（推荐）

```powershell
# 标准编译
.\build.ps1

# 最小体积编译
.\build_mini.ps1
```

### 方法二：使用批处理脚本

```batch
:: 标准编译
build.bat

:: 最小体积编译
build_mini.bat
```

### 方法三：使用 CMake

```bash
mkdir build
cd build
cmake ..
cmake --build . --config Release
```

### 方法四：使用 Make（MinGW）

```bash
make        # 标准版本
make mini   # 最小体积版本
```

### 方法五：直接编译

**MSVC:**
```batch
cl /O2 /EHsc main.cpp /Fe:自动化计算工具.exe /link user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib /SUBSYSTEM:WINDOWS
```

**MinGW:**
```bash
g++ -O2 -s -mwindows main.cpp -o 自动化计算工具.exe -luser32 -lgdi32 -lcomctl32 -lcomdlg32 -lshell32 -lole32
```

## GitHub Actions 自动编译

本项目已配置GitHub Actions，推送到GitHub后会自动编译：

1. Push代码到 `main` 或 `master` 分支
2. GitHub Actions 自动触发编译
3. 编译完成后自动创建Release，包含：
   - `自动化计算工具.exe` - 标准版本
   - `自动化计算工具_mini.exe` - 最小体积版本

### 手动触发编译

在GitHub仓库页面：
1. 点击 `Actions` 标签
2. 选择 `Build Windows Executable`
3. 点击 `Run workflow`

## 文件结构

```
cpp_version/
├── main.cpp           # 主程序源码
├── json.h             # JSON解析器（单文件）
├── resource.h         # 资源定义
├── CMakeLists.txt     # CMake配置
├── Makefile           # Make配置
├── build.ps1          # PowerShell编译脚本
├── build_mini.ps1     # 最小体积编译脚本
├── build.bat          # 批处理编译脚本
├── build_mini.bat     # 最小体积批处理脚本
└── .github/
    └── workflows/
        └── build.yml  # GitHub Actions配置
```

## 编译器要求

需要以下任一编译器：

1. **Visual Studio** (推荐)
   - VS 2017 或更高版本
   - 需安装"使用C++的桌面开发"工作负载

2. **MinGW-w64**
   - 下载：https://www.mingw-w64.org/

3. **LLVM/Clang**
   - 下载：https://llvm.org/

## 进一步压缩体积

安装 UPX 后运行 `build_mini.ps1` 可进一步压缩：

```powershell
# 下载UPX
winget install upx

# 或从 https://upx.github.io 下载

# 编译并压缩
.\build_mini.ps1
```

## 开发者

德龙轧钢自动化团队

## 版本

v3.0 (C++ Win32版)
