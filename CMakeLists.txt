cmake_minimum_required(VERSION 3.15)
project(自动化计算工具 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 设置为Windows GUI程序
if(WIN32)
    set(CMAKE_WIN32_EXECUTABLE TRUE)
endif()

# 源文件
add_executable(AutoCalcTool WIN32 main.cpp)

# 链接库
target_link_libraries(AutoCalcTool PRIVATE
    user32
    gdi32
    comctl32
    comdlg32
    shell32
    ole32
)

# Release优化
if(MSVC)
    target_compile_options(AutoCalcTool PRIVATE
        $<$<CONFIG:Release>:/O2>
        $<$<CONFIG:MinSizeRel>:/O1 /Os /Oy /GF /GL /Gy>
    )
    target_link_options(AutoCalcTool PRIVATE
        $<$<CONFIG:MinSizeRel>:/LTCG /OPT:REF /OPT:ICF /MERGE:.rdata=.text>
    )
elseif(MINGW)
    target_compile_options(AutoCalcTool PRIVATE
        $<$<CONFIG:Release>:-O2>
        $<$<CONFIG:MinSizeRel>:-Os -s -ffunction-sections -fdata-sections>
    )
    target_link_options(AutoCalcTool PRIVATE
        $<$<CONFIG:MinSizeRel>:-Wl,--gc-sections -Wl,--merge-sections>
    )
endif()

# 设置输出名称
set_target_properties(AutoCalcTool PROPERTIES
    OUTPUT_NAME "自动化计算工具"
)

# 安装规则
install(TARGETS AutoCalcTool DESTINATION bin)
