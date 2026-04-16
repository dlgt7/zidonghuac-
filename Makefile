# Makefile for 自动化计算工具
# 使用 MinGW-w64 (g++) 编译

CXX = g++
CXXFLAGS = -std=c++17 -mwindows -DNDEBUG -DUNICODE -D_UNICODE
LDFLAGS = -luser32 -lgdi32 -lcomctl32 -lcomdlg32 -lshell32 -lole32

TARGET = 自动化计算工具.exe
TARGET_MINI = 自动化计算工具_mini.exe
SRC = main.cpp

# 默认编译
all: $(TARGET)

# 标准版本
$(TARGET): $(SRC)
	$(CXX) $(CXXFLAGS) -O2 -s $(SRC) -o $(TARGET) $(LDFLAGS)

# 最小体积版本
mini: $(SRC)
	$(CXX) $(CXXFLAGS) -Os -s -ffunction-sections -fdata-sections \
		-Wl,--gc-sections -Wl,--merge-sections \
		$(SRC) -o $(TARGET_MINI) $(LDFLAGS)

# 调试版本
debug: $(SRC)
	$(CXX) -std=c++17 -g -DUNICODE -D_UNICODE $(SRC) -o debug_$(TARGET) $(LDFLAGS)

# 清理
clean:
	del /q *.exe *.obj 2>nul || true

# 帮助
help:
	@echo "可用目标:"
	@echo "  all   - 编译标准版本 (默认)"
	@echo "  mini  - 编译最小体积版本"
	@echo "  debug - 编译调试版本"
	@echo "  clean - 清理生成的文件"

.PHONY: all mini debug clean help
