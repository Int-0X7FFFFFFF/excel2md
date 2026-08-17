# excel2md — 构建 / 运行 / 清理 的快捷入口
# 用法：make help 查看所有命令
# Linux/本机：直接 make build；Windows 走 CI（.github/workflows/build-windows.yml）

PYTHON ?= .venv/bin/python
# onefile 产物为单文件，直接落在 build/ 下（区别于 standalone 的 build/main.dist/）
EXE    := build/excel2md.exe

.PHONY: all sync build run smoke clean help

all: build           ## 同步依赖并构建

sync:                ## 按 uv.lock 同步依赖
	uv sync

build: sync          ## Nuitka onefile 构建（产物 build/excel2md.exe）
	$(PYTHON) build.py

run:                 ## 用源码运行转换工具（input/ → output/）
	$(PYTHON) main.py

smoke: build         ## 用编译产物对 input/ 冒烟，输出到 /tmp（不污染工作区）
	rm -rf /tmp/excel2md_smoke_out /tmp/excel2md_smoke_state.json
	$(EXE) --input input --output /tmp/excel2md_smoke_out --state /tmp/excel2md_smoke_state.json
	@echo "-> 输出目录: /tmp/excel2md_smoke_out"

clean:               ## 清理构建产物（build/）
	rm -rf build

help:                ## 列出所有命令
	@grep -E '^[a-zA-Z_-]+:.*## ' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*## "}; {printf "  %-8s %s\n", $$1, $$2}'
