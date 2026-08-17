"""excel2md Nuitka onefile 打包脚本（在 Windows 上运行）。

用法:
    python build.py

产物: 单个可执行文件 excel2md.exe（--onefile，运行时才解包到临时目录，
       ‌自带依赖与数据文件；替代旧的 standalone 目录产物 build/main.dist/）。

注意:
    - Nuitka 不支持交叉编译，必须在目标平台（Windows）上执行。
    - 需在 Windows 上建一个 Python 3.13 venv 并可编辑安装本 workspace
      （uv sync 或 pip install -e markitdown/packages/markitdown[xlsx]）。
    - onefile + multiprocessing(spawn)：Windows 下每个 worker 子进程会重新解包
      一次 onefile（启动略慢但可用）；main.py 已调用 freeze_support()。
    - magika 通过 __file__ 相对路径加载模型文件，必须 --include-package-data。
    - onnxruntime 有原生 DLL；其 transformers/training 子包依赖 torch 等重依赖，
      用 --nofollow-import-to 排除。
"""

import os
import subprocess
import sys

# 并行编译 job 数：取 CPU 核数（上限 8）。Nuitka 默认可能单核，
# 显式设置能大幅加速最耗时的 C 编译阶段；内存紧张可调低上限。
JOBS = max(1, min(os.cpu_count() or 2, 8))

FLAGS = [
    sys.executable,
    "-m",
    "nuitka",
    "--onefile",
    f"--jobs={JOBS}",
    "--assume-yes-for-downloads",  # 允许 Nuitka 在 Windows 下载 MinGW64/MSVC
    "--output-dir=build",
    "--output-filename=excel2md.exe",
    # markitdown/__init__ 已惰性化 _markitdown，magika/onnxruntime/requests 在
    # 运行时不再被导入；这里显式不跟随，避免 Nuitka 沿惰性路径仍去编译它们
    # （onnxruntime/magika 是最大的编译单元）。
    "--nofollow-import-to=magika",
    "--nofollow-import-to=onnxruntime",
    "--nofollow-import-to=requests",
    # numpy 仅被 openpyxl.compat.numbers 以 try 守卫方式可选导入，非必需；
    # pandas 已移除，故 numpy 也不需要，排除后可进一步缩小构建/产物。
    "--nofollow-import-to=numpy",
    # Excel 读取（pandas 已不再被转换器使用，故不包含 numpy/pandas）
    "--include-package=openpyxl",
    "--include-package-data=openpyxl",
    "--include-package=xlrd",
    # markitdown xlsx 转换链（HtmlConverter → markdownify/bs4）
    "--include-package=charset_normalizer",
    "--include-package-data=charset_normalizer",
    "--include-package=markdownify",
    "--include-package=bs4",
    "--include-package=soupsieve",
    "--include-package=defusedxml",
    # 本仓库
    "--include-package=markitdown",
    "main.py",
]


def main() -> int:
    return subprocess.call(FLAGS)


if __name__ == "__main__":
    raise SystemExit(main())
