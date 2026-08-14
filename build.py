"""excel2md Nuitka standalone 打包脚本（在 Windows 上运行）。

用法:
    python build.py

产物: build/excel2md.dist/excel2md.exe（standalone，自带依赖与数据文件）。

注意:
    - Nuitka 不支持交叉编译，必须在目标平台（Windows）上执行。
    - 需在 Windows 上建一个 Python 3.13 venv 并可编辑安装本 workspace
      （uv sync 或 pip install -e markitdown/packages/markitdown[xlsx]）。
    - magika 通过 __file__ 相对路径加载模型文件，必须 --include-package-data。
    - onnxruntime 有原生 DLL；其 transformers/training 子包依赖 torch 等重依赖，
      用 --nofollow-import-to 排除。
"""

import subprocess
import sys

FLAGS = [
    sys.executable,
    "-m",
    "nuitka",
    "--standalone",
    "--assume-yes-for-downloads",  # 允许 Nuitka 在 Windows 下载 MinGW64/MSVC
    "--output-dir=build",
    "--output-filename=excel2md.exe",
    "--enable-plugin=numpy",
    # magika: 模型数据文件（config/*.json + models/standard_v3_3/*.onnx）
    "--include-package=magika",
    "--include-package-data=magika",
    # magika 依赖 onnxruntime（含原生 DLL）
    "--include-package=onnxruntime",
    "--include-package-data=onnxruntime",
    "--nofollow-import-to=onnxruntime.transformers",
    "--nofollow-import-to=onnxruntime.training",
    # numpy + pandas（编译扩展 + 共享库）
    "--include-package=numpy",
    "--include-package-data=numpy",
    "--include-package=pandas",
    "--include-package-data=pandas",
    # Excel 读取
    "--include-package=openpyxl",
    "--include-package-data=openpyxl",
    "--include-package=xlrd",
    # requests + CA 证书
    "--include-package=requests",
    "--include-package-data=requests",
    "--include-package=certifi",
    "--include-package-data=certifi",
    # markitdown 依赖
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
