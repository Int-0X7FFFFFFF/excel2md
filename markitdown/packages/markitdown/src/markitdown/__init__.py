# SPDX-FileCopyrightText: 2024-present Adam Fourney <adamfo@microsoft.com>
#
# SPDX-License-Identifier: MIT

from .__about__ import __version__
from ._base_converter import DocumentConverterResult, DocumentConverter
from ._stream_info import StreamInfo
from ._exceptions import (
    MarkItDownException,
    MissingDependencyException,
    FailedConversionAttempt,
    FileConversionException,
    UnsupportedFormatException,
)

# ``_markitdown`` 在导入时会拉起 magika/onnxruntime/requests 等重依赖。本工具
# 只用 xlsx/xls 转换器、从不实例化 MarkItDown，因此改为惰性加载，让 import 链
# 和 Nuitka 构建都不再包含这些大块头。``from markitdown import MarkItDown``
# 仍可通过模块 ``__getattr__`` 正常使用。
_LAZY_ATTRS = ("MarkItDown", "PRIORITY_SPECIFIC_FILE_FORMAT", "PRIORITY_GENERIC_FILE_FORMAT")


def __getattr__(name: str):
    if name in _LAZY_ATTRS:
        from . import _markitdown

        return getattr(_markitdown, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

__all__ = [
    "__version__",
    "MarkItDown",
    "DocumentConverter",
    "DocumentConverterResult",
    "MarkItDownException",
    "MissingDependencyException",
    "FailedConversionAttempt",
    "FileConversionException",
    "UnsupportedFormatException",
    "StreamInfo",
    "PRIORITY_SPECIFIC_FILE_FORMAT",
    "PRIORITY_GENERIC_FILE_FORMAT",
]
