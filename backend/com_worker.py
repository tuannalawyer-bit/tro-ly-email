"""Chạy toàn bộ lời gọi Outlook COM trên MỘT thread STA duy nhất.

pywebview gọi mỗi phương thức js_api trên một thread khác nhau. COM object của
Outlook thuộc về apartment (STA) nơi nó được tạo ra; dùng nó từ thread khác mà
không marshalling sẽ ném lỗi "The application called an interface that was
marshalled for a different thread".

Giải pháp chuẩn: một ThreadPoolExecutor đúng 1 worker, gọi CoInitialize() ngay
khi thread khởi động. Mọi thao tác Outlook đều đi qua com_call().
"""
from __future__ import annotations

import atexit
import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable

import pythoncom

logger = logging.getLogger(__name__)


def _initializer() -> None:
    pythoncom.CoInitialize()
    logger.info("Thread COM đã khởi tạo apartment STA.")


_executor = ThreadPoolExecutor(
    max_workers=1,                      # BẮT BUỘC là 1
    initializer=_initializer,
    thread_name_prefix="outlook-com",
)


def com_call(fn: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
    """Thực thi fn trên đúng thread sở hữu các COM object.

    Chặn cho tới khi có kết quả; ngoại lệ được ném lại nguyên vẹn cho phía gọi.
    """
    return _executor.submit(fn, *args, **kwargs).result()


@atexit.register
def _shutdown() -> None:
    _executor.shutdown(wait=False)
