"""统一日志与 GUI 日志队列。

日志同时写入标准 logging，并通过队列由主线程刷新到 Tkinter 日志窗口。
"""

import logging
import queue


log_queue = queue.Queue()


class GuiLogHandler(logging.Handler):
    """把格式化后的日志消息放入队列，避免后台线程直接操作 Tkinter。"""

    def emit(self, record):
        log_queue.put(self.format(record))


logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", datefmt="%H:%M:%S")
gui_handler = GuiLogHandler()
gui_handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))
logging.getLogger().addHandler(gui_handler)


def print_log(msg):
    """统一日志入口，写入标准 logging 和 GUI 日志队列。"""
    logging.info(msg)
