import sys
import os


def resource_path(relative_path: str) -> str:
    """
    获取资源文件的绝对路径，兼容开发环境和 PyInstaller 打包后的环境。

    * 在打包模式（存在 ``sys._MEIPASS``）下，返回可执行文件同级目录中
      的资源路径。使用 ``os.path.dirname(sys.executable)`` 获取 exe
      所在目录，而非使用 ``sys._MEIPASS``，以便将数据库和其他可写
      资源文件保存在可执行文件所在目录。
    * 在开发环境下，返回当前模块文件所在目录与 ``relative_path`` 的
      拼接路径，这样资源文件位于仓库代码结构中，便于开发调试。

    参数:
        relative_path: 相对资源文件名或路径。

    返回:
        str: 资源文件的绝对路径。
    """
    if hasattr(sys, '_MEIPASS'):
        # 在 PyInstaller 打包环境下使用可执行文件目录
        base_path = os.path.dirname(sys.executable)
    else:
        # 在开发模式下使用当前文件所在目录
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)
