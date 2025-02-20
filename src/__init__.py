# -*- coding: utf-8 -*-

"""
@File    : __init__.py.py
@Time    : 2025/2/20 17:49
@Author  : lenovo
@Contact : 
@Desc    : 这里填写文件的描述信息
"""

# 导入必要的库
import os
import sys

# 定义常量
CONSTANT_VALUE = 42


# 定义函数
def example_function(param1, param2):
    """
    这是一个示例函数，用于演示如何编写函数文档字符串。

    :param param1: 第一个参数
    :param param2: 第二个参数
    :return: 返回两个参数的和
    """
    return param1 + param2


# 主程序
if __name__ == "__main__":
    # 调用示例函数
    result = example_function(1, 2)
    print(f"结果: {result}")
