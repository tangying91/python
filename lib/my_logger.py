
import logging
import os
import time

# 日志文件
log_path = 'log.txt'

# 如果日志文件存在，则删除
if os.path.exists(log_path):
    os.remove(log_path)

# 定义日志目录
logging.basicConfig(filename=log_path, level=logging.INFO)


# 输出日志
def print_log(content):
    print(content)
    logging.info(time.strftime('%Y-%m-%d %H:%M:%S') + content)
    return
