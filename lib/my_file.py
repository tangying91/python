import glob
import os
import warnings

warnings.filterwarnings('ignore')


# 根据路径和文件格式获取所有文件
def get_files(output_file_path, file_suffix):
    search_pattern = os.path.join(output_file_path, file_suffix)
    return glob.glob(search_pattern)

