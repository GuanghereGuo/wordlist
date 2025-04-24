import shutil
import os
import sys
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
import logging


def setup_logging(log_level='INFO'):
    """初始化日志配置"""
    logging.basicConfig(
        filename='word_processor.log',
        level=getattr(logging, log_level.upper()),
        format='%(asctime)s - %(levelname)s - %(message)s'
    )


def create_backup(file_path):
    """创建文件备份"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = os.path.basename(file_path)
    dir_name = os.path.dirname(file_path)
    if not os.path.exists(f"{dir_name}/backup/"):
        os.makedirs(f"{dir_name}/backup/")
    backup_path = (f"{dir_name}/backup/"
                   + f"{file_name}_backup_{timestamp}.xlsx")
    shutil.copy(file_path, backup_path)
    logging.info(f"备份已创建: {backup_path}")
    return backup_path

def set_format(file_path, config):
    wb = load_workbook(file_path)

    # 设置全局字体
    # if config['format']['sort']:
    #     sort_excel(file_path, config)

    font = Font(name=config['format']['font']['name'], size=config['format']['font']['size'])
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                cell.font = font
                cell.alignment = Alignment(wrap_text=True)

    # 保存修改
    wb.save(file_path)
    logging.info(f"字体已设置: {file_path}")
