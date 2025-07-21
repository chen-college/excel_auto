import pandas as pd
import numpy as np
import os
import time
import multiprocessing
import sys
import logging
from concurrent.futures import ProcessPoolExecutor, as_completed
from collections import defaultdict
from openpyxl.utils import get_column_letter
from datetime import datetime
from concurrent_log_handler import ConcurrentRotatingFileHandler  # 关键修复

def setup_logging(log_dir):
    """配置多进程安全的日志系统"""
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"excel_scan_{timestamp}.log")
    
    # 创建日志记录器
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # 创建多进程安全的文件处理器[6,8,10](@ref)
    file_handler = ConcurrentRotatingFileHandler(
        log_file, 
        mode='a', 
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    
    # 创建控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    
    # 设置日志格式
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

def scan_sheet_for_empty_cells(args):
    """扫描单个sheet页的空单元格并返回结果"""
    file_path, sheet_name, logger = args
    empty_cells = []
    
    try:
        # 使用pandas批量读取整个sheet页数据
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=None,
            dtype=str,
            engine='openpyxl'
        )
        
        # 将空白字符串转换为NaN
        df = df.replace(r'^\s*$', np.nan, regex=True)
        
        # 获取空值位置
        empty_positions = np.argwhere(df.isnull().values)
        
        # 转换位置为Excel坐标
        for row_idx, col_idx in empty_positions:
            row_num = row_idx + 1
            col_letter = get_column_letter(col_idx + 1)
            cell_ref = f"{col_letter}{row_num}"
            empty_cells.append(cell_ref)
            
    except Exception as e:
        logger.error(f"处理工作表 '{sheet_name}' 时出错: {str(e)}")
    
    return sheet_name, empty_cells

def parallel_static_data_check(file_path, logger):
    """全sheet页并行静态数据检查"""
    start_time = time.time()
    empty_cells_dict = defaultdict(list)
    
    try:
        if not os.path.exists(file_path):
            logger.error(f"文件不存在: {file_path}")
            return empty_cells_dict
        
        # 获取所有sheet名称
        xl = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_names = xl.sheet_names
        xl.close()
        
        logger.info(f"工作簿包含 {len(sheet_names)} 个工作表，启动并行扫描...")
        
        # 进程池并行处理
        max_workers = min(multiprocessing.cpu_count(), len(sheet_names))
        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有sheet扫描任务
            futures = {
                executor.submit(
                    scan_sheet_for_empty_cells, 
                    (file_path, name, logger)
                ): name for name in sheet_names
            }
            
            # 收集结果
            for future in as_completed(futures):
                sheet_name, empty_cells = future.result()
                if empty_cells:
                    empty_cells_dict[sheet_name] = empty_cells
    
    except Exception as e:
        logger.error(f"处理Excel时发生全局错误: {str(e)}")
    
    # 输出结果
    logger.info("\n" + "="*60)
    logger.info("Excel空值扫描结果汇总")
    logger.info("="*60)
    
    total_empty = 0
    for sheet_name, cells in empty_cells_dict.items():
        logger.info(f"\n工作表: '{sheet_name}'")
        logger.info(f"  空单元格数量: {len(cells)}")
        
        # 每行打印5个单元格位置
        for i in range(0, len(cells), 5):
            line_cells = cells[i:i+5]
            logger.info("   " + "     ".join(f"{cell_ref:<5}" for cell_ref in line_cells))
        
        total_empty += len(cells)
    
    logger.info(f"\n总计发现空单元格: {total_empty} 个")
    
    # 计算执行时间
    end_time = time.time()
    execution_time = end_time - start_time
    logger.info(f"\n扫描总耗时: {execution_time:.4f} 秒")
    
    return dict(empty_cells_dict)

if __name__ == "__main__":
    # Windows系统必需设置
    multiprocessing.freeze_support()
    
    # 文件路径
    file_path = r'C:\Users\wjy17\Desktop\Excel_Scripts\#竞技场神兽配置表战力.xlsx'
    
    # 日志目录
    log_dir = r'C:\Users\wjy17\Desktop\Excel_Scripts'
    
    # 设置日志系统
    logger = setup_logging(log_dir)
    
    # 记录启动信息
    logger.info(f"开始Excel扫描任务: {file_path}")
    logger.info(f"日志文件将保存在: {log_dir}")
    
    # 执行并行扫描
    result = parallel_static_data_check(file_path, logger)
    
    # 记录完成信息
    logger.info("Excel扫描任务完成")