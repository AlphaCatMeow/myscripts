import numpy as np
import pandas as pd
from loguru import logger
import os
import easygui

# 获取脚本所在目录
script_dir = os.path.dirname(os.path.abspath(__file__))
logger.info(f"脚本所在目录: {script_dir}")

# 构建Excel文件的绝对路径
file_path = os.path.join(script_dir, 'excel', '永劫无间-征神之路-灵玦图鉴-收集情况.xlsx')
logger.info(f"尝试读取的文件路径: {file_path}")

# 检查文件是否存在
if not os.path.exists(file_path):
    logger.error(f"文件不存在: {file_path}")
    exit(1)

try:
    # 步骤1: 读取excel目录下的Excel文件
    df = pd.read_excel(file_path, sheet_name='Sheet1', header=0)

    # 步骤2: 显示原始数据的前几行
    logger.info("原始数据:")
    columns = df.columns.tolist()
    for column in columns:
        logger.info(column)
    logger.info(df.head())
    
    while True:
        name = easygui.enterbox(msg="请输入灵玦名称，例如：引雷·青", title="输入灵玦名称")
        
        if name:
            # 使用 DataFrame.map 替代 DataFrame.applymap
            mask = df.map(lambda x: name in str(x) if isinstance(x, str) else False)
            if mask.any().any():
                matched_rows, matched_cols = np.where(mask)
                uncollected = []
                all_collected = True
                
                for row, col in zip(matched_rows, matched_cols):
                    exist_value = df.iloc[row, col + 1]
                    if exist_value != "✓":
                        uncollected.append(df.iloc[row, col])
                        all_collected = False
                
                if all_collected:
                    easygui.msgbox(msg="该灵玦已全部收集，请选择其他灵玦！", title="提示")
                else:
                    uncollected_str = "，".join(uncollected)
                    if easygui.ynbox(msg=f'"{uncollected_str}"未收集，可以选择当前灵玦！\n是否收集？', title="提示"):
                        # 用户点击确定后，创建提示框让用户在uncollected的灵玦中选择当前要收集的是哪个灵玦，允许多选
                        selected_uncollected = easygui.multchoicebox(msg="请选择要收集的灵玦：", title="选择灵玦", choices=uncollected)
                        if selected_uncollected:
                            for row, col in zip(matched_rows, matched_cols):
                                if df.iloc[row, col] in selected_uncollected:
                                    df.iloc[row, col + 1] = "✓"
                                    logger.info(f"{df.iloc[row, col]} 已标记为收集")
                            
                            # 更新到Excel文件中
                            with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
                                df.to_excel(writer, sheet_name='Sheet1', index=False)
                            logger.info("\n处理后的数据已保存到 '永劫无间-征神之路-灵玦图鉴-收集情况.xlsx'")
                        else:
                            logger.info("用户未选择任何灵玦，程序继续")
            else:
                easygui.msgbox(msg=f"未找到 {name}，请确认输入是否正确！", title="提示")
        else:
            logger.info("用户取消输入，程序退出")
            break

except Exception as e:
    logger.error(f"处理Excel文件时发生错误: {str(e)}")