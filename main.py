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
        # 使用enterbox()函数创建输入框，让用户直接输入灵玦名称
        name = easygui.enterbox(msg="请输入灵玦名称，1级请在名称后加数字1，例如：引雷·青2", title="输入灵玦名称")
        
        # 检查用户是否输入了内容
        if name:
            # 在整个DataFrame中查找匹配的灵玦名称
            mask = df.applymap(lambda x: x == name if isinstance(x, str) else False)
            if mask.any().any():
                row_index, col_index = mask.values.nonzero()
                row_index, col_index = row_index[0], col_index[0]
                exist_value = df.iloc[row_index, col_index + 1]
                if exist_value == "✓":
                    easygui.msgbox(msg=f"{name} 已收集", title="提示")
                else:
                    if easygui.ynbox(msg=f"是否收集 {name}?", title="提示"):
                        df.iloc[row_index, col_index + 1] = "✓"
                        logger.info(f"{name} 已标记为收集")
                        # 更新到Excel文件中
                        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
                            df.to_excel(writer, sheet_name='Sheet1', index=False)
                        logger.info("\n处理后的数据已保存到 '永劫无间-征神之路-灵玦图鉴-收集情况.xlsx'")
            else:
                easygui.msgbox(msg=f"未找到 {name}，请确认输入是否正确！", title="提示")
        else:
            logger.info("用户取消输入，程序退出")
            break

except Exception as e:
    logger.error(f"处理Excel文件时发生错误: {str(e)}")