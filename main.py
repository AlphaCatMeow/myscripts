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
        # 使用enterbox()函数创建输入框，并将结果保存到name变量，同时增加下拉菜单，用于选择对应灵玦类型，菜单数据来自columns列表的偶数位
        choices = [columns[i] for i in range(0, len(columns), 2)]
        selected_type = easygui.choicebox(msg="请选择灵玦类型", title="选择类型", choices=choices)
        
        if selected_type:
            # 使用enterbox()函数创建输入框，让用户输入具体的灵玦名称
            name = easygui.enterbox(msg=f"请输入{selected_type}类型的灵玦名称，1级请在名称后加数字1，例如：引雷·青2", title="输入灵玦名称")
            
            # 检查用户是否输入了内容
            if name:
                if name in df[selected_type].values:
                    index = df[df[selected_type] == name].index[0]
                    exist_value = df.iloc[index, df.columns.get_loc(selected_type) + 1]
                    if exist_value == "✓":
                        easygui.msgbox(msg=f"{name} 已收集", title="提示")
                        continue
                    else:
                        if easygui.ynbox(msg=f"是否收集 {name}?", title="提示"):
                            df.iloc[index, df.columns.get_loc(selected_type) + 1] = "✓"
                            logger.info(f"{name} 已标记为收集")
                            # 更新到Excel文件中
                            with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
                                df.to_excel(writer, sheet_name='Sheet1', index=False)
                        continue
                else:
                    easygui.msgbox(msg=f"在{selected_type}中未找到 {name}", title="提示")
                    continue
            else:
                logger.error("输入内容为空!")
                break
        else:
            break

    # 步骤3: (您的处理逻辑)

    # 步骤5: 将处理后的数据保存到新的Excel文件
    # (您的保存逻辑)

    logger.info("\n处理后的数据已保存到 'output.xlsx'")
except Exception as e:
    logger.error(f"处理Excel文件时发生错误: {str(e)}")
