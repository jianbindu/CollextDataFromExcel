import os
import pandas as pd

# 设置包含Excel文件的主目录路径
directory = r'Z:\33 licenses'

# 配置不同类型文件的输出文件和关键字
configurations = [
    {
        "keyword": "UNDERGROUND OP LICENSE",
        "output_file": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\UGoperator.xlsx',
        "collected_files_output": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\UGoperatorcollected.xlsx'
    },
    {
        "keyword": "TEST AND MOVE",
        "output_file": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\TestMove.xlsx',
        "collected_files_output": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\TestMovecollected.xlsx'
    },
    {
        "keyword": "WORK AT HEIGHT",
        "output_file": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\TagWAH.xlsx',
        "collected_files_output": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\TagWAHcollected.xlsx'
    },
    {
        "keyword": "UNDERGROUND LDV",
        "output_file": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\UGLDV.xlsx',
        "collected_files_output": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\UGLDVcollected.xlsx'
    },
    {
        "keyword": "SURFACE OP",
        "output_file": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\SurfaceOP.xlsx',
        "collected_files_output": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\SurfaceOPcollected.xlsx'
    },
    {
        "keyword": "HOT WORK",
        "output_file": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\HotWork.xlsx',
        "collected_files_output": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\HotWorkcollected.xlsx'
    },
    {
        "keyword": "SMELTER",
        "output_file": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\Smelter.xlsx',
        "collected_files_output": r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\Smeltercollected.xlsx'
    }
]

# 预期的表头（包含可能的列名变体）
expected_columns = [
    'No', 'Company', 'Department', 'Position', 'Name', 'ID', 'Date',
    'lisence 1', 'lisence 2', 'lisence 3', 'lisence 4', 'FT', 'VC'
]

# 列名映射（处理类似的列名，例如 'lisence 3.1' -> 'lisence 3'）
column_mapping = {
    'lisence 3.1': 'lisence 3',
}

# 打开log文件以记录处理失败的文件
log_file_path = r'C:\Users\jianbinD\Documents\02 UNDERGROUND OP LICENSE\log.txt'
with open(log_file_path, 'w') as log_file:
    log_file.write("Failed to process the following files:\n")

    # 遍历每个配置
    for config in configurations:
        keyword = config['keyword']
        output_file = config['output_file']
        collected_files_output = config['collected_files_output']

        # 创建空的数据框用于收集结果
        result_df = pd.DataFrame(columns=expected_columns)
        collected_files_list = []

        # 遍历目录及所有子目录
        for root, dirs, files in os.walk(directory):
            for filename in files:
                # 仅处理文件名包含特定关键字的Excel文件
                if keyword in filename and (filename.endswith(".xlsx") or filename.endswith(".xls")):
                    if filename.startswith("~$"):
                        continue  # 忽略临时文件

                    file_path = os.path.join(root, filename)
                    try:
                        # 读取Excel文件
                        df = pd.read_excel(file_path, skiprows=2)

                        # 删除未命名列和不需要的列
                        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                        if 'Photo' in df.columns:
                            df.drop(columns=['Photo'], inplace=True)

                        # 重命名列以匹配预期的列名
                        df.rename(columns=column_mapping, inplace=True)

                        # 处理重复列名，确保列名唯一
                        seen = {}
                        new_columns = []
                        for col in df.columns:
                            if col in seen:
                                seen[col] += 1
                                new_columns.append(f"{col}_{seen[col]}")
                            else:
                                seen[col] = 0
                                new_columns.append(col)
                        df.columns = new_columns

                        # 打印实际的列名以便检查
                        print(f"File: {filename}")
                        print(f"Actual Columns: {list(df.columns)}")

                        # 过滤掉不在预期列表中的列（去掉后缀后进行比较）
                        df = df[[col for col in df.columns if col.split('_')[0] in expected_columns]]

                        # 检查实际的列是否在预期的列名中（允许额外列，但不允许缺少预期列）
                        base_columns = [col.split('_')[0] for col in df.columns]
                        if set(expected_columns).issubset(base_columns):
                            # 过滤掉 Name 为空的行
                            df_filtered = df[df['Name'].notna()]
                            result_df = pd.concat([result_df, df_filtered], ignore_index=True)
                            # 将成功抓取数据的文件路径添加到列表
                            collected_files_list.append(file_path)
                        else:
                            print(f"Header mismatch in file: {filename}")
                            log_file.write(f"Header mismatch in file: {file_path}\n")
                    except Exception as e:
                        print(f"Failed to process {filename}: {e}")
                        log_file.write(f"Failed to process {file_path}: {e}\n")

        # 保存结果到Excel文件
        result_df.to_excel(output_file, index=False)
        # 保存已成功抓取数据的文件路径到Excel文件
        collected_files_df = pd.DataFrame(collected_files_list, columns=['File Path'])
        collected_files_df.to_excel(collected_files_output, index=False)

        print(f"Data extraction and export completed for files with keyword: {keyword}")
