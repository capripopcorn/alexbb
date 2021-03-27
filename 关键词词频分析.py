#!/usr/bin/env python
# -*- coding: utf-8 -*-

from os.path import isfile, join
from os import listdir
import os
import pandas as pd
import sys
from pathlib import Path

assert sys.version_info >= (3, 6), "运行python版本必需高于3.6"

# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    dir_path = os.path.dirname(sys.executable)
elif __file__:
    dir_path = os.path.dirname(os.path.realpath(__file__))

workbench_path = f'{dir_path}/workbench'
Path(workbench_path).mkdir(parents=True, exist_ok=True)
result_path = f'{dir_path}/results'
Path(result_path).mkdir(parents=True, exist_ok=True)

xlsx_paths = [f for f in listdir(workbench_path) if isfile(
    join(workbench_path, f)) and f.endswith('.xlsx')]
print(xlsx_paths)

assert len(xlsx_paths) > 0, f"请将xlsx文件移动到目录{workbench_path}, 再重新执行应用程序。生成的结果将会被保存到:\n{result_path}"

for xlsx_path in xlsx_paths:
    df = pd.read_excel(join(workbench_path, xlsx_path),
                       index_col=0, engine='openpyxl')
    result = df["商品标题"].str.split(expand=True).stack().value_counts()
    keywords = result.index
    score_result = []
    for keyword in keywords:
        keyword_score = 0
        keyword_count = 0
        for row_index, row_val in enumerate(df["商品标题"].values):
            if keyword in row_val:
                # print(f"key:{keyword}\nrow_val: {row_val}\nscore: {df['评分数'].values[row_index]}")
                keyword_count += 1
                try:
                    keyword_score = keyword_score + \
                        int(df['评分数'].values[row_index])
                except:
                    pass
        # print(f"final score for key: {keyword} is => {keyword_score}")
        score_result.append([keyword, keyword_score, keyword_count])
    print(f'score_result ==> {score_result}')
    final_result = pd.DataFrame(score_result, columns=['关键词', '分数', '词频'])
    # 1.去掉包括符号的关键词
    regex_to_replace = r"[-!$%^&*()_+|~=`{}\[\]:\";'<>?,.\/]"
    final_result.关键词 = final_result.关键词.replace({regex_to_replace:''}, regex=True)
    # 2.统一变成小写
    final_result.关键词 = final_result.关键词.str.lower()
    # 3.复数 -> 单数，除特殊单复数, eg: people -x-> person (扭曲过多关键词传达的信息，即people和person在关键词分析中可能不代表一种关键词)
    
    # 4.去重, 叠加重复关键词的分数和词频
    final_result = final_result.groupby('关键词')['分数','词频'].agg('sum').reset_index()
    final_result_sorted = final_result.sort_values(['词频'], ascending=[False]).reset_index(drop=True)
    final_result_path = f'{join(result_path, os.path.splitext(xlsx_path)[0])}_分析结果.xlsx'
    writer = pd.ExcelWriter(final_result_path)
    print(f"关键词 - 词频加权(+评分数)分析结果 - {xlsx_path}:\n{final_result_sorted}")
    final_result_sorted.to_excel(writer, startcol=0, startrow=1)
    worksheet = writer.sheets['Sheet1']
    worksheet.write_string(0, 0, '关键词 - 词频加权(+评分数)分析')
    writer.save()

print(f"完成! 生成的结果已被保存到:\n{result_path}")