import pandas as pd
import streamlit as st
import openpyxl

st.set_page_config(layout="wide")

st.title('RPA reports')
uploaded_file = st.file_uploader('Upload an Excel file', type=['xlsx','xls'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    # st.write(df)
else:
    st.write('Please upload an Excel file.')


if uploaded_file is not None:

    df_data = df[['分公司','一级网点名称','网点代码','交费方式']]
    df_data = df_data.rename(columns={'分公司': 'branch', '一级网点名称': 'agentcomname', '网点代码': 'agentcom','交费方式': 'frequency'})
    # 筛选期交数据
    df_data = df_data[df_data['frequency'] == '期交']
    # 替换"分公司"字符串为空
    df_data['branch'] = df_data['branch'].str.replace('分公司', '', regex=False) 
    # 数据去重
    df_data = df_data.drop_duplicates() 
    # 透视表计算
    df_pivot = df_data.pivot_table(index='branch', columns='agentcomname', values='agentcom', aggfunc='count', fill_value=0)  
    # 重置索引
    df_pivot = df_pivot.reset_index().rename_axis(None, axis=0).rename(columns={'branch': '机构'}) 


    # 获取所有的列名list
    columns_list = [col for col in df_pivot.columns if col != '机构']
    # 计算"合计"
    df_pivot['合计'] =  df_pivot.loc[:, columns_list].sum(axis=1) 
    # 计算双邮
    df_pivot['双邮'] =  df_pivot['中国邮政储蓄银行'] + df_pivot['邮政局'] 
    # 重命名银行简称
    df_pivot = df_pivot.rename(columns={'中国农业银行': '农行', '中国工商银行': '工行', '中国建设银行': '建行','中国银行': '中行','中信银行': '中信','交通银行': '交行', '光大银行': '光大', '上海浦发银行': '浦发','兴业银行': '兴业','招商银行': '招行','广东发展银行': '广发','民生银行': '民生'})
    # 获取重命名后的列名list
    renname_columns_list = df_pivot.columns.tolist()
    # 展示的银行列名
    bank_columns_list = ['双邮','农行', '工行', '建行', '中行', '中信', '交行', '光大', '浦发', '兴业', '招行', '广发', '民生']
    # 获取其他银行列名list
    other_columns_list = [x for x in renname_columns_list if x not in ['合计','机构','中国邮政储蓄银行','邮政局'] and x not in bank_columns_list] 
    # 计算"其他"
    df_pivot['其他'] =  df_pivot.loc[:, other_columns_list].sum(axis=1) 
    # 筛选有数据的字段
    existing_columns = [col for col in bank_columns_list if col in df_pivot.columns] 
    # 展示字段list
    existing_columns = ['机构'] + existing_columns + ['其他','合计']
    # 展示有数据的字段
    df_pivot = df_pivot[existing_columns]


    # 绘制最终展示表格
    # 定义行索引和列名  
    rows = ['广东', '山东', '河南', '安徽', '湖南', '陕西', '四川', '江苏', '河北', '内蒙古', '江西', '浙江', '云南', '青岛', '上海', '宁波', '东莞', '深圳', '天津', '北京', '海南', '黑龙江', '苏州', '无锡']  
    cols = bank_columns_list +  ['其他','合计'] 

    # 创建一个空的DataFrame，索引为行名，列为列名  
    df_show = pd.DataFrame(index=rows, columns=cols)
    # 重命名索引名为rows
    df_pivot = df_pivot.set_index('机构').rename_axis('rows') 
 
    # 匹配值到展示表中
    for col in df_pivot.columns:  
        df_show.loc[df_pivot.index, col] = df_pivot[col] 

    # 更新空值为0
    df_show.fillna(0, inplace=True)
    # 重置索引
    df_show = df_show.reset_index().rename(columns={'index': '机构'})
    # 按照"合计"降序排列
    df_show = df_show.sort_values(by='合计', ascending=False)
    

    # 计算所有列的合计
    totals = df_show[cols].sum(axis=0)
    df_aug = pd.DataFrame([['汇总']+totals.values.tolist()],columns= df_show.columns)
    df_show_total = pd.concat([df_show,df_aug],axis=0)


    # 调整除了"其他""合计"列放在最后，剩下的列按照合计从大到小 从左往后排列

    # 筛选没有合计行的数据
    df_without_total = df_show_total[df_show_total['机构'] != '合计']
    # 筛选合计行的数据
    df_total = df_show_total[df_show_total['机构'] == '合计']  

    # 对列按照合计行从大到小 从左到右 排序
    sorted_cols = df_without_total.loc[:, bank_columns_list].sum(axis=0).sort_values(ascending=False).index  
    # 添加 机构 其他 合计 三列，保持机构在第一列，其他和合计在最后
    sorted_cols_all = ['机构'] + sorted_cols.tolist() +  ['其他','合计'] 
    # 对数据列重新排序
    df_sorted = df_without_total[sorted_cols_all]  
    # 合并"合计"行
    df_final = pd.concat([df_sorted, df_total]).reset_index(drop=True)
    
 
    # 展示结果数据
    st.dataframe(data=df_final.reset_index(drop=True),use_container_width=True)


    # 文字生成

    all = df_show['合计'].sum(axis=0)
    first_branch = df_show.iloc[0,0]
    second_branch = df_show.iloc[1,0]
    third_branch = df_show.iloc[2,0]
    first_branch_value = df_show.iloc[0,-1]
    second_branch_value = df_show.iloc[1,-1]
    third_branch_value = df_show.iloc[2,-1]

    string = " 【⻰庭战报】截⽌⽬前当⽇期交开单⽹点数为 " + str(all) + " 个，前三名分别为：" + str(first_branch) + str(first_branch_value) + "个，" + str(second_branch) + str(second_branch_value) + "个，" + str(third_branch) + str(third_branch_value) + "个，望各机构紧盯渠道每⽇期交开单⽹点数，确保期交保费达成。"


