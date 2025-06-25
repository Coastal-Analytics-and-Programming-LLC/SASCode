import sys
import pandas as pd
import numpy as np
import csv
import matplotlib.pyplot as plt
import xlsxwriter
import seaborn as sns
import io
from PIL import Image
from openpyxl.utils import get_column_letter
#from openpyxl.drawing.image import Image
import itertools

import glob
import os
from pathlib import Path

from myfunctions import unique_string_list, add_blank_lines, get_header, filter_by_string, coalesce_col, \
    create_report_column, add_totals_w_paid, add_totals_wo_paid, format_currency, get_excel_col_letters
	
	
	
def disruption_process(file_path, acct_name, excel_file):

    drops  = ['row_id']
    string_to_remove = ' (Match)'
     
       
    filename = file_path
    acct_name = acct_name
    excel_file = excel_file
    
    
    print(filename)
    print(acct_name)
    
    df2 = pd.read_csv(filename)

    num_rows = len(df2)
    #header_list=get_header(filename)
      
    keeps1 = [
        "Some Dental Program 1",
        "Some Dental Program 2",
		"Some Dental Program 3",
        "Some Dental Program 4",
		"Some Dental Program 5",
        "Some Dental Program 6",
		"Some Dental Program 7",
        "Some Dental Program 8",
		"Some Dental Program 9",
        "Some Dental Program 10",
    ]

    keeps2 = [
        "ID",
        "TIN",
        "NPI",
        "First Name",
        "Last Name",
        "Business Name",
        "Address",
        "City",
        "State",
        "Zip",
        "Paid Amount",
        "Members",
        "Services",
        "Claims",
        "clientid"
    ]

    keeps3 = [
        "Incumbent_1",
        "Incumbent_2",
        "Incumbent_3"
    ]
 
    header_list=get_header(filename)
    
    amt_paid_list = []
    if 'Paid_Amount' in header_list: 
        paid = True
        df2['Paid Amount']  = df2['Paid_Amount'].replace( '[\$,)]','', regex=True ).replace( '[(]','-',   regex=True ).astype(float)
        df2.info()
        amt_paid_list = ['Paid Amount']

    else:
        paid = False
        amt_paid_list = []
      
        
    comp_list=header_list[15:18]

    keeps_match      = filter_by_string(header_list, '(Match)')
    keeps_confidence = filter_by_string(header_list, '(Match Confidence)')

    confidence_list = [x for x in keeps_confidence if x.split(sep=" (")[0] in keeps1]

    drop_list1 = filter_by_string(header_list, 'Criteria')
    drop_list2 = filter_by_string(header_list, '(LocationID)')
    drop_list3 = filter_by_string(header_list, '(Matching Pass ID)')
    drop_list4 = filter_by_string(header_list, '(ProviderID)')
    drop_list5 = [x for x in keeps_confidence if x.split(sep=" (")[0] not in keeps1]


    calcs_network  = [x for x in keeps_match if x.split(sep=" (")[0] in keeps1]
    calcs_compare1 = [x for x in comp_list if x.split(sep=" (")[0] not in keeps3]
    calcs_compare2 = [x for x in keeps_match if x.split(sep=" (")[0] not in keeps1]

    calcs_compare = calcs_compare1 + calcs_compare2

    drop_list = drop_list1 + drop_list2 + drop_list3 + drop_list4 + drop_list5 + drops

    df3 = df2.drop(columns=drop_list)   
    
    #Replacing the target string with a blank string
    column_name_mapping = {old_name: old_name.split(' (Match)')[0] for old_name in df3.columns}
    df3 = df3.rename(columns=column_name_mapping)
    
    
    df_confidence = df3.copy()

    df_confidence = df3[confidence_list]

    for col in confidence_list:
        
        df_confidence[col] = df_confidence[col].apply(lambda x: 'D: Weak' if x == 'Low' else 'C: Moderate' if x == "Medium" else 'B: Strong' if x == 'High' else 'A: No Matches')
    
    confidence_freq = pd.DataFrame()

    for col in range(len(confidence_list)):
        summary = df_confidence.groupby([confidence_list[col]]).size().reset_index(name='COUNT')
        summary['REPORT'] =f"{confidence_list[col]}"
        
        confidence_freq = pd.concat([confidence_freq,summary],ignore_index=True)

    confidence_freq1 = coalesce_col(confidence_freq, confidence_list, 'CONFIDENCE')

    confidence_freq1 = confidence_freq1[['REPORT','CONFIDENCE','COUNT']]
    confidence_freq1['REPORT'] = confidence_freq1['REPORT'].str.rsplit(' (', n=1).str[0]

    confidence_freq1 = confidence_freq1.sort_values(by=['REPORT', 'CONFIDENCE'])


    df_detail = df3.copy()

    for col in confidence_list:
        df_detail[col]   = df3[col].apply(lambda x: 'Weak' if x == 'Low' else 'Moderate' if x == "Medium" else 'Strong' if x == 'High' else 'No Matches') 
        
    new_calcs_network = [item.replace(string_to_remove,'') for item in calcs_network]
    new_calcs_compare = [item.replace(string_to_remove,'') for item in calcs_compare]
           
    freq_list = new_calcs_network + new_calcs_compare
    df_freqs = df3.copy()
    df_freqs = df3.loc[:, df3.columns.isin(freq_list)]
        
    groupdata = pd.DataFrame()
    for col in range(len(new_calcs_network)):
        for x in range(len(new_calcs_compare)):
            summary = df_freqs.groupby([new_calcs_network[col],  new_calcs_compare[x]]).size().reset_index(name='COUNT')
            summary['REPORT'] =f"{new_calcs_network[col]} v {new_calcs_compare[x]}"
            groupdata = pd.concat([groupdata,summary],ignore_index=True)

    groupdata1 = coalesce_col(groupdata, new_calcs_network, 'UCD')
    groupdata2 = coalesce_col(groupdata1, new_calcs_compare, 'COMPETITOR')
    groupdata2 = groupdata2[['REPORT','UCD','COMPETITOR','COUNT']]
 
    conditions = [
    (groupdata2['UCD'] == 'Yes') & (groupdata2['COMPETITOR'] == 'Yes'),
    (groupdata2['UCD'] == 'No')  & (groupdata2['COMPETITOR'] == 'No') ,
    (groupdata2['UCD'] == 'Yes') & (groupdata2['COMPETITOR'] == 'No') ,
    (groupdata2['UCD'] == 'No')  & (groupdata2['COMPETITOR'] == 'Yes')
    ]
   
    values = ['No Disrupt: UCD=Y Comp=Y', 'No Disrupt: UCD=N Comp=N', 'Benefit: UCD=Y Comp=N', 'Disrupt: UCD=N Comp=Y']

    groupdata2['MATCH COMPARE'] = np.select(conditions, values, default = 'Other')

    groupdata3 = groupdata2;

    final_cols = ['REPORT', 'MATCH COMPARE', 'COUNT']
    groupdata_final = groupdata3[final_cols]
    groupdata_final = groupdata_final.sort_values(by=['REPORT', 'MATCH COMPARE'])
    
    df_paid = pd.DataFrame()
    paid_list = []

    if paid:
        df_paid = df3.copy()
        paid_list = freq_list + amt_paid_list
    
        df_paid = df_paid.loc[:, df3.columns.isin(paid_list)]
    
    paid_agg = pd.DataFrame()

    if paid:
        for r in range(len(new_calcs_network)):
            for p in range(len(new_calcs_compare)):
                agg = df_paid.groupby([new_calcs_network[r],  new_calcs_compare[p]]).agg({'Paid Amount':'sum'}).reset_index()
                agg['REPORT'] = f"{new_calcs_network[r]} v {new_calcs_compare[p]}"
        
                paid_agg = pd.concat([paid_agg, agg], ignore_index=True)

    paid_agg1 = pd.DataFrame()
    paid_agg2 = pd.DataFrame()

    if paid:
        paid_agg1 = coalesce_col(paid_agg, new_calcs_network, 'UCD')
        paid_agg2 = coalesce_col(paid_agg1, new_calcs_compare, 'COMPETITOR')
        paid_agg2 = paid_agg2[['REPORT','UCD','COMPETITOR','Paid Amount']]
      
        
    if paid == True:
        paid_agg2['MATCH COMPARE'] = np.select(conditions, values, default = 'Other')
        final_cols_paid = ['REPORT', 'MATCH COMPARE', 'Paid Amount']
        paid_agg2 = paid_agg2[final_cols_paid]
        paid_agg2 = paid_agg2.sort_values(by=['REPORT', 'MATCH COMPARE'])
       
        
    merged_df = pd.DataFrame()

    if paid == True:
        merged_df = pd.merge(paid_agg2, groupdata_final, on=['REPORT', 'MATCH COMPARE'], how='inner')
        merged_agg = merged_df.groupby('REPORT').agg({'Paid Amount': 'sum', 'COUNT': 'sum'}).reset_index()
        merged_agg = merged_agg.rename(columns={'Paid Amount': 'Total Paid', 'COUNT': 'Total Count'})
        merged_df = pd.merge(merged_df, merged_agg, on=['REPORT'], how='inner')
        merged_df['Total % Paid'] = merged_df['Paid Amount'] / merged_df['Total Paid'].round(1)
        merged_df['Total % Dentists'] = merged_df['COUNT'] / merged_df['Total Count'].round(1)
    else:
        merged_df = groupdata_final.copy()
        merged_agg = merged_df.groupby('REPORT').agg({'COUNT': 'sum'}).reset_index()
        merged_agg = merged_agg.rename(columns={'COUNT': 'Total Count'})
        merged_df = pd.merge(merged_df, merged_agg, on=['REPORT'], how='inner')
        merged_df['Total % Dentists'] = merged_df['COUNT'] / merged_df['Total Count'].round(1) 
        
    if paid == True:
        columns_to_round = ['rest_paid', 'rest_dentists', 'Total % Paid', 'Total % Dentists']
    else:
        columns_to_round = ['rest_dentists', 'Total % Dentists']

    metric1 = pd.DataFrame()
    metric2 = pd.DataFrame()
    metric3 = pd.DataFrame()
    metric4 = pd.DataFrame()
    metric5 = pd.DataFrame()
    metric6 = pd.DataFrame()

    merged_df[['ucd', 'comp']] = merged_df['REPORT'].str.split(' v ', expand=True)    
    merged_df['Description'] = merged_df.apply(create_report_column, axis=1)
    stats1 = merged_df.copy()
    stats1 = merged_df[merged_df['MATCH COMPARE'] == 'Disrupt: UCD=N Comp=Y'][['REPORT', 'ucd', 'Total % Paid', 'Total % Dentists']]
    stats1['Total % Dentists'] = 100 * stats1['Total % Dentists']
    stats1['rest_dentists'] = (100 - stats1['Total % Dentists'])

    if paid == True:
        stats1['Total % Paid'] = 100 * stats1['Total % Paid']
        stats1['rest_paid'] = (100 - stats1['Total % Paid'])

    stats1[columns_to_round] = stats1[columns_to_round].round(1)
    
    print(stats1)
 
    metric1 = stats1['REPORT']
    metric2 = stats1['ucd']
    metric4 = stats1['Total % Dentists']
    metric6 = stats1['rest_dentists']

    if paid == True:
        metric3 = stats1['Total % Paid']
        metric5 = stats1['rest_paid']
        m3 = metric3.to_list()
        m5 = metric5.to_list()

    m1 = metric1.to_list()
    m2 = metric2.to_list()
    m4 = metric4.to_list()
    m6 = metric6.to_list()  

    if paid == True:
        merged_df = merged_df[['REPORT', 'Description', 'Paid Amount' , 'Total % Paid', 'COUNT', 'Total % Dentists']]
    else:
        merged_df = merged_df[['REPORT', 'Description', 'COUNT', 'Total % Dentists']]
    
    def format_desc(desc):
        parts = desc.split('*')
        return '\n'.join(parts)

    merged_df['Description'] = merged_df['Description'].apply(format_desc)
    
    if paid == True:
        df_with_totals = add_totals_w_paid(merged_df)
    else:
        df_with_totals = add_totals_wo_paid(merged_df)  
        
        
    df_with_totals = df_with_totals.rename(columns={'COUNT': 'Number of Dentists'})
    df_with_totals['Number of Dentists'] = df_with_totals['Number of Dentists'].apply(lambda x: f"{x:,}")
    df_with_totals['Total % Dentists'] = df_with_totals['Total % Dentists'] * 100
    df_with_totals['Total % Dentists'] = df_with_totals['Total % Dentists'].apply(lambda x: f"{x:.2f}%")

    if paid == True:
        df_with_totals['Paid Amount'] = df_with_totals['Paid Amount'].apply(format_currency)
        df_with_totals['Total % Paid'] = df_with_totals['Total % Paid'] * 100
        df_with_totals['Total % Paid'] = df_with_totals['Total % Paid'].apply(lambda x: f"{x:.2f}%")


    df_with_blanks = add_blank_lines(df_with_totals, 'Description', 'TOTAL', 6)

    if paid == True:
        df_with_blanks = df_with_blanks[['Description', 'Paid Amount', 'Total % Paid', 'Number of Dentists', 'Total % Dentists']]
    else:
        df_with_blanks = df_with_blanks[['Description', 'Number of Dentists', 'Total % Dentists']]
    
     
    
    df_disrupt = groupdata3.copy()

    conditions1 = [
        ((df_disrupt['UCD'] == 'Yes') & (df_disrupt['COMPETITOR'] == 'Yes')) |
        ((df_disrupt['UCD'] == 'No')  & (df_disrupt['COMPETITOR'] == 'No'))  |
        ((df_disrupt['UCD'] == 'Yes') & (df_disrupt['COMPETITOR'] == 'No')) ,
         (df_disrupt['UCD'] == 'No')  & (df_disrupt['COMPETITOR'] == 'Yes')
    ]
 

    values1 = ['No Disruption', 'Disruption']

    df_disrupt['DISRUPTION'] = np.select(conditions1, values1, default = 'Other')
    disrupt_agg = df_disrupt.groupby(['REPORT', 'DISRUPTION'])['COUNT'].sum().reset_index(name='SUM COUNT')
    disrupt_agg['PERCENT'] = disrupt_agg['SUM COUNT'] / num_rows * 100;
    disrupt_agg['PERCENT'] = disrupt_agg['PERCENT'].round(0).astype(int)
    disrupt_y = disrupt_agg[disrupt_agg['DISRUPTION'] == 'Disruption']
    disrupt_n = disrupt_agg[disrupt_agg['DISRUPTION'] == 'No Disruption']
    disrupt_y = disrupt_y['PERCENT']
    disrupt_n = disrupt_n['PERCENT']

    perc_y = disrupt_y.to_list()
    perc_n = disrupt_n.to_list()

    disrupt_agg = disrupt_agg[['REPORT', 'DISRUPTION', 'PERCENT']]

    unique_list = unique_string_list(disrupt_agg,'REPORT')

    
    df_detail = df_detail.drop(columns ='Paid Amount', axis=1)
    num_cols = len(df_detail.columns)
    num_cols = num_cols - 1
    
    i = 1
    k = 10
    b = 9
        
    
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
   
    
    formatted_num = f'{num_rows:,}'

    cell_width = 5
    cell_height = 8

    x_scale = cell_width
    y_scale = cell_height

    workbook = writer.book
    worksheet= workbook.add_worksheet('Summary')

    red_format  = workbook.add_format({'bold': True, 'font_color': 'red', 'font_size': 12})
    blue_format = workbook.add_format({'bold': True, 'font_color': '#0000FF', 'font_size': 14})

    worksheet.insert_image('A1',"/c01/home/lid4otx/projects/ddp_assets/logos/UCD.png", {'x_scale': x_scale,'y_scale': y_scale})

    worksheet.hide_gridlines([1])

    for _ in range(len(unique_list)):
        worksheet.write(f'A{k}', f"Report: {unique_list[_]}.", blue_format)
        worksheet.write(f'B{k+1}', f"Based on {formatted_num} dentists, there was a something of {perc_y[_]}% and {perc_n[_]}% Non-Something.", red_format)
        k += 3
    
    r_counter = 10

    explode = (0, .1)
    
    dis_group_1 = disrupt_agg.groupby('REPORT')

    for name, group in dis_group_1:
        # Aggregate data for pie chart
        pie_data = group.groupby('DISRUPTION')['PERCENT'].sum()

        # Create pie chart
        plt.figure(figsize=(6, 4))
        plt.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%', startangle=90, explode=explode, shadow=True, textprops={'color': 'white', 'weight': 'bold'}, wedgeprops={'edgecolor': 'black', 'linewidth': 2, 'linestyle': 'solid'})
    
        plt.title(f"Disruption: {name}")
        plt.axis('equal')
        plt.tight_layout()
    
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png')
        img_buffer.seek(0)
    
        img = Image.open(img_buffer)
        worksheet.insert_image(f'O{r_counter}', f"img_{r_counter}", {'image_data': img_buffer})
        # Save pie chart as image
    
        r_counter += 19
    
        plt.close() # Close the figure
    
    
    ordered_categories = ['No Matches', 'Strong', 'Moderate', 'Weak']

    workbook = writer.book
    worksheet= workbook.add_worksheet('Confidence Summary')
    
    worksheet.hide_gridlines([1])    
        
    def create_conf_bar_charts_to_excel(df):
   
        grouped = df.groupby(df.columns[0])
        sheet_name = 'Confidence Summary'
        writer.sheets[sheet_name] = worksheet
    
        row_counter = 1

        for group_name, group_data in grouped:
            fig, ax = plt.subplots(figsize=(5, 5))
            sns.barplot(x=df.columns[1], y=df.columns[2], data=group_data, ax=ax)
            ax.set_title(group_name)
            ax.set_xlabel(df.columns[1])
            ax.set_ylabel(df.columns[2])
            ax.bar_label(ax.containers[0])
               
            plt.tight_layout()

            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png')
        
            img_buffer.seek(0)

            img = Image.open(img_buffer)
            worksheet.insert_image(f'A{row_counter}', f"img_{row_counter}", {'image_data': img_buffer})
        
            row_counter += 24 # Adjust row spacing for each plot

        
            plt.close(fig) # Close the figure    
            
    create_conf_bar_charts_to_excel(confidence_freq1)            
    
    
    
    workbook = writer.book
    worksheet= workbook.add_worksheet('Disruption Summary')
    
    worksheet.hide_gridlines([1])    
           
    def create_disrupt_bar_charts_to_excel(df):
   
        grouped = df.groupby(df.columns[0])

        sheet_name = 'Disruption Summary New'
        writer.sheets[sheet_name] = worksheet
        row_counter = 1
        
        for group_name, group_data in grouped:
            fig, ax = plt.subplots(figsize=(10, 5))
            sns.barplot(x=df.columns[1], y=df.columns[2], data=group_data, ax=ax , hue=df.columns[1] , palette=['green', 'red', 'blue', 'blue'])
            ax.set_title(group_name)
            ax.set_xlabel(df.columns[1])
            ax.set_ylabel(df.columns[2])
            ax.bar_label(ax.containers[0])
        
                        
            plt.tight_layout()

            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png')
        
            img_buffer.seek(0)

            img = Image.open(img_buffer)
            worksheet.insert_image(f'A{row_counter}', f"img_{row_counter}", {'image_data': img_buffer})
        
            row_counter += 24 # Adjust row spacing for each plot

            plt.close(fig) # Close the figure
        
               
    create_disrupt_bar_charts_to_excel(groupdata_final)   
    
    
    
    rows = num_rows + 11

    cell_width = 5
    cell_height = 8

    x_scale = cell_width
    y_scale = cell_height

    df_detail.to_excel(writer, sheet_name='Detail', index=False, startrow=12, header=False)
    workbook = writer.book
    worksheet= writer.sheets['Detail']
    worksheet.insert_image('A1',"/c01/home/logos/SomeLogo.png", {'x_scale': x_scale,'y_scale': y_scale})

    header_row = workbook.add_format({
        'bold': True,
        'bg_color': 'black',  
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12
    })    


    worksheet.write(f'A8', "Disruption Analysis", blue_format)
    worksheet.write(f'A9', f"{acct_name}", blue_format)
    
    for col_num, value in enumerate(df_detail.columns.values):
        worksheet.write(11, col_num, value, header_row)
    
    worksheet.hide_gridlines([1])
    border_format = workbook.add_format({'border': 1})

    worksheet.conditional_format(12, 1, rows, num_cols, {'type': 'blanks', 'format': border_format})
    worksheet.conditional_format(12, 1, rows, num_cols, {'type': 'no_blanks', 'format': border_format})
    worksheet.autofit()
    df_detail.info()  

    
    
    rows = num_rows + 11

    cell_width = 5
    cell_height = 8

    x_scale = cell_width
    y_scale = cell_height

    center_format = workbook.add_format({'align': 'center', 'valign': 'vbottom'})

    df_with_blanks.to_excel(writer, sheet_name='OVERALL SUMMARY', index=False, startrow=13, header=False)
    workbook = writer.book
    worksheet= writer.sheets['OVERALL SUMMARY']
    worksheet.insert_image('A1',"/c01/home/logos/SomeLogo.png", {'x_scale': x_scale,'y_scale': y_scale})

    for col_num, value in enumerate(df_with_blanks.columns.values):
        worksheet.write(11, col_num, value, header_row)

    worksheet.hide_gridlines([1])
    
    border_format = workbook.add_format({'border': 1})

    text_wrap_format = workbook.add_format({'text_wrap': True})

    worksheet.conditional_format('A12:E80', {'type': 'no_blanks', 'format': border_format})


    worksheet.set_column(0, 0, 60, text_wrap_format)
    worksheet.set_row(0, None, text_wrap_format)
    worksheet.set_column(1, 1, 20, cell_format = center_format)
    worksheet.set_column(2, 2, 20, cell_format = center_format)
    worksheet.set_column(3, 3, 20, cell_format = center_format)
    worksheet.set_column(4, 4, 20, cell_format = center_format)

    worksheet.freeze_panes(12,0)

    for _ in range(len(m1)):
        worksheet.write(f'G{b}', "Disruption Analysis", blue_format)
        worksheet.write(f'G{b+1}', f"{m1[_]}" , blue_format)
    
        if paid == True:
            worksheet.write(f'G{b+2}', f"Based on the data to the left, a move to the {m2[_]} network will do something {m3[_]}% of claims paid to" , red_format)
            worksheet.write(f'G{b+3}', f"{m4[_]}% of the dentists. A move will not do anything {m5[_]}% of claims paid to {m6[_]}% of the dentists.", red_format)
        else:
            worksheet.write(f'G{b+2}', f"Based on the data to the left, a move to the {m2[_]} network will do something else " , red_format)
            worksheet.write(f'G{b+3}', f"{m4[_]}% of the dentists. A move will not do anything at all to {m6[_]}% of the dentists.", red_format)
        
        b += 11

    workbook.close()
    writer.close()
    
    
def main():
    
    def find_padl_files(directory):
    
        pattern = os.path.join(directory, "*.csv")  
        file_list = glob.glob(pattern)  
        return file_list
  
    new_dir = "/n01/data/ucd/some_dir/in"
    out_dir = "/n01/data/ucd/some_dir/out/"
    
    files = find_padl_files(new_dir)
 
    print(files)
    
    file_list = []
    account_list = []
    if files:
        print("Files found:")
        for file in files:
            file_list.append(file)
            split_item = file.split('-')[1]
            account_list.append(split_item)
        
    else:
        print("No files matching the pattern found.")
 
   
    if files:
        for file_path in files: 
            acct_name = file_path.split('-')[1]
            excel_file = file_path.split('/')[6]
            excel_file = excel_file.split('.')[0]
            excel_file = out_dir + excel_file.replace("In", "Out.xlsx")
            print(excel_file)
            disruption_process(file_path, acct_name, excel_file)
            
    
    
#if __name__ == 'main':
main()
    