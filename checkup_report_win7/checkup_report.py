#! /usr/bin/env python3.6
'''
    Author      : Coslate
    Date        : 2018/07/07
    Description :
        This program will examine the input excel whether have the job number, and
        concatenate it to the total check excel output. It will also highlight the
        one that has repeated job number in multiple input excels.
'''
import argparse
import numpy as np
import pandas as pd
from   pandas import ExcelWriter
import ntpath
import shutil
import os
import sys
import datetime

#########################
#     Main-Routine      #
#########################
def main():
    (file_name_path, total_num, org_checked_file) = ArgumentParser()

    #Get the input excel file data
    file_name = ntpath.basename(file_name_path)
    file_data = pd.read_excel(file_name_path)

    original_check, total_num_origin = ProcessOrigCheckFile(org_checked_file, total_num)

    #Build the new dataframe
    total_check = original_check.copy()
    check_num_list = file_data["工號"].values.tolist()
    total_check[file_name] = [0 for x in range(0, total_num_origin)]
    columns = total_check.columns.tolist()

    for check_num in range(0, total_num):
        total_check['工號'][check_num] = check_num
        if(check_num > (total_num_origin-1)):
            break
        elif(check_num in check_num_list):
            total_check['做過健檢'][check_num] = 1
            total_check[file_name][check_num] = 1

    total_check = total_check.drop(total_check.index[total_num:total_num_origin:1])
    if(total_num > total_num_origin):
        for check_num in range(total_num_origin, total_num):
            total_check.loc[check_num] = [0 for x in range(len(columns))]
            total_check['工號'][check_num] = check_num


    #Write out to the new excel(total checked excel)
    writer = ExcelWriter(org_checked_file)
    total_check.to_excel(writer,'Checkup',index=False)
    workbook = writer.book
    worksheet = writer.sheets['Checkup']
    HightLightRepeat(total_check, workbook, worksheet, columns)
    writer.save()

#########################
#     Sub-Routine       #
#########################
def ArgumentParser():
    total_num = 0
    file_name_path = ""
    org_checked_file = ""

    parser = argparse.ArgumentParser()
    parser.add_argument("--file_name_path"   , "-file_to_check"        , help="The name of the input excel to do the examining.")
    parser.add_argument("--org_checked_file" , "-org_checked_file"     , help="The name of the original checked file.")
    parser.add_argument("--total_num"        , "-total_job_number"     , help="The maximum number of the job number.")

    args = parser.parse_args()
    if args.file_name_path:
        file_name_path = args.file_name_path

    if args.org_checked_file:
        org_checked_file = args.org_checked_file

    if args.total_num:
        total_num = int(args.total_num)

    return (file_name_path, total_num, org_checked_file)

def HightLightRepeat(in_data_frame, workbook, worksheet, columns):
    positiveFormat = workbook.add_format({
        'bold': 'true',
        'font_color': 'red'
    })

    columns_num = len(columns)
    for index, row in in_data_frame.iterrows():
        repeat_num = 0
        for x in range(2, columns_num):
            if(row[columns[x]] == 1):
                repeat_num += 1

        if(repeat_num > 1):
            worksheet.set_row(index+1, 15, positiveFormat)

def ProcessOrigCheckFile(org_checked_file, total_num):
    if os.path.exists(org_checked_file):
        #Get the original total checked excel
        original_check = pd.read_excel(org_checked_file)
        original_check_num_list = original_check["做過健檢"].values.tolist()
        total_num_origin = len(original_check_num_list)

        #Backup the org_checked_file
        org_checked_file_name = ntpath.basename(org_checked_file)
        org_checked_file_path = os.path.dirname(org_checked_file)
        now = datetime.datetime.now()
        time_now = str(now.year)+"_"+str(now.month)+'_'+str(now.day)+'_'+str(now.hour)+'_'+str(now.minute)+'_'+str(now.second)
        backup_checked_file_name = org_checked_file_name+".bk"+"."+time_now+".xlsx"
        backup_file = org_checked_file_path+"/"+backup_checked_file_name
        shutil.copy(org_checked_file, backup_file)
    else:
        total_num_origin = total_num
        index = [x for x in range(0, total_num)]
        columns = ['工號', '做過健檢']

        original_check  = pd.DataFrame(index=index, columns=columns)
        original_check  = original_check.fillna(0)

    return original_check, total_num_origin

#-----------------Execution------------------#
if __name__ == '__main__':
    main()
