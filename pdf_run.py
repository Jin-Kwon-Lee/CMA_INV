import os
import pandas as pd


def read_excel_file(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        return df
    except FileNotFoundError:
        print("File not found.")
        return None
    except Exception as e:
        print("An error occurred.:", e)
        return None

# PDF 파일이 저장된 디렉토리 경로
working_path = os.getcwd() 
pdf_directory = working_path + '/document'
excel_file_path = working_path + '/file_mapping.xlsx'

sheet_name = 'raw_data'

# 엑셀 파일에서 파일명 매칭 정보 읽어오기
# df_mapping = pd.read_excel_file(excel_file_path,sheet_name)  # 엑셀 파일 읽어오기

df = read_excel_file(excel_file_path,sheet_name)

ori_file_list = df['Invoice Ref.'].tolist()
# print(ori_file_list)
for ori_name in ori_file_list:
    # print(ori_name)
    shipment = df[(df['Invoice Ref.'] == ori_name)]['Shipment Ref.'].iloc[0]
    currency = df[(df['Invoice Ref.'] == ori_name)]['Currency'].iloc[0]
    new_name = shipment + ' ' + currency

    original_file_path  = pdf_directory + '/' + ori_name + '.pdf'
    new_file_path       = pdf_directory + '/' + new_name + '.pdf'

    if os.path.exists(original_file_path):
        os.rename(original_file_path, new_file_path)
        print(f'{ori_name}의 파일명이 {new_name}으로 변경되었습니다.')
    else:
        print(f'{ori_name}을 찾을 수 없습니다.')







# 원래 파일명과 새로운 파일명 매칭하여 변경
# for index, row in df_mapping.iterrows():
    # original_filename = row['ORIGINAL_FILENAME'] + '.pdf'  # 원래 파일명
    # new_filename = row['NEW_FILENAME'] + '.pdf'  # 새로운 파일명

    # 파일 경로
    # original_file_path = os.path.join(pdf_directory, original_filename)
    # new_file_path = os.path.join(pdf_directory, new_filename)

    # # 파일명 변경
    # if os.path.exists(original_file_path):
    #     os.rename(original_file_path, new_file_path)
    #     print(f'{original_filename}의 파일명이 {new_filename}으로 변경되었습니다.')
    # else:
    #     print(f'{original_filename}을 찾을 수 없습니다.')
