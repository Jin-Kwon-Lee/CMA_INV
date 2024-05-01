import os
import pandas as pd
import fitz
# 출력 옵션 설정
pd.set_option('display.max_colwidth', None)

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

def read_lines_pdf(path):
    doc = fitz.open(path)
    text = ''
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text()
        page = doc.load_page(page_num)
        lines = page.get_text().split('\n')  # 줄 단위로 텍스트 분할
    return lines

def _rename_file(df, ori_file_list, pdf_directory):
    for ori_name in ori_file_list:
        shipment = df[(df['Invoice Ref.'] == ori_name)]['Shipment Ref.'].iloc[0]
        currency = df[(df['Invoice Ref.'] == ori_name)]['Currency'].iloc[0]
        new_name = shipment + ' ' + currency
    
        original_file_path  = pdf_directory + '/' + ori_name + '.pdf'
        new_file_path       = pdf_directory + '/' + new_name + '.pdf'
    
        if os.path.exists(original_file_path):
            os.rename(original_file_path, new_file_path)
        else:
            print(f'{ori_name}을 찾을 수 없습니다.')

def KRW_new_name(df, ori_file_list, pdf_directory):
    KRW_name_list = []
    for ori_name in ori_file_list:
        shipment = df[(df['Invoice Ref.'] == ori_name)]['Shipment Ref.'].iloc[0]
        currency = df[(df['Invoice Ref.'] == ori_name)]['Currency'].iloc[0]
        new_name = shipment + ' ' + currency

        new_file_path       = pdf_directory + '/' + new_name + '.pdf'
        if os.path.exists(new_file_path):
            if currency == 'KRW' : KRW_name_list.append(new_name)
                
    return KRW_name_list
        
def summary_table(df, krw_list, pdf_directory):
    total_df = pd.DataFrame()
    for new_name in krw_list:
        new_file_path = pdf_directory + '/' + new_name + '.pdf'
        lines = read_lines_pdf(new_file_path)
        
        cont_qty_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "Qty")
        bl_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "Payment before delivery of Bill Of Lading (Export) or containers (Import)")
        total_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "Total")
        first_uni_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "UNI")
        total_cnt_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "Total Amount:")

        desc_cnt = len(lines[total_idx + 1 : first_uni_idx])
        
        qty = lines[cont_qty_idx + 4]
        bl = lines[bl_idx + 1]
        desc = lines[total_idx + 1 : first_uni_idx]
        amount = lines[total_cnt_idx - desc_cnt : total_cnt_idx]
        mapping = dict(zip(desc, amount))
        df = pd.DataFrame([mapping], index = [bl])
        df['QTY'] = qty
        
        total_df = pd.concat([total_df, df], axis=0)

    return total_df



# PDF 파일이 저장된 디렉토리 경로
working_path = os.getcwd() 
pdf_directory = working_path + '/document'
excel_file_path = working_path + '/file_mapping.xlsx'

sheet_name = 'raw_data'

df = read_excel_file(excel_file_path,sheet_name)
ori_file_list = df['Invoice Ref.'].tolist()

_rename_file(df, ori_file_list, pdf_directory)

krw_list = KRW_new_name(df, ori_file_list, pdf_directory)

df = summary_table(df, krw_list, pdf_directory)


qty_column = df.pop('QTY')
df['QTY'] = qty_column

df.to_excel('result.xlsx')
