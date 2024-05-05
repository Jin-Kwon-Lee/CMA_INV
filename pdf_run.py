import os
import pandas as pd
import fitz

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
    amount_df= pd.DataFrame()
    
    for new_name in krw_list:
        new_file_path = pdf_directory + '/' + new_name + '.pdf'
        lines = read_lines_pdf(new_file_path)
        

        cont_qty_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "Qty")
        bl_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "Payment before delivery of Bill Of Lading (Export) or containers (Import)")
        total_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "Total")
        first_uni_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "UNI")
        total_cnt_idx = next(idx for idx, line in enumerate(lines) if line.rstrip() == "Total Amount:")
        krw_idx = len(lines) - 1 - next(idx for idx, line in enumerate(reversed(lines)) if line.rstrip() == "KRW")

        desc_cnt = len(lines[total_idx + 1 : first_uni_idx])
        
        qty = lines[cont_qty_idx + 4]
        bl = lines[bl_idx + 1].strip()
        desc = lines[total_idx + 1 : first_uni_idx]
        amount = lines[krw_idx + 1 : krw_idx + 1 + desc_cnt]
        amount = [float(x.strip().replace(',','')) for x in amount]
        
        amount_df = df.loc[(df['Shipment Ref.'] == bl),['Amount','Currency']]
        
        krw_amount = amount_df[(amount_df['Currency'] == 'KRW')]['Amount'].iloc[0]
        usd_amount = amount_df[(amount_df['Currency'] == 'USD')]['Amount'].iloc[0]
       
        mapping = dict(zip(desc, amount))
        summary_df = pd.DataFrame([mapping], index = [bl])
        
        
        col_list = list(summary_df.columns)
        col_list.remove('Export Documentation Fee')

        summary_df[col_list] = summary_df[col_list] * int(qty)
 
        summary_df['QTY'] = qty
        summary_df['KRW_amount'] = krw_amount
        summary_df['USD_amount'] = usd_amount

        total_df = pd.concat([total_df, summary_df], axis=0)

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
total_df = summary_table(df, krw_list, pdf_directory)


# DataFrame에서 마지막에 위치시킬 열을 제외한 열 선택하여 새로운 DataFrame 생성
last_columns = ['QTY','KRW_amount','USD_amount']
other_columns = [col for col in total_df.columns if col not in last_columns]
total_df = total_df[other_columns + last_columns]

total_df.to_excel('result.xlsx')
