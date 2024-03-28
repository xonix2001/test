import re

import Office操作脚本 as of
import sort_docx_v2 as sd
from extract_name import (
    extract_names_with_new_rule as en, 
    extract_first_chinese_phrase as ep, 
    extract_all_ages_in_order as ea, 
    extract_reason_or_complaint as ec, 
    extractor, 
    extract_all_complaints, 
    extract_exam_data, 
    extract_exam_data_refined
)

def remove_spaces_and_newlines(text):
    """Remove spaces and newlines from a string using regex."""
    return re.sub(r'\s+', '', text) if text and isinstance(text, str) else 'None'

def check(n, col, list_total):
    """Check for data availability and return the value and its location."""
    for l, i in enumerate(list_total[8 - col:]):
        if i[n] and isinstance(i[n], str):
            return i[n], col - l
    return None, ''

def from_basic():
    for n, item in enumerate(columns[0]):  # Assuming the last item in columns corresponds to list1 in original script
        item_cleaned = remove_spaces_and_newlines(item)
        
        # Extract and process data
        gender = extractor(r'男|女', item_cleaned)
        name = en(item_cleaned)
        age = ea(item_cleaned)
        chief=ec(item_cleaned)
        case_num=extractor(r"(?:病历号|病例号|病案号|住院号|病案)[：:]?(\d+[\*\d]*)",item_cleaned)
        image_num=extractor(r"(?:影像号|影响号)[：:]?(\d+)",item_cleaned)
        video_num=extractor(r"(?:脑电图号|脑电号)[：:]?(\d+[-\d]*)",item_cleaned)
        check_num=extractor(r"检查号[：:]?(\d+)",item_cleaned)
        phone_num=extractor(r"(?:电话号|电话)[：:]\D*(\d+)",item_cleaned)
        
        data_list=[name,gender,age,case_num,image_num,video_num,check_num,phone_num]
        
        # Append extracted data to corresponding lists in the dictionary
        for n,i in enumerate(data_list):
            lists[f"list{2*n+1}"].append(i)
            lists[f"list{2*(n + 1)}"].append(len(i))
        
        #chief
        if chief:
            lists['list17'].append(chief)
            lists['list18'].append(len(chief))
        else:
            cc=remove_spaces_and_newlines(columns[1][n])
            cc=extract_all_complaints(cc)
            lists['list17'].append(cc)
            lists['list18'].append(len(cc))
    
def from_filename():
    #姓名    
    lists['name_from_file']=[]
    for i in columns[8]:
        i=remove_spaces_and_newlines(i)
        name=ep(i)
        lists['name_from_file'].append(name)

def from_exam():
    #辅助检查
    for n,i in enumerate(columns[7]):
        checked=check(n,8) 
        i=checked[0] if checked else None
        location=checked[1] if checked else ''
        i=remove_spaces_and_newlines(i)
        e=i.split('$')
        a=[]
        num=0
        for b in e:
            if b and b!='辅助检查：':
                num+=1
            exam_data=extract_exam_data_refined(b)
            if exam_data:
                for i in exam_data:
                    a.append(i)
        lists['list19'].append(a)
        lists['list20'].append(num-len(a))
        lists['list21'].append(location)

def from_pexam():
    #查体
    for n,i in enumerate(columns[6]):
        checked=check(n,7) 
        i=checked[0] if checked else None
        location=checked[1] if checked else ''
        i=remove_spaces_and_newlines(i)
        e=i.split('$')
        a=[]
        num=0
        for b in e:
            if b and all(key not in b for key in ['MRI','EEG','PET','MEG','脑电图','脑磁图','血药浓度','视野检查','CT','ECG','MRS']):
                num+=1
            pexam_data=extract_exam_data_refined(b)
            if pexam_data:
                for i in pexam_data:
                    a.append(i)
        lists['list22'].append(a)
        lists['list23'].append(num-len(a))
        lists['list24'].append(location)
           

if __name__ == "__main__":
    #path
    folder_path = r'E:\病例整理\先兆\病历摘要\病历摘要（未融合）'
    file_path = r"C:\Users\11\Desktop\database\main\files\分类结果.xlsx"
    
    #convert to excel
    sd.process_folder(folder_path, file_path)
    
    #excel_script Variable setting
    sheet_name = 'Sheet1'
    excel = of.excel_extract(file_path, sheet_name)
    excel1 = of.excel_writer(file_path, sheet_name)
    
    # Extract columns from Excel and reverse to match original logic
    columns = [excel.col_byindex(i, None)[1] for i in range(1, 10)]
    
    # Initialize a dictionary for managing lists
    lists = {f"list{i}": [] for i in range(1, 37)}
    
    #extract data
    from_basic()
    from_filename()
    from_exam()
    from_pexam()
     
    # Writing data back to Excel from the dictionary
    excel1.write_list_to_column(11, lists['name_from_file'])
    for i in range(1, 37):
        excel1.write_list_to_column(i + 11, lists[f"list{i}"])
