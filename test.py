import Office操作脚本 as of
from extract_name import extract_names_with_new_rule as en
from extract_name import extract_first_chinese_phrase as ep
from extract_name import extract_all_ages_in_order as ea
from extract_name import extract_reason_or_complaint as ec
from extract_name import extractor 
from extract_name import extract_all_complaints
from extract_name import extract_exam_data
from extract_name import extract_exam_data_refined


import re

def remove_spaces_and_newlines(text):
    # 使用正则表达式去除空格和换行符
    if text and isinstance(text, str):
        cleaned_text = re.sub(r'\s+', '', text)
        return cleaned_text
    else:
        return 'None'


if __name__ =="__main__":
    file_path=r"D:\python code\sort_docx\files\分类结果.xlsx"
    sheet_name='Sheet1'
    excel=of.excel_extract(file_path,sheet_name)
    excel1=of.excel_writer(file_path,sheet_name)
    list1=excel.col_byindex(1,None)[1]
    list2=excel.col_byindex(9,None)[1]
    list_cc=excel.col_byindex(2,None)[1]
    list_exam=excel.col_byindex(8,None)[1]

    list3=[]
    list4=[]
    list5=[]
    list6=[]
    list7=[]
    list8=[]
    list9=[]
    list10=[]
    list11=[]
    list12=[]
    list13=[]
    list14=[]
    list15=[]
    list16=[]
    list17=[]
    list18=[]
    list19=[]
    list20=[]
    list21=[]
    list22=[]
    list23=[]
    list24=[]








    for n,i in enumerate(list1):
        i=remove_spaces_and_newlines(i)
        # gender =extractor(r'男|女',i)
        # name=en(i)
        # age=ea(i)
        # case_num=extractor(r"(?:病历号|病例号|病案号|住院号|病案)[：:]?(\d+[\*\d]*)",i)
        # image_num=extractor(r"(?:影像号|影响号)[：:]?(\d+)",i)
        # video_num=extractor(r"(?:脑电图号|脑电号)[：:]?(\d+[-\d]*)",i)
        # check_num=extractor(r"检查号[：:]?(\d+)",i)
        # phone_num=extractor(r"(?:电话号|电话)[：:]\D*(\d+)",i)
        chief=ec(i)
        
        
        #添加
        # list6.append(name)
        # list7.append(len(name))
        # list4.append(gender)
        # list5.append(len(gender))
        # list8.append(age)
        # list9.append(len(age))
        # list10.append(case_num)
        # list11.append(len(case_num))
        # list12.append(image_num)
        # list13.append(len(image_num))
        # list14.append(video_num)
        # list15.append(len(video_num))   
        # list16.append(check_num)
        # list17.append(len(check_num))
        # list18.append(phone_num)
        # list19.append(len(phone_num))   
        if chief:
            list20.append(chief)
            list21.append(len(chief))
        else:
            cc=remove_spaces_and_newlines(list_cc[n])
            cc=extract_all_complaints(cc)
            list20.append(cc)
            list21.append(len(cc))
            
        # list22.append(phone_num)
        # list23.append(len(phone_num))   
       
    #print(list20)    
        
    for i in list2:
        i=remove_spaces_and_newlines(i)
        name=ep(i)
        list3.append(name)

    for n,i in enumerate(list_exam):
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
        list22.append(a)
        list23.append(num-len(a))
        
    #写入

    # excel1.write_list_to_column(11,list2)
    # excel1.write_list_to_column(11,list3)
    # excel1.write_list_to_column(12,list6)
    # excel1.write_list_to_column(13,list7)


    # excel1.write_list_to_column(14,list4)
    # excel1.write_list_to_column(15,list5)

    # excel1.write_list_to_column(16,list8)
    # excel1.write_list_to_column(17,list9)

    # excel1.write_list_to_column(18,list10)
    # excel1.write_list_to_column(19,list11)

    # excel1.write_list_to_column(20,list12)
    # excel1.write_list_to_column(21,list13)
    # excel1.write_list_to_column(22,list14)
    # excel1.write_list_to_column(23,list15)

    # excel1.write_list_to_column(24,list16)
    # excel1.write_list_to_column(25,list17)
    # excel1.write_list_to_column(26,list18)
    # excel1.write_list_to_column(27,list19)
    #excel1.write_list_to_column(28,list20)
    #excel1.write_list_to_column(29,list21)
    excel1.write_list_to_column(30,list22)
    excel1.write_list_to_column(31,list23)
    #excel1.write_list_to_column(32,list24)
