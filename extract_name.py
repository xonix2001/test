import re

def extract_names_with_new_rule(text):
    # 为了提高灵活性，对正则表达式进行调整，使其能够更准确地匹配不同格式的姓名
    patterns =[r'患者(.*?)[,，]|姓名(.*?)性别|讨论\$(.*?)[女男年龄]|摘要\$(.*?)[女男年龄]',r'患者(.*?)[,，]|姓名(.*?)性别|讨论\$(.*?)[女男年龄]|摘要\$(.*?)[女男年龄]|^(.*?)[女男]\d',r'摘要\$?(.*?)男|摘要\$?(.*?)女|患者(.*?)[,，]|姓名(.*?)性别|讨论\$(.*?)[女男年龄]|摘要\$(.*?)[女男年龄]|^(.*?)[女男]\d']
    pattern=patterns[2]
    names = []
    

    # 查找所有匹配的姓名
    matches = re.findall(pattern, text, re.DOTALL)
    # 提取姓名并添加到结果列表中
    
    for match in matches:
        # 由于使用了多个匹配组，合并并去除空字符串
        name = ''.join(match).strip()
        # 移除多余的匹配后缀
        name = re.sub(r'\d', '', name)
        name = re.sub(r'[:：；。，,$-]', '', name)
        pattern1 = r'姓名|安返病房|患者|病例|病历|摘要|讨论|自发以来|因失眠|左利手|意识清楚|开始|停药|间断|调药|双颞钻孔|自行停药|出现|发作|头向左转|出现持续|癫痫再次|入院以来|再次上述|反复|个月大时|次数增加'
        pattern2 = r'年前突发|再次上述|呕吐|再次大|形式改变|发病以来|全麻未醒|继发少尿|愣神|意识丧失|于年月日|次数增多|偶有白天|岁时|神志清楚|后又再次|每天次|足月顺产|意识不清|再次上述|于年前|突然倒地|仍癫痫|脑外伤史|每日次|性抽搐年|形式变化|自觉秒|主因|高热'

        name = re.sub(r'性别.*', '', name)
        
        name = re.sub(pattern1, '', name)
        name = re.sub(pattern2, '', name)
        #name = re.sub(r'[姓名患者病例历摘要讨论自发以来因失眠左利手意识清楚开始停药间断调药双颞钻孔自行停药出现发作头向左转出现持续]', '', name)
        # 移除性别及其后的内容
        if name and 1<len(name)<5:
            names.append(name)
    
                
    return names


# 重新定义函数，目的是提取首个中文词组作为姓名
def extract_first_chinese_phrase(filename):
    # 使用正则表达式匹配首个中文词组
    # 假设中文字符范围为\u4e00-\u9fa5
    match = re.search(r'[\u4e00-\u9fa5]+', filename)
    if match:
        return match.group(0)  # 返回匹配到的整个中文词组
    else:
        return ""


def extract_all_ages_in_order(text):
    # 调整正则表达式以更广泛地匹配年龄信息，保持文本出现的顺序
    pattern = r'年龄[:：]?\s*(\d{1,2})\s*岁?|(\d{1,2})岁'
    ages = []

    matches = re.findall(pattern, text)
    for match in matches:
        # 由于使用了两个匹配组，需要遍历并选择非空的匹配结果
        age = next((m for m in match if m), None)
        if age and age.isdigit():  # 确保是数字
            ages.append(age)

    return ages

def extractor(pattern,text):
    # 查找所有匹配的检查号
    matches = re.findall(pattern, text)
    
    # 提取检查号并添加到结果列表中
    check_numbers = []
    for match in matches:
        # 去除空字符串
        if match.strip():
            check_numbers.append(match.strip())
    
    return check_numbers

def extract_reason_or_complaint(text):

    #提取文本中的“因”或“主诉”及其描述。

    pattern = r'因[“"][.]?(.+?)[”"]|(?:主诉|主因)[：:](.+?)[。\$]'
    matches = re.finditer(pattern, text)
    extracted_texts = [match.group(1) if match.group(1) else match.group(2) for match in matches]
    return extracted_texts

def extract_all_complaints(text):
    pattern = r'主诉[：:][\$]?(.*?)\$'
    matches = re.findall(pattern, text)
    if matches:
        check_numbers = []
        for match in matches:
            # 去除空字符串
            if match.strip():
                check_numbers.append(re.sub(r"主诉?[。$.“”\"]", '', match))

        return check_numbers
    elif re.search(r'[\u4e00-\u9fff]', text) is not None:
        text = re.sub(r"[。$.“”\"]", '', text)
        return [text]
    else:
        return matches

def extract_exam_data(text):
    pattern = r"([\u4e00-\u9fff]*MRI|脑血流灌注显像|脑磁图|脑电图|PET|MEG|[V-]?EEG|发作期|发作间期)[：:]?[\(（]?(.*?)[\)）]?[：:]?(.*?)"
    matches = re.findall(pattern, text)
    return matches

def extract_exam_data_refined(text):
    pattern = r"([\u4e00-\u9fff]*[f]?MRI[平扫]{0,2}|[\u4e00-\u9fff]*MRS|脑血流灌注显像|[\u4e00-\u9fff]*[SPE]*CT|脑磁图|脑电图|[\u4e00-\u9fff]*PET|MEG|EMG|[V-]*EEG|ECG|病理|发作期|发作间期|运动诱发|体感诱发|癫痫定位|基因检测报告|视野检查|皮层电刺激|运动诱发电位|刺激左侧皮层|刺激右侧皮层|血常规|神经心理检查|盆腔Ｂ超|血药浓度)[：:]?\s*([\(（](.*?)[\)）])?[：:]?\s*(.*?)$"
    matches = re.findall(pattern, text)

    refined_matches = []
    for match in matches:
        exam_type, _, location, details = match
        # If location is empty, it means there were no parentheses, and what was captured as details are actually the details
        if not location and details:
            location = ''  # Ensure location is explicitly empty if no parentheses were present
        refined_matches.append((exam_type, details if details else '', location))

    return refined_matches

def extract_pexam_data_refined():
    




if __name__ == '__main__':
    text='头颅MRI（我院）：右侧海马硬化改变'
    results=extract_exam_data_refined(text)
    print(results,len(results),)
    
    
    
    
    
    # import Office操作脚本 as of
    
    # def remove_spaces_and_newlines(text):
    #     # 使用正则表达式去除空格和换行符
    #     if text and isinstance(text, str):
    #         cleaned_text = re.sub(r'\s+', '', text)
    #         return cleaned_text
    #     else:
    #         return 'None'

    
    # file_path=r"D:\python code\sort_docx\files\分类结果.xlsx"
    # sheet_name='Sheet1'
    # excel=of.excel_extract(file_path,sheet_name)
    # excel1=of.excel_writer(file_path,sheet_name)
    # # list1=excel.col_byindex(1,None)[1]
    # # list2=excel.col_byindex(9,None)[1]
    # # list_cc=excel.col_byindex(2,None)[1]
    
    # list_exam=excel.col_byindex(8,None)[1]
    
    
    # list3=[]
    # list4=[]
    # list5=[]
    # list6=[]
    # list7=[]
    # list8=[]
    # list9=[]
    # list10=[]
    # list11=[]
    # list12=[]
    # list13=[]
    # list14=[]
    # list15=[]
    # list16=[]
    # list17=[]
    # list18=[]
    # list19=[]
    # list20=[]
    # list21=[]
    # list22=[]
    # list23=[]
    # list24=[]
    
    # for n,i in enumerate(list_exam):
    #     i=remove_spaces_and_newlines(i)
    #     print(extract_exam_data(i))