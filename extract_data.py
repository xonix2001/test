# 再次调整函数以确保正确提取所有字段，特别是修正姓名字段的正则表达式问题
# 调整函数以改进信息提取逻辑，特别是针对姓名和性别的提取
import re
def extract_patient_info_v3(text):

    # 调整正则表达式模式，考虑到性别和姓名可能混合的情况
    patterns = {
        "姓名": r"病例讨论\$(\w+)",
        "性别": r"([男女])",
        "年龄": r"(\d+)岁",
        "主诉": "未提供",  # 示例中未提供主诉信息
        "病历号": r"病历号：(\d+)",
        "影像号": r"影像号：([^$\n]+)",
        "脑电图号": "未提供",  # 示例中未提供脑电图号信息
        "住址": r"岁([^$]+)\$"
    }

    # 创建一个字典，用于存储提取的数据
    extracted_data = {}

    # 特殊处理姓名和性别
    name_sex_age_address = re.search(r"病例讨论\$([^$]+)\$", text)
    if name_sex_age_address:
        nsaa_split = name_sex_age_address.group(1).split("岁", 1)
        if len(nsaa_split) > 1:
            extracted_data["住址"] = nsaa_split[1]
        name_sex = re.search(r"(\w+)(男|女)", nsaa_split[0])
        if name_sex:
            extracted_data["姓名"] = name_sex.group(1)
            extracted_data["性别"] = name_sex.group(2)
    else:
        extracted_data["姓名"] = "未提供"
        extracted_data["性别"] = "未提供"
        extracted_data["住址"] = "未提供"

    # 遍历定义的模式，使用正则表达式提取其他信息
    for key, pattern in patterns.items():
        if key in ["姓名", "性别", "住址"]:
            continue
        if pattern == "未提供":
            extracted_data[key] = "未提供"
        else:
            match = re.search(pattern, text)
            # 如果匹配成功，则将结果存储在字典中
            extracted_data[key] = match.group(1) if match else "未提供"

    return extracted_data


def extract_patient_data_v6(text):
    

    # 再次调整正则表达式模式，修正姓名字段的提取问题
    patterns = {
        "姓名": [r"姓名：(\w+)", r"患者(.*?)[,，]", r"^(\w+)[男女]性",r"^(\w+)[男女]"],
        "性别": [r"性别：(\w)", r",(\w),", r"，(\w)，", r"(\w)性"],
        "年龄": [r"年龄：(\d+)岁", r",(\d+)岁,", r"，(\d+)岁，", r"(\d+)岁"],
        "主诉": [r"主因\"(.*?)\"于", r"主因“(.*?)”于", r"\"(.*?)\"$"],
        "病历号": [r"病历号：(\d+)"],
        "影像号": [r"影像号：(\d*)"],
        "脑电图号": [r"脑电图号：(\d*)"],
        "住址": [r"患者来源：(.*?)$", r"住址：(.*?)$",r"籍贯：(.*?)(?:\$|病)"] # 更新籍贯的正则表达式，以适应不同的分隔符
    }

    # 创建一个字典，用于存储提取的数据
    extracted_data = {}

    # 遍历定义的模式，使用正则表达式提取信息
    for key, patterns_list in patterns.items():
        for pattern in patterns_list:
            match = re.search(pattern, text)
            if match:
                extracted_data[key] = match.group(1).replace("性别", "")  # 特别针对姓名字段的修正
                break  # 匹配到后就跳出循环
        else:  # 如果所有模式都不匹配，则赋值None
            extracted_data[key] = None

    return extracted_data

# 使用再次调整后的函数提取信息
# if __name__ == '__main__':
#     extracted_data_v6 = extract_patient_data_v6(new_text_2)
#     print(extracted_data_v6)
