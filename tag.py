import re

import Office操作脚本 as of

def remove_spaces_and_newlines(text):
    if text and isinstance(text, str):
        cleaned_text = re.sub(r'\s+', '', text)
        return cleaned_text
    else:
        return 'None'

def seletor(list_,keywords,e_keywords):        
    a=[]
    for i in list_:
        
        i=remove_spaces_and_newlines(i)
        
        if i != 'None':
            if i and any(key in i for key in keywords) and all(e_key not in i for e_key in e_keywords): 
                a.append(1)
            else:
                a.append(0)
        else:
            a.append('')
    return a
        
        

def main():
    try:
        file_path = r"C:\Users\11\Desktop\database\main\files\分类结果.xlsx"
        sheet_name = 'Sheet2'
        excel=of.excel_extract(file_path,sheet_name)
        excel1=of.excel_writer(file_path,sheet_name)
        
        list=excel.col_byindex(16,None)[1]
        #print(list)
        keywords_list=[['腹'],['胃'],['肠'],['肚'],['幻觉']]
        empty_lists = [[] for _ in range(len(keywords_list))]
        e_keywords_list=[]
        
        excel1.write_list_to_row(1,keywords_list,None,'Sheet3')
        
        for n,keys in enumerate(keywords_list):        
            empty_lists[n]=seletor(list,keys,e_keywords_list)
            print(empty_lists[n])
            #excel1.write_list_to_column(n+1,empty_lists[n],None,'Sheet3')
    except Exception as e:
        print(e)
if __name__ =="__main__":
    main()
    # list=['腹部1','现病史：患者于23年前（8岁时）无明显诱因突发意识障碍，伴咂嘴，无四肢抽搐、口吐白沫、舌咬伤或尿便失禁，数十秒后缓解。之后反复发作，表现为意识障碍，不应不语,咂嘴,汗毛竖立，皮肤起“鸡皮疙瘩”。有时可在精神恍惚状态下继续行走，事后无法回忆。不伴四肢抽搐，尿便失禁。发作前有先兆：自感胸闷，烦躁。口服卡马西平（,Bid）,丙戊酸镁（,Bid），有效 但不能完全控制：每月2-4次发作。为求手术治疗，收入我院。$', '现病史：于3年前（2岁半时）无明显诱因玩耍时出现左眼角发作性不自主抽搐、阵挛，持续数十秒后可缓解，后发作反复出现，每日 均有发作，且逐渐出现左眼角及口角同时抽搐，发作前有左眼疼痛的先兆，发作后可回忆整个过程，但发作中不能对答，持续十余秒至数十秒不等，白天夜间均有发作，有连发倾向。曾两次因高热后出 现“大发作”，表现为四肢强直，约两分钟后缓解，具体不详，后未再出现同类发作。曾至北大第一医院就诊，诊为“癫痫”，口服卡马西平、丙戊酸钠、妥泰等多种药物不能完全控制发作，现规律口服奥 卡西平早375mg，晚450mg；开普兰1#，Bid；利必通2#，Bid，每日仍有发作。$']
    # keywords_list=[['腹'],['胃'],['肠'],['肚'],['幻觉']]
    # empty_lists = [[] for _ in range(len(keywords_list))]
    # e_keywords_list=[]

    # for n,keys in enumerate(keywords_list):        
    #     empty_lists[n]=seletor(list,keys,e_keywords_list)
    #     print(empty_lists[n])

    