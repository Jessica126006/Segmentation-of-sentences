person_id=['53f3628adabfae4b3498cb6f']
error = {}
    for i in person_id:
        # if i == '53f3628adabfae4b3498cb6f':
            old_excel = openpyxl.load_workbook(i+'.xlsx')
            # 从0开始计数
            ws = old_excel.worksheets[9]
            # 删除年份为0的行
            if ws['A2'].value==0:
                ws.delete_rows(2)
            # for cell in list(ws.columns)[3]:
            result_20 = []
            ws["F1"].value = 'key_20'
            for row in range(2,ws.max_row+1):
                # 分词(中文也可以,所有中文分为一段)
                try :
                     word=word_tokenize(ws["D%d" % (row)].value)
                except Exception as e:
                    error[i] = row
                # print(word)
                word_filter=[]
                for wo in word:
                    num=re.search('^[0-9a-zA-Z_-]{1,}$',wo)
                    # 在去除停用词前把字母变小写，并过滤掉非法字符,去掉长度小于5的关键词
                    if wo.lower() not in english_stopwords and num and len(wo)>=5:
                        word_filter.append(wo)
                print(word_filter)

                # 将分词后的结果转化为文档，统计固定搭配
                word_text = Text(word_filter)
                double_key = word_text.collocations(num=20, window_size=2)
                double_key_list=double_key.split("; ")
                print(double_key_list)

                # 将分词结果标注词性
                tag_word = pos_tag(word_filter, tagset='universal')
                VERB = []
                NOUN = []
                ADJ = []
                for a,b in tag_word:
                    if b == "NOUN":
                        NOUN.append(a)
                    elif b == "VERB":
                        VERB.append(a)
                    elif b == "ADJ":
                        ADJ.append(a)
                # print(VERB)
                # 分别统计名词和动词出现次数
                noun = {}
                verb = {}
                adj = {}
                # ------名词---------
                for no in NOUN:
                    if NOUN.count(no) >= 1:
                        noun[no] = NOUN.count(no)
                        # 对字典value(重复次数)降序排序，以列表形式返回
                noun = sorted(noun.items(), key=lambda item: item[1], reverse=True)
                NOUN_strip = []
                # 获取列表中每个元组中的第一个值，即数组中的数字（此时已是按重复次数降序）
                for item in noun:
                    NOUN_strip.append(item[0])

                # ------动词---------
                for vb in VERB:
                    if VERB.count(vb) >= 1:
                        verb[vb] = VERB.count(vb)
                        # 对字典value(重复次数)降序排序，以列表形式返回
                verb = sorted(verb.items(), key=lambda item: item[1], reverse=True)
                VERB_strip = []
                # 获取列表中每个元组中的第一个值，即数组中的数字（此时已是按重复次数降序）
                for item in verb:
                    VERB_strip.append(item[0])
                # ------形容词---------
                for ad in ADJ:
                    if ADJ.count(ad) >= 1:
                        adj[ad] = ADJ.count(ad)
                        # 对字典value(重复次数)降序排序，以列表形式返回
                adj = sorted(adj.items(), key=lambda item: item[1], reverse=True)
                ADJ_strip = []
                # 获取列表中每个元组中的第一个值，即数组中的数字（此时已是按重复次数降序）
                for item in adj:
                    ADJ_strip.append(item[0])

                print('dongci', VERB_strip)
                print('mingci', NOUN_strip)
                print('xingrongci',ADJ_strip)

                # 读取中英文关键词
                final_key = ws["E%d" % (row)].value[2:-2]
                print(final_key)
                # 将中英文关键词从字符串变数组
                result = final_key.split("', '")
                """
                对于中英文关键词小于20的从题目抽取来补充加进去，大于20的把长度小于5的剔除，含有非法字符
                （除了字母，数字，下划线，中下划线的字符）的删除，这时观察长度区间，在20-100的居多，少部分在200到600之间
                 于是20-50的就取中间部分，大于50的就随机选择20个关键词
                """
                # 过滤掉中英文关键词的非法字符,智能含有中文，或字母，数字，中下划线和下划线
                result_filter = []
                for tem in result:
                    if re.search('^[\u4e00-\u9fa5_a-zA-Z0-9- /s]+$',tem) and len(tem)>5:
                        result_filter.append(tem)
                result_filter.sort(key=lambda i: len(i), reverse=True)
                print(result_filter)
                print(len(result_filter))

                # 判断已有的关键词个数
                if len(result_filter) >= 20 and len(result_filter) <= 50:
                    index=(len(result_filter)-20)//2
                    result_filter = result_filter[index:index+20]
                if len(result_filter) > 50 :
                    # 随机选取20个关键词
                    result_filter = random.sample(result_filter,20)
                if len(result_filter) < 20:
                    for l in double_key_list:
                        if len(result_filter) < 20 and l!='':
                            result_filter.append(l)

                if len(result_filter) < 20:
                    diff = 20 - len(result_filter)
                    if len(NOUN_strip) > diff:
                        for n in NOUN_strip[:diff]:
                            result_filter.append(n)
                            result_20 = result_filter
                    else:
                        for n in NOUN_strip:
                            result_filter.append(n)

                print('kai*******',result_filter)

                if len(result_filter) < 20:
                    diff2 = 20 - len(result_filter)
                    if len(VERB_strip) > diff2:
                        for v in VERB_strip[:diff2]:
                            result_filter.append(v)
                            result_20 = result_filter
                    else:
                        for v in VERB_strip:
                            result_filter.append(v)

                if len(result_filter) < 20:
                    diff3 = 20 - len(result_filter)
                    if len(ADJ_strip) > diff3:
                        for a in ADJ_strip[:diff3]:
                            result_filter.append(a)
                            result_20 = result_filter
                    else:
                        for a in ADJ_strip:
                            result_filter.append(a)
                result_20=result_filter
                # if '' in result_20:
                #     result_20.remove('')
                ws["F%d" % (row)].value = str(result_20)
                print(result_20)
            print('-------')
            old_excel.save(i+'.xlsx')
            print(error)
