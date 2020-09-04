from openpyxl import Workbook
from openpyxl import load_workbook
from collections import OrderedDict
from parseconf import ParseConf
from directory import Directory

def readconfig():
    parseconf = ParseConf("config.ini")
    directory = {}
    rule = {}
    directory["input"] = parseconf.parseStr("directory", "input")
    directory["result"] = parseconf.parseStr("directory", "result")
    rule["function"] = parseconf.parseStr("rule", "function")
    rule["condition"] = parseconf.parseList("rule", "condition")
    rule["flag"] = parseconf.parseStr("rule", "flag")
    configdict ={"directory":directory, "rule":rule}
    return configdict

def getFilelist(path):
    filelist = []
    directory = path.strip()
    templist = directory.split(".")
    if (templist[-1] == "xlsx"):
        filelist.append(directory)
    else:
        files = Directory.getFiles(directory)
        for file in files:
            templist = file.split(".")
            if (templist[-1] == "xlsx"):
                filelist.append(file)
    return filelist

def read_excel(filename):
    '''
    提取excel文件数据
    '''
    datas = []
    wb = load_workbook(filename)
    for sheetname in wb.sheetnames:
        #datas[sheetname] = []
        top_flag = True
        title_list = []
        data_list = []
        for row in wb[sheetname].rows:
            data_dict = OrderedDict()
            if top_flag:
                for cell in row:
                    title_list.append(cell.value)
                top_flag = False
            else:
                for index in range(0, len(row)):
                    data_dict [title_list[index]] = row[index].value
                data_list.append(data_dict)
        datas.extend(data_list)
    wb.close()
    return datas

def write_excel(filename, datalist, flag):
    '''
    将合并后的数据写入excel文件
    '''
    #print(datalist[0].keys())
    wb = Workbook()
    ws = wb.active
    titlelist = []
    if flag == "1":
        for i in range(len(datalist)-1, -1,-1):
            if datalist[i]["flag"] == 1:
                del datalist[i]
            else:
                del datalist[i]["flag"]
    if (len(datalist)>0):
        titlelist = list(datalist[0].keys())
    print(datalist)
    ws.append(titlelist)
    for i in range(len(datalist)):
        ws.append(list(datalist[i].values()))
    wb.save(filename)
    wb.close()


def simplify(configdict):
    files = getFilelist(configdict["directory"]["input"])
    conditions = configdict["rule"]["condition"]
    flag = configdict["rule"]["flag"]
    count = 0
    filedict = {}
    datalist = []
    databaklist = []
    uidlist = []
    conditionlist = []
    print(files)
    for file in files:
        filelist = file.split("\\")
        filename = filelist[-1]
        uidlist = []
        datas = read_excel(file)
        datasbak = datas.copy()
        for i in range(len(datas)):
            datas[i]["index"] = i + count
            datas[i]["delindex"] = [i+count]
            tmp = ""
            for condition in conditions:
                tmp = tmp + str(datas[i][condition])
                #将要比较的字段拼接在一起
            datas[i]["condition"] = tmp
        datalist.extend(datas)
        databaklist.extend(datasbak)
        print(len(datas))
        #记录各文件大小，用例保存时按照各文件进行保存
        filedict[filename] = (count, count+len(datas))
        count = count +len(datas)
    for data in datalist:
        uidlist.append(data["uid"])
    for i in range(len(uidlist)-1, -1, -1):
        #逆序遍历，避免访问越界
        #多轮交互为方便比较，仅保留首轮交互（首轮会带上后面的交互信息），并删除其他轮交互用例
        if (uidlist.index(uidlist[i]) != i):
            j = uidlist.index(uidlist[i])
            datalist[j]["delindex"].append(i)
            #将删除用例id记录在被保留的首轮交互中
            datalist[j]["condition"] = datalist[j]["condition"] + datalist[i]["condition"]
            del datalist[i]
    for data in datalist:
        conditionlist.append(data["condition"])
    multilist = []
    #多轮合并处理为单轮后，与单轮交互一并处理判断是否重复，若重复将用例记录下来
    for i in range(len(conditionlist)):
        if (conditionlist.index(conditionlist[i]) != i):
            multilist.append(i)
    delindexlist = []
    for multi in multilist:
        #根据重复用例回推在原始列表中重复的用例索引
        delindexlist.extend(datalist[multi]["delindex"])
    print(datalist)
    print(databaklist)
    print(delindexlist)
    for i in range(len(databaklist)):
        del databaklist[i]["delindex"]
        del databaklist[i]["index"]
        del databaklist[i]["condition"]
        if i in delindexlist:
            databaklist[i]["flag"] = 1
        else:
            databaklist[i]["flag"] = ""
    count_src = len(databaklist)
    count_del = len(delindexlist)
    count_dest = count_src - count_del
    for key in filedict.keys():
        filepath = configdict["directory"]["result"] + "/" + key
        print(filedict)
        write_excel(filepath, databaklist[filedict[key][0]:filedict[key][1]], flag)
    print("###################result######################")
    print("原始用例条数：",count_src)
    print("删除用例条数：", count_del)
    print("剩余用例条数：", count_dest)
    print("###################result######################")
    return 0

def inspect(configdict):
    print('请检查config.ini参数function设置是否正确!')

def run():
    configdict = readconfig()
    function_dic = {
        '0' : simplify,

    }
    function_dic.get(configdict["rule"]["function"], inspect)(configdict)

if __name__ == "__main__":
    run()