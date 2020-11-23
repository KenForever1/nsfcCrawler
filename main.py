import requests
import json
import xlsxwriter

headers = {
    'Host': 'output.nsfc.gov.cn',
    'Origin': 'http://output.nsfc.gov.cn',
    'Connection': 'keep-alive',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/86.0.4240.75 Safari/537.36',
    'Content-Type': 'application/json;charset=UTF-8',
    'Accept': '*/*',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Dest': 'empty',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'zh-CN,zh;q=0.9',
}

years = ["2015", "2014"]
code = ["B01", "B02", "B03", "B04", "B05", "B06", "B07", "B08"]
search_data_format = {"ratifyNo": "", "projectName": "", "personInCharge": "", "dependUnit": "", "code": "B04",
                      "projectType": "429", "subPType": "", "psPType": "", "keywords": "", "ratifyYear": "",
                      "conclusionYear": "2015", "beginYear": "", "endYear": "", "checkDep": "", "checkType": "",
                      "quickQueryInput": "", "adminID": "", "pageNum": 0, "pageSize": 5, "queryType": "input",
                      "complete": "true"}

search_url = "http://output.nsfc.gov.cn/baseQuery/data/conclusionQueryResultsData"

all_result = []

row = 0

workbook = xlsxwriter.Workbook('my_table_other1.xlsx')
worksheet = workbook.add_worksheet(name="化学期刊")


def post_search(search_data):
    task_map = {}
    resp = requests.post(url=search_url, data=json.dumps(search_data), headers=headers)
    res_dict = resp.json()['data']['resultsData']
    for item in res_dict:
        print(item)
        url_path_id = item[2]
        teacher_name = item[5]
        print(url_path_id)
        print(teacher_name)
        task_map[teacher_name] = url_path_id
    # return task_map
    get_every_teacher_paper_task(task_map)


def get_every_teacher_paper_task(task):
    # task = post_search()

    task_project_techername_url = {}
    url_base_str = "http://output.nsfc.gov.cn/baseQuery/data/conclusionProjectInfo/"
    for teachername, url_path_id in task.items():
        url_path = url_base_str + url_path_id
        task_project_techername_url[teachername] = url_path
    # return task_project_techername_url
    handle_every_teacher_all_paper(task_project_techername_url)


def handle_every_teacher_all_paper(task_project_url):
    # 处理一个老师
    # 每个url 代表一个老师的一个项目
    for teacher_name, url in task_project_url.items():
        # 保存每个老师期刊名称和数量的dict
        res_dict = {}
        paper_url_list = []
        base_paper_url_str = "http://output.nsfc.gov.cn/baseQuery/data/resultsInfoData/"
        print(url)
        resp = requests.get(url, headers=headers)
        # with open('D:\\fileworkspace\\tjr\\nsfcCrawler\\html.txt', 'a') as f:
        #     f.write(str(json.dumps(resp.json()).encode("utf-8")))
        paper_url_result_list = resp.json()['data']['resultsList']
        print(paper_url_result_list)
        for item in paper_url_result_list:
            # 增加判断 只加入“会议期刊”
            paper_type = item['result'][3]
            # print(paper_type)
            if paper_type == "期刊论文":
                paper_url_id = item['result'][1]
                # print(paper_url_id)
                paper_url = base_paper_url_str + paper_url_id
                paper_url_list.append(paper_url)

        # 处理paper_url_list
        for paper_url in paper_url_list:
            resp = requests.get(paper_url, headers=headers)
            try:
                res_json = resp.json()
            except:
                # str = resp.text
                # print(str)
                with open('D:\\fileworkspace\\tjr\\nsfcCrawler\\except_other1.txt', 'a', encoding='utf-8') as f:
                    f.write(teacher_name + "\n")
                    f.write(paper_url)
                    f.write("\n")
                continue
            else:
                print(paper_url)
                journalName = res_json['data']['journalName']
                print(journalName)
                if journalName not in res_dict.keys():
                    res_dict[journalName] = 1
                else:
                    res_dict[journalName] = res_dict[journalName] + 1
        # print(res_dict)
        one_result = {}
        one_result["teacher_name"] = teacher_name
        one_result["result"] = res_dict
        print(one_result)
        all_result.append(one_result)


def sava_to_xlsx():
    global row
    for item in all_result:
        teacher_name = item["teacher_name"]
        one_res_dict = item["result"]
        worksheet.write(row, 0, teacher_name)
        for paper_name, num in one_res_dict.items():
            worksheet.write(row, 1, paper_name)  # 第i行1列
            worksheet.write(row, 2, paper_name.lower())
            worksheet.write(row, 3, num)
            row = row + 1
        worksheet.write(row, 0, "")
        row = row + 1

    worksheet.write(row, 0, "")
    row = row + 1


def get_by_projectId():
    url_base_str = "http://output.nsfc.gov.cn/baseQuery/data/conclusionProjectInfo/"
    task_map = {
        "陈传峰": "20625206",
        "陈永明": "20625412",
        "吕小兵": "20625414",
        "李永旺": "20625620",
        "张广照": "20725414",
        "霍启升": "20788101",
        "胡文兵": "20825415",
        "孙平川": "20825416",
        "赵江": "20925416",
        "夏海平": "20925208",
        "施章杰": "20925207",
        "张书圣": "21025523",
        "陆豪杰": "21025519",
        "朱新远": "21025417",
        "罗三中": "21025208",
        "雷爱文": "21025206"
    }

    task_map1 = {}
    task_map1 ={
         "吕小兵": "20625414"
    }

    get_every_teacher_paper_task(task_map1)





if __name__ == '__main__':
    # for i_code in code:
    #     # handle_every_teacher_all_paper()
    #     search_data_format["code"] = i_code
    #     search_data = search_data_format
    #     # print(search_data)
    #     post_search(search_data)
    #     # print(all_result)
    #     # # with open('D:\\fileworkspace\\tjr\\nsfcCrawler\\result.txt', 'a') as f:
    #     # #     f.write(str(all_result))

    get_by_projectId()
    sava_to_xlsx()

    workbook.close()
