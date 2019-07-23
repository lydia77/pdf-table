import os
import re
from pdfminer.pdfparser import PDFParser,PDFDocument,PDFPage
# from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter,PDFTextExtractionNotAllowed
from pdfminer.layout import *
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfparser import PDFSyntaxError
from pdfminer.psparser import PSEOF
import json
import datetime
from collections import defaultdict
import math




def extract_layout_by_page(pdf_path):

# 提取页面布局



    fp = open(pdf_path, 'rb')
    parser = PDFParser(fp)
    document = PDFDocument(parser)
    parser.set_document(document)
    document.set_parser(parser)
    document.initialize()

    # if not document.is_extractable:
    #     raise PDFTextExtractionNotAllowed


    laparams = LAParams()
    rsrcmgr = PDFResourceManager()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    layouts = []
    for page in document.get_pages():

        interpreter.process_page(page)
        layouts.append(device.get_result())

    return layouts

TEXT_ELEMENTS = [
    LTTextBox,
    LTTextBoxHorizontal,
    LTTextLine,
    LTTextLineHorizontal
]

def flatten(lst):
    """Flattens a list of lists"""
    return [subelem for elem in lst for subelem in elem]

def extract_characters(element):
    """
    Recursively extracts individual characters from
    text elements.
    """
    if isinstance(element, LTChar):
        return [element]

    if any(isinstance(element, i) for i in TEXT_ELEMENTS):
        return flatten([extract_characters(e) for e in element])

    if isinstance(element, list):
        return flatten([extract_characters(l) for l in element])

    return []


def width(rect):
    x0, y0, x1, y1 = rect.bbox
    return min(x1 - x0, y1 - y0)

def area(rect):
    x0, y0, x1, y1 = rect.bbox
    return (x1 - x0) * (y1 - y0)


def cast_as_line(rect):
    # 用最长维度的线代替矩形

    x0, y0, x1, y1 = rect.bbox

    if x1 - x0 > y1 - y0:
        return (x0, y0, x1, y0, "H")
    else:
        return (x0, y0, x0, y1, "V")

def does_it_intersect(x, xmin, xmax):
    return (x <= xmax and x >= xmin)

def find_bounding_rectangle(x, y, lines):

        # Given a collection of lines, and a point, try to find the rectangle
        # made from the lines that bounds the point. If the point is not
        # bounded, return None.
        # 寻找字符边界

        v_intersects = [l for l in lines
                        if l[4] == "V"
                        and does_it_intersect(y, l[1], l[3])]

        h_intersects = [l for l in lines
                        if l[4] == "H"
                        and does_it_intersect(x, l[0], l[2])]

        if len(v_intersects) < 2 or len(h_intersects) < 2:
            return None

        v_left = [v[0] for v in v_intersects
                  if v[0] < x]

        v_right = [v[0] for v in v_intersects
                   if v[0] > x]

        if len(v_left) == 0 or len(v_right) == 0:
            return None

        x0, x1 = max(v_left), min(v_right)

        h_down = [h[1] for h in h_intersects
                  if h[1] < y]

        h_up = [h[1] for h in h_intersects
                if h[1] > y]

        if len(h_down) == 0 or len(h_up) == 0:
            return None

        y0, y1 = max(h_down), min(h_up)

        return (x0, y0, x1, y1)

def chars_to_string(chars):

        # 将字符集转化为字符串
        if not chars:
            return ""
        rows = sorted(list(set(c.bbox[1] for c in chars)), reverse=True)
        text = ""
        for row in rows:
            sorted_row = sorted([c for c in chars if c.bbox[1] == row], key=lambda c: c.bbox[0])
            text += "".join(c.get_text() for c in sorted_row)
        return text

def boxes_to_table(box_record_dict):

        # 将单元格-字符 字典转换为行列table
        # of lists of strings. Tries to split cells into rows, then for each row
        # breaks it down into columns.
        #
        boxes = box_record_dict.keys()
        rows = sorted(list(set(b[1] for b in boxes)), reverse=True)
        table = []
        for row in rows:
            sorted_row = sorted([b for b in boxes if b[1] == row], key=lambda b: b[0])
            table.append([chars_to_string(box_record_dict[b]) for b in sorted_row])
        return table

def extract_tables(current_page):
    """
    提取所有当前页面的所有表格
    :param current_page: 当前页面
    :return: tables：嵌套列表，列表每个元素为每行表格
    """
    texts, rects = [], []
    # seperate text and rectangle elements
    for e in current_page:
        if isinstance(e, LTTextBoxHorizontal):
            texts.append(e)
        elif isinstance(e, LTRect):
            rects.append(e)

    characters = extract_characters(texts)

    lines = [cast_as_line(r) for r in rects
             if width(r) < 2 and
             area(r) > 1]

    box_char_dict = {}
    for c in characters:
        # choose the bounding box that occurs the majority of times for each of these:
        bboxes = defaultdict(int)
        l_x, l_y = c.bbox[0], c.bbox[1]
        bbox_l = find_bounding_rectangle(l_x, l_y, lines)
        bboxes[bbox_l] += 1

        c_x, c_y = math.floor((c.bbox[0] + c.bbox[2]) / 2), math.floor((c.bbox[1] + c.bbox[3]) / 2)
        bbox_c = find_bounding_rectangle(c_x, c_y, lines)
        bboxes[bbox_c] += 1

        u_x, u_y = c.bbox[2], c.bbox[3]
        bbox_u = find_bounding_rectangle(u_x, u_y, lines)
        bboxes[bbox_u] += 1

        # 以三个点为基准寻找边框，若相同则边框定
        # 若不相同以中心点边框为准
        # if all values are in different boxes, default to character center.
        # otherwise choose the majority.
        if max(bboxes.values()) == 1:
            bbox = bbox_c
        else:
            bbox = max(bboxes.items(), key=lambda x: x[1])[0]

        if bbox is None:
            continue

        if bbox in box_char_dict.keys():
            box_char_dict[bbox].append(c)
            continue

        box_char_dict[bbox] = [c]

    xmin, ymin, xmax, ymax = current_page.bbox


    for x in range(int(xmin), int(xmax), 10):
        for y in range(int(ymin), int(ymax), 10):
            bbox = find_bounding_rectangle(x, y, lines)

            if bbox is None:
                continue
            if bbox in box_char_dict.keys():
                continue

            box_char_dict[bbox] = []

    # with open('test5.txt', 'a') as f:
    #     for i in box_char_dict.items():
    #         f.write(str(i)+'\n')

    tables = boxes_to_table(box_char_dict)

    # with open('D:\work\公众公司抽取/tables/table7.xlsx', 'a') as output:
    #     for i in range(len(tables)):
    #         for j in range(len(tables[i])):
    #             if tables[i][j] is '':
    #                 continue
    #             x = str(tables[i][j]).replace(' ','')
    #             output.write(x)  # write函数不能写int类型的参数，所以使用str()转化
    #             output.write('\t')  # 相当于Tab一下，换一个单元格
    #         output.write('\n')  # 写完一行立马换行
    # output.close()
    return tables

def extract_head(path):
    """
    从路径中提取标题、证券代码和简称
    :param i: "json/000002-万科A-关于监事辞职的公告.txt"
    :return: dic
    """
    filename = os.path.split(path)[-1]
    i = os.path.splitext(filename)[0]
    p1 = re.compile(r'^\d{6}')
    a = re.search(p1, i)[0]
    p2 = re.compile(r'(?<=-).+?(?=-)')
    b = re.search(p2, i)[0]
    # p3 = re.compile(r'(?<=-).+?$')
    # title = re.search(p3, i)[0]
    # print(a, b, c)
    dic = {}
    dic["证券代码"] = a
    dic["证券简称"] = b
    return dic,i





def pdf2json(pdf_path):
    """
    从PDF提取想要的数据
    :param pdf_path: pdf文件位置
    :return: json
    """
    page_layouts = extract_layout_by_page(pdf_path)
    all_tables = []
    # 提取所有表格为嵌套列表
    for current_page in page_layouts[15:]:
        tables = extract_tables(current_page)
        all_tables.extend(tables)
    # "现金流量表（母公司）": {
    #     "单位": "元",
    #     "项目": [{},{}]

    asset1, asset2, profit1, profit2, cash1, cash2 = [], [], [], [], [], []
    #    六个表的项目内容列表
    # 查找六个表的开始位置
    position = []
    i = 0
    for table in all_tables:
        # print(table)
        if table==['项目 ', '附注 ', '期末余额 ', '期初余额 '] or table==['项目 ', '附注 ', '本期金额 ', '上期金额 ']:
            position.append(i)
            # print(table)
        i+=1
    # [176, 274, 354, 409, 444, 502]
    # print(position)
    p1=position[0]
    p2 = position[1]
    p3 = position[2]
    p4 = position[3]
    p5 = position[4]
    p6 = position[5]
    for table in all_tables[p1+1:p2]:
        if len(table)==4:
            dic={
            "名称": table[0],
            "附注": table[1],
            "年初至报告期末": table[2],
            "上年年初至报告期末": table[3]}
            asset1.append(dic)
        else:
            print(table)

    for table in all_tables[p2+1:p3]:
        if len(table)==4:
            dic={
            "名称": table[0],
            "附注": table[1],
            "年初至报告期末": table[2],
            "上年年初至报告期末": table[3]}
            asset2.append(dic)
        else:
            print(table)

    for table in all_tables[p3+1:p4]:
        if len(table)==4:
            dic={
            "名称": table[0],
            "附注": table[1],
            "年初至报告期末": table[2],
            "上年年初至报告期末": table[3]}
            profit1.append(dic)
        else:
            print(table)

    for table in all_tables[p4+1:p5]:
        if len(table)==4:
            dic={
            "名称": table[0],
            "附注": table[1],
            "年初至报告期末": table[2],
            "上年年初至报告期末": table[3]}
            profit2.append(dic)
        else:
            print(table)

    for table in all_tables[p5+1:p6]:
        if len(table)==4:
            dic={
            "名称": table[0],
            "附注": table[1],
            "年初至报告期末": table[2],
            "上年年初至报告期末": table[3]}
            cash1.append(dic)
        else:
            print(table)
    for table in all_tables[p6+1:p6+40]:
        if len(table) == 4:
            dic = {
                "名称": table[0],
                "附注": table[1],
                "年初至报告期末": table[2],
                "上年年初至报告期末": table[3]}
            cash2.append(dic)
        else:
            print(table)

    a1={"资产负债表（合并）": {"单位": "元","项目": asset1}}
    a2={"资产负债表（母公司）":{"单位": "元","项目": asset2}}
    pro1={"利润表（合并）": {"单位": "元","项目": profit1}}
    pro2={"利润表（母公司）": {"单位": "元", "项目": profit2}}
    c1={"现金流量表（合并）": {"单位": "元","项目": cash1}}
    c2={"现金流量表（母公司）": {"单位": "元","项目": cash2}}
    dic, title = extract_head(pdf_path)

    result={"证券代码":dic["证券代码"],
            "证券简称":dic["证券简称"],
            "现金流量表（母公司）":c2["现金流量表（母公司）"],
            "现金流量表（合并）":c1["现金流量表（合并）"],
            "利润表（母公司）":pro2["利润表（母公司）"],
            "利润表（合并）":pro1["利润表（合并）"],
            "资产负债表（母公司）":a2["资产负债表（母公司）"],
            "资产负债表（合并）":a1["资产负债表（合并）"]}
    # result={}
    # result.update(dic)


    final_result={title:result}
    # print(final_result)
    # output_filename='tables/'+title+'.json'
    # with open(output_filename,'w',encoding='utf-8')as f:
    #     json.dump(final_result,f,ensure_ascii=False)
    final_result=json.dumps(final_result,indent=1,ensure_ascii=False)
    return final_result




if __name__=='__main__':
    # test

    print("start",datetime.datetime.now())

    pdf_path='CCKS评测任务5数据及说明/年报财务报表训练数据/财务报表_pdf_2017年报/430189-摩点文娱-2017年年度报告.pdf'

    result=pdf2json(pdf_path)

    print(result)

    print("end",datetime.datetime.now())

