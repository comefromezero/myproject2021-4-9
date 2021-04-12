#_*_ coding: utf-8 _*_
import xlwt

import requests

from pyquery import PyQuery as pq

import chardet
import xlwt
import sys
class Excel:
    def __init__(self,style):
        self.style = style;
        self.workbook = xlwt.Workbook(encoding='utf-8');
        self.sheet={};
    def new_sheet(self,name):
        self.sheet[name] = self.workbook.add_sheet(name);

    def write(self,sheet_name,row,col,data,style=0):
        mystyle = 0;
        if 0 != style:
            mystyle = style;
        else:
            mystyle = self.style;
        self.sheet[sheet_name].write(row,col,data,mystyle);
    def save(self,file_name):
        self.workbook.save(file_name);


class Style:
    def __init__(self):
        self.style = xlwt.XFStyle();
        self.font = xlwt.Font();
        self.font.name = "楷体"; #设置字体
        self.font.height = 20*12;#20号字体 #设置字体大小,20为基本单位，20表示1号字体，20*1：1号，20*2：2号。t
        self.alignment = xlwt.Alignment();
        self.alignment.horz = self.alignment.HORZ_CENTER;
        self.alignment.vert = self.alignment.VERT_CENTER;
        self.style.font = self.font;
        self.style.alignment = self.alignment;
class __Autonomy__(object):
    """ 自定义变量的write方法 """
    def __init__(self):
        """ init """
        self._buff = ""
    def write(self, out_stream):
        """ :param out_stream: :return: """
        self._buff += out_stream
try:
    #初始化表格
    style = Style();
    excel = Excel(style.style);
    sheet_name = "手机信息";
    excel.new_sheet(sheet_name);
    excel.write(sheet_name,0,0,"手机名");
    excel.write(sheet_name,0,1,"上市时间");
    excel.write(sheet_name,0,2,"屏幕尺寸");
    excel.write(sheet_name,0,3,"屏幕分辨率");
    excel.write(sheet_name,0,4,"操作系统");
    excel.write(sheet_name,0,5,"运行内存");
    excel.write(sheet_name,0,6,"CPU型号");
    file_name = "E:\\SmartPhoneSpider\\phone.xls";
    excel.save(file_name);
    r = requests.get("https://product.cnmo.com/",timeout = 10);

    if 200 == r.status_code:
        doc_nav = pq(r.text);
    else:
        print("首页http请求失败，错误码：",r.status);
    a_array = doc_nav(".navbox .navDiv a");
    print("Get all a elements of nav!");
    print(a_array);
    for item in a_array.items():
        if "手机大全" == item.text():
            smart_phone_url = item.attr("href");
            print("Have getted smart_phone_url:");
            print("https:" + smart_phone_url);
            break;
        else:
            print("not get 手机大全 url！");
            print("current url:",item.text());
    next_page_url = smart_phone_url;
    i = 1;
    while(next_page_url):
        #print(next_page_url);
        current_html = requests.get("https:" + next_page_url,timeout = 10);
        if 200 == current_html.status_code:
            cur_doc = pq(current_html.text);
            #print("xxxxx");
        else:
            print("http请求失败，错误码：",current_html.status,"https:" + next_page_url);
        next_page_url = cur_doc("div.all-con-con div.all-con-shaix div.page-up a.pnext").attr("href");
        print(next_page_url);
        all_cur_html_a = cur_doc("ul.all-con-con-ul div.info");
        url_list = [];
        phone_info_list = [];
        #print("test",all_cur_html_a);
        for item in all_cur_html_a.items():
            url_tmp=[];
            url_tmp.append(item("a").eq(0).attr("href"));
            url_tmp.append(item("a").eq(1).attr("href"));
            url_list.append(url_tmp);
        #print(url_list);
        #开始遍历当前页的所有手机的信息
        #手机的信息包括：
        #手机名
        #上市时间
        #屏幕尺寸
        #屏幕分辨率
        #操作系统
        #CPU型号
        #运行内存
        for phone in url_list:
            phone_info_map = {};
            phone_info_one = requests.get("https:" + phone[0],timeout=10);
            phone_info_doc_one = pq(phone_info_one.text);
            phone_info_two = requests.get("https:" + phone[1],timeout=10); 
            phone_info_doc_two = pq(phone_info_two.text);
            try:
                phone_info_map["手机名"] = phone_info_doc_one("div.cell-con-tit h2.pro-part-tit a").text();
                #print(phone_info_map["手机名"]);
                phone_info_map["上市时间"] = phone_info_doc_one("div.cell-con-tit div.cell-price p.p-left").eq(1).text().split("：")[1];
                #print(phone_info_map["上市时间"]);
#############################################################################################################################################################################################################################
                temp_p = phone_info_doc_two("div.cell-con-ul#cell-con-table ul").eq(2)("p");
                for item in temp_p.items():
                    if "屏幕尺寸" == item.attr("paramname"):
                        phone_info_map["屏幕尺寸"] = item.text().split("英寸")[0].lstrip().rstrip();                          #phone_info_doc_two("div.cell-con-ul#cell-con-table ul").eq(1)("p").eq(1).text().split("：")[1].split("英寸")[0].lstrip().rstrip();
                        break;
                    else:
                        phone_info_map["屏幕尺寸"] = "--";
                #print(phone_info_map["屏幕尺寸"]);
                for item in temp_p.items():
                    if "屏幕分辨率" == item.attr("paramname"):
                        phone_info_map["屏幕分辨率"] = item.text().split("像素")[0].lstrip().rstrip();                        #phone_info_doc_two("div.cell-con-tit ul.cell_phone_param li").eq(1)("p");#.text().split("：")[1].split("像素")[0].lstrip().rstrip();
                        break;
                    else:
                        phone_info_map["屏幕分辨率"] = "--";
                #print(phone_info_map["屏幕分辨率"]);

                temp_p = phone_info_doc_two("div.cell-con-ul#cell-con-table ul").eq(3)("p");
                for item in temp_p.items():
                    if "操作系统" == item.attr("paramname"):
                        phone_info_map["操作系统"] = item.text().lstrip().rstrip();                     #phone_info_doc_two("div.cell-con-tit ul.cell_phone_param li").eq(3)("p").eq(1).text().split("：")[1].split("；")[0].lstrip().rstrip();
                        break;
                    else:
                        phone_info_map["操作系统"] = "--";
                #print(phone_info_map["操作系统"]);
                for item in temp_p.items():
                    if "运行内存" == item.attr("paramname"):
                        phone_info_map["运行内存"] = item.text().lstrip().rstrip();                                  #phone_info_doc_two("div.cell-con-tit ul.cell_phone_param li").eq(3)("p").eq(1).text().split("：")[2].lstrip().rstrip();
                        break;
                    else:
                        phone_info_map["运行内存"] = "--";
                #print(phone_info_map["运行内存"]);
                for item in temp_p.items():
                    if "CPU型号" == item.attr("paramname"):
                        phone_info_map["CPU型号"] = item.text().lstrip().rstrip();                                         #phone_info_doc_two("div.cell-con-tit ul.cell_phone_param li").eq(3)("p").eq(2).text().split("；")[0].split("：")[1].lstrip().rstrip();
                        break;
                    else:
                        phone_info_map["CPU型号"] = "--";
                #print(phone_info_map["CPU型号"]);
            except Exception:
                phone_info_map["屏幕尺寸"] = "--";
                phone_info_map["屏幕分辨率"] = "--"; 
                phone_info_map["操作系统"] = "--";
                phone_info_map["运行内存"] = "--";
                phone_info_map["CPU型号"] = "--";
            phone_info_list.append(phone_info_map);
            #print(phone_info_list);
            #break;
        print("start write info to table!")
        print(phone_info_list);
        for item in phone_info_list:
            try:
                excel.write(sheet_name,i,0,item["手机名"]);
                excel.write(sheet_name,i,1,item["上市时间"]);
                excel.write(sheet_name,i,2,item["屏幕尺寸"]);
                excel.write(sheet_name,i,3,item["屏幕分辨率"]);
                excel.write(sheet_name,i,4,item["操作系统"]);
                excel.write(sheet_name,i,5,item["运行内存"]);
                excel.write(sheet_name,i,6,item["CPU型号"]);
            except Exception as ex:
                redirect_out = __Autonomy__();
                cur_stdout = sys.stdout;
                sys.stdout = redirect_out;
                print(ex);
                sys.stdout = cur_stdout;
                #print(redirect_out._buff);
                temp_ex = redirect_out._buff.split("'")[1];
                print(temp_ex);
                if "屏幕尺寸" == temp_ex:
                    #print(redirect_out._buff);
                    excel.write(sheet_name,i,2,"--");
                    excel.write(sheet_name,i,3,"--");
                    excel.write(sheet_name,i,4,"--");
                    excel.write(sheet_name,i,5,"--");
                    excel.write(sheet_name,i,6,"--");                
                else:
                    print("xxxxxxxxxxxxxxxxxxxxxxxxxxxx");
                if "屏幕分辨率" == ex:
                    excel.write(sheet_name,i,3,"--");
                    excel.write(sheet_name,i,4,"--");
                    excel.write(sheet_name,i,5,"--");
                    excel.write(sheet_name,i,6,"--");  
                if "操作系统" == ex:
                    excel.write(sheet_name,i,4,"--");
                    excel.write(sheet_name,i,5,"--");
                    excel.write(sheet_name,i,6,"--"); 
                if "运行内存" == ex:
                    excel.write(sheet_name,i,5,"--");
                    excel.write(sheet_name,i,6,"--"); 
                if "CPU型号" == ex:
                    excel.write(sheet_name,i,6,"--"); 
            excel.save(file_name);
            i = i + 1;
            print(i);
        #break;
    
    
#    all_smart_phone_current_html = doc_smart_phone("div.all-con-con");
#    print(all_smart_phone_current_html);
#    all_smart_phone_current_html = "https://product.cnmo.com/all/product_t1_p165.html#allConShaix";
#    r_test = requests.get(all_smart_phone_current_html);
#    r_test_pq = pq(r_test.text);
#    a = r_test_pq("div .all-con-con .all-con-shaix div .page-up a .pnext").attr("href");
#    if not a:
#        print(a);
except Exception as e:
    print(e);


