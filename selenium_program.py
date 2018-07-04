from selenium import webdriver
import time
import json
from openpyxl import Workbook
import requests

class Sele():
    def __init__(self, ident, series_id, series_name):
        self.driver = webdriver.PhantomJS()
        self.driver.maximize_window()
        self.ident = ident
        self.series_id = series_id
        self.series_name = series_name
        self.comms = []

    def open_url(self):
        url = "http://db.auto.sohu.com/%s/%s/dianping.html" % (self.ident, self.series_id)
        print(url)
        self.driver.get(url)
    def get_tagA(self):
        koubei_box = self.driver.find_element_by_class_name("koubei-box")
        tab = koubei_box.find_element_by_class_name("tab")
        tagA = tab.find_elements_by_tag_name("a")
        return tagA
    def click_a(self, tagA):
        for a in tagA:
            a.click()
            time.sleep(1)
            self.comms.append(self.get_current_comment())
    def get_current_comment(self):
        part = self.driver.find_element_by_css_selector(".koubei-tabcon.cur")
        items = part.find_elements_by_class_name("short-comm")
        part_comm = []
        for item in items:
            content = item.text
            if content not in part_comm:
                part_comm.append(content)
        return part_comm
        # flag = self.check_get_more(part)
        # self.get_more(part ,part_comm, flag)
    def check_get_more(self, part):
        get_more = part.find_element_by_class_name("get-more")
        return "display: none" in get_more.get_attribute("style")

    def get_more(self, part, part_comm, flag):
        if not flag:
            part.find_element_by_class_name("get-more").click()
            items = part.find_elements_by_class_name("short-comm")
            for item in items[len(part_comm)+1:]:
                content = item.text
                if not self.check_exist(content):
                    part_comm.append(content)
                else:
                    break
            else:
                self.get_more(part, part_comm, flag)
            return part_comm
        else:
            return part_comm
    def check_exist(self, content):
        for comm in self.comms:
            return content in comm
    def close_driver(self):
        self.driver.close()
def get_comms(ident,series_id, series_name):
    sele = Sele(ident,series_id,series_name)
    r = requests.get("http://db.auto.sohu.com/%s/%s/dianping.html"%(ident, series_id), allow_redirects = False)
    if 200 != r.status_code:
        return None
    else:
        flag = sele.open_url()
        tagA = sele.get_tagA()
        sele.click_a(tagA)
        sele.close_driver()
        return sele.comms

def main():
    wb = Workbook()
    f = open("seriesinfoall.json","r",encoding="utf-8")
    seriesinfo = json.load(f)
    for each in seriesinfo:
        ident = each["ident"]
        table_name = each["name"]
        for series in each["series"]:
            series_id = series["d"]
            series_name = series["n"]
            title = table_name+"-"+series_name
            comms = get_comms(ident, series_id, series_name)
            if comms is None:
                continue
            else:
                save_excel(wb, title, comms)
    wb.save("souhu.xlsx")
def save_excel(wb, title, comms):
    sheet = wb.create_sheet(title=title)
    sheet["A1"] = "短口碑"
    sheet["B1"] = "最满意"
    sheet["C1"] = "最不满意"
    sheet["D1"] = "外观"
    sheet["E1"] = "内饰"
    sheet["F1"] = "空间"
    sheet["G1"] = "动力"
    sheet["H1"] = "操控"
    sheet["I1"] = "油耗"
    sheet["J1"] = "舒适性"
    sheet["K1"] = "性价比"
    for i in range(len(comms)):
        comm = comms[i]
        for j in range(len(comm)):
            sheet["%s%d"%(chr(65+i),j+2)] = comm[j]



if __name__ == '__main__':
    main()
