# 如要读取自定义的invoice.xlsx内容，请删除img_ins文件夹中的所有文件，并删除invoice_auto.xlsx文件
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ChromeOptions
import os
import time
from datetime import datetime
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from image_util import image_crop_invoice_paper, weixin_screenshots_to_texts, weixin_texts_to_elements, weixin_elements_to_excel
from docx import Document


class InvoiceChecker:
    target_url = 'https://inv-veri.chinatax.gov.cn/'
    invoice_xlsx = 'invoice.xlsx'
    invoice_auto_xlsx = 'invoice_auto.xlsx'
    invoice_word = 'invoice.docx'

    weixin_screenshots_src = os.getcwd() + '\\img_ins\\'

    raw_path = os.getcwd() + '\\screenshot_raw\\'
    os.makedirs(raw_path, exist_ok=True)
    crop_path = os.getcwd() + '\\screenshot_crop\\'
    os.makedirs(crop_path, exist_ok=True)
    ratio_y0, ratio_y1, ratio_x0, ratio_x1 = 0.05, 0.75, 0.21, 0.79

    def __init__(self, max_workers=3):
        self.max_workers = max_workers
        self.browser_type = None

    @classmethod
    def start(cls, max_workers=3):
        object_ins = cls(max_workers=max_workers)

        while object_ins.browser_type not in ['C', 'F', 'E']:
            browser_type = input('\n\n============== 请选择 ==============\n使用Chrome浏览器【C】操作\n使用FireFox浏览器【F】操作\n退出程序【E】Exit\n\n请输入：')
            object_ins.browser_type = browser_type.upper()
            print('====================================\n输入的内容是：%s' % object_ins.browser_type)

        if object_ins.browser_type == 'E':
            return

        object_ins._start()
        return object_ins

    def _start(self):
        start_time = datetime.now()
        print('==========================（1）汇总微信截图 ==================================')
        list_text = weixin_screenshots_to_texts(self.weixin_screenshots_src)
        list_obj = weixin_texts_to_elements(list_text)
        weixin_elements_to_excel(list_obj)
        one_time = datetime.now()
        print(f"解图用时: {one_time - start_time}")
        print(f"解图完成: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        print('==========================（2）开始爬取网站 ==================================')
        self.start_master()
        two_time = datetime.now()
        print(f"爬取用时: {two_time - one_time}")
        print(f"爬取完成: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        print('==========================（3）开始裁剪图片 ==================================')
        image_crop_invoice_paper(self.raw_path, self.crop_path, self.ratio_y0, self.ratio_y1, self.ratio_x0, self.ratio_x1)
        three_time = datetime.now()
        print(f"裁剪用时: {three_time - two_time}")
        print(f"裁剪完成: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        #print('==========================（4）制作WORD文件 ==================================')
        #document = Document()
        #list_pic = os.listdir(self.crop_path)
        #for pic in list_pic:
        #    pic_fullname = os.path.join(self.crop_path, pic)
        #    document.add_picture(pic_fullname)
        #document.save(self.invoice_word)

    def start_master(self):
        invoice_ins = self.process_check_list()
        if invoice_ins:
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                results = executor.map(self.handle_request, invoice_ins)

                for result in results:
                    print('完成result为：', result)
        else:
            print('已没有新的发票需要下载网站截图，或者请在img_ins文件夹里放入微信扫描发票的结果的截图，或者在根目录下手工添加规定格式的invoice.xlsx文件')

    def process_check_list(self):
        list_file = os.listdir(os.getcwd())
        if self.invoice_auto_xlsx in list_file:
            df = pd.read_excel(self.invoice_auto_xlsx, sheet_name=0, converters={'发票代码': str, '发票号码': str, '开票日期': str, '校验码后六位': str})
        elif self.invoice_xlsx in list_file:
            df = pd.read_excel(self.invoice_xlsx, sheet_name=0, converters={'发票代码': str, '发票号码': str, '开票日期': str, '校验码后六位': str})
        else:
            return None

        df.rename(columns={'发票代码': 'fpdm', '发票号码': 'fphm', '开票日期': 'kprq', '校验码后六位': 'kjje'}, inplace=True)
        check_list = []
        for i in df.index.values:
            row_data = df.loc[i, ['fpdm', 'fphm', 'kprq', 'kjje']].to_dict()
            check_list.append(row_data)

        pic_set = set()
        for root, dirs, files in os.walk(self.raw_path):
            for name in files:
                name = os.path.splitext(name)[0]
                pic_set.add(name)

        list_need = []
        for one in check_list:
            filename = one['fpdm'] + '-' + one['fphm']
            if filename not in pic_set:
                list_need.append(one)

        return list_need

    def handle_request(self, invoice_in):
        fpdm = invoice_in['fpdm']
        fphm = invoice_in['fphm']
        kprq = invoice_in['kprq']
        kjje = invoice_in['kjje']

        options = ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-automation'])

        if self.browser_type == 'F':
            browser = webdriver.Firefox()
        else:
            browser = webdriver.Chrome(options=options)

        try:
            browser.get(self.target_url)
            wait = WebDriverWait(browser, 3600)
            time.sleep(1)
            browser.maximize_window()

            browser.find_element_by_css_selector('#fpdm').send_keys(fpdm)
            browser.find_element_by_css_selector('#fphm').send_keys(fphm)
            browser.find_element_by_css_selector('#kprq').send_keys(kprq)
            element_kjje = browser.find_element_by_css_selector('#kjje')
            if element_kjje:
                element_kjje.send_keys(kjje)

            if self.browser_type == 'F':
                wait.until(EC.presence_of_element_located((By.ID, 'print_area')))  # 使用FireFox时需要留意
            else:
                wait.until(EC.presence_of_element_located((By.ID, 'dialog-body')))

            time.sleep(2)
        except BaseException as pic_msg:
            print("截图失败：%s" % pic_msg)
        finally:
            if self.browser_type == 'C':
                browser.switch_to.frame('dialog-body')

            response_content = browser.page_source
            filename = fpdm + '-' + fphm
            browser.save_screenshot(self.raw_path + filename + '.png')

            time.sleep(2)
            browser.close()

            with open(self.raw_path + filename + '.html', 'w', encoding='utf-8') as f:
                f.write(response_content)

        return filename


if __name__ == "__main__":
    InvoiceChecker.start()
