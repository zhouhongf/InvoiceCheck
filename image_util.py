import os
import cv2
from fnmatch import fnmatch
import pytesseract as pt
from PIL import Image, ImageEnhance
import re
import pandas as pd


# 用于裁剪从发票查验网站上截屏过来的图片
# ratio_y0, ratio_y1, ratio_x0, ratio_x1 = 0.05, 0.75, 0.21, 0.79
def image_crop_invoice_paper(src_path, dest_path, ratio_y0, ratio_y1, ratio_x0, ratio_x1):
    # src_path = os.getcwd() + '\\invoice_pic\\'
    # dest_path = os.getcwd() + '\\invoice_pic_cropped\\'
    os.makedirs(dest_path, exist_ok=True)
    # print(os.listdir(src_path))

    for file in os.listdir(src_path):
        if fnmatch(file, '*.png') or fnmatch(file, '*.jpeg') or fnmatch(file, '*.jpg'):
            filename = os.path.join(src_path, file)
            img = cv2.imread(filename)

            img_shape = img.shape
            len_y = img_shape[0]
            len_x = img_shape[1]

            if 0 < ratio_y0 < 1:
                y0 = int(len_y * ratio_y0)
            else:
                y0 = 0

            if 0 < ratio_y1 < 1 and ratio_y0 < ratio_y1:
                y1 = int(len_y * ratio_y1)
            else:
                y1 = len_y

            if 0 < ratio_x0 < 1:
                x0 = int(len_x * ratio_x0)
            else:
                x0 = 0

            if 0 < ratio_x1 < 1 and ratio_x0 < ratio_x1:
                x1 = int(len_x * ratio_x1)
            else:
                x1 = len_x

            cropped = img[y0:y1, x0:x1]
            filename_new = os.path.join(dest_path, file)
            cv2.imwrite(filename_new, cropped)


# !!! 用于裁剪微信发票, 但是裁剪后内容识别不出来
def image_crop_by_ratio(src_path, dest_path, ratio_y0, ratio_y1, ratio_x0, ratio_x1):
    os.makedirs(dest_path, exist_ok=True)
    print('准备裁剪图片：', os.listdir(src_path))
    for file in os.listdir(src_path):
        if fnmatch(file, '*.png') or fnmatch(file, '*.jpeg') or fnmatch(file, '*.jpg'):
            filename = os.path.join(src_path, file)
            img = Image.open(filename)
            img_shape = img.size
            len_y = img_shape[1]
            len_x = img_shape[0]

            if 0 < ratio_y0 < 1:
                y0 = int(len_y * ratio_y0)
            else:
                y0 = 0

            if 0 < ratio_y1 < 1 and ratio_y0 < ratio_y1:
                y1 = int(len_y * ratio_y1)
            else:
                y1 = len_y

            if 0 < ratio_x0 < 1:
                x0 = int(len_x * ratio_x0)
            else:
                x0 = 0

            if 0 < ratio_x1 < 1 and ratio_x0 < ratio_x1:
                x1 = int(len_x * ratio_x1)
            else:
                x1 = len_x

            img_cropped = img.crop((x0, y0, x1, y1))
            filename_new = os.path.join(dest_path, file)
            img_cropped.save(filename_new)


# 用于从微信发票截图中读取信息
def weixin_screenshots_to_texts(src_path):
    list_text = []
    for file in os.listdir(src_path):
        if fnmatch(file, '*.png') or fnmatch(file, '*.jpeg') or fnmatch(file, '*.jpg'):
            filename = os.path.join(src_path, file)
            image = Image.open(filename)
            image = image.convert('L')
            text = pt.image_to_string(image=image, lang='chi_sim')
            list_text.append(text)
    return list_text


# 根据正则表达式，将微信图片中读取出来的文本，提取出关键信息
def weixin_texts_to_elements(list_text: list):
    pattern_fpdm = re.compile(r'代码\s+([0-9]+)')
    pattern_fphm = re.compile(r'号码\s+([0-9]+)')
    pattern_amount_notax = re.compile(r'金额\s+([0-9]+(\.[0-9]+)*)')
    pattern_kprq = re.compile(r'日期\s+(20[0-9]{2}\.[0-9]{1,2}\.[0-9]{1,2})')
    pattern_check_code = re.compile(r'验码\s+([0-9]+)')

    list_obj = []
    for text in list_text:
        obj = {'fpdm': '', 'fphm': '', 'amount_notax': '', 'kprq': '', 'kjje': ''}
        fpdm_text = pattern_fpdm.search(text)
        if fpdm_text:
            fpdm = fpdm_text.group(1)
            obj['fpdm'] = fpdm if fpdm else ''

        fphm_text = pattern_fphm.search(text)
        if fphm_text:
            fphm = fphm_text.group(1)
            obj['fphm'] = fphm if fphm else ''

        amount_notax_text = pattern_amount_notax.search(text)
        if amount_notax_text:
            amount_notax = amount_notax_text.group(1)
            obj['amount_notax'] = amount_notax if amount_notax else ''

        kprq_text = pattern_kprq.search(text)
        if kprq_text:
            kprq_temp = kprq_text.group(1)
            if kprq_temp:
                kprq = re.sub('\.', '', kprq_temp)
                obj['kprq'] = kprq

        kjje_text = pattern_check_code.search(text)
        if kjje_text:
            kjje = kjje_text.group(1)
            if kjje:
                obj['kjje'] = kjje[-6:]

        if obj['fpdm'] and obj['fphm'] and obj['amount_notax'] and obj['kprq'] and obj['kjje']:
            list_obj.append(obj)
        else:
            obj = weixin_text_to_elements_extra(text)
            list_obj.append(obj)

    return list_obj


def weixin_text_to_elements_extra(text: str):
    obj = {'fpdm': '', 'fphm': '', 'amount_notax': '', 'kprq': '', 'kjje': ''}

    pattern_kprq = re.compile(r'20[0-9]{2}\.[0-9]{1,2}\.[0-9]{1,2}')
    pattern_fpdm = re.compile(r'0[0-9]{11}')
    pattern_check_code = re.compile(r'[0-9]{20}')

    pattern_amount_notax = re.compile(r'[0-9]+\.[0-9]+')
    pattern_num = re.compile(r'[0-9]+')
    pattern_fphm = re.compile(r'[0-9]{8}')

    kprq_text = pattern_kprq.search(text)
    if kprq_text:
        kprq_temp = kprq_text.group(0)
        kprq = re.sub('\.', '', kprq_temp)
        obj['kprq'] = kprq
        text = re.sub(kprq_temp, '', text)

    fpdm_text = pattern_fpdm.search(text)
    if fpdm_text:
        fpdm = fpdm_text.group(0)
        obj['fpdm'] = fpdm
        text = re.sub(fpdm, '', text)

    kjje_text = pattern_check_code.search(text)
    if kjje_text:
        kjje = kjje_text.group(0)
        obj['kjje'] = kjje[-6:]
        text = re.sub(kjje, '', text)

    amount_notax_text = pattern_amount_notax.search(text)
    if amount_notax_text:
        amount_notax = amount_notax_text.group(0)
        obj['amount_notax'] = amount_notax
        text = re.sub(amount_notax, '', text)

        if obj['kprq'] and obj['fpdm'] and obj['kjje']:
            fphm_text = pattern_num.search(text)
            if fphm_text:
                fphm = fphm_text.group(0)
                obj['fphm'] = fphm
                return obj
    else:
        fphm_text = pattern_fphm.search(text)
        if fphm_text:
            fphm = fphm_text.group(0)
            obj['fphm'] = fphm
            text = re.sub(fphm, '', text)

            amount_notax_text = pattern_num.search(text)
            if amount_notax_text:
                amount_notax = amount_notax_text.group(0)
                obj['amount_notax'] = amount_notax

            return obj
    return obj


def weixin_elements_to_excel(list_obj, filename='invoice_auto.xlsx'):
    if list_obj:
        data = pd.DataFrame(list_obj)
        data_new = data.drop(columns='amount_notax')
        data_new.rename(columns={'fpdm': '发票代码', 'fphm': '发票号码', 'kprq': '开票日期', 'kjje': '校验码后六位'}, inplace=True)
        data_new.to_excel(filename, index=None)


# 1、对比度：白色画面(最亮时)下的亮度除以黑色画面(最暗时)下的亮度；
# 2、色彩饱和度：：彩度除以明度，指色彩的鲜艳程度，也称色彩的纯度；
# 3、色调：向负方向调节会显现红色，正方向调节则增加黄色。适合对肤色对象进行微调；
# 4、锐度：是反映图像平面清晰度和图像边缘锐利程度的一个指标。
def augument(image_path, parent):
    # 读取图片
    image = Image.open(image_path)

    image_name = os.path.split(image_path)[1]
    name = os.path.splitext(image_name)[0]

    # 变亮
    # 亮度增强,增强因子为0.0将产生黑色图像；为1.0将保持原始图像。
    image_brightened1 = ImageEnhance.Brightness(image).enhance(1.5)
    image_brightened1.save(os.path.join(parent, '{}_bri1.jpg'.format(name)))

    # 变暗
    image_brightened2 = ImageEnhance.Brightness(image).enhance(0.8)
    image_brightened2.save(os.path.join(parent, '{}_bri2.jpg'.format(name)))

    # 色度,增强因子为1.0是原始图像
    # 色度增强
    image_colored1 = ImageEnhance.Color(image).enhance(1.5)
    image_colored1.save(os.path.join(parent, '{}_col1.jpg'.format(name)))

    # 色度减弱
    image_colored1 = ImageEnhance.Color(image).enhance(0.8)
    image_colored1.save(os.path.join(parent, '{}_col2.jpg'.format(name)))

    # 对比度，增强因子为1.0是原始图片
    # 对比度增强
    image_contrasted1 = ImageEnhance.Contrast(image).enhance(1.5)
    image_contrasted1.save(os.path.join(parent, '{}_con1.jpg'.format(name)))

    # 对比度减弱
    image_contrasted2 = ImageEnhance.Contrast(image).enhance(0.8)
    image_contrasted2.save(os.path.join(parent, '{}_con2.jpg'.format(name)))

    # 锐度，增强因子为1.0是原始图片
    # 锐度增强
    image_sharped1 = ImageEnhance.Sharpness(image).enhance(3.0)
    image_sharped1.save(os.path.join(parent, '{}_sha1.jpg'.format(name)))

    # 锐度减弱
    image_sharped2 = ImageEnhance.Sharpness(image).enhance(0.8)
    image_sharped2.save(os.path.join(parent, '{}_sha2.jpg'.format(name)))


# 自适应阈值二值化
def _get_dynamic_binary_image(filedir, img_name):
    dest_path = os.getcwd() + '\\out_img\\'
    os.makedirs(dest_path, exist_ok=True)
    filename = dest_path + img_name.split('.')[0] + '-binary.jpg'

    img_name = filedir + '/' + img_name

    im = cv2.imread(img_name)
    im = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)

    th1 = cv2.adaptiveThreshold(im, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 21, 1)
    cv2.imwrite(filename, th1)
    return th1


# 去除边框
def clear_border(img, img_name):
    dest_path = os.getcwd() + '\\out_img\\'
    os.makedirs(dest_path, exist_ok=True)
    filename = dest_path + img_name.split('.')[0] + '-clearBorder.jpg'

    h, w = img.shape[:2]
    for y in range(0, w):
        for x in range(0, h):
            # if y ==0 or y == w -1 or y == w - 2:
            if y < 4 or y > w - 4:
                img[x, y] = 255
            # if x == 0 or x == h - 1 or x == h - 2:
            if x < 4 or x > h - 4:
                img[x, y] = 255

    cv2.imwrite(filename, img)
    return img


# 干扰线降噪
def interference_line(img, img_name):
    dest_path = os.getcwd() + '\\out_img\\'
    os.makedirs(dest_path, exist_ok=True)
    filename = dest_path + img_name.split('.')[0] + '-interferenceline.jpg'

    h, w = img.shape[:2]

    # ！！！opencv矩阵点是反的， img[1,2] 1:图片的高度，2：图片的宽度
    for y in range(1, w - 1):
        for x in range(1, h - 1):
            count = 0
            if img[x, y - 1] > 245:
                count = count + 1
            if img[x, y + 1] > 245:
                count = count + 1
            if img[x - 1, y] > 245:
                count = count + 1
            if img[x + 1, y] > 245:
                count = count + 1
            if count > 2:
                img[x, y] = 255
    cv2.imwrite(filename, img)
    return img


# 点降噪
# 9邻域框,以当前点为中心的田字框,黑点个数
def interference_point(img, img_name, x=0, y=0):
    dest_path = os.getcwd() + '\\out_img\\'
    os.makedirs(dest_path, exist_ok=True)
    filename = dest_path + img_name.split('.')[0] + '-interferencePoint.jpg'

    # 判断图片的长宽度下限
    cur_pixel = img[x, y]  # 当前像素点的值
    height, width = img.shape[:2]

    for y in range(0, width - 1):
        for x in range(0, height - 1):
            if y == 0:  # 第一行
                if x == 0:  # 左上顶点,4邻域
                    # 中心点旁边3个点
                    sum = int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y + 1])
                    if sum <= 2 * 245:
                        img[x, y] = 0
                elif x == height - 1:  # 右上顶点
                    sum = int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x - 1, y]) \
                          + int(img[x - 1, y + 1])
                    if sum <= 2 * 245:
                        img[x, y] = 0
                else:  # 最上非顶点,6邻域
                    sum = int(img[x - 1, y]) \
                          + int(img[x - 1, y + 1]) \
                          + int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y + 1])
                    if sum <= 3 * 245:
                        img[x, y] = 0
            elif y == width - 1:  # 最下面一行
                if x == 0:  # 左下顶点
                    # 中心点旁边3个点
                    sum = int(cur_pixel) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y - 1]) \
                          + int(img[x, y - 1])
                    if sum <= 2 * 245:
                        img[x, y] = 0
                elif x == height - 1:  # 右下顶点
                    sum = int(cur_pixel) \
                          + int(img[x, y - 1]) \
                          + int(img[x - 1, y]) \
                          + int(img[x - 1, y - 1])

                    if sum <= 2 * 245:
                        img[x, y] = 0
                else:  # 最下非顶点,6邻域
                    sum = int(cur_pixel) \
                          + int(img[x - 1, y]) \
                          + int(img[x + 1, y]) \
                          + int(img[x, y - 1]) \
                          + int(img[x - 1, y - 1]) \
                          + int(img[x + 1, y - 1])
                    if sum <= 3 * 245:
                        img[x, y] = 0
            else:  # y不在边界
                if x == 0:  # 左边非顶点
                    sum = int(img[x, y - 1]) \
                          + int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x + 1, y - 1]) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y + 1])

                    if sum <= 3 * 245:
                        img[x, y] = 0
                elif x == height - 1:  # 右边非顶点
                    sum = int(img[x, y - 1]) \
                          + int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x - 1, y - 1]) \
                          + int(img[x - 1, y]) \
                          + int(img[x - 1, y + 1])

                    if sum <= 3 * 245:
                        img[x, y] = 0
                else:  # 具备9领域条件的
                    sum = int(img[x - 1, y - 1]) \
                          + int(img[x - 1, y]) \
                          + int(img[x - 1, y + 1]) \
                          + int(img[x, y - 1]) \
                          + int(cur_pixel) \
                          + int(img[x, y + 1]) \
                          + int(img[x + 1, y - 1]) \
                          + int(img[x + 1, y]) \
                          + int(img[x + 1, y + 1])
                    if sum <= 4 * 245:
                        img[x, y] = 0
    cv2.imwrite(filename, img)
    return img
