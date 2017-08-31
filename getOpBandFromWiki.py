#!/usr/bin/env python
# -*- coding: utf-8 -*-

import traceback
import os
import sys
from urllib import request
import xlrd
import xlwt
import re
import time

import codecs
import configparser
import logging


#read config : zh_name, timezone, bands of all RATs
config = configparser.ConfigParser()
path = 'config.conf'
config.readfp(codecs.open(path, "r", "utf-8-sig"))

#output logs
localtime = time.asctime( time.localtime(time.time()) )
filehandler = logging.FileHandler(filename='./log--'+localtime.replace(':','.')+'.txt',encoding="utf-8")
fmter = logging.Formatter(fmt='%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s - %(message)s',datefmt="%Y-%m-%d %H:%M:%S")
filehandler.setFormatter(fmter)
loger = logging.getLogger(__name__)
loger.addHandler(filehandler)
loger.setLevel(logging.DEBUG)


class Operator_band_info():
    '''operator infos of country'''
    operator_name = ''
    mcc = ''
    mnc = ''
    gwtc_band = ''
    lte_band = ''

class Country_band_info():
    '''country info and operators band info'''
    country_name = ''
    zh_name = ''
    timezone = ''
    operators = []

class Lte_band_info():
    '''for decoding lte band info '''
    operator_name = ''
    lte_band = ''


def translate_country_name(en_name):
    '''translate en_name 2 zh_name and timezone'''
    try:
        zh_name = config.get('country', en_name)
        timezone = config.get('timezone', en_name)
    except Exception:
        loger.error('translate Country or Timezone error: ' + en_name)
        zh_name = en_name
        timezone = '?'
    
    return zh_name,timezone


def split_countries_decode_lte_band_data(country_band_infos_gwtc):
    '''decode lte band data'''
    #read origin data from wiki
    print("split_countries_decode_lte_info: request lte band info from wiki")
    lte_band_info_other = request.urlopen("https://en.wikipedia.org/wiki/List_of_LTE_networks").read().decode('utf8')
    lte_band_info_europe = request.urlopen("https://en.wikipedia.org/wiki/List_of_LTE_networks_in_Europe").read().decode('utf8')
    lte_band_info_asia = request.urlopen("https://en.wikipedia.org/wiki/List_of_LTE_networks_in_Asia").read().decode('utf8')
    lte_band_info_asia = lte_band_info_asia.replace("<br />\n(SAR)",'').replace(' (SAR)','')  #a <br />\n(SAR) following Hong Kong ; a  (SAR) following Macau
    lte_band_infos = lte_band_info_other + lte_band_info_europe + lte_band_info_asia
    print("response lte band info from wiki")
    
    #all operator items
    all_operator_band_struc = r"<tr>\n<td[\s\S]*?</td>\n</tr>"
    #band items
    band_item_struct = r"<td[\s\S]*?</td>"
    
    #filter operator items
    pattern = re.compile(all_operator_band_struc)
    all_operator_band_info = pattern.findall(lte_band_infos)
    
    #circle of operator items
    i=0
    while i < len(all_operator_band_info):
        #filter first operator of some country
        if(all_operator_band_info[i].find("span class=\"flagicon\"") < 0):
            loger.warning('non-first operator info: ' + all_operator_band_info[i])
            i=i+1
        else:
            #all operator info of some country
            lte_bands = []
            #get country name : delete splits bottom then get the last split item
            country_name = all_operator_band_info[i].split("</a></td>")[0].split(">")[-1]
            
            #filter items(operator_name, frequency, band...) without HTML Tags
            items = []
            pattern_item = re.compile(band_item_struct)
            for item in pattern_item.findall(all_operator_band_info[i]):
                dr = re.compile(r'<[^>]+>',re.S)
                dd = dr.sub('',item)
                items.append(dd)
            
            #add first operator info and its country info
            i=i+1
            if len(items) == 11:  #some countries do not contain an opetaror info, eg.Russia
                if items[2].find('♠') >=0:  #filter band info in format: 7003180000000000000♠</span>1800
                    items[2] = items[2].split("♠")[1]
                lte_band_info = Lte_band_info()
                lte_band_info.operator_name = items[1].replace('\n', '')
                lte_band_info.lte_band = "B"+items[3].split('\n')[0]
                lte_bands.append(lte_band_info)
                loger.info('process country :' + country_name + '  first operator lte band info: ' + lte_band_info.operator_name+'  '+lte_band_info.lte_band)
            
            #other operator infos basides the first operator info(do not contain country_name line)
            if(all_operator_band_info[i-1].find("td rowspan=") >= 0):
                #get operators' NO.
                count = int(all_operator_band_info[i-1].split("rowspan=\"")[1].split("\"")[0])
                for j in range(2, count+1):  #count+1 non-included in range()
                    #filter items(operator_name, frequency, band...) without HTML Tags
                    items = []
                    for item in pattern_item.findall(all_operator_band_info[i]):
                        dr = re.compile(r'<[^>]+>',re.S)
                        dd = dr.sub('',item)
                        items.append(dd)
                        
                    #add/append other operator info and its country info
                    i = i+1
                    if len(items) == 10:
                        if items[1].find('♠') >=0:
                            items[1] = items[1].split("♠")[1]
                        
                        #append operator LTE band info
                        add_flag = True
                        for k in range(0,len(lte_bands)):  #following out of index - 1
                            if lte_bands[k].operator_name == items[0].replace('\n', ''):
                                lte_bands[k].lte_band = lte_bands[k].lte_band + ",B"+items[2].split('\n')[0]
                                add_flag = False
                                loger.info('process country :' + country_name + 'append operator lte band info: ' + lte_bands[k].operator_name+'  '+lte_bands[k].lte_band)
                        #add new operator LTE band info
                        if add_flag:
                            lte_band_info = Lte_band_info()
                            lte_band_info.operator_name = items[0].replace('\n', '')
                            lte_band_info.lte_band = "B"+items[2].split('\n')[0]
                            lte_bands.append(lte_band_info)
                            loger.info('process country :' + country_name + 'add operator lte band info: ' + lte_band_info.operator_name+'  '+lte_band_info.lte_band)
            
            #append to existing countries or add new country
            insert_flag = True
            #index the index of country 
            country_index = 0
            for country_info in country_band_infos_gwtc:  # 定位国家位置
                if country_name == country_info.country_name or \
                    (country_name == 'United States' and country_info.country_name == 'United States of America') or \
                    (country_name == 'Russia' and country_info.country_name == 'Russian Federation'): #in these countries: name different bw GWTC and LTE band page
                    insert_flag = False
                    break;
                else:
                    country_index = country_index + 1
                    
            #insert new country info
            if insert_flag :
                operators = []
                for lte_band_info in lte_bands:
                    operator = Operator_band_info()
                    operator.operator_name = lte_band_info.operator_name
                    operator.lte_band = lte_band_info.lte_band
                    operators.append(operator)
                country_info = Country_band_info()
                country_info.country_name = country_name
                country_info.zh_name,country_info.timezone = translate_country_name(country_info.country_name)
                country_info.operators = operators
                country_band_infos_gwtc.append(country_info)
                loger.info('process band_list add lte info : insert country band info: ' + country_name + str(len(country_info.operators)) + 'operators')
            #append LTE info to existing country
            else:
                for lte_band_info in lte_bands:
                    #append to existing operators
                    add_flag = True
                    for j in range(0,len(country_band_infos_gwtc[country_index].operators)): # if following out of index : - 1
                        loger.info(country_band_infos_gwtc[country_index].operators[j].operator_name+'-----------'+lte_band_info.operator_name.split("\n")[0]+ \
                            '  find result: '+str(country_band_infos_gwtc[country_index].operators[j].operator_name.find(lte_band_info.operator_name.split("\n")[0])))
                        # some LTE country names are different with gwtc country names
                        if country_band_infos_gwtc[country_index].operators[j].operator_name.lower().find(lte_band_info.operator_name.split(" (")[0].lower()) >= 0 or \
                            country_band_infos_gwtc[country_index].operators[j].operator_name.lower().find(lte_band_info.operator_name.split(" /")[0].lower()) >= 0 or \
                            country_band_infos_gwtc[country_index].operators[j].operator_name.lower().find(lte_band_info.operator_name.split("(")[0].lower()) >= 0 :
                            add_flag = False
                            country_band_infos_gwtc[country_index].operators[j].lte_band = country_band_infos_gwtc[country_index].operators[j].lte_band+','+lte_band_info.lte_band
                            loger.info('process band_list add lte info : set operator band info: ' + country_name + '  operator: ' + \
                                country_band_infos_gwtc[country_index].operators[j].operator_name+'  '+lte_band_info.lte_band)
                            break;
                    #add new operators
                    if add_flag:
                        operator_lte_band = Operator_band_info()
                        operator_lte_band.operator_name = lte_band_info.operator_name
                        operator_lte_band.lte_band = lte_band_info.lte_band
                        country_band_infos_gwtc[country_index].operators.append(operator_lte_band)
                        loger.info('process band_list add lte info : add operator band info: ' + country_name + '  operator: ' + operator_lte_band.operator_name)
                        
    print(len(country_band_infos_gwtc), 'countries after add lte band info')
    
    return country_band_infos_gwtc
    

def split_countries_decode_gwtc_info():
    '''split 23G band infos of wiki'''
    print("split_countries_decode_gwtc_info: request https://en.wikipedia.org/wiki/Mobile_country_code")
    gwtc_band_infos = request.urlopen("https://en.wikipedia.org/wiki/Mobile_country_code").read().decode('utf8')
    print("response https://en.wikipedia.org/wiki/Mobile_country_code")
    
    country_band_infos_gwtc = []
    
    coyntry_band_struc = r"<h3>[\s\S]*?<table[\s\S]*?</table>"
    no_band_info_country = r"<h3>[\s\S]*?</h3>\n<p>See [\s\S]*?</p>"
    pattern = re.compile(coyntry_band_struc)  #正则表达式查找
    
    print('decode gwtc infos')
    for item in pattern.findall(gwtc_band_infos):
        band_info = re.sub(no_band_info_country,"",item)
        
        band_data = Country_band_info() #国家英文, 运营商名, MCC, MNC, GSM850, 900, 1800, 1900, WCDMA 10个, TDS 2个, CDMA 3个, CDMA2000 6个, 
        
        operator_struct = r"<tr>\n<td>[\s\S]*?</tr>"
        operator_item_struct = r"<td>[\s\S]*?</td>"
    
        #country_name
        country_name = band_info.split('</a>')[0].split('<a href=')[1].split('>')[1]
        band_data.country_name = country_name
        band_data.zh_name,band_data.timezone = translate_country_name(country_name)
    
        #Operator, MCC, MNC  --  in operational
        operator_infos = []
        pattern = re.compile(operator_struct)
        for operator_item in pattern.findall(band_info):
            infos = []
            pattern1 = re.compile(operator_item_struct)
            for item in pattern1.findall(operator_item):
                dr = re.compile(r'<[^>]+>',re.S)
                dd = dr.sub('',item)
                infos.append(dd) #append:  mcc  mnc  brand  operator  Status  Bands (MHz)  References and notes
            loger.info('process gwtc band info : ' + country_name + 'operator: ' + infos[3]+' ('+infos[2]+')' + ' gwtc band: ' + infos[5])
            
            add_flag = True
            for operator_item in operator_infos:
                if(operator_item.operator_name == infos[3]+' ('+infos[2]+')' and infos[4] == 'Operational'): # add " and infos[4] == 'Operational'"
                    add_flag = False
                    if(operator_item.mcc.find(infos[0]) < 0):  #Japan has 2 mcc
                        operator_item.mcc = operator_item.mcc + ',' + infos[0]
                    if(operator_item.mnc.find(infos[1]) < 0):
                        operator_item.mnc = operator_item.mnc + ',' + infos[1]
                    operator_item.gwtc_band = operator_item.gwtc_band + infos[5]
                    loger.info('process gwtc band info : ' + country_name + 'append operator: ' + infos[3]+' ('+infos[2]+')' + ' gwtc band: ' + infos[5])
                    break
                
            if(infos[4] == 'Operational' and add_flag):
                operator_info = Operator_band_info()
                operator_info.operator_name = infos[3]+' ('+infos[2]+')'
                operator_info.mcc = infos[0]
                operator_info.mnc = infos[1]
                operator_info.gwtc_band = infos[5]
                operator_infos.append(operator_info)
                loger.info('process gwtc band info : ' + country_name + 'add operator: ' + infos[3]+' ('+infos[2]+')' + ' gwtc band: ' + infos[5])
    
        band_data.operators = operator_infos
        
        #at least one operator's info needed to append country_info
        if (len(band_data.operators) >= 1):
            country_band_infos_gwtc.append(band_data)
    
    return country_band_infos_gwtc


def set_xlwt_style(color):
    stylei= xlwt.XFStyle() 
    
    #set color
    patterni= xlwt.Pattern() 
    patterni.pattern=1 
    patterni.pattern_fore_colour=color  
    #color = 0 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, the list goes on...
    patterni.pattern_back_colour=35 
    
    #alignment
    alignment = xlwt.Alignment() # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_LEFT # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    
    #border
    borders = xlwt.Borders() # Create Borders
    borders.left = xlwt.Borders.THIN # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED......
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    stylei.pattern=patterni #
    stylei.alignment = alignment # Add Alignment to Style
    stylei.borders = borders # Add Borders to Style
    return stylei


def write_rows(sheet, country_band_data, row, col, style):
    '''write band info to rows'''
    #read bands of all RATs from config file
    gwt_bands = [
    config.get('band', 'gsm').split(","),
    config.get('band', 'wcdma').split(","),
    config.get('band', 'tds').split(",")
    ]
    cdma_bands = config.get('band', 'cdma').split(",")
    c2k_bands = config.get('band', 'cdma2000').split(",")
    lte_bands = config.get('band', 'lte').split(",")
    
    #write merged country_name, zh_name, timezone to excel
    sheet.write_merge(row, row + col - 1, 0, 0, country_band_data.country_name, style)
    sheet.write_merge(row, row + col - 1, 1, 1, country_band_data.zh_name, style)
    sheet.write_merge(row, row + col - 1, 2, 2, country_band_data.timezone, style)
    loger.info('write file : ' + country_band_data.country_name + country_band_data.zh_name)
    
    #circle for operators in country_band_data
    for i in range(0, len(country_band_data.operators)):
        #write operator_name, mcc. mnc
        sheet.write(row+i, 3, country_band_data.operators[i].operator_name.replace('&amp;','&'), style)
        sheet.write(row+i, 4, country_band_data.operators[i].mcc, style)
        sheet.write(row+i, 5, country_band_data.operators[i].mnc.replace('&#160;',''), style)
        loger.info('write file : operator: ' + country_band_data.operators[i].operator_name + 'gwtc_band: ' + \
            country_band_data.operators[i].gwtc_band + 'lte_band: ' + country_band_data.operators[i].lte_band)
        
        #write band infos
        count_col = 5
        for rat_bands in gwt_bands:  #write gwt bands
            for band in rat_bands:
                count_col = count_col + 1
                if country_band_data.operators[i].gwtc_band.replace('GSM900', 'GSM 900').find(band) >= 0:
                    sheet.write(row+i, count_col, band.replace('TD-SCDMA','TDS'), style)
                else:
                    sheet.write(row+i, count_col, "", style)
        for band in cdma_bands:  #write cdma bands
            count_col = count_col + 1
            #filter CDMA info without WCDMA info  &&  if CDMA2000 supported, CDMA also set to supported.
            if (country_band_data.operators[i].gwtc_band.find(band) >= 0 and country_band_data.operators[i].gwtc_band.find('TD-S'+band) < 0) or \
                    country_band_data.operators[i].gwtc_band.find(band.replace('CDMA','CDMA2000')) >= 0:
                sheet.write(row+i, count_col, band, style)
            else:
                sheet.write(row+i, count_col, "", style)
        for band in c2k_bands:  #write cdma2000 bands
            count_col = count_col + 1
            if country_band_data.operators[i].gwtc_band.find(band) >= 0 :
                sheet.write(row+i, count_col, band.replace('CDMA2000','C2K'), style)
            else:
                sheet.write(row+i, count_col, "", style)
        for band in lte_bands:  #write lte bands
            count_col = count_col + 1
            if band in country_band_data.operators[i].lte_band.split(','):  # a in str_b means 'ab in abc,ad' is true; a in array_b means 'ab in ['abc','ad']' is false.
                sheet.write(row+i, count_col, band, style)
            else:
                sheet.write(row+i, count_col, "", style)
        loger.info('flag: operator: ' + country_band_data.operators[i].operator_name + 'gwtc_band: ' + \
            country_band_data.operators[i].gwtc_band + 'lte_band: ' + country_band_data.operators[i].lte_band)


def write_title(sheet):
    '''write title'''
    style_title = set_xlwt_style(29)
    sheet.write_merge(0, 1, 0, 1, '国家', style_title)
    sheet.write_merge(0, 1, 2, 2, '时区', style_title)
    sheet.write_merge(0, 1, 3, 3, '运营商', style_title)
    sheet.write_merge(0, 1, 4, 4, 'MCC', style_title)
    sheet.write_merge(0, 1, 5, 5, 'MNC', style_title)
    sheet.write_merge(0, 0, 6, 9, 'GSM', style_title)
    sheet.write_merge(0, 0, 10, 16, 'WCDMA', style_title)
    sheet.write_merge(0, 0, 17, 18, 'TDS', style_title)
    sheet.write_merge(0, 0, 19, 24, 'CDMA', style_title)
    sheet.write_merge(0, 0, 25, 32, 'CDMA2000', style_title)
    sheet.write_merge(0, 0, 33, 64, 'LTE', style_title)
    
    bands = ['850(B5)','900(B8)','1800(B3)','1900(B2)','800(B6)','850(B5)','900(B8)','1700(B4)','1800(B3)','1900(B2)','2100(B1)','1900(B39F)','2000(B34A)','450','800','850','1900',
        '2000','2100','450','800','850','1700','1800','1900','2000','2100','B38','B39','B40','B41','B42','B43','B1','B2','B3','B4','B5','B7','B8','B9','B11','B12','B13','B17','B18',
        'B19','B20','B21','B25','B26','B27','B28','B29','B30','B31','B32','B34','B66']
    index = 6
    for band in bands:
        sheet.write(1, index, band, style_title)
        index = index+1


def set_xlwt_width(sheet):
    '''set width'''
    sheet.col(0).width = 150*20
    sheet.col(1).width = 150*20
    sheet.col(2).width = 150*20
    sheet.col(3).width = 600*20
    sheet.col(4).width = 100*20
    sheet.col(5).width = 170*20
    for i in range(6, 68):
        sheet.col(i).width = 200*20


def process_band_info_list():
    '''write excel & set style'''
    # decode gwtc band info
    country_band_infos_gwtc = split_countries_decode_gwtc_info()
    print ('countries with gwtl band info',len(country_band_infos_gwtc))
    
    # decode lte band info
    country_band_infos_with_lte_band = split_countries_decode_lte_band_data(country_band_infos_gwtc)
    print ('countries with lte band info',len(country_band_infos_with_lte_band))
    
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(u'国家频段')
    #write title
    write_title(sheet)
    #write content
    index = 2
    style_alter = True
    styles = [set_xlwt_style(1),set_xlwt_style(3)]
    for country_band_data in country_band_infos_with_lte_band:
        height = len(country_band_data.operators)
        if height > 0:
            if style_alter:  # set color/border/alignment
                write_rows(sheet, country_band_data, index, height, styles[0])
                style_alter = False
            else:
                write_rows(sheet, country_band_data, index, height, styles[1])
                style_alter = True
            index = index + height
    #set width of content
    set_xlwt_width(sheet)
    
    workbook.save(u'运营商与对应频段.xls')


'''main'''
process_band_info_list()
print (os.system('svn commit -m \"每周自动更新上传\" 运营商与对应频段.xls'))
os.system("pause")


