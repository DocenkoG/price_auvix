# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import sys
import configparser
import time
import shutil
import openpyxl                      # Для .xlsx
#import xlrd                          # для .xls
from   price_tools import getCellXlsx, getCell, quoted, dump_cell, currencyType, openX, sheetByName
import csv
import requests, lxml.html



def getXlsString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item]
        if item in ('закупка','продажа','цена со скидкой','цена_') :
            if getCell(row=i, col=j, isDigit='N', sheet=sh) == '' :       # .find('Звоните') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCell(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCell(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def getXlsxString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item]
        #if item in ('закупка','продажа','цена','цена1') :
        if item.find('цена') >= 0 :
            if getCellXlsx(row=i, col=j, isDigit='N', sheet=sh).find('Звоните') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCellXlsx(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCellXlsx(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def convert_csv2csv( cfg ):
    inFfileName  = cfg.get('basic', 'filename_in')
    outFfileNameRUR = cfg.get('basic', 'filename_out_RUR')
    outFfileNameEUR = cfg.get('basic', 'filename_out_EUR')
    outFfileNameUSD = cfg.get('basic', 'filename_out_USD')
    inFile  = open( inFfileName,  'r', newline='', encoding='CP1251', errors='replace')
    outFileRUR = open( outFfileNameRUR, 'w', newline='')
    outFileEUR = open( outFfileNameEUR, 'w', newline='')
    outFileUSD = open( outFfileNameUSD, 'w', newline='')
    
    outFields = cfg.options('cols_out')
    csvReader = csv.DictReader(inFile, delimiter=';')
    csvWriterRUR = csv.DictWriter(outFileRUR, fieldnames=cfg.options('cols_out'))
    csvWriterEUR = csv.DictWriter(outFileEUR, fieldnames=cfg.options('cols_out'))
    csvWriterUSD = csv.DictWriter(outFileUSD, fieldnames=cfg.options('cols_out'))

    print(csvReader.fieldnames)
    csvWriterRUR.writeheader()
    csvWriterEUR.writeheader()
    csvWriterUSD.writeheader()
    recOut = {}
    for recIn in csvReader:
        for outColName in outFields :
            shablon = cfg.get('cols_out',outColName)
            for key in csvReader.fieldnames:
                if shablon.find(key) >= 0 :
                    shablon = shablon.replace(key, recIn[key])
            if outColName in('закупка','продажа'):
                if shablon.find('Звоните') >=0 :
                    shablon = '0.1'
            recOut[outColName] = shablon
        if recOut['валюта'] == 'Рубль' :
            csvWriterRUR.writerow(recOut)
        elif recOut['валюта'] == 'Евро' :
            csvWriterEUR.writerow(recOut)
        elif recOut['валюта'] == 'USD' :
            csvWriterUSD.writerow(recOut)
        else :
            log.error('нераспознана валюта "%s" для товара "%s"', recOut['валюта'], recOut['код производителя'] )
    log.info('Обработано '+ str(csvReader.line_num) +'строк.')
    inFile.close()
    outFileRUR.close()
    outFileEUR.close()
    outFileUSD.close()



def config_read( cfgFName ):
    cfg = configparser.ConfigParser(inline_comment_prefixes=('#'))
    if  os.path.exists('private.cfg'):     
        cfg.read('private.cfg', encoding='utf-8')
    if  os.path.exists(cfgFName):     
        cfg.read( cfgFName, encoding='utf-8')
    else: 
        log.debug('Нет файла конфигурации '+cfgFName)
    return cfg



def download( cfg ):
    retCode     = False
    filename_new= cfg.get('download','filename_new')
    filename_old= cfg.get('download','filename_old')
    login       = cfg.get('download','login'    )
    password    = cfg.get('download','password' )
    url_lk      = cfg.get('download','url_lk'   )
    url_file    = cfg.get('download','url_file' )
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.0; rv:14.0) Gecko/20100101 Firefox/14.0.1',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Language':'ru-ru,ru;q=0.8,en-us;q=0.5,en;q=0.3',
               'Accept-Encoding':'gzip, deflate',
               'Connection':'keep-alive',
               'DNT':'1'
              }

    try:
        s = requests.Session()
        r = s.get(url_lk,  headers = headers)  # auth=(login,password),(И без него сработало, но где-то может понадобиться)
        page = lxml.html.fromstring(r.text)
        form = page.forms[0]
        #print(form)
        form.fields['USER_LOGIN'] = login
        form.fields['USER_PASSWORD'] = password
        r = s.post(url_lk+ form.action, data=form.form_values())
        #print('<<<',r.text,'>>>')
        print('       ==================================================')

        log.debug('Авторизация на %s   --- code=%d', url_lk, r.status_code)
        r = s.get(url_file)
        log.debug('Загрузка файла %16d bytes   --- code=%d', len(r.content), r.status_code)
        retCode = True
    except Exception as e:
        log.debug('Exception: <' + str(e) + '>')

    if os.path.exists( filename_new) and os.path.exists( filename_old): 
        os.remove( filename_old)
        os.rename( filename_new, filename_old)
    if os.path.exists( filename_new) :
        os.rename( filename_new, filename_old)
    f2 = open(filename_new, 'wb')                                  # Теперь записываем файл
    f2.write(r.content)
    f2.close()
    if filename_new[-4:] == '.zip':                                # Архив. Обработка не завершена
        log.debug( 'Zip-архив. Разархивируем.')
        #dir_befo_download = set(os.listdir(os.getcwd()))
        os.system('unzip -oj ' + filename_new)
        #dir_afte_download = set(os.listdir(os.getcwd()))
        #new_files = list( dir_afte_download.difference(dir_befo_download))
    return retCode



def is_file_fresh(fileName, qty_days):
    qty_seconds = qty_days *24*60*60 
    if os.path.exists( fileName):
        price_datetime = os.path.getmtime(fileName)
    else:
        log.error('Не найден файл  '+ fileName)
        return False

    if price_datetime+qty_seconds < time.time() :
        file_age = round((time.time()-price_datetime)/24/60/60)
        log.error('Файл "'+fileName+'" устарел!  Допустимый период '+ str(qty_days)+' дней, а ему ' + str(file_age) )
        return False
    else:
        return True



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def main(dealerName):
    """ Обработка прайсов выполняется согласно файлов конфигурации.
    Для этого в текущей папке должны быть файлы конфигурации, описывающие
    свойства файла и правила обработки. По одному конфигу на каждый
    прайс или раздел прайса со своими правилами обработки
    """
    make_loger()
    log.info('          ' + dealerName)

    rc_download = False
    '''
    '''
    if os.path.exists('getting.cfg'):
        cfg = config_read('getting.cfg')
        filename_new_1 = cfg.get('basic','filename_new_1')
        filename_new_2 = cfg.get('basic','filename_new_2')
        if cfg.has_section('download'):
            rc_download = download(cfg)
        if not(rc_download==True or is_file_fresh( filename_new_1, int(cfg.get('basic','срок годности')))):
            return False
    for cfgFName in os.listdir("."):
        if cfgFName.startswith("cfg") and cfgFName.endswith(".cfg"):
            log.info('----------------------- Processing '+cfgFName )
            cfg = config_read(cfgFName)
            filename_in = cfg.get('basic','filename_in')
            if rc_download==True or is_file_fresh( filename_in, int(cfg.get('basic','срок годности'))):
                convert_excel2csv(cfg)


if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main( myName)
