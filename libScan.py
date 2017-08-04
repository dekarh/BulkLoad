# -*- coding: utf-8 -*-

#
# Библиотека функций
#

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from configparser import ConfigParser

from mysql.connector import MySQLConnection, Error
import datetime
from datetime import datetime
import time
import string
import sys
from random import uniform

# DRIVER_PATH = 'drivers/chromedriver.exe'
#DRIVER_PATH = 'drivers/chromedriver'

B = {
    'menuIII'   : {'t': 'i', 's': 'compact-mode-btn'},
    'select_in' : {'t': 'x', 's': '//A[@class="changeBoxContainer"][text()="Вход"]'},
    'login'     : {'t': 'x', 's': '//INPUT[@id="inputLogin"]'},
    'password'  : {'t': 'x', 's': '//INPUT[@id="inputPassword"]'},
    'a-button'  : {'t': 'x', 's': '//INPUT[@name="submit_login"]'},
    'tiles'     : {'t': 'x', 's': '//DIV[@class="tiles-list-wrapper"]//A[@href]'},
    'tiles-name': {'t': 'x', 's': '//DIV[@class="tiles-list-wrapper"]//A[@href]', 'a': 'text'},
    'tiles-href': {'t': 'x', 's': '//DIV[@class="tiles-list-wrapper"]//A[@href]', 'a': 'href'},
    'tiles-img' : {'t': 'x', 's': '//DIV[@class="tiles-list-wrapper"]//IMG[@src]', 'a': 'src'},
    'tiles-offl': {'t': 'x', 's': '//I[@class="icn icn-circleOnline-small icn-small info-online js-info-online"]'
                                  '/..', 'a': 'href'},
    'tiles-onln': {'t': 'x', 's': '//I[@class="icn icn-circleOnline-small icn-small info-online js-info-online show"]'
                                  '/..', 'a': 'href'},
  'anketa-about': {'t': 'x', 's': '//DIV[@class="b-anketa_inset b-anketa_inset-info"]', 'a': 'outerHTML'},
   'anketa-msg' : {'t': 'x', 's': '//DIV[@class="b-profile-cloud-inner__message alien"]', 'a': 'text'},
 'anketa-favour': {'t': 'x', 's': '//DIV[@class="in clearFix"]', 'a': 'text'},
    'anketa-btn': {'t': 'x', 's': '//A[@class="button button-blue first  _openChateg "]', 'a': 'href'},
'anketa-locator': {'t': 'x', 's': '//SPAN[@class="info info-misc__distance"]', 'a': 'text'},
    'back-find' : {'t': 'x', 's': '//A[@class="widget-title js-widget-title"][text()="Результаты поиска"]'},



    'menuRegs'  : {'t': 'x', 's': '(//SPAN[@class="controls-DropdownList__text"])[1]'},
    'menuCats'  : {'t': 'x', 's': '(//SPAN[@class="controls-DropdownList__text"])[2]'},
    'firms_x'   : {'t': 'x', 's': '//DIV[@class="Contragents-CommonRenders__InnCorner '
                                                    'Contragents-CommonRenders__Inn ws-ellipsis"]'},
    'firms_tr'  : {'t': 'x', 's': '//DIV[@sbisname="contragentsBrowser"]//TR[@data-id]'},
    'firms_trA' : {'t': 'x', 's': '//SPAN[@class="Contragents-CommonRenders__Name"]/../../../../..', 'a': 'data-id'},
    'data_id'   : {'t': 'x', 's': '//TR[@class="controls-DataGridView__tr controls-ListView__item '
                                                          'js-controls-ListView__item"][@data-id="'},
    'data_idA'  : {'t': 'x', 's': '//TR[@class="controls-DataGridView__tr controls-ListView__item '
                                    'js-controls-ListView__item"][@data-id="', 'a' : 'data-id'},
    'close'     : {'t': 'x', 's': '//DIV[@class="sbisname-window-title-close ws-button-classic ws-component '
                        'ws-control-inactive ws-enabled ws-field-button ws-float-close-right ws-no-select"]'},
    'first'     : {'t': 'x', 's': '(//I[@sbisname="PagingBegin"])[1]'},
    'next'      : {'t': 'x', 's': '(//I[@sbisname="PagingNext"])[1]'},
    'prev'      : {'t': 'x', 's': '(//I[@sbisname="PagingPrev"])[1]'},
    'innA'      : {'t': 'x', 's': '//INPUT[@name="СтрокаИНН"]', 'a' : 'value'},
    'kppA'      : {'t': 'x', 's': '//INPUT[@name="СтрокаКПП"]', 'a': 'value'},
    'familyA'   : {'t': 'x', 's': '//INPUT[@name="СтрокаФамилия"]', 'a': 'value'},
    'nameA'     : {'t': 'x', 's': '//INPUT[@name="СтрокаИмя"]', 'a': 'value'},
    'surnameA'  : {'t': 'x', 's': '//INPUT[@name="СтрокаОтчество"]', 'a': 'value'},
'firm_full_nameA':{'t': 'x', 's': '//INPUT[@name="СтрокаПолноеНазвание"]', 'a': 'value'},
   'act_num1000': {'t': 'x', 's': '//DIV[@class="custom-select-option"][@value="1000"]'},
    'about'     : {'t': 'x', 's': '//SPAN[@class="ContragentCard_RightAccordion-content"][text()="О компании"]'},
    'contacts'  : {'t': 'x', 's': '//SPAN[@class="ContragentCard_RightAccordion-content"][text()="Контактные данные"]'},
    'rekv'      : {'t': 'x', 's': '//SPAN[@class="ContragentCard_RightAccordion-content"][text()="Реквизиты"]'},
    'owners'    : {'t': 'x', 's': '//SPAN[@class="ContragentCard_RightAccordion-content"][text()="Владельцы"]'},
    'summA'     : {'t': 'x', 's': '//SPAN[@class="Contragents-ContragentCardRatingBanner__title-Revenue  '
                                  'ctrg-subseparator"][text()="Выручка: "]/SPAN', 'a': 'text'},
    'costA'     : {'t': 'x', 's': '//SPAN[@class="Contragents-ContragentCardRatingBanner__title-Cost  '
                                  'ctrg-subseparator"][text()="Стоимость бизнеса: "]/SPAN', 'a': 'text'},
    'rat_sumA'  : {'t': 'x', 's': '//DIV[@class="ctrg-half-left"]//DIV[@class="Contragents-'
                                  'ContragentCardRatingBanner__positions"]/DIV/SPAN', 'a': 'title'},
    'rat_costA' : {'t': 'x', 's': '//DIV[@class="ctrg-half-right"]//DIV[@class="Contragents-'
                                  'ContragentCardRatingBanner__positions"]/DIV/SPAN', 'a': 'title'},
    'phonesA'   : {'t': 'x', 's': '//DIV[@sbisname="Таблица телефонов"]//DIV[@class="crm-phone-number crm-noicon '
                                  'ContragentCardPhones-Ellipsis"]', 'a': 'text'},
   'phones_typA': {'t': 'x', 's': '//DIV[@sbisname="Таблица телефонов"]//SPAN[@class="crm-phone-comment '
                                  'ContragentCardPhones-Ellipsis"]', 'a': 'text'},
         'warnA': {'t': 'x', 's': '//DIV[@sbisname="Реестры"]//DIV[@class="Contragents-ContragentCardIndicators'
                                  '__itemTitle "]//SPAN[@data-component="SBIS3.CONTROLS.Link"]', 'a': 'text'},
    'warn_dataA': {'t': 'x', 's': '//DIV[@sbisname="Реестры"]//DIV[@class="Contragents-ContragentCardIndicators_'
                                  '_listItemFooter"]//SPAN[@data-component="SBIS3.CONTROLS.Link"]', 'a': 'text'},
    'filialsA'  : {'t': 'x', 's': '//DIV[@sbisname="Таблица филиалов"]//DIV[@title]', 'a': 'text'},
    'ogrnA'     : {'t': 'x', 's': '//DIV[@sbisname="СтрокаОГРН"]//SPAN[@title]', 'a': 'text'},
    'okpoA'     : {'t': 'x', 's': '//DIV[@sbisname="СтрокаОКПО"]//SPAN[@title]', 'a': 'text'},
    'oktmoA'    : {'t': 'x', 's': '//DIV[@sbisname="СтрокаОКТМО"]//SPAN[@title]', 'a': 'text'},
    'reg_N_pfrA': {'t': 'x', 's': '//DIV[@sbisname="СтрокаРегНомерПФ"]//SPAN[@title]', 'a': 'text'},
    'reg_comp'  : {'t': 'x', 's': '//DIV[@class="ContragentCardRegistration-Data"]', 'a': 'text'},
    'reg_org'   : {'t': 'x', 's': '//DIV[@class="ContragentCardRegistrationGosOrg-Data"]', 'a': 'text'},
    'uchred'    : {'t': 'x', 's': '//DIV[@sbisname="brwУчредители"]//DIV[@class="ws-browser-cell-paddings"]'
                                  '/DIV[@title]', 'a': 'title'},
    'uchredsA'  : {'t': 'x', 's': '//DIV[@sbisname="brwУчредители"]//TR[@rowkey]', 'a': 'rowkey'},
 'uchred_innAD' : {'t': 'x', 's': '//DIV[@sbisname="brwУчредители"]//TR[@rowkey="', 'a': 'text',
                                  'e' :'//SPAN[@class="Contragents-ContragentCardPersons__Inn"]'},
   'uchred_%AD' : {'t': 'x', 's': '//DIV[@sbisname="brwУчредители"]//TR[@rowkey="', 'a': 'text',
                                  'e': '//DIV[@class="Contragents-ContragentCardPersons__cellSharePercentage"]'},
'uchred_nameAD' : {'t': 'x', 's': '//DIV[@sbisname="brwУчредители"]//TR[@rowkey="', 'a': 'text',
                                  'e': '//DIV[@class="ws-ellipsis"]/SPAN[not(@class)]'},
    'dochki'    : {'t': 'x', 's': '//DIV[@sbisname="brwДочерниеКомпании"]//DIV[@class="ws-browser-cell-paddings"]'
                                  '/DIV[@title]', 'a': 'title'},
    'inn_spisA' : {'t': 'x', 's': '//DIV[@sbisname="contragentsBrowser"]//DIV[@class="Contragents-CommonRenders_'
                                  '_InnCorner Contragents-CommonRenders__Inn ws-ellipsis"]', 'a': 'text'},
    'search'    : {'t': 'x', 's': '//DIV[@sbisname="strSearch"]//INPUT'},
    'search_reg': {'t': 'x', 's': '//INPUT[@placeholder="Код или название региона"]'},
 'cats_all_link': {'t': 'x', 's': '//DIV[@class="controls-DropdownList__buttonsBlock"]//SPAN[text()="Еще..."]'},
    'okved-tab' : {'t': 'x', 's': '//DIV[@data-component="SBIS3.CONTROLS.TabButton"][@data-id="ОКВЭД"]'},
    'sbis-tab'  : {'t': 'x', 's': '//DIV[@data-component="SBIS3.CONTROLS.TabButton"][@data-id="Категории"]'},
   'okved-listA': {'t': 'x', 's': '//DIV[@sbisname="okvedSelector"]//TR[@data-id]//DIV[@title]', 'a': 'title'},
   'okved-listD': {'t': 'x', 's': '//DIV[@sbisname="okvedSelector"]//TR[@data-id]//DIV[@title="'},
    'sbis-listA': {'t': 'x', 's': '//DIV[@sbisname="vdSelector"]//TR[@data-id]//DIV[text()]', 'a': 'text'},
    'sbis-listD': {'t': 'x', 's': '//DIV[@sbisname="vdSelector"]//TR[@data-id]//DIV[text()="'},
    'more_uchr' : {'t': 'x', 's': '//DIV[@sbisname="btnЕщеУчредители"]'},

    'cats'      : {'t': 'c', 's': 'controls-DropdownList__item-text'},
    'firms_c'   : {'t': 'c', 's': 'controls-DataGridView__tr'},
   'ch_surnameA': {'t': 'c', 's': 'Contragents-ContragentCard__Chief__surname', 'a': 'text'},
    'ch_nameA'  : {'t': 'c', 's': 'Contragents-ContragentCard__Chief__name', 'a': 'text'},
    'ch_titleA' : {'t': 'c', 's': 'Contragents-ContragentCard__Chief__title', 'a': 'text'},
    'gen_infoA' : {'t': 'c', 's': 'Contragents-ContragentCardGeneralInfo__State', 'a': 'text'},
    'act_link'  : {'t': 'c', 's': 'Contragents-ContragentCardGeneralInfo__ActivityTypes__title'},
  'act_by_count': {'t': 'c', 's': 'custom-select-text'},
    'acts'      : {'t': 'c', 's': 'ws-browser-table-row'},
    'act_numA'  : {'t': 'c', 's': 'Contragents-ContragentCardGeneralInfo__ActivityTypes__counter', 'a': 'text'},
    'emp_qtyA'  : {'t': 'c', 's': 'Contragents-ContragentCard__EmployeesQuantity__qty', 'a': 'text'},
    'addressA'  : {'t': 'c', 's': 'ContragentCardAddresses-blackLink', 'a': 'text'},
    'predstavA' : {'t': 'c', 's': 'user-info-cell', 'a': 'text'},
    'reg_filter': {'t': 'c', 's': 'Clrs__regionSelectorItemName'},

}

propusk = ['-', 'aoзт', 'а', 'автономн%', 'администрац%', 'ао', 'ассоциац%', 'астрахан%', 'агентств%',
           'база', 'бутик', 'бюро', 'в', 'г', 'групп', 'группа', 'дом', 'дополнительн%', 'завод',
           'и', 'или', 'институ%', 'информацион%', 'ип',
           'кафе', 'кировск%', 'клиника', 'комбинат', 'комитет', 'коммерческ%', 'компан%', 'комплекс', 'кондитерск%',
           'консалтин%'
           'ленинск%', 'ломбард', 'магази%', 'мастерск%', 'миграц%', 'мир', 'многопрофильн%', 'муниципальн%',
           'на', 'научно-исследовательск%', 'некоммерческ%',
           'оао', 'област%', 'образован%', 'обществен%', 'общество', 'объединен%', 'ограниченной', 'оказан%', 'ооо',
           'организац%', 'от', 'ответственностью', 'отделен%', 'охранн%',
           'пао', 'плюс', 'по', 'портал', 'под', 'помощ%', 'поселок', 'предприят%', 'производствен%', 'профессиональн%',
           'рабочий', 'ремонтн%', 'российск%', 'рф',
           'с', 'сеть', 'сервиc', 'служб%', 'советск%', 'строительн%', 'субфилиал', 'специализирован%',
           'терминал', 'торгов%', 'транспортн%', 'трусовск%',
           'управление', 'управляющ%', 'участок', 'фабрик%', 'фармацевтическ%', 'филиал', 'фирм%', 'частн%',
           'центр', 'цех', 'школ%', 'юридическ%'
           ]
LINK = [
    'Не начинал',
    'Не интересен',
    'Её интерес',
    'Переписка',
    'Симпатия',
    'Ждет встречу',
    'Встреча',
    'Доверие',
]

PEOPLE = [
    'Упырь',
    'Продажа',
    'Барыга',
    'Недалекая',
    'Нет КПД',
    'Не верит',
    'Нет места',
    'Неизвестная',
    'Услышала',
    'Проводник',
]

ONLINE = [
    'OFFLINE',
    'ONLINE',
    'Все равно'
]

def unique(lst):
    seen = set()
    j = 0
    while j < len(lst)-1:
        for i, x in enumerate(lst):
            j = i
            if x.lower() in seen:
                lst.pop(i)
                seen = set()
                break
            seen.add(x.lower())
    return

def wr():
    time.sleep(uniform(0.5, 5))

def read_config(filename='config.ini', section='mysql'):
    """ Read database configuration file and return a dictionary object
    :param filename: name of the configuration file
    :param section: section of database configuration
    :return: a dictionary of database parameters
    """
    # create parser and read ini configuration file
    parser = ConfigParser()
    parser.read(filename)

    # get section, default to mysql
    db = {}
    if parser.has_section(section):
        items = parser.items(section)
        for item in items:
            db[item[0]] = item[1]
    else:
        raise Exception('{0} not found in the {1} file'.format(section, filename))

    return db


def l(a):
    try:
        if a != None:
            a = str(a).strip()
            if  a != '':
                a = ''.join([char for char in a if char in string.digits])
                if len(a) > 0:
                    return int(a)
                else:
                    return 0
        return 0
    except TypeError:
        return 0

def s(a):
    try:
        if a != None:
            return str(a).strip().replace(u"\xa0", u" ")
        return ''
    except TypeError:
        return ''

def append_words(a, n_words):
    try:
        if a != None:
            for a1 in a.split(','):
                for a2 in a1.split(';'):
                    for a3 in a2.split(' '):
                        if a3 != '':
                            n_words.append(a3)
            unique(n_words)
            unique(n_words)
        return
    except TypeError:
        return

def norm_phone(tel):
    tel = str(tel).strip()
    if tel == '' or tel == None:
        return None
    else:
        tel = ''.join([char for char in tel if char in string.digits])
        if len(tel) == 11:
            if tel[0] in ['8', '9']:
                return int('7' + tel[1:])
            elif tel[0] == '7':
                return int(tel)
            else:
                return None
        elif len(tel) == 10:
            return int('7' + tel)
        elif len(tel) == 6:
            return int('78512' + tel)
        else:
            return None


def wj(driver):  # Ждем, пока динамическая ява завершит все свои процессы
    WebDriverWait(driver, 50).until(lambda driver: driver.execute_script("return jQuery.active == 0"))
    """
    Еще варианты фреймворков/библиотек:
    "return Ajax.activeRequestCount == 0"
    "return dojo.io.XMLHTTPTransport.inFlight.length == 0"
    Ожидание пока все набранные буквы отработют явой:
    element = WebDriverWait(ff, 10).until(EC.presence_of_element_located((By.ID, "keywordSuggestion")))
    """
    return

def wa(driver): # Типа ловит анимацию. Здесь не ловит :(
    WebDriverWait(driver, 10).until(lambda driver: driver.find_element(By.ID, 'new - element') and
                                                   driver.find_elements(By.ID, 'spinner') == 0)
    return
"""
Прокрутка
driver.execute_script("return arguments[0].scrollIntoView();", elem) # Здесь не точно
driver.execute_script("window.scrollTo(0, 911)") # Здесь вообще не прокручивает

"""
def chk(d, t, s, f = '', a = ''): # Проверка наличия элемента, не вызывающая исключения
    wj(d)
    try:
        if   t == 'i':
            d.find_element(By.ID, s)
        elif t == 'c':
            d.find_element(By.CLASS_NAME, s)
        elif t == 'x':
            d.find_element(By.XPATH, s)
    except NoSuchElementException:
        return False
    return True
"""
^^^
|||
Потому что EC.presence_of_element_located((By.XPATH, "xpath"))) возвращает объект, не нашел где там результат
try:
    assert EC.presence_of_element_located((By.XPATH, '//*[@id="Waldo"]')) is not True
except AssertionError, e:
    self.verificationErrors.append('presence_of_element_located returned True for Waldo')
"""

def p(d, t, f, s, a = '', data_id = '', e = ''):
    wj(d)
    if data_id != '':
        data_id += '"]'
    if t == 'i':
        if   f == 'c':
            foo = WebDriverWait(d, 20).until(EC.element_to_be_clickable((By.ID, s)))
            wj(d)
            if a == '':
                return foo
            else:
                if a == 'text':
                    return foo.text
                else:
                    return foo.get_attribute(a)
        elif f == 'v':
            foo = WebDriverWait(d, 20).until(EC.visibility_of_element_located((By.ID, s)))
            wj(d)
            if a == '':
                return foo
            else:
                if a == 'text':
                    return foo.text
                else:
                    return foo.get_attribute(a)
        elif f == 'vs':
            foo = WebDriverWait(d, 20).until(EC.visibility_of_any_elements_located((By.ID, s)))
            wj(d)
            if a == '':
                return foo
            else:
                if a == 'text':
                    return [atr.text for atr in foo]
                else:
                    return [atr.get_attribute(a) for atr in foo]
        elif f == 'vv':
            if chk(d = d, t = t, s = s):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.ID, s)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return [atr.text for atr in foo if atr.is_displayed()]
                    else:
                        return [atr.get_attribute(a) for atr in foo if atr.is_displayed()]
            else:
                if a == '':
                    return []
                else:
                    return ['']

        elif f == 'p':
            if chk(d = d, t = t, s = s):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_element_located((By.ID, s)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return foo.text
                    else:
                        return foo.get_attribute(a)
            else:
                if a == '':
                    return
                else:
                    return ''
        elif f == 'ps':
            if chk(d = d, t = t, s = s):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.ID, s)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return [atr.text for atr in foo]
                    else:
                        return [atr.get_attribute(a) for atr in foo]
            else:
                if a == '':
                    return []
                else:
                    return ['']
        else:
            return
    elif t == 'x':
        if   f == 'c':
            foo = WebDriverWait(d, 20).until(EC.element_to_be_clickable((By.XPATH, s + data_id + e)))
            wj(d)
            if a == '':
                return foo
            else:
                if a == 'text':
                    return foo.text
                else:
                    return foo.get_attribute(a)
        elif f == 'v':
            foo = WebDriverWait(d, 20).until(EC.visibility_of_element_located((By.XPATH, s + data_id + e)))
            wj(d)
            if a == '':
                return foo
            else:
                if a == 'text':
                    return foo.text
                else:
                    return foo.get_attribute(a)
        elif f == 'vs':
            foo = WebDriverWait(d, 20).until(EC.visibility_of_any_elements_located((By.XPATH, s + data_id + e)))
            wj(d)
            if a == '':
                return foo
            else:
                if a == 'text':
                    return [atr.text for atr in foo]
                else:
                    return [atr.get_attribute(a) for atr in foo]
        elif f == 'vv':
            if chk(d = d, t = t, s = s + data_id + e):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.XPATH, s + data_id + e)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return [atr.text for atr in foo if atr.is_displayed()]
                    else:
                        return [atr.get_attribute(a) for atr in foo if atr.is_displayed()]
            else:
                if a == '':
                    return []
                else:
                    return ['']
        elif f == 'p':
            if chk(d = d, t = t, s = s + data_id + e):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_element_located((By.XPATH, s + data_id + e)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return foo.text
                    else:
                        return foo.get_attribute(a)
            else:
                if a == '':
                    return
                else:
                    return ''
        elif f == 'ps':
            if chk(d = d, t = t, s = s + data_id + e):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.XPATH, s + data_id + e)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return [atr.text for atr in foo]
                    else:
                        return [atr.get_attribute(a) for atr in foo]
            else:
                if a == '':
                    return []
                else:
                    return ['']
        else:
            return
    elif t == 'c':
        if   f == 'c':
            foo = WebDriverWait(d, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, s)))
            if a == '':
                return foo
            else:
                if a == 'text':
                    return foo.text
                else:
                    return foo.get_attribute(a)
        elif f == 'v':
            foo = WebDriverWait(d, 20).until(EC.visibility_of_element_located((By.CLASS_NAME, s)))
            if a == '':
                return foo
            else:
                if a == 'text':
                    return foo.text
                else:
                    return foo.get_attribute(a)
        elif f == 'vs':
            foo = WebDriverWait(d, 20).until(EC.visibility_of_any_elements_located((By.CLASS_NAME, s)))
            wj(d)
            if a == '':
                return foo
            else:
                if a == 'text':
                    return [atr.text for atr in foo]
                else:
                    return [atr.get_attribute(a) for atr in foo]
        elif f == 'vv':
            if chk(d = d, t = t, s = s):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.CLASS_NAME, s)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return [atr.text for atr in foo if atr.is_displayed()]
                    else:
                        return [atr.get_attribute(a) for atr in foo if atr.is_displayed()]
            else:
                if a == '':
                    return []
                else:
                    return ['']
        elif f == 'p':
            if chk(d = d, t = t, s = s):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_element_located((By.CLASS_NAME, s)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return foo.text
                    else:
                        return foo.get_attribute(a)
            else:
                if a == '':
                    return
                else:
                    return ''
        elif f == 'ps':
            if chk(d = d, t = t, s = s):
                wj(d)
                foo = WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.CLASS_NAME, s)))
                wj(d)
                if a == '':
                    return foo
                else:
                    if a == 'text':
                        return [atr.text for atr in foo]
                    else:
                        return [atr.get_attribute(a) for atr in foo]
            else:
                if a == '':
                    return []
                else:
                    return ['']
        else:
            return

def authorize(driver, login, password, authorize_page=''):
    time.sleep(1)
    if authorize_page != '':
        driver.get(authorize_page)
    time.sleep(1)
    sel = p(d = driver, f = 'c', **B['select_in'])
    sel.click()
    # Ввод логина
    log = p(d = driver, f = 'c', **B['login'])
    time.sleep(1)
    log.send_keys(login)
    # Ввод пароля
    passwd = p(d = driver, f = 'c', **B['password'])
    time.sleep(1)
    passwd.send_keys(password)
    # Отправка формы нажатием кнопки
    cl = p(d = driver, f = 'c', **B['a-button'])
    cl.click()
    return

def to_spisok(driver):
    err_count = 0
    g = 0
    while g < 1000:
        try:
            menu = p(d = driver, f = 'c', **B['menuIII']) # Три палочки
            wj(driver)
            company = p(d = driver, f = 'p', **B['menu>>'])  # >>
            wj(driver)
            menu.click()
            wj(driver)
            if not company.is_displayed():
                wj(driver)
                continue
            company.click()
            wj(driver)
            if chk(d = driver, **B['menuCats']):
                wj(driver)
                if p(d = driver, f = 'p', **B['menuCats']).is_displayed():
                    return
            continue
        except Exception as ee:
            print(datetime.strftime(datetime.now(), "%H:%M:%S"), 'Ошибка в to_spisok', ee)
            driver.close()
            sys.exit()
            continue

def set_filter(driver, use_category = 'True', type_category = 'СБИС', category = 'Страхование, пенсионное обеспечение',
               use_region = 'False', region = '30'):
    g = 0
    err_count = 0
    while g < 1000:
        try:
            if use_region == 'True' and int(region) > 0 and int(region) < 100:
                drop = p(d=driver, f='c', **B['menuRegs'])  # Открываем дроплист
                wj(driver)
                drop.click()
                wj(driver)
                cats_all_link = p(d=driver, f='vs', **B['cats_all_link'])  # Переходим ко всем категориям
                cats_all_link[0].click()
                wj(driver)
                time.sleep(4)
                search = p(d = driver, f = 'c', **B['search_reg'])
                wj(driver)
                search.clear()
                wj(driver)
                search.send_keys(region.strip())
                wj(driver)
                time.sleep(2)
                reg_filter = p(d = driver, f = 'c', **B['reg_filter'])
                wj(driver)
                reg_filter.click()
                wj(driver)
            else:
                print(datetime.strftime(datetime.now(), "%H:%M:%S"), " Регион не найден\n\nНЕ МОГУ ПРИМЕНИТЬ ФИЛЬТР !!!")
                return
            if use_category == 'True':
                drop = p(d = driver, f = 'c', **B['menuCats']) # Открываем дроплист
                wj(driver)
                drop.click()
                wj(driver)
                cats_all_link = p(d = driver, f = 'vs', **B['cats_all_link']) # Переходим ко всем категориям
                cats_all_link[0].click()
                wj(driver)
                time.sleep(4)
                if type_category == 'ОКВЭД':
                    category = category.strip() + ' '
                    okved_tab = p(d = driver, f = 'c', **B['okved-tab'])
                    wj(driver)
                    okved_tab.click()
                    search = p(d = driver, f = 'c', **B['search'])
                    wj(driver)
                    search.clear()
                    wj(driver)
                    search.send_keys(category.strip())
                    wj(driver)
                    time.sleep(2)
                    okved_list = p(d = driver, f = 'ps', **B['okved-listA'])
                    wj(driver)
                    for okved_str in okved_list:
                        wj(driver)
                        if okved_str[:(len(category))] == category:
                            okved = p(d = driver, f = 'c', **B['okved-listD'], data_id=okved_str)
                            okved.click()
                            wj(driver)
                            time.sleep(4)
                            return
                elif type_category == 'СБИС':
                    sbis_tab = p(d = driver, f = 'c', **B['sbis-tab'])
                    wj(driver)
                    sbis_tab.click()
                    search = p(d = driver, f = 'c', **B['search'])
                    wj(driver)
                    search.clear()
                    wj(driver)
                    search.send_keys(category.strip())
                    wj(driver)
                    time.sleep(2)
                    sbis_list = p(d = driver, f = 'vs', **B['sbis-listA'])
                    wj(driver)
                    for sbis_str in sbis_list:
                        if sbis_str.strip() == category.strip():
                            sbis = p(d = driver, f = 'c', **B['sbis-listD'], data_id=sbis_str.strip())
                            sbis.click()
                            wj(driver)
                            time.sleep(4)
                            return
                else:
                    print(datetime.strftime(datetime.now(), "%H:%M:%S")," Категория (ОКВЭД или СБИС) не найдена,\n\n"
                                                                        "НЕ МОГУ ПРИМЕНИТЬ ФИЛЬТР !!!")
                    return
            else:
                return
        except Exception as ee:
            print(datetime.strftime(datetime.now(), "%H:%M:%S"), ' Ошибка в set_filter:\n', ee)
            driver.close()
            sys.exit()
            continue
