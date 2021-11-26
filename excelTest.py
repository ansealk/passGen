# Генерация паспорта сервиса из групп Zabbix
# Андросов С.А.
# 26.11.2021 - Версия 0.1
#   - Генерация списка хостов
#   - Генерация списка элементов по инфраструктуре
#   - Генерация списка триггеров по инфраструктуре

# Настройка параметров
zUser = 'api'
zPass = 'api'
ZABBIX_SERVER = 'http://zabbix.ansealk.ru'
serviceName = 'Linux Servers'
groupPrefix = ''
passportPath = 'c:\\Users\\ansea\\PycharmProjects\\pasGen\\'


import xlsxwriter
from pyzabbix import ZabbixAPI

wsHosts             = 'Узлы сети'
wsInfraItems        = 'Инфраструктура - Элементы'
wsInfraTriggers     = 'Инфраструктура - Триггеры'
wsBMetricItems      = 'Бизнес-метрики - Элементы'
wsBMetricTriggers   = 'Бизнес-метрики - Триггеры'
wsAlerts            = 'Оповещения'
worksheets=[]

hostsHeader=[
    ['Имя узла сети', 30],
    ['Видимое имя', 30],
    ['Тип', 10],
    ['IP-адрес / DNS-имя', 20],
    ['Описание', 40],
    ['Присоединенные шаблоны', 40 ]
]

infraItemsHeader=[
    ['Имя узла сети', 30],
    ['Элемент данных', 30],
    ['Описание элемента данных', 30],
    ['Ключ', 20],
    ['Тип', 10],
    ['Интервал', 10 ]
]

infraTriggersHeader=[
    ['Имя узла сети', 30],
    ['Название триггера', 30],
    ['Описание триггера', 40],
    ['Важность', 20],
    ['Выражение', 40],
    ['Выражение восстановления', 40]
]

itemsType={
    0: 'Zabbix agent',
    2: 'Zabbix trapper',
    3: 'Simple check',
    5: 'Zabbix internal',
    7: 'Zabbix agent (active)',
    9: 'Web item',
    10: 'External check',
    11: 'Database monitor',
    12: 'IPMI agent',
    13: 'SSH agent',
    14: 'Telnet agent',
    15: 'Calculated',
    16: 'JMX agent',
    17: 'SNMP trap',
    18: 'Dependent item',
    19: 'HTTP agent',
    20: 'SNMP agent',
    21: 'Script'
}

triggerPriority={
    0: 'не классифицировано',
    1: 'информационный',
    2: 'предупреждение',
    3: 'средний',
    4: 'высокий',
    5: 'чрезвычайный'
}

zapi = ZabbixAPI(ZABBIX_SERVER)
zapi.login(zUser, zPass)

# Получает id группы сервиа
def getHostGroupId(zabbix, group):
    hostgroup = zapi.hostgroup.get(
        output='groupid',
        search = {'name': [groupPrefix + group]
        },
    )
    return hostgroup[0]['groupid']

# Возвращает имя хоста по его Id
def searchHostNameById(id, hostlist):
    for host in hostlist:
        if host['hostid'] == id:
            return host['name']
    return ''


# Создает список хостов по сервису (группе в Zabbix)
def createHostsList(zabbix, group):
    hosts = zapi.host.get(
        output=['host', 'name', 'description', 'interfaces', 'hostid'],
        selectInterfaces=['interfaceid','hostid','dns','port','type','main','ip','useip'],
        selectParentTemplates=['name'],
        seletcTags=True,
        sortfield='hostid',
        groupids=group,
        expandComment=True,
        expandDescription=True,
        expandExpression=True,
    )
    return hosts

# Создает список элементов по сервису (группе в Zabbix)
def createItemsList(zabbix,groupid):
    items = zapi.item.get(
        output=['name', 'type', 'hostid', 'key_', 'delay', 'description'],
        groupids=groupid,
        sortfield='name',
        expandComment=True,
        expandDescription=True,
        expandExpression=True,
    )
    return items

# Создает список триггеров по сервису (группе в Zabbix)
def createTriggersList(zabbix,groupid):
    triggers = zapi.trigger.get(
        output=['hostname', 'description', 'expression', 'recovery_expression', 'priority', 'comments'],
        groupids=groupid,
        sortfield='hostname',
        expandComment=True,
        expandDescription=True,
        expandExpression=True,
    )
    return triggers

# Создает файл паспорта и  вкладки
def createPassport(name):
    workbook = xlsxwriter.Workbook(passportPath + name + '.xlsx')
    worksheets.append(workbook.add_worksheet(wsHosts))
    worksheets.append(workbook.add_worksheet(wsInfraItems))
    worksheets.append(workbook.add_worksheet(wsInfraTriggers))
    worksheets.append(workbook.add_worksheet(wsBMetricItems))
    worksheets.append(workbook.add_worksheet(wsBMetricTriggers))
    worksheets.append(workbook.add_worksheet(wsAlerts))
    return workbook

# Создание заголовка таблицы
def createPageHeader(wb,page,headerList):
    row = 1
    col = 1
    page.set_row(row, 30)
    page.outline_settings(True, False, False, True)
    page.autofilter('B2:G2')
    for header in (headerList):
        page.write(row, col, header[0], headerFormat)
        page.set_column(col, col, header[1])
        col+=1

# Создает таблицу с данными хостов
def createHostsTable(wb,page,dataList):
    row = 2
    col = 1
    for host in dataList:
        page.write(row, col, host['interfaces'][0]['dns'], dataFormat)
        page.write(row, col+1, host['name'], dataFormat)
        page.write(row, col+2, 'ZBX', dataFormat)
        page.write(row, col+3, host['interfaces'][0]['ip'], dataFormat)
        page.write(row, col+4, host['description'], dataFormat)
        templatesString = ''
        for template in host['parentTemplates']:
                templatesString = templatesString  + template['name'] + ',\n '
        page.write(row, col + 5, templatesString[:-2], dataFormat)
        row+=1

# Подготавливает таблицу айтемов
def prepareItemsTable(dataList):
    list=[]
    for data in dataList:
        item = []
        item.append(data['hostid'])
        item.append(searchHostNameById(data['hostid'],hostsList))
        item.append(data['name'])
        item.append(data['description'])
        item.append(data['key_'])
        item.append(itemsType[int(data['type'])])
        item.append(data['delay'])
        list.append(item)

    list.sort()
    return list



# Создает таблицу с айтемами
def createItemsTable(wb,page,list):
    row = 2
    col = 1
    currentId=''
    for item in list:
        if currentId != item[0]:     # если начался новый хост
            page.write(row, col,item[1],hostFormat)
            page.write_blank(row, col+1,None ,hostFormat)
            page.write_blank(row, col+2,None ,hostFormat)
            page.write_blank(row, col+3,None ,hostFormat)
            page.write_blank(row, col+4,None ,hostFormat)
            page.write_blank(row, col+5,None ,hostFormat)
            row+=1
            currentId = item[0]
        else:
            page.write(row, col, item[1], dataFormat)
            page.write(row, col+1, item[2], dataFormat)
            page.write(row, col+2, item[3], dataFormat)
            page.write(row, col+3, item[4], dataFormat)
            page.write(row, col+4, item[5], dataFormat)
            page.write(row, col+5, item[6], dataFormat)
            page.set_row(row, None, None, {'level': 1, 'collapsed': True, 'hidden': True})
            row+=1


# Создает таблицу с триггерами
def createTriggersTable(wb,page,list):
    row = 2
    col = 1
    currentHost=''
    for trigger in list:
        if currentHost != trigger['hostname']:      # Если начался новый хост
            page.write(row, col,trigger['hostname'],hostFormat)
            page.write_blank(row, col+1,None ,hostFormat)
            page.write_blank(row, col+2,None ,hostFormat)
            page.write_blank(row, col+3,None ,hostFormat)
            page.write_blank(row, col+4,None ,hostFormat)
            page.write_blank(row, col+5,None ,hostFormat)
            row+=1
            currentHost = trigger['hostname']
        else:
            page.write(row, col, trigger['hostname'], dataFormat)
            page.write(row, col+1, trigger['description'], dataFormat)
            page.write(row, col+2, trigger['comments'], dataFormat)
            page.write(row, col+3, trigger['priority'], dataFormat)
            page.write(row, col+3, triggerPriority[int(trigger['priority'])], triggerFormat[int(trigger['priority'])])

            page.write(row, col+4, trigger['expression'], dataFormat)
            page.write(row, col+5, trigger['recovery_expression'], dataFormat)
            page.set_row(row, None, None, {'level': 1, 'collapsed': True, 'hidden': True})
            row += 1

def createTriggerFormatCell(wb):
    format = []
    formatItem = wb.add_format({'font_size': 9})
    formatItem.set_text_wrap()
    formatItem.set_align('left')
    formatItem.set_align('vtop')
    formatItem.set_border(1)
    formatItem.set_fg_color('#97AAB3')
    format.append(formatItem)
    formatItem = wb.add_format({'font_size': 9})
    formatItem.set_text_wrap()
    formatItem.set_align('left')
    formatItem.set_align('vtop')
    formatItem.set_border(1)
    formatItem.set_fg_color('#7499FF')
    format.append(formatItem)
    formatItem = wb.add_format({'font_size': 9})
    formatItem.set_text_wrap()
    formatItem.set_align('left')
    formatItem.set_align('vtop')
    formatItem.set_border(1)
    formatItem.set_fg_color('#FFC859')
    format.append(formatItem)
    formatItem = wb.add_format({'font_size': 9})
    formatItem.set_text_wrap()
    formatItem.set_align('left')
    formatItem.set_align('vtop')
    formatItem.set_border(1)
    formatItem.set_fg_color('#FFA059')
    format.append(formatItem)
    formatItem = wb.add_format({'font_size': 9})
    formatItem.set_text_wrap()
    formatItem.set_align('left')
    formatItem.set_align('vtop')
    formatItem.set_border(1)
    formatItem.set_fg_color('#E97659')
    format.append(formatItem)
    formatItem = wb.add_format({'font_size': 9})
    formatItem.set_text_wrap()
    formatItem.set_align('left')
    formatItem.set_align('vtop')
    formatItem.set_border(1)
    formatItem.set_fg_color('#E45959')
    format.append(formatItem)
    return format


workbook=createPassport(serviceName)

triggerFormat = createTriggerFormatCell(workbook)

headerFormat = workbook.add_format({'bold': True, 'font_size' : 11})
headerFormat.set_text_wrap()
headerFormat.set_align('center')
headerFormat.set_align('vcenter')
headerFormat.set_border(2)

dataFormat = workbook.add_format({'font_size' : 9})
dataFormat.set_text_wrap()
dataFormat.set_align('left')
dataFormat.set_align('vtop')
dataFormat.set_border(1)

hostFormat = workbook.add_format({'bold': True,'font_size' : 9,'fg_color': '#97AAB3'})
hostFormat.set_text_wrap()
hostFormat.set_align('left')
hostFormat.set_align('vtop')
hostFormat.set_border(1)

hostGroup= getHostGroupId(zapi,serviceName)
createPageHeader(workbook,worksheets[0],hostsHeader)
createPageHeader(workbook,worksheets[1],infraItemsHeader)
createPageHeader(workbook,worksheets[2],infraTriggersHeader)
hostsList = createHostsList(zapi,hostGroup)
itemsList = createItemsList(zapi,hostGroup)
triggersList = createTriggersList(zapi, hostGroup)

createHostsTable(workbook,worksheets[0],hostsList)
createItemsTable(workbook,worksheets[1],prepareItemsTable(itemsList))
createTriggersTable(workbook,worksheets[2],triggersList)

workbook.close()
