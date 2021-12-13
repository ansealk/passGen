from pyzabbix import ZabbixAPI

# IP-адрес сервера zabbix
ZABBIX_SERVER = 'http://zabbix.ansealk.ru'

zapi = ZabbixAPI(ZABBIX_SERVER)
zapi.login('', '')

hosts = zapi.host.get(
    output=['name','description','interfaces'],
    selectInterfaces=["ip"],
    selectParentTemplates=['name'],
    seletcTags=True,
    selectGroups=['Linux Servers'],
    expandComment=True,
    expandDescription=True,
    expandExpression=True,
)
print(hosts)

