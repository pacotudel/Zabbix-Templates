# Zabbix-Templates

## Conf
Descubrimiento
UserParameter=Ovt.Discover[*],cscript.exe c:\zabbix\scripts\discover.vbs $1 //NoLogo
UserParameter=Ovt.PingCheck[*],cscript.exe c:\zabbix\scripts\pingcheck.vbs $1 //NoLogo
UserParameter=Ovt.Service[*],cscript.exe c:\zabbix\scripts\Service_Status.vbs $1 //NoLogo
UserParameter=Ovt.NetworkShare[*],cscript.exe c:\zabbix\scripts\networkshare.vbs $1 $2 $3 $4 $5 //NoLogo
UserParameter=Ovt.NetworkShareFileAge[*],cscript.exe c:\zabbix\scripts\file_age_remote.vbs $1 $2 $3 $4 $5 $6 //NoLogo
UserParameter=Ovt.Snmpget[*],c:\zabbix\scripts\snmp_get.cmd $1 $2 $3