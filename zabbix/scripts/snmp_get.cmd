@setLocal EnableDelayedExpansion
@echo off
@rem Valores
@rem %1 IP
@rem %2 OID
@rem %3 Cadena de conexion (public)

:: @echo IP = %1%
@SET IP=%1%
:: @echo OID = %2%
@SET OID=%2%
:: @echo COMUNITY = %3%
@SET COMUNITY=%3%


@set v1=0
for /f "tokens=2 delims=: " %%a in ('c:\zabbix\scripts\bin\snmpget\SNMPGET -M "c:\zabbix\scripts\bin\snmpget\mibs" -v 1 -c %COMUNITY% -O v %IP% %OID%') do (
	@set /a N+=1
	@set v!N!=%%a
)
@echo %v1%