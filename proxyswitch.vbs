' ProxySwitch v1.0
' Towsif Ahmed Omi
' 7 September 2016

OPTION EXPLICIT 
DIM WSHShell, strSetting
SET WSHShell = WScript.CREATEOBJECT("WScript.Shell")
'Determine current proxy setting and toggle to opposite setting
strSetting = wshshell.regread("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")

IF strSetting = 1 THEN 
NoProxy
 ELSE Proxy
END IF

'Subroutine to Toggle Proxy Setting to ON
SUB Proxy
set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", "182.160.104.213:8080"

WshShell.Popup "Proxy is ON!" & vbCrLf &"     - Created by OMI",0,"ProxySwitch",64
END SUB

'Subroutine to Toggle Proxy Setting to OFF
SUB NoProxy
set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", "0.0.0.0:80"
WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"

WshShell.Popup "ProxySwitch OFF!",1,"ProxySwitch",64
END SUB