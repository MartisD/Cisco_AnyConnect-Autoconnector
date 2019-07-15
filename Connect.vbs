
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"""

WScript.Sleep 3000

WshShell.AppActivate "Cisco AnyConnect Secure Mobility Client"

WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 1000
  
WshShell.SendKeys "^%{TAB}" 
WScript.Sleep 500
  
WshShell.SendKeys "{ENTER}"
WScript.Sleep 500

WshShell.SendKeys "password"
WshShell.SendKeys "{ENTER}"
