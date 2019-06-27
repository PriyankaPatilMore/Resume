'Cross-site scripting off
'Allow pop ups

Call Main

Function Main
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
 
    'Login
    IE.Navigate "http://job.prtr.com/user/login" '********Login URL here*************
    Wait IE
    set username = IE.document.getElementByID("email_address") 'Enter User name in the input field
    username.value = "karri.suryarao5@gmail.com" '********Type your User name*************
    Wait IE
            
            set password = IE.document.getElementByID("password") 'Enter Passeord in the input field
            password.value = "Surya@2019" '********Type your password*************
            Wait IE
            
    set loginbuttons = IE.document.getElementsByName("vc_apply")
            
    for each loginbutton in loginbuttons
                msgbox loginbutton.innerhtml
                loginbutton.click
            next
            
    
    'Login End

Wait2 IE

WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "^+{ENTER}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{ENTER}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{ENTER}"
            Wait3 IE
            WshShell.SendKeys "^w"
            



End Function

Sub Wait(IE)
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait1(IE)
  Do
    WScript.Sleep 60000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait2(IE)
  Do
    WScript.Sleep 10000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait3(IE)
  Do
    WScript.Sleep 4000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub