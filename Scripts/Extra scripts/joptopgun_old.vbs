'Cross-site scripting off
'Allow pop ups

Call Main

Function Main
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
 
    'Login
    IE.Navigate "https://www.jobtopgun.com/?view=index&locale=en_TH" '********Login URL here*************
    Wait IE
    set objButtons = IE.document.getElementsByClassName("btn btn-block btn-custom btn-xs font_14 jskloginBtn")
    for each objButton in objButtons
            objButton.click
            Wait IE
            set usernames = IE.document.getElementsByName("username") 'Enter User name in the input field
            for each username in usernames
                username.value = "more.piyapatil@gmail.com" '********Type your User name*************
            Wait IE
            next
            set passwords = IE.document.getElementsByName("password") 'Enter Passeord in the input field
            for each password in passwords
                password.value = "Piyapavi12" '********Type your password*************
            Wait IE
            next
            set loginbuttons = IE.document.getElementsByClassName("btn font_18 whiteColor no_border modalLoginBtn normalBtn")
            for each loginbutton in loginbuttons
                strText = loginbutton.innerhtml
                Set re = New RegExp
                re.Pattern = "^\s+|\s+$"
                re.Global  = True
                CustomTrim = re.Replace(strText, "")
                if(CustomTrim = "Sign In") then
                    loginbutton.click
                end if

                Wait IE
            next
    Next
    'Login End


    'Apply Jobs
    Wait2 IE
    IE.Navigate "https://www.jobtopgun.com/%E0%B8%AB%E0%B8%B2%E0%B8%87%E0%B8%B2%E0%B8%99/IT/Computer%20Jobs/jobfield/6"
    'msgbox "hai"
    Wait1 IE
    Dim x
    Dim Pagination
    Dim flag
    Dim a
    a=2
    flag=1

    pagination=2
    x=0


Do While pagination<=60
            
'Apply Jobs start
set searchs1 = IE.document.getElementsByClassName("color0060CF font_16")
                    for each search1 in searchs1
                        strText = search1.innerhtml
                        strPostcode = Replace(strText, " ", "")
                        strPostcode = Replace(strPostCode, chr(10), "")
                        if(InStr(strPostcode,"Applynow") <> 0) then 
                            search1.click
                            Wait3 IE

                            'Page1
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
            WshShell.SendKeys " "
            Wait IE
            'msgbox "a"
            WshShell.SendKeys "{TAB}"
            Wait IE
            'msgbox "a"
            WshShell.SendKeys "{TAB}"
            Wait IE
            WshShell.SendKeys "{TAB}"
            Wait IE
            'WshShell.SendKeys "{TAB}"
            'Wait IE
            WshShell.SendKeys "{ENTER}"
            Wait3 IE
            'page2
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
            Wait3 IE
            'page3
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
            Wait3 IE
            WshShell.SendKeys "^w"
            Wait IE
            

                        end if
                    Next
'Apply Jobs end

'pagination changing
    set pageNoChecks = IE.document.getElementsByClassName("color5D5D5D font_16")
    for each pageNoCheck in pageNoChecks
    temppageno =  pageNoCheck.innerhtml
    if(temppageno = CStr(pagination)) then
        pageNoCheck.Click
        Wait3 IE
    end if
    Next

    pagination = pagination+1

Loop






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
    WScript.Sleep 8000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub