'*************************Note*****************************
'Login Before
'**********************************************************
Call Main

Function Main
	Set objFSO=CreateObject("Scripting.FileSystemObject")
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True
    
Dim x
x=15

Dim Finalmailarr
Finalmailarr = "aa"
Do While x<=37
    
    'PageLoad
     IE.Navigate "http://yend.org/?s&category&location=141&a=true&paged="&x
    WScript.sleep 100
    Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  


    set listOfCompanies = IE.document.getElementsByClassName("main-link")
    for each listOfCompany in listOfCompanies
            if(InStr(listOfCompany.href, "http://yend.org/item/") <> "0") then
                'msgbox listOfCompany.href
                IE2.Navigate listOfCompany.href
                WScript.sleep 100
    Do While IE2.Busy or IE2.ReadyState <> 4: WScript.Sleep 100: Loop 
                set listOfEmails = IE2.document.getElementsByTagName("a")
                   for each listOfEmail in listOfEmails
                    if(isNull(listOfEmail.href) = false)then
                        if(InStr(listOfEmail.href, "mailto:") <> "0" And InStr(listOfEmail.href, "ying.org") = "0") then
                                mailarr = Split(listOfEmail.href,":")
                                mail = mailarr(1)
                                'msgbox mail
                        end if
                    end if
                    Next
                    Finalmailarr = Finalmailarr&"%"&mail 
            end if
    Next

    'msgbox "Final emails"&Finalmailarr

x=x+1

Wscript.Echo Finalmailarr

Loop

Wscript.Echo Finalmailarr

msgbox "Done"
    

End Function

Sub Wait(IE)
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait1(IE)
  Do
    WScript.Sleep 25000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait2(IE)
  Do
    WScript.Sleep 120000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait3(IE)
  Do
    WScript.Sleep 5000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub WaitIE210(IE2)
  Do
    WScript.Sleep 10000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub