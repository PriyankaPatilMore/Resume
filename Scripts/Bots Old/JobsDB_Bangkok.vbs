Call Main

Function Main
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
 
    'Login
    IE.Navigate "https://th.jobsdb.com/th/en/login/jobseekerlogin?from=header#start"
    Wait IE
    With IE.Document
        .getElementByID("c_JbSrP1LnItDap_El0").value = "karri.suryarao5@gmail.com"
        .getElementByID("c_JbSrP1LnItDap_Pd0").value = "surya1223"
        .getElementByID("reg-login-button").Click
    End With
    
    'Jobs list with Filter URL
    Wait3 IE
    'IE.Navigate "https://th.jobsdb.com/TH/EN/Search/FindJobs?KeyOpt=COMPLEX&JSRV=1&RLRSF=1&JobCat=131&Career=4&JSSRC=CASOP&posFix=1&keepExtended=1&recentSelected=94"
    IE.Navigate "https://th.jobsdb.com/TH/EN/Search/FindJobs?KeyOpt=COMPLEX&JSRV=1&RLRSF=1&JobCat=131&Career=4&JSSRC=CSB&posFix=1&keepExtended=1"
    Wait1 IE

Dim x
Dim pageno
Dim applied
applied = 1
pageno =1
x=1
Do While x<=50

    'Click on the job
    With IE.Document
        .getElementByID("cp"&x).Click
    End With
    Wait IE
    set jobapplying = IE.document.getElementsByClassName("result-sherlock-cell applied selected") ' Check job applied or not Start
            for each jobapplied in jobapplying     
                if(InStr(jobapplied.innerhtml,"Applied") <> "0") then
                    applied = 0
                end if
            Next ' Check job applied or not End
    
    if(applied = 1) then
    Wait IE
    With IE.Document
        .getElementByID("Btn_Apply").Click
    End With

    Wait2 IE

    'Apply job and exit the Tab
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
    Wait3 IE
    WshShell.SendKeys "^w"
    Wait IE
    end if
    
    applied = 1

    'Change the Pagination
    if (x=50) then
        pageno = pageno+1
        set objButtons = IE.document.getElementsByClassName("pagebox")
        for each objButton in objButtons
            strText = objButton.innerhtml
            strPostcode = Replace(strText, " ", "")
            strPostcode = Replace(strPostCode, chr(10), "")
            if (strPostcode = "<span>"&pageno&"</span>") then
                'objButton.click
                x=0
                IE.Navigate objButton.href
                Wait1 IE
                Exit For
            end if
        Next
    end if

x=x+1
if pageno=50 then
 Exit Do
end if
Loop



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
    WScript.Sleep 10000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait3(IE)
  Do
    WScript.Sleep 8000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub