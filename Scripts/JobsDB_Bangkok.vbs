Call Main

Function Main
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True
 
    'Login
    IE.Navigate "https://th.jobsdb.com/th/en/login/jobseekerlogin?from=header#start"
    Wait IE
    With IE.Document
        .getElementByID("c_JbSrP1LnItDap_El0").value = "more.piyapatil@gmail.com"
        .getElementByID("c_JbSrP1LnItDap_Pd0").value = "Piyapavi12"
        .getElementByID("reg-login-button").Click
    End With
    
    'Jobs list with Filter URL
    Wait3 IE
    'IE.Navigate "https://th.jobsdb.com/TH/EN/Search/FindJobs?KeyOpt=COMPLEX&JSRV=1&RLRSF=1&JobCat=131&Career=4&JSSRC=CASOP&posFix=1&keepExtended=1&recentSelected=94"
    IE.Navigate "https://th.jobsdb.com/TH/EN/Search/FindJobs?KeyOpt=COMPLEX&JSRV=1&RLRSF=1&JobCat=131&Career=4,3&JSSRC=CASOP&posFix=1&keepExtended=1&recentSelected=94"
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
    set jobHref = IE.Document.getElementByID("cp"&x)
    getapplyurls=Split(jobHref.href, "-")
    for each getapplyurltemp in getapplyurls
      getapplyurl = Split(getapplyurltemp, "?")
      'IE2.Navigate getapplyurl(0)
      if(IsNumeric(getapplyurl(0))) then
        IE2.Navigate "https://th.jobsdb.com/TH/en/Job/SelectCoverLetterAndResume?jobAdIdList="&getapplyurl(0)&"&trackData=%7B%22ApplySource%22%3A7%2C%22ABTest%22%3A%220%22%7D&IsInPopupPage=true&applySource=2&solAppTrackData=%7B%22jobID%22%3A%22300003001983185%22%2C%22jobTitle%22%3A%22AutoCad%20Programmer%2F%E0%B8%9E%E0%B8%99%E0%B8%B1%E0%B8%81%E0%B8%87%E0%B8%B2%E0%B8%99%E0%B9%80%E0%B8%82%E0%B8%B5%E0%B8%A2%E0%B8%99%E0%B9%81%E0%B8%9A%E0%B8%9A%20%E0%B8%AD%E0%B8%AD%E0%B9%82%E0%B8%95%E0%B9%89%E0%B9%81%E0%B8%84%E0%B8%94%22%2C%22rank%22%3A7%2C%22page%22%3A6%2C%22clickType%22%3A%22organic%22%2C%22searchID%22%3A%22fa0a8d15-c844-4bcb-b23c-ed27429be430%22%2C%22pageType%22%3A%22P%22%7D"
        WaitIE22 IE2
        set applybtns = IE2.document.getElementsByClassName("btn btn-primary")
        for each applybtn in applybtns
          if(applybtn.innerhtml = "Apply now") then
            IE2.document.getElementByID("c_PeAySyItDap_EdSy0").value = "35,000"
            WaitIE2 IE2
            IE2.document.getElementByID("c_PeAySyItDap_JbSrIo0").value = "Sir, I just started my career. While doing Masters I work as a part-timer in software field and I worked as an individual developer and learned new tech simultaneously for the development. Presently I am looking for an opportunity to in Software development."
            WaitIE2 IE2
            applybtn.click
            WaitIE22 IE2
          end if
        Next
      end if
    Next
    
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


Sub WaitIE2(IE2)
  Do
    WScript.Sleep 500
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub

Sub WaitIE22(IE2)
  Do
    WScript.Sleep 2000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub