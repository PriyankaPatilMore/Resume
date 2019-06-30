Call Main

Function Main
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True
 
    'Login
    IE.Navigate "https://hk.jobsdb.com/hk/en/login/jobseekerlogin?from=header"
    Wait IE
    With IE.Document
        .getElementByID("c_JbSrP1LnItDap_El0").value = "karri.suryarao5@gmail.com"
        .getElementByID("c_JbSrP1LnItDap_Pd0").value = "surya1223"
        .getElementByID("reg-login-button").Click
    End With
    
    'Jobs list with Filter URL
    Wait3 IE
    'IE.Navigate "https://th.jobsdb.com/TH/EN/Search/FindJobs?KeyOpt=COMPLEX&JSRV=1&RLRSF=1&JobCat=131&Career=4&JSSRC=CASOP&posFix=1&keepExtended=1&recentSelected=94"
    

Dim pageno
Dim applied
applied = 1
pageno =1
Do While pageno<=300
	IE.Navigate "https://hk.jobsdb.com/hk/jobs/information-technology/entry-level/"&pageno
    Wait1 IE

	set jobsDivs = IE.document.getElementsByClassName("fTybHge KXwQiR_")
        for each jobsDiv in jobsDivs ' Each Job div start
        	jobsDiv.click
            Wait2 IE

            set jobapplying = IE.document.getElementsByClassName("_2z0xbFq") ' Check job applied or not Start
            for each jobapplied in jobapplying 
                if(InStr(jobapplied.innerhtml,"applied") = "0") then
                    applied = 0
                end if
            Next ' Check job applied or not End

                if(applied = 1) then ' Tab and apply job Start
                    set jobsApplyBtns = IE.document.getElementsByClassName("_37Yu17M _2a81uiN DC5GtQ9 _16YdUsX _2nPU7y8")
                for each jobsApplyBtn in jobsApplyBtns ' Apply button click start
                    'msgbox jobsApplyBtn.href
                    IE2.Navigate jobsApplyBtn.href
                    WaitIE22 IE2
                    set btns = IE2.document.getElementsByClassName("btn btn-primary")
                    for each btn in btns
                      if(btn.innerhtml = "Apply now") then
                        btn.click
                        WaitIE22 IE2
                      end if
                    Next
                Next
                    
                end if ' Tab and apply job End
                applied = 1
        Next  ' Each Job div End

        'Pagination
        'set jobsNextPages = IE.document.getElementsByClassName("_3RBy0Im")
        'for each jobsNextPage in jobsNextPages
        '	if(jobsNextPage.innerhtml = "Next") then
        '		jobsNextPage.click
        '		Wait1 IE
        '	end if
        'Next
        
    pageno = pageno+1

Loop



End Function


Sub Wait(IE)
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait1(IE)
  Do
    WScript.Sleep 5000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait2(IE)
  Do
    WScript.Sleep 2000
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

Sub WaitIE24(IE2)
  Do
    WScript.Sleep 4000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub

Sub WaitIE210(IE2)
  Do
    WScript.Sleep 10000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub

Sub WaitIE25(IE2)
  Do
    WScript.Sleep 5000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub