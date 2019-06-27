Call Main

Function Main
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
 
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
    IE.Navigate "https://hk.jobsdb.com/hk/jobs/information-technology/entry-level/1"
    Wait1 IE

Dim pageno
Dim applied
applied = 1
pageno =1
Do While pageno<=50
	
	set jobsDivs = IE.document.getElementsByClassName("fTybHge KXwQiR_")
        for each jobsDiv in jobsDivs ' Each Job div start
        	jobsDiv.click
            Wait IE

            set jobapplying = IE.document.getElementsByClassName("_2z0xbFq") ' Check job applied or not Start
            for each jobapplied in jobapplying 
                if(InStr(jobapplied.innerhtml,"applied") = "0") then
                    applied = 0
                end if
            Next ' Check job applied or not End

                if(applied = 1) then ' Tab and apply job Start
                    set jobsApplyBtns = IE.document.getElementsByClassName("_37Yu17M _2a81uiN DC5GtQ9 _16YdUsX _2nPU7y8")
                for each jobsApplyBtn in jobsApplyBtns ' Apply button click start
                    jobsApplyBtn.click
        
                    'Apply job and exit the Tab
                    Wait3 IE
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

                Next ' Apply button click end

                Wait IE


                end if ' Tab and apply job End
                applied = 1
        Next  ' Each Job div End

        'Pagination
        set jobsNextPages = IE.document.getElementsByClassName("_3RBy0Im")
        for each jobsNextPage in jobsNextPages
        	if(jobsNextPage.innerhtml = "Next") then
        		jobsNextPage.click
        		Wait1 IE
        	end if
        Next
        
    pageno = pageno+1
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