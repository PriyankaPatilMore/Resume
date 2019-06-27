Call Main

Function Main
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
 
    'Login
    'IE.Navigate "https://angel.co/login"
    'Wait3 IE
    'With IE.Document
      '  .getElementByID("user_email").value = "more.piyapatil@gmail.com"
     '   .getElementByID("user_password").value = "Piyapavi12"
    'End With

    'set loginBtns = IE.document.getElementsByClassName("c-button c-button--blue s-vgPadLeft1_5 s-vgPadRight1_5")
    'for each loginBtn in loginBtns
     ' loginBtn.click
    'Next
    
    'Jobs list with Filter URL
    IE.Navigate "https://angel.co/jobs#find/f!%7B%22locations%22%3A%5B%221644-Hong%20Kong%22%5D%2C%22roles%22%3A%5B%22Software%20Engineer%22%5D%7D"
    Wait25 IE

	
	set jobsDivs = IE.document.getElementsByClassName("startup-link")
        for each jobsDiv in jobsDivs ' Each Job div start
        	jobsDiv.click
            Wait IE
        		set jobsApplyBtns = IE.document.getElementsByClassName("g-button blue apply-now-button")
        		for each jobsApplyBtn in jobsApplyBtns ' Apply button click start
        			          jobsApplyBtn.click
        			          Wait2 IE
        			          set clickapplies = IE.document.getElementsByClassName("fontello-paper-plane")
        			          for each clickapplie in clickapplies
        			          	clickapplie.click
        			          	Wait2 IE
        			          	set closeapplies = IE.document.getElementsByClassName("c-button c-button--blue")
        			          	for each closeapplie in closeapplies
        			          	if(closeapplies.innerhtml = "Close") then
        			          		closeapplie.click
        			          	end if
        			          	Next
        			          Next
                        
        		Next ' Apply button click end

        	Wait IE
        Next  ' Each Job div End

        



End Function

Sub Wait(IE)
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait1(IE)
  Do
    WScript.Sleep 15000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait2(IE)
  Do
    WScript.Sleep 2000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait25(IE)
  Do
    WScript.Sleep 25000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait3(IE)
  Do
    WScript.Sleep 8000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub