'Cross-site scripting off
'Allow pop ups

Call Main

Function Main
Set objFSO=CreateObject("Scripting.FileSystemObject")
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True
 
    'Login
    IE.Navigate "https://www.jobtopgun.com/?view=index&locale=en_TH" '********Login URL here*************
    wscript.sleep 100
Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  
    set objButtons = IE.document.getElementsByClassName("btn btn-block btn-custom btn-xs font_14 jskloginBtn")
    for each objButton in objButtons
            objButton.click
            wscript.sleep 100
Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  
            set usernames = IE.document.getElementsByName("username") 'Enter User name in the input field
            for each username in usernames
                username.value = "more.piyapatil@gmail.com" '********Type your User name*************
            wscript.sleep 100
Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  
            next
            set passwords = IE.document.getElementsByName("password") 'Enter Passeord in the input field
            for each password in passwords
                password.value = "Piyapavi12" '********Type your password*************
            wscript.sleep 100
Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  
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

                wscript.sleep 100
Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  
            next
    Next
    'Login End


    'Apply Jobs
    Wait2 IE
    IE.Navigate "https://www.jobtopgun.com/%E0%B8%AB%E0%B8%B2%E0%B8%87%E0%B8%B2%E0%B8%99/IT/Computer%20Jobs/jobfield/6"
    IE2.Navigate "https://www.jobtopgun.com/%E0%B8%AB%E0%B8%B2%E0%B8%87%E0%B8%B2%E0%B8%99/IT/Computer%20Jobs/jobfield/6"
    
   	wscript.sleep 100
Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  
    Dim x
    Dim Pagination
    Dim flag
    Dim a
    a=2
    flag=1

    pagination=2
    x=0


Do While pagination<=60
  
  set applybtns = IE.document.getElementsByClassName("col-sm-12 col-xs-6 text-left")

  for each applybtn in applybtns
  if((InStr(applybtn.innerhtml, "register.superresume.com") <> "0") And (InStr(applybtn.innerhtml, "Apply now") <> "0")) then
    getapplybtn = Split(applybtn.innerhtml, Chr(34))
  
  	'Window2 start
    IE2.Navigate Replace(getapplybtn(3), "amp;", "")
  	wscript.sleep 100
Do While IE2.Busy or IE2.ReadyState <> 4: WScript.Sleep 100: Loop  
    
    'button click
    if(IsNull(IE2.document.getElementByID("applyButton")) = False) then
  	IE2.document.getElementByID("applyButton").click
  	wscript.sleep 100
Do While IE2.Busy or IE2.ReadyState <> 4: WScript.Sleep 100: Loop  
    end if

    'button click on final submission
  	if(IsNull(IE2.document.getElementByID("btn_submit")) = False) then
    IE2.document.getElementByID("btn_submit").click
  	wscript.sleep 100
Do While IE2.Busy or IE2.ReadyState <> 4: WScript.Sleep 100: Loop  
    end if
  	set finalbtns = IE2.document.getElementsByTagName("button")
  	for each finalbtn in finalbtns
  		if(finalbtn.ClassName = "btn btn-block btn-primary") then
  			finalbtn.click
  			wscript.sleep 100
Do While IE2.Busy or IE2.ReadyState <> 4: WScript.Sleep 100: Loop  
  			wscript.sleep 100
Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  
  		end if
  	Next
  	'Window2 End
    
  end if
  Next

    'pagination changing
    set pageNoChecks = IE.document.getElementsByClassName("color5D5D5D font_16")
    for each pageNoCheck in pageNoChecks
    temppageno =  pageNoCheck.innerhtml
    if(temppageno = CStr(pagination)) then
        pageNoCheck.Click
        wscript.sleep 100
Do While IE.Busy or IE.ReadyState <> 4: WScript.Sleep 100: Loop  
    end if
    Next

    pagination = pagination+1
Loop






End Function

Sub Wait(IE)
  While IE.readyState <> 4 Or IE.Busy: DoEvents: Wend
End Sub

Sub Wait1(IE)
  Do
    WScript.Sleep 60000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait2(IE)
  Do
    WScript.Sleep 4000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait3(IE)
  Do
    WScript.Sleep 8000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub


Sub Wait25(IE)
  Do
    WScript.Sleep 25000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub WaitIE21(IE2)
  Do
    WScript.Sleep 60000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub

Sub WaitIE2(IE2)
 While IE.readyState <> 4 Or IE.Busy: DoEvents: Wend
End Sub

Sub WaitIE23(IE2)
  Do
    WScript.Sleep 8000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub