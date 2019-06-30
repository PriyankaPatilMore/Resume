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
    'IE2.Navigate "https://www.jobtopgun.com/%E0%B8%AB%E0%B8%B2%E0%B8%87%E0%B8%B2%E0%B8%99/IT/Computer%20Jobs/jobfield/6"
    
   	Wait3s IE
    Dim dups(100)
	Dim count
Dim flag
flag=1
count = 1
dups(0)= "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
    Dim Pagination
    pagination=2


Do While pagination<=1000
  count=1
  set applybtns = IE.document.getElementsByClassName("col-sm-12 col-xs-6 text-left")

  for each applybtn in applybtns
  if((InStr(applybtn.innerhtml, "register.superresume.com") <> "0") And (InStr(applybtn.innerhtml, "Apply now") <> "0")) then
    getapplybtn = Split(applybtn.innerhtml, Chr(34))
	for each dup in dups
		if(InStr(dup, getapplybtn(3))<> "0") then
			flag = 0
		end if
	Next

  dups(count) = getapplybtn(3)
count = count+1

  	'Window2 start
if(flag = 1) then
    IE2.Navigate Replace(getapplybtn(3), "amp;", "")
  	WaitIE22 IE2
    
    'button click
    if(IsNull(IE2.document.getElementByID("applyButton")) = False) then
  	IE2.document.getElementByID("applyButton").click
  	WaitIE22 IE2
    end if

    'button click on final submission
  	if(IsNull(IE2.document.getElementByID("btn_submit")) = False) then
    IE2.document.getElementByID("btn_submit").click
  	WaitIE22 IE2
    end if
	if(IsNull(IE2.document.getElementsByTagName("button")) = False) then 'finalbtns check start
  	set finalbtns = IE2.document.getElementsByTagName("button")
  	for each finalbtn in finalbtns		
	  if(IsNull(finalbtn.ClassName) = False) then
  		if(finalbtn.ClassName = "btn btn-block btn-primary") then
  			finalbtn.click
  			WaitIE22 IE2
  			Wait IE
  		end if
	end if
  	Next
	end if 'finalbtns check end

	end if
  	'Window2 End
flag = 1
    
  end if
  Next

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
    WScript.Sleep 6000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait2(IE)
  Do
    WScript.Sleep 4000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait3(IE)
  Do
    WScript.Sleep 5000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait3s(IE)
  Do
    WScript.Sleep 80000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub


Sub WaitIE21(IE2)
  Do
    WScript.Sleep 60000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub

Sub WaitIE22(IE2)
  Do
    WScript.Sleep 5000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub

Sub WaitIE23(IE2)
  Do
    WScript.Sleep 8000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub