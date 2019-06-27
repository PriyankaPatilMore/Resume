'*************************Note*****************************
'Login Before
'**********************************************************
Call Main

Function Main
	Set objFSO=CreateObject("Scripting.FileSystemObject")
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE3 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True
    IE3.Visible = True

Dim x
x=7


Do While x<=76
    
    'PageLoad
    'IE.Navigate "https://www.jobthai.com/searchjob/Computer-IT-Programmer,p"&x&".html"
     IE.Navigate "https://www.jobthai.com/home/job_list.php?l=en&JobType=Computer&selSort=companyname&station[]=1&station[]=2&station[]=3&station[]=4&station[]=5&station[]=6&p="&x
    IE2.Navigate "https://www.jobthai.com/searchjob/Computer-IT-Programmer.html"
    Wait2 IE

Dim tempemails
Dim tempcount
tempcount = 1
tempemails = "a"

    set searchs1 = IE.document.getElementsByTagName("a")
        for each search1 in searchs1
            tempcount = 1
        if(search1.className = "searchjob") then
        'Wait IE
        	if(InStr(search1,"/job/") <> 0) then 
        		IE2.Navigate search1.href ' Second page to get the emails Start
        		Wait3 IE2

        		set geturls = IE2.document.getElementsByClassName("linkbutton")
                for each geturl in geturls
                    getapplyurls=Split(geturl.href, "'")
                    if(tempcount = 1) then
                        IE3.Navigate "https://www.jobthai.com/home/apply_job.php?l=en&comcode="&getapplyurls(1)&"&jobcode="&getapplyurls(3)&"&resumecode="&getapplyurls(1)&"&fromapply="
                        Wait IE3
                                                    'IE3.document.getElementByID("bt_send").click
                                                    'IE3.Document.getElementsByName("btnK").Item(0).Click

                        if(IsNull(IE3.document.getElementById("bt_send"))) then
                        tempcount = 0
                         else

                            IE3.document.getElementByID("noteresume").value = "After completion of my Masters study I return back to India due to completion of my student visa. If my profile is consider it would be great for me to attend online interview. I value your feedback."
                            Wait IE3
                            IE3.document.parentWindow.sendresume()
                            end if
                         Wait2 IE3
                        tempcount = 0
                    end if  
                Next
        	end if
        end if
        'IE2.Navigate search1.href
        'Wait3 IE2

        Next

x=x+1
Wait2 IE

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
    WScript.Sleep 5000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub