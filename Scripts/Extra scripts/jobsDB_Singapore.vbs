Call Main

Function Main
	Set objFSO=CreateObject("Scripting.FileSystemObject")
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True

    'PageLoad
    IE.Navigate "https://www.jobstreet.com.sg/en/job-search/job-vacancy.php?area=1&specialization=191&position=5%2C4&job-type=5&experience-min=-1&experience-max=-1&classified=1&salary-option=on&job-posted=0&src=1&ojs=4"
    IE2.Navigate "https://www.jobthai.com/searchjob/Computer-IT-Programmer.html"
    Wait IE
    Wait IE2

Dim x
Dim page
page=2

Do While page<=300

x=1
Do While x<=20

    set searchs1 = IE.document.getElementByID("position_title_"&x)
    IE2.Navigate searchs1.href
    Wait3 IE2
    set applybtn = IE2.Document.getElementByID("apply_button")
    applybtn.click
    Wait3 IE2
             'IE2.Navigate IE2.LocationURL
             'Wait3 IE2
    if(InStr(IE2.LocationURL, "myjobstreet.jobstreet.com.sg") <> 0) then
        IE2.document.getElementByID("pitch_text").value = "Sir, I just started my career. While doing UG and Masters I work as a part-timer(1* year) in software field and I worked as an individual developer and learned new tech simultaneously to develop Web/Mobile app and API's. Presently I am looking for an opportunity to in Software development."
        Wait IE2
        IE2.document.getElementByID("btnAction").click
    end if
    Wait3 IE2

x=x+1
Loop

IE.Document.getElementByID("page_"&page).click
Wait3 IE
page=page+1

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