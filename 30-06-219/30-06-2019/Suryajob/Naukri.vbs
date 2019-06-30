Call Main

Function Main
	Set objFSO=CreateObject("Scripting.FileSystemObject")
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True
    'PageLoad
    IE.Navigate "https://www.naukri.com/information-technology-jobs-in-hyderabad-secunderabad-"&x
    Wait5 IE
    IE.document.getElementByID("exp_ddHid").value = "0"
    Wait IE
    IE.document.getElementByID("qsbFormBtn").click
    Wait4 IE

Dim x
x=1

Do While x<=76
  'Next Page button click
  set nextbtns = IE.document.getElementsByClassName("grayBtn")
  for each nextbtn in nextbtns
    if(nextbtn.innerhtml = "Next") then
      nextbtn.click
      Wait10 IE
    end if
  Next
    'Apply jobs start
    set aTags = IE.document.getElementsByTagName("a")
    for each aTag in aTags
    if(isNull(aTag.id) = False) then
            if(aTag.id = "jdUrl") then
              IE2.Navigate aTag.href
              WaitIE22 IE2
              if(isNull(IE2.document.getElementByID("trig1")) = False) then
                IE2.document.getElementByID("trig1").click
                WaitIE22 IE2
                if(isNull(IE2.document.getElementByID("skip_qup")) = False) then
                IE2.document.getElementByID("skip_qup").click
                WaitIE22 IE2
                end if
              end if
            end if
        end if    
    Next
    'Apply jobs End



x=x+1

Loop

End Function

Sub Wait(IE)
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait2(IE)
  Do
    WScript.Sleep 2000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait4(IE)
  Do
    WScript.Sleep 4000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait10(IE)
  Do
    WScript.Sleep 10000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait5(IE)
  Do
    WScript.Sleep 5000
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