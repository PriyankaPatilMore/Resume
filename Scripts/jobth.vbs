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



    Dim x
    Dim Pagination
    
    pagination=1
Dim tempemails
tempemails = "a"

Do While pagination<=89

    'Apply Jobs
    Wait2 IE
    IE.Navigate "https://www.jobth.com/searchjob2.php?typejob=000003&typejob2=&city=&province=&jobmoney=&jobmoney2=&keyword=&page="&pagination
    Wait3 IE



  set applybtns = IE.document.getElementsByClassName("w3-large LinkVisited")

  for each applybtn in applybtns
    IE2.Navigate applybtn.href
    Wait22 IE2
    
    set emails = IE2.document.getElementsByClassName("w3-container w3-left-align w3-medium w3-theme-l5")
    for each email in emails
      email1=Split(email.innerhtml, " ")
      for each x in email1
        if(InStr(x, "@") <> "0") then
            if(InStr(x, ".") <> "0") then
              email2 = Split(x, "<br>")  
              for each y in email2
                if(InStr(y, "@") <> "0") then
                  tempemails = tempemails&"%"&y

                end if
              Next

            end if
        end if
      next
      Next


  Next

'How to read a file
strFile = "F:\jobthsample1.txt"
Set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine= objFile.ReadLine
Loop
objFile.Close
          
'Write to the file
fileName = "F:\jobthsample1.txt"
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set out = fso.CreateTextFile(fileName, True, True)
 out.WriteLine (tempemails)
 out.close


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
    WScript.Sleep 60000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait2(IE)
  Do
    WScript.Sleep 4000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait22(IE2)
  Do
    WScript.Sleep 4000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub

Sub Wait3(IE)
  Do
    WScript.Sleep 8000
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub