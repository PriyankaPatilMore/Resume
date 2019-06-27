Call Main

Function Main
    Set objFSO=CreateObject("Scripting.FileSystemObject")
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True
 
    
    

Dim x
Dim pageno
Dim tempemails
x=2
Do While x<=1000

'Jobs list with Filter URL
    IE.Navigate "https://www.jobbkk.com/jobs/lists/"&x&"/%E0%B8%AB%E0%B8%B2%E0%B8%87%E0%B8%B2%E0%B8%99,%E0%B8%97%E0%B8%B1%E0%B9%89%E0%B8%87%E0%B8%AB%E0%B8%A1%E0%B8%94,%E0%B8%81%E0%B8%A3%E0%B8%B8%E0%B8%87%E0%B9%80%E0%B8%97%E0%B8%9E%E0%B8%A1%E0%B8%AB%E0%B8%B2%E0%B8%99%E0%B8%84%E0%B8%A3,%E0%B8%97%E0%B8%B1%E0%B9%89%E0%B8%87%E0%B8%AB%E0%B8%A1%E0%B8%94.html?province_id=246&keyword_type=1"
    Wait1 IE




tempemails="aaaaaaaaaa"
    set jobapplying = IE.document.getElementsByTagName("a") ' Check job applied or not Start
            for each jobapplied in jobapplying 
              if(IsNull(jobapplied.href) = False) then  
                if(InStr(jobapplied.href,"https://www.jobbkk.com/jobs/detail") <> "0") then
                    IE2.Navigate jobapplied.href
                    WaitIE22 IE2
                    set emailsections = IE2.document.getElementsByClassName("job-detail-content")

                    for each emailsection in emailsections
                          emailarrs = Split(emailsection.innerhtml, " ")
                          for each emailarr in emailarrs
                          'msgbox emailarr
                          if(InStr(emailarr,"@") <> "0") then
                            emailfinal = Split(emailarr, "<br>")
                            tempemails = tempemails&"%"&emailfinal(0)
                          end if
                          Next
                    Next

                end if
              end if  
            Next ' Check job applied or not End
    


'How to read a file
strFile = "E:\Piya\PriyankaGit\Scripts\JobBkk.txt"
Set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine= objFile.ReadLine
Loop
objFile.Close

' How to write file
outFile="E:\Piya\PriyankaGit\Scripts\JobBkk.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write strLine&"%"&tempemails & vbCrLf
objFile.Close

  
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
    WScript.Sleep 15000
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
    WScript.Sleep 5000
  Loop While IE2.ReadyState < 4 And IE2.Busy
End Sub