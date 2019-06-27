Call Main

Function Main
	Set objFSO=CreateObject("Scripting.FileSystemObject")
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    Set IE2 = WScript.CreateObject("InternetExplorer.Application", "IE_")
    IE.Visible = True
    IE2.Visible = True

Dim x
x=56


Do While x<=76
    
    'PageLoad
    'IE.Navigate "https://www.jobthai.com/searchjob/Computer-IT-Programmer,p"&x&".html"
     IE.Navigate "https://www.jobthai.com/home/job_list.php?l=en&JobType=Computer&selSort=companyname&station[]=1&station[]=2&station[]=3&station[]=4&station[]=5&station[]=6&p="&x
    IE2.Navigate "https://www.jobthai.com/searchjob/Computer-IT-Programmer.html"
    Wait2 IE

Dim tempemails
tempemails = "a"

    set searchs1 = IE.document.getElementsByTagName("a")
        for each search1 in searchs1
        if(search1.className = "searchjob") then
        'Wait IE
        	if(InStr(search1,"/job/") <> 0) then 
        		
        		IE2.Navigate search1.href ' Second page to get the emails Start
        		Wait3 IE2
        		set attr2 = IE2.document.getElementsByTagName("a")
        		for each attr in attr2 'attr foreach start
        			if(InStr(attr.className, "searchjob") <> "0") then ' Get all attr with classname searchjob Start
        				if(InStr(attr.innerhtml, "@trustmail") <> "0") then ' Get only email Start
        					tempemails = tempemails&"%"&attr.innerhtml
        				end if  ' Get only End
        			end if  ' Get all attr with classname searchjob End
        		Next 'attr foreach start 
        		'Second page to get the emails End

        	end if
        end if
        'IE2.Navigate search1.href
        'Wait3 IE2

        Next

x=x+1
Wait2 IE

'How to read a file
strFile = "F:\JobThaiCompanies2.txt"
Set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine= objFile.ReadLine
Loop
objFile.Close

' How to write file
outFile="F:\JobThaiCompanies2.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write strLine&"%"&tempemails & vbCrLf
objFile.Close


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