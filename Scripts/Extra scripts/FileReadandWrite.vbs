Set objFSO=CreateObject("Scripting.FileSystemObject")

'How to read a file
strFile = "F:\sample.txt"
Set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine= objFile.ReadLine
     temp = strLine
Loop
objFile.Close

' How to write file
outFile="F:\sample.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write temp & vbCrLf
objFile.Close
