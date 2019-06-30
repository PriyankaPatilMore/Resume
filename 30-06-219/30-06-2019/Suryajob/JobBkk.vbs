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
Dim FirstOrSecond
Dim allCompanies(100000)
Dim tempdelcompanies
Dim allCompaniesCount
FirstOrSecond = 0
Dim allCompanies1 : allCompanies1= Array(45052,57346,24528,176707,24356,175525,33627,182881,48733,20868,45312,32465,178848,168551,47965,44707,183534,169848,40538,174790,40475,44981,12544,8786,14156,48910,34948,164304,161172,24252,181121,5879,169315,47727,169299,169603,159479,169539,175068,173802,43695,7768,41314,182550,12121,39360,48296,169811,12837,166115,30884,167342,159080,176277,184061,7688,56757,181841,38714,43586,53888,34594,47746,6628,46587,163805,49317,177123,27718,166615,55283,157666,183264,44413,46905,39768,51918,173839,6142,50407,21658,30074,37389,182198,54268,160756,46827,172746,54342,32379,47375,57415,183589,177319,168471,36707,44282,50640,21025,35281,158880,48492,15452,12718,180567,27043,52578,161187,18002,46413,33559,44929,23719,46759,40302,24445,157408,51945,14004,13716,14063,168959,41582,45248,38949,38439,37631,50557,24888,56389,29009,183519,26418,171147,30780,54000,24632,55666,41961,42249,20347,177351,183789,174452,177630,19031,184063,182760,178493,184185,175711,165195,12076,42459,19186,41736,31666,18462,167642,156814,39254,172120,159440,16975,18253,17631,57473,179926,33965,58627,176167,184143,169207,34797,170342,39361,25985,38603,164035,35260,28347,47265,182387,25498,184225,51837,44989,18911,34084,170993,179981,182922,170657,175124,47590,39173,37887,51788,16940,169234,184039,39757,177800,167511,31752,55618,165789,177051,40294,184216,184229,169230,21560,53959,31352,169912,172198,166698,23324,27049,33004,162288,180804,23651,25029,23105,50365,47875,25173,3010,39182,31787,161100,50610,35062,183920,184167,29737,184187,25614,44650,49472,33268,167578,26071,165638,43316,54768,36223,164321,43489,51914,26381,44831,36726,36279,40864,164211,40710,38904,177670,175540,44055,169798,53173,52249,159778,23072,14100,55986,47346,2887,177431,163105,42161,20377,38304,182886,182565,30754,28977,184157,50804,51241,31847,57184,182282,21004,53814,28838,26894,182494,10829,31873,14368,37838,15036,165390,27235,42013,28410,50009,167171,52743,7061,7137,19529,163768,55927,47745,24328,51356,164922,39737,160890,13879,178452,14677,172604,157332,169862,173163,38721,30643,41630,28269,181590,33325,174191,162685,10670,48580,181202,21035,37340,41767,30000,158781,34259,182990,170589,179530,182337,44477,36608,46699,170348,26755,46688,30797,34216,19806,173518,184203,27421,30923,32643,45072,167135,36616,163945,32245,184147,162407,159223,55391,27648,174795,53904,37684,23865,160802,56516,179742,52541,18493,44314,12144,41872,36938,178777,45476,29872,31827,177580,170494,51039,180281,15082,52058,29959,52758,183455,28826,19196,32302,18269,11600,174212,46931,37785,36619,25553,24675,2064,37800,44626,178153,173849,54792,167371,47445,27780,36549,169952,181154,25089,174673,183323,181337,21738,173362,57784,50217,55732,178154,55869,36997,9821,43790,30735,165735,161166,20549,25401,53948,29333,39953,183441,184102,44680,22123,53472,36274,54222,45371,22662,160298,20128,166693,44060,54995,40904,33310,12972,183775,25792,182759,34120,33322,29629,41232,25219,36923,39301,48724,162811,46032,14267,39778,6339,42460,180948,169931,183862,27277,22487,42121,16342,174585,46083,46290,20029,17449,171561,23471,164578,38568,33809,51263,40341,38156,14577,40281,28017,19487,30289,57388,178660,45745,34977,167264,183903,14318,174001,33951,14691,45825,163440,30737,19158,165054,20624,22938,22607,7537,176480,32693,159168,184026,24796,162961,165935,27770,36651,12297,51669,156733,178032,43965,22936,178381,35689,54687,10788,12097,41006,174700,57384,28684,45707,160072,180508,49501,33763,55529,160317,27391,56732,43823,157083,38695,46481,49175,40564,56474,41621,23485,53611,183520,162357,37432,15084,33168,48264,23069,34004,157154,29499,55226,20229,47709,19059,180060,44653,177277,32243,174579,177297,181516,170626,54246,8994,175675,31500,40466,56876,16491,44363,30826,46995,173205,38324,176304,56948,176084,50456,173980,53116,21575,178896,177855,158749,49346,27020,48987,47932,181636,39240,57638,52781,41944,176595,173180,182454,33928,20594,22252,52163,39146,11649,183057,38462,18135,174428,170376,27751,49067,21975,163295,41480,21396,180266,55884,167972,58254,58196,58287,15047,157689,170012,23855,175733,38822,32685,173526,15114,44836,179569,30978,55893,28241,37146,15209,31042,34749,29338,26085,158109,20832,31377,50085,54960,36371,51187,184184,53882,46153,58325,39580,32682,38401,53778,161144,39270,164012,50045,55093,58622,44839,173556,165506)


tempdelcompanies = "1"
x=306
tempdel = 0
for each allCompany1 in allCompanies1
allCompanies(tempdel) = allCompany1
tempdel = tempdel + 1
Next

allCompaniesCount = tempdel




Do While x<=1100
'Jobs list with Filter URL
    IE.Navigate "https://www.jobbkk.com/jobs/lists/"&x&"/%E0%B8%AB%E0%B8%B2%E0%B8%87%E0%B8%B2%E0%B8%99,%E0%B8%97%E0%B8%B1%E0%B9%89%E0%B8%87%E0%B8%AB%E0%B8%A1%E0%B8%94,%E0%B8%81%E0%B8%A3%E0%B8%B8%E0%B8%87%E0%B9%80%E0%B8%97%E0%B8%9E%E0%B8%A1%E0%B8%AB%E0%B8%B2%E0%B8%99%E0%B8%84%E0%B8%A3,%E0%B8%97%E0%B8%B1%E0%B9%89%E0%B8%87%E0%B8%AB%E0%B8%A1%E0%B8%94.html?province_id=246&keyword_type=1"
    Wait1 IE

tempemails="aaaaaaaaaa"
    set jobapplying = IE.document.getElementsByTagName("a") ' Check job applied or not Start
            for each jobapplied in jobapplying 
              if(IsNull(jobapplied.href) = False) then  
                if(InStr(jobapplied.href,"https://www.jobbkk.com/jobs/detail") <> "0") then
                	checkEmail = Split(jobapplied.href, "/")
                	'msgbox checkEmail(5)

                	If Ubound(Filter(allCompanies, checkEmail(5))) > -1 Then
        				    'nothing
        			    Else
           				  allCompanies(allCompaniesCount) = checkEmail(5)
           				  allCompaniesCount = allCompaniesCount + 1
						tempdelcompanies = tempdelcompanies&"%"&checkEmail(5)
                    IE2.Navigate jobapplied.href
                    WaitIE22 IE2
                    set emailsections = IE2.document.getElementsByClassName("job-detail-content")
                    for each emailsection in emailsections
                        emailarrs = Split(emailsection.innerhtml, " ")
                          for each emailarr in emailarrs
                            'msgbox emailarr
                            if(InStr(emailarr,"@") <> "0") then
                              emailfinal = Split(emailarr, "<br>")
				if(FirstOrSecond = 1) then
					tempdelemails = tempdelemails&","&emailfinal(0)
				end if
				if(FirstOrSecond = 0) then
					tempdelemails = emailfinal(0)
					FirstOrSecond = 1 
				end if
				
                            end if
                          Next
			
                    Next
                  End If
		tempemails = tempemails&"%"&tempdelemails
			FirstOrSecond = 0
tempdelemails="0"
                end if

              end if  
		
            Next ' Check job applied or not End
    
if(x<=310) then
strFile = "F:\Suryajob\email1.txt"
outFile="F:\Suryajob\email1.txt"
end if

if(x>310 And x<=320) then
strFile = "F:\Suryajob\email2.txt"
outFile="F:\Suryajob\email2.txt"
end if

if(x>320 And x<=340) then
strFile = "F:\Suryajob\email3.txt"
outFile="F:\Suryajob\email3.txt"
end if

if(x>340 And x<=470) then
strFile = "F:\Suryajob\email4.txt"
outFile="F:\Suryajob\email4.txt"
end if

if(x>370 And x<=400) then
strFile = "F:\Suryajob\email5.txt"
outFile="F:\Suryajob\email5.txt"
end if

if(x>430 And x<=470) then
strFile = "F:\Suryajob\email6.txt"
outFile="F:\Suryajob\email6.txt"
end if

if(x>470) then
strFile = "F:\Suryajob\email7.txt"
outFile="F:\Suryajob\email7.txt"
end if




'How to read a file
'strFile = "F:\Suryajob\email1.txt"
Set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine= objFile.ReadLine
Loop
objFile.Close

' How to write file
'outFile="F:\Suryajob\email1.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write strLine&"%"&tempemails & vbCrLf
objFile.Close



' How to write file
outFileComp="F:\Suryajob\Companyids.txt"
Set objFileComp = objFSO.CreateTextFile(outFileComp,True)
objFileComp.Write  tempdelcompanies & vbCrLf
objFileComp.Close


  
x=x+1
Loop


End Function

Sub Wait(IE)
  Do
    WScript.Sleep 500
  Loop While IE.ReadyState < 4 And IE.Busy
End Sub

Sub Wait1(IE)
  Do
    WScript.Sleep 10000
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