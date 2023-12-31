Public BrowserCreationTimeIndex
Public tExecutionStartTime
Sub ErrComponentCreation(sTestcaseId)
    
	ExecuteGlobal "Dim QTPReporter"
	Set QTPReporter = Reporter
	ExecuteGlobal "Dim Reporter"
	
	Set Reporter = New CustomReporter
	Reporter.tExcelFilePath=EXCEL_RESULT_FILE
	
End Sub 


'*******************



''Help


'************************
Sub ReportWriter(sEventStatus,sReportStepName,sDetails,bExitRun)
	Reporter.ReportEvent sEventStatus,sReportStepName,sDetails,bExitRun
End Sub


'***********
'help





'*************************
function ReportStart(sTestName)
		TName = sTestName
	Reporter.StartTest(TName)
End function


'*****************************

'help




'********************************
function ReportEnd()
	Reporter.EndTest
End function

Class CustomReporter
	Public ReportSheet
	Public ReportSheetName
	Public iIndex
	Public sName
	Public sDes
	Public sStatus
	Public sScreen
	Public sScreenShotPath
	Public clsTestCaseID
	'Public tExecutionStartTime
	Public tExecutionEndTime
	Public tExecutionTimeZone
	Public tExecutionTime
	Public CurrentFunctionName
	Public FunctionCallFlow
	Public eLogFilePath
	Public cErrLogFolderPath
	Public eImageFilePath
	Public ExecutionData
	Public ExecutionEnvironment
	Public tHtmlResultFilePath
	Public TrackStatus
	Public lGroupStatus
	Public TestCasePriority
	Public tExcelFilePath
	Public clsErrDescription
	Public cErrLogName
	'Public BrowserCreationTimeIndex
	Public uExpInputSuggest
	Public CreateHtmlResultFilePath
	Public vCurTestName
	Public cHTMLFileName
	Public vTestStepNo
	Public sScriptExecutionStatus
	Public sFail
	Public lScenarioStatus
	
'########################################################################################
   	Private Sub Class_Initialize ' this will initialize the class
		ReportSheetName="eReport"
		Set ReportSheet=DataTable.AddSheet("eReport")
		Set sName=ReportSheet.AddParameter("StepName","")
		Set sDes=ReportSheet.AddParameter("Description","")
		Set sStatus=ReportSheet.AddParameter("Status","")
		Set sScreen=ReportSheet.AddParameter("ScreenShot","")
		
		clsTestCaseID=Environment("TestName")
		iIndex=1
		sFail=0
		eLogFilePath=CreateLogFile
        ExecutionEnvironment=GetExecutionEnvironment()
	End Sub
'########################################################################################


'***************************************************************

''help
'Elseif sEventStatus="2" Then
			'sEventStatus="DONE"
			

''*************************************************************
	Sub ReportEvent(sEventStatus,sReportStepName,sDetails,bExitRun)

		If sEventStatus="0" Then
			sEventStatus="PASS"
		Elseif sEventStatus="1" Then
			sEventStatus="FAIL"
		Elseif sEventStatus="2" Then
			sEventStatus="DONE"
		Elseif sEventStatus="4" Then
			sEventStatus="INFO"
		Elseif sEventStatus="5" Then
			sEventStatus="PASS"
'		Elseif sEventStatus="5" Then
'			sEventStatus="PASSCAPTURE"
		End If

		
		ReportSheet.SetCurrentRow(iIndex)
		sName.ValuebyRow(iIndex)=sReportStepName
		sDes.ValuebyRow(iIndex)=sDetails
		sStatus.ValuebyRow(iIndex)=sEventStatus

		If ucase(sEventStatus)="FAIL" or ucase(sEventStatus)="PASS" Then
			eImageFilePath=CreateImageFilePath

			If BrowserCreationTimeIndex="" Then
				BrowserCreationTimeIndex=0
			End If
			'If Browser("creationtime:="&BrowserCreationTimeIndex).Exist(2) Then
			'	Browser("creationtime:="&BrowserCreationTimeIndex).CaptureBitmap eImageFilePath,true
			'	Else
				Desktop.CaptureBitmap eImageFilePath,true	
			'End If
			
		End If
		sScreen.ValuebyRow(iIndex)=eImageFilePath

		If TrackStatus=true Then
			If lGroupStatus="" then
				lGroupStatus=UCase(sEventStatus) 
			ElseIf	lGroupStatus= "PASS" and UCase(sEventStatus) = "PASS" then
				lGroupStatus= "PASS"
			ElseIf	lGroupStatus= "PASS" and UCase(sEventStatus) = "FAIL" then
				lGroupStatus= "FAIL"
			End If
		End If
		
		If TrackStatus=true Then
			If lScenarioStatus="" then
				lScenarioStatus=UCase(sEventStatus) 
			ElseIf	lScenarioStatus= "PASS" and UCase(sEventStatus) = "PASS" then
				lScenarioStatus= "PASS"
			ElseIf	lScenarioStatus= "PASS" and UCase(sEventStatus) = "FAIL" then
				lScenarioStatus= "FAIL"
			ElseIf	lScenarioStatus= "DONE" and UCase(sEventStatus) = "FAIL" then
				lScenarioStatus= "FAIL"
			ElseIf	lScenarioStatus= "FAIL" and UCase(sEventStatus) = "FAIL" then
				lScenarioStatus= "FAIL"
			End If
		End If

		If UCase(sEventStatus) = "PASS" or UCase(sEventStatus) = "PASSCAPTURE" Then 
			QTPReporter.ReportEvent MicPass,sReportStepName,sDetails
		ElseIf UCase(sEventStatus) = "FAIL" Then
			QTPReporter.ReportEvent MicFail,sReportStepName,sDetails
			If bExitRun = True Or bExitRun = 1 Then
				clsErrDescription=Err.Description
				'CloseWindow ""
				'call KillProcessJunifer 
			End If 
		End If 
		
	iIndex=iIndex+1
	eImageFilePath=""
	End Sub
'########################################################################################
	Function QTPRunStatus()
	
			If QTPReporter.RunStatus=micPass Then
				QTPRunStatus="PASS"
			ElseIf QTPReporter.RunStatus=micFail And LCase(uExpInputSuggest)="pass" Then
				QTPRunStatus="PASS"
			ElseIf QTPReporter.RunStatus=micFail And LCase(uExpInputSuggest)<>"pass" Then
				QTPRunStatus="FAIL"
			ElseIf QTPReporter.RunStatus=micDone Then
				QTPRunStatus="DONE"
			ElseIf QTPReporter.RunStatus=micWarning Then
				QTPRunStatus="WARNING"
			End If
	End Function
'########################################################################################

''******************************************

'help




''*******************************************
	Function BuildFolderPath(byval fldPath)
		Dim bFso
		Dim PathArray
		Dim rfldPath
		Dim pIndex
		Dim fPath
		
		Set bFso=CreateObject("scripting.filesystemobject")
	
		If instr(1,fldPath,"\\")=1 Then
			rfldPath=mid(fldPath,3,len(fldPath)-2)
			PathArray=Split(rfldPath,"\")
			fPath="\\"&PathArray(0)
		Else
			PathArray=Split(fldPath,"\")
			fPath=PathArray(0)
		End If
	
			For pIndex=1 to ubound(PathArray)
		
			fPath=fPath&"\"&PathArray(pIndex)
		
				If not bFso.FolderExists(fPath) then
					bFso.CreateFolder(fPath)
				End If

			Next
		
		Set bFso=nothing
	
	End Function
'########################################################################################

''''************************************************



'help


''***********************************************
	Function CreateLogFile()
		cErrLogName = clsTestCaseID&"-"&Replace(Replace(Replace(Now,"/",""),":","")," ","-")
		cErrLogFolderName=clsTestCaseID&"-"&replace(date,"/","_")
'		cErrLogFolderPath=TEMP_DIR&"\"&cErrLogName
		cErrLogFolderPath=TEMP_DIR&cErrLogFolderName
		BuildFolderPath cErrLogFolderPath
		CreateLogFile=cErrLogFolderPath&"\"&cErrLogName&".txt"
	End Function
'########################################################################################

''***********************************************************



'help


''***********************************************************
	Function CreateImageFilePath()
			CreateImageFilePath=cErrLogFolderPath&"\"&cErrLogName&"_Step"&iIndex&" Error"&".png"

	End Function
'########################################################################################

''************************************






''*************************************
	Function CreateHtmlResultFile()
			    
		Set hFso=CreateObject("scripting.filesystemobject")
		Set oDataFolder = hFso.GetFolder (PROJECT_DIR&"HtmlResults")
		Set oDataFile=oDataFolder.Files
		Existfiles = oDataFile.Count
		
		cHTMLFileName=Environment("TestName")&"_"&vCurTestName & replace(date,"/","_")&"_"&replace(time,":","_")&".html"
		tHtmlResultFilePath=PROJECT_DIR&"HtmlResults\"&cHTMLFileName
				
				CreateHtmlResultFile=tHtmlResultFilePath
					
				
	End Function
'########################################################################################

'''******************************************************






''*****************************************************
	Function GenerateHTMLFile()
	
	dtRowCount=ReportSheet.GetRowCount
	If dtRowCount <> vTestStepNo-1 Then
		
		Set hFso=CreateObject("scripting.filesystemobject")
		tHtmlResultFilePath=CreateHtmlResultFile
		Set objHtmlFile=hFso.CreateTextFile(tHtmlResultFilePath)

		clsResultName=clsTestCaseID
		clsUser=Environment("UserName")
		tExecutionTimeZone=GetTimeZone
		tExecutionEndTime=Now
		tExecutionTime=ScriptExecutionTimeInSeconds(tExecutionStartTime,tExecutionEndTime)

		'Write HTML Header
		objHtmlFile.WriteLine ("<html><body  bgcolor=	#E6F0B8>")
		objHtmlFile.WriteLine ("<table align=center width=900 style=""font-family: Georgia, Arial;"" border=0><tr><td align=center bgcolor=#4A9586><font color=white><b>"&clsResultName&" Test Results</b></font></td></tr></table>")
		objHtmlFile.WriteLine ("<br>")
		objHtmlFile.WriteLine ("<table align=center width=900 border=0>")
		'Row1 (Test ScriptId, Start Time)
		objHtmlFile.WriteLine ("<tr><td bgcolor=#4A9586 ><font color=white><b>Test Script ID:</b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#D6C485><font color=#000000> "&clsTestCaseID&"</font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586 ><font color=white><b>Start Time</b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#D6C485><font color=#000000>"&tExecutionStartTime&"</font></td></tr>")
		''objHtmlFile.WriteLine ("<td bgcolor=#D6C485><font color=#000000>"&tExecutionStartTime&" "& tExecutionTimeZone&"</font></td></tr>")
		'Row2 (Environment, End Time)
		objHtmlFile.WriteLine ("<tr><td bgcolor=#4A9586 ><font color=white><b>Environment</b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#D6C485><font color=#000000>" &ExecutionEnvironment&"</font></td> ") 
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586 ><font color=white><b>End Time </b></font></td>")
		objHtmlFile.WriteLine ("<td id=""endtime"" bgcolor=#D6C485><font color=#000000>"&tExecutionEndTime&" </font></td></tr>")
		'Row3 (Module Name, Execution Time)
		objHtmlFile.WriteLine ("<tr><td bgcolor=#4A9586 ><font color=white><b>Module Name </b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#D6C485><font color=#000000>"&MODULE_NAME&"</font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586 ><font color=white><b>Execution Time </b></font></td>")
		objHtmlFile.WriteLine ("<td id=""etime"" bgcolor=#D6C485><font color=#000000>"&tExecutionTime&" Sec</font></td></tr>")
		'Row4 (Project Name, Execution Status)
		objHtmlFile.WriteLine ("<tr><td bgcolor=#4A9586 ><font color=white><b>Project Name</b></font></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#D6C485><font color=#000000>IP Tool</font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586 ><font color=white><b>Execution Status</b></font></font></td>")
		
			Select case sScriptExecutionStatus
				case "WARNING"
					fColor="FF9900"
				Case "DONE"
					fColor="#000000"
				case "FAIL"
					fColor="CC0000"
				case "PASS"
					fColor="#009900"
			End Select
		'objHtmlFile.WriteLine ("<td id=""status"" bgcolor=#D6C485><font color="&fColor&"><b>"&QTPRunStatus&"</b></font></td></tr>")
		objHtmlFile.WriteLine ("<td id=""status"" bgcolor=#D6C485><font color="&fColor&"><b>"&sScriptExecutionStatus&"</b></font></td></tr>")
		
		'Row5 (Machine Name,Browser Version)
		objHtmlFile.WriteLine ("<tr><td bgcolor=#4A9586 ><font color=white><b>OS Name </b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#D6C485><font color=#000000>"&GetMachineVersion&"</font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586 ><font color=white><b>Browser </b></font></td>")
		objHtmlFile.WriteLine ("<td id=""etime"" bgcolor=#D6C485><font color=#000000>"&GetBrowserVersion&" </font></td></tr></table><p> </p> ")
		
		'Result Columns
		objHtmlFile.WriteLine ("<table align=center  id="&chr(34)&clsTestCaseID&chr(34)&" style=""font-family: Georgia, Arial;"" width= 900>")
		objHtmlFile.WriteLine ("<tr bgcolor=#4A9586>")
		objHtmlFile.WriteLine ("<td><font color=white><b>S.No</b></font></td>")
		objHtmlFile.WriteLine ("<td><font color=white><b>Step Name</b></font></td>")
		objHtmlFile.WriteLine ("<td><font color=white><b>Description</b></font></td>")
		objHtmlFile.WriteLine ("<td><font color=white><b>Status</b></font></td>")
		objHtmlFile.WriteLine ("<td><font color=white><b>ScreenShot</b></font></td></tr>")
		
		dtRowCount=ReportSheet.GetRowCount
		dtColumnCount=ReportSheet.GetParameterCount
		
			If  vTestStepNo = " " or vTestStepNo = 0 Then
				vTestStepNo = 1
			End If
		
		For dtIndex=vTestStepNo to dtRowCount

			ReportSheet.SetCurrentRow(dtIndex)
			dtStepName=sName.ValuebyRow(dtIndex)
			dtDescription=sDes.ValuebyRow(dtIndex)
			dtStatus=UCase(sStatus.ValuebyRow(dtIndex))
			dtScreen=sScreen.ValuebyRow(dtIndex)

		
				Select case dtStatus
					case "WARNING"
						reportColor="FF9900"
					Case "DONE"
						reportColor="#000000"
					case "FAIL"
						reportColor="CC0000"
					case "PASS"
						reportColor="#009900"
					Case "INFO"
						reportColor="CCCCCC"
				End Select
		If InStr(1,dtDescription,"Began")=1 Then

			HTMLStepStart="<TR bgColor=#99CCFF>"
			HTMLStepIndex="<td align='center' ><font color=#7A297A>"&dtIndex&"</font></td>"
			HTMLStepName="<td ><font color=#7A297A>"&dtStepName&"</font></td>"
			HTMLStepDescription="<td ><font color=#7A297A>"&dtDescription&"</font></td>"
			HTMLStepStatus="<td ><font color=#7A297A>"&dtStatus&"</font></td>"
			HTMLStepScreenPath="<td><font color=#7A297A>-</font></td>"
		Else
			HTMLStepStart="<TR bgColor=#D6C485>"
			HTMLStepIndex="<td align='center' bgcolor='#D6C485'><font color="&reportColor&">"&dtIndex&"</font></td>"
			HTMLStepName="<td bgcolor='#D6C485'><font color="&reportColor&">"&dtStepName&"</font></td>"
			HTMLStepDescription="<td bgcolor=#D6C485><font color="&reportColor&">"&dtDescription&"</font></td>"
			HTMLStepStatus="<td bgcolor=#D6C485><font color=	"&reportColor&">"&dtStatus&"</font></td>"
	
			If dtScreen<>"" Then
				
				RelImgPath=replace(dtScreen,PROJECT_DIR,"file:../")
				HTMLStepScreenPath="<td bgcolor=#D6C485><a href='"&dtScreen&"' target=""_blank"">View Screenshot</a></td>"
			else
				HTMLStepScreenPath="<td bgcolor=#D6C485>-</td>"
			End If
	
			HTMLStepEnd="</TR>"
		
		End If

		HTMLStep=HTMLStepStart&HTMLStepIndex&HTMLStepName&HTMLStepDescription&HTMLStepStatus&HTMLStepScreenPath&HTMLStepEnd
		objHtmlFile.WriteLine(HTMLStep)
		Next
		vTestStepNo = dtRowCount+1
	    	outFile=PROJECT_DIR&"\HtmlResults\HghLelRptSts.txt"
			Set objFSO=CreateObject("Scripting.FileSystemObject")
'			Set objFile = objFSO.OpenTextFile(outFile,8,True)
'			objFile.Writeline clsTestCaseID &","& tExecutionEndTime &","& vCurTestName &","& sScriptExecutionStatus &","& tHtmlResultFilePath
'			objFile.Close
	    
			
		objHtmlFile.WriteLine ("</table></body></html>")
		objHtmlFile.Close					
	
	End If
					
	End Function
''########################################################################################
'Function SetStartIndex(xindex)
'
'	vTestStepNo = xindex
'	
'End Function
'########################################################################################
Function getpiechart()

SystemUtil.CloseProcessByName "EXCEL.EXE"
Const xlDataLabelsShowPercent = 3
Const xlHtml = 44

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1,1) = "No. of Test cases"
objWorksheet.Cells(2,1) = "No. of Passed cases"
objWorksheet.Cells(3,1) = "No. of Failed  cases"

    	i = 0
		p = 0
		f = 0
		
		outFile1=PROJECT_DIR&"HtmlResults\HighLevelReports\HghLelRptSts.txt"
		Set objFSO1=CreateObject("Scripting.FileSystemObject")
		Set objFile1 = objFSO1.OpenTextFile(outFile1,1,True)
		
		do until objFile1.AtEndOfStream
		    strLine= objFile1.ReadLine()
		    strline1 = split(strLine,",")
			tcName=strline1(0)
			stus=strline1(1)
			refLinkRep = strline1(2)
			i=i+1
			
		    If stus = "PASS" Then
		     p=p+1
		    Else
		     f=f+1
	        End IF	
	        
			dtIndex=i
		Loop
objWorksheet.Cells(1,2) = "Execution Details"
objWorksheet.Cells(2,2) = p
objWorksheet.Cells(3,2) = f
Set objRange = objWorksheet.UsedRange
objRange.Select
Set colCharts = objExcel.Charts
colCharts.Add()

Set objChart = colCharts(1)
objChart.Activate

objChart.ChartType = 70
''objChart.Elevation = 30
''objChart.Rotation = 80
''objExcel.ActiveChart.Width = 300

objChart.ApplyDataLabels xlDataLabelsShowPercent

objChart.PlotArea.Fill.Visible = False
objChart.PlotArea.Border.LineStyle = -4142

objChart.SeriesCollection(1).DataLabels.Font.Size = 14
objChart.SeriesCollection(1).DataLabels.Font.ColorIndex = 2

objChart.ChartArea.Fill.ForeColor.SchemeColor = 49
objChart.ChartArea.Fill.BackColor.SchemeColor = 23
objChart.ChartArea.Fill.TwoColorGradient 1,1

objChart.ChartTitle.Font.Size = 24
objChart.ChartTitle.Font.ColorIndex = 2

objChart.Legend.Shadow = True

	Set hFso=CreateObject("scripting.filesystemobject")
		
	''	Set hFso=CreateObject("scripting.filesystemobject")
		Set oDataFolder1 = hFso.GetFolder (PROJECT_DIR&"HtmlResults\HighLevelReports")
		Set oDataFile=oDataFolder1.Files
		Existfiles1 = oDataFile.Count
If  hFso.FileExists(PROJECT_DIR&"HtmlResults\\HighLevelReports\Piechart.html") Then
hFso.MoveFile PROJECT_DIR&"HtmlResults\HighLevelReports\Piechart.html",PROJECT_DIR&"HtmlResults\HighLevelReports\Piechart"&Existfiles1&".html"
End If
cHTMLFileName1="Piechart"&".html"	
''tHtmlResultFilePath=cResFolderPath&"\"&cHTMLFileName
tHtmlResultFilePath1=PROJECT_DIR&"HtmlResults\HighLevelReports\"&cHTMLFileName1

objExcel.ActiveWorkbook.SaveAs tHtmlResultFilePath1,  xlHtml
''objExcel.ActiveWorkbook.SaveAs PROJECT_DIR&"HtmlResults\HighLevelReports\piechartex.png",  xlHtml

End Function
'########################################################################################

	
	'************************************************
	
	'help
	
	
	
	
	'***********************************************
	Function WritingTcStatuInTheTextFile()
	
	cResFolderPath=PROJECT_DIR&"HtmlResults"
	
		Set hFso=CreateObject("scripting.filesystemobject")
		Set oDataFolder = hFso.GetFolder (cResFolderPath)
		Set oDataFile=oDataFolder.Files
		Existfiles = oDataFile.Count
				
'		BuildFolderPath cResFolderPath
		If  hFso.FileExists(PROJECT_DIR&"\HtmlResults\HighLevelReport.html") Then
		      hFso.MoveFile PROJECT_DIR&"\HtmlResults\HighLevelReport.html",PROJECT_DIR&"\HtmlResults\HighLevelReport"&Existfiles&".html"
		End If
		cHTMLFileName="HighLevelReport"&".html"	
		tHtmlResultFilePath=PROJECT_DIR&"\HtmlResults\"&cHTMLFileName
        
		Set objHtmlFile=hFso.CreateTextFile(tHtmlResultFilePath)

		clsResultName=clsTestCaseID
		clsUser=Environment("UserName")
		tExecutionTimeZone=GetTimeZone
		tExecutionEndTime=Now
		tExecutionTime=ScriptExecutionTimeInSeconds(tExecutionStartTime,tExecutionEndTime)
		objHtmlFile.WriteLine ("<html>")		
		objHtmlFile.WriteLine ("<head><script type=""text/javascript"">")
		objHtmlFile.WriteLine ("    function showhide(id) {")
        objHtmlFile.WriteLine ("var vDivs = document.getElementsByTagName('div');")
        objHtmlFile.WriteLine ("if (vDivs != null && vDivs.length > 0) {")
        objHtmlFile.WriteLine ("for (var i = 0; i < vDivs.length; i++) {")
        objHtmlFile.WriteLine ("var e = document.getElementById(vDivs[i].id);")
        objHtmlFile.WriteLine ("e.style.display = (e.id == id) ? 'block' : 'none'; } } }")
        objHtmlFile.WriteLine ("</script>")
        objHtmlFile.WriteLine ("<title>Prolifics Automation Report</title></head>")
        
        ''Logo part
        objHtmlFile.WriteLine ("<table align=""center"" border=2 bordercolor=#000000 id=table1 width=1200 height=31 bordercolorlight=#000012>")
        objHtmlFile.WriteLine ("<tr><td COLSPAN =3 bgcolor = #45AEE3>")
    		objHtmlFile.WriteLine ("<p><img src="&Chr(34)&PROJECT_DIR &"Data\ClientLogo.PNG" &Chr(34)&" align =left><img src="&Chr(34)&PROJECT_DIR &"Data\ClientLogo.PNG" &Chr(34)&" align =right></p><p align=center><font color=""white"" size=6 face= ""Copperplate Gothic Bold"">&nbsp;IP Tool Automation Execution Summary </font><font face= ""Copperplate Gothic Bold""></font> </p>")
    		objHtmlFile.WriteLine ("</td></tr></table>")
    	
    	'An Empty table for space
    	objHtmlFile.WriteLine ("<table align=""center"" width=""1200"" style=""font-family: Georgia, Arial;"" border=""0"">")
        objHtmlFile.WriteLine ("<tr><td align=""center"" bgcolor=""white"">")
        objHtmlFile.WriteLine ("<font color=""white""><b></b></font>")
        objHtmlFile.WriteLine ("</td></tr></table>")
        
        ''Module Name table
        objHtmlFile.WriteLine ("<table align=""center"" width=""1200"" border=""0"">")
        objHtmlFile.WriteLine ("<tr><td bgcolor=""#45AEE3""><font color=""white""><b>Module Name</b></font></td>")
        objHtmlFile.WriteLine ("<td bgcolor=""#45AEE3""><font color=""white""><b>Total Test Executed</b></font></td>")
        objHtmlFile.WriteLine ("<td bgcolor=""#45AEE3""><font color=""white""><b>No.of  PASS</font></td>")
        objHtmlFile.WriteLine ("<td bgcolor=""#45AEE3""><font color=""white""><b>No.of  FAIL</b></font></td>")

		sModuleNames = ""
		sTotalTCCnt = 0
		sPassTCCnt = 0
		sFailTCCnt = 0	
		outFile1=PROJECT_DIR&"HtmlResults\HghLelRptSts.txt"
		Set objFSO1=CreateObject("Scripting.FileSystemObject")
		Set objFile1 = objFSO1.OpenTextFile(outFile1,1,True)

	do until objFile1.AtEndOfStream
		strLine= objFile1.ReadLine()
		If not strLine = ""  Then	
				    strline1 = split(strLine,",")
				    moduleName=strline1(0)
					tcName=strline1(2)
					stus=strline1(3)
					refLinkRep = strline1(4)
			if objFSO1.FileExists(refLinkRep) Then
				If instr(1, sModuleNames,moduleName)  Then
						sTotalTCCnt=sTotalTCCnt+1					
				    	If stus = "PASS" Then
				    		sPassTCCnt=sPassTCCnt+1
				   		 Else
				     		sFailTCCnt=sFailTCCnt+1
			        	End IF
			        Else
				        If OldmoduleName<>"" Then
							objHtmlFile.WriteLine ("<tr bgcolor=""#99CCFF""><td><font color=""#7A297A"">"& OldmoduleName &"</font></td>")
					        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sTotalTCCnt&"</font></td>")
					        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sPassTCCnt&"</font></td>")
					        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sFailTCCnt&"</font></td>")
				        End If
			        sModuleNames =moduleName & ", "
			        sTotalTCCnt=1
			        	If stus = "PASS" Then
				    		sPassTCCnt=1
				    		sFailTCCnt=0
				   		 Else
				     		sFailTCCnt=1
				     		sPassTCCnt=0
			        	End IF			
				End If
					OldmoduleName = moduleName
			End if
		End If
	Loop
		objFile1.Close
			
	objHtmlFile.WriteLine ("<tr bgcolor=""#99CCFF""><td><font color=""#7A297A"">"& OldmoduleName &"</font></td>")
        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sTotalTCCnt&"</font></td>")
        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sPassTCCnt&"</font></td>")
        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sFailTCCnt&"</font></td>")
    	
    	objHtmlFile.WriteLine ("<body  bgcolor=	white>")
        objHtmlFile.WriteLine ("<table align=""center"" width=""900"" style=""font-family: Georgia, Arial;"" border=""0"">")
        objHtmlFile.WriteLine ("<tr><td align=""center"" bgcolor=""#4A9586""><font color=white><b>Automation Script Execution Status</b></font></td></tr></table>")		
      
		objHtmlFile.WriteLine (" <table align=""center"" width=""900"" border=""0"">")
		objHtmlFile.WriteLine ("<tr><td bgcolor=""#4A9586"" ><font color=white><b>S.No</b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586><font color=white><b>Date&Time</b></font></td>") 
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586><font color=white><b>ScriptName</b></font></td>") 
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586 ><font color=white><b>Status</b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586><font color=white><b>Results</b></font></td></tr>")	

		i = 0
			
		outFile2=PROJECT_DIR&"HtmlResults\HghLelRptSts.txt"
		Set objFSO2=CreateObject("Scripting.FileSystemObject")
		Set objFile2 = objFSO2.OpenTextFile(outFile2,1,True)
		
		do until objFile2.AtEndOfStream
		    strLine= objFile2.ReadLine()
			If not strLine = ""  Then	
		    strline1 = split(strLine,",")
			'MsgBox "Reading hello world:"&strline1(0)&strline1(1)&strline1(2)&strline1(3)
			ModName=strline1(0)
			tcEndTime = strline1(1)
			tcName=strline1(2)
			stus=strline1(3)
			refLinkRep = strline1(4)
			i=i+1
			dtIndex=i
			
			Set objFSO3=CreateObject("Scripting.FileSystemObject")
			Set objReadFile = objFSO3.OpenTextFile(refLinkRep, 1, False)

			'Read file contents
			contents = objReadFile.ReadAll
			
			
			'Close file
			objReadFile.Close
			Set objReadFile = Nothing
			Set objFSO3 = Nothing 
		
			objHtmlFile.WriteLine ("<TR bgColor=#99CCFF>")
			objHtmlFile.WriteLine ("<td align='center' ><font color=#7A297A>"&dtIndex&"</font></td>")
			objHtmlFile.WriteLine ("<td ><font color=#7A297A>"&tcEndTime&"</font></td>")
			objHtmlFile.WriteLine ("<td ><font color=#7A297A>"&tcName&"</font></td>")
			If ucase(stus)="PASS" Then
				objHtmlFile.WriteLine ("<td ><font color=#268147><b>"&stus&"</b></font></td>")
				Else
				objHtmlFile.WriteLine ("<td ><font color=#FF5733><b>"&stus&"</b></font></td>")
			End If	
			objHtmlFile.WriteLine ("<td bgcolor=#99CCFF><a href=""javascript:showhide('Div"&dtIndex&"')"">View Results</a></td></tr>")
	        objHtmlFile.WriteLine ("<tr><td colspan =""5""><Div id='Div"&dtIndex&"' style=""display:none;;border-style:solid;border-width:thin;border-color:Navy"">")
	        objHtmlFile.WriteLine (contents)
		objHtmlFile.WriteLine ("</div> </td></tr>")
		End if
		Loop
	
		
		objFile2.Close
		objHtmlFile.WriteLine ("</table></body></html>")
		objHtmlFile.Close
	End Function
	
	'########################################################################################
	''******************************************************
	
	
	'help
	
	
	
	''*********************************************************
Function DeleteFile()
	If filesys.FileExists(PROJECT_DIR&"HtmlResults\HghLelRptSts.txt") Then
       filesys.DeleteFile PROJECT_DIR&"HtmlResults\HghLelRptSts.txt"
End If 
   End Function

'########################################################################################


''***************************************************

'help




'''****************************************************
	Function GetTimeZone()
	
		Set ws=CreateObject("WScript.Shell")
	    
	    If InStr(GetMachineVersion(),"7")>0 Then
	        IntVersion=ws.RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\TimeZoneInformation\TimeZoneKeyName")
	    Else
	        IntVersion=ws.RegRead ("HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\TimeZoneInformation\StandardName")
	    End If
	    
	    If IntVersion="Pacific Standard Time" Then
	        GetTimeZone="PST"
	    ElseIf IntVersion="India Standard Time" Then
	        GetTimeZone="IST"
	    Else
	        GetTimeZone=IntVersion
	    End If  
	   Set ws=Nothing
   End Function

'########################################################################################

'''************************************************************

'help




''****************************************************************
	Function ScriptExecutionTimeInSeconds(StartTime,EndTime)
		StartHour = Hour(StartTime)
		StartMin = Minute(StartTime)
		StartSec = Second(StartTime)
		EndHour = Hour(EndTime)
		EndMin = Minute(EndTime)
		EndSec = Second(EndTime)
		
		StartingSeconds = (StartSec + (StartMin * 60) + (StartHour * 3600))
		EndingSeconds = (EndSec + (EndMin * 60) + (EndHour * 3600))
	
		ScriptExecutionTimeInSeconds = EndingSeconds - StartingSeconds
	End Function
'########################################################################################
	
	''**************************************************
	
	''Help
	
	
	
	
	''*******************************************************
	Function AppendDataToLogFile()

	Dim lFso
	' Get instance of FileSystemObject.
	Set lFso = CreateObject("Scripting.FileSystemObject")
	Set lMyFile = lFso.CreateTextFile (eLogFilePath, True)
	''MsgBox "eLogFilePath" & eLogFilePath
	lMyFile.WriteLine("*************************"&clsTestCaseID&"*************************")

	dtRowCount=ReportSheet.GetRowCount
	dtColumnCount=ReportSheet.GetParameterCount
	
	For dtIndex=1 to dtRowCount

		ReportSheet.SetCurrentRow(dtIndex)
		dtStepName=sName.ValuebyRow(dtIndex)
		dtDescription=sDes.ValuebyRow(dtIndex)
		dtStatus=sStatus.ValuebyRow(dtIndex)
		dtScreen=sScreen.ValuebyRow(dtIndex)
		lMyFile.WriteLine ("Step-"& CInt(dtIndex) &vbTab & dtStatus&vbTab & dtDescription & vbTab & Time )

		If UCase(dtStatus) = "FAIL" Or clsErrDescription<>"" Then
			If dtStatus="FAIL" Then
				sFail=sFail+1
			End If
			If CurrentFunctionName<>"" Then
				ErrFunction=vbnewline&"Error at Function: "&CurrentFunctionName
				lMyFile.WriteLine ErrFunction
			ElseIf clsErrDescription<>"" Then
				lMyFile.WriteLine clsErrDescription			
			End If
		End If
		If sFail>0 Then
				sScriptExecutionStatus="FAIL"
		Else
				sScriptExecutionStatus="PASS"
		End If
		If Instr(dtDescription,"Began of Scenario")=1 Then
				sFail=0
				lScenarioStatus = ""
				
		End If

	Next

	lMyFile.Close

	Set lFso=Nothing
	Set lMyFile=Nothing

	End Function
'########################################################################################

	Function StoreTestdata(eData)
		If not isobject(eData) Then
			   eArray=Split(eData,",")
				For each edValue in eArray
					If  ExecutionData<>"" Then
						ExecutionData=ExecutionData&vbnewline&edValue
					Else
						ExecutionData=edValue
					End If
				Next
		Else
	
			dKeys=eData.Keys
			dItems=eData.Items
	
			For dItemCount=0 to eData.Count-1
					If  ExecutionData<>"" Then
						ExecutionData=ExecutionData&vbnewline&dKeys(dItemCount)&":"&dItems(dItemCount)
					Else
						ExecutionData=dKeys(dItemCount)&":"&dItems(dItemCount)
					End If
			Next
				
		End If
	
	End Function
'########################################################################################

''*************************************************


''Help



''****************************************************
Function WriteTestResultsIntoExcel ()
	SystemUtil.CloseProcessByName "EXCEL.EXE"	
	Set fs=CreateObject("Scripting.FileSystemObject")
	Set objExcel=CreateObject ("Excel.Application")
	bFound = False

	If (fs.FileExists(tExcelFilePath)) Then
        Set objWorkBook=objExcel.Workbooks.Open(tExcelFilePath)
		sSheetName = replace(date,"/","_")

		For i = 1 to objWorkBook.Worksheets.Count

			If  objWorkBook.Worksheets(i).Name = sSheetName Then

				Set sNewSheet=objWorkBook.Worksheets(sSheetName)
				bFound = True
				Exit for
				
            End If
		Next
		If bFound = False Then
			Set sNewSheet=objWorkBook.Worksheets.Add
			sNewSheet.Name=sSheetName
		End If
	Else
		Set objWorkBook=objExcel.Workbooks.Add
		Set sNewSheet=objWorkBook.Worksheets.Add
		sNewSheet.Name=replace(date,"/","_")
		objWorkBook.SaveAs tExcelFilePath
	End If
'======	
	Set objSheet=objExcel.Worksheets(sSheetName)
	aa.DisplayAlerts=False
	rc=sNewSheet.UsedRange.Rows.count
	
		Dim MyDate
		MyDate = Date   
		DateTime = Date & "  "&  Time
		
		rc=rc+1
		
		sNewSheet.Cells(1,1) = "Script ID"
		sNewSheet.Cells(1,1).Font.Bold = True
		sNewSheet.Cells(1,1).Font.ColorIndex = 1
		sNewSheet.Cells(1,1).Interior.ColorIndex=48
		
		
		sNewSheet.Cells(1,2) = "Priority"
		sNewSheet.Cells(1,2).Font.Bold = True
		sNewSheet.Cells(1,2).Font.ColorIndex = 1
		sNewSheet.Cells(1,2).Interior.ColorIndex=48
		
		sNewSheet.Cells(1,3) = "Run Status"
		sNewSheet.Cells(1,3).Font.Bold = True
		sNewSheet.Cells(1,3).Font.ColorIndex = 1
		sNewSheet.Cells(1,3).Interior.ColorIndex=48
		
		sNewSheet.Cells(1,4) = "Execution Time"
		sNewSheet.Cells(1,4).Font.Bold = True
		sNewSheet.Cells(1,4).Font.ColorIndex = 1
		sNewSheet.Cells(1,4).Interior.ColorIndex=48
	
		sNewSheet.Cells(1,5) = "DateTime"
		sNewSheet.Cells(1,5).Font.Bold = True
		sNewSheet.Cells(1,5).Font.ColorIndex = 1
		sNewSheet.Cells(1,5).Interior.ColorIndex=48
		sNewSheet.Cells(1,5).Interior.ColorIndex=48
	
		sNewSheet.Cells(1,6) = "TestData"
		sNewSheet.Cells(1,6).Font.Bold = True
		sNewSheet.Cells(1,6).Font.ColorIndex = 1
		sNewSheet.Cells(1,6).Interior.ColorIndex=48
	
		sNewSheet.Cells(1,7) = "Execution Environment"
		sNewSheet.Cells(1,7).Font.Bold = True
		sNewSheet.Cells(1,7).Font.ColorIndex = 1
		sNewSheet.Cells(1,7).Interior.ColorIndex=48
	
		sNewSheet.Cells(1,8) = "DetailedResult"
		sNewSheet.Cells(1,8).Font.Bold = True
		sNewSheet.Cells(1,8).Font.ColorIndex = 1
		sNewSheet.Cells(1,8).Interior.ColorIndex=48
	
	 'Store Test Name in Excel
		sNewSheet.Cells(rc,1).value=Environment.Value("TestName")
	sNewSheet.Cells(rc, 1).Font.Bold = TRUE
		iStatus= QTPRunStatus

	'Specify Priority
		sNewSheet.Cells(rc,2).value=TestCasePriority

	 'Store Test status in Excel	
		If iStatus = "FAIL"  Then
			sNewSheet.Cells(rc,3).Font.Bold = TRUE
			'Font color to Red
			sNewSheet.Cells(rc, 2).Font.ColorIndex = 1
			sNewSheet.Cells(rc, 3).Interior.ColorIndex=3
			
			sNewSheet.Cells(rc,3).value=iStatus
		Else
			sNewSheet.Cells(rc,3).Font.Bold = True
'			Font green color
			sNewSheet.Cells(rc, 2).Font.ColorIndex = 4
'			Fill background color to green
			sNewSheet.Cells(rc, 3).Interior.ColorIndex=4
			sNewSheet.Cells(rc,3).value=iStatus
		End If
	
	 'Store Execution Time	
		sNewSheet.Cells(rc,4) = tExecutionTime
	
	 'Store Date and Time
		sNewSheet.Cells(rc,5) = DateTime
	
	 'Store Execution Data
		sNewSheet.Cells(rc,6) = ExecutionData
	
	 'Store Execution Environment
		sNewSheet.Cells(rc,7) = ExecutionEnvironment
	
'	 Specify Detailed Result Link
		sNewSheet.Cells(rc,8) = "View Result"
		sNewSheet.Cells(rc,8).select
		sNewSheet.Hyperlinks.Add objExcel.selection,tHtmlResultFilePath
	
		sNewSheet.Columns("A:A").EntireColumn.AutoFit
		sNewSheet.Columns("B:B").EntireColumn.AutoFit
		sNewSheet.Columns("C:C").EntireColumn.AutoFit
		sNewSheet.Columns("D:D").EntireColumn.AutoFit
		sNewSheet.Columns("E:E").EntireColumn.AutoFit
		sNewSheet.Columns("F:F").EntireColumn.AutoFit
		sNewSheet.Columns("G:G").EntireColumn.AutoFit
		sNewSheet.Columns("H:H").EntireColumn.AutoFit

		objWorkBook.Save
		objExcel.Quit		
		
	If Dialog("regexpwndtitle:=Microsoft Excel").WinButton("text:=&Yes").Exist(5) Then		
		SystemUtil.CloseProcessByName "EXCEL.EXE"
		Dialog("regexpwndtitle:=Microsoft Excel").WinButton("text:=&Yes").Highlight
		Dialog("regexpwndtitle:=Microsoft Excel").WinButton("text:=&Yes").Click
	End If
			
		Set objExcel=Nothing
		Set objSheet = Nothing
		Set sSheetName = Nothing
		Set objWorkBook= nothing
	End Function
'########################################################################################
	Function GetExecutionEnvironment()
	   EnvArray=split(EXECUTION_ENVIRONMENT,"-")
		GetExecutionEnvironment= EnvArray(ubound(EnvArray))
	End Function
'########################################################################################
	Function StartStatusTrack()
		TrackStatus=True
	End Function
'########################################################################################
	Function EndStatusTrack()
		TrackStatus=False
	End Function
	
'########################################################################################
	Function StartTest(vCTestName)
	
	vCurTestName = vCTestName

	End Function	
'########################################################################################
	Function EndTest()
		 lScenarioStatus= "DONE"
		ExitEachTest
	End Function
	'########################################################################################
Private Sub ExitEachTest
		AppendDataToLogFile()
			
		If UCase(sScriptExecutionStatus) = "FAIL" Or clsErrDescription<>"" Then
			If CurrentFunctionName<>"" Then
				ErrFunction="Error at Function: "&CurrentFunctionName
				ReportEvent "Fail","Function Failure",ErrFunction,INDEX_VALUE_ZERO
				
			ElseIf clsErrDescription<>"" Then
				ReportEvent "Fail","Script Failure",clsErrDescription,INDEX_VALUE_ZERO
				
			ElseIf Err.Description<>"" Then
				ReportEvent "Fail","Script Failure",Err.Description,INDEX_VALUE_ZERO
			End If
		End If
		
		If UCase(GENERATE_HTML_REPORT)="TRUE" Then
			GenerateHTMLFile()
'			WritingTcStatuInTheTextFile()
			Rem getpiechart()
'			WriteTestResultsIntoExcel ()
		End If
		
'		DataTable.DeleteSheet "eReport"
	End Sub
'########################################################################################
	Private Sub Class_Terminate

	GenerateHTMLFile()
	WritingTcStatuInTheTextFile()	
	End Sub
'########################################################################################
Function GetBrowserVersion()

If BROWSER_NAME="IE" Then
   
    Set ws=CreateObject("WScript.Shell")
    IntVersion=ws.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")
    IntVersion= Mid(IntVersion,1,3)
    IntVersion=BROWSER_NAME&" "&IntVersion&".x"
    GetBrowserVersion=IntVersion
    Set ws=Nothing

ElseIf BROWSER_NAME="FF" Then

	Set ws=CreateObject("WScript.Shell")
	IntVersion=ws.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Mozilla\Mozilla Firefox\CurrentVersion")
	IntVersion= Mid(IntVersion,1,3)
	IntVersion=BROWSER_NAME&" "&IntVersion&".x"
	GetBrowserVersion=IntVersion
	Set ws=Nothing
End If

End Function
'########################################################################################
Function GetMachineVersion()

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
 
Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOS in colOSes
  
  If InStr(UCase(objOS.Caption),"XP")<>0 Then
  	OSName="WinXP"
  ElseIf InStr(UCase(objOS.Caption),"VISTA")<>0 Then
  	OSName="Win Vista"
  ElseIf InStr(UCase(objOS.Caption),"7")<>0 Then
  	OSName="Windows 7"
  Else
  	OSName=objOS.Caption
  End If
  
  spVersion=objOS.ServicePackMajorVersion
Next

OSName=OSName& " SP"& spVersion
GetMachineVersion=OSName

End Function
'########################################################################################
Function StartFunction(cFunName,oSendReport)
	CurrentFunctionName=cFunName
	If LCase(oSendReport)="true" Then
		ReportEvent "Done",cFunName, "Began of "& CurrentFunctionName,0
	End If
End Function
'########################################################################################
Function EndFunction
	CurrentFunctionName=""
End Function

'########################################################################################

End Class

	Function WritingTcStatuInTheTextFileMahesh()
	
	cResFolderPath=PROJECT_DIR&"HtmlResults"
	
		Set hFso=CreateObject("scripting.filesystemobject")
		Set oDataFolder = hFso.GetFolder (cResFolderPath)
		Set oDataFile=oDataFolder.Files
		Existfiles = oDataFile.Count
				
'		BuildFolderPath cResFolderPath
		If  hFso.FileExists(PROJECT_DIR&"\HtmlResults\HighLevelReport.html") Then
		      hFso.MoveFile PROJECT_DIR&"\HtmlResults\HighLevelReport.html",PROJECT_DIR&"\HtmlResults\HighLevelReport"&Existfiles&".html"
		End If
		cHTMLFileName="HighLevelReport"&".html"	
		tHtmlResultFilePath=PROJECT_DIR&"\HtmlResults\"&cHTMLFileName
        
		Set objHtmlFile=hFso.CreateTextFile(tHtmlResultFilePath)

		clsResultName=clsTestCaseID
		clsUser=Environment("UserName")
		tExecutionTimeZone=GetTimeZone
		tExecutionEndTime=Now
'		tExecutionTime=ScriptExecutionTimeInSeconds(tExecutionStartTime,tExecutionEndTime)
		objHtmlFile.WriteLine ("<html>")		
		objHtmlFile.WriteLine ("<head><script type=""text/javascript"">")
		objHtmlFile.WriteLine ("    function showhide(id) {")
        objHtmlFile.WriteLine ("var vDivs = document.getElementsByTagName('div');")
        objHtmlFile.WriteLine ("if (vDivs != null && vDivs.length > 0) {")
        objHtmlFile.WriteLine ("for (var i = 0; i < vDivs.length; i++) {")
        objHtmlFile.WriteLine ("var e = document.getElementById(vDivs[i].id);")
        objHtmlFile.WriteLine ("e.style.display = (e.id == id) ? 'block' : 'none'; } } }")
        objHtmlFile.WriteLine ("</script>")
        objHtmlFile.WriteLine ("<title>Prolifics Automation Report</title></head>")
        
        ''Logo part
        objHtmlFile.WriteLine ("<table align=""center"" border=2 bordercolor=#000000 id=table1 width=1200 height=31 bordercolorlight=#000012>")
        objHtmlFile.WriteLine ("<tr><td COLSPAN =3 bgcolor = #45AEE3>")
        objHtmlFile.WriteLine ("<p><img src="&PROJECT_DIR &"\Data\ProlificsLogo.bmp" & " align =left><img src="&Chr(34)&PROJECT_DIR &"\Data\ClientLogo.bmp" &Chr(34)&" align =right></p><p align=center><font color=""white"" size=6 face= ""Copperplate Gothic Bold"">&nbsp;Prolifics Automation Execution Summary </font><font face= ""Copperplate Gothic Bold""></font> </p>")
    	objHtmlFile.WriteLine ("</td></tr></table>")
    	
    	'An Empty table for space
    	objHtmlFile.WriteLine ("<table align=""center"" width=""1200"" style=""font-family: Georgia, Arial;"" border=""0"">")
        objHtmlFile.WriteLine ("<tr><td align=""center"" bgcolor=""white"">")
        objHtmlFile.WriteLine ("<font color=""white""><b></b></font>")
        objHtmlFile.WriteLine ("</td></tr></table>")
        
        ''Module Name table
        objHtmlFile.WriteLine ("<table align=""center"" width=""1200"" border=""0"">")
        objHtmlFile.WriteLine ("<tr><td bgcolor=""#45AEE3""><font color=""white""><b>Module Name</b></font></td>")
        objHtmlFile.WriteLine ("<td bgcolor=""#45AEE3""><font color=""white""><b>Total Test Executed</b></font></td>")
        objHtmlFile.WriteLine ("<td bgcolor=""#45AEE3""><font color=""white""><b>No.of Sub Modules PASS</font></td>")
        objHtmlFile.WriteLine ("<td bgcolor=""#45AEE3""><font color=""white""><b>No.of Sub Modules  FAIL</b></font></td>")
        objHtmlFile.WriteLine ("<td bgcolor=""#45AEE3""><font color=""white""><b>No.of test steps Executed</b></font></td></tr>")

		sModuleNames = ""
		sTotalTCCnt = 0
		sPassTCCnt = 0
		sFailTCCnt = 0	
		SpTotalStepCnt = 0
		outFile1=PROJECT_DIR&"HtmlResults\HghLelRptSts.txt"
		Set objFSO1=CreateObject("Scripting.FileSystemObject")
		Set objFile1 = objFSO1.OpenTextFile(outFile1,1,True)

	do until objFile1.AtEndOfStream
		strLine= objFile1.ReadLine()
		If not strLine = ""  Then	
				    strline1 = split(strLine,",")
				    moduleName=strline1(0)
					tcName=strline1(2)
					stus=strline1(3)
					refLinkRep = strline1(4)
					SpStepCnt = strline1(5)
			
			if objFSO1.FileExists(refLinkRep) Then
				If instr(1, sModuleNames,moduleName)  Then
						SpTotalStepCnt = SpTotalStepCnt+SpStepCnt
						sTotalTCCnt=sTotalTCCnt+1					
				    	If stus = "PASS" Then
				    		sPassTCCnt=sPassTCCnt+1
				   		 Else
				     		sFailTCCnt=sFailTCCnt+1
			        	End IF
			        Else
				        If OldmoduleName<>"" Then
							objHtmlFile.WriteLine ("<tr bgcolor=""#99CCFF""><td><font color=""#7A297A"">"& OldmoduleName &"</font></td>")
					        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sTotalTCCnt&"</font></td>")
					        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sPassTCCnt&"</font></td>")
					        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sFailTCCnt&"</font></td>")
					      '' MsgBox "upp"& SpTotalStepCnt
		                             objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&SpTotalStepCnt&"</font></td></tr>")
		                             SpTotalStepCnt=0
				        End If
			        sModuleNames =moduleName & ", "
			        SpTotalStepCnt = SpTotalStepCnt+SpStepCnt
			        sTotalTCCnt=1
			        	If stus = "PASS" Then
				    		sPassTCCnt=1
				    		sFailTCCnt=0
				   		 Else
				     		sFailTCCnt=1
				     		sPassTCCnt=0
			        	End IF			
				End If
					OldmoduleName = moduleName

			End if
		End If
	Loop
		objFile1.Close
			
	objHtmlFile.WriteLine ("<tr bgcolor=""#99CCFF""><td><font color=""#7A297A"">"& OldmoduleName &"</font></td>")
        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sTotalTCCnt&"</font></td>")
        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sPassTCCnt&"</font></td>")
        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&sFailTCCnt&"</font></td>")
        ''MsgBox "down"& SpTotalStepCnt
        objHtmlFile.WriteLine ("<td><font color=""#7A297A"">"&SpTotalStepCnt&"</font></td></tr>")

    	
    	objHtmlFile.WriteLine ("<body  bgcolor=	white>")
        objHtmlFile.WriteLine ("<table align=""center"" width=""900"" style=""font-family: Georgia, Arial;"" border=""0"">")
        objHtmlFile.WriteLine ("<tr><td align=""center"" bgcolor=""#4A9586""><font color=white><b>Automation Script Execution Status</b></font></td></tr></table>")		
      
		objHtmlFile.WriteLine (" <table align=""center"" width=""900"" border=""0"">")
		objHtmlFile.WriteLine ("<tr><td bgcolor=""#4A9586"" ><font color=white><b>S.No</b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586><font color=white><b>Date&Time</b></font></td>") 
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586><font color=white><b>ScriptName</b></font></td>") 
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586 ><font color=white><b>Status</b></font></td>")
		objHtmlFile.WriteLine ("<td bgcolor=#4A9586><font color=white><b>Results</b></font></td></tr>")	

		i = 0
			
		outFile2=PROJECT_DIR&"HtmlResults\HghLelRptSts.txt"
		Set objFSO2=CreateObject("Scripting.FileSystemObject")
		Set objFile2 = objFSO2.OpenTextFile(outFile2,1,True)
		
		do until objFile2.AtEndOfStream
		    strLine= objFile2.ReadLine()
			If not strLine = ""  Then	
		    strline1 = split(strLine,",")
			'MsgBox "Reading hello world:"&strline1(0)&strline1(1)&strline1(2)&strline1(3)
			ModName=strline1(0)
			tcEndTime = strline1(1)
			tcName=strline1(2)
			stus=strline1(3)
			refLinkRep = strline1(4)
			i=i+1
			dtIndex=i
			
			Set objFSO3=CreateObject("Scripting.FileSystemObject")
			Set objReadFile = objFSO3.OpenTextFile(refLinkRep, 1, False)

			'Read file contents
			contents = objReadFile.ReadAll
			
			
			'Close file
			objReadFile.Close
			Set objReadFile = Nothing
			Set objFSO3 = Nothing 
		
			objHtmlFile.WriteLine ("<TR bgColor=#99CCFF>")
			objHtmlFile.WriteLine ("<td align='center' ><font color=#7A297A>"&dtIndex&"</font></td>")
			objHtmlFile.WriteLine ("<td ><font color=#7A297A>"&tcEndTime&"</font></td>")
			objHtmlFile.WriteLine ("<td ><font color=#7A297A>"&tcName&"</font></td>")
			If ucase(stus)="PASS" Then
				objHtmlFile.WriteLine ("<td ><font color=#268147><b>"&stus&"</b></font></td>")
				ElseIf ucase(stus)="FAIL" Then
				objHtmlFile.WriteLine ("<td ><font color=#FF5733><b>"&stus&"</b></font></td>")
				Else
				objHtmlFile.WriteLine ("<td ><font color=#bd33ff><b>TRUNCATED</b></font></td>")
			End If	
			objHtmlFile.WriteLine ("<td bgcolor=#99CCFF><a href=""javascript:showhide('Div"&dtIndex&"')"">View Results</a></td></tr>")
	        objHtmlFile.WriteLine ("<tr><td colspan =""5""><Div id='Div"&dtIndex&"' style=""display:none;;border-style:solid;border-width:thin;border-color:Navy"">")
	        objHtmlFile.WriteLine (contents)
		objHtmlFile.WriteLine ("</div> </td></tr>")
		End if
		Loop
	
		
		objFile2.Close
		objHtmlFile.WriteLine ("</table></body></html>")
		objHtmlFile.Close
	End Function


