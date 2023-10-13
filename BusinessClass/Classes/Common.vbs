''Creating sap object
set oControls=new clsControlsSAP

'Set clsControlsSAP = new clsControlsSAP


Function FioriAppLaunch
	
'	SystemUtil.Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe","https://s4hana5.mydomain.com:44300/sap/bc/ui2/flp"    ---- RDP location
'	SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://s4hana5.mydomain.com:44300/sap/bc/ui2/flp"
'	SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://s4hana5.nivedasoft.com:44300/sap/bc/ui2/flp"
	SystemUtil.Run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe","https://s4hana5.nivedasoft.com:44300/sap/bc/ui2/flp"
	
	If Browser("name:=Logon").Exist(2) Then
		
		Wait 5
		ReportWriter micPass, "Fiori LaunchPad Check", " Fiori LaunchPad Opened Successfully", 0 
	 else 
		 ReportWriter micFail, "Fiori LaunchPad Check", " Fiori LaunchPad Is Not Opened", 1
	
	End If
	
End Function

Function LoginDetails(InputData)

       wait 1
	Browser("name:=Logon").Page("title:=Logon").WebEdit("name:=sap-user").Set EXECUTION_ENVIRONMENT_3
	Browser("name:=Logon").Page("title:=Logon").WebEdit("name:=sap-password").SetSecure EXECUTION_ENVIRONMENT_4
	
	ReportWriter micPass, "Login Check Check", EXECUTION_ENVIRONMENT_3 &" Login Successfully", 0 
	
	Browser("name:=Logon").Page("title:=Logon").WebButton("outertext:=Log On").Click
	
	
End Function


Function SelectMyApp(MyApp)
	
	Browser("name:=Home").Page("title:=Home").WebButton("acc_name:=Home - Show All My Apps").Click
	Wait 5
	
	Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesList-listUl").WebElement("innerhtml:="&MyApp).Highlight
'
	Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesList-listUl").WebElement("innerhtml:="&MyApp).Click
	
	ReportWriter micPass, "Select My App", MyApp& " Selected Successfully", 0 
	
End Function


Function SelectTCode(InputData,TCode)
	
	Browser("name:=Home").Page("title:=Home").WebList("html id:=oItemsContainerlist-listUl").Select TCode ' "Create Sales OrdersVA01"
	Wait 1
	ReportWriter micPass, "Select T-Code", TCode& " Selected Successfully", 0 

End Function


Function OrderDetails(InputData)
	Wait 3
	Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPEdit("logical name:=Order Type").Set InputData("OrderType") '"OR"
	Wait 1
	ReportWriter micPass, "Enter Order Type",InputData("OrderType")& " Order Type Entered Successfully", 0 
	Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Sales Organization").SAPEdit("logical name:=Sales Organization").Set InputData("SalesOrganization") '"ABCB"
	Wait 1
	ReportWriter micPass, "Enter Sales Organization",InputData("SalesOrganization")& " Sales Organization Entered Successfully", 0
	Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Distribution Channel").SAPEdit("logical name:=Distribution Channel").Set InputData("DistributionChannel") ' "A1"
	Wait 1
	ReportWriter micPass, "Enter Distribution Channel ",InputData("DistributionChannel")& " Distribution Channel Entered Successfully", 0
	Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Division").SAPEdit("logical name:=Division").Set InputData("Division") ' "B1"
	Wait 1
	ReportWriter micPass, "Enter Division ",InputData("Division")& " Division Entered Successfully", 0

	Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPButton("name:=Continue").Click
	Wait 3
'	Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Sold-To Party").SAPEdit("logical name:=Sold-To Party").Set "1000014"
''Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Ship-to party").SAPEdit("logical name:=Ship-To Party").Highlight
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Ship-to party").SAPEdit("logical name:=Ship-To Party").Set "1000014"
'	Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Customer Reference").SAPEdit("logical name:=Cust. Reference").Set "Test"
End Function



Function EnterSalesOrderDetails(InputData)
	
	Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Sold-To Party").SAPEdit("logical name:=Sold-To Party").Set InputData("SoldToParty") ' "1000014"
	Set Keys = CreateObject("WScript.Shell")
		Keys.SendKeys("{ENTER}")
		Wait 1
	Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Customer Reference").SAPEdit("logical name:=Cust. Reference").Set InputData("CustReference") '"Test"
	Set Keys = CreateObject("WScript.Shell")
	Keys.SendKeys("{ENTER}")
	Wait 1
	Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebTable("innertext:=Terms of payment key").SAPEdit("logical name:=Pyt Terms").Set InputData("Payment")   ' "0001"
	Set Keys = CreateObject("WScript.Shell")
		Keys.SendKeys("{ENTER}")
	Wait 1
	Dim oDesc,iCounter
 
	Set oDesc = Description.Create
 
	oDesc("micclass").value = "SAPTable"
 
	'Find all the SAPTables in a Page
 
	Set objChkBox = Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").ChildObjects(oDesc)
	'objChkBox(2).highlight 
	objChkBox(2).SetCellData 2,3,InputData("Material") '"102"
	objChkBox(3).SetCellData 2,5,InputData("Units") '"1"
	objChkBox(3).SetCellData 2,14,InputData("Plant") '"1710"	
	ReportWriter micPass, "Enter Division ",InputData("SoldToParty")& " Division Entered Successfully", 0
	
End Function


Function CaptureOrderNumber
	
	''Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").SAPButton("logical name:=Save").Highlight
	Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").SAPButton("logical name:=Save").Click
'
	OrderNumber =Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebElement("outertext:=Standard Order \d.*").GetROProperty("outertext") 
'
'
	value = split (OrderNumber," ")
	'MsgBox value (0)
	SalesOrderNumver = value (2)
	MsgBox SalesOrderNumver
	
End Function

'***************************************************************************************************************************************************************
Function SAPLogon (InputData)

	'@ Help                    :
	'@ Group                   : Common
	'@ Method Name             : SAPLogon
	'@ Pre - Condition/Screen  : None
	'@ Purpose                 : Login to the SAP system
	'@ Paramerters	           : None
	'@ Returns                 : None

	'Start Reporting 
	Reporter.StartStatusTrack	
	Reporter.StartFunction "SAP Login","True"
	On Error resume next
	'Closing all existing session of SAP to launch new session for login
	REM MsgBox "Login in"
	CloseExistingSessions
	REM MsgBox aLoginData(0,1)&" "& aLoginData(0,2)&" "& aLoginData(0,3)&" "& aLoginData(0,4)&" "& aLoginData(0,5)
	''SAPGuiUtil.AutoLogon aLoginData(0,1), aLoginData(0,4), aLoginData(0,2), aLoginData(0,3), aLoginData(0,5)
	SAPGuiUtil.AutoLogon aLoginData(0,1), aLoginData(0,2), aLoginData(0,3), aLoginData(0,4), aLoginData(0,5)
	REM MsgBox err.description
    If err.description = "" Then
'    		ReportWriter micPass, "Select Status Check", sStatus&" Status Selected Successfully", 0 
		ReportWriter micPass, "Logging in to SAP application","Successfully logged into '" & aLoginData(0,1) & "' system",0
	Else
		ReportWriter MicFail, "Logging in to SAP application","Unable logged into '" & aLoginData(0,1) & "' system",1
	End If
    'End Reporting 
	Reporter.EndFunction
	Reporter.EndStatusTrack
End Function
'***************************************************************************************************************************************************************
Sub CloseExistingSessions ()
	' @HELP
	' @class	:	clsControlsSAP
	' @method	:   CloseExistingSessions ()
	' @returns	:  
	' @parameter:	
	' @notes:   	Closes all the existing SAP sessions
	' @END
	SAPGuiUtil.CloseConnections
	ProExist = False 
	Set AllProcess = getobject("winmgmts:")
	For Each Process In AllProcess.InstancesOf("Win32_process") 
		If Instr ((Process.Name),ProcessName) = 1 Then 
			SystemUtil.CloseProcessByName ("pcsfe.exe") 
			ProExist = True 
			Exit for 
		End If 
	Next
	ProExist = False 
	Set AllProcess = getobject("winmgmts:")
	For Each Process In AllProcess.InstancesOf("Win32_process") 
		If Instr ((Process.Name),ProcessName) = 1 Then 
			SystemUtil.CloseProcessByName ("pcsws.exe") 
			ProExist = True 
			Exit for 
		End If 
	Next 
	ProExist = False 
	Set AllProcess = getobject("winmgmts:")
	For Each Process In AllProcess.InstancesOf("Win32_process") 
		If Instr ((Process.Name),ProcessName) = 1 Then 
			SystemUtil.CloseProcessByName ("pcscm.exe") 
			ProExist = True 
			Exit for 
		End If 
	Next 
	ProExist = False 
	Set AllProcess = getobject("winmgmts:")
	For Each Process In AllProcess.InstancesOf("Win32_process") 
		If Instr ((Process.Name),ProcessName) = 1 Then 
			SystemUtil.CloseProcessByName ("TeRun.exe") 
			ProExist = True 
			Exit for 
		End If 
	Next
End Sub
'***************************************************************************************************************************************************************
Function fnVA01SO_WithRef()

	''Starting the reporting
	Reporter.StartStatusTrack	
	Reporter.StartFunction "Creating Sales Order with referance","True"
	'' Ã‰ntering T code
	oControls.SetSAPOKCode ("/nVA01")
	''Entering Order Type
	oControls.SetSAPEdit "VBAK-AUART", "ZUS1"
	''Entering Sales Organization
	oControls.SetSAPEdit "VBAK-VKORG", "1710"
	''Entering Distribution Channel
	oControls.SetSAPEdit "VBAK-VTWEG", "10"
	''Entering Division
	oControls.SetSAPEdit "VBAK-SPART", "10"	
	''Clicking on Create with referance button
	oControls.ClickSAPButton "btn\[8\]"
	''Selecting Order tab
	oControls.SelectSAPTabStrip "MYTABSTRIP", "Order"
	''Entering order number
	oControls.SetSAPEdit "LV45C-VBELN", "6206241"
	oControls.ClickSAPButton "Copy"
	oControls.ClickSAPButton "btn\[0\]"
	oControls.ClickSAPButton "btn\[0\]"
	REM oControls.ClickSAPButton "btn\[0\]"
	REM oControls.ClickSAPButton "btn\[0\]"
	''Entering CC Type
	oControls.SetSAPEdit "CCDATA-CCINS", "VISA"	
	''Entering CC no
	oControls.SetSAPEdit "CCDATA-CCNUM", "-E803-1111-FWZW7FT203ME6W"		
	''Entering CC exp
	oControls.SetSAPEdit "CCDATE-EXDATBI", "07/20"		
	''Entering CC CVV
	oControls.SetSAPEdit "CCARD_CVV-CVVAL", "123"	
	''Clicking on Display header details button
	oControls.ClickSAPButton "BT_HEAD"
	''Selecting Additional data B tab
	oControls.SelectSAPTabStrip "TAXI_TABSTRIP_HEAD", "Additional data B"
	''Entering Account Number
	oControls.SetSAPEdit "VBAK-ZZACCOUNTID", "12345"		
	''Entering Consultant ID
	oControls.SetSAPEdit "VBAK-ZZCONSULTID", "1234"		
	''Clicking on save button
	oControls.ClickSAPSaveButton()
	''Clicking on save button
	statusVA01=oControls.GetSAPStatusBarInfo("item2")
	fnVA01SO_WithRef=statusVA01
End function

'=============================================================================================================
Function VA03_DisplaySO(sSO)
	''Starting the reporting
	Reporter.StartStatusTrack	
	Reporter.StartFunction "Display Sales Order VA03","True"
	'' Ã‰ntering T code
	oControls.SetSAPOKCode ("/nVA03")
	''Entering Sales Order

	oControls.SetSAPEdit "VBAK-VBELN", sSO
	''Enter
	oControls.ClickSAPEnter ()
	''Get the status bar information
	statusVA03=oControls.GetSAPStatusBarInfo("itemscount")
	
	If statusVA03=0 Then
      oControls.UpdateExecutionReport micPass, "GetSAPStatusBarInfo", "Sales Order was successfully displayed in VA03"
     Else
	 varTmpstsVA03n=oControls.GetSAPStatusBarInfo("text")
	 oControls.UpdateExecutionReport micFail, "GetSAPStatusBarInfo", "Unable to display Sales order in VA03-Error message '"&varTmpstsVA03n&"'" 
	End if
End function
'=============================================================================================================
Function VL01N_CreateDel(sSO)
''Starting the reporting
	Reporter.StartStatusTrack	
	Reporter.StartFunction "Create Delivery VL01N","True"
	'' Ã‰ntering T code
	oControls.SetSAPOKCode ("/nVL01N")
	
	oControls.SetSAPEdit "LIKP-VSTEL", "1711"
	
	oControls.SetSAPEdit "LV50C-DATBI", "02/23/2017"

	oControls.SetSAPEdit "LV50C-VBELN", sSO
	
	oControls.ClickSAPEnter ()
	
	oControls.ClickSAPSaveButton()
	
	VL01N_CreateDel=oControls.GetSAPStatusBarInfo("item2")


End Function

'=============================================================================================================
Function VL03N_DisplayOBDel(sOBDel)
''Starting the reporting
	Reporter.StartStatusTrack	
	Reporter.StartFunction "Display Outbound Delivery VL03N","True"
	'' Ã‰ntering T code
	oControls.SetSAPOKCode ("/nVL03N")
	
	oControls.SetSAPEdit "LIKP-VBELN", sOBDel
	oControls.ClickSAPEnter ()
	
	oControls.SelectSAPMenuItem "Extras;Delivery Output;Header"
	oControls.SelectSAPTableRow "SAPDV70ATC_NAST3", 1
	''Clicking on Processing log
	oControls.ClickSAPButton "btn\[26\]"
	varIdocNo=oControls.GetSAPLabelContent("wnd[1]/usr/lbl[6,6]")
	oControls.ClickSAPButton "btn\[0\]"
	VL03N_DisplayOBDel=varIdocNo
	
End Function

Function WE19_IdocProcess(sDelNo)
''Starting the reporting
	Reporter.StartStatusTrack	
	Reporter.StartFunction "Idoc Processing WE19","True"
	'' Ã‰ntering T code
	oControls.SetSAPOKCode ("/nWE19")
	oControls.SetSAPEdit "MSED7START-EXIDOCNUM", "10036635"
	''Clicking on Execute
	oControls.ClickSAPButton "btn\[8\]"
	
	oControls.SetSAPLabelFocus "E1EDL20"
	oControls.SendSAPkey F2
	oControls.SetSAPEdit "VSTEL", sDelNo
	oControls.SetSAPEdit "BOLNR", "600712342224"
	oControls.ClickSAPButton "btn\[0\]"
	

End Function

Function sampletestcapturevalue
	
	value1 = 28
'	MsgBox value1
'	Browser("Google").Page("Practice Page").WebElement("28").Click

End Function





