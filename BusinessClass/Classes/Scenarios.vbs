Function fnSampleFunction ()

	'@ Help                    :
	'@ Group                   : Common
	'@ Method Name             : SAPLogon
	'@ Pre - Condition/Screen  : None
	'@ Purpose                 : Login to the SAP system
	'@ Paramerters	           : None
	'@ Returns                 : None

	'Start Reporting 
	Reporter.StartStatusTrack	
	Reporter.StartFunction "fnSampleFunction","True"
	SAPLogon (aLoginData)
	REM Closing all existing session of SAP to launch new session for login
	REM CloseExistingSessions ()
	REM SAPGuiUtil.AutoLogon aLoginData(), aLoginData(), aLoginData(), aLoginData(), aLoginData()
    MsgBox "sample function"
    'End Reporting 
	Reporter.EndFunction
	Reporter.EndStatusTrack
End Function

Function fnOTCStandard_Flow()

	'@ Help                    :
	'@ Group                   : Common
	'@ Method Name             : SAPLogon
	'@ Pre - Condition/Screen  : None
	'@ Purpose                 : Login to the SAP system
	'@ Paramerters	           : None
	'@ Returns                 : None

	'Start Reporting 
	Reporter.StartStatusTrack	
	Reporter.StartFunction "fnOTCFlow with referance","True"
	''Loggin in to SAP system
	SAPLogon (aLoginData)
	''creating sales order with referance
	varStndSO=fnVA01SO_WithRef()	
	VA03_DisplaySO(varStndSO)
	varDelNo=VL01N_CreateDel(varStndSO)
	sIdocNo=VL03N_DisplayOBDel(varDelNo)
	WE19_IdocProcess(varDelNo)
    'End Reporting 
	Reporter.EndFunction
	Reporter.EndStatusTrack
End Function

Function sTestMethod(InputData)
	'Start Reporting 
	Reporter.StartStatusTrack	
	Reporter.StartFunction "Test Sample","True"
	
	wait 1
	msgbox "in"
	sampletestcapturevalue
	
	''UpdateExecutionReport micPass, "Logging in to SAP application","Successfully logged in sys"
	ReportWriter micPass,"Validate Finished Wizard opens after finish", "Finished Wizard Appears as expected", 0
	
	Reporter.EndFunction
	Reporter.EndStatusTrack
	
End Function


Function SampleM(InputData)
	'Start Reporting 
	Reporter.StartStatusTrack	
	Reporter.StartFunction "Test Sample1","True"
	
	wait 1
	msgbox "in"
	sampletestcapturevalue

	'UpdateExecutionReport micPass, "Logging in to SAP application","Successfully logged in sys"
	ReportWriter micPass,"Validate Finished Wizard opens after finish", "Finished Wizard Appears as expected", 0
	
	Reporter.EndFunction
	Reporter.EndStatusTrack
	
End Function
