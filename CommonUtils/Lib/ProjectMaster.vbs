'Initialize the Variables
Dim oWshShell
Dim oWshEnvironment
Dim arrAutomationPath
Dim sScriptPath
Dim sProjectPath
Dim TC_ID

'Initialize the frame work.
sTestCaseID=Environment.Value("TestName")
sAutomationPath=GetAutomationSuitePath("BusinessClass")
sProjectPath=GetAutomationSuitePath("Scripts")
COMMON_LIB_DIR	= sAutomationPath	& "\CommonUtils\Lib\"

''Load Common libraries 
'LoadFunctionLibrary COMMON_LIB_DIR &"DB.vbs"
'LoadFunctionLibrary COMMON_LIB_DIR &"ErrorLog.vbs"
'LoadFunctionLibrary COMMON_LIB_DIR &"FileUtility.vbs"
'LoadFunctionLibrary COMMON_LIB_DIR &"Functions.vbs"
'LoadFunctionLibrary COMMON_LIB_DIR &"Utils.vbs"
'LoadFunctionLibrary COMMON_LIB_DIR &"SwfControls.vbs"
'LoadFunctionLibrary COMMON_LIB_DIR &"Constants.vbs"
'LoadFunctionLibrary COMMON_LIB_DIR &"Common_Junifer.qfl"
'
''Load Repo
'RepositoriesCollection.Add COMMON_REPO_DIR &"CommonRepo_Junifer.tsr"

'AddSearchFolderToQTP
' REM MsgBox "include"
' ExecuteFile sAutomationPath &"\CommonUtils\Lib\Include.vbs"
' REM MsgBox "Constants"
' ExecuteFile sAutomationPath &"\BusinessClass\Lib\Constants.vbs"
' REM MsgBox "Common"
' ExecuteFile sAutomationPath &"\BusinessClass\Classes\Common_QL.vbs"
' REM MsgBox "Scenarios"
' ExecuteFile sAutomationPath &"\BusinessClass\Classes\Scenarios_Complaints.vbs"
' REM MsgBox "end"

'*************
'**************************************
Function GetAutomationSuitePath(sFolderName)
Dim fso
Dim pFolderPath
Dim CommonExists
Dim pFolderName

Set fso=CreateObject("scripting.filesystemobject")
'pFolderPath=Environment("TestDir")
pFolderPath=Environment.Value("TestDir")
pFolderName=fso.GetParentFolderName(pFolderPath)
CommonExists=False

While not CommonExists
	If fso.FolderExists (pFolderName&"\"&sFolderName) Then
		CommonExists=true
	Else
		pFolderName=fso.GetParentFolderName(pFolderPath)
		pFolderPath=pFolderName
	End If
	
Wend
GetAutomationSuitePath= pFolderName

End Function
'***************************************************
Function AddSearchFolderToQTP()

Dim qtApp
Set qtApp = CreateObject("QuickTest.Application")
qtApp.Folders.RemoveAll
qtApp.Folders.Add sProjectPath&"\Lib", 1
Set qtApp=Nothing

End Function
'***************************************************
Function GetModuleName()
	mPart=Split(Environment("TestDir"),"Scripts\")
	mSubPart=Split(mPart(1),"\")
	GetModuleName=mSubPart(0)
End Function
