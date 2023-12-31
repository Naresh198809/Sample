'****************************************************************************************************************************
' $Filename:		WebControls.vbs
' $Description: 	WebControls 
' $Copyright: 		Arsin Corporation
'****************************************************************************************************************************
Dim oObjectDescriptions
Dim oSesObjDescriptions
Dim BaseWindow

Public Const Short_Interval = 2
Public Const Long_Interval = 2

Function SetSwfWindowObjectDescription(sVal)
		Set oWindowObjDescriptions = Description.Create ()
		oWindowObjDescriptions("text").Value = sVal
End Function

Function GetObjectDescriptions ()
		Set oSesObjDescriptions = Description.Create ()
		oSesObjDescriptions("micClass").Value = "Browser"
		
		Set oObjectDescriptions = Description.Create ()
		oObjectDescriptions("micClass").Value = "Page"
End Function  
'****************************************************************************************************************************************************************************************************************************************************************************************
Function SwfTableSelectRow(mWindow, sPropAndVal, sRow)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfTableSelectRow(mWindow, sPropAndVal, sRow)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sRow
		' @notes	: Selects a row from SwfTable
		' @END
		
		If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
		Else
			If mWindow.SwfTable(sPropAndVal).Exist(5) Then
				mWindow.SwfTable(sPropAndVal).SelectRow sRow
				ReportWriter micDone, "Row Selection from Table", "Row selected : "&sRow,0
			Else
				ReportWriter MicFail, "Row Selection from Table", "Unable to select Row : "&sRow,1
			End If
		End If
		
End Function

Function SwfTableSelectRowDblClick(mWindow, sPropAndVal, sRow)
		' @HELP new
		' @group	: SwfControls	
		' @funcion	: SwfTableSelectRow(mWindow, sPropAndVal, sRow)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sRow
		' @notes	: Selects a row from SwfTable
		' @END
		
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfTable(sPropAndVal).Exist(5) Then
			mWindow.SwfTable(sPropAndVal).SelectRow sRow
			wait 1
			mWindow.SwfTable(sPropAndVal).DblClick
			ReportWriter micDone, "Row Selection from Table", "Row selected : "&sRow,0
		Else
			ReportWriter MicFail, "Row Selection from Table", "Unable to select Row : "&sRow,1
		End If
	End If
	
End Function

Function JUN_SwfTableSelectRow(mWindow, sWinId, sPropAndVal, sRow)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfTableSelectRow(mWindow, sPropAndVal, sRow)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sRow
		' @notes	: Selects a row from SwfTable
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
		Exit Function
	Else
		'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
		If MainObj.SwfTable(sPropAndVal).Exist(5) Then
			MainObj.SwfTable(sPropAndVal).SelectRow sRow
			ReportWriter micDone, "Row Selection from Table", "Row selected : "&sRow,0
		Else
			ReportWriter MicFail, "Row Selection from Table", "Unable to select Row : "&sRow,1
		End If
	End If
	
End Function
'****************************************************************************************************************************************************************************************************************************************************************************************
Function SwfClickButton(mWindow, sPropAndVal)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfClickButton(mWindow, sPropAndVal, sRow)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sRow
		' @notes	: Click SwfButton
		' @END
'	If Reporter.lScenarioStatus = "FAIL"  Then
'			Exit Function
'	Else
		If mWindow.SwfButton(sPropAndVal).Exist(10) Then
			mWindow.SwfButton(sPropAndVal).Click
			ReportWriter micDone, "Button Click", "Button Clicked sucessfully",0 '"Uncomment this"
		Else
			ReportWriter MicFail, "Button Click", "Unable to click button",1 '"Uncomment this" Set value as '1' for failure after code update
		End If
	'End If
End Function

'****************************************************************************************************************************************************************************************************************************************************************************************
Function SwfSetvalueEdit(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfSetvalueEdit(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
'	If Reporter.lScenarioStatus = "FAIL"  Then
'			Exit Function
'	Else
		If mWindow.SwfEdit(sPropAndVal).Exist(5) Then
			mWindow.SwfEdit(sPropAndVal).set sText
			ReportWriter micPass, "Set value in edit box", "Set value in edit box sucessfull - "&sText,0        '"Uncomment This"
		Else
			ReportWriter MicFail, "Set value in edit box", "Unable to set value in edit box : "&sText,1         '"Uncomment This"
		End If
		Reporter.EndFunction			 
'	End If
End Function

Function SetPasswordInEdit(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfSetvalueEdit(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
If Reporter.lScenarioStatus = "FAIL"  Then

		Exit Function

	Else
		If mWindow.SwfEdit(sPropAndVal).Exist(5) Then
		wait 3
			mWindow.SwfEdit(sPropAndVal).set sText
			ReportWriter micDone, "Set Password value in edit box", "Set value in edit box sucessfull ",0        '"Uncomment This"
		Else
			ReportWriter micFail, "Set Password value in edit box", "Set value in edit box Failed ",1      '"Uncomment This"
		End If	
End If
End Function

Function SetPasswordSecure(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfSetvalueEdit(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
If Reporter.lScenarioStatus = "FAIL"  Then
		Exit Function
	Else
		If mWindow.SwfEdit(sPropAndVal).Exist(5) Then
		wait 3
			'mWindow.SwfEdit(sPropAndVal).set sText
			mWindow.SwfEdit(sPropAndVal).SetSecure sText
			ReportWriter micDone, "Enter Password  Secure", "Entered Password in secure mode sucessfull  ",0        '"Uncomment This"
		Else
			ReportWriter micFail, "Enter Password  Secure", " Password not entered in secure mode ",1        '"Uncomment This"
		End If
End If
End Function

'****************************************************************************************************************************************************************************************************************************************************************************************
Function SwfEditTypeTab(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfEditTypeTab(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
		Exit Function
	Else
		If mWindow.SwfEdit(sPropAndVal).Exist(5) Then
		mWindow.SwfEdit(sPropAndVal).Click
		mWindow.SwfEdit(sPropAndVal).Type sText
		wait 1
		mWindow.SwfEdit(sPropAndVal).Type micTab
			ReportWriter micDone, "Type value in edit box", "Set value in edit box sucessfull - "&sText,0
		Else
			ReportWriter MicFail, "Type value in edit box", "Unable to set value in edit box : "&sText,1
		End If
	End If
End Function

'****************************************************************************************************************************************************************************************************************************************************************************************
Function SwfEditClearTypeTab(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfEditTypeTab(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
		Exit Function
	Else
		If mWindow.SwfEdit(sPropAndVal).Exist(5) Then
			mWindow.SwfEdit(sPropAndVal).Object.SelectAll
			mWindow.SwfEdit(sPropAndVal).Type sText
			mWindow.SwfEdit(sPropAndVal).Type micTab
			ReportWriter micDone, "Type value in edit box", "Set value in edit box sucessfull - "&sText,0
		Else
			ReportWriter MicFail, "Type value in edit box", "Unable to set value in edit box : "&sText,1
		End If
	End If
End Function

'****************************************************************************************************************************************************************************************************************************************************************************************
Function SelValSwfComboBox(mWindow, sPropAndVal, sIndex)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SelValSwfComboBox(mWindow, sPropAndVal, sIndex)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Select a value from ComboBox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfComboBox(sPropAndVal).Exist(5) Then
			mWindow.SwfComboBox(sPropAndVal).click
			mWindow.SwfComboBox(sPropAndVal).Select sIndex
			sSelection = mWindow.SwfComboBox(sPropAndVal).GetSelection
	
			ReportWriter micDone, "Set value in SwfComboBox", "Select value in ComboBox box sucessfull - "&sSelection,0
		Else
			ReportWriter MicFail, "Set value in SwfComboBoxx", "Unable to select value in ComboBox box : "&sIndex,1
		End If
	End If
End Function

'****************************************************************************************************************************************************************************************************************************************************************************************
Function SelValSwfComboBoxByValue(mWindow, sPropAndVal, sItem)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SelValSwfComboBoxByValue(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Select a value from ComboBox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfComboBox(sPropAndVal).Exist(5) Then
			a = mWindow.SwfComboBox(sPropAndVal).GetContent()
			'MsgBox a
			iIndex = ""
				arrLines = Split(a, vbCrLf)
				For i = 0 To ubound(arrLines)
					If arrLines(i) = sItem  Then
						iIndex = i
						Exit for
					End If
				Next
			If iIndex <>"" Then
				mWindow.SwfComboBox(sPropAndVal).click
				mWindow.SwfComboBox(sPropAndVal).Select iIndex
				sSelection = mWindow.SwfComboBox(sPropAndVal).GetSelection	
				ReportWriter micDone, "Set value in SwfComboBox", "Select value in ComboBox box sucessfull - "&sSelection,0
			Else
				ReportWriter MicFail, "Set value in SwfComboBox", "Unable to select value in ComboBox box - "&sSelection,1
			End If
	
		Else
			ReportWriter MicFail, "Set value in SwfComboBoxx", "Unable to select value in ComboBox box : "&sIndex,1
		End If
	End If
	
End Function

Function SelValSwfComboBoxByText(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls - Combo box
		' @funcion	: SelValSwfComboBoxByText(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Select a value from ComboBox by Text present in it
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfComboBox(sPropAndVal).Exist(5) Then
			mWindow.SwfComboBox(sPropAndVal).Click
			wait 1
			mWindow.SwfComboBox(sPropAndVal).Select sText
			'sSelection = mWindow.SwfComboBox(sPropAndVal).GetSelection
	
			ReportWriter micDone, "Set value in SwfComboBox", "Select value in ComboBox box sucessfull - "&sSelection,0
		Else
			ReportWriter MicFail, "Set value in SwfComboBoxx", "Unable to select value in ComboBox box : "&sSelection,1
		End If
	End If
	
End Function

Function SelValSwfComboBoxByTextDblClick(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls - Combo box
		' @funcion	: SelValSwfComboBoxByTextDblClick(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Performs double click on dropdown to activate the object and Select a value from ComboBox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		'To Create object based on Window ID 'Additional$$$
		wait 8
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
	
		If mWindow.SwfComboBox(sPropAndVal).Exist(5) Then
			mWindow.SwfComboBox(sPropAndVal).DblClick
			wait 1
			mWindow.SwfComboBox(sPropAndVal).Select sText
'			sSelection = mWindow.SwfComboBox(sPropAndVal).GetSelection	
			ReportWriter micDone, "Set value in SwfComboBox", "Select value in ComboBox box sucessfull - "&sSelection,0
		Else
			ReportWriter MicFail, "Set value in SwfComboBoxx", "Unable to select value in ComboBox box : "&sSelection,1
		End If
	End If
	
End Function

'****************************************************************************************************************************************************************************************************************************************************************************************
Function GetValSwfComboBox(mWindow, sPropAndVal)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SelValSwfComboBox(mWindow, sPropAndVal, sIndex)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Select a value from ComboBox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfComboBox(sPropAndVal).Exist(5) Then
		sValue = mWindow.SwfComboBox(sPropAndVal).GetSelection
		GetValSwfComboBox = sValue
			ReportWriter micDone, "Get value in SwfComboBox", "Getting value in ComboBox sucessfull - "&sValue,0
		Else
			ReportWriter MicFail, "Get value in SwfComboBox", "Unable to get value in ComboBox box : ",1
		End If
	End If
End Function

Function ClickSwfEditor(mWindow, sPropAndVal)
		' @HELP
		' @group	: ClickSwfEditor	
		' @funcion	: ClickSwfEditor(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal
		' @notes	: Click on SwfEditor
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfEditor(sPropAndVal).Exist(5) Then
		mWindow.SwfEditor(sPropAndVal).Click
		'mWindow.SwfEditor(sPropAndVal).DblClick
			ReportWriter micDone, "SwfEditor Click", "SwfEditor Clicked sucessfully",0
		Else
			ReportWriter MicFail, "SwfEditor Click", "Unable to click SwfEditor",1
		End If
	End If
End Function

Function SetValSwfEditor(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: ClickSwfEditor	
		' @funcion	: ClickSwfEditor(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal
		' @notes	: Click on SwfEditor
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfEditor(sPropAndVal).Exist(5) Then
			mWindow.SwfEditor(sPropAndVal).SetText sText
			ReportWriter micDone, "SwfEditor Set Value", "SwfEditor Set sucessfully",0
		Else
			ReportWriter MicFail, "SwfEditor Set Value", "Unable to Set value in SwfEditor",1
		End If
	End If
	
End Function

Function ClickSwfTab(mWindow, sPropAndVal, sTab)
		' @HELP new
		' @group	: ClickSwfEditor	
		' @funcion	: ClickSwfEditor(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal
		' @notes	: Click on SwfEditor
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfTab(sPropAndVal).Exist(10) Then
		mWindow.SwfTab(sPropAndVal).Select sTab
'		mWindow.SwfTab(sPropAndVal).DblClick
			ReportWriter micDone, "SwfTab Click", "SwfTab Selected - "&sTab,0
		Else
			ReportWriter MicFail, "SwfTab Click", "Unable to select SwfTab - "&sTab,1
		End If
	End If
	
End Function

Function SelectItemSwfToolbar(mWindow,sPropAndVal,TabName,RibbonItemName)
	' @HELP
	' @group	: SwfToolbar	
	' @funcion	: SelectItemSwfToolbar(mWindow,sPropAndVal,TabName,RibbonItemName)
	' @returns	: None
	' @parameter: mWindow,sPropAndVal,TabName,RibbonItemName
	' @notes	: Select page from ribbon and navigate to provided Tab under page
	' @END
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	If mWindow.SwfToolbar(sPropAndVal).Exist(10) Then
		If TabName<>"" Then
			mWindow.SwfToolbar(sPropAndVal).SelectPage TabName
		End If
		Set toolbarObj = mWindow.SwfToolbar(sPropAndVal)
		toolbarObj.WaitProperty "enabled","true",20000
		toolbarObj.Press RibbonItemName '"Home;Customers;Account"
		ReportWriter micDone, "SwfTab Click", "SwfTab Selected - "&sTab,0
	Else 
		ReportWriter micDone, "SwfTab Click", "Unable to select SwfTab Selected - "&sTab,0
	End If
End If

End Function

Function SwfCalendarSetDate(mWindow,sPropAndVal,sText)
	' @HELP
	' @group	: SwfCalendarSetDate	
	' @funcion	: SwfCalendarSetDate(mWindow, sPropAndVal, sText)
	' @returns	: None
	' @parameter: mWindow, sPropAndVal, sText
	' @notes	: Set Date in Calendar box
	' @END
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	If mWindow.SwfCalendar(sPropAndVal).Exist(10) Then
		mWindow.SwfCalendar(sPropAndVal).Click
		mWindow.SwfCalendar(sPropAndVal).SetDate sText
'		mWindow.SwfTab(sPropAndVal).DblClick
		ReportWriter micDone, "SwfTab Click", "SwfTab Selected - "&sTab,0
	Else
		ReportWriter MicFail, "SwfTab Click", "Unable to select SwfTab - "&sTab,1
	End If
End If
End Function

Function SwfCheckBox(mWindow,sPropAndVal)
	' @HELP
	' @group	: SwfCalendarSetDate	
	' @funcion	: SwfCalendarSetDate(mWindow, sPropAndVal, sText)
	' @returns	: None
	' @parameter: mWindow, sPropAndVal, sText
	' @notes	: Set Date in Calendar box
	' @END
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	If mWindow.SwfCheckBox(sPropAndVal).Exist(10) Then
		mWindow.SwfCheckBox(sPropAndVal).Click
		ReportWriter micDone, "SwfTab Click", "SwfTab Selected - "&sTab,0
	Else
		ReportWriter MicFail, "SwfTab Click", "Unable to select SwfTab - "&sTab,1
	End If
End If
End Function

Function SwfObjClick(mWindow, sPropAndVal)
		' @HELP new
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.WinObject(sPropAndVal).Exist(5) Then
			mWindow.WinObject(sPropAndVal).Click
			wait 1
			ReportWriter micDone, "Click on a Winobject", "Click on a Winobject successfull - ",0
		Else
			ReportWriter MicFail, "Click on a Winobject", "Unable to Click on a Winobject",1
		End If
	End If
End Function

'Function SwfObjClick(mWindow, sPropAndVal)
'		' @HELP - new
'		' @group	: SwfControls	
'		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
'		' @returns	: None
'		' @parameter: mWindow, sPropAndVal, sText
'		' @notes	: Set value in SwfEditbox
'		' @END
'
'		If mWindow.WinObject(sPropAndVal).Exist(5) Then
'			mWindow.WinObject(sPropAndVal).Click
'			wait 1
'			ReportWriter micDone, "Click on a Winobject", "Click on a Winobject successfull - ",0
'		Else
'			ReportWriter MicFail, "Click on a Winobject", "Unable to Click on a Winobject",0
'		End If
'End Function

Function SelectCellSwfTreeView(mWindow, sPropAndVal, sText, sText2)
		' @HELP - new
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.SwfTreeView(sPropAndVal).Exist(5) Then
			mWindow.SwfTreeView(sPropAndVal).SelectCell sText,sText2
			wait 1
			ReportWriter micPass, "Click on a Tree Node", "Node "&sText&" has been selected",0
		Else
			ReportWriter MicFail, "Click on a Tree Node", "Unable to Click on a Tree Node",1
		End If
	End If
End Function

Function SelectCellSwfTreeViewDblClick(mWindow, sPropAndVal, sText, sText2)
		' @HELP - new
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		'To Create object based on Window ID
'If sWinId<>"" and Isnumeric(sWinId) Then
'	Set MainObj = JUN_WINDOW.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
'ElseIf sWinId<>"" Then
'	Set MainObj = JUN_WINDOW.SwfWindow(sWinId)
'Else
'	Set MainObj = JUN_WINDOW
'End If							  
		If mWindow.SwfTreeView(sPropAndVal).Exist(5) Then
			mWindow.SwfTreeView(sPropAndVal).SelectCell sText,sText2
			mWindow.SwfTreeView(sPropAndVal).DblClick
			wait 1
			ReportWriter micDone, "Click on a Winobject", "Click on a Winobject successfull - ",0
		Else
			ReportWriter MicFail, "Click on a Winobject", "Unable to Click on a Winobject",1
		End If
	End If
End Function

Function ClickOnTreeNode(InputData, sWinId, sPropAndVal, sVal)
If Reporter.lScenarioStatus = "FAIL"  Then
	Exit Function
Else
'To Create object based on Window ID
If sWinId<>"" and Isnumeric(sWinId) Then
	Set MainObj = JUN_WINDOW.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
ElseIf sWinId<>"" Then
	Set MainObj = JUN_WINDOW.SwfWindow(sWinId)
Else
	Set MainObj = JUN_WINDOW
End If
		
'TREE_PRODLIST1= sPropAndVal
MainObj.SwfTreeView(sPropAndVal).WaitProperty "enabled","True",500
MainObj.SwfTreeView(sPropAndVal).Highlight
Set tree = MainObj.SwfTreeView(sPropAndVal)
ccount=tree.GetItemsCount

'CellValue=InputData("AccNo")
For i=0 to ccount-1
    celval=tree.GetItem(i)
    
    If i=0 Then
    	CellValue=celval
    Else
    	CellValue  = CellValue&";"&celval	
    End If
    
    If Instr(1,celval,sVal)>0 Then
  	JUN_Window_SelectCellSwfTreeViewDblClick JUN_WINDOW, sWinId, sPropAndVal,CellValue,0  ' MPAN
  	wait 1
  	ReportWriter micPass, "Validate "&celval&" available under Expected Node", celval&" : is available under node - "&tree.GetItem(0),0
        Exit for
    End If
     
Next
End If 
End  Function 
	
Function VbSetvalueEdit(mWindow, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: VbSetvalueEdit(mWindow, sPropAndVal, sText)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.WinEdit(sPropAndVal).Exist(5) Then
'		mWindow.WinEdit(sPropAndVal).Click
		mWindow.WinEdit(sPropAndVal).highlight
		mWindow.WinEdit(sPropAndVal).Set sText
		wait 1
		mWindow.WinEdit(sPropAndVal).Type micTab
			ReportWriter micDone, "Set value in VB edit box", "Set value in VB edit box sucessfull - "&sText,0
		Else
			ReportWriter MicFail, "Set value in VB edit box", "Unable to set value in VB edit box : "&sText,1
		End If
	End If
End Function

Function VbWInObjectClick(mWindow, sPropAndVal)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else	
		If mWindow.WinObject(sPropAndVal).Exist(5) Then
		mWindow.WinObject(sPropAndVal).Click
		wait 1
			ReportWriter micDone, "Click on a Winobject", "Click on a Winobject successfull - ",0
		Else
			ReportWriter MicFail, "Click on a Winobject", "Unable to Click on a Winobject",1
		End If
	End if
End Function

Function VbWInDialogObjectClick(mWindow, sDialog, sPropAndVal)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		If mWindow.Dialog(sDialog).WinObject(sPropAndVal).Exist(5) Then
			mWindow.Dialog(sDialog).Click
			mWindow.Dialog(sDialog).WinObject(sPropAndVal).Click
		wait 1
			ReportWriter micDone, "Click on a Winobject", "Click on a Winobject successfull - ",0
		Else
			ReportWriter MicFail, "Click on a Winobject", "Unable to Click on a Winobject",1
		End If
	End If
End Function

Function JUN_SwfEditClearTypeTab(sWinId,sPropAndVal, sText)

If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	Set MainObj = SwfWindow("micClass:=SwfWindow").SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	
	If MainObj.SwfEdit(sPropAndVal).Exist(5) Then
	       MainObj.SwfEdit(sPropAndVal).Object.SelectAll
	       MainObj.SwfEdit(sPropAndVal).Type sText
	       MainObj.SwfEdit(sPropAndVal).Type micTab
		ReportWriter micDone, "Type value in edit box", "Set value in edit box sucessfull - "&sText,0
	Else
		ReportWriter MicFail, "Type value in edit box", "Unable to set value in edit box : "&sText,1
	End If
End If

End Function

Function JUN_SwfEditorClearTypeTab(sWinId,sPropAndVal, sText)

If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	Set MainObj = SwfWindow("micClass:=SwfWindow").SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	
	If MainObj.SwfEditor(sPropAndVal).Exist(5) Then
	       MainObj.SwfEditor(sPropAndVal).Object.SelectAll
	       MainObj.SwfEditor(sPropAndVal).Type sText
	       MainObj.SwfEditor(sPropAndVal).Type micTab
		ReportWriter micDone, "Type value in edit box", "Set value in edit box sucessfull - "&sText,0
	Else
		ReportWriter MicFail, "Type value in edit box", "Unable to set value in edit box : "&sText,1
	End If
End If

End Function

Function JUN_TXTWINDOW_SwfEditClearTypeTab(mWindow,sWinId,sPropAndVal, sText)

If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else

	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
	
	If MainObj.SwfEdit(sPropAndVal).Exist(5) Then
		MainObj.SwfEdit(sPropAndVal).Click
		MainObj.SwfEdit(sPropAndVal).Set sText
		ReportWriter micDone, "Type value in edit box", "Set value in edit box sucessfull - "&sText,0
	Else
		ReportWriter MicFail, "Type value in edit box", "Unable to set value in edit box : "&sText,1
	End If
End If

End Function

Function JUN_SwfEditorClearTypeTab_MultiWindow(mWindow,sWinId,sPropAndVal, sText)

If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	
	'To Create object based on Window ID
	If sWinId<>"" and Isnumeric(sWinId) Then
		Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	ElseIf sWinId<>"" Then
		Set MainObj = mWindow.SwfWindow(sWinId)
	Else
		Set MainObj = mWindow
	End If
		
	If MainObj.SwfEditor(sPropAndVal).Exist(5) Then
	       MainObj.SwfEditor(sPropAndVal).Object.SelectAll
	       MainObj.SwfEditor(sPropAndVal).Type sText
	       MainObj.SwfEditor(sPropAndVal).Type micTab
		ReportWriter micDone, "Type value in edit box", "Set value in edit box sucessfull - "&sText,0
	Else
		ReportWriter MicFail, "Type value in edit box", "Unable to set value in edit box : "&sText,1
	End If
End If

End Function

'SwfWindow("Junifer Systems Ltd [PREP]").SwfWindow("Test Enterprises One [10212987]").SwfTab("tabControl").Select "Agreements"

Function JUN_SwitchSwfTab(sWinId, sPropAndVal, sText)
'		' @HELP new
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	Set mWindow = SwfWindow("micClass:=SwfWindow")
	
	'To Create object based on Window ID
	If sWinId<>"" and Isnumeric(sWinId) Then
		Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	ElseIf sWinId<>"" Then
		Set MainObj = mWindow.SwfWindow(sWinId)
	Else
		Set MainObj = mWindow
	End If
	
	If MainObj.SwfTab(sPropAndVal).Exist(5) Then
	
	      ' MainObj.SwfTab(sPropAndVal).Object
	       MainObj.SwfTab(sPropAndVal).Select sText
	       
	       'Wait property
		MainObj.SwfTab(sPropAndVal).WaitProperty "enabled","true",200
		wait 1
		ReportWriter micDone, "Switch tab", "Switched to "&sText&" tab successful",0
	Else
		ReportWriter MicFail, "Switch tab", "Unable to switch tab",1
	End If
End If
	
End Function

Function SwitchSwfTab(mWindow,sWinId, sPropAndVal, sText)

If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	'To Create object based on Window ID
	If sWinId<>"" and Isnumeric(sWinId) Then
		Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	ElseIf sWinId<>"" Then
		Set MainObj = mWindow.SwfWindow(sWinId)
	Else
		Set MainObj = mWindow
	End If
	
	If MainObj.SwfTab(sPropAndVal).Exist(5) Then
	       MainObj.SwfTab(sPropAndVal).Select sText
	       wait 1
		ReportWriter micDone, "Switch tab", "Switch tab successful",0
	Else
		ReportWriter MicFail, "Switch tab", "Unable to switch tab",1
	End If
End If

End Function
	
Function JUN_SelectComboBox(sWinId, sPropAndVal, sText)
'		' @HELP new
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	Set MainObj = SwfWindow("micClass:=SwfWindow").SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	
	If MainObj.SwfComboBox(sPropAndVal).Exist(5) Then
	      ' MainObj.SwfTab(sPropAndVal).Object
	      MainObj.SwfComboBox(sPropAndVal).Click
	       MainObj.SwfComboBox(sPropAndVal).Select sText
	       wait 1
		ReportWriter micDone, "Switch tab", "Switch tab successful",0
	Else
		ReportWriter MicFail, "Switch tab", "Unable to switch tab",1
	End If
End If

End Function

Function JUN_Window_SelectComboBox(mWindow, sWinId, sPropAndVal, sText)
'		' @HELP new
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
If Reporter.lScenarioStatus = "FAIL"  Then
	Exit Function
Else
	'To Create object based on Window ID
	If sWinId<>"" and Isnumeric(sWinId) Then
		Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	ElseIf sWinId<>"" Then
		Set MainObj = mWindow.SwfWindow(sWinId)
	Else
		Set MainObj = mWindow
	End If
	
	If MainObj.SwfComboBox(sPropAndVal).Exist(5) Then
		MainObj.SwfComboBox(sPropAndVal).DblClick
		wait 1
	       MainObj.SwfComboBox(sPropAndVal).Select sText
		ReportWriter micPass, "Select Value from Combo box", sText&" selected from combo box",0
	Else
		ReportWriter micFail, "Select Value from Combo box", "Unable to select "&sText&" from combo box",1
	End If
End If

End Function

Function JUN_Window_SwfCalendarSetDate(mWindow,sWinId,sPropAndVal,sText)
	' @HELP
	' @group	: SwfCalendarSetDate	
	' @funcion	: SwfCalendarSetDate(mWindow, sPropAndVal, sText)
	' @returns	: None
	' @parameter: mWindow, sPropAndVal, sText
	' @notes	: Set Date in Calendar box
	' @END
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	'To Create object based on Window ID
	If sWinId<>"" and Isnumeric(sWinId) Then
		Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	ElseIf sWinId<>"" Then
		Set MainObj = mWindow.SwfWindow(sWinId)
	Else
		Set MainObj = mWindow
	End If
	
	If MainObj.SwfCalendar(sPropAndVal).Exist(10) Then
		MainObj.SwfCalendar(sPropAndVal).Click
		MainObj.SwfCalendar(sPropAndVal).SetDate sText
'		mWindow.SwfTab(sPropAndVal).DblClick
		ReportWriter micDone, "SwfTab Click", "SwfTab Selected - "&sTab,0
	Else
		ReportWriter MicFail, "SwfTab Click", "Unable to select SwfTab - "&sTab,1
	End If
End If

End Function

Function JUN_Window_SelectCellSwfTreeView(mWindow, sWinId, sPropAndVal, sText, sText2)
		' @HELP - new
		' @group	: SwfControls	
		' @funcion	: JUN_Window_SelectCellSwfTreeView(mWindow, sWinId, sPropAndVal, sText, sText2)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Select Node from Swf Tree based on "sText" from parameter
		' @END
	
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else
		'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
		Set toolbarObj = mWindow.SwfTreeView(sPropAndVal)
		toolbarObj.WaitProperty "enabled","true",20000
		If MainObj.SwfTreeView(sPropAndVal).Exist(5) Then
			cnt=MainObj.SwfTreeView(sPropAndVal).GetItemsCount
			If cnt>0 Then
				sFinalText=MainObj.SwfTreeView(sPropAndVal).GetItem(0)
				'sContent=MainObj.SwfTreeView(sPropAndVal).GetContent(0)
				For i = 1 To cnt-1
					sFoundText=MainObj.SwfTreeView(sPropAndVal).GetItem(i)
					sFinalText = sFinalText&";"&sFoundText
					If Strcomp(Trim(UCase(sFoundText)),Trim(UCase(sText)))=0 Then
						'Report Gen
						ReportWriter micDone, "Select Node from SwfTreeView", "Select Node from SwfTreeView Successful",0
						Exit for
					End If
				Next	
			End If

			MainObj.SwfTreeView(sPropAndVal).SelectCell sFinalText,sText2
			'MainObj.SwfTreeView(sPropAndVal).Click "","",micRightBtn
			wait 1
			ReportWriter micDone, "Select Node from SwfTreeView", "Select Node from SwfTreeView Successful",0
		Else
			ReportWriter MicFail, "Select Node from SwfTreeView", "Unable to Select Node from SwfTreeView Successful",1
		End If
	End If
	
End Function
		   
Function JUN_Window_SelectCellSwfTreeViewDblClick(mWindow, sWinId, sPropAndVal, sText, sText2)
		' @HELP - new - updated 08-11
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sWinId, sPropAndVal, sText, sText2 (need to pass 'sText' as semicolon separated values for multiple validation eg: parameter1, text1;text2;text3, parameter 3
		' @notes	: Used to validate multiple hierarchy text available within single tree structure in order
		' @END
	
		
									  
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
		Set toolbarObj = mWindow.SwfTreeView(sPropAndVal)
		toolbarObj.WaitProperty "enabled","true",20000
		If MainObj.SwfTreeView(sPropAndVal).Exist(5) Then
			cnt=MainObj.SwfTreeView(sPropAndVal).GetItemsCount
			If cnt>0 Then
'				sFinalText=MainObj.SwfTreeView(sPropAndVal).GetItem(0)
'				sContent=MainObj.SwfTreeView(sPropAndVal).GetContent
				sSplit=Split(sText,";")
				For z = 0 to cnt-1
					sFounFlag=0
					sFoundText=""
					sFinalText=""
					sFinalText=MainObj.SwfTreeView(sPropAndVal).GetItem(z)
					For i = 1 To UBound(sSplit)						
						sFoundText=MainObj.SwfTreeView(sPropAndVal).GetItem(z+i)
						sFinalText = sFinalText&";"&sFoundText
						If Strcomp(Trim(UCase(sFinalText)),Trim(UCase(sText)))=0 Then
							ReportWriter MicPass, "Select Node from SwfTreeView using expected text", "Expected Text is available in SwfTreeView",0
							sFounFlag=1
							Exit for
						End If
					Next
					
					If sFounFlag=1 Then
						Exit for
					End If
					
				Next					
			End If

			MainObj.SwfTreeView(sPropAndVal).SelectCell sFinalText,sText2
			''MainObj.SwfTreeView(sPropAndVal).DblClick sFinalText,sText2
			MainObj.SwfTreeView(sPropAndVal).DblClick sFinalText,sText2
'			MainObj.SwfTreeView(sPropAndVal).SelectCell sText,sText2
'			MainObj.SwfTreeView(sPropAndVal).DblClick sText,sText2
			wait 1
			ReportWriter micDone, "Select Node from SwfTreeView", "Select Node from SwfTreeView Successful",0
		Else
			ReportWriter MicFail, "Select Node from SwfTreeView", "Unable to Select Node from SwfTreeView Successful",1
		End If
	
	
End Function

Function JUN_SwfClickButton(mWindow, sWinId, sPropAndVal)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfClickButton(mWindow, sPropAndVal, sRow)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sRow
		' @notes	: Click SwfButton
		' @END
'	If Reporter.lScenarioStatus = "FAIL"  Then
'			Exit Function
'	Else		
		'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
	
		If MainObj.SwfButton(sPropAndVal).Exist(10) Then
			Wait 3
			MainObj.SwfButton(sPropAndVal).Click
			
			ReportWriter MicPass, "Button Click", "Button Clicked sucessfully",0 '"Uncomment this"
		Else
'			ReportWriter MicFail, "Button Click", "Unable to click button",1 '"Uncomment this" Set value as '1' for failure after code update
		End If
'	End If
	
End Function

Function JUN_Window_Dialog_BTNCLICK(mWindow, sDialog, sPropAndVal)
		' @HELP new
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
'If Reporter.lScenarioStatus = "FAIL"  Then
'			Exit Function
'Else
'	If Reporter.lScenarioStatus = "FAIL"  Then
'		Exit Function
'	Else
		Set MainObj = mWindow.Dialog(sDialog)
		
		If MainObj.SwfButton(sPropAndVal).Exist(5) Then
		      ' MainObj.SwfTab(sPropAndVal).Object
		       MainObj.Click
		       MainObj.SwfButton(sPropAndVal).Click
		       wait 1
			ReportWriter MicPass, "Button click", "Button click successful",0
		Else
			ReportWriter MicFail, "Button Click", "Unable to Click button",1
		End If
'	End If
'End If

End Function

Function JUN_Dialog_BTNCLICK(mWindow, sWinId, sDialog, sPropAndVal)
		' @HELP new
		' @group	: SwfControls	
		' @funcion	: VbWInObjectClick(mWindow, sPropAndVal)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sText
		' @notes	: Set value in SwfEditbox
		' @END
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'Set MainObj = mWindow.Dialog(sDialog)
	If sWinId<>"" and Isnumeric(sWinId) Then
		Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId).Dialog(sDialog)
	ElseIf sWinId<>"" Then
		Set MainObj = mWindow.SwfWindow(sWinId).Dialog(sDialog)
	Else
		Set MainObj = mWindow.Dialog(sDialog)
	End If
	
	If MainObj.SwfButton(sPropAndVal).Exist(5) Then
	      ' MainObj.SwfTab(sPropAndVal).Object
	       MainObj.Click
	       ReportWriter micPass, "Button click successful on question dialog box", "Button click successful on Question Dialog box",0
	       MainObj.SwfButton(sPropAndVal).Click
	Else
		ReportWriter MicFail, "Button Click", "Unable to Click button",1
	End If
End If

End Function

Function SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Row Number - Global Variable (sRowNo)
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
'If Reporter.lScenarioStatus = "FAIL"  Then
'			Exit Function
'Else	
	'To Create object based on Window ID
	Wait 3
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
		For i = 0 To sRows-1
			sval=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName)
			If Strcomp(Trim(UCase(sval)),Trim(UCase(sExpectedVal)))=0 Then
				MainObj.SwfTable(sPropAndVal).MakeCellVisible i,1
				MainObj.SwfTable(sPropAndVal).SelectRow i
				x=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFieldName,"x")
				y=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFieldName,"y")
				sRowNo= i
				Exit for
			End If
		Next
		
		If sRowNo<>"" Then
			ReportWriter micPass, "Select value from table", "Selected Value from Webtable "&sval&" match with in expected value - "&sExpectedVal,0
		End If
		
	Else
		ReportWriter micFail, "Select value from Table", "Table Unavailable",1
	End If
'End If

End Function

Function SelectRowByValSwfTable_StartRow(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal,sStartRow)
		' @HELP - New
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Row Number - Global Variable (sRowNo)
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
		For i = sStartRow To sRows
			sval=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName)
			If Strcomp(Trim(UCase(sval)),Trim(UCase(sExpectedVal)))=0 Then
				MainObj.SwfTable(sPropAndVal).SelectRow i
				x=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFieldName,"x")
				y=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFieldName,"y")
				sRowNo= i
				Exit for
			End If
		Next
		
		If sRowNo<>"" Then
			ReportWriter micPass, "Select value from table", "Selected Value from Webtable "&sval&" match with in expected value - "&sExpectedVal,0
		Else
			ReportWriter micFail, "Select value from table", "Unable to select row since Expected Value - "&sExpectedVal&", not available in the Table",1		
		End If
		
	Else
		ReportWriter micFail, "Select value from Table", "Table Unavailable",1
	End If
End If

End Function

Function SelectRowByValSwfTable_MultiFieldValidation(mWindow,sWinId,sPropAndVal,sFieldName1,sExpectedVal1,sFieldName2,sExpectedVal2)
		' @HELP - New
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Row Number - Global Variable (sRowNo)
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
		For i = 0 To sRows-1
			sval1=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName1)
			sval2=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName2)
			If Strcomp(Trim(UCase(sval1)),Trim(UCase(sExpectedVal1)))=0 and Strcomp(Trim(UCase(sval2)),Trim(UCase(sExpectedVal2)))=0 Then
				MainObj.SwfTable(sPropAndVal).SelectRow i
				x=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFieldName,"x")
				y=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFieldName,"y")
				sRowNo= i
				Exit for
			End If
		Next
		
		If sRowNo<>"" Then
			ReportWriter micPass, "Select value from table", "Selected Value from Webtable "&sval&" match with in expected value - "&sExpectedVal,0
		Else
			ReportWriter micFail, "Select value from table", "Unable to select row since Expected Value - "&sExpectedVal&", not available in the Table",1		
		End If
		
	Else
		ReportWriter micFail, "Select value from Table", "Table Unavailable",1
	End If
End If

End Function

Function FetchValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Field Value
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
		For i = 0 To sRows-1
			sval=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName)
			MainObj.SwfTable(sPropAndVal).SelectRow i
			Exit for
		Next
		''ReportWriter micPass, "Fetch value from ", "Value fetch from Webtable "&sval&" match with in expected value - "&sExpectedVal,0
	Else
		''ReportWriter micFail, "Fetch value from ", sExpectedVal&" not available in Swf Table",1
	End If
	
	FetchValSwfTable = sval
	
End If

End Function

Function FetchValSwfTable_ExpectedVal(mWindow,sWinId,sPropAndVal,sFindFieldName,sExpectedFieldVal,sFieldValToFetch)
		' @HELP - New
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Field Value
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
		For i = 0 To sRows-1
			sval=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFindFieldName)
			If Instr(1,UCase(sval),UCase(sExpectedFieldVal))>0 Then
				MainObj.SwfTable(sPropAndVal).SelectRow i	
				sRVal=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldValToFetch)
				x=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFindFieldName,"x")
				y=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFindFieldName,"y")
				sRowNo= i
				FetchValSwfTable_ExpectedVal = sRVal
				Exit for	
			End If
		Next
		';ReportWriter micPass, "Fetch value from "&sExpectedFieldVal, "Expected field value: "&sExpectedFieldVal&" match with available value: - "&sval&", data - "&sRVal&" retrieved from field - "&sFieldValToFetch,0
	Else
		''ReportWriter micFail, "Fetch value from "&sExpectedFieldVal&" not available in Swf Table",1
	End If
	
End If
End Function

Function SelectRowByValSwfTableDoubleClickandSelectTab(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal,sTabName)
	
If Reporter.lScenarioStatus = "FAIL"  Then
	Exit Function
Else
	wait 5
	Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
	Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
		For i = 0 To sRows-1
			sval=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName)
			If Strcomp(Trim(UCase(sval)),Trim(UCase(sExpectedVal)))=0 Then
				MainObj.SwfTable(sPropAndVal).SelectRow i
				
				x=MainObj.SwfTable(sPropAndVal).GetROProperty ("abs_x")
				y=MainObj.SwfTable(sPropAndVal).GetROProperty("abs_y")
				
'				MainObj.SwfTable(sPropAndVal).MouseClick x,y,2
					DeviceReplay.MouseMove x,y
					DeviceReplay.MouseDblClick x,y,0
				wait 5 
				SwfWindow("text:=Junifer Systems Ltd.*").SwfObject("swfname:=TicketDashboard").SwfTab("swfname:=tabControl").select sTabName
					Exit for
				
			End If
		Next
		ReportWriter micPass, "Fetch value from Table ", "Value fetch from Webtable "&sval&" match with in expected value - "&sExpectedVal,0
	Else
		ReportWriter micDone, "Fetch value from Table ", sExpectedVal&" not available in Swf Table",0
	End If
End If	



End Function

Function FetchValSwfTable_ByRowNo(mWindow,sWinId,sPropAndVal,sFieldValToFetch,sRowNo)
		' @HELP - New
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Field Value
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
			i=sRowNo
			MainObj.SwfTable(sPropAndVal).SelectRow i
			sRVal=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldValToFetch)
			x=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFindFieldName,"x")
			y=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFindFieldName,"y")
			FetchValSwfTable_ByRowNo = sRVal
				
	''	ReportWriter micPass, "Fetch value from "&sExpectedFieldVal, "Expected field value: "&sExpectedFieldVal&" match with available value: - "&sval&", data - "&sRVal&" retrieved from field - "&sFieldValToFetch,0
	Else
	''	ReportWriter micFail, "Fetch value from "&sExpectedFieldVal&" not available in Swf Table",1
	End If
	
End If

End Function

Function CloseWindow(sWinId)
'	If Reporter.lScenarioStatus = "FAIL"  Then
'	Exit Function
'Else
	'Handle parent object
	Set mWindow = SwfWindow("micclass:=SwfWindow")
	
	'To Create object based on Window ID
	If sWinId<>"" Then
		'To Create object based on Window ID
		If Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		Else
			Set MainObj = mWindow.SwfWindow(sWinId)
		End If
		
		MainObj.Close()
	Else
		'Create object
		Dim oDesc
		Set oDesc = Description.Create
		oDesc("micclass").Value="SwfWindow"
	
		Set ChldWindows = 	SwfWindow("micclass:=SwfWindow").ChildObjects(oDesc)
		
		'Iteration to close windows
		For z = ChldWindows.Count-1 to 0 step -1
			'Set MainObj = SwfWindow("micclass:=SwfWindow","index:="&z)
			If ChldWindows(z).Exist(5) Then
		       	ChldWindows(z).Close()
				ReportWriter micDone, "Close Window "&sWinId, "Window closed successfully",0
			End If
		Next
	End If
'End If 
End Function

Function SelectRowByVal_Instr(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Row Number - Global Variable (sRowNo)
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
		For i = 0 To sRows-1
			sval=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName)
			If Instr(Trim(UCase(sval)),Trim(UCase(sExpectedVal)))>0 Then
				MainObj.SwfTable(sPropAndVal).MakeCellVisible i,1
				wait 1
				MainObj.SwfTable(sPropAndVal).SelectRow i
				x=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFieldName,"x")
				y=MainObj.SwfTable(sPropAndVal).GetCellProperty (i,sFieldName,"y")
				sRowNo= i
				Exit for
			End If
		Next
		ReportWriter micPass, "Select row based on value", "Row "&i&" has been selected based on value "&sval&"  which match with in expected value - "&sExpectedVal,0
	Else
		ReportWriter micFail, "Select row based on value", sExpectedVal&" not available in Swf Table",1
	End If
End If

End Function

Function SwfTable_SetCellData(mWindow,sWinId,sPropAndVal,RowNo,ColNo,sType)
		' @HELP new
		' @group	: SwfControls	
		' @funcion	: SwfTable_SetCellData(mWindow,sWinId,sPropAndVal,RowNo,ColNo,sType)
		' @returns	: None
		' @parameter: mWindow, sWinId, sPropAndVal, RowNo, ColNo, sType
		' @notes	: Set Data on Cell based on Row No and Col No (Eg: Set Check box True or False - sType)
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
				
		'To click cell based on row and column
		If MainObj.SwfTable(sPropAndVal).Exist(5) Then
			MainObj.SwfTable(sPropAndVal).ActivateCell RowNo,ColNo
			wait 1
			MainObj.SwfTable(sPropAndVal).SetCellData RowNo,ColNo,sType
			sRowNo= ""
			ReportWriter micPass, "Set Cell value in table", "Cell value set as : "&sType ,0
		Else
			ReportWriter micFail, "Set Cell value in table", "Unable to Set Cell Value as : "&sType ,1
		End If
		
End If

End Function

Function SwfTable_SelectCheckBox(mWindow,sWinId,sPropAndVal,RowNo)
		' @HELP new
		' @group	: SwfControls	
		' @funcion	: SwfTable_SelectCell(mWindow,sWinId,sPropAndVal,RowNo,ColNo)
		' @returns	: None
		' @parameter: mWindow, sWinId, sPropAndVal, RowNo, ColNo
		' @notes	: Select Cell based on Row No and Col No
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
		Dim oDesc
		Set oDesc=Description.Create
		oDesc("swftypename").Value="Junifer.Thor.UI.Editors.BaseCheckEdit"
		
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
				
		'To click cell based on row and column
		If MainObj.SwfTable(sPropAndVal).Exist(5) Then
			Set oChkBox = MainObj.ChildObjects(oDesc)
			oChkBox(RowNo).Click
			sRowNo= ""
			ReportWriter micDone, "Select Cell from Table", "The Cell on Row: "&RowNo&" and Column: "&ColNo,0
		Else
			ReportWriter micDone, "Select Cell from Table", sExpectedVal&" not available in Swf Table",1
		End If
		
End If

End Function



'Function JUN_SwitchSwfTab_MultiWnd(sWinId,sPropAndVal, sText)
'
'	Set MainObj = SwfWindow(micClass:=SwfWindow).SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
'	
'	If MainObj.SwfTab(sPropAndVal).Exist(5) Then
'	       MainObj.SwfTab(sPropAndVal).Object.Select sText
'		ReportWriter micDone, "Switch to "&sText&" tab", "Switch tab successful"
'	Else
'		ReportWriter MicFail, "Switch to "&sText&" tab", "Unable to switch tab"
'	End If
'	
'End Function
'
Function SwfSpin(mWindow,sPropAndVal,sText)
	' @HELP
	' @group	: SwfCalendarSetDate	
	' @funcion	: SwfCalendarSetDate(mWindow, sPropAndVal, sText)
	' @returns	: None
	' @parameter: mWindow, sPropAndVal, sText
	' @notes	: Set Date in Calendar box
	' @END
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	If mWindow.SwfSpin(sPropAndVal).Exist(10) Then
		mWindow.SwfSpin(sPropAndVal).Click
		mWindow.SwfSpin(sPropAndVal).Set sText
'		mWindow.SwfTab(sPropAndVal).DblClick
		ReportWriter micDone, "SwfTab Click", "SwfTab Selected - "&sTab,0
	Else
		ReportWriter MicFail, "SwfTab Click", "Unable to select SwfTab - "&sTab,1
	End If
End If

End Function

Function JUN_SwfSpin(mWindow,sWinId,sPropAndVal,sText)
	' @HELP
	' @group	: SwfCalendarSetDate	
	' @funcion	: SwfCalendarSetDate(mWindow, sPropAndVal, sText)
	' @returns	: None
	' @parameter: mWindow, sPropAndVal, sText
	' @notes	: Set Date in Calendar box
	' @END

'If Reporter.lScenarioStatus = "FAIL"  Then
'	Exit Function
'Else
		'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
	
	If MainObj.SwfSpin(sPropAndVal).Exist(10) Then
		MainObj.SwfSpin(sPropAndVal).Click
		MainObj.SwfSpin(sPropAndVal).Set sText
'		mWindow.SwfTab(sPropAndVal).DblClick
		ReportWriter micPass, "SwfSpin - Set Value in Edit Box", "Value "&sText&" has set",0
	Else
		ReportWriter MicFail, "SwfSpin - Set Value in Edit Box", "Unable to select SwfSpin",1
	End If
'End if

End Function

Function FetchValSwfComboBox(mWindow,sWinId,sPropAndVal,sExpectedVal)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Field Value
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfComboBox(sPropAndVal).Exist(5) Then
		MainObj.SwfComboBox(sPropAndVal).Highlight
		sVal=MainObj.SwfComboBox(sPropAndVal).GetROProperty("text")
		If Strcomp(sVal,sExpectedVal)=0 Then
			ReportWriter micPass, "Fetch value from Combo Box", "Value fetched from Combo box "&sval&" match with in expected value - "&sExpectedVal,0
		Else
			ReportWriter micFail, "Fetch value from Combo Box", sExpectedVal&" not available in Combo box",1			
		End If
		
	Else
		ReportWriter micFail, "Fetch value from Combo Box", "Combo box not available",1
	End If
	
	FetchValSwfComboBox=sVal
	
End If

End Function

Function FetchAllValSwfTable_Dictionary(mWindow,sWinId,sPropAndVal,sKey,sItem)
		' @HELP - Jagdeesh NEw
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Field Value
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
'If Reporter.lScenarioStatus = "FAIL"  Then
'			Exit Function
'Else
	Set dict = CreateObject("Scripting.Dictionary")
	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
		For i = 0 To sRows-1
			sKeyVal=MainObj.SwfTable(sPropAndVal).GetCellData(i,sKey)
			sItemVal=MainObj.SwfTable(sPropAndVal).GetCellData(i,sItem)
			MainObj.SwfTable(sPropAndVal).SelectRow i
			dict.Add sKeyVal,sItemVal
		Next
		ReportWriter micPass, "Fetch value from SwfTable", "Value fetch from Webtable "&sval&" match with in expected value - "&sExpectedVal,0
	Else
		ReportWriter micFail, "Fetch value from SwfTable", sExpectedVal&" not available in Swf Table",1
	End If
	
'End If

End Function

Function ValidateAllValSwfTable_Dictionary(mWindow,sWinId,sPropAndVal,sFieldName1,sFieldName2)
		' @HELP - Jagdeesh NEw
		' @group	: SwfControls	
		' @funcion	: ValidateAllValSwfTable_Dictionary(mWindow,sWinId,sPropAndVal,sFieldName1,sFieldName2)
		' @returns	: None
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Works only with "dict" dictionary value agaist table. sFieldName1 represents Key and sFieldName2 represents Item. key and item will be collected as array separate
					' and both array value will be validated against table fields in swf table.. sFieldName1 is Field1 and sFieldName2 is Field2
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
	
	sKeys=dict.Keys
	sItems=dict.items
			
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		
		For z = 0 To UBound(sKeys)
			
			sRows=MainObj.SwfTable(sPropAndVal).GetRowsCount
			
			For i = 0 To sRows-1
				sVal1=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName1)
				sVal2=MainObj.SwfTable(sPropAndVal).GetCellData(i,sFieldName2)
				MainObj.SwfTable(sPropAndVal).SelectRow i
				If Instr(sVal1,sKeys(z))>=0 and Instr(sVal2,sItems(z))>=0 Then
					ReportWriter micPass, "Validate value on SwfTable", "Value fetch from Webtable "&sVal1&"-"&sVal2&" match with in expected value - "&sKeys(i)&"-"&sItems(i),0
					Exit for
				Else
					ReportWriter micFail, "Validate value on SwfTable", "Value fetch from Webtable "&sVal1&"-"&sVal2&" does not match with in expected value - "&sKeys(i)&"-"&sItems(i),1
				End If
			Next
		Next
		
	Else
		ReportWriter micFail, "Fetch value from SwfTable", "No Rows found on Swf Table",1
	End If
	
End If

End Function

Function ValidateValSwfTable_ByRowNo(mWindow,sWinId,sPropAndVal,sRowNo,sFieldName1,sExpectedVal)
		' @HELP - Jagdeesh NEw
		' @group	: SwfControls	
		' @funcion	: ValidateAllValSwfTable_Dictionary(mWindow,sWinId,sPropAndVal,sFieldName1,sFieldName2)
		' @returns	: None
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Works only with "dict" dictionary value agaist table. sFieldName1 represents Key and sFieldName2 represents Item. key and item will be collected as array separate
					' and both array value will be validated against table fields in swf table.. sFieldName1 is Field1 and sFieldName2 is Field2
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
			
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		sVal1=MainObj.SwfTable(sPropAndVal).GetCellData(sRowNo,sFieldName1)
		MainObj.SwfTable(sPropAndVal).SelectRow sRowNo
		If Instr(Trim(UCase(sVal1)),Trim(UCase(sExpectedVal)))>=0 Then
			ReportWriter micPass, "Validate value on SwfTable By Row NO", "Value fetch from Webtable "&sVal1&" match with in expected value - "&sExpectedVal,0
		Else
			ReportWriter micFail, "Validate value on SwfTable By Row NO", "Value fetch from Webtable "&sVal1&" does not match with in expected value - "&sExpectedVal,1
		End If	
	Else
		ReportWriter micFail, "Fetch value from SwfTable By Row NO", "No Rows found on Swf Table",1
	End If
	
End If

End Function

Function SelectRowByNumSwfTable(mWindow,sWinId,sPropAndVal,sRW)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SelectRowByValSwfTable(mWindow,sWinId,sPropAndVal,sFieldName,sExpectedVal)
		' @returns	: Row Number - Global Variable (sRowNo)
		' @parameter: mWindow, sWinId, sPropAndVal, sFieldName, sExpectedVal
		' @notes	: Select row from table based on Field name "sFieldName" and Value "sExpectedVal" 
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else	
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
	If MainObj.SwfTable(sPropAndVal).Exist(5) Then
		
		MainObj.SwfTable(sPropAndVal).SelectRow sRW
		x=MainObj.SwfTable(sPropAndVal).GetCellProperty (sRW,sFieldName,"x")
		y=MainObj.SwfTable(sPropAndVal).GetCellProperty (sRW,sFieldName,"y")
		'sRowNo= i
				
'		If sRowNo<>"" Then
		ReportWriter micPass, "Select value from table", "Selected Value from Webtable from row"&sval&" match with in expected value - "&sExpectedVal,0
'		Else
'			ReportWriter micFail, "Select value from table", "Unable to select row since Expected Value - "&sExpectedVal&", not available in the Table",1		
'		End If
		
	Else
		ReportWriter micFail, "Select value from Table", "Table Unavailable",1
	End If
	
End If

End Function

Function JUN_SwfRadioButton(mWindow, sWinId, sPropAndVal, sText)
		' @HELP
		' @group	: SwfControls	
		' @funcion	: SwfClickButton(mWindow, sPropAndVal, sRow)
		' @returns	: None
		' @parameter: mWindow, sPropAndVal, sRow
		' @notes	: Click SwfButton
		' @END
	If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
	Else		
		'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
	
		If MainObj.SwfRadioButton(sPropAndVal).Exist(10) Then
			MainObj.SwfRadioButton(sPropAndVal).Select sText
			ReportWriter micPass, "Select Radio button", sText&" Radio Button Clicked sucessfully",0 '"Uncomment this"
		Else
			ReportWriter MicFail, "Select Radio button", "Unable to select "&sText&" Radio Button",1 '"Uncomment this" Set value as '1' for failure after code update
		End If
	End If
	
End Function

Function JUN_RightClickTableandSelectMenu(mWindow,sWinId,sPropAndVal)
	' @HELP
	' @group	: SwfCalendarSetDate	
	' @funcion	: SwfCalendarSetDate(mWindow, sPropAndVal, sText)
	' @returns	: None
	' @parameter: mWindow, sPropAndVal, sText
	' @notes	: Set Date in Calendar box
	' @END
	
	'To Create object based on Window ID
If Reporter.lScenarioStatus = "FAIL"  Then
		Exit Function
	Else
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
		
			If MainObj.SwfSpin(sPropAndVal).Exist(10) Then
				 
'				 JUN_WINDOW.SwfWindow("window id:=2").SwfTable(TICKET_TABLE).Click x,y,micRightBtn
				 
				  MainObj.SwfTable(sPropAndVal).Click x,y,micRightBtn
				  
				ReportWriter micDone, "SwfSpin Click", "SwfSpin Selected",0
			Else
				ReportWriter MicFail, "SwfSpin Click", "Unable to select SwfSpin",1
			End If
End If
End Function



''*******************************************
'C Purushotham
'25 -10-2022
''*******************************************
Function SwfTable_SelectCheckBox_Billing(mWindow,sWinId,sPropAndVal)
		' @HELP new
		' @group	: SwfControls	
		' @funcion	: SwfTable_SelectCell(mWindow,sWinId,sPropAndVal,RowNo,ColNo)
		' @returns	: None
		' @parameter: mWindow, sWinId, sPropAndVal, RowNo, ColNo
		' @notes	: Select Cell based on Row No and Col No
		' @END
		
If Reporter.lScenarioStatus = "FAIL"  Then
			Exit Function
Else
		Dim oDesc
		Set oDesc=Description.Create
		oDesc("swftypename").Value="Junifer.Thor.UI.Editors.BaseCheckEdit"		
	'To Create object based on Window ID
		If sWinId<>"" and Isnumeric(sWinId) Then
			Set MainObj = mWindow.SwfWindow("micClass:=SwfWindow","window id:="&sWinId)
		ElseIf sWinId<>"" Then
			Set MainObj = mWindow.SwfWindow(sWinId)
		Else
			Set MainObj = mWindow
		End If
				
		'To click cell based on row and column
		If MainObj.SwfTable(sPropAndVal).Exist(5) Then
			MainObj.SwfTable(sPropAndVal).highlight
			wait 2
			MainObj.SwfTable(sPropAndVal).click
			wait 3
			ReportWriter micPass, "UnSelect Check Box", "Step Passed",0
		Else
			ReportWriter micDone, "UnSelect Check Box", "Step Failed",1
		End If
		
End If

End Function


