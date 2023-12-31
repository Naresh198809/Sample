'***********************************************************************************************************************************************************************************************
'**  Project		:	ATLaS
'**  File Name		:	Controls_SAP.vbs
'**  Version		:	1.0
'**  Created on		:	20th September 2012
'**  Updated on		:	20th September 2012
'**  Description	:	Common Methods for executing the SAP functionality in the scripts across Business Units
'**  Copyright		:	Arsin Corporation.
'**  Author			:	
'***********************************************************************************************************************************************************************************************
Class clsControlsSAP
'***********************************************************************************************************************************************************************************************
'''''*****************Methods related to Data arrays - Extract Data from Effecta Database *****************************************
''''''*****************************************************************************************************************************
	Function EffectaDateToStandardDateConvertor (sDateString)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	EffectaDateToStandardDateConvertor (sDateString)
			' @returns	:	EffectaDateToStandardDateConvertor 	: DateString
			' @parameter:	sDateString							: Date in Effecta Format
			' @notes	:	This method converts Date in Effecta format to standard date format
			' @END
			sDateSeperator = "/"
			If  IsNumeric(Mid(sDateString,8))Then
				If (UCase(Mid(sDateString,1,4)) = "DATE") Then
					If (UCase(Mid(sDateString,5,2)) = "TP") Then 
						Select Case UCase(Mid(sDateString,7,1)) 
							Case "D"
								sDateString = DateAdd("d",Mid(sDateString,8),date)
							Case "M"
								sDateString = DateAdd("m",Mid(sDateString,8),date)
							Case "Y"
								sDateString = DateAdd("yyyy",Mid(sDateString,8),date)
						End Select
						' aPCDateArray = Split(sDateString, "/", -1, 1)
						' if len(aPCDateArray(0)) = 1 then
						' aPCDateArray(0) = "0"&aPCDateArray(0)
						' end if
						' if len(aPCDateArray(1)) = 1 then
						' aPCDateArray(1) = "0"&aPCDateArray(1)
						' end if
						' sPCdate = aPCDateArray(0) & "/" & aPCDateArray(1) & "/" & aPCDateArray(2)
						sPCdate = month(sDateString) & sDateSeperator & Day(sDateString) & sDateSeperator & year(sDateString)
						EffectaDateToStandardDateConvertor = sPCdate	
					ElseIf (UCase(Mid(sDateString,5,2)) = "TM") Then 
						Select Case UCase(Mid(sDateString,7,1)) 
							Case "D"
								sDateString = DateAdd("d",-Mid(sDateString,8),date)
							Case "M"
								sDateString = DateAdd("m",-Mid(sDateString,8),date)
							Case "Y"
								sDateString = DateAdd("yyyy",-Mid(sDateString,8),date)
						End Select
						' aPCDateArray = Split(sDateString, "/", -1, 1)
						' if len(aPCDateArray(0)) = 1 then
						' aPCDateArray(0) = "0"&aPCDateArray(0)
						' end if
						' if len(aPCDateArray(1)) = 1 then
						' aPCDateArray(1) = "0"&aPCDateArray(1)
						' end if
						' sPCdate = aPCDateArray(0) & "/" & aPCDateArray(1) & "/" & aPCDateArray(2)
						sPCdate = month(sDateString) & sDateSeperator & Day(sDateString) & sDateSeperator & year(sDateString)
						EffectaDateToStandardDateConvertor = sPCdate	
					else
						EffectaDateToStandardDateConvertor = sDateString
					End If			 
					Else
					EffectaDateToStandardDateConvertor = sDateString
				End if
			Else
				EffectaDateToStandardDateConvertor = sDateString
			End If
		End Function
		'***************************************************************************************************************************
		Function ConvertCustomDataToActualInputData (aDataSetData)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	ConvertCustomDataToActualInputData (sDateString)
			' @returns	:	ConvertCustomDataToActualInputData 	: DateString
			' @parameter:	sDateString							: Date in Effecta Format
			' @notes	:	This method converts Date in Effecta format to standard date format
			' @END
			sDateSeperator = "/"
			If (UCase(Mid(aDataSetData,1,4)) = "DATE") Then
				If (UCase(Mid(aDataSetData,5,2)) = "TP") Then 
					Select Case UCase(Mid(aDataSetData,7,1)) 
						Case "D"
							aDataSetData = DateAdd("d",Mid(aDataSetData,8),date)
						Case "M"
							aDataSetData = DateAdd("m",Mid(aDataSetData,8),date)
						Case "Y"
							aDataSetData = DateAdd("yyyy",Mid(aDataSetData,8),date)
					End Select
					REM aPCDateArray = Split(aDataSetData, "/", -1, 1)
					REM sPCdate = aPCDateArray(0) & "/" & aPCDateArray(1) & "/" & aPCDateArray(2)
					sPCdate =  month(aDataSetData)& sDateSeperator & Day(aDataSetData)& sDateSeperator &  year(aDataSetData)
					ConvertCustomDataToActualInputData = sPCdate	
				ElseIf (UCase(Mid(aDataSetData,5,2)) = "TM") Then 
					Select Case UCase(Mid(aDataSetData,7,1)) 
						Case "D"
							aDataSetData = DateAdd("d",-Mid(aDataSetData,8),date)
						Case "M"
							aDataSetData = DateAdd("m",-Mid(aDataSetData,8),date)
						Case "Y"
							aDataSetData = DateAdd("yyyy",-Mid(aDataSetData,8),date)
					End Select
					REM aPCDateArray = Split(aDataSetData, "/", -1, 1)
					REM sPCdate = aPCDateArray(0) & "/" & aPCDateArray(1) & "/" & aPCDateArray(2)
					' sPCdate =  month(aDataSetData)& "/" & Day(aDataSetData)& "/" &  year(aDataSetData)
					sPCdate =  month(aDataSetData)& sDateSeperator & Day(aDataSetData)& sDateSeperator &  year(aDataSetData)
					ConvertCustomDataToActualInputData = sPCdate	
				else
					ConvertCustomDataToActualInputData = aDataSetData
				End If			 
			else
				ConvertCustomDataToActualInputData = aDataSetData
			end if
			If (UCase(Mid(aDataSetData,1,6)) = "RANDOM") Then
					aRandom = Split(aDataSetData,"_")
					iLength = aRandom(2) - Len(aRandom(1))
					Maximum = 9 
					For i = 1 to iLength-1
						Maximum = Maximum & 9
					Next
				Randomize 
			   ConvertCustomDataToActualInputData = aRandom(1) &  Int((Maximum * Rnd) + 1)   ' Generate random value 
			End If
			Select Case UCase(aDataSetData)
				Case "CURRENTYEAR"
					ConvertCustomDataToActualInputData = Year(now)
				Case "LASTYEAR"
					ConvertCustomDataToActualInputData = Year(now) - 1
				Case "NEXTYEAR"
					ConvertCustomDataToActualInputData = Year(now) + 1
				Case "CURRENTYEARMINUS2"
					ConvertCustomDataToActualInputData = Year(now) - 2
				Case "YEAREND"
					ConvertCustomDataToActualInputData  = "12/31/" & Year(now)
				Case "LASTYEAREND"
					ConvertCustomDataToActualInputData  = "12/31/" & Year(now)-1
				Case "NEXTYEAREND"
					ConvertCustomDataToActualInputData  = "12/31/" & Year(now)+1
				Case "BEGINOFCURRENTYEAR"
					ConvertCustomDataToActualInputData  = "01/01/" & Year(now)
				Case "BEGINOFLASTYEAR"
					ConvertCustomDataToActualInputData  = "01/01/" & Year(now)-1
				Case "BEGINOFNEXTYEAR"
					ConvertCustomDataToActualInputData  = "01/01/" & Year(now)+1
			End Select
		End Function
		'***************************************************************************************************************************
		Function GetHeaderScreenData (sScreenName, aInputData)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetScreenData (sScreenName, aInputData)
			' @returns	:	GetScreenData 	: Array
			' @parameter:	sScreenName		: Name of the screen which gets activated on selecting the function
			' @parameter:	alnputData 		: Array that contains component data
			' @notes	:	This method returns Input data of a particular screen from a component data array
			' @END
			''On error Resume Next
			 Dim aData ()
			Dim iCount,idCount
			idCount = 0
			For iRow = 0 to UBound(aInputData)-1
				If  (aInputData(iRow,0)= sScreenName )Then
					idCount = idCount  + 1
				End If
			Next
			
			ReDim aData (idCount,4)
			
			iCount = 0
			For iRow = 0 to UBound(aInputData)-1
				aInputData(iRow,6) = ConvertCustomDataToActualInputData (aInputData(iRow,6))
				If (aInputData(iRow,0) = sScreenName ) Then
					If aInputData(iRow,1)="O" Then 
						aData (iCount,0) = aInputData (iRow,0)
						aData (iCount,1) = aInputData (iRow,5)
						aData (iCount,2) = aInputData (iRow,10)
						aData (iCount,3) = aInputData (iRow,6)
						aData (iCount,4) = aInputData (iRow,11)
						iCount = iCount  + 1
					Else
						aData (iCount,0) = aInputData (iRow,0)
						aData (iCount,1) = aInputData (iRow,4)
						aData (iCount,2) = aInputData (iRow,3)
						aData (iCount,3) = aInputData (iRow,6)
						aData (iCount,4) = aInputData (iRow,11)
						iCount = iCount  + 1
					End If 
				End If		    
			Next
			GetHeaderScreenData = aData
			Err.Clear
			On Error Goto 0
		End Function
		'***************************************************************************************************************************
		Function GetLineItemScreenData (sScreenName, aInputData)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetLineItemData (sScreenName, aInputData)
			' @returns	:	GetLineItemData :	Array
			' @parameter:   sScreenName  	:	Name of the screen which gets activated on selecting the function
			' @parameter:   sInputData   	:	Array that contains component data 
			' @notes	:	Returns LineItemData of a particular screen from Component data
			' @END
			On error resume next
			Dim aData ()
			Dim iCount,idCount
			For iRow = 0 to UBound(aInputData)-1
				If  (aInputData(iRow,0)= sScreenName ) and (aInputData (iRow,1) = "L") Then
				idCount = idCount  + 1
				End If
			Next
			ReDim aData(idCount-1,6)
			For iRow = 0 to UBound(aInputData)-1
				aInputData(iRow,6) = ConvertCustomDataToActualInputData(aInputData(iRow,6))
				If  (aInputData(iRow,0)= sScreenName ) and aInputData (iRow,1) = "L" Then
						aData (iCount,0) = aInputData (iRow,0)
						aData (iCount,1) = aInputData (iRow,4)
						aData (iCount,2) = aInputData (iRow,2)
						aData (iCount,3) = aInputData (iRow,5)
						aData (iCount,4) = aInputData (iRow,6)
						iCount = iCount  + 1
				End If
			Next
			GetLineItemScreenData = aData
			On error goto 0		
		End Function
		'***************************************************************************************************************************
		Function GetOffSetFieldsData (sScreenName, aInputData)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetOffSetFieldsData (sScreenName, aInputData)
			' @returns	:	GetOffSetFieldsData :	Array
			' @parameter:	sScreenName			:	Name of the screen which gets activated on selecting the function
			' @parameter:	alnputData			:	Array that contains component data
			' @notes	:	This method returns Offset data of a particular screen from a component data array
			' @END	
			Dim aData ()
			On error resume next
			Dim iCount,idCount
			For iRow = 0 to UBound(aInputData)-1
				If  (aInputData(iRow,0)= sScreenName )Then
					idCount = idCount  + 1
				End If
			Next
			ReDim aData(idCount-1,6)
			For iRow = 0 to UBound(aInputData)-1
				If  (aInputData(iRow,0)= sScreenName )Then
					If aInputData(iRow,1)="O" Then 
						aData (iCount,0) = aInputData (iRow,1)
						aData (iCount,1) = aInputData (iRow,4)
						aData (iCount,2) = aInputData (iRow,2)
						aData (iCount,3) = aInputData (iRow,5)
						aData (iCount,4) = aInputData (iRow,9)
						aData (iCount,5) = aInputData (iRow,10)								
						aData (iCount,6) = EffectaDateToStandardDateConvertor(aInputData (iRow,6))
					Else
						aData (iCount,0) = aInputData (iRow,0)
						aData (iCount,1) = aInputData (iRow,4)
						aData (iCount,2) = aInputData (iRow,2)
						aData (iCount,3) = aInputData (iRow,6)				
						aData (iCount,4) = aInputData (iRow,4)
						aData (iCount,5) = aInputData (iRow,3)
						aData (iCount,6) = EffectaDateToStandardDateConvertor(aInputData (iRow,6))
					End If 
					iCount = iCount  + 1
				End If
			Next
			GetOffSetFieldsData = aData
			Err.Clear
			On error goto 0
		End Function
		'***************************************************************************************************************************
		Function GetFieldValue (sFieldName, aInputData)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetFieldValue (sFiledName, aInputData)
			' @returns	:	GetFiledValue	:	String  
			' @parameter:   sFieldName		:	Name of the field to which value is to retreived 
			' @parameter:   alnputData		:	DataArray that contains InputData
			' @notes	:	This Method returns a field value for a given Field name from a data array	 
			' @END
			On error resume next
			For iRow = 0 to UBound(aInputData)
				If  (aInputData(iRow,1)= sFieldName)Then
					GetFieldValue = aInputData(iRow,3)
					exit for
				End If
			Next
			Err.Clear
			On error goto 0	
		End Function
		''''''************** End of Extract Functions ****************************************************************************************

		'''''''Methods related to performing actions on group of Objects on single SAP Screen
		'*************************************************************************************************************************************
		Sub SetSAPHeaderScreenData (sDataArray)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPScreenData (sDataArray)
			' @parameter:   sDataArray	:	Array that contains Screen data
			' @notes	:	Inserts data on SAP Screen
			' @END
			For iRow = 0 to UBound(sDataArray)
				If (UCase(Mid(sDataArray(iRow,3),1,4)) = "DATE") then
					If (UCase(Mid(sDataArray(iRow,3),5,2)) = "TP") then 
						Select Case UCase(Mid(sDataArray(iRow,3),7,1)) 
							Case "D"
								sDataArray(iRow,3) = DateAdd("d",Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
							Case "M"
								sDataArray(iRow,3) = DateAdd("m",Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
							Case "Y"
								sDataArray(iRow,3) = DateAdd("yyyy",Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))		
						End Select	
					ElseIf (UCase(Mid(sDataArray(iRow,3),5,2)) = "TM") then 
						Select Case UCase(Mid(sDataArray(iRow,3),7,1)) 
							Case "D"
								sDataArray(iRow,3) = DateAdd("d",-Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
							Case "M"
								sDataArray(iRow,3) = DateAdd("m",-Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
							Case "Y"
								sDataArray(iRow,3) = DateAdd("yyyy",-Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
						End Select			
					End If 	
				End if
				Select Case sDataArray(iRow,2)
					Case 34 
						SetSAPComboBox sDataArray(iRow,1),sDataArray(iRow,4),sDataArray(iRow,3)
					Case 32,31
						SetSAPEdit sDataArray(iRow,1),sDataArray(iRow,4),sDataArray(iRow,3)
					Case 42
						SetSAPCheckBox sDataArray(iRow,1),sDataArray(iRow,4),sDataArray(iRow,3)
					Case 41
						SelectSAPRadioButton sDataArray(iRow,1),sDataArray(iRow,3)
				End Select 
			Next
		End Sub	

		'*****************************************************************************************************************************
		
		Sub SetSAPModalScreenData (sDataArray)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPScreenData (sDataArray)
			' @parameter:   sDataArray	:	Array that contains Screen data
			' @notes	:	Inserts data on SAP Screen
			' @END

			For iRow = 0 to UBound(sDataArray)
				If (UCase(Mid(sDataArray(iRow,3),1,4)) = "DATE") then
					If (UCase(Mid(sDataArray(iRow,3),5,2)) = "TP") then 
						Select Case UCase(Mid(sDataArray(iRow,3),7,1)) 
							Case "D"
								sDataArray(iRow,3) = DateAdd("d",Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
							Case "M"
								sDataArray(iRow,3) = DateAdd("m",Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
							Case "Y"
								sDataArray(iRow,3) = DateAdd("yyyy",Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))		
						End Select	
					ElseIf (UCase(Mid(sDataArray(iRow,3),5,2)) = "TM") then 
						Select Case UCase(Mid(sDataArray(iRow,3),7,1)) 
							Case "D"
								sDataArray(iRow,3) = DateAdd("d",-Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
							Case "M"
								sDataArray(iRow,3) = DateAdd("m",-Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
							Case "Y"
								sDataArray(iRow,3) = DateAdd("yyyy",-Mid(sDataArray(iRow,3),8),CDate(Environment.Value("ExecDate")))
						End Select			
					End If 	
				End if
				Select Case sDataArray(iRow,2)
					Case 34 
						SetSAPModalComboBox sDataArray(iRow,1),sDataArray(iRow,4),sDataArray(iRow,3)
					Case 32,31 
						SetSAPModalEdit sDataArray(iRow,1),sDataArray(iRow,4),sDataArray(iRow,3)
					Case 42
						SetSAPModalCheckBox sDataArray(iRow,1),sDataArray(iRow,4),sDataArray(iRow,3)
				End Select 
			Next
            CaptureSAPScreenShot()
		End Sub
		
		'*****************************************************************************************************************************
		
		Sub SetSAPTableData (sSAPTable,aDataArray)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPTableData (sSAPTable,aDataArray)
			' @parameter:   sSAPTable   :	Name of the Table
			' @parameter:   aDataArray  :	Array of fields and field values
			' @notes	: 	Sets the data into SAP Table
			' @END 
			''On Error Resume Next

			If InStr(sSAPTable,"_") > 0 or InStr(sSAPTable,"TABLE") > 0 or InStr(sSAPTable,"SAP") then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable				
			End if
			' MsgBox "SetSAPTableData"
			err.clear
			iTableRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).Rowcount'GetSAPTableRowCount (sSAPTable)
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			' MsgBox UBound(aDataArray)
			For iRow = 0 to UBound(aDataArray)
				
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).SetCellData aDataArray(iRow,2),aDataArray(iRow,3),aDataArray(iRow,4)  			
				
				if err.Description = "" then
				
					UpdateExecutionReport micPass, "SetSAPTableData", "Value '"&aDataArray(iRow,4)&"' set in Column '"&aDataArray(iRow,3)&"' of Table '"&sSAPTable&"' of Screen '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "SetSAPTableData", "Value '"&aDataArray(iRow,4)&"' couldnot be set in Column '"&aDataArray(iRow,3)&"' of Table '"&sSAPTable&"' of Screen '"&sScreenName&"' ### Error Description:"&err.description
				end if
				err.clear
				If iRow < UBound(aDataArray) Then
					If (aDataArray(iRow,2) <> aDataArray(iRow+1,2)) Then								
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SendKey ENTER
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SendKey ENTER
						If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").exist(1)	Then
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Close
						end	if
						iTableRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).Rowcount'GetSAPTableRowCount (sSAPTable)
					End If	
				Else							
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SendKey ENTER
				End If
			Next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").exist(1)	Then
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Close
			end	if
			err.clear
			On error Goto 0
		End Sub
		
		'************************************************************************************************************************************************************************************
		
		Sub SetSAPModalTableData (sSAPTable,aDataArray)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPModalTableData (sSAPTable,aDataArray)
			' @parameter:   sSAPTable   :	Name of the Table
			' @parameter:   aDataArray  :	Array of fields and field values
			' @notes	: 	Sets the data into SAP Table
			' @END 
			On Error Resume Next
			If InStr(sSAPTable,"_") > 0 or InStr(sSAPTable,"TABLE") > 0 or InStr(sSAPTable,"SAP") then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			err.clear
			iTableRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).Rowcount'GetSAPTableRowCount (sSAPTable)
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			err.clear
			For iRow = 0 to UBound(aDataArray)		
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).SetCellData aDataArray(iRow,2),aDataArray(iRow,3),aDataArray(iRow,4)
				if err.Description = "" then
					UpdateExecutionReport micPass, "SetSAPModalTableData", "Value '"&aDataArray(iRow,4)&"' set in Column '"&aDataArray(iRow,3)&"' of Table '"&sSAPTable&"' of Screen '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "SetSAPModalTableData", "Value '"&aDataArray(iRow,4)&"' couldnot be set in Column '"&aDataArray(iRow,3)&"' of Table '"&sSAPTable&"' of Screen '"&sScreenName&"' ### Error Description:"&err.description
				end if
				err.clear
				If iRow < UBound(aDataArray) Then
					If (aDataArray(iRow,2) <> aDataArray(iRow+1,2)) Then								
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SendKey ENTER
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SendKey ENTER
						Rem If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").exist(1)	Then
							Rem SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Close
						Rem end	if
						iTableRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).Rowcount'GetSAPTableRowCount (sSAPTable)
					End If	
				Else							
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SendKey ENTER
				End If
				err.clear
			Next
			Rem If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").exist(1)	Then
				Rem SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Close
			Rem end	if
			err.clear
			On error Goto 0
            CaptureSAPScreenShot()
		End Sub
		'*****************************************************************************************************************************
		Sub SetSAPOffsetData ( aDataArray)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPOffsetData ( aDataArray)
			' @parameter:	sWindow		:	Name of the Session
			' @parameter:   aDataArray	:	Array of Fields and field values 
			' @notes	:	Sets data in to Stepped loop fields 
			' @END 
			On error resume next

			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			For iRow = 0 to UBound(aDataArray)
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="& aDataArray(iRow,4),"guicomponenttype:="& aDataArray(iRow,5),"index:="& aDataArray(iRow,2)-1).Set aDataArray(iRow,6)
			
			Next
			If err.description = "" Then
			UpdateExecutionReport micPass, "SetSAPOffsetData","Screen:"&sScreenName&"### Data set into OffSet fields successfully"
			else
			UpdateExecutionReport micFail, "SetSAPOffsetData","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0	
            CaptureSAPScreenShot()
		End Sub
		
		'*************************************************************************************************************************************
		''***********************End of group action Functions********************************************************************************

		'Methods related to performing actions on Individual Objects on SAP Screens
''***********************************************************************************************************************************************************************************************
Rem Function SAPButtonRelatedFunctions()
Function ClickSAPButton(sButton)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	ClickSAPButton( sButton)
	' @parameter:   sButton	:	name/Text/Tooltip property of the Button
	' @notes	:	Clicks the button on main window using Text/Tooltip property	 	
	' @END
'            CaptureSAPScreenShot()
	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist(1) Then
	set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
	sScreenName = oSAPObj.GetROProperty("text")
		If oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).Exist Then 
			sDes=oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).GetRoProperty("text")
			If sDes="" Then
				sDes=oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).GetRoProperty("tooltip")
			End If
			oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).click
			UpdateExecutionReport micPass, "ClickSAPButton","Click action performed on Button '"&sDes&"' in Screen '"&sScreenName&"'"
		ElseIf oSAPObj.SAPGuiButton("type:=GuiButton","text:="&sButton).Exist Then 
			sDes=oSAPObj.SAPGuiButton("type:=GuiButton","text:="&sButton).GetRoProperty("text")
			If sDes="" Then
				sDes=oSAPObj.SAPGuiButton("type:=GuiButton","text:="&sButton).GetRoProperty("tooltip")
			End If
			oSAPObj.SAPGuiButton("type:=GuiButton","text:="&sButton).click
			UpdateExecutionReport micPass, "ClickSAPButton","Click action performed on Button '"&sDes&"' in Screen '"&sScreenName&"'"
		ElseIf oSAPObj.SAPGuiButton("type:=GuiButton","tooltip:="&sButton).Exist Then 
			sDes=oSAPObj.SAPGuiButton("type:=GuiButton","tooltip:="&sButton).GetRoProperty("text")
			If sDes="" Then
				sDes=oSAPObj.SAPGuiButton("type:=GuiButton","tooltip:="&sButton).GetRoProperty("tooltip")
			End If
			oSAPObj.SAPGuiButton("type:=GuiButton","tooltip:="&sButton).click
			UpdateExecutionReport micPass, "ClickSAPButton","Click action performed on Button '"&sDes&"' in Screen '"&sScreenName&"'"
		else
			UpdateExecutionReport micFail, "ClickSAPButton","Click action couldnot be performed on Button '"&sDes&"' in Screen '"&sScreenName&"'"
		End If
	ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
	sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
	UpdateExecutionReport micFail, "ClickSAPButton","Click action couldnot be performed on Button '"&sButton&"' as Active Screen is a Modal Window '"&sScreenName&"'"
	End If
	If err.description <> "" Then
	UpdateExecutionReport micFail, "ClickSAPButton","Button:"&sButton&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
	End If
	Err.Clear
	On error goto 0
End Function

'***********************************************************************************************************************************************************************************************

Function GetSAPButtonProperty(sButton,sROProperty)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	GetSAPButtonProperty(sButton,sROProperty)
	' @parameter:   sButton	:	name/property of the Button
	' @notes	:	Clicks the button on main window using Text/Tooltip property	 	
	' @END
	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
	Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
	sScreenName = oSAPObj.GetROProperty("text")
		If oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).Exist Then 
		GetSAPButtonProperty = oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).GetROProperty(sROProperty)
		UpdateExecutionReport micPass, "GetSAPButtonProperty","Value: "& GetSAPButtonProperty &" for Property  "& sROProperty &" found for Button with name " &sButton&"' in Screen '"& sScreenName
		End if
	ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
	Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		sScreenName = oSAPObj.GetROProperty("text")
		If oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).Exist Then 
		GetSAPButtonProperty = oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).GetROProperty(sROProperty)
		UpdateExecutionReport micPass, "GetSAPButtonProperty","Value '"&GetSAPButtonProperty&"'for Property '"&sROProperty&"' found for Button with name '" & sButton&"' in Screen '"&sScreenName&"'"
		End if
	End If
	If err.description <> "" Then
		UpdateExecutionReport micFail, "GetSAPButtonProperty","Button with Index '"&iIndex&" and Tooltip '"&sTooltip&"' ### Error Description:"&err.description
	End If
	Err.Clear
	On error goto 0
End Function

'***********************************************************************************************************************************************************************************************

Function VerifySAPButton( sButton)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	VerifySAPButton( sButton)
	' @parameter:   sButton	:	Text/Tooltip property of the Button
	' @notes:		Verify  the button existance on Main window using Text/Tooltip property	 	
	' @END
	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
	Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
	sScreenName = oSAPObj.GetROProperty("text")
		If oSAPObj.SAPGuiButton("type:=GuiButton","name:="&sButton).Exist Then 
			VerifySAPButton = True
			UpdateExecutionReport micPass, "VerifySAPButton","Button '"&sButton&"' exists in Screen '"&sScreenName&"'"
		ElseIf oSAPObj.SAPGuiButton("type:=GuiButton","text:="&sButton).Exist Then 
			VerifySAPButton = True
			UpdateExecutionReport micPass, "VerifySAPButton","Button '"&sButton&"' exists in Screen '"&sScreenName&"'"
		ElseIf oSAPObj.SAPGuiButton("type:=GuiButton","tooltip:="&sButton).Exist Then 
			VerifySAPButton = True
			UpdateExecutionReport micPass, "VerifySAPButton","Button '"&sButton&"' ' exists in Screen '"&sScreenName&"'"
		else
			VerifySAPButton = False
			UpdateExecutionReport micPass, "VerifySAPButton","Button '"&sButton&"' doesnot exist in Screen '"&sScreenName&"'"
		End If
	ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
		sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
		UpdateExecutionReport micFail, "VerifySAPButton","Button: '"&sButton&"' ### Active Window is Modal Window '"&sScreenName&"'"
	End If
	If err.description <> "" Then
		UpdateExecutionReport micFail, "VerifySAPButton","Button:'"&sButton&"' ### Screen: '"&sScreenName&"' ###Error Description:"&err.description
	End If
	Err.Clear
	On error goto 0
End Function


'**************************************************************************************************************************************************************************		
Rem End Function 
'***********************************************************************************************************************************************************************************************

	Rem Function SAPCheckBoxRelatedFunctions()
		Function VerifySAPCheckBox (sTechName)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPCheckBox ( sTechName)
			' @parameter:   sTechName	:	name property of the Checkbox
			' @notes:		Verify  the Checkbox existance on Main window using name property	 	
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheckBox,"guicomponenttype:=42").Exist Then
					VerifySAPCheckBox = True
					UpdateExecutionReport micPass, "VerifySAPCheckBox", "SAP Checkbox '"&sTechName&"' exist in Main Window '"&sScreenName&"'"
				Else
					VerifySAPCheckBox = False
					UpdateExecutionReport micPass, "VerifySAPCheckBox", "SAP Checkbox '"&sTechName&"' doesnot exist in Main Window '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "VerifySAPCheckBox", "SAP Checkbox: '"&sTechName&"' ### Active Window is Modal Window '"&sScreenName&"'"
			End If
			If err.description <> "" Then
				UpdateExecutionReport micFail, "VerifySAPCheckBox","SAP Checkbox: '"&sTechName&"' ### Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0
		End Function

		'***********************************************************************************************************************************************************************************************

		Function VerifySAPModalCheckBox (sTechName)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPModalCheckBox ( sTechName)
			' @parameter:   sTechName	:	name property of the Checkbox
			' @notes:		Verify  the Checkbox existance on Modal window using name property	 	
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheckBox,"guicomponenttype:=42").Exist Then
					VerifySAPModalCheckBox = True
					UpdateExecutionReport micPass, "VerifySAPModalCheckBox", "SAP Checkbox '"&sTechName&"' exist in Modal Window '"&sScreenName&"'"
				Else
					VerifySAPModalCheckBox = False
					UpdateExecutionReport micPass, "VerifySAPModalCheckBox", "SAP Checkbox '"&sTechName&"' doesnot exist in Modal Window '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "VerifySAPModalCheckBox", "SAP Checkbox: '"&sTechName&"' ### Active Window is Main Window '"&sScreenName&"'"
			End If
			If err.description <> "" Then
				UpdateExecutionReport micFail, "VerifySAPModalCheckBox","SAP Checkbox: '"&sTechName&"' ### Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPCheckBox ( sSAPGuiCheckBox,sDescription,sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPCheckBox ( sSAPGuiCheckBox,sDescription, sValue)
			' @parameter:   sSAPGuiCheckBox		:	name property of the Checkbox
			' @parameter:   sDescription	:	Description of CheckBox
			' @parameter:   sValue			:	ON/OFF
				' @notes:		Set the Checkbox existance on Main window using name property	 	
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheckBox,"guicomponenttype:=42").Exist() Then
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheckBox,"guicomponenttype:=42").Set sValue
				End If
				If err.description = "" Then
					UpdateExecutionReport micPass, "SetSAPCheckBox", "Value '"&sValue&"'set in SAP Checkbox '"& sDescription &"' in Main Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("text:="&sSAPGuiCheckBox,"guicomponenttype:=42").Exist() Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("text:="&sSAPGuiCheckBox,"guicomponenttype:=42").Set sValue
						UpdateExecutionReport micPass, "SetSAPCheckBox", "Value '"&sValue&"'set in SAP Checkbox '"& sDescription &"' in Main Window '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SetSAPCheckBox","SAP Checkbox: '"& sDescription &"' ### Error Description:"&err.description
					End If	
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPCheckBox", "SAP Checkbox: '"& sDescription &"' ### Active Window is Modal Window '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0
		End Function
		
		'***********************************************************************************************************************************************************************************************
		
		Function GetSAPCheckBoxProperty (sSAPGuiCheck,sProperty)
				' @HELP
				' @class	:	Controls_SAP
				' @method	: 	SelectSAPCheckBox ( sSAPGuiCheck,sOption)
				' @parameter:   sSAPGUICheck:	Name of the Check box
				' @parameter:   sOption     :	Option (ON/OFF)
				' @parameter:   sIndex      :   Index of the Check ox on that screen
				' @notes	: 	Selects the Check box based on the option usig index property
				' @END 
				
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42").Exist then
					GetSAPCheckBoxProperty = 	SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42").GetROProperty (sProperty)
			
					End If 
				End If
		End Function


		'***********************************************************************************************************************************************************************************************

		Function SetSAPModalCheckBox ( sSAPGuiCheckBox,sDescription, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPCheckBox ( sSAPGuiCheckBox,sDescription,sValue)
			' @parameter:   sSAPGuiCheckBox	:	name property of the Checkbox
			' @parameter:   sDescription	:	Description of the Checkbox
			' @parameter:   sValue		:	ON/OFF
			' @notes:		Set the Checkbox existance on Main window using name property	 	
			' @END
			On error resume next
			
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
						
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiCheckBox("name:="&sSAPGuiCheckBox,"guicomponenttype:=42").Exist(0) Then
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiCheckBox("name:="&sSAPGuiCheckBox,"guicomponenttype:=42").Set sValue
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiCheckBox("text:="&sDescription,"guicomponenttype:=42").Exist(0) Then 
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiCheckBox("text:="&sDescription,"guicomponenttype:=42").Set sValue
				End If
				
				
				If err.description = "" Then
					UpdateExecutionReport micPass, "SetSAPModalCheckBox", "Value '"&sValue&"'set in SAP Checkbox '"&sDescription&"' in Modal Window '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "SetSAPModalCheckBox","SAP Checkbox: '"&sDescription&"' ### Error Description:"&err.description
				End If
			
				
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPModalCheckBox", "SAP Checkbox: '"&sDescription&"' ### Active Window is Main Window '"&sScreenName&"'"
			End If
			
			Err.Clear
			On error goto 0
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPModalCheckBoxUsingIndex ( iIndex, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPCheckBox ( sTechName)
			' @parameter:   sTechName	:	name property of the Checkbox
			' @parameter:   sValue		:	ON/OFF
			' @notes:		Set the Checkbox existance on Main window using name property	 	
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiCheckBox("index:="&iIndex,"guicomponenttype:=42").Set sValue
				If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPModalCheckBoxUsingIndex", "Value '"&sValue&"'set in SAP Checkbox with index '"&iIndex&"' in Modal Window '"&sScreenName&"'"
				else
				UpdateExecutionReport micFail, "SetSAPModalCheckBoxUsingIndex","SAP Checkbox with index '"&iIndex&"' ### Error Description:"&err.description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPModalCheckBoxUsingIndex", "SAP Checkbox with index '"&iIndex&"' ### Active Window is Main Window '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPCheckBoxUsingIndex ( iIndex, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPCheckBox ( sTechName)
			' @parameter:   iIndex		:	Index property of the Checkbox
			' @parameter:   sValue		:	ON/OFF
			' @notes	:	Set the Checkbox existance on Main window using Index property	 	
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("index:="&iIndex,"guicomponenttype:=42").Set sValue
				If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPCheckBoxUsingIndex", "Value '"&sValue&"'set in SAP Checkbox with index '"&iIndex&"' in Main Window '"&sScreenName&"'"
				else
				UpdateExecutionReport micFail, "SetSAPCheckBoxUsingIndex","SAP Checkbox with index '"&iIndex&"' ### Error Description:"&err.description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPCheckBoxUsingIndex", "SAP Checkbox with index '"&iIndex&"' ### Active Window is Modal Window '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPCheckBoxUsingNameAndIndex ( sSAPGuiCheckBox,sDescription,iIndex, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPCheckBox ( sTechName)
			' @parameter:   sTechName	:	name property of the Checkbox
			' @parameter:   sValue		:	ON/OFF
			' @notes:		Set the Checkbox existance on Main window using name property	 	
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheckBox,"index:="&iIndex,"guicomponenttype:=42").Set sValue
				If err.description = "" Then
					UpdateExecutionReport micPass, "SetSAPCheckBoxUsingNameAndIndex", "Value '"&sValue&"'set in SAP Checkbox with index '"&iIndex&"' and name '"&sDescription&"' in Main Window '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "SetSAPCheckBoxUsingNameAndIndex","SAP Checkbox with index '"&iIndex&"' and name '"&sDescription&"' ### Error Description:"&err.description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPCheckBoxUsingNameAndIndex", "SAP Checkbox with index '"&iIndex&"' and name '"&sDescription&"' ### Active Window is Modal Window '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0
		End Function
	Rem End Function
	'***********************************************************************************************************************************************************************************************

'***********************************************************************************************************************************************************************************************
Rem Function SAPComboBoxRelatedFunctions ()
Function SetSAPComboBox ( sSAPGuiComboBox,sValue)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	SetSAPModalComboBox ( sSAPGuiComboBox,sDescription,sValue)
	' @parameter:   sSAPGUIComboBox	:	Name of the combo box
	' @parameter:   sValue          :  	Value to be set 
	' @notes	: 	Sets the data into SAPGuiComboBox in Modal Window
	' @END 	
	On error resume next
	Err.Clear
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
	Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
	sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").Select sValue
		sDescription=oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").GetRoProperty("selecteditem")
		If err.description = "" Then
			UpdateExecutionReport micPass, "SetSAPModalComboBox","Value '"& sValue &"' set in Combobox '"&sDescription&"'  in Screen '"&sScreenName&"'"
		else
			UpdateExecutionReport micFail, "SetSAPModalComboBox","Screen:"&sScreenName&"###Error Description:"&err.description
		End If
	ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
	Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
	sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").Select sValue
		sDescription=oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").GetRoProperty("selecteditem")
		If err.description = "" Then
			UpdateExecutionReport micPass, "SetSAPModalComboBox","Value '"& sValue &"' set in Combobox '"&sDescription&"'  in Screen '"&sScreenName&"'"
		else
			UpdateExecutionReport micFail, "SetSAPModalComboBox","Screen:"&sScreenName&"###Error Description:"&err.description
		End If
	End if
	Err.Clear
	On error goto 0
End Function

'************************************************************************************************************************************************************
Function SelectSAPComboBox (sSAPGuiComboBox,sValue)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	SetSAPModalComboBox ( sSAPGuiComboBox,sDescription,sValue)
	' @parameter:   sSAPGUIComboBox	:	Name of the combo box
	' @parameter:   sValue          :  	Value to be set 
	' @notes	: 	Sets the data into SAPGuiComboBox in Modal Window
	''Modified on : 12/20/2016--Mahesh
	' @END 	
	On error resume next
	GetSAPEditValue=""
	Err.clear
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").Exist Then
		Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").Select sValue
		sAttachedText=oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").GetROProperty("selecteditem")
	Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").Exist Then
		Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
		sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").Select sValue
		sAttachedText=oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").GetROProperty("selecteditem")
	End if

	If 	err.description ="" Then
		UpdateExecutionReport micPass, "SelectSAPComboBox","Value '"& sValue &"' set in Combobox '"&sAttachedText&"'  in Screen '"&sScreenName&"'"
	Else
		UpdateExecutionReport micFail, "SelectSAPComboBox","Screen:"&sScreenName&"###Error Description:"&err.description
	End if
	Err.Clear
	On error goto 0
End Function
'************************************************************************************************************************************************************

Function SelectSAPComboBoxKey (sSAPGuiComboBox,sValue)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	SetSAPComboBox ( sSAPGuiComboBox,sValue)
	' @parameter:   sSAPGUIComboBox	:	Name of the combo box
	' @parameter:   sValue          :  	Value to be set 
	' @notes	: 	Sets the data into SAPGuiComboBox in Modal Window
	' @END 	
'	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
        	Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
			sScreenName = oSAPObj.GetROProperty("text")
			oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").SelectKey sValue
            sAttachedText=oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").GetROProperty("attachedtext")
			UpdateExecutionReport micPass, "SelectSAPComboBoxKey","KeyValue '"& sValue &"' set in Combobox '"& sAttachedText &"'  in Screen '"&sScreenName&"'"
		Else
        	Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
			sScreenName = oSAPObj.GetROProperty("text")
			oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").SelectKey sValue
            sAttachedText=oSAPObj.SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").GetROProperty("attachedtext")
			UpdateExecutionReport micPass, "SelectSAPComboBoxKey","KeyValue '"& sValue &"' set in Combobox '"& sAttachedText &"'  in Screen '"&sScreenName&"'"
	End if
	
		If err.description <> "" Then
			UpdateExecutionReport micFail, "SelectSAPComboBoxKey","Screen: "& sScreenName & "@ Screen Field : " & sAttachedText & "@ Field Value : " & sValue & " ### Error Description: " &err.description
		End If
	Err.Clear
	On error goto 0
End Function
'************************************************************************************************************************************************************
		Function VerifySAPComboBox( sSAPGuiComboBox)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPComboBox( sSAPGuiComboBox)
			' @parameter:   sSAPGUIComboBox	:	Name of the combo box
			' @notes	: 	Verify the existance of SAPGuiComboBox in Main Window
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").Exist Then 
				VerifySAPComboBox = True
				UpdateExecutionReport micPass, "VerifySAPComboBox","ComboBox '"& sSAPGuiComboBox &"' exists in Screen '"&sScreenName&"'"
				else
				VerifySAPComboBox = False
				UpdateExecutionReport micPass, "VerifySAPComboBox","ComboBox '"& sSAPGuiComboBox &"' doesnot exist in Screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "VerifySAPComboBox","ComboBox: '"& sSAPGuiComboBox &"'### Active Window is Modal Window '"&sScreenName&"'"
			End If
			If err.description <> "" Then
				UpdateExecutionReport micFail, "VerifySAPComboBox","ComboBox:"& sSAPGuiComboBox &"' ### Screen:"&sScreenName&" ### Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0
		End Function

		'***********************************************************************************************************************************************************************************************

		Function VerifySAPModalComboBox( sSAPGuiComboBox)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPComboBox( sSAPGuiComboBox)
			' @parameter:   sSAPGUIComboBox	:	Name of the combo box
			' @notes	: 	Verify the existance of SAPGuiComboBox in Main Window
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").Exist Then 
				VerifySAPModalComboBox = True
				UpdateExecutionReport micPass, "VerifySAPModalComboBox","ComboBox '"& sSAPGuiComboBox &"' exists in Screen '"&sScreenName&"'"
				else
				VerifySAPModalComboBox = False
				UpdateExecutionReport micPass, "VerifySAPModalComboBox","ComboBox '"& sSAPGuiComboBox &"' doesnot exist in Screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "VerifySAPModalComboBox","ComboBox: '"& sSAPGuiComboBox &"'### Active Window is Main Window '"&sScreenName&"'"
			End If
			If err.description <> "" Then
				UpdateExecutionReport micFail, "VerifySAPModalComboBox","ComboBox:"& sSAPGuiComboBox &"' ###Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPComboBoxValue ( sSAPGuiComboBox)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPComboBoxValue ( sSAPGuiComboBox)
			' @parameter:   sSAPGUIComboBox	:	Name of the combo box
			' @notes	: 	returns a value set in SAPGuiComboBox in Main Window
			' @END 	
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				GetSAPComboBoxValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").GetROProperty ("selecteditem")
				If err.description = "" Then
				UpdateExecutionReport micPass, "GetSAPComboBoxValue","Selected item '"&GetSAPComboBoxValue&"' found in Combobox '"&sSAPGuiComboBox&"' in Screen '"&sScreenName&"'"
				else
				UpdateExecutionReport micFail, "GetSAPComboBoxValue","Screen:"&sScreenName&"###Error Description:"&err.description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPComboBoxValue", "SAP Checkbox: '"& sSAPGuiComboBox &"' ### Active Window is Modal Window '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalComboBoxValue ( sSAPGuiComboBox)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalComboBoxValue ( sSAPGuiComboBox)
			' @parameter:   sSAPGUIComboBox	:	Name of the combo box
			' @notes	: 	returns a value set in SAPGuiComboBox in Modal Window
			' @END 	
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				GetSAPModalComboBoxValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiComboBox("name:="&sSAPGuiComboBox,"guicomponenttype:=34").GetROProperty ("selecteditem")
				If err.description = "" Then
					UpdateExecutionReport micPass, "GetSAPModalComboBoxValue","Selected item '"&GetSAPModalComboBoxValue&"' found in Combobox '"&sSAPGuiComboBox&"' in Screen '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "GetSAPModalComboBoxValue","Screen:"&sScreenName&"###Error Description:"&err.description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPModalComboBoxValue", "SAP Checkbox: '"& sSAPGuiComboBox &"' ### Active Window is Main Window '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0
		End Function
	Rem End Function 
'***********************************************************************************************************************************************************************************************

Rem Function SAPEditBoxRelatedFunctions()
Function SetSAPEdit ( sSAPGUIEdit,sValue)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	SetSAPEdit ( sSAPGUIEdit,sEditDesc,sValue)
	' @parameter:   sSAPGUIEdit	:	Name of the GuiEdit field   
	' @parameter:   sEditDesc	:	Component type of the GuiEdit Field
	' @parameter:   sValue		:	Value to be set
	' @notes	:	Sets the data into a SAPGuiEdit
	' @END
	On error resume next
	Err.Clear
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
		Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		If oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Exist then
			oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Set sValue
'			oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Set InputData("OrderType")
			sEditDesc=oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").GetROproperty("attachedtext")
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPEdit","Value '"& sValue &"' set in Edit '" & sEditDesc & "'  in Screen '"&sScreenName&"'"
			Else
				UpdateExecutionReport micFail, "SetSAPEdit","Edit '" & sEditDesc & "' ### Screen:"&sScreenName&"###Error Description:"&err.description
			End If	
		Elseif oSAPObj.SAPGuiEdit("attachedtext:="&sSAPGUIEdit,"guicomponenttype:=32").Exist then
			oSAPObj.SAPGuiEdit("attachedtext:="&sSAPGUIEdit,"guicomponenttype:=32").Set sValue
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPEdit","Value '"& sValue &"' set in Edit '" & sSAPGUIEdit & "'  in Screen '"&sScreenName&"'"
			Else
				UpdateExecutionReport micFail, "SetSAPEdit","Edit '" & sSAPGUIEdit & "' ### Screen:"&sScreenName&"###Error Description:"&err.description
			End If
		ElseIf oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Exist then
			oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Set sValue
			sEditDesc=oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").GetROproperty("attachedtext")
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPEdit","Value '"& sValue &"' set in Edit '" & sEditDesc & "'  in Screen '"&sScreenName&"'"
			Else
				UpdateExecutionReport micFail, "SetSAPEdit","Edit '" & sEditDesc & "' ### Screen:"&sScreenName&"###Error Description:"&err.description
			End If
		Elseif oSAPObj.SAPGuiEdit("attachedtext:="&sSAPGUIEdit,"guicomponenttype:=31").Exist then
			oSAPObj.SAPGuiEdit("attachedtext:="&sSAPGUIEdit,"guicomponenttype:=31").Set sValue
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPEdit","Value '"& sValue &"' set in Edit '" & sSAPGUIEdit & "'  in Screen '"&sScreenName&"'"
			Else
				UpdateExecutionReport micFail, "SetSAPEdit","Edit '" & sSAPGUIEdit & "' ### Screen:"&sScreenName&"###Error Description:"&err.description
			End If
		End if
	ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
		Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
				If oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Exist then
			oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Set sValue
			sEditDesc=oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").GetROproperty("attachedtext")
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPEdit","Value '"& sValue &"' set in Edit '" & sEditDesc & "'  in Screen '"&sScreenName&"'"
			Else
				UpdateExecutionReport micFail, "SetSAPEdit","Edit '" & sEditDesc & "' ### Screen:"&sScreenName&"###Error Description:"&err.description
			End If	
		Elseif oSAPObj.SAPGuiEdit("attachedtext:="&sSAPGUIEdit,"guicomponenttype:=32").Exist then
			oSAPObj.SAPGuiEdit("attachedtext:="&sSAPGUIEdit,"guicomponenttype:=32").Set sValue
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPEdit","Value '"& sValue &"' set in Edit '" & sSAPGUIEdit & "'  in Screen '"&sScreenName&"'"
			Else
				UpdateExecutionReport micFail, "SetSAPEdit","Edit '" & sSAPGUIEdit & "' ### Screen:"&sScreenName&"###Error Description:"&err.description
			End If
		ElseIf oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Exist then
			oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Set sValue
			sEditDesc=oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").GetROproperty("attachedtext")
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPEdit","Value '"& sValue &"' set in Edit '" & sEditDesc & "'  in Screen '"&sScreenName&"'"
			Else
				UpdateExecutionReport micFail, "SetSAPEdit","Edit '" & sEditDesc & "' ### Screen:"&sScreenName&"###Error Description:"&err.description
			End If
		Elseif oSAPObj.SAPGuiEdit("attachedtext:="&sSAPGUIEdit,"guicomponenttype:=31").Exist then
			oSAPObj.SAPGuiEdit("attachedtext:="&sSAPGUIEdit,"guicomponenttype:=31").Set sValue
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPEdit","Value '"& sValue &"' set in Edit '" & sSAPGUIEdit & "'  in Screen '"&sScreenName&"'"
			Else
				UpdateExecutionReport micFail, "SetSAPEdit","Edit '" & sSAPGUIEdit & "' ### Screen:"&sScreenName&"###Error Description:"&err.description
			End If
		End if
	End If
	Err.Clear
	On error goto 0
End Function
''------------------------------------------------------------------------------------------------------------
		Function SetSAPEdit_UsingAttachedText(sAttachedText,sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPEdit_UsingAttachedText(sAttachedText,sValue)
			' @parameter:   sAttachedText	:	AttachedText of the GuiEdit field  
			' @parameter:   sValue		:	Value to be set
			' @notes	:	Sets the data into a SAPGuiEdit
			' @END
			on error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			 sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("guicomponenttype:=32","attachedtext:="&sAttachedText).Exist Then
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("guicomponenttype:=32","attachedtext:="&sAttachedText).Set sValue
					If err.Description = "" Then
						UpdateExecutionReport micPass,"SetSAPEdit_UsingAttachedText","Value '"& sValue &"' set into SAPEdit '"& sAttachedText &"' of MainScreen '"&sScreenName&"'"
					else
						UpdateExecutionReport micFail,"SetSAPEdit_UsingAttachedText","Value '"& sValue &"' not set into SAPEdit '"& sAttachedText &"' of MainScreen '"&sScreenName&"'"
					End If
				elseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("guicomponenttype:=31","attachedtext:="&sAttachedText).Exist Then
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("guicomponenttype:=31","attachedtext:="&sAttachedText).Set sValue
					If err.Description = "" Then
						UpdateExecutionReport micPass,"SetSAPEdit_UsingAttachedText","Value '"& sValue &"' set into SAPEdit '"& sAttachedText &"' of MainScreen '"&sScreenName&"'"
					else
						UpdateExecutionReport micFail,"SetSAPEdit_UsingAttachedText","Value '"& sValue &"' not set into SAPEdit '"& sAttachedText &"' of MainScreen '"&sScreenName&"'"
					End If
				End If
			Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail,"SetSAPEdit_UsingAttachedText","SAPEdit = '"&sAttachedText&"' ### Active Window is Modal Window '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0
		End Function
		'***********************************************************************************************************************************************************************************************
			
Function SetSAPModalEdit_UsingAttachedText(sAttachedText,sValue)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	SetSAPModalEdit_UsingAttachedText(sAttachedText,sValue)
				' @parameter:   sAttachedText	:	AttachedText of the GuiEdit field  
				' @parameter:   sValue		:	Value to be set
				' @notes	:	Sets the data into a SAPGuiEdit
				' @END
				on error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				 sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("guicomponenttype:=31","attachedtext:="&sAttachedText).Exist Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("guicomponenttype:=31","attachedtext:="&sAttachedText).Set sValue
						If err.Description = "" Then
							UpdateExecutionReport micPass,"SetSAPModalEdit_UsingAttachedText","Value '"& sValue &"' set into SAPEdit '"& sAttachedText &"' of MainScreen '"&sScreenName&"'"
						else
							UpdateExecutionReport micFail,"SetSAPModalEdit_UsingAttachedText","Value '"& sValue &"' not set into SAPEdit '"& sAttachedText &"' of MainScreen '"&sScreenName&"'"
						End If
					elseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("guicomponenttype:=32","attachedtext:="&sAttachedText).Exist Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("guicomponenttype:=32","attachedtext:="&sAttachedText).Set sValue
						If err.Description = "" Then
							UpdateExecutionReport micPass,"SetSAPModalEdit_UsingAttachedText","Value '"& sValue &"' set into SAPEdit '"& sAttachedText &"' of MainScreen '"&sScreenName&"'"
						else
							UpdateExecutionReport micFail,"SetSAPModalEdit_UsingAttachedText","Value '"& sValue &"' not set into SAPEdit '"& sAttachedText &"' of MainScreen '"&sScreenName&"'"
						End If
					End If
				Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail,"SetSAPModalEdit_UsingAttachedText","SAPEdit = '"&sAttachedText&"' ### Active Window is Modal Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0
			End Function

			'***********************************************************************************************************************************************************************************************

			Function SetSAPModalEdit ( sSAPGUIEdit,sEditDesc,sValue)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	SetSAPModalEdit ( sSAPGUIEdit,sEditDesc,sValue)
				' @parameter:   sSAPGUIEdit	:	Name of the GuiEdit field   
				' @parameter:   sEditDesc	:	Component type of the GuiEdit Field
				' @parameter:   sValue		:	Value to be set
				' @notes	:	Sets the data into a SAPGuiEdit on a Modal Window
				' @END
				On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Exist() then
					
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Set sValue
					
					End if
					If err.description = "" Then
						UpdateExecutionReport micPass, "SetSAPModalEdit","Value '"& sValue &"' set in Modal Edit '"&sEditDesc&"' in Screen '"&sScreenName&"'"
					End if
						'UpdateExecutionReport micFail, "SetSAPModalEdit","Screen:"&sScreenName&"###Error Description:"&err.description
					'End If
					If err.description <> "" Then
						If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Exist() then
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Set sValue
							UpdateExecutionReport micPass, "SetSAPModalEdit","Value '"& sValue &"' set in Edit '" & sEditDesc & "'  in Screen '"&sScreenName&"'"
						Else
							UpdateExecutionReport micFail, "SetSAPModalEdit","Screen:"&sScreenName&"###Error Description:"&err.description
						End If
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "SetSAPModalEdit", "SAP Edit: '"& sSAPGUIEdit &"' ### Active Window is Main Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0
			End Function

			'***********************************************************************************************************************************************************************************************
			Sub SetSAPEdit_UsingIndex ( sSAPGUIEdit,sDescription,iIndex,sValue)
				' @HELP
				' @class	:   Controls_SAP
				' @method	:   SetSAPEditUsingIndex ( sSAPGUIEdit,sDescription,iIndex,sValue)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPGUIEdit	:	Name of the field   
				' @parameter:   sDescription   :	Description of the SAP Edit Field 
				' @parameter:   iIndex         :	Index of the field
				' @parameter:   sValue      :	Value to be set in to field
				' @notes	:  	Sets data in to Gui Edit field , using index  property to identify the field       
				' @END

			   ' On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
							If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31", "index:="&iIndex).Exist(0) Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31", "index:="&iIndex).Set sValue
					End If
					
					If err.description = "" Then
						UpdateExecutionReport micPass, "SetSAPEditUsingIndex","Value '"& sValue &"' set in Edit '" &sDescription & "'  in Screen '"&sScreenName&"'"
					End If	

					If err.description <>  "" Then
						If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32","index:="&iIndex).Exist(0) then
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32","index:="&iIndex).Set sValue
							UpdateExecutionReport micPass, "SetSAPEditUsingIndex","Value '"& sValue &"' set in Edit '" & sEditDesc & "'  in Screen '"&sScreenName&"'"
						Else
						UpdateExecutionReport micFail, "SetSAPEditUsingIndex","SAPGUIEdit '"&sDescription&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
					End If	
						
				End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SetSAPEditUsingIndex","Value '"&sValue&"' couldnot be set in SAPGUIEdit '"&sDescription&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SetSAPEditUsingIndex","SAPGUIEdit '"&sDescription&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0		
			End Sub

			'**************************************************************************************************************************************************************************
			
			Sub SetSAPModalEdit_UsingIndex ( sSAPGUIEdit,sDescription,iIndex,sValue)
				' @HELP
				' @class	:   Controls_SAP
				' @method	:   SetSAPModalEdit_UsingIndex ( sSAPGUIEdit,sDescription,iIndex,sValue)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPGUIEdit	:	Name of the field   
				' @parameter:   sDescription   :	Description of the SAP Edit Field 
				' @parameter:   iIndex         :	Index of the field
				' @parameter:   sValue      :	Value to be set in to field
				' @notes	:  	Sets data in to Gui Edit field , using index  property to identify the field       
				' @END

			   ' On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
							If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31", "index:="&iIndex).Exist(0) Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31", "index:="&iIndex).Set sValue
					End If
					
					If err.description = "" Then
						UpdateExecutionReport micPass, "SetSAPModalEdit_UsingIndex","Value '"& sValue &"' set in Edit '" &sDescription & "'  in Screen '"&sScreenName&"'"
					End If	

					If err.description <>  "" Then
						If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32","index:="&iIndex).Exist(0) then
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32","index:="&iIndex).Set sValue
							UpdateExecutionReport micPass, "SetSAPModalEdit_UsingIndex","Value '"& sValue &"' set in Edit '" & sEditDesc & "'  in Screen '"&sScreenName&"'"
						Else
						UpdateExecutionReport micFail, "SetSAPModalEdit_UsingIndex","SAPGUIEdit '"&sDescription&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
					End If	
						
				End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "SetSAPModalEdit_UsingIndex","Value '"&sValue&"' couldnot be set in SAPGUIEdit '"&sDescription&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SetSAPModalEdit_UsingIndex","SAPGUIEdit '"&sDescription&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0		
			End Sub



	'********************************************************************************************************************************************************************************************************

			Function VerifySAPEdit ( sSAPGUIEdit,sDescription)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	VerifySAPEdit ( sSAPGUIEdit,sDescription)
				' @parameter:   sSAPGUIEdit	:	Name of the GuiEdit field   
				' @parameter:   sDescription	:	Description of the GuiEdit Field
				' @notes	:	Verify the existance of SAP Edit in Main Window
				' @END
				On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Exist Then
						VerifySAPEdit = True
					ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Exist Then
						VerifySAPEdit = True
					Else
						VerifySAPEdit = False
					End If
					If err.description = "" Then
						UpdateExecutionReport micPass, "VerifySAPEdit","SAP Edit '"& sDescription &"'  field exist in modal window '"&sScreenName&"'"
					End If	
					If err.description <>  "" Then
						If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31","index:="&iIndex).Exist(0) then
							UpdateExecutionReport micPass, "VerifySAPEdit","SAP Edit '"& sDescription &"'  field exist in modal window '"&sScreenName&"'"
							Err.Clear
						Else
							UpdateExecutionReport micPass, "VerifySAPEdit","SAP Edit '"& sDescription &"'  field does not exist in modal window '"&sScreenName&"'"
							
					End If	
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "VerifySAPEdit", "SAP Edit: '"&sDescription&"' ### Active Window is Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "VerifySAPEdit","SAP Edit: '"&sDescription&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Function

			'***********************************************************************************************************************************************************************************************

			Function VerifySAPModalEdit ( sSAPGUIEdit,sDescription)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	VerifySAPModalEdit ( sSAPGUIEdit,sCompType)
				' @parameter:   VerifySAPModalEdit	:	Name of the GuiEdit field   
				' @parameter:   sDescription				:	Description of the GuiEdit Field
				' @notes	:	Verify the existance of SAP Edit in Main Window
				' @END
				On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Exist(0) Then
					VerifySAPModalEdit = True
					UpdateExecutionReport micPass, "VerifySAPModalEdit","SAP Modal Edit '"& sDescription &"'  field exist in modal window '"&sScreenName&"'"
					ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Exist(0) then
						VerifySAPModalEdit = True
						UpdateExecutionReport micPass, "VerifySAPModalEdit","SAP Modal Edit '"& sDescription &"'  field exist in modal window '"&sScreenName&"'"
					Else
						UpdateExecutionReport micPass, "VerifySAPModalEdit","SAP Modal Edit '"& sDescription &"'  field does not exist in modal window '"&sScreenName&"'"						
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "VerifySAPModalEdit", "SAP Edit: '"&sDescription&"' ### Active Window is Main Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "VerifySAPModalEdit","SAP Edit: '"&sDescription&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Function

			'***********************************************************************************************************************************************************************************************

Function GetSAPEditValue ( sSAPGUIEdit)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	GetSAPEditValue ( sSAPGUIEdit)
	' @Returns	:	Returns Value set in SAP Edit
	' @parameter:   sSAPGUIEdit	:	Name of the combo box
	' @notes	: 	Sets the data into SAPGuiComboBox in Main Window
	' @END 	
	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
		Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		sScreenName = oSAPObj.GetROProperty("text")
		If oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Exist then
			GetSAPEditValue = oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").GetROProperty ("value")
		Elseif oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Exist Then
			GetSAPEditValue = oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").GetROProperty ("value")
		End If
		If err.description = "" Then
			UpdateExecutionReport micPass, "GetSAPEditValue","Selected item '"&GetSAPEditValue&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
		End If
	Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
		Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		sScreenName = oSAPObj.GetROProperty("text")
		If oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Exist then
			GetSAPEditValue = oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").GetROProperty ("value")
		Elseif oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Exist Then
			GetSAPEditValue = oSAPObj.SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").GetROProperty ("value")
		End If
		If err.description = "" Then
			UpdateExecutionReport micPass, "GetSAPEditValue","Selected item '"&GetSAPEditValue&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
		End If
	End if
	Err.Clear
	On error goto 0
End Function

			'***********************************************************************************************************************************************************************************************

			Function GetSAPOffsetEditValue ( sSAPGUIEdit,sDescription,iIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	GetSAPOffsetEditValue (sSAPGUIEdit,sDescription,iIndex)
				' @parameter:   sSAPGUIEdit	:	Name of the SAP OffsetEdit
				' @parameter:   sDescription:	Description of the SAP OffsetEdit
				' @parameter:   iIndex      :  	Index of the OffsetEdit
				' @notes	: 	Sets the data into SAPGuiComboBox in Main Window
				' @END 	
				On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					GetSAPOffsetEditValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31","index:="&iIndex).GetROProperty ("value")
					If err.description = "" Then
						UpdateExecutionReport micPass, "GetSAPOffsetEditValue","Selected item '"&GetSAPOffsetEditValue&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					else
						UpdateExecutionReport micFail, "GetSAPOffsetEditValue","Screen:"&sScreenName&"###Error Description:"&err.description
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "GetSAPOffsetEditValue", "SAP Edit: '"& sSAPGUIEdit &"' ### Active Window is Modal Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0
			End Function

			'***********************************************************************************************************************************************************************************************
			
			Function GetSAPModalOffsetEditValue ( sSAPGUIEdit,sDescription,iIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	GetSAPOffsetEditValue (sSAPGUIEdit,sDescription,iIndex)
				' @parameter:   sSAPGUIEdit	:	Name of the SAP OffsetEdit
				' @parameter:   sDescription:	Description of the SAP OffsetEdit
				' @parameter:   iIndex      :  	Index of the OffsetEdit
				' @notes	: 	Sets the data into SAPGuiComboBox in Main Window
				' @END 	
				On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					GetSAPModalOffsetEditValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31","index:="&iIndex).GetROProperty ("value")
					If err.description = "" Then
						UpdateExecutionReport micPass, "GetSAPModalOffsetEditValue","Selected item '"&GetSAPModalOffsetEditValue&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					else
						UpdateExecutionReport micFail, "GetSAPModalOffsetEditValue","Screen:"&sScreenName&"###Error Description:"&err.description
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "GetSAPModalOffsetEditValue", "SAP Edit: '"& sSAPGUIEdit &"' ### Active Window is Main Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0
			End Function

			'***********************************************************************************************************************************************************************************************

			Function GetSAPModalEditValue ( sSAPGUIEdit,sDescription)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	GetSAPModalEditValue ( sSAPGUIEdit,sDescription)
				' @parameter:   sSAPGUIEdit		:	Name of the combo box
				' @parameter:   sDescription    :  	Description of the SAP Edit
				' @notes	: 	Gets the data from  SAPGuiEdit field in Modal Window
				' @END 	
				On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					If GetSAPModalEditValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").Exist() Then
						GetSAPModalEditValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32").GetROProperty ("value")
					End If
					If err.description = "" Then
						UpdateExecutionReport micPass, "GetSAPModalEditValue","Selected item '"&GetSAPModalEditValue&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					End If
					If err.description <> "" Then
						If GetSAPModalEditValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").Exist() Then
							GetSAPModalEditValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31").GetROProperty ("value")
							UpdateExecutionReport micPass, "GetSAPModalEditValue","Selected item '"&GetSAPModalEditValue&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
						Else	
							UpdateExecutionReport micFail, "GetSAPModalEditValue","Screen: "&sScreenName&"###Error Description:"&err.description
						End If	
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "GetSAPModalEditValue", "SAP Edit: '"& sSAPGUIEdit &"' ### Active Window is not a Modal Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0
			End Function
			
				'***********************************************************************************************************

			Function GetFieldValue_UsingDesc (sFieldDesc, aInputData)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	GetFieldValue_UsingDesc (sFiledName, aInputData)
				' @returns	:	GetFieldValue_UsingDesc	:	String  
				' @parameter:   sFieldDesc				:	Description of the field to which value is to retreived 
				' @parameter:   alnputData				:	DataArray that contains InputData
				' @notes	:	This Method returns a field value for a given Field description from a data array	 
				' @END
				For iRow = 0 to UBound(aInputData)
					If  (aInputData(iRow,1)= sFieldDesc)Then
						GetFieldValue_UsingDesc = aInputData(iRow,3)
						exit for
					ElseIf (aInputData(iRow,4)= sFieldDesc)Then
						GetFieldValue_UsingDesc = aInputData(iRow,3)
						exit For
					End If 
				Next
			End Function

			'***************************************************************************************************************************************
			Function GetSAPModalEditValue_UsingIndex ( sSAPGUIEdit,sCompType,iIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	GetSAPModalEditValue ( sSAPGUIEdit,sCompType)
				' @parameter:   sSAPGUIEdit		:	Name of the combo box
				' @parameter:   sCompType       :  	GUIComponentType of SAP Edit
				' @notes	: 	Sets the data into SAPGuiComboBox in Main Window
				' @END 	
				On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22","index:="&iIndex).Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22","index:="&iIndex).GetROProperty("text")
					GetSAPModalEditValue_UsingIndex = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType,"index:="&iIndex).GetROProperty ("value")
					
					If err.description = "" Then
					
						UpdateExecutionReport micPass, "GetSAPModalEditValue_UsingIndex","Selected item '"&GetSAPModalEditValue_UsingIndex&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					else
						UpdateExecutionReport micFail, "GetSAPModalEditValue_UsingIndex","Screen:"&sScreenName&"###Error Description:"&err.description
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21","index:="&iIndex).GetROProperty("text")
					UpdateExecutionReport micFail, "GetSAPModalEditValue_UsingIndex", "SAP Edit: '"& sSAPGUIEdit &"' ### Active Window is Main Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0
			End Function


			'***************************************************************************************************************************************

			Function GetSAPEditText_UsingIndex ( sSAPGUIEdit,iIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	GetSAPEditText_UsingIndex ( sSAPGUIEdit,iIndex)
				' @parameter:   sSAPGUIEdit		:	Name of the combo box
				' @parameter:   iIndex       :  	Index NUmber of the SAP Edit
				' @notes	: 	Get the label text data from SAPGuiEdit field
				' @END 	
				On error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21","index:="&iIndex).Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					GetSAPEditText_UsingIndex = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=31","index:="&iIndex).GetROProperty ("value")
					
					If err.description = "" Then
					
						UpdateExecutionReport micPass, "GetSAPEditText_UsingIndex","Selected item '"&GetSAPEditText_UsingIndex&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					else
						UpdateExecutionReport micFail, "GetSAPEditText_UsingIndex","Screen:"&sScreenName&"###Error Description:"&err.description
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22","index:="&iIndex).GetROProperty("text")
					UpdateExecutionReport micFail, "GetSAPEditText_UsingIndex", "SAP Edit: '"& sSAPGUIEdit &"' ### Active Window is Main Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0
			End Function
			
			'**************************************************************************************************************************************************************************

			Sub DblClickSAPEdit ( sSAPGUIEdit,sCompType)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	DblClickSAPEdit ( sSAPGUIEdit,sCompType)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPGUIEdit	:	Name of the field   
				' @parameter:   sCompType  	:	Component Type of the field
				' @notes	:	Double clicks on the specified SAP Edit
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType).Exist Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType).SetFocus
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SendKey F2
						UpdateExecutionReport micPass, "DblClickSAPEdit","Double Click Action performed on SAPGUIEdit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					Else
					UpdateExecutionReport micFail, "DblClickSAPEdit","Double Click Action not performed on SAPGUIEdit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "DblClickSAPEdit","Double Click Action not performed on SAPGUIEdit '"&sSAPGUIEdit&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "DblClickSAPEdit","SAPGUIEdit '"&sSAPGUIEdit&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0		
			End Sub 
           '*************************************************************************************************************************************************************************************
		   
		   Sub SelectSAPEdit ( sSAPGUIEdit,sCompType)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	DblClickSAPEdit ( sSAPGUIEdit,sCompType)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPGUIEdit	:	Name of the field   
				' @parameter:   sCompType  	:	Component Type of the field
				' @notes	:	Double clicks on the specified SAP Edit
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType).Exist Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType).SetFocus
						
						UpdateExecutionReport micPass, "SelectSAPEdit","Double Click Action performed on SAPGUIEdit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					Else
					UpdateExecutionReport micFail, "SelectSAPEdit","Double Click Action not performed on SAPGUIEdit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPEdit","Double Click Action not performed on SAPGUIEdit '"&sSAPGUIEdit&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPEdit","SAPGUIEdit '"&sSAPGUIEdit&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0		
			End Sub 

			'**************************************************************************************************************************************************************************
             
			  Sub SelectSAPModalEdit ( sSAPGUIEdit,sCompType)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	DblClickSAPEdit ( sSAPGUIEdit,sCompType)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPGUIEdit	:	Name of the field   
				' @parameter:   sCompType  	:	Component Type of the field
				' @notes	:	Double clicks on the specified SAP Edit
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType).Exist Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType).SetFocus
						
						UpdateExecutionReport micPass, "SelectSAPModalEdit","Double Click Action performed on SAPGUIEdit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					Else
					UpdateExecutionReport micFail, "SelectSAPModalEdit","Double Click Action not performed on SAPGUIEdit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPModalEdit","Double Click Action not performed on SAPGUIEdit '"&sSAPGUIEdit&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPModalEdit","SAPGUIEdit '"&sSAPGUIEdit&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0		
			End Sub 
			
			
'*******************************************************************************************************************************************************************


			Sub SetSAPEditUsingId ( sSAPGUIEdit,sId,sValue)
				' @HELP
				' @class	:   Controls_SAP
				' @method	:   SetSAPEditUsingId ( sSAPGUIEdit,sCompType,sId,sValue)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPGUIEdit	:	Name of the field   
				' @parameter:   sCompType   :	Component type of the field
				' @parameter:   sId         :	Id of the field
				' @parameter:   sValue      :	Value to be set in to field
				' @notes	:  	Sets data in to Gui Edit field , using id property to identify the field       
				' @END
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:=32", "id:="&sId).Set sValue
					UpdateExecutionReport micPass, "SetSAPEditUsingId","Value '"&sValue&"' set in SAPGUIEdit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SetSAPEditUsingId","Value '"&sValue&"' couldnot be set in SAPGUIEdit '"&sSAPGUIEdit&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SetSAPEditUsingId","SAPGUIEdit '"&sSAPGUIEdit&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0		
			End Sub

			'**************************************************************************************************************************************************************************

			Sub SetSAPModalEditUsingId ( sSAPGUIEdit,sCompType,sId,sValue)
				' @HELP
				' @class	:   Controls_SAP
				' @method	:   SetSAPModalEditUsingId ( sSAPGUIEdit,sCompType,sId,sValue)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPGUIEdit	:	Name of the field   
				' @parameter:   sCompType   :	Component type of the field
				' @parameter:   sId         :	Id of the field
				' @parameter:   sValue      :	Value to be set in to field
				' @notes	:  	Sets data in to Gui Edit field , using id property to identify the field       
				' @END
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType, "id:="&sId).Set sValue
					UpdateExecutionReport micPass, "SetSAPEditUsingId","Value '"&sValue&"' set in SAPGUIEdit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "SetSAPEditUsingId","Value '"&sValue&"' couldnot be set in SAPGUIEdit '"&sSAPGUIEdit&"' as Active Screen is a Main Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SetSAPEditUsingId","SAPGUIEdit '"&sSAPGUIEdit&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0		       
			End Sub

			'**************************************************************************************************************************************************************************

			Sub SetFocusOnSAPOffsetEdit ( sFieldName,sCompType,iIndex)
				' @HELP
				' @class	:  	Controls_SAP
				' @method	:  	SetFocusOnSAPOffsetEdit ( sFieldName,sCompType,iIndex)
				' @parameter:   sFieldName	:	Name of the field name 
				' @parameter:   sCompType   :	Component type of the field
				' @parameter:   iIndex     	:	Index number of the field
				' @notes	:  	Sets the focus on to the stepped loop field
				' @END
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sFieldName,"guicomponenttype:="&sCompType,"index:="&iIndex).Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sFieldName,"guicomponenttype:="&sCompType,"index:="&iIndex).SetFocus
						UpdateExecutionReport micPass, "SetFocusOnSAPOffsetEdit","SetFocus action performed on SAPGUIOffsetEdit '"&sFieldName&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SetFocusOnSAPOffsetEdit","SetFocus action not performed on SAPGUIOffsetEdit '"&sFieldName&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SetFocusOnSAPOffsetEdit","SetFocus action couldnot be performed on SAPGUIOffsetEdit '"&sFieldName&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SetFocusOnSAPOffsetEdit","SAPGUIOffsetEdit '"&sFieldName&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0		
			End Sub 
			
			'**************************************************************************************************************************************************************************

			Function OpenSAPOffsetEditPossibleEntries (sFieldName,sCompType,iIndex)
				On Error Resume Next
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sFieldName,"guicomponenttype:="&sCompType,"index:="&iIndex).OpenPossibleEntries
				If err.description = "" Then
					UpdateExecutionReport micPass, "OpenSAPOffsetEditPossibleEntries","Possible Entries displayed for OffSetEdit Name = "&sFieldName&" ### Guicomponenttype = "&sCompType&" ### Index = "&iIndex
				Else
					UpdateExecutionReport micFail, "OpenSAPOffsetEditPossibleEntries","OffSetEdit Name = "&sFieldName&" ### Guicomponenttype = "&sCompType&" ### Index = "&iIndex &" ### Error Description : "&err.description
				End If
				Err.Clear
				On error goto 0	
               CaptureSAPScreenShot()
			End Function

			'**************************************************************************************************************************************************************************
			
			Function OpenSAPModalEditPossibleEntries (sFieldName)
				On Error Resume Next
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sFieldName,"guicomponenttype:=32").OpenPossibleEntries
				If err.description = "" Then
					UpdateExecutionReport micPass, "OpenSAPOffsetEditPossibleEntries","Possible Entries displayed for Modal Edit Name = "&sFieldName&" ### Guicomponenttype = "&sCompType
				Else
					UpdateExecutionReport micFail, "OpenSAPOffsetEditPossibleEntries","OffSetEdit Name = "&sFieldName&" ### Guicomponenttype = "&sCompType&" ### Error Description : "&err.description
				End If
				Err.Clear
				On error goto 0	
               CaptureSAPScreenShot()
			End Function

			'**************************************************************************************************************************************************************************

			Sub SetSAPOffsetEdit ( sSAPGUIEdit,sCompType,iIndex,sValue)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	SetSAPOffsetEdit ( sSAPGUIEdit,sCompType,iIndex,sValue)
				' @parameter:   sSAPGUIEdit	:	Name of the Field
				' @parameter:   sCompType   :	Component type of the Field
				' @parameter:  	iIndex      :	Index number of the Field
				' @parameter:   sValue      :   Value to e set in to field
				' @notes	:	Sets data in to Stepped loop field 
				' @END 
				On error resume next
				If  SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType,"index:="&iIndex).Set sValue
				UpdateExecutionReport micPass, "SetSAPOffsetEdit","Screen:"&sScreenName&"### Data '"&sValue&"'set into OffSet field '"&sSAPGUIEdit&"' successfully"
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then  
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPOffsetEdit","OffSet field '"&sSAPGUIEdit&"' ### Active Window is Modal Window '"&sScreenName&"'"
				End If 
				If err.description <> "" Then
				UpdateExecutionReport micFail, "SetSAPOffsetEdit","OffSet field '"&sSAPGUIEdit&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0	
			End Sub

			'**************************************************************************************************************************************************************************

			Sub SetSAPModalOffsetEdit ( sSAPGUIEdit,sCompType,iIndex,sValue)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	SetSAPOffsetEdit ( sSAPGUIEdit,sCompType,iIndex,sValue)
				' @parameter:   sSAPGUIEdit	:	Name of the Field
				' @parameter:   sCompType   :	Component type of the Field
				' @parameter:  	iIndex      :	Index number of the Field
				' @parameter:   sValue      :   Value to e set in to field
				' @notes	:	Sets data in to Stepped loop field 
				' @END 
				On error resume next
				If  SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("name:="&sSAPGUIEdit,"guicomponenttype:="&sCompType,"index:="&iIndex).Set sValue
				UpdateExecutionReport micPass, "SetSAPModalOffsetEdit","Screen:"&sScreenName&"### Data '"&sValue&"'set into OffSet field '"&sSAPGUIEdit&"' successfully"
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then  
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPModalOffsetEdit","OffSet field '"&sSAPGUIEdit&"' ### Active Window is Main Window '"&sScreenName&"'"
				End If 
				If err.description <> "" Then
				UpdateExecutionReport micFail, "SetSAPModalOffsetEdit","OffSet field '"&sSAPGUIEdit&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0	
			End Sub
	Rem End Function
	'***********************************************************************************************************************************************************************************************

Rem Function SAPGridRelatedFunctions ()
	

Sub SetSAPGridData (sSAPGrid,aDataArray)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	SetSAPGridData (sSAPTable,aDataArray)
	' @parameter:   sSAPTable   :	Name of the GiuGrid
	' @parameter:   aDataArray  :	Array of fields and field values
	' @notes	: 	Sets the data into SAPGuiGrid
	' @Modified by/on	: 	Mahesh/ 01/2016
	' @END 
	On Error Resume Next
    Err.clear
	If InStr(sSAPGrid,"_") > 0 or InStr(sSAPGrid,"TABLE") > 0 or InStr(sSAPGrid,"SAP") then
		Set oDesc = Description.Create()
        oDesc("name").Value = sSAPGrid
		oDesc("guicomponenttype").Value = 201	
		oDesc("Index").Value = 1		
	Else
		Set oDesc = Description.Create()
		oDesc("name").Value = sSAPGrid			
		oDesc("guicomponenttype").Value = 201	
		oDesc("Index").Value = 1		
	End if
    If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
		Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		sScreenName = oSAPObj.GetROProperty("text")
		If oSAPObj.SAPGuiGrid(oDesc).Exist(1) Then
			For iRow = 0 to UBound(aDataArray)
				oSAPObj.SAPGuiGrid(oDesc).SetCellData aDataArray(iRow,2),aDataArray(iRow,3),aDataArray(iRow,4)
				If err.Description = "" then
					UpdateExecutionReport micPass, "SetSAPGridData", "Value '"&aDataArray(iRow,4)&"' set in Column '"&aDataArray(iRow,3)&"' of Table '"&sSAPGrid&"' of Screen '"&sScreenName&"'"
                Else
                    UpdateExecutionReport micFail, "SetSAPGridData", "Value '"&aDataArray(iRow,4)&"' couldnot be set in Column '"&aDataArray(iRow,3)&"' of Table '"&sSAPGrid&"' of Screen '"&sScreenName&"' ### Error Description:"&err.description
                End If
			Next
		End If
	Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
				sScreenName = oSAPObj.GetROProperty("text")
		If oSAPObj.SAPGuiGrid(oDesc).Exist(1) Then
			For iRow = 0 to UBound(aDataArray)
				oSAPObj.SAPGuiGrid(oDesc).SetCellData aDataArray(iRow,2),aDataArray(iRow,3),aDataArray(iRow,4)
				If err.Description = "" then
					UpdateExecutionReport micPass, "SetSAPGridData", "Value '"&aDataArray(iRow,4)&"' set in Column '"&aDataArray(iRow,3)&"' of Table '"&sSAPGrid&"' of Screen '"&sScreenName&"'"
                Else
                    UpdateExecutionReport micFail, "SetSAPGridData", "Value '"&aDataArray(iRow,4)&"' couldnot be set in Column '"&aDataArray(iRow,3)&"' of Table '"&sSAPGrid&"' of Screen '"&sScreenName&"' ### Error Description:"&err.description
                End If
			Next
		End If
	End if
	Err.Clear
    On error Goto 0
End Sub
'***********************************************************************************************************************************************************************************************
		Function SetSAPGridCellData ( sRow, sColumn, sData )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPGridCellData ( sRow,sColumn,sData)
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").SetCellData sRow,sColumn,sData
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SetSAPGridCellData", "Value '"&sData&"' set into Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SetSAPGridCellData", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPGridCellData", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPGridCellData_UsingIndex ( sRow, sColumn,iIndex,sData )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPGridCellData ( sRow,sColumn,sData)
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			Err.Clear
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).SetCellData sRow,sColumn,sData
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SetSAPGridCellData_UsingIndex", "Value '"&sData&"' set into Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SetSAPGridCellData_UsingIndex", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPGridCellData_UsingIndex", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPModalGridCellData ( sRow, sColumn, sData )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPModalGridCellData ( sRow,sColumn,sData)
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell").SetCellData sRow,sColumn,sData
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SetSAPModalGridCellData", "Value '"&sData&"' set into Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SetSAPModalGridCellData", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPModalGridCellData", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPGridCellData ( sRow, sColumn )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPGridCellData ( sRow,sColumn,sData)
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				GetSAPGridCellData = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").GetCellData (sRow,sColumn)
				If Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPGridCellData", "Value '"&GetSAPGridCellData&"' captured from Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "GetSAPGridCellData", "Value could not be captured of Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPGridCellData", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPGridCellData_UsingIndex ( iIndex,sRow, sColumn )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPGridCellData_UsingIndex ( sRow,sColumn,sData)
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				GetSAPGridCellData_UsingIndex = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).GetCellData (sRow,sColumn)
				If Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPGridCellData_UsingIndex", "Value '"&GetSAPGridCellData_UsingIndex&"' captured from Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "GetSAPGridCellData_UsingIndex", "Value could not be captured of Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPGridCellData_UsingIndex", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
		
		'*******************************************************************************************************************************************************************************
		
		Function GetSAPGridData (sProperty )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPGridData (sProperty)
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
		      GetSAPGridData = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201").GetROProperty(sProperty)
				If Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPGridData", "Value '"&GetSAPGridCellData_UsingIndex&"' captured from Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "GetSAPGridData", "Value could not be captured of Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPGridData", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		

		End Function


		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalGridCellData ( sRow, sColumn )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalGridCellData ( sRow,sColumn,sData)
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				GetSAPModalGridCellData = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell").GetCellData (sRow,sColumn)
				If Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPModalGridCellData", "Value '"&GetSAPModalGridCellData&"' captured from Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "GetSAPModalGridCellData", "Value could not be captured of Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPModalGridCellData", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************
		Function SetSAPModalGridCheckbox ( sRow, sColumn,sValue )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalGridCellData ( sRow,sColumn,sData)
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("name:=shell","type:=GuiShell","guicomponenttype:=201").SetCheckbox  sRow, sColumn,sValue
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SetSAPModalGridCheckbox", "Value '"&sValue&"' set into Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SetSAPModalGridCheckbox", "Value '"&sValue&"' couldnot be set into Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPModalGridCheckbox", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function 
		
		'***********************************************************************************************************************************************************************************************
		Function SetSAPGridCheckbox ( sRow, sColumn,sValue )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalGridCellData ( sRow,sColumn,sData)
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell","guicomponenttype:=201").SetCheckbox  sRow, sColumn,sValue
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SetSAPGridCheckbox", "Value '"&sValue&"' set into Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SetSAPGridCheckbox", "Value '"&sValue&"' couldnot be set into Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SetSAPGridCheckbox", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function 
		
		'***********************************************************************************************************************************************************************************************

		Function SelectSAPGridRow ( sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   selects specific row of a grid
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell", "Index:=1").selectRow sRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SelectSAPGridRow", "Grid Row '"&sRow&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPGridRow", "Grid Row '"&sRow&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPGridRow", "Grid Row '"&sRow&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function


		'***********************************************************************************************************************************************************************************************
		
			Function ExtandSAPGridRow ( sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	ExtandSAPGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Select  alternate row
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").ExtendRow  sRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "ExtandSAPGridRow", "Grid Row '"&sRow&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "ExtandSAPGridRow", "Grid Row '"&sRow&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "ExtandSAPGridRow", "Grid Row '"&sRow&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************
		Function SelectSAPGridColumn ( sColumn )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").SelectColumn sColumn
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SelectSAPGridColumn", "Grid Column '"&sColumn&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPGridColumn", "Grid Column '"&sColumn&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPGridColumn", "Grid Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
		
	Rem End Function
	
	  '*************************************************************************************************************************************************************************
	     Function SelectSAPTableColumn ( sName,sColumn )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPTableColumn ( sName,sColumn )
			' @parameter:	sColumn	:	Column  Name
			' @parameter:	sName	:	Name of the table
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("name:="&sName).SelectColumn sColumn
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SelectSAPTableColumn", "Table Column '"&sColumn&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPTableColumn", "Table Column '"&sColumn&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPTableColumn", "Table Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
		
		'***********************************************************************************************************************************************************************************************
		Sub ClickSAPGridCell_UsingIndex (iIndex, iRow, sColumn)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	ClickSAPGridCell ( iRow, sColumn)
				' @parameter:   iRow	:	Row number
				' @parameter:   sColumn :   Column Name
				' @notes	:	Clicks on the Specified cell in the SAP GUI Grid on SAP Main Window
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).ClickCell iRow, sColumn
						UpdateExecutionReport micPass, "ClickSAPGridCell_UsingIndex","Click action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "ClickSAPGridCell_UsingIndex","Click action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "ClickSAPGridCell_UsingIndex","Click action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "ClickSAPGridCell_UsingIndex","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub 
		'***********************************************************************************************************************************************************************************************
		Function SelectSAPGridRow_UsingIndex ( iIndex,sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).selectRow sRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SelectSAPGridRow_UsingIndex", "Grid Row '"&sRow&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPGridRow_UsingIndex", "Grid Row '"&sRow&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPGridRow_UsingIndex", "Grid Row '"&sRow&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************
		Function SelectSAPGridRowsRange_UsingIndex ( iIndex,iFRomRow,iToRow)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).SelectRowsRange iFRomRow,iToRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SelectSAPGridRow_UsingIndex", "Grid Rows from '"&iFRomRow&"' to '"&iToRow&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPGridRow_UsingIndex", "Grid Rows from '"&iFRomRow&"' to '"&iToRow&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPGridRow_UsingIndex", "Grid Rows from '"&iFRomRow&"' to '"&iToRow&"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
		'***********************************************************************************************************************************************************************************************
		Function SelectSAPGridRowsRange (iFRomRow,iToRow)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Inserts given data into specified grid cell
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").SelectRowsRange iFRomRow,iToRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SelectSAPGridRowsRange", "Grid Rows from '"&iFRomRow&"' to '"&iToRow&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPGridRowsRange", "Grid Rows from '"&iFRomRow&"' to '"&iToRow&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPGridRowsRange", "Grid Rows from '"&iFRomRow&"' to '"&iToRow&"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SelectSAPModalGridRow ( sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPModalGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Select SAPModalGridRow of a given row
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell").selectRow sRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SelectSAPModalGridRow", "Grid Row '"&sRow&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPModalGridRow", "Grid Row '"&sRow&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPModalGridRow", "Grid Row '"&sRow&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
         '**********************************************************************************************************************************************************************************
		 
		 Function SelectSAPModalGridRow_UsingIndex (Iindex, sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPModalGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Select SAPModalGridRow of a given row
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&Iindex).selectRow sRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SelectSAPModalGridRow_UsingIndex", "Grid Row '"&sRow&"' selected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPModalGridRow_UsingIndex", "Grid Row '"&sRow&"' couldnot be selected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPModalGridRow_UsingIndex", "Grid Row '"&sRow&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
		
		'***********************************************************************************************************************************************************************************************

		Function DeselectSAPGridRow ( sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	DeselectSAPGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Deselects a specified grid row
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").selectRow sRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "DeSelectSAPGridRow", "Grid Row '"&sRow&"' deselected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "DeSelectSAPGridRow", "Grid Row '"&sRow&"' couldnot be deselected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "DeSelectSAPGridRow", "Grid Row '"&sRow&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
        '**************************************************************************************************************************************************************************
		
		
		
Function DeselectSAPGridRow_UsingIndex  ( sRow ,iIndex)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	DeselectSAPGridRow_UsingIndex ( sRow ,iIndex)
			' @parameter:	sRow	:	Row number
			' @notes	:   Deselects a specified grid row
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index;="&iIndex).DeselectRow  sRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "DeselectSAPGridRow_UsingIndex", "Grid Row '"&sRow&"' deselected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "DeselectSAPGridRow_UsingIndex", "Grid Row '"&sRow&"' couldnot be deselected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "DeselectSAPGridRow_UsingIndex", "Grid Row '"&sRow&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
		
		'***********************************************************************************************************************************************************************************************

		Function DeselectSAPModalGridRow ( sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	DeselectSAPModalGridRow ( sRow )
			' @parameter:	sRow	:	Row number
			' @notes	:   Deselects a specified modal grid row
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell").selectRow sRow
				If Err.Description = "" then
					UpdateExecutionReport micPass, "DeSelectSAPModalGridRow", "Grid Row '"&sRow&"' deselected in MainWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "DeSelectSAPModalGridRow", "Grid Row '"&sRow&"' couldnot be deselected ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "DeSelectSAPModalGridRow", "Grid Row '"&sRow&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPGridRowCount ( )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPGridRowCount ( )
			' @notes	:   Returns rowcount of a grid
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPGridRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").Rowcount
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPGridRowCount", "Rows '"&GetSAPGridRowCount&"' exist for Grid of MainWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "GetSAPGridRowCount","MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPGridRowCount", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPGridRowCount_UsingIndex (iIndex )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPGridRowCount_UsingIndex ( )
			' @notes	:   Returns rowcount of a grid
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPGridRowCount_UsingIndex = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).Rowcount
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPGridRowCount_UsingIndex", "Rows '"&GetSAPGridRowCount_UsingIndex &"' exist for Grid of MainWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "GetSAPGridRowCount_UsingIndex","MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPGridRowCount_UsingIndex", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalGridRowCount ( )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalGridRowCount ( )
			' @notes	:   Returns rowcount of a modal grid
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				GetSAPModalGridRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell").Rowcount
				If Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPModalGridRowCount", "Rows '"&GetSAPGridRowCount&"' exist for Grid of ModalWindow '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "GetSAPModalGridRowCount","ModalWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPModalGridRowCount", "Grid Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPGridRowbyCellContent ( sColumn, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetTableRowbyContent ( sSAPTable, sColumn, sValue)
			' @returns	:	GetTableRowbyContent:	Reurns the row number of the table cell matches with the Content of the column.			 
			' @parameter:	sWindow				:	Name of the Session
			' @parameter:   sSAPTable  			:	Name of the table				
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Reurns the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				GetSAPGridRowbyCellContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").FindRowByCellContent(sColumn,sValue)	
				if Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPGridRowbyCellContent", "Content '"&sValue&"' found in Row No:"&GetSAPGridRowbyCellContent&" and Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "GetSAPGridRowbyCellContent", "Content '"&sValue&"' not found in Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPGridRowbyCellContent", "Value: '"&sValue&"'### Grid Column: '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPGridRowbyCellContent_UsingIndex ( iIndex, sColumn, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetTableRowbyContent ( sSAPTable, sColumn, sValue)
			' @returns	:	GetTableRowbyContent:	Reurns the row number of the table cell matches with the Content of the column.			 
			' @parameter:	sWindow				:	Name of the Session
			' @parameter:   sSAPTable  			:	Name of the table				
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Reurns the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				GetSAPGridRowbyCellContent_UsingIndex = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).FindRowByCellContent(sColumn,sValue)	
				if Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPGridRowbyCellContent_UsingIndex", "Content '"&sValue&"' found in Row No:"&GetSAPGridRowbyCellContent_UsingIndex&" and Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "GetSAPGridRowbyCellContent_UsingIndex", "Content '"&sValue&"' not found in Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPGridRowbyCellContent_UsingIndex", "Value: '"&sValue&"'### Grid Column: '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalGridRowbyCellContent ( sColumn, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetTableRowbyContent ( sSAPTable, sColumn, sValue)
			' @returns	:	GetTableRowbyContent:	Reurns the row number of the table cell matches with the Content of the column.			 
			' @parameter:	sWindow				:	Name of the Session
			' @parameter:   sSAPTable  			:	Name of the table				
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Reurns the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				GetSAPModalGridRowbyCellContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell").FindRowByCellContent(sColumn,sValue)	
				if Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPModalGridRowbyCellContent", "Content '"&sValue&"' found in Row No:"&GetSAPModalGridRowbyCellContent&" and Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "GetSAPModalGridRowbyCellContent", "Content '"&sValue&"' not found in Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPModalGridRowbyCellContent", "Value: '"&sValue&"'### Grid Column: '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalGridRowbyCellContent_UsingIndex ( iIndex,sColumn, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetTableRowbyContent ( sSAPTable, sColumn, sValue)
			' @returns	:	GetTableRowbyContent:	Reurns the row number of the table cell matches with the Content of the column.			 
			' @parameter:	sWindow				:	Name of the Session
			' @parameter:   sSAPTable  			:	Name of the table				
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Reurns the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				GetSAPModalGridRowbyCellContent_UsingIndex = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).FindRowByCellContent(sColumn,sValue)	
				if Err.Description = "" then
					UpdateExecutionReport micPass, "GetSAPModalGridRowbyCellContent_UsingIndex", "Content '"&sValue&"' found in Row No:"&GetSAPModalGridRowbyCellContent_UsingIndex&" and Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				else
					UpdateExecutionReport micFail, "GetSAPModalGridRowbyCellContent_UsingIndex", "Content '"&sValue&"' not found in Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPModalGridRowbyCellContent_UsingIndex", "Value: '"&sValue&"'### Grid Column: '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SelectSAPModalGridRowbyCellContent (sColumn,sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPModalGridRowbyCellContent (sColumn, sValue)
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Select the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				iGridRowbyCellContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell").FindRowByCellContent(sColumn,sValue)	
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("guicomponenttype:=201","name:=shell").SelectRow iGridRowbyCellContent
				If Err.Description = "" Then
					UpdateExecutionReport micPass, "SelectSAPModalGridRowbyCellContent", "Content '"&sValue&"' found in Row No:"&iGridRowbyCellContent&" and Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPModalGridRowbyCellContent", "Content '"&sValue&"' not found in Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPModalGridRowbyCellContent", "Value: '"&sValue&"'### Grid Column: '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
	Rem End Function
	'*************************************************************************************************************
	
	Function SelectSAPGridRowbyCellContent (sColumn,sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPModalGridRowbyCellContent (sColumn, sValue)
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Select the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				iGridRowbyCellContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").FindRowByCellContent(sColumn,sValue)	
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").SelectRow iGridRowbyCellContent
				If Err.Description = "" Then
					UpdateExecutionReport micPass, "SelectSAPGridRowbyCellContentt", "Content '"&sValue&"' found in Row No:"&iGridRowbyCellContent&" and Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SelectSAPGridRowbyCellContentt", "Content '"&sValue&"' not found in Column '"&sColumn&"' in Grid of screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SelectSAPGridRowbyCellContentt", "Value: '"&sValue&"'### Grid Column: '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
		
	'**************************************************************************************************************************************

	Rem Function SAPLabelRelatedFunctions()
		Function VerifySAPLabelExistance (sRelativeId)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPLabelExistance (sRelativeId)
			' @returns	:	VerifySAPLabelExistance	:	True/False
			' @parameter:	sRelativeId	:	Relative ID of the GUI Label
			' @notes	:	Returns the existance of the SAPGuiLabel in Main Window
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sSAPGuiLabelRelativeID).Exist Then
					VerifySAPLabelExistance = True
					UpdateExecutionReport micPass, "VerifySAPLabelExistance", "SAPLabel with RelativeId: '"&sRelativeId&"'### found in Main Window: '"&sScreenName&"'"
				else
					VerifySAPLabelExistance = False
					UpdateExecutionReport micPass, "VerifySAPLabelExistance", "SAPLabel with RelativeId: '"&sRelativeId&"'### not found in Main Window: '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "VerifySAPLabelExistance", "RelativeId: '"&sRelativeId&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
'**************************************************************************************************************************************
		Function VerifySAPLabelExistance_UsingContent (sContent)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPLabelExistance_UsingContent (sRelativeId)
			' @returns	:	VerifySAPLabelExistance_UsingContent	:	True/False
			' @parameter:	sRelativeId	:	Relative ID of the GUI Label
			' @notes	:	Returns the existance of the SAPGuiLabel in Main Window
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sContent).Exist Then
					VerifySAPLabelExistance_UsingContent = True
					UpdateExecutionReport micPass, "VerifySAPLabelExistance_UsingContent", "SAPLabel with Content: '"&sContent&"'### found in Main Window: '"&sScreenName&"'"
				else
					VerifySAPLabelExistance_UsingContent = False
					UpdateExecutionReport micPass, "VerifySAPLabelExistance_UsingContent", "SAPLabel with Content: '"&sContent&"'### not found in Main Window: '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "VerifySAPLabelExistance_UsingContent", "Content: '"&sContent&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
'**************************************************************************************************************************************
		Function VerifySAPModalLabelExistance_UsingContent (sContent)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPLabelExistance_UsingContent (sRelativeId)
			' @returns	:	VerifySAPLabelExistance_UsingContent	:	True/False
			' @parameter:	sRelativeId	:	Relative ID of the GUI Label
			' @notes	:	Returns the existance of the SAPGuiLabel in Main Window
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sContent).Exist Then
					VerifySAPModalLabelExistance_UsingContent = True
					UpdateExecutionReport micPass, "VerifySAPModalLabelExistance_UsingContent", "SAPLabel with Content: '"&sContent&"'### found in ModalWindow: '"&sScreenName&"'"
				else
					VerifySAPModalLabelExistance_UsingContent = False
					UpdateExecutionReport micPass, "VerifySAPModalLabelExistance_UsingContent", "SAPLabel with Content: '"&sContent&"'### not found in ModalWindow: '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "VerifySAPModalLabelExistance_UsingContent", "Content: '"&sContent&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function


'***********************************************************************************************************************************************************************************************

		Function VerifySAPModalLabelExistance (sRelativeId)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPModalLabelExistance (sRelativeId)
			' @returns	:	VerifySAPModalLabelExistance	:	True/False
			' @parameter:	sRelativeId						:	Relative ID of the GUI Label
			' @notes	:	Returns the existance of the SAPGuiLabel in Modal Window
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sSAPGuiLabelRelativeID).Exist Then
					VerifySAPModalLabelExistance = True
					UpdateExecutionReport micPass, "VerifySAPModalLabelExistance", "SAPLabel with RelativeId: '"&sRelativeId&"'### found in ModalWindow: '"&sScreenName&"'"
				else
					VerifySAPModalLabelExistance = False
					UpdateExecutionReport micPass, "VerifySAPModalLabelExistance", "SAPLabel with RelativeId: '"&sRelativeId&"'### not found in ModalWindow: '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				UpdateExecutionReport micFail, "VerifySAPModalLabelExistance", "RelativeId: '"&sRelativeId&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

'***********************************************************************************************************************************************************************************************

		Function GetSAPLabelContent ( sSAPGuiLabelRelativeID)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPLabelContent ( sRelativeId)
			' @returns	:	GetSAPLabelContent	:	String
			' @parameter:	sRelativeId				:	Relative ID of the GUI Label
			' @parameter:	sProperty				:	Name of the Property
			' @notes	:	Returns the content of the SAPGuiLabel in Main Window
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPLabelContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sSAPGuiLabelRelativeID).GetROProperty("content")	
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPLabelContent", "Content of SAPLabel with RelativeId '"& sSAPGuiLabelRelativeID &"' is :"&sProperty&" in MainWindow:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPLabelContent", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPLabelContent", "RelativeId: '"&sSAPGuiLabelRelativeID&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************
		
		Function GetSAPLabelRelativeID_UsingContent ( sContent)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPLabelRelativeID_UsingContent ( sContent)
			' @returns	:	GetSAPLabelRelativeID_UsingContent	:	String
			' @parameter:	sRelativeId				:	Relative ID of the GUI Label
			' @parameter:	sProperty				:	Name of the Property
			' @notes	:	Returns the content of the SAPGuiLabel in Main Window
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPLabelRelativeID_UsingContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sContent).GetROProperty("relativeid")	
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPLabelRelativeID_UsingContent", "Relative id of SAPLabel with Content '"& sContent &"' is :"&sProperty&" in MainWindow:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPLabelRelativeID_UsingContent", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPLabelRelativeID_UsingContent", "Content '"& sContent&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalLabelContent ( sRelativeId)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPLabelContent ( sRelativeId)
			' @returns	:	GetSAPLabelContent	:	String
			' @parameter:	sRelativeId				:	Relative ID of the GUI Label
			' @parameter:	sProperty				:	Name of the Property
			' @notes	:	Returns the content of the SAPGuiLabel		
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			GetSAPModalLabelContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sRelativeId).GetROProperty("content")	
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPModalLabelContent", "Content of SAPLabel with RelativeId '"& sRelativeId &"' is :"&sProperty&" in ModalWindow:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPModalLabelContent", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPModalLabelContent", "RelativeId: '"&sRelativeId&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalLabelContent_UsingId ( sId)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalLabelContent_UsingId ( sId)
			' @returns	:	GetSAPModalLabelContent_UsingId	:	String
			' @parameter:	sId				:	ID of the GUI Label
			' @notes	:	Returns the content of the SAPGuiLabel		
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			GetSAPModalLabelContent_UsingId = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","id:="&sId).GetROProperty("content")	
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPModalLabelContent_UsingId", "Content of SAPLabel with id '"& sId &"' is :"&sProperty&" in ModalWindow:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPModalLabelContent_UsingId", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPModalLabelContent_UsingId", "Id: '"&sId&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalLabelContent_RelativeId (sRelativeId)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalLabelContent_RelativeId (sRelativeId)
			' @returns	:	GetSAPModalLabelContent_RelativeId	:	String
			' @parameter:	sId				:	ID of the GUI Label
			' @notes	:	Returns the content of the SAPGuiLabel		
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			GetSAPModalLabelContent_RelativeId = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sRelativeId).GetROProperty("content")	
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPModalLabelContent_RelativeId", "Content of SAPLabel with id '"& sId &"' is :"&sProperty&" in ModalWindow:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPModalLabelContent_RelativeId", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPModalLabelContent_RelativeId", "Id: '"&sId&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************
Function SetSAPLabelFocus (sLabel)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	GetSAPLabelContent ( sRelativeId)
	' @returns	:	GetSAPLabelContent	:	String
	' @parameter:	sRelativeId				:	Relative ID of the GUI Label
	' @parameter:	sProperty				:	Name of the Property
	' @notes	:	Returns the content of the SAPGuiLabel		
	' @END
	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
       Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
	   sScreenName = oSAPObj.GetROProperty("text")
	   If oSAPObj.SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sLabel).Exist Then
	   		oSAPObj.SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sLabel).SetFocus
	   	ElseIf oSAPObj.SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sLabel).Exist Then
	   		oSAPObj.SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sLabel).SetFocus
	   End If
	ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
       Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
	   sScreenName = oSAPObj.GetROProperty("text")
	   If oSAPObj.SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sLabel).Exist Then
	   		oSAPObj.SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sLabel).SetFocus
	   	ElseIf oSAPObj.SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sLabel).Exist Then
	   		oSAPObj.SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sLabel).SetFocus
	   End If
	End if
	If Err.Description = "" then
	UpdateExecutionReport micPass, "SetSAPLabelFocus", "Focus set on SAPLabel with RelativeId '"& sRelativeId &"' in MainWindow:'"&sScreenName&"'"
	else
	UpdateExecutionReport micFail, "SetSAPLabelFocus", "Error Description:"&Err.Description
	End If
	Err.Clear
	On error goto 0	
End Function
		'***********************************************************************************************************************************************************************************************

		Function SetSAPLabelFocusUsingContent ( sContent)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPLabelContent ( sRelativeId)
			' @returns	:	GetSAPLabelContent	:	String
			' @parameter:	sRelativeId				:	Relative ID of the GUI Label
			' @parameter:	sProperty				:	Name of the Property
			' @notes	:	Returns the content of the SAPGuiLabel		
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sContent).SetFocus
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SetSAPLabelFocusUsingContent", "Focus set on SAPLabel with Content '"& sContent &"' in MainWindow:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "SetSAPLabelFocusUsingContent", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPLabelFocusUsingContent", "Content: '"&sContent&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	

		End Function
		
   '***********************************************************************************************************************************************************************************************		

	Function SetSAPModalGRidCellData_UsingF4Help(sCurrentCellColumn,sLabelName)
		
		If Not IsObject(application) Then
		   Set SapGuiAuto  = GetObject("SAPGUI")
		   Set application = SapGuiAuto.GetScriptingEngine
		End If
		If Not IsObject(connection) Then
		   Set connection = application.Children(0)
		End If
		If Not IsObject(SAPSession) Then
		   Set SAPSession    = connection.Children(0)
		End If
		If IsObject(WScript) Then
		   WScript.ConnectObject SAPSession,     "on"
		   WScript.ConnectObject application, "on"
		End If
		
		SAPSession.findById("wnd[1]/shellcont[1]/shell").currentCellColumn = sCurrentCellColumn
		SAPSession.findById("wnd[1]/shellcont[1]/shell").pressF4
		SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sLabelName).SetFocus
		SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SendKey ENTER

		Set SAPSession    = Nothing
		Set connection = Nothing
		Set SapGuiAuto  = Nothing

	End Function
	'***********************************************************************************************************************************************************************************************		
	Function SetSAPEdit_SelectingFromF4Help (sId,sLabelName)
		If Not IsObject(application) Then
		   Set SapGuiAuto  = GetObject("SAPGUI")
		   Set application = SapGuiAuto.GetScriptingEngine
		End If
		If Not IsObject(connection) Then
		   Set connection = application.Children(0)
		End If
		If Not IsObject(session) Then
		   Set SAPSession    = connection.Children(0)
		End If
		If IsObject(WScript) Then
		   WScript.ConnectObject session,     "on"
		   WScript.ConnectObject application, "on"
		End If
		SAPSession.findById("wnd[0]").maximize
		SAPSession.findById(sId).setFocus
		SAPSession.findById("wnd[0]").sendVKey 4
		SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sLabelName).SetFocus
		SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SendKey ENTER
		Set SAPSession    = Nothing
		Set connection = Nothing
		Set SapGuiAuto  = Nothing
	End Function
	'***********************************************************************************************************************************************************************************************		
    Function SetSAPLabelFocusUsingContent_Index ( sContent,iIndex)
	        ' @HELP
	        ' @class	:	Controls_SAP
	        ' @method	:	GetSAPLabelContent ( sRelativeId)
	        ' @returns	:	GetSAPLabelContent	:	String
	        ' @parameter:	sRelativeId				:	Relative ID of the GUI Label
	        ' @parameter:	sProperty				:	Name of the Property
	        ' @notes	:	Returns the content of the SAPGuiLabel		
	        ' @END
	        On error resume next
	        If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
    	    sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
	        SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sContent,"index:="&iIndex).SetFocus
	        If Err.Description = "" then
	        UpdateExecutionReport micPass, "SetSAPLabelFocusUsingContent", "Focus set on SAPLabel with Content '"& sContent &"' in MainWindow:'"&sScreenName&"'"
	        else
   	        UpdateExecutionReport micFail, "SetSAPLabelFocusUsingContent", "Error Description:"&Err.Description
    	    End If
	        ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
	        sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
	        UpdateExecutionReport micFail, "SetSAPLabelFocusUsingContent", "Content: '"&sContent&"'### Active window is Modal Window: '"&sScreenName&"'"
	        End If
	        Err.Clear
	        On error goto 0	
    
	    End Function
		
    '***********************************************************************************************************************************************************************************************

		Function SetSAPModalLabelFocusUsingContent ( sContent)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPLabelContent ( sRelativeId)
			' @returns	:	GetSAPLabelContent	:	String
			' @parameter:	sRelativeId				:	Relative ID of the GUI Label
			' @parameter:	sProperty				:	Name of the Property
			' @notes	:	Returns the content of the SAPGuiLabel		
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","content:="&sContent).SetFocus
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SetSAPModalLabelFocusUsingContent", "Focus set on SAPLabel with Content '"& sContent &"' in Modal Window:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "SetSAPModalLabelFocusUsingContent", "Label Content = "&sContent&" ### Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPModalLabelFocusUsingContent", "Content: '"&sContent&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPModalLabelFocus ( sRelativeId)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPLabelContent ( sRelativeId)
			' @returns	:	GetSAPLabelContent	:	String
			' @parameter:	sRelativeId				:	Relative ID of the GUI Label
			' @parameter:	sProperty				:	Name of the Property
			' @notes	:	Returns the content of the SAPGuiLabel		
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sRelativeId).SetFocus
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SetSAPModalLabelFocus", "Focus set on SAPLabel with RelativeId '"& sRelativeId &"' in Modal Window:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "SetSAPModalLabelFocus", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPModalLabelFocus", "RelativeId: '"&sRelativeId&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPLabel ( sRelativeId, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPLabel ( sRelativeId)
			' @returns	:	GetSAPLabelContent	:	String
			' @parameter:	sRelativeId				:	Relative ID of the GUI Label
			' @parameter:	sProperty				:	Name of the Property
			' @notes	:	Returns the content of the SAPGuiLabel		
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sSAPGuiLabelRelativeID).Set sValue
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SetSAPLabel", "Value '"&sValue&"' set in SAPLabel with RelativeId '"& sRelativeId &"' in MainWindow:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "SetSAPLabel", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPLabel", "RelativeId: '"&sRelativeId&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPModalLabel ( sRelativeId, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPLabelContent ( sRelativeId)
			' @returns	:	GetSAPLabelContent	:	String
			' @parameter:	sRelativeId				:	Relative ID of the GUI Label
			' @parameter:	sProperty				:	Name of the Property
			' @notes	:	Returns the content of the SAPGuiLabel		
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sSAPGuiLabelRelativeID).Set sValue
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SetSAPModalLabel", "Value '"&sValue&"' set in SAPLabel with RelativeId '"& sRelativeId &"' in Modal Window:'"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "SetSAPModalLabel", "Error Description:"&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPModalLabel", "RelativeId: '"&sRelativeId&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
		
			'**********************************************************************************************************************************************************************************************************
		Function GetSAPLableId_UsingContent ( sContent)
				 On error resume next
				
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					
					GetSAPLableId_UsingContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("content:="&sContent,"guicomponenttype:=30").GetROProperty ("id")
					
					If err.description = "" Then
					   UpdateExecutionReport micPass, "GetSAPLableId_UsingContent","Selected item '"&GetSAPLableId_UsingContent&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					else
						UpdateExecutionReport micFail, "GetSAPLableId_UsingContent","Screen:"&sScreenName&"###Error Description:"&err.description
				End If
				   ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22","index:="&iIndex).GetROProperty("text")
					UpdateExecutionReport micFail, "GetSAPLableId_UsingContent", "SAP Edit: '"& sSAPGUIEdit &"' ### Active Window is Main Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0

		End Function
		
		'************************************************************************************************************************************************************************
		
		Function GetSAPLableRelativeid_UsingContent ( sContent)
				 On error resume next
				
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					
					GetSAPLableRelativeid_UsingContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("content:="&sContent,"guicomponenttype:=30").GetROProperty ("relativeid")
					
					If err.description = "" Then
					   UpdateExecutionReport micPass, "GetSAPLableRelativeid_UsingContent","Selected item '"&GetSAPLableId_UsingContent&"' found in Edit '"&sSAPGUIEdit&"' in Screen '"&sScreenName&"'"
					else
						UpdateExecutionReport micFail, "GetSAPLableRelativeid_UsingContent","Screen:"&sScreenName&"###Error Description:"&err.description
				End If
				   ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22","index:="&iIndex).GetROProperty("text")
					UpdateExecutionReport micFail, "GetSAPLableRelativeid_UsingContent", "SAP Edit: '"& sSAPGUIEdit &"' ### Active Window is Main Window '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0

		End Function


	Rem End Function
	'***********************************************************************************************************************************************************************************************
	'***********************************************************************************************************************************************************************************************
	Rem Function SAPRadioButtonRelatedFunctions()
		Function SelectSAPRadioButton ( sSAPGuiRadioButton,sRadioButtonDesc)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPRadioButton ( sSAPGuiRadioButton)
			' @parameter:   sSAPGUIRadioButton 	:	Name of the Radio button 	
			' @notes	:	Selects Radio Button on Main Window
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41").Exist(0) Then
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41").Set
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("attachedtext:="&sRadioButtonDesc,"guicomponenttype:=41").Exist(0) Then
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("attachedtext:="&sRadioButtonDesc,"guicomponenttype:=41").Set 
			End If	
			If err.description = "" Then
			UpdateExecutionReport micPass, "SelectSAPRadioButton","ScreenName = "&sScreenName&"; RadioButton '"&sRadioButtonDesc&"' selected"
			else
			UpdateExecutionReport micFail, "SelectSAPRadioButton","Screen:'"&sScreenName&"' ### RadioButton:'"&sRadioButtonDesc&"'###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPRadioButton", "Property: '"&sProperty&"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SelectSAPModalRadioButton ( sSAPGuiRadioButton,iIndex)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPRadioButton ( sSAPGuiRadioButton)
			' @parameter:   sSAPGUIRadioButton 	:	Name of the Radio button 	
			' @notes	:	Selects Radio Button on Main Window
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41","index:="&iIndex).Set
			If err.description = "" Then
			UpdateExecutionReport micPass, "SelectSAPModalRadioButton","ScreenName = "&sScreenName&"; RadioButton '"&sSAPGuiRadioButton&"' selected"
			else
			UpdateExecutionReport micFail, "SelectSAPModalRadioButton","Screen:"&sScreenName&"### RadioButton:'"&sSAPGuiRadioButton&"'###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPModalRadioButton", "Property: '"&sProperty&"' ### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function VerifySAPRadioButtonExistance ( sSAPGuiRadioButton)
		
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPRadioButton ( sSAPGuiRadioButton)
			' @parameter:   sSAPGUIRadioButton 	:	Name of the Radio button 	
			' @notes	:	Verify the existance of Radio Button on Main Window
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41").Exist Then
			VerifySAPRadioButtonExistance = True
			UpdateExecutionReport micPass, "VerifySAPRadioButtonExistance","ScreenName = "&sScreenName&"; RadioButton '"&sSAPGuiRadioButton&"' Exist"
				Else
			VerifySAPRadioButtonExistance = False
			UpdateExecutionReport micPass, "VerifySAPRadioButtonExistance","ScreenName = "&sScreenName&"; RadioButton '"&sSAPGuiRadioButton&"' doesnot exist"
			End If
			If err.description <> "" Then
				UpdateExecutionReport micFail, "VerifySAPRadioButtonExistance","Screen:"&sScreenName&"### RadioButton:'"&sSAPGuiRadioButton&"'###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "VerifySAPRadioButtonExistance", "SAPGuiRadioButton: '"&sSAPGuiRadioButton&"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On Error goto 0
			
		End Function

		'***********************************************************************************************************************************************************************************************

		Function VerifySAPModalRadioButton ( sSAPGuiRadioButton)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifySAPModalRadioButton ( sSAPGuiRadioButton)
			' @parameter:   sSAPGUIRadioButton 	:	Name of the Radio button 	
			' @notes	:	Verify the existance of Radio Button on Main Window
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41").Exist Then
			VerifySAPModalRadioButton = True
			UpdateExecutionReport micPass, "VerifySAPModalRadioButton","ScreenName = "&sScreenName&"; RadioButton '"&sSAPGuiRadioButton&"' Exist"
			Else
			VerifySAPModalRadioButton = False
			UpdateExecutionReport micPass, "VerifySAPModalRadioButton","ScreenName = "&sScreenName&"; RadioButton '"&sSAPGuiRadioButton&"' doesnot exist"
			End If
			If err.description <> "" Then
			UpdateExecutionReport micFail, "VerifySAPModalRadioButton","Screen:"&sScreenName&"### RadioButton:'"&sSAPGuiRadioButton&"'###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "VerifySAPModalRadioButton", "Property: '"&sProperty&"' ### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
		
		'**************************************************************************************************************************************************************************
		Sub SelectSAPRadioButtonUsingIndex ( sSAPGuiRadioButton,iIndex)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPRadioButton_ByIndex ( sSAPGuiRadioButton,iIndex)
			' @parameter:   sSAPGUIRadioButton	:	Name of the Radio Button
			' @parameter:   iIndex				:	Index of the Radio Button
			' @notes	:	Selects the Specified SAPGUIRadioButton on SAP Main Window
			' @END
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41","index:="&iIndex).Exist Then 
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41","index:="&iIndex).Set
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41","index:="&iIndex).SetFocus 
							UpdateExecutionReport micPass, "SelectSAPRadioButtonUsingIndex","Select action performed on SAPRadioButton of name '"&sSAPGuiRadioButton&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingIndex","Select action couldnot be performed on SAPRadioButton of name '"&sSAPGuiRadioButton&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingIndex","Select action couldnot be performed on SAPRadioButton of name '"&sSAPGuiRadioButton&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingIndex","SAPRadioButton of name '"&sSAPGuiRadioButton&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
		End Sub

	'**************************************************************************************************************************************************************************

		Sub SelectSAPRadioButtonUsingId (iId)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPRadioButton_ById ( iId)
			' @parameter:   iId		:	Id of the SAPGUIRadioButton on which action is to be performed
			' @notes	:	Clicks the specified SAPGUIRadioButton on SAP Main window 
			' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("guicomponenttype:=41","id:="&iId).Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("guicomponenttype:=41","id:="&iId).Set
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("guicomponenttype:=41","id:="&iId).SetFocus
						UpdateExecutionReport micPass, "SelectSAPRadioButtonUsingId","Select action performed on SAPRadioButton of ID '"&iId&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingId","Select action couldnot be performed on SAPRadioButton of ID '"&iId&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingId","Select action couldnot be performed on SAPRadioButton of ID '"&iId&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingId","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
		End Sub
		
		Sub SelectSAPRadioButtonUsingNameAndIndex ( sSAPGuiRadioButton,sindex)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPRadioButton_ByName ( sSAPGuiRadioButton)
			' @parameter:	sWindow				:	Name of the Session
			' @parameter:   sSAPGUIRadioButton	:	Name of the radio button
			' @notes:	 	Selects Radiobutton using name property 
			' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41","index:="&sindex).Exist then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41","index:="&sindex).Set
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiRadioButton("name:="&sSAPGuiRadioButton,"guicomponenttype:=41","index:="&sindex).SetFocus 
						UpdateExecutionReport micPass, "SelectSAPRadioButtonUsingNameAndIndex","SAPRadioButton with index '"&sIndex&"' selected in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingNameAndIndex","SAPRadioButton with index '"&sIndex&"' not selected in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingNameAndIndex","SAPRadioButton with index '"&sIndex&"' not selected as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPRadioButtonUsingNameAndIndex","SAPRadioButton with index '"&sIndex&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				on error goto 0
		End Sub
	Rem End Function
	'***********************************************************************************************************************************************************************************************

'Rem Function SAPTabStripRelatedFunctions()
Function SelectSAPTabStrip ( sSAPTabStrip,sTab)
	' @HELP
	' @class:		Controls_SAP
	' @method:		SelectSAPTabStrip (sSAPTabStrip, sTab)
	' @parameter:	sWindow		:	Name of the Session
	' @parameter:   sSAPTabStrip	:	Name of the TabStrip  
	' @parameter:   sTab            :	Name of the Tab to be selected
	' @notes:		Selects the required Tab of the specified TabStrip	 
	' @END 
	CaptureSAPScreenShot()
	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
		Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiTabStrip("name:="&sSAPTabStrip,"type:=GuiTabStrip","guicomponenttype:=90").Select sTab
		If err.description = "" Then
		UpdateExecutionReport micPass,"SelectSAPTabStrip","ScreenName: '"&sScreenName&"'; Tabstrip '"& sSAPTabStrip &"' with Tab '"&sTab&"' selected"
		End if
	Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
		Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
		sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiTabStrip("name:="&sSAPTabStrip,"type:=GuiTabStrip","guicomponenttype:=90").Select sTab
		If err.description = "" Then
		UpdateExecutionReport micPass,"SelectSAPTabStrip","ScreenName: '"&sScreenName&"'; Tabstrip '"& sSAPTabStrip &"' with Tab '"&sTab&"' selected"
		End if
	End if
	If err.description <> "" Then
		UpdateExecutionReport micFail,"SelectSAPTabStrip","Screen: "&sScreenName & ";Tabstrip '"& sSAPTabStrip &"' and Tab '"&sTab&"' selected ###ErrorDescription:"& err.description
	End If
	Err.Clear
	On error goto 0	
End Function

	'***********************************************************************************************************************************************************************************************

		Function SelectSAPModalTabStrip ( sSAPTabStrip,sTab)
			' @HELP
			' @class:		Controls_SAP
			' @method:		SelectSAPModalTabStrip (sSAPTabStrip, sTab)
			' @parameter:   sSAPTabStrip	:	Name of the TabStrip  
			' @parameter:   sTab            :	Name of the Tab to be selected
			' @notes:		Selects the required Tab of the specified TabStrip	 
			' @END 
            CaptureSAPScreenShot()
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTabStrip("name:="&sSAPTabStrip,"type:=GuiTabStrip","guicomponenttype:=90").Select sTab	
			If err.description = "" Then
			UpdateExecutionReport micPass,"SelectSAPModalTabStrip","ScreenName: '"&sScreenName&"'; Tabstrip '"& sSAPTabStrip &"' with Tab '"&sTab&"' selected"
			else
			UpdateExecutionReport micFail, "SelectSAPModalTabStrip","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPModalTabStrip", "SAPTabStrip: '"&sSAPTabStrip&"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function VerifySAPTabStripExistance ( sSAPTabStrip)
			' @HELP
			' @class:		Controls_SAP
			' @method:		VerifySAPTabStrip (sSAPTabStrip)
			' @parameter:   sSAPTabStrip	:	Name of the TabStrip
			' @notes:		Selects the required Tab of the specified TabStrip	 
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			if SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTabStrip("name:="&sSAPTabStrip,"type:=GuiTabStrip","guicomponenttype:=90").Exist Then
			VerifySAPTabStripExistance = True
			UpdateExecutionReport micPass,"VerifySAPTabStripExistance","ScreenName: '"&sScreenName&"'; Tabstrip '"& sSAPTabStrip &"' Exist"
			else
			VerifySAPTabStripExistance = False
			UpdateExecutionReport micPass,"VerifySAPTabStripExistance","ScreenName: '"&sScreenName&"'; Tabstrip '"& sSAPTabStrip &"' doesnot exist"
			End If
			If err.description <> "" Then
			UpdateExecutionReport micFail, "VerifySAPTabStripExistance","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "VerifySAPTabStripExistance", "SAPTabStrip: '"&sSAPTabStrip&"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		Function VerifySAPModalTabStripExistance ( sSAPTabStrip)
			' @HELP
			' @class:		Controls_SAP
			' @method:		VerifySAPTabStrip (sSAPTabStrip)
			' @parameter:   sSAPTabStrip	:	Name of the TabStrip
			' @notes:		Selects the required Tab of the specified TabStrip	 
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			if SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTabStrip("name:="&sSAPTabStrip,"type:=GuiTabStrip","guicomponenttype:=90").Exist Then
			VerifySAPModalTabStrip = True
			UpdateExecutionReport micPass,"VerifySAPModalTabStripExistance","ScreenName: '"&sScreenName&"'; Tabstrip '"& sSAPTabStrip &"' Exist"
			else
			VerifySAPModalTabStrip = False
			UpdateExecutionReport micPass,"VerifySAPModalTabStripExistance","ScreenName: '"&sScreenName&"'; Tabstrip '"& sSAPTabStrip &"' doesnot exist"
			End If
			If err.description <> "" Then
			UpdateExecutionReport micFail, "VerifySAPModalTabStripExistance","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "VerifySAPModalTabStripExistance", "SAPTabStrip: '"&sSAPTabStrip&"' ### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
	Rem End Function
	'***********************************************************************************************************************************************************************************************

'**************************************************************************************************************************************************************************
Rem Function SAPTAbleRelatedFunctions()
Function SetSAPTableCellData ( sSAPTable, sRow, sColumn, sData )
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	SetSAPTableCellData ( sSAPTable, sRow, sColumn, sData )
	' @parameter:	sRow	:	Row number
	' @parameter:	sColumn	:	Coloumn Name
	' @parameter:	sData	:   Value to be inserted 
	' @parameter:	sSAPTable  :	Name of the table
	' @notes	:   Inserts given data into specified Table cell
	' @END
	
	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then	
	Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
	sScreenName = oSAPObj.GetROProperty("text")
		Set oDesc = Description.Create()
		oDesc("name").Value = sSAPTable
		if SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).Exist then
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).SetCellData sRow,sColumn,sData			
		Else
			Set oDesc = Description.Create()
			oDesc("text").Value = sSAPTable
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).SetCellData sRow,sColumn,sData
		End if	
	Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then	
	Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
	sScreenName = oSAPObj.GetROProperty("text")
		Set oDesc = Description.Create()
		oDesc("name").Value = sSAPTable
		if SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).Exist then
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).SetCellData sRow,sColumn,sData			
		Else
			Set oDesc = Description.Create()
			oDesc("text").Value = sSAPTable
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).SetCellData sRow,sColumn,sData
		End if
	End if
	If Err.Description = "" then
		UpdateExecutionReport micPass, "SetSAPTableCellData", "Value '"&sData&"' set into Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
	Else
		UpdateExecutionReport micFail, "SetSAPTableCellData", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
	End If
	Err.Clear
	On error goto 0		
End Function
        
		'***********************************************************************************************************************************************************************************************

		Function SelectSAPTableCell ( sSAPTable, sRow, sColumn)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPTableCellData ( sSAPTable, sRow, sColumn, sData )
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   Inserts given data into specified Table cell
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).SelectCell sRow,sColumn
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SelectSAPTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'has been selected"
			Else
			UpdateExecutionReport micFail, "SelectSAPTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SelectSAPModalTableCell ( sSAPTable, sRow, sColumn)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPTableCellData ( sSAPTable, sRow, sColumn, sData )
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   Inserts given data into specified Table cell
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).SelectCell sRow,sColumn
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SelectSAPModalTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"' of Modal Window '"&sScreenName&"'has been selected"
			Else
			UpdateExecutionReport micFail, "SelectSAPModalTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Modal Window: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPModalTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function SetSAPModalTableCellData ( sSAPTable, sRow, sColumn, sData )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPModalTableCellData ( sSAPTable, sRow, sColumn, sData )
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sData	:   Value to be inserted 
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   Inserts given data into specified Table cell
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 or  InStr(sSAPTable,"SAP") then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).SetCellData sRow,sColumn,sData
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SetSAPModalTableCellData", "Value '"&sData&"' set into Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "SetSAPModalTableCellData", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Modal Window: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPModalTableCellData", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPTableCellData ( sSAPTable, sRow, sColumn )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPTableCellData ( sSAPTable, sRow, sColumn )
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   Returns value of a speciied Table Cell
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 or InStr(sSAPTable,"SAP") then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPTableCellData = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).GetCellData (sRow,sColumn)
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPTableCellData", "Value '"&GetSAPTableCellData&"' captured from Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "GetSAPTableCellData", "Value could not be captured of Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPTableCellData", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPTableCellData_ByValue ( sSAPTable,sValue,sReferenceColumn,sRequiredColumn)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPTableCellData ( sSAPTable, sRow, sColumn )
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   Returns value of a speciied Table Cell
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPTableRowbyCellContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).FindRowByCellContent(sReferenceColumn,sValue)	
			GetSAPTableCellData_ByValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).GetCellData(SAPTableRowbyCellContent,sRequiredColumn)	
			If Err.Description = "" Then
			UpdateExecutionReport micPass, "GetSAPTableCellData_ByValue", "Value '"&GetSAPTableCellData_ByValue&"' captured from Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "GetSAPTableCellData_ByValue", "Value could not be captured of Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### MainWindow: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPTableCellData_ByValue", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'*****************************************************************************************************************************
		Function GetSAPModalTableCellData ( sSAPTable, sRow, sColumn )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalTableCellData ( sSAPTable, sRow, sColumn )
			' @parameter:	sRow	:	Row number
			' @parameter:	sColumn	:	Coloumn Name
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   Returns value of a speciied Table Cell
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 or InStr(sSAPTable,"TABLE") > 0 or InStr(sSAPTable,"SAP") then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			GetSAPModalTableCellData = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).GetCellData (sRow,sColumn)
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPModalTableCellData", "Value '"&GetSAPModalTableCellData&"' captured from Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"' of Modal Window '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "GetSAPModalTableCellData", "Value could not be captured of Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Modal Window: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPModalTableCellData", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

Function SelectSAPTableRow ( sSAPTable, sRow )
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	SelectSAPTableRow ( sSAPTable, sRow )
	' @parameter:	sRow	:	Row number
	' @parameter:	sSAPTable  :	Name of the table
	' @notes	:   Selects a Table Row
	' @END 
	On error resume next
	If InStr(sSAPTable,"_") > 0 then
		Set oDesc = Description.Create()
		oDesc("name").Value = sSAPTable
	Else
		Set oDesc = Description.Create()
		oDesc("text").Value = sSAPTable
	End if

	If 	SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then	
		Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiTable(oDesc).selectRow sRow
		If Err.Description = "" then
		UpdateExecutionReport micPass, "SelectSAPTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' selected in MainWindow '"&sScreenName&"'"
		Else
		UpdateExecutionReport micFail, "SelectSAPTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' couldnot be selected ### Error Description : "&Err.Description
		End If
	ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then	
		Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
		sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiTable(oDesc).selectRow sRow
		If Err.Description = "" then
		UpdateExecutionReport micPass, "SelectSAPTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' selected in MainWindow '"&sScreenName&"'"
		Else
		UpdateExecutionReport micFail, "SelectSAPTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' couldnot be selected ### Error Description : "&Err.Description
		End If
	End if	
	Err.Clear
	On error goto 0		
End Function

		'***********************************************************************************************************************************************************************************************

		Function SelectSAPModalTableRow ( sSAPTable, sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPModalTableRow ( sSAPTable, sRow )
			' @parameter:	sRow	:	Row number
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   Selects a Table Row
			' @END 
			On error resume next

			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).selectRow sRow
			If Err.Description = "" then
			UpdateExecutionReport micPass, "SelectSAPModalTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' selected in Modal Window '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "SelectSAPModalTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' couldnot be selected ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPModalTableRow", "Row '"&sRow&"'###Table '"&sSAPTable&"'###Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function DeSelectSAPTableRow ( sSAPTable, sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	DeSelectSAPTableRow ( sSAPTable, sRow )
			' @parameter:	sRow	:	Row number
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   DeSelects a Table Row
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).DeselectRow sRow
			If Err.Description = "" then
			UpdateExecutionReport micPass, "DeSelectSAPTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' deselected in MainWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "DeSelectSAPTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' couldnot be deselected ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "DeSelectSAPTableRow", "Row '"&sRow&"'###Table '"&sSAPTable&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function DeSelectSAPModalTableRow ( sSAPTable, sRow )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	DeSelectSAPModalTableRow ( sSAPTable, sRow )
			' @parameter:	sRow	:	Row number
			' @parameter:	sSAPTable  :	Name of the table
			' @notes	:   DeSelects a Table Row
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).DeselectRow sRow
			If Err.Description = "" then
			UpdateExecutionReport micPass, "DeSelectSAPModalTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' deselected in MainWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "DeSelectSAPModalTableRow", "Row '"&sRow&"' of Table '"&sSAPTable&"' couldnot be deselected ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "DeSelectSAPModalTableRow", "Row '"&sRow&"'###Table '"&sSAPTable&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPTableRowCount ( sSAPTable )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPTableRowCount ( sSAPTable )
			' @parameter:   sSAPTable  :	Name of the table
			' @notes	:   Returns rowcount of Table
			' @END
			On error resume next
			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPTableRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).Rowcount
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPTableRowCount", "Rows '"&GetSAPTableRowCount&"' exist for Table '"&sSAPTable&"' of MainWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "GetSAPTableRowCount","MainWindow: '"&sScreenName&"'### Table '"&sSAPTable&"'### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPTableRowCount", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPModalTableRowCount ( sSAPTable )
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetSAPModalTableRowCount ( sSAPTable )
			' @parameter:   sSAPTable  			:	Name of the table
			' @notes	:   Returns rowcount of Table
			' @END
			On error resume next
			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			GetSAPModalTableRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).Rowcount
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPModalTableRowCount", "Rows '"&GetSAPModalTableRowCount&"' exist for Table '"&sSAPTable&"' of ModalWindow '"&sScreenName&"'"
			Else
			UpdateExecutionReport micFail, "GetSAPModalTableRowCount","ModalWindow: '"&sScreenName&"' ### Table '"&sSAPTable&"' ### Error Description : "&Err.Description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPModalTableRowCount", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function

		'***********************************************************************************************************************************************************************************************

		Function GetSAPTableRowbyCellContent ( sSAPTable, sColumn, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetTableRowbyContent ( sSAPTable, sColumn, sValue)
			' @returns	:	Reurns the row number of the table cell matches with the Content of the column.	
			' @parameter:   sSAPTable  			:	Name of the table				
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Reurns the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 or InStr(sSAPTable,"SAP") then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPTableRowbyCellContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).FindRowByCellContent(sColumn,sValue)	
			if Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPTableRowbyCellContent", "Content '"&sValue&"' found in Row No:"&GetSAPTableRowbyCellContent&" and Column '"&sColumn&"' in Table '"&sSAPTable&"' of screen '"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPTableRowbyCellContent", "Content '"&sValue&"' not found in Column '"&sColumn&"' in Table '"&sSAPTable&"' of screen '"&sScreenName&"'"
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPTableRowbyCellContent", "Value: '"&sValue&"'### Table '"&sSAPTable&"'### Column: '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
         
		 
		 '******************************************************************************************************************************************************************************
		 
		 

Function GetSAPModalTableRowbyCellContent ( sSAPTable, sColumn, sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetTableRowbyContent ( sSAPTable, sColumn, sValue)
			' @returns	:	Reurns the row number of the table cell matches with the Content of the column.	
			' @parameter:   sSAPTable  			:	Name of the table				
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Reurns the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 or InStr(sSAPTable,"SAP") then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPModalTableRowbyCellContent = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable(oDesc).FindRowByCellContent(sColumn,sValue)	
			if Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPModalTableRowbyCellContent", "Content '"&sValue&"' found in Row No:"&GetSAPTableRowbyCellContent&" and Column '"&sColumn&"' in Table '"&sSAPTable&"' of screen '"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPModalTableRowbyCellContent", "Content '"&sValue&"' not found in Column '"&sColumn&"' in Table '"&sSAPTable&"' of screen '"&sScreenName&"'"
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "GetSAPModalTableRowbyCellContent", "Value: '"&sValue&"'### Table '"&sSAPTable&"'### Column: '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
		
		'***********************************************************************************************************************************************************************************************
		Function SetSAPTableRowByCellContent (sSAPTable,sRefColumn,sRefValue,sActualColumn,sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetTableRowbyContent ( sSAPTable, sColumn, sValue)
			' @returns	:	Reurns the row number of the table cell matches with the Content of the column.	
			' @parameter:   sSAPTable  			:	Name of the table				
			' @parameter:   sColumn    			:	Name of the column in which value to be verified against the cell content.	  
			' @parameter:   sValue     			:	Value to be varified 
			' @notes	:	Reurns the row number of the table cell matches with the Content of the column.
			' @END 
			On error resume next
			If InStr(sSAPTable,"_") > 0 or InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			iRowNumber = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).FindRowByCellContent(sRefColumn,sRefValue)	
			if Err.Description = "" then
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).SetCellData iRowNumber,sActualColumn,sValue
			UpdateExecutionReport micPass, "SetSAPTableRowByCellContent", "Content '"&sValue&"' set in Row No:"&iRowNumber&" and Column '"&sActualColumn&"' in Table '"&sSAPTable&"' of screen '"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "SetSAPTableRowByCellContent", "Content '"&sRefValue&"' not found in Column '"&sColumn&"' in Table '"&sSAPTable&"' of screen '"&sScreenName&"'"
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPTableRowByCellContent", "Value: '"&sValue&"'### Table '"&sSAPTable&"'### Column: '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function

		'***********************************************************************************************************************************************************************************************

		
		
		Function VerifySAPTableCellEditable( sSAPTable,iRow,sColumn)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	IsTableCellEditable( sSAPTable,iRow,sColumn)
			' @returns	:	IsTableCellEditable	:	TRUE/FALSE
			' @parameter:	sWindow				:	Name of the Session
			' @parameter:   sSAPTable			:	Name of the Column	
			' @parameter:   iRow				:  	Row Number
			' @parameter:   sColumn				: 	Name of the column
			' @notes	:	Returns TRUE or FALSE on whether the specified column in the specified row is editable or not
			' @END 

			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			VerifySAPTableCellEditable = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).IsCellEditable(iRow,sColumn)
			If Err.Description = "" then
			UpdateExecutionReport micPass, "VerifySAPTableCellEditable", "Editable Status of TableCell of Table '"&sSAPTable&"', Row '"&iRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"' is '"&VerifySAPTableCellEditable&"'"
			Else
			UpdateExecutionReport micFail, "VerifySAPTableCellEditable", "Table '"&sSAPTable&"', Row '"&iRow&"' & Column '"&sColumn&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
		End Function

		'-----------------------------------------------------------------------------------------------------------------------------------------	


		Function GetSAPTableFirstEditableRowNumber ( sSAPTable,sColumn)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	GetFirstEditableRowNumber ( sSAPTable,sColumn)
			' @returns	:	GetFirstEditableRowNumber	:	Integer
			' @parameter:	sWindow						:	Name of the Session
			' @parameter:	sSAPTable					:	Name of the Table
			' @parameter:	sColumn						:	Name of the Column
			' @notes	:	Returns the first editable row ,compares the first column to see if it is editable		
			' @END 

			If InStr(sSAPTable,"_") > 0 then
				Set oDesc = Description.Create()
				oDesc("name").Value = sSAPTable
			Else
				Set oDesc = Description.Create()
				oDesc("text").Value = sSAPTable
			End if
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			iRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).RowCount
			bVal = False
			iCount = 1
			Do While bVal = False
			For iRow = 3 to iRowCount
				If (SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).IsCellEditable(iRow,sColumn)) then 
				
					If (SAPGuiSession(oSesObjDescriptions).SAPGuiWindow(oObjectDescriptions).SAPGuiTable(oDesc).GetCellData (iRow,sColumn) = "") then  
						iCount =iRow
						bVal =True        
						Exit do 
					End If 
				End if
			Next
			Loop 
			
			GetSAPTableFirstEditableRowNumber = iCount
			If Err.Description = "" then
			UpdateExecutionReport micPass, "GetSAPTableFirstEditableRowNumber", "First Editable Row of TableCell of Table '"&sSAPTable&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"' is '"&GetSAPTableFirstEditableRowNumber&"'"
			Else
			UpdateExecutionReport micFail, "GetSAPTableFirstEditableRowNumber", "Table '"&sSAPTable&"'& Column '"&sColumn&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
			End If
		End Function
		
		Sub ClickSAPTableCell ( sSAPTable,iRow,sColumn)
			' @HELP
			' @class:  		Controls_SAP
			' @method:  	ClickSAPTableCell ( sSAPTable,iRow,sColumn)
			' @parameter:   sSAPTable   :	Name of the field
			' @parameter:   iRow        :	Row number
			' @parameter:   sColumn     :	Column name
			' @notes: 		Clicks on the Specified SAPGUITableCell on SAP Main Screen
			' @END
			
			If InStr(sSAPTable,"_") > 0 then
			Set oDesc = Description.Create()
			oDesc("name").Value = sSAPTable
			Else
			Set oDesc = Description.Create()
			oDesc("text").Value = sSAPTable
			End if
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).Exist Then 
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).ClickCell iRow, sColumn
					UpdateExecutionReport micPass, "ClickSAPTableCell","Click action performed on SAPGUITable '"&sSAPTable&"' of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "ClickSAPTableCell","Select action couldnot be performed on SAPGUITable '"&sSAPTable&"' of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "ClickSAPTableCell","Select action couldnot be performed on SAPGUITable '"&sSAPTable&"' of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Modal Window '"&sScreenName&"'"
			End If
			If err.description <> "" Then
				UpdateExecutionReport micFail, "ClickSAPTableCell","SAPGUITable '"&sSAPTable&"' of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0			
		End Sub 
	'-------------------------------------------------------------------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------------------------------------------
			'Used to check if table with the specified name exists (needs the 'Name' to be passed as parameter).
			Function VerifySAPTableExistance( sSAPTable)
				' @HELP
				' @class:		Controls_SAP
				' @method:      SAPGuiTableExists(sSAPTable)
				' @returns:     True if container with specified Id exists and false if it does not
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:	sId	: Id of the Table                       
				' @notes:          
				' @END            
				If InStr(sSAPTable,"_") > 0 then
					Set oDesc = Description.Create()
					oDesc("name").Value = sSAPTable
				Else
					Set oDesc = Description.Create()
					oDesc("text").Value = sSAPTable
				End if 
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).Exist Then
				VerifySAPTableExistance = True
				Else
				VerifySAPTableExistance = False
				End If
				If Err.Description = "" then
				UpdateExecutionReport micPass, "VerifySAPTableExistance", "Existance of SAPGUITable '"&sSAPTable&"' in MainWindow '"&sScreenName&"' is '"&VerifySAPTableExistance&"'"
				Else
				UpdateExecutionReport micFail, "VerifySAPTableExistance", "SAPGUITable '"&sSAPTable&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			End Function

			'-----------------------------------------------------------------------------------------------------------------------------------------
			''''''To get Table Items Count
			Function GetSAPTableRowsPopulated ( sSAPTable,sColumn)
				' @HELP
				' @class:		Controls_SAP
				' @method:   	GetItemsCount (sSAPTable,sColumn)
				' @returns:  	GetItemsCount	:	Integer
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPTable		: 	Name of the Table
				' @parameter:   sColumn			:	Name of the Column
				' @notes:   	Returns the populated Items Count of the given table
				' @END  
				If InStr(sSAPTable,"_") > 0 then
					Set oDesc = Description.Create()
					oDesc("name").Value = sSAPTable
				Else
					Set oDesc = Description.Create()
					oDesc("text").Value = sSAPTable
				End if
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				iRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).RowCount
				bVal = False
				iCount = 0
				Do while bVal = False
				For iRow = 1 to iRowCount
				If (SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).GetCellData (iRow,sColumn) <> "") then  
				iCount =iRow
				Else 
				bVal =True        
				Exit do 
				End if 
				Next
				Loop
				GetSAPTableRowsPopulated = iCount
				If Err.Description = "" then
				UpdateExecutionReport micPass, "GetSAPTableRowsPopulated", "Total Rows '"&GetSAPTableRowsPopulated&"' populated for SAPGUITable '"&sSAPTable&"' in MainWindow '"&sScreenName&"'"
				Else
				UpdateExecutionReport micFail, "GetSAPTableRowsPopulated", "SAPGUITable '"&sSAPTable&"' ### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			End Function

			'-----------------------------------------------------------------------------------------------------------------------------------------
			'**********************************************************************************************

			Function ActivateSAPTableCell ( sSAPTable, sRow, sColumn )
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	ActivateSAPTableCell ( sSAPTable, sRow, sColumn )
				' @parameter:	sRow	:	Row number
				' @parameter:	sColumn	:	Coloumn Name
				' @parameter:	sData	:   Value to be inserted 
				' @parameter:	sSAPTable  :	Name of the table
				' @notes	:   Inserts given data into specified Table cell
				' @END 
				On error resume next
				If InStr(sSAPTable,"_") > 0 then
					Set oDesc = Description.Create()
					oDesc("name").Value = sSAPTable
				Else
					Set oDesc = Description.Create()
					oDesc("text").Value = sSAPTable
				End if
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).ActivateCell sRow,sColumn
				If Err.Description = "" then
				UpdateExecutionReport micPass, "ActivateSAPTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"' has been activated"
				Else
				UpdateExecutionReport micFail, "ActivateSAPTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "ActivateSAPTableCell", "Table '"&sSAPTable&"' Cell of Row '"&sRow&"' & Column '"&sColumn&"'### Active window is Modal Window: '"&sScreenName&"'"
				End If
				Err.Clear
				On error goto 0		
			End Function

			'-----------------------------------------------------------------------------------------------------------------------------------------

			Function GetSAPTableEditableRow ( sSAPTable,sColumn)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	GetEditableRow ( sSAPTable,sColumn)
				' @returns	:	GetEditableRow	:	String
				' @parameter:	sWindow			:	Name of the Session
				' @parameter:   sSAPTable		:	Name of the Table  
				' @parameter:   sColumn			:	Name of the column in a Table
				' @notes	:	Returns the first editable row number of a given table. 
				' @END 

				If InStr(sSAPTable,"_") > 0 then
					Set oDesc = Description.Create()
					oDesc("name").Value = sSAPTable
				Else
					Set oDesc = Description.Create()
					oDesc("text").Value = sSAPTable
				End if
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				iRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).RowCount
				bVal = False
				iCount = 1
				Do while bVal = False
				For iRow = 1 to iRowCount
				If (SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).IsCellEditable(iRow,sColumn) and SAPGuiSession(oSesObjDescriptions).SAPGuiWindow(oObjectDescriptions).SAPGuiTable(oDesc).GetCellData (iRow,sColumn) = "") then  
				iCount =iRow
				bVal =True        
				Exit do 
				End If    
				Next
				Loop
				GetSAPTableEditableRow = iCount
				If Err.Description = "" then
				UpdateExecutionReport micPass, "GetSAPTableEditableRow", "Editable Row of Table '"&sSAPTable&"' & Column '"&sColumn&"' of MainWindow '"&sScreenName&"' is '"&GetSAPTableEditableRow&"'"
				Else
				UpdateExecutionReport micFail, "GetSAPTableEditableRow", "Table '"&sSAPTable&"' & Column '"&sColumn&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			End Function
			'-----------------------------------------------------------------------------------------------------------------------

			'--------------------------------------------------------------------------------------------------------------------------------- 	
			Function GetSAPTableValidRow ( sSAPTable,iRow)
				' @HELP
				' @class:  		Controls_SAP
				' @method:   	GetModalTableValidRow (sSAPTable, iRow)
				' @returns:   	GetModalTableValidRow: String
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPTable  :	Name of the Table
				' @parameter:   iRow       :	Row number
				' @notes:  
				' @END   

				If InStr(sSAPTable,"_") > 0 then
					Set oDesc = Description.Create()
					oDesc("name").Value = sSAPTable
				Else
					Set oDesc = Description.Create()
					oDesc("text").Value = sSAPTable
				End if
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				GetSAPTableValidRow = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable(oDesc).ValidRow(iRow)
				If Err.Description = "" then
				UpdateExecutionReport micPass, "GetSAPTableValidRow", "Validity of Row '"&iRow&"' of Table '"&sSAPTable&"' in MainWindow '"&sScreenName&"' is '"&GetSAPTableValidRow&"'"
				Else
				UpdateExecutionReport micFail, "GetSAPTableValidRow", "Row '"&iRow&"' of Table '"&sSAPTable&"' ### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If

			End Function

			Sub SetSAPTableDataIntoSpecificRow ( sSAPTable,aDataArray,iSpecificRow)
				' @HELP
				' @class	:   Controls_SAP
				' @method	:   SetItemTableData_IntoSpecificRow ( sSAPTable,aDataArray,iSpecificRow)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPTable	:	Name of the Table
				' @parameter:   aDataArray  :	Array of fields and field values
				' @parameter:   iSpecificRow:	Row number from which data to be set 
				' @notes	:	Sets data into SAPGuiTable into a specific Row        
				' @END     
				For iRow = 0 to UBound(aDataArray)
				SetSAPTableCellData  sSAPTable,iSpecificRow,aDataArray(iRow,3),aDataArray(iRow,4)
				Next     
			End Sub	

	'************************************************************************************************************************************************
	'-----------------------------------------------------------------------------------------------------------------------------------------
	Rem End Function
	'***********************************************************************************************************************************************************************************************

	Rem Function SAPCheckBoxRelatedFunctions()
			'----------------------------------------------------------------------------------------------------------------------------------------- 
			'It will return the selected property of the check box
			Function IsSAPCheckBoxSelected ( sSAPGuiCheck)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	IsSAPCheckBoxSelected ( sSAPGuiCheck)
				' @returns	:  	IsSAPCheckBoxSelected	:	ON/OFF
				' @parameter:	sWindow					:	Name of the Session
				' @parameter:   sSAPGuiCheck			:	Name of the CheckBox
				' @notes	:  	It will return the selected property of the check box
				' @END 
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				IsSAPCheckBoxSelected = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42").GetROProperty("Selected",1)
				If Err.Description = "" then
				UpdateExecutionReport micPass, "IsSAPCheckBoxSelected", "Status of SAPCheckbox '"&sSAPGuiCheckBox&"' of MainWindow '"&sScreenName&"' is '"&IsSAPCheckBoxSelected&"'"
				Else
				UpdateExecutionReport micFail, "IsSAPCheckBoxSelected", "SAPCheckbox '"&sSAPGuiCheckBox&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			End Function
			'Returns TRUE or FALSE ased on whether the check box existed or not using ID
			Function VerifySAPCheckBoxExistanceUsingIndex ( iIndex)
				' @HELP
				' @class:	  	Controls_SAP
				' @method:  	IsSAPCheckBoxExist_Index (iIndex)
				' @returns:		IsSAPCheckBoxExist_Index:	TRUE/FALSE
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:  	iIndex	:ID of the CheckBox
				' @notes:  		Returns TRUE or FALSE ased on whether the check box existed or not using ID
				' @END 
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				VerifySAPCheckBoxExistanceUsingIndex = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("index:="&iIndex,"guicomponenttype:=42").Exist
				If Err.Description = "" then
				UpdateExecutionReport micPass, "VerifySAPCheckBoxExistanceUsingIndex", "Existance of SAPGUICheckbox with index '"&iIndex&"' in MainWindow '"&sScreenName&"' is '"&VerifySAPCheckBoxExistanceUsingIndex&"'"
				Else
				UpdateExecutionReport micFail, "VerifySAPCheckBoxExistanceUsingIndex", "SAPGUICheckbox with index '"&iIndex&"' ### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			End Function
			'**************************************************************************************************************************************************************************

			Sub SelectSAPCheckBoxUsingIndex ( sSAPGuiCheck,sOption,sIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	: 	SelectSAPCheckBox_Index ( sSAPGuiCheck,sOption,sIndex)
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:   sSAPGUICheck:	Name of the Check box
				' @parameter:   sOption     :	Option (ON/OFF)
				' @parameter:   sIndex      :   Index of the Check ox on that screen
				' @notes	: 	Selects the Check box based on the option usig index property
				' @END 		
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42","index:="&sIndex).Exist then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42","index:="&sIndex).Set sOption
						'SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42","index:="&sIndex).SetFocus 
						UpdateExecutionReport micPass, "SelectSAPCheckBoxUsingIndex","SAPCheckBox with index '"&sIndex&"' selected in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPCheckBoxUsingIndex","SAPCheckBox with index '"&sIndex&"' not selected in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPCheckBoxUsingIndex","SAPCheckBox with index '"&sIndex&"' not selected as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPCheckBoxUsingIndex","SAPCheckBox with index '"&sIndex&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				on error goto 0
			End Sub
	''''''''**********************************************************************************************************************************************
			Sub SelectSAPCheckBox ( sSAPGuiCheck,sOption)
				' @HELP
				' @class	:	Controls_SAP
				' @method	: 	SelectSAPCheckBox ( sSAPGuiCheck,sOption)
				' @parameter:   sSAPGUICheck:	Name of the Check box
				' @parameter:   sOption     :	Option (ON/OFF)
				' @parameter:   sIndex      :   Index of the Check ox on that screen
				' @notes	: 	Selects the Check box based on the option usig index property
				' @END 		
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42").Exist then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42").Set sOption
						'SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheck,"guicomponenttype:=42").SetFocus 
						
						UpdateExecutionReport micPass, "SelectSAPCheckBox","SAPCheckBox   -" & sSAPGuiCheck & " set to " & sOption & " in Screen "&sScreenName
					ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("attachedtext:="&sSAPGuiCheck,"guicomponenttype:=42").Exist then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("attachedtext:="&sSAPGuiCheck,"guicomponenttype:=42").Set sOption
						'SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("attachedtext:="&sSAPGuiCheck,"guicomponenttype:=42").SetFocus 
						UpdateExecutionReport micPass, "SelectSAPCheckBox","SAPCheckBox   -" & sSAPGuiCheck & " set to " & sOption & " in Screen "&sScreenName
						
					Else 
					
						UpdateExecutionReport micFail, "SelectSAPCheckBox","SAPCheckBox   -" & sSAPGuiCheck & "not found in Screen "&sScreenName
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPCheckBox","SAPCheckBox - " &sSAPGuiCheck& " not selected as Active Screen is a Modal Window "&sScreenName
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPCheckBox","SAPCheckBox -"&sSAPGuiCheck&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				on error goto 0
			End Sub
	''''******************************************************************************************************************************************
			Sub SelectSAPCheckBoxUsingId ( sId,sOption)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	SelectSAPCheckBox_Id ( sId,sOption)
				' @parameter:	sWindow			:	Name of the Session
				' @parameter:   sSAPGuiCheck	:	Name of the CheckBox
				' @parameter:   sOption         :	Option (ON/OFF)
				' @notes	:	Selects a Checkox option using ID property to identify the checkbox
				' @END
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
						sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
						If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("guicomponenttype:=42","id:="&sId).Exist then
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("guicomponenttype:=42","id:="&sId).Set sOption
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("guicomponenttype:=42","id:="&sId).SetFocus
							UpdateExecutionReport micPass, "SelectSAPCheckBoxUsingId","SAPCheckBox with ID '"&sId&"' selected in Screen '"&sScreenName&"'"
						Else
							UpdateExecutionReport micFail, "SelectSAPCheckBoxUsingId","SAPCheckBox with ID '"&sId&"' not selected in Screen '"&sScreenName&"'"
						End If
					ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
						sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
						UpdateExecutionReport micFail, "SelectSAPCheckBoxUsingId","SAPCheckBox with ID '"&sId&"' not selected as Active Screen is a Modal Window '"&sScreenName&"'"
					End If
					If err.description <> "" Then
						UpdateExecutionReport micFail, "SelectSAPCheckBoxUsingId","SAPCheckBox with ID '"&sId&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
					End If
					Err.Clear
					on error goto 0
			End Sub 
	''''''*********************************************************************************************************************************************************		
			Function VerifySAPCheckBoxExistance (sSAPGuiCheckBox)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	IsSAPCheckBoxExist ( sSAPGuiCheckBox)
				' @returns	:  	IsSAPCheckBoxExist	:	TRUE/FALSE
				' @parameter:	sWindow				:	Name of the Session
				' @parameter:   sSAPGuiCheckBox		:	Name of the CheckBox
				' @notes	:  	Return True or False whether the CheckBox is Checked or not
				' @END 
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				VerifySAPCheckBoxExistance = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("name:="&sSAPGuiCheckBox,"guicomponenttype:=42").Exist
				If Err.Description = "" then
				UpdateExecutionReport micPass, "VerifySAPCheckBoxExistance", "Existance of SAPCheckbox '"&sSAPGuiCheckBox&"' of MainWindow '"&sScreenName&"' is '"&VerifySAPCheckBoxExistance&"'"
				Else
				UpdateExecutionReport micFail, "VerifySAPCheckBoxExistance", "SAPCheckbox '"&sSAPGuiCheckBox&"'### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			End Function

	Rem End Function
	'**************************************************************************************************************************************************************************
	'**************************************************************************************************************************************************************************
	Rem Function SAPGRidRelatedFunctions()
			Sub SelectSAPGridAllRows ()
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	SelectAllGridRow (sWindow)
				' @parameter:	sWindow	:	Name of the Session
				' @notes	:	Selects all row of a Grid
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").SelectAll ()
						UpdateExecutionReport micPass, "SelectSAPGridAllRows","All rows of a SAP Grid selected in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPGridAllRows","All rows of a SAP Grid not selected in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPGridAllRows","All rows of a SAP Grid not selected in SAP Main Window as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPGridAllRows","SAP GUI Grid ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
			End Sub
			
			Sub SelectSAPGridAllRows_UsingIndex (iIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	SelectAllGridRow (sWindow)
				' @parameter:	sWindow	:	Name of the Session
				' @notes	:	Selects all row of a Grid
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).SelectAll ()
						UpdateExecutionReport micPass, "SelectSAPGridAllRows","All rows of a SAP Grid selected in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPGridAllRows","All rows of a SAP Grid not selected in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPGridAllRows","All rows of a SAP Grid not selected in SAP Main Window as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPGridAllRows","SAP GUI Grid ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
			End Sub

			'-----------------------------------------------------------------------------------------------------------------------------------------
			Sub SelectSAPGridAllRowsUsingIndex (iIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	SelectAllGridRow (sWindow)
				' @parameter:	sWindow	:	Name of the Session
				' @notes	:	Selects all row of a Grid
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).Exist then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).SelectAll ()
						UpdateExecutionReport micPass, "SelectSAPGridAllRows","All rows of a SAP Grid selected in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPGridAllRows","All rows of a SAP Grid not selected in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPGridAllRows","All rows of a SAP Grid not selected in SAP Main Window as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPGridAllRows","SAP GUI Grid ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
			End Sub

			'-----------------------------------------------------------------------------------------------------------------------------------------
			Function VerifySAPGridExistance( )
				' @HELP
				' @class:		Controls_SAP
				' @method:      SAPGuiTableExists(sSAPTable)
				' @returns:     True if container with specified Id exists and false if it does not
				' @parameter:	sWindow		:	Name of the Session
				' @parameter:	sId	: Id of the Table                       
				' @notes:          
				' @END             
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell").Exist Then
				VerifySAPGridExistance = True
				Else
				VerifySAPGridExistance = False
				End If
				If Err.Description = "" then
				UpdateExecutionReport micPass, "VerifySAPGridExistance", "Existance of SAPGUIGrid in MainWindow '"&sScreenName&"' is '"&VerifySAPGridExistance&"'"
				Else
				UpdateExecutionReport micFail, "VerifySAPGridExistance", "SAPGUIGrid ### Main Window: '"&sScreenName&"' ### Error Description : "&Err.Description
				End If
			End Function


			'**************************************************************************************************************************************************************************

			Sub SelectSAPGridCell ( iRow, sColumn)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	SelectSAPGridCell ( iRow, sColumn)
				' @parameter:   iRow		:	Row number
				' @parameter:   sColumn		:	Name of the Column
				' @notes	:	Selects cell of a SAPGUIGrid on SAP Main Window
				' @END
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
					sScreenName = oSAPObj.GetROProperty("text")
					If oSAPObj.SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						oSAPObj.SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").SelectCell iRow, sColumn
						UpdateExecutionReport micPass, "SelectSAPGridCell","Select action performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPGridCell","Select action couldnot be performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
					sScreenName = oSAPObj.GetROProperty("text")
					If oSAPObj.SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						oSAPObj.SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").SelectCell iRow, sColumn
						UpdateExecutionReport micPass, "SelectSAPGridCell","Select action performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPGridCell","Select action couldnot be performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPGridCell","SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub

			'**************************************************************************************************************************************************************************

			Sub SelectSAPModalGridCell ( iRow, sColumn)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	SelectSAPModalGridCell ( iRow, sColumn)
				' @parameter:   iRow		:	Row number
				' @parameter:   sColumn		:	Name of the Column
				' @notes	:	Selects cell of a SAPGUIGrid on SAP Main Window
				' @END
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").SelectCell iRow, sColumn
						UpdateExecutionReport micPass, "SelectSAPModalGridCell","Select action performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPModalGridCell","Select action couldnot be performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPModalGridCell","Select action couldnot be performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Main Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPGridCell","SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub

			'**************************************************************************************************************************************************************************
			Sub ActivateSAPGridCell ( iRow, sColumn)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	ActivateSAPGridCell ( iRow, sColumn)
				' @parameter:   sSAPTable   :	Name of the field
				' @parameter:   iRow        :	Row number
				' @parameter:   sColumn     :   Column Name
				' @notes	:	Double Clicks on the Specified cell in SAPGUIGrid on a SAP Main Window
				' @END   		
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").ActivateCell iRow, sColumn
						UpdateExecutionReport micPass, "ActivateSAPGridCell","Activate action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "ActivateSAPGridCell","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "ActivateSAPGridCell","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "ActivateSAPGridCell","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub 
			
			'*****************************************************************************************************************************************************************************
			
			

			'**************************************************************************************************************************************************************************
			
			Sub ActivateSAPModalGridCell ( iRow, sColumn)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	ActivateSAPModalGridCell ( iRow, sColumn)
				' @parameter:   sSAPTable   :	Name of the field
				' @parameter:   iRow        :	Row number
				' @parameter:   sColumn     :   Column Name
				' @notes	:	Double Clicks on the Specified cell in SAPGUIGrid on a SAP Main Window
				' @END   		
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").ActivateCell iRow, sColumn
						UpdateExecutionReport micPass, "ActivateSAPGridCell","Activate action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in ModalScreen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "ActivateSAPGridCell","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in ModalScreen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "ActivateSAPGridCell","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Main Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "ActivateSAPGridCell","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub 

			'**************************************************************************************************************************************************************************

			Sub ClickSAPGridCell ( iRow, sColumn)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	ClickSAPGridCell ( iRow, sColumn)
				' @parameter:   iRow	:	Row number
				' @parameter:   sColumn :   Column Name
				' @notes	:	Clicks on the Specified cell in the SAP GUI Grid on SAP Main Window
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
					sScreenName = oSAPObj.GetROProperty("text")
					If oSAPObj.SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						oSAPObj.SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").ClickCell iRow, sColumn
						UpdateExecutionReport micPass, "ClickSAPGridCell","Click action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "ClickSAPGridCell","Click action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
					sScreenName = oSAPObj.GetROProperty("text")
					If oSAPObj.SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						oSAPObj.SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").ClickCell iRow, sColumn
						UpdateExecutionReport micPass, "ClickSAPGridCell","Click action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "ClickSAPGridCell","Click action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "ClickSAPGridCell","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub 

			''**********************************************************************************************************************************************************************************

			Sub ClickSAPMOdalGridCell ( iRow, sColumn)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	ClickSAPMOdalGridCell ( iRow, sColumn)
				' @parameter:   iRow	:	Row number
				' @parameter:   sColumn :   Column Name
				' @notes	:	Clicks on the Specified cell in the SAP GUI Grid on SAP Main Window
				' @END 
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").ClickCell iRow, sColumn
						UpdateExecutionReport micPass, "ClickSAPMOdalGridCell","Click action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Modal SAP Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "ClickSAPMOdalGridCell","Click action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "ClickSAPMOdalGridCell","Click action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Main Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "ClickSAPMOdalGridCell","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub 

			''**********************************************************************************************************************************************************************************

			Sub ActivateSAPGridCell_UsingIndex ( iRow, sColumn,iIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	ActivateGridCell ( iRow, sColumn)
				' @parameter:   sSAPTable   :	Name of the field
				' @parameter:   iRow        :	Row number
				' @parameter:   sColumn     :   Column Name
				' @notes	:	Double Clicks on the Specified cell in SAPGUIGrid on a SAP Main Window
				' @END   		
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201","index:="&iIndex).Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201","index:="&iIndex).ActivateCell iRow, sColumn
						UpdateExecutionReport micPass, "ActivateSAPGridCell_UsingIndex","Activate action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "ActivateSAPGridCell_UsingIndex","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "ActivateSAPGridCell_UsingIndex","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "ActivateSAPGridCell_UsingIndex","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub 

			''************************************************************************************************************************************************************************************************************************************

			Sub SelectSAPGridCell_UsingIndex ( iRow, sColumn,iIndex)
				' @HELP
				' @class	:	Controls_SAP
				' @method	:  	SelectSAPGridCell ( iRow, sColumn)
				' @parameter:   iRow		:	Row number
				' @parameter:   sColumn		:	Name of the Column
				' @notes	:	Selects cell of a SAPGUIGrid on SAP Main Window
				' @END
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201","index:="&iIndex).Exist Then 
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201","index:="&iIndex).SelectCell iRow, sColumn
						UpdateExecutionReport micPass, "SelectSAPGridCell","Select action performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					Else
						UpdateExecutionReport micFail, "SelectSAPGridCell","Select action couldnot be performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
					End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					UpdateExecutionReport micFail, "SelectSAPGridCell","Select action couldnot be performed on SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Modal Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "SelectSAPGridCell","SAPGUIGrid of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
				End If
				Err.Clear
				On error goto 0
			End Sub
	Rem End Function
	
	'***************************************************************************************************************************************************************************************************************
		Sub  OpenPossibleEntriesForSAPGridCell ( iRow, sColumn)
					' @HELP
					' @class	:	Controls_SAP
					' @method	:  	ActivateGridCell ( iRow, sColumn)
					' @parameter:   sSAPTable   :	Name of the field
					' @parameter:   iRow        :	Row number
					' @parameter:   sColumn     :   Column Name
					' @notes	:	Double Clicks on the Specified cell in SAPGUIGrid on a SAP Main Window
					' @END   		
                    CaptureSAPScreenShot()
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
						sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
						If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").Exist Then 
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("name:=shell","type:=GuiShell", "guicomponenttype:=201").OpenPossibleEntries  iRow, sColumn
							UpdateExecutionReport micPass, "SAPOpenPossibleEntries","Activate action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
						Else
							UpdateExecutionReport micFail, "SAPOpenPossibleEntries","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
						End If
					ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
						sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
						UpdateExecutionReport micFail, "SAPOpenPossibleEntries","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Modal Window '"&sScreenName&"'"
					End If
					If err.description <> "" Then
						UpdateExecutionReport micFail, "SAPOpenPossibleEntries","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
					End If
					Err.Clear
					On error goto 0
					
		End Sub 

		Sub  OpenSAPGridCellPossibleEntries_UsingIndex ( iIndex, iRow, sColumn)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:  	ActivateGridCell ( iRow, sColumn)
			' @parameter:   sSAPTable   :	Name of the field
			' @parameter:   iRow        :	Row number
			' @parameter:   sColumn     :   Column Name
			' @notes	:	Double Clicks on the Specified cell in SAPGUIGrid on a SAP Main Window
			' @END
            CaptureSAPScreenShot()
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).Exist Then 
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("guicomponenttype:=201","name:=shell","index:="&iIndex).OpenPossibleEntries  iRow, sColumn
					UpdateExecutionReport micPass, "SAPOpenPossibleEntries","Activate action performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
				Else
					UpdateExecutionReport micFail, "SAPOpenPossibleEntries","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' in Screen '"&sScreenName&"'"
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "SAPOpenPossibleEntries","Activate action couldnot be performed on GridCell of Row '"&iRow&"' and Column '"&sColumn&"' as Active Screen is a Modal Window '"&sScreenName&"'"
			End If
			If err.description <> "" Then
				UpdateExecutionReport micFail, "SAPOpenPossibleEntries","GridCell of Row '"&iRow&"' and Column '"&sColumn&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0					
		End Sub 


	'**************************************************************************************************************************************************************************************************************************

		Function SelectSAPMenuItem ( sItem)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPMenuItem (sItem)
			' @parameter:   sItem	:	Name of the menu item
			' @notes	:	Selects the menu bar item	 
			' @END 
            CaptureSAPScreenShot()
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiMenubar("name:=mbar").Select sItem
			If err.description = "" Then
				UpdateExecutionReport micPass, "SelectSAPMenuItem","'"& sItem &"' selected from MenuBar in Screen '"&sScreenName&"'"
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sStatusText = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("text")
			sMsgType = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("messagetype")
			If sStatusText <> "" Then
			UpdateExecutionReport micPass,"SelectSAPMenuItem-StatusBar Info",sMsgType&":"&sStatusText
			End if
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sModalWindowName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micPass,"SelectSAPMenuItem-ModalWindowInfo","S:ModalWindow populated with text '"&sModalWindowName&"'"
			End If
			else
			UpdateExecutionReport micFail, "SelectSAPMenuItem","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPMenuItem", "MenuItem: '"& sItem &"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
		'***********************************************************************************************************************************************************************************************
		Function SetSAPOKCode (sTCode)
		         
			' @HELP
			' @class:	 	Controls_SAP
			' @method:	 	SetTCode(sTCode )
			' @parameter:   sTCode	:	Value of the TCode   
			' @notes:	 	Sets the TCode in the OKCode field
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			
			if mid(lcase(sTCode),1,2) <> "/n" then
				sTCode = "/n" & sTCode
			end if
			
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiOKCode("guicomponenttype:=35","type:=GuiOkCodeField").Set sTCode
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SendKey ENTER
			If err.description = "" Then
			UpdateExecutionReport micPass, "SetSAPTCode","Tcode '"&sTCode&"' has been set"
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sStatusText = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("text")
			sMsgType = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("messagetype")
			If sStatusText <> "" Then
			UpdateExecutionReport micPass,"SetSAPTCode-StatusBarInfo",sMsgType&":"&sStatusText
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sModalWindowName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micPass,"SetSAPTCode-ModalWindowInfo","S:ModalWindow populated with text '"&sModalWindowName&"'"				
			End If
			else
			UpdateExecutionReport micFail, "SetSAPTCode","Screen: "&sScreenName&" ###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPTCode", "TCode: '"&sTCode&"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function 
		'***********************************************************************************************************************************************************************************************
		'***********************************************************************************************************************************************************************************************
		Function SetSAPTCode (sTCode)
		         
			' @HELP
			' @class:	 	Controls_SAP
			' @method:	 	SetTCode(sTCode )
			' @parameter:   sTCode	:	Value of the TCode   
			' @notes:	 	Sets the TCode in the OKCode field
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			
			if mid(lcase(sTCode),1,2) <> "/n" then
				sTCode = "/n" & sTCode
			end if
			
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiOKCode("guicomponenttype:=35","type:=GuiOkCodeField").Set sTCode
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SendKey ENTER
			If err.description = "" Then
			UpdateExecutionReport micPass, "SetSAPTCode","Tcode '"&sTCode&"' has been set"
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sStatusText = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("text")
			sMsgType = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("messagetype")
			If sStatusText <> "" Then
			UpdateExecutionReport micPass,"SetSAPTCode-StatusBarInfo",sMsgType&":"&sStatusText
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sModalWindowName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micPass,"SetSAPTCode-ModalWindowInfo","S:ModalWindow populated with text '"&sModalWindowName&"'"				
			End If
			else
			UpdateExecutionReport micFail, "SetSAPTCode","Screen: "&sScreenName&" ###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SetSAPTCode", "TCode: '"&sTCode&"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function 
		'************************************************************************************************************************************************
		Function GetSAPOKCodeProperty (sProperty)
			' @HELP
			' @class:	 	Controls_SAP
			' @method:	 	SetTCode(sTCode )
			' @parameter:   sTCode	:	Value of the TCode   
			' @notes:	 	Sets the TCode in the OKCode field
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				GetSAPOKCodeProperty = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiOKCode("guicomponenttype:=35","type:=GuiOkCodeField").GetROProperty("text")
				If err.description = "" Then
					UpdateExecutionReport micPass, "GetSAPOKCodeProperty","SAPOKCode Property Name "&sProperty &" ### Property Value'"&GetSAPOKCodeProperty&"' has been captured"
				else
					UpdateExecutionReport micFail, "GetSAPOKCodeProperty","Screen: "&sScreenName&" ###Error Description:"&err.description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "GetSAPOKCodeProperty", "SAPOKCode Property Name "&sProperty &" ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
Function GetSAPStatusBarInfo (sProperty)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:	GetStatusBarInfo (sProperty)
	' @returns	:	GetStatusBarInfo:	String
	' @parameter:	sProperty  		:	Property value of the status bar info to be returned.
	' @notes	:	Returns the required part of the status bar message
	' @END	
	On error resume next
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").exist Then
		sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
		If (sProperty = "text") then
			StatusBarMsg = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("text")
			MsgType = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("MessageType")
			If (MsgType <> "") Then
				GetSAPStatusBarInfo = MsgType & ":" & StatusBarMsg
				If MsgType = "E" Then
					UpdateExecutionReport micWarning,"GetSAPStatusBarInfo","Error Status Message displayed"
				End If
			Else
				GetSAPStatusBarInfo = "W:No Status Bar Message" 
			End If		
		Else 
			GetSAPStatusBarInfo = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty(sProperty)	
		End If
	Else
		sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
		UpdateExecutionReport micFail, "GetSAPStatusBarInfo", "SAP Modal Window Exists"
	End If
	If Err.Description = "" then
		UpdateExecutionReport micPass, "GetSAPStatusBarInfo", "Value '"& GetSAPStatusBarInfo &"' found for Property '"&sProperty&"' of Statusbar in Screen '"&sScreenName&"'"
	Else
		UpdateExecutionReport micFail, "GetSAPStatusBarInfo", "Value not found for Property'"&sProperty&"' of Statusbar in Screen '"&sScreenName&"'"
	End If
	Err.Clear
	On Error Goto 0	
	CaptureSAPScreenShot()
End Function
'***********************************************************************************************************************************************************************************************
Function PressSAPToolBarButton( sGuiCompType, iIndex, sButton)
	' @HELP
	' @class	:	Controls_SAP
	' @method	:  	PressSAPToolBarButton(sGuiCompType, iIndex, sButton)
	' @parameter:	sGuiCompType	:	Gui Component Type
	' @parameter:   iIndex	:	Index number of the Toolbar
	' @parameter:   sButton :	Name of the button	
	' @notes	:  	To press the button on a Tool bar
	'@Modified:-    Mahesh------01/2017
	' @END 
	On error resume next
	CaptureSAPScreenShot()
	If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
        Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
		sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiToolBar("guicomponenttype:="&sGuiCompType ,"index:="&iIndex).PressButton sButton
		UpdateExecutionReport micPass, "PressSAPToolBarButton","ScreenName = "&sScreenName&"### toolbar Button '"& sButton &"' pressed"
	Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
        Set oSAPObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
		sScreenName = oSAPObj.GetROProperty("text")
		oSAPObj.SAPGuiToolBar("guicomponenttype:="& sGuiCompType,"index:="&iIndex).PressButton sButton
		UpdateExecutionReport micPass, "PressSAPToolBarButton","ScreenName = "&sScreenName&"### toolbar Button '"& sButton &"' pressed"
	End if
	If err.description <> "" Then
	UpdateExecutionReport micFail, "PressSAPToolBarButton","ScreenName = "&sScreenName&"### toolbar Button = "& sButton &" ###Error Description:"&err.description
	End If
	Err.Clear
	On error goto 0		
End Function
'***********************************************************************************************************************************************************************************************
		Function ClickSAPToolBarButtonAndSelectMenuItem( iIndex,sContextButton,sMenuItem)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:  	ClickToolBarButton( iIndex,sContextButton,sMenuItem)
			' @parameter:	sWindow			:	Name of the Session
			' @parameter:   iIndex			:	Index number of the tool bar
			' @parameter:   sContextButton  :	Name of the Button
			' @parameter:   sMenuItem       :	Name of the menu item
			' @notes	:  	Clicks button on Tool bar
			' @END
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiToolbar("guicomponenttype:=202","index:="&iIndex).PressContextButton sContextButton
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiToolbar("guicomponenttype:=202","index:="&iIndex).SelectMenuItem sMenuItem
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			If err.description = "" Then
			UpdateExecutionReport micPass, "ClickToolBarButtonAndSelectMenuItem","ScreenName = "&sScreenName&"; toolbar Button '"& sButton &"' pressed and MenuItem '"&sMenuItem&"' selected"
			else
			UpdateExecutionReport micFail, "ClickToolBarButtonAndSelectMenuItem","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "ClickToolBarButtonAndSelectMenuItem", "ToolBar Index: '"&iIndex&"'###Toolbar Button '"& sButton &"' ### MenuItem '"&sMenuItem&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
		'**************************************************************************************************************************************************************************
		Function ClickSAPToolBarButtonAndSelectMenuItemById( iIndex,iGUICompType,sContextButton,sMenuItem)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:  	ClickToolBarButton( iIndex,sContextButton,sMenuItem)
			' @parameter:	sWindow			:	Name of the Session
			' @parameter:   iIndex			:	Index number of the tool bar
			' @parameter:   sContextButton  :	Name of the Button
			' @parameter:   sMenuItem       :	Name of the menu item
			' @notes	:  	Clicks button on Tool bar
			' @END
			On error resume next
			err.Clear
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then			
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiToolbar("guicomponenttype:="&iGUICompType,"index:="&iIndex).PressContextButton sContextButton
				err.Clear
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiToolbar("guicomponenttype:="&iGUICompType,"index:="&iIndex).SelectMenuItemById sMenuItem
				If err.description = "" Then
					UpdateExecutionReport micPass, "ClickSAPToolBarButtonAndSelectMenuItemById","ScreenName = "&sScreenName&"; toolbar Button '"& sButton &"' pressed and MenuItem '"&sMenuItem&"' selected"
				else
					UpdateExecutionReport micFail, "ClickSAPToolBarButtonAndSelectMenuItemById","Screen:"&sScreenName&"###Error Description:"&err.description
				End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "ClickToolBarButtonAndSelectMenuItem", "ToolBar Index: '"&iIndex&"'###Toolbar Button '"& sButton &"' ### MenuItem '"&sMenuItem&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0		
		End Function
		'**************************************************************************************************************************************************************************

		Sub SelectSAPTreeNodeAndOpenContextMenu( iIndex,sNode, sMenuItem)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:  	SelectSAPTreeNodeAndOpenContextMenu( iIndex,sNode, sMenuItem)
			' @parameter:   iIndex		:	Index of the context menu	
			' @parameter:   sNode       :	Name of the node	
			' @parameter:   sMenuItem   :	Name of the menu item	
			' @notes	:  	Clicks on a Menuitem by id of a node of a context menu
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","index:="&iIndex).OpenNodeContextMenu sNode
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","index:="&iIndex).SelectMenuItemById sMenuItem
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			If err.description = "" Then
			UpdateExecutionReport micPass, "SelectSAPTreeNodeAndOpenContextMenu","ScreenName = "&sScreenName&"; Node '"& sNode &"' & MenuItem '"& sMenuItem &"'selected in Tree"
			else
			UpdateExecutionReport micFail, "SelectSAPTreeNodeAndOpenContextMenu","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPTreeNodeAndOpenContextMenu", "Tree Index: '"&iIndex&"'### Node '"& sNode &"' ### MenuItem '"&sMenuItem&"'### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Sub
		'**************************************************************************************************************************************************************************
		Sub SelectSAPTreeNode( iIndex,sNode)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:  	SelectSAPTreeNode( iIndex,sNode)
			' @parameter:   iIndex		:	Index of the table tree control	
			' @parameter:   sNode		:	Item to select
			' @notes	:  	Selects a item from table tree control
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","index:="&iIndex).SelectNode sNode
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			If err.description = "" Then
			UpdateExecutionReport micPass, "SelectSAPTreeNode","ScreenName = "&sScreenName&"; Node '"& sNode &"' selected in Tree"
			else
			UpdateExecutionReport micFail, "SelectSAPTreeNode","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "SelectSAPTreeNode", "Tree Index: '"&iIndex&"'### ParentItem '"& sNode &"' ### Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Sub
	'**************************************************************************************************************************************************************************	
		Function ClickSAPEnter ()
			' @HELP
			' @class:	 	Controls_SAP
			' @method:	 	ClickSAPEnter( )
			' @notes:	 	Clicks Enter on any normal window
			' @END 
            CaptureSAPScreenShot()
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SendKey ENTER
				If err.description = "" Then
					UpdateExecutionReport micPass, "ClickSAPEnter","'ENTER' key pressed in Screen '"&sScreenName&"'"
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
						sStatusText = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("text")
						sMsgType = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("messagetype")
						If sStatusText <> "" and sMsgType = "S" Then
							UpdateExecutionReport micDone,"ClickSAPEnter-StatusBarInfo",sMsgType&":"&sStatusText
							ElseIf 	sStatusText <> "" and sMsgType <> "E" Then
							UpdateExecutionReport micDone,"ClickSAPEnter-StatusBarInfo",sMsgType&":"&sStatusText
							ElseIf 	sStatusText <> "" and sMsgType = "E" Then
							ClickEnter()
						End if
						ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
						sModalWindowName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
						UpdateExecutionReport micPass,"ClickSAPEnter-ModalWindow Info","S:ModalWindow populated with text '"&sModalWindowName&"'"
					End If
				Else
					UpdateExecutionReport micFail, "ClickSAPEnter","Screen:"&sScreenName&"###Error Description:"&err.description
				End If
				ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
				UpdateExecutionReport micFail, "ClickSAPEnter", "Active window is Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
		'***********************************************************************************************************************************************************************************************
		'***********************************************************************************************************************************************************************************************
		Function ClickSAPSaveButton()
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	ClickSAPSaveButton( )
			' @notes	:	Clicks the  tool bar  button on main window.	 	
			' @END
            CaptureSAPScreenShot()
            sButton = "btn[11]"
			sButton = Replace(sButton,"[","\[",1,1)

			sButton = Replace(sButton,"]","\]",1,1)
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("type:=GuiButton","name:="&sButton).Exist Then 
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("type:=GuiButton","name:="&sButton).click
				UpdateExecutionReport micPass, "ClickSAPSaveButton","Click action performed on Button 'SAVE' in Screen '"&sScreenName&"'"
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("type:=GuiButton","tooltip:="&sButton).Exist Then 
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("type:=GuiButton","tooltip:="&sButton).click
				UpdateExecutionReport micPass, "ClickSAPSaveButton","Click action performed on Button 'SAVE' in Screen '"&sScreenName&"'"
			else
				UpdateExecutionReport micFail, "ClickSAPSaveButton","Click action couldnot be performed on Button 'SAVE' in Screen '"&sScreenName&"'"
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "ClickSAPSaveButton","Click action couldnot be performed on Button '"&sButton&"' as Active Screen is a Modal Window '"&sScreenName&"'"
			End If
			If err.description <> "" Then
			UpdateExecutionReport micFail, "ClickSAPSaveButton","Button:'SAVE'' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0
		End Function

	'****************************************************************************************************************************
		Function ClickSAPToolBarButton( sButton)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	ClickSAPToolBarButton( )
			' @notes	:	Clicks the  tool bar  button on main window.	 	
			' @END
            CaptureSAPScreenShot()
			sButton = Replace(sButton,"[","\[",1,1)

			sButton = Replace(sButton,"]","\]",1,1)
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("type:=GuiButton","name:="&sButton).Exist Then 
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("type:=GuiButton","name:="&sButton).click
				UpdateExecutionReport micPass, "ClickSAPToolBarButton","Click action performed on Button '"&sButton&"' in Screen '"&sScreenName&"'"
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("type:=GuiButton","tooltip:="&sButton).Exist Then 
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("type:=GuiButton","tooltip:="&sButton).click
				UpdateExecutionReport micPass, "ClickSAPToolBarButton","Click action performed on Button '"&sButton&"' in Screen '"&sScreenName&"'"
			else
				UpdateExecutionReport micFail, "ClickSAPToolBarButton","Click action couldnot be performed on Button '"&sButton&"' in Screen '"&sScreenName&"'"
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			UpdateExecutionReport micFail, "ClickSAPToolBarButton","Click action couldnot be performed on Button '"&sButton&"' as Active Screen is a Modal Window '"&sScreenName&"'"
			End If
			If err.description <> "" Then
			UpdateExecutionReport micFail, "ClickSAPToolBarButton","Button:"&sButton&"' ### Screen:'"&sScreenName&"' ### Error Description:"&err.description
			End If
			Err.Clear
			On error goto 0
		End Function

	'****************************************************************************************************************************
		Function ClickSAPModalEnter ()
			' @HELP
			' @class:	 	Controls_SAP
			' @method:	 	ClickSAPModalEnter( )
			' @parameter:	sWindow		:	Name of the Session
			' @notes:	 	Clicks Enter on any Modal Window
			' @END 
            CaptureSAPScreenShot()
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SendKey ENTER
			If err.description = "" Then
			UpdateExecutionReport micPass, "ClickSAPModalEnter","'ENTER' key pressed in Screen '"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "ClickSAPModalEnter","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "ClickSAPModalEnter", "Active window is Main Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
		'**********************************************************************************************************************************************************
	
		
		''''*************************************
Function SendSAPkey (sKey)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SendSAPKey ( sKey)
			' @parameter:   sKey	:	Function Key/Combination of Keys
			' @notes	:	Presses any Function key or Combination of keys	 
			' @END 
            CaptureSAPScreenShot()
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
                Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
				sScreenName =oSAPObj.GetROProperty("text")
				oSAPObj.Activate
				oSAPObj.SendKey sKey
			Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
                Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
				sScreenName =oSAPObj.GetROProperty("text")
				oSAPObj.SendKey sKey
			End If
				If err.description = "" Then
					UpdateExecutionReport micPass, "SendSAPKey","Key '"& sKey &"' pressed in Screen '"&sScreenName&"'"
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
						sStatusText = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("text")
						sMsgType = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("name:=sbar","guicomponenttype:=103","type:=GuiStatusbar").GetROProperty("messagetype")
						If sStatusText <> "" and sMsgType = "S" Then
							UpdateExecutionReport micDone,"SendSAPKey-StatusBarInfo",sMsgType&":"&sStatusText
						ElseIf sStatusText <> "" and sMsgType <> "E" Then
							UpdateExecutionReport micDone,"SendSAPKey-StatusBarInfo",sMsgType&":"&sStatusText
						ElseIf sStatusText <> "" and sMsgType = "E" Then
							ClickEnter()
						End if
					ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
						sModalWindowName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
						UpdateExecutionReport micPass,"SendSAPKey-ModalWindowInfo","S:ModalWindow populated with text '"&sModalWindowName&"'"
					End If
				else
				UpdateExecutionReport micFail, "SendSAPFunctionKey","Screen:"&sScreenName&"###Error Description:"&err.description
				End If
			Err.Clear
			On error goto 0	
		End Function
		'**************************************************************************************************************************************************************************
		Function SendSAPModalKey ( sKey)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SendSAPModalKey ( sKey)
			' @parameter:	sWindow	:	Name of the Session
			' @parameter:   sKey	:	Function Key/Combination of Keys
			' @notes	:	Presses any Function key or Combination of keys	 
			' @END 
            CaptureSAPScreenShot()
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SendKey sKey
			If err.description = "" Then
			UpdateExecutionReport micPass, "SendSAPModalKey","Key '"& sKey &"' pressed in Screen '"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "SendSAPModalKey","Screen:"&sScreenName&"###Error Description:"&err.description
			End If
			ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			UpdateExecutionReport micFail, "SendSAPModalKey", "Active window is not a Modal Window: '"&sScreenName&"'"
			End If
			Err.Clear
			On error goto 0	
		End Function
		'**************************************************************************************************************************************************************************
		Sub CloseSAPSession ()
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SAPCloseSession ()
			' @notes	:	closes SAP session
			' @END 
			SAPGuiUtil.CloseConnections
			UpdateExecutionReport micPass, "CloseSAPSession","All open SAP Sessions are closed"
		End Sub
	
		'**************************************************************************************************************************************************************************
		Sub SAPLogon (aLoginData,iStepNo)
			' @HELP
			' @class:		Controls_SAP
			' @method:	 	SAPLogon (aLoginData, iStepNo)
			' @parameter:   aLoginData	:	Array of Login details
			' @parameter:   iStepNo     :	Step number of the component in the Plan
			' @notes:		Logs into SAP system	 
			' @END	    	      
			SAPGuiUtil.AutoLogon aLoginData(iStepNo-1,5), aLoginData(iStepNo-1,1), aLoginData(iStepNo-1,3), aLoginData(iStepNo-1,4), aLoginData(iStepNo-1,2)
			if err.description = "" then
				UpdateExecutionReport micPass, "SAPLogon","Login into "&aLoginData(iStepNo-1,5)&" successful with UserId '"&aLoginData(iStepNo-1,3)&"'"
			else
				UpdateExecutionReport micPass, "SAPLogon","Login into "&aLoginData(iStepNo-1,5)&" successful with UserId '"&aLoginData(iStepNo-1,3)&"'"
			end if
		End Sub
		'**************************************************************************************************************************************************************************
		Sub SAPLogOut (aLoginData,iStepNo)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SAPLogout (aLoginData, iStepNo)
			' @parameter:   aLoginData  :	Array of Login details
			' @parameter:   iStepNo     :	Step number of the component in the Plan
			' @notes	: 	Logs out of the SAP system
			' @END 
			If  (aLoginData(iStepNo,3)  =  aLoginData(iStepNo-1,3) ) = False Then
			SAPGuiUtil.CloseConnections
			SAPLogon aLoginData,iStepNo + 1
			End If	
		End Sub
		''''''******************************************************************************************************************************************************************
		'***********************************************************************************************************************************************************************************************
		Function VerifySAPModalWindowExistance ( )
			' @HELP
			' @class:		Controls_SAP
			' @method:		VerifySAPModalWindowExistance (sSAPTabStrip)
			' @notes:		returns modal window existance  
			' @END 
			On error resume next
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
			VerifySAPModalWindowExistance = True
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			'UpdateExecutionReport micPass, "VerifySAPModalWindowExistance", "Active window is Modal Window: '"&sScreenName&"'"
			Reporter.ReportEvent micPass, "VerifySAPModalWindowExistance", "Active window is Modal Window: '"&sScreenName&"'"
			else
			VerifySAPModalWindowExistance = False
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			'UpdateExecutionReport micWarning, "VerifySAPModalWindowExistance", "Active window is Main Window: '"&sScreenName&"'"
			End If
			If err.description <> "" Then
			Reporter.ReportEvent micFail, "VerifySAPModalWindowExistance","Error Description: " & err.description
			End If
			Err.Clear
			On error goto 0	
		End Function
		'***********************************************************************************************************************************************************************************************
		Function GetSAPWindowProperty (sProperty)
			' @HELP
			' @class:		Controls_SAP
			' @method:	 	GetWindowProperty (sProperty)
			' @returns:	 	GetWindowProperty : 
			' @parameter:   sProperty  : 	Name of the Property
			' @notes:	 
			' @END 
			On error resume next
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			GetSAPWindowProperty = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty(sProperty)
			If err.description = "" Then
			UpdateExecutionReport micPass,"GetSAPWindowProperty","'"& GetSAPWindowProperty & "' found for Property '"&sProperty&"' of SAP MainWindow '"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPWindowProperty","Error Description:"&err.description
			End If
		End Function
		'**********************************************************************************************
		
		Function GetSAPSessionProperty (sProperty)
			' @HELP
			' @class:		Controls_SAP
			' @method:	 	GetSessionProperty (sProperty)
			' @returns:	 	GetWindowProperty : 
			' @parameter:   sProperty  : 	Name of the Property
			' @notes:	 
			' @END 
			'On error resume next
		
			GetSAPSessionProperty = SAPGuiSession("guicomponenttype:=12").GetROProperty(sProperty)
		
		End Function
		
		'**********************************************************************************************************************************
		Function GetSAPModalWindowProperty (sProperty)
			' @HELP
			' @class:		Controls_SAP
			' @method:	 	GetWindowProperty (sProperty)
			' @returns:	 	GetWindowProperty : 
			' @parameter:   sProperty  : 	Name of the Property
			' @notes:	 
			' @END 
			On error resume next
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
			GetSAPModalWindowProperty = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty(sProperty)
			If err.description = "" Then
			UpdateExecutionReport micPass,"GetSAPModalWindowProperty","'"& GetSAPWindowProperty & "' found for Property '"&sProperty&"' of SAP MainWindow '"&sScreenName&"'"
			else
			UpdateExecutionReport micFail, "GetSAPModalWindowProperty","Error Description:"&err.description
			End If
		End Function	 	
		'********************************************************************************************************************************************************************************************

		Function GetNumberFromText(text)
			ss=Split(text," ")
			For i=0 To UBound(ss)
				If IsNumeric(ss(i))  Then
				 no=ss(i)
				 Exit For
				End If 
			Next
			GetNumberFromText=no
		End Function 

		'********************************************************************************************************************************************************************************************

		Function VerifyText (sActual, sExpected)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	VerifyText (sActual, sExpected)
			' @returns	:	VerifyText : String
			' @parameter:   sActual    : Actual value to be compared with expected value 
			' @parameter:   sExpected  : Expected value to be compared with actual value
			' @notes	:	This method verifies the actual text with expected text and returns the result (string comparision - returns 0 if comparable and 1 if NOT COMPARABLE)
			' @END
			If strComp(sActual,sExpected,1) = 0 Then
			VerifyText = 0
			Else 
			VerifyText = 1 
			End If 
		End Function
		'-----------------------------------------------------------------------------------------------------------------------------------------	
		'Sets the screen to the Easy Access screen (from any screen)
			Sub SetToSAPEasyAccess ()
				' @HELP
				' @class:		Controls_SAP
				' @method:	 	SetToEasyAccess( )
				' @notes:		Sets the session to Easy Access screen 	 
				' @END 
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiOKCode("guicomponenttype:=35","type:=GuiOkCodeField").Set "/n"
				SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SendKey ENTER
				If Err.Description = "" then
					UpdateExecutionReport micPass, "SetToSAPEasyAccess", "SAP Main Window set to SAPEasyAccess"
				Else
					UpdateExecutionReport micFail, "SetToSAPEasyAccess", "Error Description : "&Err.Description
				End If
			End Sub

		'**************************************************************************************************************************************************************************
			Sub CloseSAPModalWindow ()
				' @HELP
				' @class	:	Controls_SAP
				' @method	:	CloseModalWindow ()
				' @notes	:	Closes modal window
				' @END 
				on error resume next
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").GetROProperty("text")
					SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Close
					UpdateExecutionReport micPass, "CloseSAPModalWindow","SAPModalWindow of Text '"&sScreenName&"' has been closed"
				Else
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					UpdateExecutionReport micFail, "CloseSAPModalWindow","No SAPModalWindow Exists, Active Window is SAP Main Window '"&sScreenName&"'"
				End If
				If err.description <> "" Then
					UpdateExecutionReport micFail, "CloseSAPModalWindow","SAPModalWindow ### Error Description:"&err.description
				End If
				Err.Clear
				on error goto 0
			End Sub
		'**************************************************************************************************************************************************************************

			Function ActivateSAPTreeNode (sNode,iIndex )
					' @HELP
					' @class	:	Controls_SAP
					' @method	:	ActivateSAPNode (  sNode,iIndex )
					' @parameter:	sNode	:	Name of the node 
					' @parameter:	iIndex	:	Index number
					' @notes	:   Activate The Node
					' @END 
					
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","index:="&iIndex).OpenNodeContextMenu sNode
						SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","index:="&iIndex).ActivateNode  sNode
					End if 
			End Function
			
	    '*******************************************************************************************************************************************************************
		
		Function SelectSAPTreeCheckBox  (sNode,CheckBoxName ,sOption )
					' @HELP
					' @class	:	Controls_SAP
					' @method	:	 SelectSAPTreeCheckBox  (sNode,CheckBoxName ,sOption )
					' @parameter:	sNode	:	Name of the node 
					' @parameter:	CheckBoxName	:	Name of the Checkbox (Recored)
					' @parameter :  sOption : ON / OF
					' @notes	:  Select the Check box  
					' @END 	
					On Error Resume Next
					If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
						sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
						 If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200").Exist then
							SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200").Set sNode,CheckBoxName ,sOption
							UpdateExecutionReport micPass,"SelectSAPTreeCheckBox","CheckBoxName = "&CheckBoxName&" set with Option = "&sOption&" in Node = "&sNode
                        End if
					Else
						UpdateExecutionReport micFail,"SelectSAPTreeCheckBox","CheckBoxName = "&CheckBoxName&" not set with Option = "&sOption&" in Node = "&sNode
					End If
					On Error Goto 0
		    End Function
			
		''''''''************************** &&&&&&&&&&&&&&&&&& Add new functions here &&&&&&&&&&&&&&&&&&&&&&&&&&&&&************************************
			Sub  SelectSAPTabStripUsingChildObject (sSAPTabStrip,sTab)
				' @HELP
				' @class:		Controls
				' @method:		SelectSAPTabStripUsingChildObject (sSAPTabStrip,sTab)
				' @parameter:   sSAPTabStrip	:	Name of the TabStrip  
				' @parameter:   sTab            :	Name of the Tab to be selected
				' @notes:		Selects the required Tab of the specified TabStrip	 
				' @END 
                CaptureSAPScreenShot()
				On Error Resume Next
				Reporter.Filter = rfDisableAll
				iSelectSAPTabStripUsingChildObject = 0
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then 
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
					Set allobj =SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").ChildObjects()
					For chobjcnt = 0 to allobj.Count-1
						sName = allobj(chobjcnt).getROProperty("name")
						If  sName = sSAPTabStrip  Then
							allobj(chobjcnt).Select sTab
							iSelectSAPTabStripUsingChildObject = 1
							Exit For
						End If 
					Next
				End If
				Reporter.Filter = rfEnableAll
				If iSelectSAPTabStripUsingChildObject = 1 Then
					UpdateExecutionReport micPass,"SelectSAPTabStripUsingChildObject","Tab "&sTab&" selected from Tabstrip "&sSAPTabStrip
				Else
					UpdateExecutionReport micFail,"SelectSAPTabStripUsingChildObject","Tab "&sTab&" couldnot be selected from Tabstrip "&sSAPTabStrip
				End If
				On Error Goto 0
        End Sub

'*****************************************************************************************************************************************************

		Function GetSAPTableTreeSelectedProperty (iIndex,sProperty)
			' @HELP
			' @class:		Controls
			' @method:   	GetTableTreeSelectedProperty(,iIndex,sProperty)
			' @returns:  	GetTableTreeSelectedProperty	: String/Integer
			' @parameter:	sWindow							: Name of the Window
			' @parameter:   iIndex							: Index of the Table tree control
			' @parameter:   sProperty						: Property to be retreived
			' @notes:   	Returs the Table Tree Property selected 
			' @END
			If SAPGuiSession("guicomponenttype:=12").SAPGuiTree("guicomponenttype:=200","index:="&iIndex).Exist Then
				GetSAPTableTreeSelectedProperty = SAPGuiSession("guicomponenttype:=12").SAPGuiTree("guicomponenttype:=200","index:="&iIndex).GetROProperty(sProperty)
			End If
		End Function

'*********************************************************************************************************************************************************************

		Function GetSAPLableContent_id (iId)	
			' @HELP
			' @class	:  Controls
			' @method	:  GetSAPIDValue()
			' @returns	:  GetSAPEditValue  : To get the id value of the window 
			' @notes	:  To get the id value of the window
			' @END  
				If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then				
					sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")					
					GetSAPLableContent_id = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("id:="&iId,"guicomponenttype:=30").GetROProperty ("content")
				End If
		End Function

'************************************************************************************************************************************************************************************
'Commented by Dinakar
			Rem Function IsSAPLabelContentExists_Id (sSAPRelativeID)
			Rem ' @HELP
			Rem ' @class	:  	Controls
			Rem ' @method	:  	IsSAPLabelContentExists_Id (sSAPRelativeID)
			Rem ' @returns	:  	IsLabelContentExists_Id :	String
			Rem ' @parameter:   sSAPRelativeID  		:	Relative ID of the Label
			REM ' @notes	:  	Returns the Content of the Label
			REM ' @END 
			REM If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			Rem sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			REM If IsSAPLabelContentExists_Id = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("guicomponenttype:=30","type:=GuiLabel","relativeid:="&sSAPRelativeID).Exist(1) Then 	
			REM IsSAPLabelContentExists_Id = "True"
			REM Else  
			REM IsSAPLabelContentExists_Id = "False" 
			REM End If
			REM ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then        
			REM Set allobj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").ChildObjects()
			REM For chobjcnt = 0 to allobj.Count-1
			Rem iGuiComponentType =  allobj(chobjcnt).getROProperty("guicomponenttype")
			Rem iType = allobj(chobjcnt).getROProperty("type")
			Rem iRelativeId = allobj(chobjcnt).getROProperty("relativeid")
			Rem If allobj(chobjcnt).getROProperty("relativeid") = sSAPRelativeID Then
			REM IsSAPLabelContentExists_Id = "True"
			REM Exit for
			REM Else
			REM IsSAPLabelContentExists_Id = "False"
			REM End If 
			REM Next
			REM End If 
			Rem End Function
	
	'****************************************************************************************************************************************************************************************
	
		Function GetSAPLable_UsingRelativeid (Relativeid)
			' @HELP
			' @class	:  Controls
			' @method	:  GetSAPIDValue()
			' @returns	:  GetSAPEditValue  : To get the id value of the window 
			' @notes	:  To get the id value of the window
			' @END  
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
				GetSAPLable_UsingRelativeid = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiLabel("relativeid:="&Relativeid,"guicomponenttype:=30").GetROProperty ("content")
			End If
		End Function


   '*************************************************************************************************************************************************************************************************
     
		Function GetSAPIDValue()
			' @HELP
			' @class	:  Controls
			' @method	:  GetSAPIDValue()
			' @returns	:  GetSAPEditValue  : To get the id value of the window 
			' @notes	:  To get the id value of the window
			' @END  
			IDValueData = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("id")
			MyArray = Split(IDValueData, "/", -1, 1)
			sRequiredID    ="/"&MyArray(1)&"/"&MyArray(2)&"/"&MyArray(3)&"/"
			sRequiredID = replace(sRequiredID,"/","\/")
			sRequiredID = replace(sRequiredID,"[","\[")
			GetSAPIDValue = replace(sRequiredID,"]","\]")    
		End Function

'*************************************************************************************************************************************************************************************
     Function GetSAPTextAreaValue ( sName)
		GetSAPTextAreaValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTextArea("name:="&sName,"guicomponenttype:=203").getroproperty("value")
	 End Function
	 
	 Sub SetSAPTextArea (sName,sValue)
	  ' @HELP
	  ' @class	:  	Controls
	  ' @method	:  	SetSAPTextArea (sName,sValue)
	  ' @parameter:   sName   :	Tech Name of the field
	  ' @parameter:   sValue      :   Value to be set in to Text area field
	  ' @notes	:  	Sets data into a SAPGuiEdit Text Area
	  ' @END 
		   If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
			 SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTextArea("name:="&sName,"guicomponenttype:=203").Set sValue
		   End If                            
	  End Sub
	
   '***************************************************************************************************************************************************************
Sub UpdateExecutionReport(sResult,sExpeMessage,sActualMessage)
	sActualMessage = Replace(sActualMessage,"'","")
	Select Case sResult
		Case 0
			ReportWriter "Pass",sExpeMessage,sActualMessage,0
			REM oAccess.UpdateExecutionReportToDB sExpeMessage,sActualMessage,"Pass"
		Case 1
			ReportWriter "Fail",sExpeMessage,sActualMessage,0
			REM oAccess.UpdateExecutionReportToDB sExpeMessage,sActualMessage,"Fail"
		Case 3
			REM oAccess.UpdateExecutionReportToDB sExpeMessage,sActualMessage,"Warning"
	End Select
End Sub
			
		REM Sub UpdateExecutionReport(sResult,sExpeMessage,sActualMessage)
			REM sActualMessage = Replace(sActualMessage,"'","")
			
			REM Select Case sResult
				REM Case 0
					REM UpdateExecutionReportToDB sExpeMessage,sActualMessage,"Pass"
				REM Case 1
					REM REM iRunID          =   Environment.Value("RunId")
					REM REM sExecuteID		=	Environment.Value("ExecuteID")
					REM REM sRunID			=	Environment.Value("RunId")
					REM REM iDataSetNo		=	Environment.Value("DataSetNo")
					REM REM iTestCaseID		=	Environment.Value("TestCaseID")
					REM REM sResultPath		=	Environment.Value("ResultPath")
					REM REM sScreenShotPath	=	replace(sResultPath,"HTMLReports","ScreenShots")
					REM REM sNow			=	Now
					REM REM sNow			=	replace(sNow," ","")
					REM REM sNow			=	replace(sNow,"/","")
					REM REM sNow			=	replace(sNow,":","")
					REM REM sNow			=	replace(sNow,"AM","1")
					REM REM sNow			=	replace(sNow,"PM","2")
					REM REM sScreenShot		=	sScreenShotPath & iTestCaseID & "@" & iDataSetNo & "@" & sExecuteID & "@" & sRunID & "@" & sNow & ".png"
					REM REM on error resume next
					REM REM Reporter.Filter = rfDisableAll
					REM REM If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
						REM REM Set oSAPSessionObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
					REM REM ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
						REM REM Set oSAPSessionObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
					REM REM End If
					REM REM oSAPSessionObj.Activate
					REM REM oSAPSessionObj.Highlight
					REM REM sScreenName = oSAPSessionObj.GetROProperty("text")
					REM REM oSAPSessionObj.CaptureBitmap sSCreenShot,True
					REM REM Reporter.Filter = rfEnableAll
					REM REM ExecutionComponentTracker sExpeMessage,sActualMessage&" ###Begin###"&sSCreenShot&"###End###","Fail",Now
					REM REM Set oSAPSessionObj = Nothing
					REM on error goto 0					
				REM Case 3
					REM iRunID          =   Environment.Value("RunId")
					REM sExecuteID		=	Environment.Value("ExecuteID")
					REM sRunID			=	Environment.Value("RunId")
					REM iDataSetNo		=	Environment.Value("DataSetNo")
					REM iTestCaseID		=	Environment.Value("TestCaseID")
					REM sResultPath		=	Environment.Value("ResultPath")
					REM sScreenShotPath	=	replace(sResultPath,"HTMLReports","ScreenShots")
					REM sNow			=	Now
					REM sNow			=	replace(sNow," ","")
					REM sNow			=	replace(sNow,"/","")
					REM sNow			=	replace(sNow,":","")
					REM sNow			=	replace(sNow,"AM","1")
					REM sNow			=	replace(sNow,"PM","2")
					REM sScreenShot		=	sScreenShotPath & iTestCaseID & "@" & iDataSetNo & "@" & sExecuteID & "@" & sRunID & "@" & sNow & ".png"
					REM on error resume next
					REM Reporter.Filter = rfDisableAll
					REM If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
						REM Set oSAPSessionObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
					REM ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
						REM Set oSAPSessionObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
					REM End If
					REM oSAPSessionObj.Activate
					REM oSAPSessionObj.Highlight
					REM sScreenName = oSAPSessionObj.GetROProperty("text")
					REM oSAPSessionObj.CaptureBitmap sSCreenShot,True
					REM Reporter.Filter = rfEnableAll
					REM ExecutionComponentTracker sExpeMessage,sActualMessage&" ###Begin###"&sSCreenShot&"###End###","Warning",Now
					REM Set oSAPSessionObj = Nothing
					REM on error goto 0	
			REM End Select
		REM End Sub
		
		Sub ExecutionComponentTracker (sEffectaStep, sEffectaStepDescription, sEffectaStepResult, sEffectaTimeStamp)
			sXLFileName = Environment.Value("TempResultsFile")
			strSQLStatement = "INSERT INTO [EffectaActions$] (EffectaStep, EffectaStepDescription, EffectaStepResult, EffectaTimeStamp) VALUES ('"& sEffectaStep & "', '" & sEffectaStepDescription & "', '" & sEffectaStepResult & "', '" & sEffectaTimeStamp &"')"
			Const adOpenStatic		= 3
			Const adLockOptimistic	= 3
			Const adCmdText			= 1
			Set objExcelConnection	= CreateObject("ADODB.Connection")  
			objExcelConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sXLFileName & ";Extended Properties=""Excel 8.0;HDR=Yes;"";" 
			Set objCreateCommand	= CreateObject("ADODB.Command")  
			Set objCreateCommand.ActiveConnection = objExcelConnection  
			objCreateCommand.CommandText = strSQLStatement
			objCreateCommand.Execute , , adCmdText
			Set objCreateCommand = Nothing
			Set objExcelConnection = Nothing
		End Sub
			
'***************************************************************************************************************************************************************
Function ClickSAPWebElement( sFrame,sWebElement)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SelectSAPGuiTreeItem( iIndex,sPath,sItem)
			' @parameter:	sFrame		:	Name of the Frame
			' @parameter:   sWebElement		: innerhtml of WebElement					
			' @notes	:  	ClickSAPWebElement
			' @END
	On error resume next
	Err.Clear
			If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
              		Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
					sScreenName = oSAPObj.GetROProperty("text")
					oSAPObj.Page("micClass:=Page").Frame("name:="&sFrame).WebElement("innerhtml:="&sWebElement).Click
                    UpdateExecutionReport micPass, "ClickSAPWebElement","SAPWebElement '"& sWebElement &"' click in Screen '"&sScreenName&"'"
			Elseif SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
              		Set oSAPObj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
					sScreenName = oSAPObj.GetROProperty("text")
					oSAPObj.Page("micClass:=Page").Frame("name:="&sFrame).WebElement("innerhtml:="&sWebElement).Click
					UpdateExecutionReport micPass, "ClickSAPWebElement","SAPWebElement '"& sWebElement &"' click in Screen '"&sScreenName&"'"
			End if

			If Err.description <>"" Then
				UpdateExecutionReport micFail, "ClickSAPWebElement","SAPWebElement:"& sWebElement &"' ### Screen:"&sScreenName&" ### Error Description:"&err.description
			End if
	Err.Clear
	On error goto 0
End Function
'***************************************************************************************************************************************************************
		
	  Function SetSAPWebEdit ( sPagetitle,sFrameName,sWebEditName,sValue)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	SetSAPwebEdit ( sPagetitle,sFrameName,sWebEditName,sValue)
			' @parameter:   sPagetitle	:	Page Title of the Edit field   
			' @parameter:   sFrameName	:	Techiniclal Name Of the Frame
			' @parameter:   sWebEditName : Techinical Name of the  WebEdit
			' @parameter:   sValue		:	Value to be set
			' @notes	:	Sets the data into a SAPGuiEdit
			' @END
			On Error Resume Next
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Page("title:="&sPagetitle).Frame("name:="&sFrameName).WebEdit("name:="&sWebEditName).Set  sValue
			If err.description = "" Then
				UpdateExecutionReport micPass, "SetSAPWebEdit","Value "&sValue&" set in SAPwebEdit "& sWebEditName &" on SAP Window "&sScreenName 
			Else
				UpdateExecutionReport MicFail, "SetSAPWebEdit","Value "&sValue&" not set in SAPwebEdit "& sWebEditName &" ### SAP Window = "&sScreenName & " Error = " &err.description
			End If
			Err.Clear
			on error goto 0
		End Function
			
   '***************************************************************************************************************************************************************
	
		Function ClickSAPWebButton ( sPagetitle,sFrameName,sWebButton)
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	ClickSAPWebButton ( sPagetitle,sFrameName,sWebButton)
			' @parameter:   sPagetitle	:	 Title of the Page
			' @parameter:   sFrameName	:	Techiniclal Name Of the Frame
			' @parameter:   sWebEditName : Techinical Name of the  WebButton
			' @notes	:	Click The WebButton
			' @END
			On Error Resume Next
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Page("title:="&sPagetitle).Frame("name:="&sFrameName).WebButton("name:="&sWebButton).Click
			If err.description = "" Then
				UpdateExecutionReport micPass, "ClickSAPWebButton","Click Action performed on WebButton "& sWebButton &" on SAP Window "&sScreenName 
			Else
				UpdateExecutionReport MicFail, "ClickSAPWebButton","WebButton = "& sWebButton &" ### Frame name:="&sFrameName&" ### SAP Window = "&sScreenName & " Error = " &err.description
			End If
			Err.Clear
			on error goto 0
		End Function
	
   '***************************************************************************************************************************************************************
	
		Function ClickSAPWebLink ( sPagetitle,sFrameName,sLink)			  
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	ClickSAPWebButton ( sPagetitle,sFrameName,sWebButton)
			' @parameter:   sPagetitle	:	 Title of the Page
			' @parameter:   sFrameName	:	Techiniclal Name Of the Frame
			' @parameter:   sLink : Techinical Name of the  WebLink
			' @notes	:	Click The WebLink
			' @END
			On Error Resume Next
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Page("title:="&sPagetitle).Frame("name:="&sFrameName).Link("name:="&sLink).Click
			If err.description = "" Then
				UpdateExecutionReport micPass, "ClickSAPWebLink","Click Action performed on Link "& sLink &" on SAP Window "&sScreenName 
			Else
				UpdateExecutionReport MicFail, "ClickSAPWebLink","Link = "& sLink &" ### Frame name:="&sFrameName&" ### SAP Window = "&sScreenName & " Error = " &err.description
			End If
			Err.Clear
			on error goto 0
		End Function
	
   '***************************************************************************************************************************************************************
				
		Function ClickSAPWebImage ( sPagetitle,sFrameName,sFileName)			  
			' @HELP
			' @class	:	Controls_SAP
			' @method	:	ClickSAPWebImage ( sPagetitle,sFrameName,sFileName)
			' @parameter:   sPagetitle	:	 Title of the Page
			' @parameter:   sFrameName	:	Techiniclal Name Of the Frame
			' @parameter:   sFileName : FileName  of the  WebImage
			' @notes	:	Click The WebImage
			' @END
			On Error Resume Next		
			sScreenName = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").GetROProperty("text")
			SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Page("title:="&sPagetitle).Frame("name:="&sFrameName).Image("file name:="&sFileName).Click
			If err.description = "" Then
				UpdateExecutionReport micPass, "ClickSAPWebImage","Click Action performed on Image "& sFileName &" on SAP Window "&sScreenName 
			Else
				UpdateExecutionReport MicFail, "ClickSAPWebImage","Image = "& sFileName &" ### Frame name:="&sFrameName&" ### SAP Window = "&sScreenName & " Error = " &err.description
			End If
			Err.Clear
			on error goto 0
	   End Function
	
   '***************************************************************************************************************************************************************
	
		   
		Function GetFirstDateOfCurrentMonth(sDate)
			' @HELP
			' @class	:	Controls
			' @method	:   GetFirstDateofMonth(sDate)
			' @returns	:  	GetFirstDateofMonth : Date
			' @parameter:	sDate				: Required Date
			' @notes:   	Returns first date of month for a given date
			' @END
			sDay = Day(now)
			If CDbl(sDay) <> 1 Then
				sDay1 = CDbl(sDay) - 1 
				iStartDay = CDbl(sDay) - CDbl(sDay1)
			Else
				iStartDay = CDbl(sDay)
			End If
			If len(iStartDay) = 1 Then
				iStartDay = "0"&iStartDay
			End If
			iMonth = Month(now)
			If len(iMonth) <> 2 Then
				iMonth = "0"&iMonth
			End If
			GetFirstDateOfCurrentMonth= iMonth &"/" & iStartDay & "/" & Year(sDate)
		End Function
	
   '***************************************************************************************************************************************************************
	
		Function GetLastDateofCurrentMonth(sDate)
			' @HELP
			' @class	:	Controls
			' @method	:   GetLastDateOfMonth(sDate)
			' @returns	:  	GetLastDateOfMonth : Date
			' @parameter:	sDate			   : Required Date
			' @notes:   	Returns last date of month for a given date
			' @END
			iFirstDayNextMonth = DateSerial(Year(sDate),Month(sDate) + 1, 1)
			iLastDay = Day(DateAdd ("d", -1, iFirstDayNextMonth))
			iMonth = Month(now)
			If len(iMonth) <> 2 Then
				iMonth = "0"&iMonth
			End If
			GetLastDateofCurrentMonth = iMonth &"/"&iLastDay&"/"&Year(sDate)
		End Function
			
   '***************************************************************************************************************************************************************
	
		Function CaptureSAPScreenShot()
			' iRunID          =   Environment.Value("RunId")
			' sExecuteID		=	Environment.Value("ExecuteID")
			' sRunID			=	Environment.Value("RunId")
			' iDataSetNo		=	Environment.Value("DataSetNo")
			' iTestCaseID		=	Environment.Value("TestCaseID")
			' sResultPath		=	Environment.Value("ResultPath")
			' sScreenShotPath	=	replace(sResultPath,"HTMLReports","ScreenShots")
			' sNow			=	Now
			' sNow			=	replace(sNow," ","")
			' sNow			=	replace(sNow,"/","")
			' sNow			=	replace(sNow,":","")
			' sNow			=	replace(sNow,"AM","1")
			' sNow			=	replace(sNow,"PM","2")
			' sScreenShot		=	sScreenShotPath & iTestCaseID & "@" & iDataSetNo & "@" & sExecuteID & "@" & sRunID & "@" & sNow & ".png"
			' on error resume next
			' If SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").Exist Then
				' Set oSAPSessionObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
			' ElseIf SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").Exist Then
				' Set oSAPSessionObj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
			' End If
			' oSAPSessionObj.Activate
			' sScreenName = oSAPSessionObj.GetROProperty("text")
			' oSAPSessionObj.CaptureBitmap sSCreenShot,True
			' reporter.reportevent micDone,"||EffectaReport||CaptureSAPScreenShot","ScreenShot of SAP Window '"&sScreenName&"' ###Begin###"&sSCreenShot&"###End###"
			' Set oSAPSessionObj = Nothing
			' on error goto 0
		End Function
'***************************************************************************************************************************************************************

		
   '***************************************************************************************************************************************************************
		

		'*********************************************************************************************************************************************************
	
		''''''''***********************************************End of New Functions ******************************************************************

		
		''''''''****************&&&&&&&&&&&&&&& Begin of Functions being moved to Obsolete &&&&&&&&&&&&&&&&***********************************************
	
		''''''''****************&&&&&&&&&&&&&&& Begin of Functions being moved to Obsolete &&&&&&&&&&&&&&&&**********************************************
		End Class
