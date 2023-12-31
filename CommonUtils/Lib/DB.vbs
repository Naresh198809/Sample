'*********************************************************************************************************************
'**  Project            : 
'**  File Name          : DB.vbs
'**  Version            : 
'**  Created on         : 
'**  Updated on         : 
'**  Description        : Common Methods for extracting the data from the Db files
'**  Copyright          : Arsin Corporation.
'**  Author             :       
'*********************************************************************************************************************

Class clsDataAccess
	
	Dim cnnAccess,rs
		
Private Sub Class_Initialize 
	ConnectAccessDB DB_PATH
End Sub
'------------------------------------------------------------------------------------------------------------------------
'		Private Sub Class_Terminate
'			DisconnectAccessDB
'		End Sub
'------------------------------------------------------------------------------------------------------------------------
' For Access Database
Function ConnectAccessDB(sFileName)

		' @HELP
		' @class	: Access
		' @method	: ConnectAccessDB(sFileName)
		' @returns	: None 
		' @parameter: sFileName: Name of the mdb file
		' @notes	: Connect to the specified database
		' @END
		REM MsgBox sFileName
			
		Set cnnAccess = CreateObject("ADODB.Connection")
		cnnAccess.Open "DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & sFileName & ";Readonly=False;Extended Properties={HDR=YES;IMEX=1}"

End Function
'------------------------------------------------------------------------------------------------------------------------
Function DisconnectAccessDB()

		' @HELP
		' @class	: clsAccess
		' @method	: DisconnectAccessDB
		' @returns	: None 
		' @parameter: None
		' @notes	: Disconnect from the specified Database.
		' @END
		
		Wait 2
		cnnAccess.Close
		
		If cnnAccess.State = 0 Then
			'MsgBox "Close"
			ReportWriter "Pass","Database Connection","Database Connection is Closed",0
		ElseIf cnnAccess.State = 1 Then
			ReportWriter "Fail","Database Connection","Database Connection is not Closed",0
		End If
		
End Function
'------------------------------------------------------------------------------------------------------------------------
Function ExecSQLStatementWithWhereClass(sSQL)

		' @HELP
		' @class	: Access
		' @method	: ExecSQLStatementWithWhereClass(sSQL)
		' @returns	: Gets row data from database
		' @parameter: sSQL- table name
		' @notes	: Returns data of the specified table of a given database.
		' @END
		
		Dim sColumnHeader,sCellValue
		Dim objDictionary
		
		Set objDictionary = CreateObject("Scripting.Dictionary")
		Set rs =CreateObject("ADODB.RECORDSET")
	   
		rs.Open sSQL,cnnAccess
		
		'While Not rs.EOF
		If Not rs.EOF Then
			sColumnHeader = ""
			sCellValue = ""
			For i = 0 To rs.Fields.Count -1
				sColumnHeader = rs.Fields(i).Name
				sCellValue = rs.Fields(i).Value
				objDictionary.Add sColumnHeader,sCellValue
			Next
			rs.MoveNext
'		Else 
			'MsgBox "No Records"
'			ReportWriter "Fail","DB Connection successful with ","No Records Found" & sSQL,0
		End If
		'Wend
		Set ExecSQLStatementWithWhereClass = objDictionary
		rs.Close
		
		Set objDictionary = Nothing
		Set rs = Nothing
		'Set cnnAccess = Nothing
			
End Function
'------------------------------------------------------------------------------------------------------------------------
Function GetSingleRowValuefromAccessDB(sSQL)
		' @HELP
		' @class	: Access
		' @method	: GetSingleRowValuefromAccessDB(sSQL)
		' @returns	: Gets single row data from database
		' @parameter: sSQL- table name
		' @notes	: Returns data of the specified table of a given database.
		' @END
		
		Dim bFlag,iColCount,iRowCount
		iColCount = 0
		iRowCount = 0
		''**************************
'		Set rs =CreateObject("ADODB.RECORDSET")
'		rs.Open sSQL,cnnAccess
		
		set rs = cnnAccess.Execute (sSQL)
		If Err.Number <> 0 Then
			ReportWriter "Fail","Read Functional scenarios","Read Data Failure,Please Check Input sheet",0
			Err.Clear 
		End If
		''***************************
		
		iRowCount = 0 
		If rs.EOF Then
			ReportWriter "Fail","Read Functional scenarios","No data Found for Search Crateria,Please Check for Flag",0
			'MsgBox "No data Found for Search Crateria,Please Check for Flag"
			Exit Function
		End if
		Do While Not rs.EOF
		iRowCount = iRowCount + 1
		rs.MoveNext
		Loop
		
		rs.MoveFirst
		
		If Not rs.EOF Then
			bFlag = True
			ReDim arrData(iRowCount-1,rs.Fields.Count -1)
			For i = 0 To iRowCount -1 
				For j = 0 To rs.Fields.Count-1 
					If IsNull(rs.Fields.Item(j)) Then
						arrData(i,j) = ""
					Else
						arrData(i,j) =  rs.Fields.Item(j)
					End If
				Next
				rs.MoveNext
			Next
		End If	
		
		If bFlag = True Then
			'ReportWriter "Pass","DB Connection successful with ",iRowCount & "Record(s) Found ",0
			'msgbox "ok"
		Else
			'ReportWriter "Fail","DB Connection successful with ","No Records Found" & sSQL,0
		End If
		GetSingleRowValuefromAccessDB = arrData
		rs.Close				
		Set rs = Nothing
End Function
'------------------------------------------------------------------------------------------------------------------------
Function InsertValuesIntoAccessDB(sSQL)
	
		' @HELP
		' @class	: Access
		' @method	: InsertValuesIntoAccessDB(sSQL)
		' @returns	: This method is used to Insert test data into the database
		' @parameter: sSQL- table name
		' @notes	: Returns data of the specified table of a given database.
		' @END
		
		Dim Retlng
		Set rs =CreateObject("ADODB.RECORDSET")

		cnnAccess.Execute sSQL,Retlng
		'MsgBox Retlng

		If Retlng >= 1 Then
			Set rs = cnnAccess.Execute ("SELECT @@Identity as Inserted_RowId")	' RecordsetSELECT @@Identity as Inserted_RowId;
			PrmKey =  rs.Fields(0).Value
			InsertValuesIntoAccessDB = PrmKey
			ReportWriter "Pass","DB Connection successful with ","record(s)"& Retlng &"inserted" ,0
		Else
			ReportWriter "Fail","DB Connection successful with ","NO record inserted " & sSQL,0
		End If
		
		rs.Close
		Set rs = Nothing
		
End Function
'------------------------------------------------------------------------------------------------------------------------
'Points To be rememberd
'Primary Key Is Compulsory In when using the InsertValuesIntoAccessDB
Function UpdateValuesIntoAccessDB(sSQL)
		
		' @HELP
		' @class	: Access
		' @method	: UpdateValuesIntoAccessDB(sSQL)
		' @returns	: This method is used to update the existing values/data in the database
		' @parameter: sSQL- table name
		' @notes	: Returns data of the specified table of a given database.
		' @END

		Dim Retlng
		
		cnnAccess.Execute sSQL,Retlng
		If Retlng >= 1 Then
			ReportWriter "Pass","DB Connection successful with ","record(s)"& Retlng &"inserted" ,0
		Else
			ReportWriter "Fail","DB Connection successful with ","NO record inserted " & sSQL,0
		End If
			
End Function
'------------------------------------------------------------------------------------------------------------------------
End Class
