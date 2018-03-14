VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "requests_db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private DB As Database
Private rst As Recordset
Private row As requests_dao
Private dateformat As String


Private Sub Class_Initialize()
    Set DB = CurrentDb
    dateformat = getConfig("FormatDate")
End Sub

Public Function barcodeExists(bc As String) As Boolean
    Set row = Nothing
    strSQL = "select " _
            & "barcode, " _
            & "kepcode, " _
            & "fname, " _
            & "lname, " _
            & "licence_code, " _
            & "print_status, " _
            & "cr_date, " _
            & "insert_date, " _
            & "last_update, " _
            & "resolve_date, " _
            & "status " _
        & "from " _
            & "requests " _
        & "where " _
            & "barcode = '" & bc & "';"
    
    Set rst = DB.OpenRecordset(strSQL)
    If (rst.RecordCount > 0) Then
        barcodeExists = True
        rst.MoveFirst
        
        Set row = New requests_dao
        row.barcode = rst.Fields("barcode").value
        row.firstName = rst.Fields("fname").value & ""
        row.lastName = rst.Fields("lname").value & ""
        row.licenceCode = rst.Fields("licence_code").value & ""
        row.printStatus = rst.Fields("print_status").value & ""
        If Not IsNull(rst.Fields("cr_date").value) Then
            row.createdDate = rst.Fields("cr_date").value
        End If
        If Not IsNull(rst.Fields("insert_date").value) Then
            row.insertedDate = rst.Fields("insert_date").value
        End If
        If Not IsNull(rst.Fields("last_update").value) Then
            row.updatedDate = rst.Fields("last_update").value
        End If
        If Not IsNull(rst.Fields("resolve_date").value) Then
            row.resolvedDate = rst.Fields("resolve_date").value
        End If
        row.status = rst.Fields("status") & ""
    Else
        barcodeExists = False
    End If
    
    Set rst = Nothing
End Function

Public Function isUpdated() As Boolean
    isUpdated = False
    If Not IsNull(row) Then
        If Len(row.firstName) > 0 And Len(row.lastName) > 0 _
            And Len(row.licenceCode) > 0 Then
            isUpdated = True
        End If
    End If

End Function

Public Sub insertBarcode(barcode As String, kep As Long)
' Check if there is already a value
    
    If validateBarCode(barcode) = False Then
        Err.raise 878, "", "Το barcode " & barcode & " δεν έχει τη σωστή μορφή."
    End If
    
    strSQLExists = "select 1 from requests where barcode = '" & barcode & "';"
    
    Set rst = DB.OpenRecordset(strSQLExists)
    If (rst.RecordCount > 0) Then
        Err.raise 778, "", "Το barcode " & barcode & " είναι ήδη καταχωρημένο."
    End If
    
    Set rst = Nothing

' Insert the new Barcode and KEP in DB

    strSQLInsert = "Insert into requests " _
        & "(barcode, kepcode, status) " _
      & "values " _
         & "('" & barcode & "', " & kep & ", 'N');"

    
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQLInsert
    DoCmd.SetWarnings True
    
End Sub

Public Function markCompleted(barcode As String)
    
    strSQL = "update requests " _
        & "set " _
            & "status = 'C', " _
            & "resolve_date = #" & Format(Now, dateformat) & "# " _
        & "where " _
            & "barcode = '" & barcode & "' " _
        & "and status <> 'C';"

    DoCmd.SetWarnings False
    DB.Execute strSQL
    DoCmd.SetWarnings True
    markCompleted = DB.RecordsAffected
End Function

Public Function massUpdate(data() As requests_dao) As Integer
    Dim existingBarcodes() As requests_dao
    Dim serial_barcodes As String
    Dim pos As Integer
    
    ' Lets find the imported barcodes
    ' Initialize the string.
    existingBarcodes = getMatchedBarcodes(data)

    DoCmd.SetWarnings False
    m_RowsUpdated = 0
    For Each r In data
        pos = isInArray(r.barcode, existingBarcodes)
        If (pos > -1) Then
            sqlupdate = createUpdateSQL(r, existingBarcodes(pos))
            If Len(sqlupdate) > 0 Then
                DB.Execute sqlupdate
                m_RowsUpdated = m_RowsUpdated + DB.RecordsAffected
            End If
        End If
    Next r
    DoCmd.SetWarnings True
    massUpdate = m_RowsUpdated
End Function

Function createUpdateSQL(request_data As Variant, Optional existing_row As requests_dao) As String
    Dim upd As String
    Dim dbrow As requests_dao
    
    ' initialize the optional parameter
    If IsMissing(existing_row) Then
        Set dbrow = New requests_dao
    Else
        Set dbrow = existing_row
    End If

    upd = ""
    If dbrow.firstName = "" Or dbrow.lastName = "" Or dbrow.licenceCode = "" Then
        upd = "lname = """ & request_data.lastName & """, " _
            & "fname = """ & request_data.firstName & """, " _
            & "licence_code = """ & request_data.licenceCode & """"
    End If
    
    If request_data.printStatus <> dbrow.printStatus Then
        If Len(upd) > 0 Then
            upd = upd & ", "
        End If
        upd = upd & "print_status = """ & request_data.printStatus & """, " _
            & "insert_date = #" & Format(request_data.insertedDate, dateformat) & "#, " _
            & "last_update = #" & Format(request_data.updatedDate, dateformat) & "# "
    End If
    
    If dbrow.status = "N" Then
        ' This check I am almost sure that is not needed...
        If Len(upd) > 0 Then
            upd = upd & ", "
        End If
        upd = upd & "status = 'P' "
    End If
    
    If Len(upd) > 0 Then
        createUpdateSQL = "update requests " _
            & "set " & upd _
            & "where barcode = '" & request_data.barcode & "';"
    Else
        createUpdateSQL = ""
    End If
    
    Debug.Print "Query to update : " & createUpdateSQL
    
End Function

Private Function getMatchedBarcodes(data() As requests_dao) As Variant
    Dim serial_barcodes As String
    Dim i As Integer
    Dim r As Variant
    Dim rows() As requests_dao
    
    ' Serialize the barcodes to make parameter in the query
    serial_barcodes = ""
    For Each r In data
        If serial_barcodes = "" Then
            serial_barcodes = """" & r.barcode & """"
        Else
            serial_barcodes = serial_barcodes & ",""" & r.barcode & """"
        End If
    Next r
    
    strSQLExists = "select barcode, fname, lname, licence_code, print_status, status from requests where barcode in (" & serial_barcodes & ") order by barcode asc;"
    
    Set rst = DB.OpenRecordset(strSQLExists)
    i = 0
    ReDim rows(0 To rst.RecordCount - 1)
    rst.MoveFirst
    Do Until rst.EOF = True
        Set rows(i) = New requests_dao
        rows(i).barcode = rst.Fields("barcode").value
        rows(i).firstName = rst.Fields("fname").value & ""
        rows(i).lastName = rst.Fields("lname").value & ""
        rows(i).licenceCode = rst.Fields("licence_code").value & ""
        rows(i).printStatus = rst.Fields("print_status").value & ""
        rows(i).status = rst.Fields("status").value & ""
        rst.MoveNext
        i = i + 1
    Loop

    Set rst = Nothing
    getMatchedBarcodes = rows
End Function


Private Sub Class_Terminate()
    Set DB = Nothing
End Sub
