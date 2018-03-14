Attribute VB_Name = "shared"
Option Compare Database

' This should be private for kinisiCSV.
' But I wish to test this and also has the potential to make this more generic if needed.
' For now this is usefull only for csvkinisi dates
Public Function stringToDate(dt As String) As Date
' Requires "Microsoft VBScript Regular Expressions 5.5"
    Dim regEx As New RegExp
    Dim datePattern As String
    Dim retDate As Date
    Dim retTime As Date
    Dim dtMatch() As Variant
    Dim dtDay As Integer
    Dim dtMonth As Integer
    Dim dtYear As Integer
    Dim dtHour As Integer
    Dim dtMin As Integer
    Dim dtSec As Integer
    Dim dtPM As String
    
    ' The code is not that generic.
    ' So, I do not think that make sense to have this in a config
    datePattern = "(\d+)/(\d+)/(20\d\d) (\d+):(\d+):(\d+) (..)"
     
    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = datePattern
    End With

    If regEx.test(dt) Then
        Set dtmatches = regEx.Execute(dt)
    Else
        Err.raise 775, "Μετατροπή ημερομηνίας", "Η ημερομηνία δεν μπορεί να αναγωνριστεί"
    End If

    dtDay = dtmatches(0).SubMatches(0)
    dtMonth = dtmatches(0).SubMatches(1)
    dtYear = dtmatches(0).SubMatches(2)
    dtHour = dtmatches(0).SubMatches(3)
    dtMin = dtmatches(0).SubMatches(4)
    dtSec = dtmatches(0).SubMatches(5)
    dtPM = dtmatches(0).SubMatches(6)
    
    ' This is a hack. The reason is the greek letters. In MS Access 2003 there is different
    ' encoding between vba and front end. "μμ" has the same characters and I use this to
    ' identify if it is pm or am
    If (Left(dtPM, 1) = Right(dtPM, 1)) Then
        dtHour = dtHour + 12
    End If

    retDate = DateSerial(dtYear, dtMonth, dtDay)
    retTime = TimeSerial(dtHour, dtMin, dtSec)

    stringToDate = retDate + retTime

End Function

Function validateBarCode(bc As String) As Boolean
    ' Requires "Microsoft VBScript Regular Expressions 5.5"
    Dim regEx As New RegExp
    Dim bcPattern As String

    bcPattern = getConfig("barcodePattern")
     
    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = bcPattern
    End With
    
    If regEx.test(bc) Then
        validateBarCode = True
    Else
        validateBarCode = False
    End If
End Function

Function getConfig(configName As String) As String
    Dim DB As Database
    Dim rst As Recordset
    Dim strSQL As String
    
    strSQL = "select configValue from tconfig where configName = '" & configName & "';"
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset(strSQL)
    
    If rst.RecordCount <> 1 Then
        Err.raise 651, "", "Πρόβλημα του συστήματος. Η παράμετρος " & confiname & " δεν υπάρχει."
    End If
    
    rst.MoveFirst
    getConfig = rst.Fields("configValue")
    
    Set rst = Nothing
    Set DB = Nothing
End Function

Function isInArray(barcode As String, arr As Variant) As Integer
    ' arr is an array of request_dao
    ' arr is sorted by barcodes

    Dim l, r, m, n As Integer
    Dim data() As requests_dao
    
    data = arr
    
    l = 0
    r = UBound(data)
    ' By default the record is not found
    isInArray = -1
    While (l <= r And isInArray = -1)
        m = CInt((l + r) / 2)
        If (arr(m).barcode < barcode) Then
            l = m + 1
        ElseIf (arr(m).barcode > barcode) Then
            r = m - 1
        ElseIf (arr(m).barcode = barcode) Then
            isInArray = m
        End If
    Wend

End Function

Function setLastprint(dt As String)
    Dim DB As Database
    Dim strSQL As String
    Set DB = CurrentDb
    strSQL = "update tconfig set configValue  = '" & dt & "' where configName = 'lastPrint'"
    DoCmd.SetWarnings False
        DB.Execute strSQL
    DoCmd.SetWarnings True
    Set DB = Nothing
End Function
