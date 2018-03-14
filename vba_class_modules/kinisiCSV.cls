VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kinisiCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
' Requires Reference "Microsoft Scripting Runtime"
'
' We assume that the fields are 9 and are tab delimeted

Private lines() As String
Private m_rowsCount As Integer
Private m_RowsUpdated As Integer
Private firstLine As Integer
Private data() As requests_dao
Private dict As Scripting.Dictionary
Private dateformat As String

Private Sub Class_Initialize()
    Set dict = New Scripting.Dictionary
    
    dict.Add "order", 0
    dict.Add "barcode", 1
    dict.Add "firstName", 2
    dict.Add "lastName", 3
    dict.Add "licenceCode", 4
    dict.Add "user", 5
    dict.Add "printStatus", 6
    dict.Add "insertDate", 7
    dict.Add "updateDate", 8

    dateformat = getConfig("FormatDate")
    m_rowsCount = 0
    m_RowsUpdated = 0
End Sub


Property Get rowsCount()
    rowsCount = m_rowsCount
End Property

Property Get RowsUpdated()
    RowsUpdated = m_RowsUpdated
End Property

Public Sub parse(dvalue As String)
    
    detectCSVBoundaries (dvalue)
    readData (dvalue)
    updateDB

End Sub

' This updates the private variables:
' - firstLine : from which line to start reading
' - m_rowsCount : how many lines are expected
Private Sub detectCSVBoundaries(dvalue As String)
    csvText = dvalue
    Dim i, index As Integer
    Dim Fields() As String

' I am going to avoid to make this too generic.
' The file has some lines in the begining and then the table with 10 fields.
' This is what I care about

' find first row of csv (tab delimited) and number of rows
    index = 0
    lines = Split(csvText, vbCrLf)
    
    For i = 0 To UBound(lines)
        Fields = Split(lines(i), vbTab)
        If ((UBound(Fields) + 1) = 9) Then
            If (Fields(dict("barcode")) = "Bar Code") Then
                ' The assumption is that we found the header
                ' The next line will have the first row with data
                firstLine = i + 1
                index = 0
            ElseIf (Fields(dict("order")) = index + 1) Then
                ' Just a sanity check that we do not have broken data
                If (firstLine + index <> i) Then
                    Err.raise 951, "Λάθος", "Τα δεδομένα της σελίδας φαίνεται να μην είναι υγιή"
                End If
                index = index + 1
            End If
        End If
    Next i
    
    ' And we found the number of rows
    m_rowsCount = index
End Sub


Private Sub readData(csvText As String)
    Dim j As Integer
    Dim Fields() As String
    ReDim data(0 To m_rowsCount - 1)
    For i = firstLine To (firstLine + m_rowsCount - 1)
        j = i - firstLine
        Set data(j) = New requests_dao
        Fields = Split(lines(i), vbTab)
        With data(j)
            .barcode = Fields(dict("barcode"))
            .firstName = Fields(dict("firstName"))
            .lastName = Fields(dict("lastName"))
            .licenceCode = Fields(dict("licenceCode"))
            .printStatus = Fields(dict("printStatus"))
            .insertedDate = stringToDate(Fields(dict("insertDate")))
            .updatedDate = Now
        End With
    Next i

End Sub


Private Sub updateDB()
    Dim m_db As requests_db
    Set m_db = New requests_db
    m_RowsUpdated = m_db.massUpdate(data)
    Set m_db = Nothing
End Sub


Function isInArray(stringToBeFound As String, arr As Variant) As Boolean
  isInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function




Private Sub Class_Terminate()
    Set dict = Nothing
End Sub
