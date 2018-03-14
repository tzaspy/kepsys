VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Completed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(Cancel As Integer)
    Dim strSQL As String
    Dim start_date As String
    Dim end_date As String
    
    start_date = Forms!frmparametersforcompleted!tbDateStart.value
    end_date = Forms!frmparametersforcompleted!tbDateEnd.value
    ' TODO
    ' For now I am leaving a gap. I should check for valid format of the date...
    ' I believe that this should not be a problem (was the last famous words...)
    strSQL = "SELECT " _
            & "requests.barcode, " _
            & "requests.fname, " _
            & "requests.lname, " _
            & "requests.licence_code, " _
            & "kep.FORMAL_TITLE " _
        & "FROM requests LEFT JOIN kep " _
        & "ON requests.kepcode=kep.p_bodyid " _
        & "where requests.status ='C' " _
        & "and requests.resolve_date between #" & start_date & "# and #" & end_date & "#;"
     Debug.Print strSQL
     Me.RecordSource = strSQL
End Sub
