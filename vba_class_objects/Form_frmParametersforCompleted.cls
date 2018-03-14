VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmParametersforCompleted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private lastPrint As String
Private dateformat As String

Private Sub Form_Load()
    lastPrint = getConfig("lastPrint")
    dateformat = getConfig("FormatDate")
    
    Forms!frmparametersforcompleted!optSelectDate.value = 1
    isEnabled = False
    start_date = lastPrint
    end_date = Format(Now(), dateformat)
    Forms!frmparametersforcompleted!tbDateStart.Enabled = isEnabled
    Forms!frmparametersforcompleted!tbDateEnd.Enabled = isEnabled
    Forms!frmparametersforcompleted!tbDateStart.value = start_date
    Forms!frmparametersforcompleted!tbDateEnd.value = end_date

End Sub

Private Sub optSelectDate_AfterUpdate()
    Dim m_option As Byte
    Dim isEnabled As Boolean
    isEnabled = False
    m_option = Forms!frmparametersforcompleted!optSelectDate.value
    
    Select Case m_option
        Case 1
            isEnabled = False
            start_date = lastPrint
            end_date = Format(Now(), dateformat)
        Case 2
            isEnabled = False
            start_date = Left(Format(Now(), dateformat), 11) & "00:00:00"
            end_date = Format(Now(), dateformat)
        Case 3
            isEnabled = True
            start_date = lastPrint
            end_date = Format(Now(), dateformat)
    End Select
    
    Forms!frmparametersforcompleted!tbDateStart.Enabled = isEnabled
    Forms!frmparametersforcompleted!tbDateEnd.Enabled = isEnabled
    Forms!frmparametersforcompleted!tbDateStart.value = start_date
    Forms!frmparametersforcompleted!tbDateEnd.value = end_date
    

End Sub
Private Sub cbOpentRptCompleted_Click()
On Error GoTo Err_cbOpentRptCompleted_Click

    Dim stDocName As String

    
    stDocName = "Completed"
    DoCmd.OpenReport stDocName, acPreview

Exit_cbOpentRptCompleted_Click:
    Exit Sub

Err_cbOpentRptCompleted_Click:
    MsgBox Err.Description
    Resume Exit_cbOpentRptCompleted_Click
    
End Sub
Private Sub btnQuit_Click()
On Error GoTo Err_btnQuit_Click


    DoCmd.Close

Exit_btnQuit_Click:
    Exit Sub

Err_btnQuit_Click:
    MsgBox Err.Description
    Resume Exit_btnQuit_Click
    
End Sub
Private Sub btnPrint2PDF_Click()
On Error GoTo Err_btnPrint2PDF_Click

    Dim stDocName As String
    Dim filename As String
    Dim b As Boolean
    Dim folder As String
    Dim dt As String
    Dim dateformat As String

    stDocName = "Completed"
    dateformat = getConfig("FormatDate")
    dt = Format(Forms!frmparametersforcompleted!tbDateEnd.value, "yyyy-mm-dd")
    folder = getConfig("reporDirectory")
    
    b = ConvertReportToPDF("Completed", vbNullString, _
folder & "/" & stDocName & "_" & dt & ".pdf", False, True, 150, "", "", 0, 0, 0)
    
    If b = True Then
        setLastprint (Format(Now(), dateformat))
    End If
        
Debug.Print b
Exit_btnPrint2PDF_Click:
    Exit Sub

Err_btnPrint2PDF_Click:
    MsgBox Err.Description
    Resume Exit_btnPrint2PDF_Click
    
End Sub

