VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmInsertUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private orgPage As String

Private Sub btImportData_Click()
On Error GoTo Err_btImportData_Click

    Dim csv As kinisiCSV
    
    If IsNull(Forms!frminsertupdates!txtCopiedPage.value) Then
        Err.Description = "Η φόρμα δεν μπορεί να είναι κενή. Διαβάστε τις οδηγίες."
        GoTo Err_btImportData_Click
    End If
    
    
    Set csv = New kinisiCSV
    
    csv.parse (Forms!frminsertupdates!txtCopiedPage.value)
    
    msg = "Διαβάστηκαν " & csv.rowsCount & " γραμμές από τον πίνακα που εισάγατε." & vbCrLf _
         & "Από αυτές αναγνωρίστηκαν και ενημερώθηκαν " & csv.RowsUpdated & " εγγραφές."

    MsgBox msg, vbInformation, "Αποτέλεσμα εισαγωγής"

Exit_btImportData_Click:
    Exit Sub

Err_btImportData_Click:
    MsgBox Err.Description
    GoTo Exit_btImportData_Click
    
End Sub


' Unfortunately MS Access trims trailing spaces after update.
' We need to temporarily save the text before the update and add at the end a character
' to prevent MS Access to trim the trailling spaces
Private Sub btnClean_Click()
  Forms!frminsertupdates!txtCopiedPage.value = Null
End Sub

' Here we re-update value of the text box. We also add a line and a space, assuming that
' our code can ignore this.
Private Sub txtCopiedPage_AfterUpdate()
 Forms!frminsertupdates!txtCopiedPage.value = orgPage & vbCrLf & "-"
End Sub

Private Sub txtCopiedPage_BeforeUpdate(Cancel As Integer)
    orgPage = Me!txtCopiedPage.Text
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
