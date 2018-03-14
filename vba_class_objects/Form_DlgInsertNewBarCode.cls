VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DlgInsertNewBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private rdb As requests_db

Private Sub Ctl_InsertNewBarcode_Click()
On Error GoTo Err_Ctl_InsertNewBarcode_Click

    Dim barcode As String
    Dim kepcode As Long
    Dim errorMessage As String
    
' Validation Rulues - Not null

    If IsNull(Forms!DlgInsertNewBarCode!barcode.value) Then
        errorMessage = "Το barcode πρέπει να έχει μια τιμή."
        GoTo Exit_DuplicateBarCode
    End If
    
    
    If IsNull(Forms!DlgInsertNewBarCode!cbKep.value) Then
        errorMessage = "Το ΚΕΠ πρέπει να έχει μια τιμή."
        GoTo Exit_DuplicateBarCode
    End If
 
   
    barcode = Forms!DlgInsertNewBarCode!barcode.value
    kepcode = CLng(Forms!DlgInsertNewBarCode!cbKep.value)

    rdb.insertBarcode barcode, kepcode
    

' Clean the form

    Forms!DlgInsertNewBarCode!barcode.value = Null
    Forms!DlgInsertNewBarCode!cbKep.value = Null


Exit_Ctl_InsertNewBarcode_Click:
    Exit Sub

Exit_DuplicateBarCode:
    MsgBox errorMessage, vbExclamation, "Ανεπιτυχής εισαγωγή"
    Exit Sub

Err_Ctl_InsertNewBarcode_Click:
    MsgBox Err.Description
    Resume Exit_Ctl_InsertNewBarcode_Click
    
End Sub

Private Sub Form_Load()
    Set rdb = New requests_db
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rdb = Nothing
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
