VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmInsertReturnedBarcodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private logDateFormat As String
Private rdb As requests_db
Private r As requests_dao

Private Sub btlnClean_Click()
    cleanBarcode
End Sub
Private Sub cleanBarcode()
    Forms!FrmInsertReturnedBarcodes!tbReturnedBarCodes.value = ""
End Sub
Private Sub Form_Load()
Set rdb = New requests_db
End Sub


Private Sub tbReturnedBarCodes_Change()
    Dim barcode As String
    barcode = Forms!FrmInsertReturnedBarcodes!tbReturnedBarCodes.Text
    If validateBarCode(barcode) Then
        log "������������� ���������� ��� barcode '" & barcode & "'."
        If rdb.barcodeExists(barcode) Then
            If rdb.markCompleted(barcode) = 1 Then
                If rdb.isUpdated Then
                    log ("�� BarCode '" & barcode & "' ������������ �� ������������.")
                Else
                    log ("�� BarCode '" & barcode & "' ������� ��� ���� ���� ��� ����� �����������. " _
                        & "������ ��� � ����������� ��� �� �����: �����, �������, ������� ������.")
                End If
            End If
        Else
            log ("�� BarCode '" & barcode & "' ��� ���� �����������.")
        End If
        cleanBarcode
    End If
    
End Sub


Private Sub log(msg As String)
    logDateFormat = "yyyy-mm-dd HH:mm:ss"
    Forms!FrmInsertReturnedBarcodes!tbLogs.value = "" _
        & Format(Now, logDateFormat) & ": " _
        & msg & vbCrLf _
        & Forms!FrmInsertReturnedBarcodes!tbLogs.value
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

Private Sub tbReturnedBarCodes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
    End If
End Sub

Private Sub tbReturnedBarCodes_KeyPress(KeyAscii As Integer)

End Sub

Private Sub tbReturnedBarCodes_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub
