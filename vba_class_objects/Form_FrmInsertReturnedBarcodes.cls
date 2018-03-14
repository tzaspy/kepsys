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
        log "Αναγνωρίστηκε καταχώρηση του barcode '" & barcode & "'."
        If rdb.barcodeExists(barcode) Then
            If rdb.markCompleted(barcode) = 1 Then
                If rdb.isUpdated Then
                    log ("Το BarCode '" & barcode & "' καταχωρήθηκε ως οληκληρωμένο.")
                Else
                    log ("Το BarCode '" & barcode & "' υπάρχει στη βάση αλλά δεν είναι ενημερωμένο. " _
                        & "Λείπει ένα η περισσότερα από τα πεδία: όνομα, επώνυμο, αριθμός άδειας.")
                End If
            End If
        Else
            log ("Το BarCode '" & barcode & "' δεν έχει καταχωρηθεί.")
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
