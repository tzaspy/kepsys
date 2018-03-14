Attribute VB_Name = "test_shared"
Option Compare Database

Sub run_all_tests()
    test_getConfig
    test_stringToDate
    test_validateBarCode
    test_isInArray
End Sub

Sub test_stringToDate()
    Dim mdate As String
    On Error GoTo error_test
    mdate1 = "26/2/2018 10:30:17 µµ"

    ' Test 1 : Recognise the pm
    mdate = "26/2/2018 10:30:17 µµ"
    If (Format(stringToDate(mdate), "YYYY-MM-DD HH:MM:SS") = "2018-02-26 22:30:17") Then
        Debug.Print "Success to recognise the date #" & mdate1 & "#"
    Else
        Debug.Print "Failed to recognise the date #" & mdate1 & "#"
        GoTo error_test
    End If
     
    ' Test 2 : Recognise the am
    mdate = "26/2/2018 10:30:17 ðµ"
    If (Format(stringToDate(mdate), "YYYY-MM-DD HH:MM:SS") = "2018-02-26 10:30:17") Then
        Debug.Print "Success to recognise the date #" & mdate1 & "#"
    Else
        Debug.Print "Failed to recognise the date #" & mdate1 & "#"
        GoTo error_test
    End If
    
    ' Test 3 : Recognize invalid date
    mdate1 = "2/26/2018 25:45:99"
    On Error GoTo expected_error
    stringToDate (mdate1)

    
    On Error GoTo error_test
    Debug.Print "========================"
    Debug.Print "SUCCESS: stringToDate"
    Debug.Print "========================"
    
    Exit Sub
expected_error:
    Debug.Print "Success: Expected error with description: " & Err.Description
    Resume Next

error_test:
    Debug.Print "########################"
    Debug.Print "FAILURE: stringToDate"
    Debug.Print "########################"
End Sub


Sub test_validateBarCode()
    Dim bc As String
    On Error GoTo error_test
    ' Successful
    bc = "0011558813213"
    If validateBarCode(bc) = False Then
        Err.raise 885, , "Failure in BarCode validate"
    End If
    
    ' Failure, the string has 14 characters
    bc = "00115588132112"
    If validateBarCode(bc) = False Then
        Debug.Print "Success: validation of barcode '" & bc & "' (14char) failed with error : " & Err.Description
    Else
        Debug.Print "Failed: validation of barcode '" & bc & "' (14char) succeded"
        GoTo error_test
    End If

    ' Failure, the barcode has 12 digits
    bc = "001155881324"
    If validateBarCode(bc) = False Then
        Debug.Print "Success: validation of barcode '" & bc & "' (12char) failed with error : " & Err.Description
    Else
        Debug.Print "Failed: validation of barcode '" & bc & "' (12char) succeded"
        GoTo error_test
    End If

    ' Failure, the barcode has a character
    bc = "001a558813213"
    If validateBarCode(bc) = False Then
        Debug.Print "Success: validation of barcode '" & bc & "' (alphanumeric character) failed with error : " & Err.Description
    Else
        Debug.Print "Failed: validation of barcode '" & bc & "' (alphanumeric character) succeded"
        GoTo error_test
    End If
    Debug.Print "========================"
    Debug.Print "SUCCESS: validateBarCode"
    Debug.Print "========================"
    
    Exit Sub
    
error_test:
    Debug.Print "########################"
    Debug.Print "FAILURE: validateBarCode"
    Debug.Print "########################"
    Exit Sub

End Sub

Sub test_getConfig()
    
    On Error GoTo error_test
    If getConfig("FormatDate") = "yyyy-mm-dd hh:mm:ss" Then
        Debug.Print "Success to recognize 'FormatDate' config"""
    Else
        Debug.Print "Failed to recognise 'FormatDate' config"
        GoTo error_test
    End If
    
    On Error GoTo expected_error:
    getConfig ("RandomConfig")
    
    Debug.Print "========================"
    Debug.Print "SUCCESS: getConfig      "
    Debug.Print "========================"
    
    Exit Sub
    
expected_error:
    Debug.Print "Success: Error has description: " & Err.Description
    Resume Next

error_test:
    Debug.Print "########################"
    Debug.Print "FAILURE: getConfig      "
    Debug.Print "########################"
    Exit Sub

End Sub


Sub test_isInArray()
    Dim arr(0 To 9) As requests_dao
    Dim i As Integer
    Dim ret As Integer
    
    For i = LBound(arr) To UBound(arr)
        Set arr(i) = New requests_dao
        arr(i).barcode = "00114647620" & Chr(48 + i) & Chr(48 + i)
    Next i
        
    ' Test 1 : The barcode exists and the table has even number of elements
    For i = LBound(arr) To UBound(arr)
        If (isInArray(arr(i).barcode, arr) = i) Then
            Debug.Print "Success: Table has even elements, barcode " & arr(i).barcode & " was found on " & i & "position"
        Else
            Debug.Print "Failed: Table has even elements, barcode " & arr(i).barcode & " wasn't found on " & i & "position"
            GoTo error_test
        End If
    Next i
        
    ' Test 2 : Te barcode doesn't exists and is smaller from the first element
    pos = isInArray("0011464761000", arr)
    If (pos = -1) Then
        Debug.Print "Success: Table has even elements, barcode was not found"
    Else
        Debug.Print "Failed: Table has even elements, barcode was found on " & pos & "position"
        GoTo error_test
    End If
        
        
    ' Test 3 : Te barcode doesn't exists and is in the range of first and last elements of the table
    pos = isInArray("0011464762078", arr)
    If (pos = -1) Then
        Debug.Print "Success: Table has even elements, barcode was not found"
    Else
        Debug.Print "Failed: Table has even elements, barcode was found on " & pos & "position"
        GoTo error_test
    End If
        
    ' Test 4 : Te barcode doesn't exists and is bigger of the last elements of the table
    pos = isInArray("0011464765078", arr)
    If (pos = -1) Then
        Debug.Print "Success: Table has even elements, barcode was not found"
    Else
        Debug.Print "Failed: Table has even elements, barcode was found on " & pos & "position"
        GoTo error_test
    End If
        
    Debug.Print "========================"
    Debug.Print "SUCCESS: isInArray      "
    Debug.Print "========================"
    
    Exit Sub
    
error_test:
    Debug.Print "########################"
    Debug.Print "FAILURE: isInArray      "
    Debug.Print "########################"
    Exit Sub


    
End Sub
