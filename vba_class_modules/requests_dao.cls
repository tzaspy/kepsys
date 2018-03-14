VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "requests_dao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private dbBarCode As String
Private dbKEPCode As String
Private dbFirstName As String
Private dbLastName As String
Private dbLicenceCode As String
Private dbPrintStatus As String
Private dbCreatedDate As Date
Private dbInsertedDate As Date
Private dbUpdatedDate As Date
Private dbResolvedDate As Date
Private dbStatus As String

Property Get barcode() As String
    barcode = dbBarCode
End Property

Property Let barcode(barcode As String)
    If validateBarCode(barcode) Then
        dbBarCode = barcode
    Else
        Err.raise 776, "", "Το barcode απέτυχε να επικυρωθεί"
    End If
End Property

Property Get kepcode() As String
    kepcode = dbKEPCode
End Property

Property Let kepcode(kepcode As String)
    dbKEPCode = kepcode
End Property

Property Get firstName() As String
    firstName = dbFirstName
End Property

Property Let firstName(firstName As String)
    dbFirstName = firstName
End Property

Property Get lastName() As String
    lastName = dbLastName
End Property

Property Let lastName(lastName As String)
    dbLastName = lastName
End Property

Property Get licenceCode() As String
    licenceCode = dbLicenceCode
End Property

Property Let licenceCode(licenceCode As String)
    dbLicenceCode = licenceCode
End Property

Property Get printStatus() As String
    printStatus = dbPrintStatus
End Property

Property Let printStatus(printStatus As String)
    dbPrintStatus = printStatus
End Property

Property Get createdDate() As Date
    createdDate = dbCreatedDate
End Property

Property Let createdDate(createdDate As Date)
    dbCreatedDate = createdDate
End Property

Property Get insertedDate() As Date
    insertedDate = dbInsertedDate
End Property

Property Let insertedDate(insertedDate As Date)
    dbInsertedDate = insertedDate
End Property

Property Get updatedDate() As Date
    updatedDate = dbUpdatedDate
End Property

Property Let updatedDate(updatedDate As Date)
    dbUpdatedDate = updatedDate
End Property


Property Get resolvedDate() As Date
    resolvedDate = dbResolvedDate
End Property

Property Let resolvedDate(resolvedDate As Date)
    dbResolvedDate = resolvedDate
End Property


Property Get status() As String
    status = dbStatus
End Property

Property Let status(status As String)
    dbStatus = status
End Property

Private Sub Class_Initialize()
    dbBarCode = ""
    dbKEPCode = -1
    dbFirstName = ""
    dbLastName = ""
    dbLicenceCode = ""
    dbPrintStatus = ""
    dbCreatedDate = 0
    dbInsertedDate = 0
    dbUpdatedDate = 0
    dbResolvedDate = 0
    dbStatus = "N"
End Sub
