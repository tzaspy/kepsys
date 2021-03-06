Attribute VB_Name = "modReportToPDF"
Option Compare Database
Option Explicit

'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 2000 through A2003
' Can be converted to A97 but you must modify the RelationSip window Blob
' structures to the A97 specific versions. You can find these structure declarations
' in the RelationShip Views project on my site.
'
'Copyright: Stephen Lebans - Lebans Holdings 1999 Ltd.


'Distribution:

' Plain and simple you are free to use this source within your own
' applications, whether private or commercial, without cost or obligation, other that keeping
' the copyright notices intact. No public notice of copyright is required.
' You may not resell this source code by itself or as part of a collection.
' You may not post this code or any portion of this code in electronic format.
' The source may only be downloaded from:
' www.lebans.com
'
'Name:      ConvertReportToPDF
'
'Version:   7.85
'
'Purpose:
'
' 1) Export report to Snapshot and then to PDF. Output exact duplicate of a Report to PDF.
'
'ญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญ 
'
'Author:    Stephen Lebans
'
'Email:     Stephen@lebans.com
'
'Web Site:  www.lebans.com
'
'Date:      May 16, 2008, 11:11:11 PM
'
'Dependencies: DynaPDF.dll  StrStorage.dll  clsCommonDialog
'
'Inputs:    See inline Comments for explanation

'Output:    See inline Comments for explanation
'
'Credits:   Anyone who wants some!
'
'BUGS:      Please report any bugs to my email address.
'
'What's Missing:
'           Enhanced Error Handling
'
'How it Works:
' A SnapShot file is created in the normal manner by code like:
'       'Export the selected Report to SnapShot format
'       DoCmd.OutputTo acOutputReport, rptName, "SnapshotFormat(*.snp)", _
'       strPathandFileName
'
' rptName is the desired Report we are working with.
' strPathandFileName can be anything, in this Class it is a
' Temporary FileName and Path created with calls to the
' GetTempPath and GetUniqueFileName API's.
'
' We then pass the FileName to the SetupDecompressOrCopyFile API.
' This will decompress the original SnapShot file into a
' Temporary file with the same name but a "tmp" extension.
'
' The decompressed Temp SnapShot file is then passed to the
' ConvertUncompressedSnapshotToPDF function exposed by StrStorage.DLL.
' The declaration for this call is at the top of this module.
' The function uses the Structured Storage API's to
' open and read the uncompressed Snapshot file. Within this file,
' there is one Enhanced Metafile for each page of the original report.
' Additionally, there is a Header section that contains, among other things,
' a copy of the Report's Printer Devmode structure. We need this to
' determine the page size of the report.

'The StrStorage DLL exposes the function:
'Public Function ConvertUncompressedSnapshotToPDF( _
'UnCompressedSnapShotName As String, _
'OutputPDFname As String = "", _
'Optional CompressionLevel As Long = 0, _
'Optional PasswordOpenAs String = "" _
'Optional PasswordOwner As String = "" _
'Optional PasswordRestrictions as Long = 0, _
'Optional ByVal PDFNoFontEmbedding As Long = 0, _
'Optional ByVal PDFUnicodeFlags As Long = 0 _
') As Boolean

' Now we call the ConvertUncompressedSnapshotToPDF funtion exposed by the StrStorage DLL.
'
'blRet = ConvertUncompressedSnapshot(sFileName as String, sPDFFileName as String)
' Please note that sFileName must include a full valid path(folder) or it will default
' to your My Documents folder. For example  "C:\MyPDFs\MonthlyReport.PDF"

' All other parameters are optional.
'
'Have Fun!
'
'

' Version 7.85
' Please note that the function signatures for both ConvertUncompressedSnapshotToPDF and ConvertReportToPDF
' have changed. An optional parameter has been added to expose the conversion of the
' Metafile to PDF. Flags now include broader support for Unicode and BiDi languages. Finer control
' over how the Metafile is interpreted is exposed as well.

' Added Security/Encryption
' Added/Exposed Flags for Unicode
' Fixed Bug in 11 x 17 paper size
' Fixed Landscape/Portrait bug
'

' Version 7.75
' Added Merge function to merge 2 PDF documents
'
' ******************************************************
#Const ConDebug = 0    ' Set to 1 to force loading of DEBUG StrStorage.DLL
#If (ConDebug = 1) Then

' This is where I screwed up the Font Embedding. Forgot to declare PDFNoFontEmbedding as ByVal!
    Public Declare Function ConvertUncompressedSnapshot Lib "C:\VisualCsource\Debug\StrStorage.dll" _
    (ByVal UnCompressedSnapShotName As String, _
    ByVal OutputPDFname As String, _
    Optional ByVal CompressionLevel As Long = 0, _
    Optional ByVal PasswordOpen As String = "", _
    Optional ByVal PasswordOwner As String = "", _
    Optional ByVal PasswordRestrictions As Long = 0, _
    Optional ByVal PDFNoFontEmbedding As Long = 0, _
    Optional ByVal PDFUnicodeFlags As Long = 0 _
    ) As Boolean


    Public Declare Function DrawTableWindow Lib "C:\VisualCsource\Debug\StrStorage.dll" _
    (ByVal TableName As String, _
    ByVal Fields As String, _
    ByVal NumFields As Long, _
    ByVal Xpos As Double, _
    ByVal Ypos As Double, _
    ByVal Width As Double, _
    ByVal Height As Double _
    ) As Long

    Public Declare Function DrawLine Lib "C:\VisualCsource\Debug\StrStorage.dll" _
    (ByVal Width As Double, _
    ByVal Width1 As Double, _
    ByVal Xpos As Double, _
    ByVal Ypos As Double, _
    ByVal Xpos1 As Double, _
    ByVal Ypos1 As Double, _
    ByVal Attrib As Long _
    ) As Long


    Public Declare Function BeginPDF Lib "C:\VisualCsource\Debug\StrStorage.dll" _
    (ByVal PDFfilename As String, _
    ByVal PageWidth As Long, _
    ByVal PageHeight As Long _
    ) As Long

    Public Declare Function EndPDF Lib "C:\VisualCsource\Debug\StrStorage.dll" _
    () As Long

    Public Declare Function MergePDFDocuments Lib "C:\VisualCsource\Debug\StrStorage.dll" _
    (ByVal PDFMaster As String, _
    ByVal PDFChild As String _
    ) As Boolean




#Else


' This is where I screwed up the Font Embedding. Forgot to declare PDFNoFontEmbedding as ByVal!
Public Declare Function ConvertUncompressedSnapshot Lib "StrStorage.dll" _
    (ByVal UnCompressedSnapShotName As String, _
    ByVal OutputPDFname As String, _
    Optional ByVal CompressionLevel As Long = 0, _
    Optional ByVal PasswordOpen As String = "", _
    Optional ByVal PasswordOwner As String = "", _
    Optional ByVal PasswordRestrictions As Long = 0, _
    Optional ByVal PDFNoFontEmbedding As Long = 0, _
    Optional ByVal PDFUnicodeFlags As Long = 0 _
    ) As Boolean
    
    
    Public Declare Function DrawTableWindow Lib "StrStorage.dll" _
    (ByVal TableName As String, _
    ByVal Fields As String, _
    ByVal NumFields As Long, _
    ByVal Xpos As Double, _
    ByVal Ypos As Double, _
    ByVal Width As Double, _
    ByVal Height As Double _
    ) As Long
    
    Public Declare Function DrawLine Lib "StrStorage.dll" _
    (ByVal Width As Double, _
    ByVal Width1 As Double, _
    ByVal Xpos As Double, _
    ByVal Ypos As Double, _
    ByVal Xpos1 As Double, _
    ByVal Ypos1 As Double, _
    ByVal Attrib As Long _
    ) As Long
    
    
    Public Declare Function BeginPDF Lib "StrStorage.dll" _
    (ByVal PDFfilename As String, _
    ByVal PageWidth As Long, _
    ByVal PageHeight As Long _
    ) As Long
    
    Public Declare Function EndPDF Lib "StrStorage.dll" _
    () As Long
    
    Public Declare Function MergePDFDocuments Lib "StrStorage.dll" _
    (ByVal PDFMaster As String, _
    ByVal PDFChild As String _
    ) As Boolean
    

#End If

' For debugging with Visual C++
'Lib "C:\VisualCsource\Debug\StrStorage.dll"

Private Declare Function ShellExecuteA Lib "shell32.dll" _
(ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" _
Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
(ByVal hLibModule As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" _
Alias "GetTempPathA" (ByVal nBufferLength As Long, _
ByVal lpBuffer As String) As Long

Private Declare Function GetTempFileName _
Lib "kernel32" Alias "GetTempFileNameA" _
(ByVal lpszPath As String, _
ByVal lpPrefixString As String, _
ByVal wUnique As Long, _
ByVal lpTempFileName As String) As Long
 
Private Declare Function SetupDecompressOrCopyFile _
Lib "setupAPI" _
Alias "SetupDecompressOrCopyFileA" ( _
ByVal SourceFileName As String, _
ByVal TargetFileName As String, _
ByVal CompressionType As Integer) As Long

Private Declare Function SetupGetFileCompressionInfo _
Lib "setupAPI" _
Alias "SetupGetFileCompressionInfoA" ( _
ByVal SourceFileName As String, _
TargetFileName As String, _
SourceFileSize As Long, _
DestinationFileSize As Long, _
CompressionType As Integer _
) As Long

 
'Compression types
Private Const FILE_COMPRESSION_NONE = 0
Private Const FILE_COMPRESSION_WINLZA = 1
Private Const FILE_COMPRESSION_MSZIP = 2

Private Const Pathlen = 256
Private Const MaxPath = 256

' Note: I converted the Enums to Constants to allow for use in Access 97.

'Enum TDocumentInfo 'Coming Soon!
 '  diAuthor
 '  diCreator
 '  diKeywords
 '  diProducer
 '  diSubject
 '  diTitle
 '  diCompany
 '  diPDFX_Ver ' GetInDocInfo() only -> The PDF/X version is set by SetPDFVersion()!
 '  diCustom   ' User defined key
'End Enum

'Enum TKeyLen
   Public Const kl40bit = 0    '  40 bit RC4 encryption (Acrobat 3 or higher)
   Public Const kl128bit = 1 ' 128 bit RC4 encryption (Acrobat 5 or higher)
   Public Const kl128bitEx = 2 ' 128 bit RC4 encryption (Acrobat 6 or higher)
'End Enum

'Enum TRestrictions
  Public Const rsDenyNothing = 0
  Public Const rsDenyAll = 3900
  Public Const rsPrint = 4
  Public Const rsModify = 8
  Public Const rsCopyObj = 16
  Public Const rsAddObj = 32
  ' 128 bit encryption only -> these values are ignored if 40 bit encryption is used
  Public Const rsFillInFormFields = 256
  Public Const rsExtractObj = 512
  Public Const rsAssemble = 1024
  Public Const rsPrintHighRes = 2048
  Public Const rsExlMetadata = 4096      ' PDF 1.5 -> can be used with kl128bitEx only
'End Enum



Public Type POINTAPI
   x As Long
   Y As Long
End Type

Public Type RECTL
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type


Public Const AAAlength = 12
Public Const FFFlength = 8
Public Const Padding = 12
Public Const NameLengthMax = 128
' 64 Char MAX for a DAO Table Name * 2 = Unicode

Public Type RelBlob
    Sig As Long
    AAAs(1 To AAAlength) As Byte
    RelWinX1  As Long
    RelWinY1 As Long
    RelWinX2  As Long
    RelWinY2 As Long
    Blank As Long
    FFFs(1 To FFFlength) As Byte
    ClientRectX As Long
    ClientRectY As Long
    'Pad(1 To Padding) As Byte
    ' These next 2 long values represent the Horiz and Vert ScrollBar positions(if any).
    ' These values must be added to the window coordinates stored in this Blob.
    ScrollBarYoffset As Long
    ScrollBarXoffset As Long
    Pad1 As Long
    NumWindows As Long
End Type

Public Type RelWindow
    RelWinX1  As Long
    RelWinY1 As Long
    RelWinX2  As Long
    RelWinY2 As Long
    Junk As Long
    WinName As String * NameLengthMax
    Junk1 As Long
    WinNameMaster As String * NameLengthMax
    'Pad(1 To Padding) As Byte
    Junk2 As Long
End Type

Public Type RelWindowMin
    RelWinX1  As Long
    RelWinY1 As Long
    RelWinX2  As Long
    RelWinY2 As Long
    Column As Long
    WinName As String
End Type

Public Declare Function ScreenToClient Lib "user32" _
(ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function FindWindowEx Lib "user32" Alias _
"FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function apiGetWindow Lib "user32" _
Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function GetWindowRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECTL) As Long

Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

' Create an Information Context
Private Declare Function apiCreateIC Lib "gdi32" Alias "CreateICA" _
(ByVal lpDriverName As String, ByVal lpDeviceName As String, _
ByVal lpOutput As String, lpInitData As Any) As Long

Private Declare Function apiGetDeviceCaps Lib "gdi32" _
Alias "GetDeviceCaps" (ByVal hDC As Long, ByVal nIndex As Long) As Long


Private Declare Function apiDeleteDC Lib "gdi32" _
  Alias "DeleteDC" (ByVal hDC As Long) As Long


' SetWindowPos() Constants
Public Const SWP_SHOWWINDOW = &H40

' GetWindow() Constants
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

'  Device Parameters for GetDeviceCaps()
Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

' ***********************************************
'       Font, DC and TextWidth stuff

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
 
Private Const LF_FACESIZE = 32
 
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" _
(ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
 
Private Declare Function apiCreateFontIndirect Lib "gdi32" Alias _
        "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
 
Private Declare Function apiSelectObject Lib "gdi32" Alias "SelectObject" _
(ByVal hDC As Long, _
ByVal hObject As Long) As Long
 
Private Declare Function apiDeleteObject Lib "gdi32" _
  Alias "DeleteObject" (ByVal hObject As Long) As Long
 
Private Declare Function apiMulDiv Lib "kernel32" Alias "MulDiv" _
(ByVal nNumber As Long, _
ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
 
Private Declare Function apiGetDC Lib "user32" _
  Alias "GetDC" (ByVal hwnd As Long) As Long
 
Private Declare Function apiReleaseDC Lib "user32" _
 Alias "ReleaseDC" (ByVal hwnd As Long, _
 ByVal hDC As Long) As Long
  
Private Declare Function apiDrawText Lib "user32" Alias "DrawTextA" _
(ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function CreateDCbyNum Lib "gdi32" Alias "CreateDCA" _
(ByVal lpDriverName As String, ByVal lpDeviceName As String, _
ByVal lpOutput As Long, ByVal lpInitData As Long) As Long  'DEVMODE) As Long

  
Declare Function GetProfileString Lib "kernel32" _
   Alias "GetProfileStringA" _
  (ByVal lpAppName As String, _
   ByVal lpKeyName As String, _
   ByVal lpDefault As String, _
   ByVal lpReturnedString As String, _
   ByVal nSize As Long) As Long




' CONSTANTS
Private Const TWIPSPERINCH = 1440
' Used to ask System for the Logical pixels/inch in X & Y axis
'Private Const LOGPIXELSY = 90
'Private Const LOGPIXELSX = 88
 
' DrawText() Format Flags
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_EDITCONTROL = &H2000&
Private Const DT_NOCLIP = &H100



' Font stuff
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_PRECIS = 4
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_OUTLINE_PRECIS = 8

Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_MASK = &HF
Private Const CLIP_LH_ANGLES = 16
Private Const CLIP_TT_ALWAYS = 32
Private Const CLIP_EMBEDDED = 128

Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255

' ***********************************************




' Allow user to set FileName instead
' of using API Temp Filename or
' popping File Dialog Window
Private mSaveFileName As String

' Full path and name of uncompressed SnapShot file
Private mUncompressedSnapFile As String

' Name of the Report we ' working with
Private mReportName As String

' Instance returned from LoadLibrary calls
Private hLibDynaPDF As Long
Private hLibStrStorage As Long


Public Function ConvertReportToPDF( _
Optional RptName As String = "", _
Optional SnapshotName As String = "", _
Optional OutputPDFname As String = "", _
Optional ShowSaveFileDialog As Boolean = False, _
Optional StartPDFViewer As Boolean = True, _
Optional CompressionLevel As Long = 0, _
Optional PasswordOpen As String = "", _
Optional PasswordOwner As String = "", _
Optional PasswordRestrictions As Long = 0, _
Optional PDFNoFontEmbedding As Long = 0, _
Optional PDFUnicodeFlags As Long = 0 _
) As Boolean


' RptName is the name of a report contained within this MDB
' SnapshotName is the name of an existing Snapshot file
' OutputPDFname is the name you select for the output PDF file
' ShowSaveFileDialog is a boolean param to specify whether or not to display
' the standard windows File Dialog window to select an exisiting Snapshot file
' CompressionLevel - not hooked up yet
' PasswordOwner  - not hooked up yet
' PasswordOpen - not hooked up yet
' PasswordRestrictions - not hooked up yet
' PDFNoFontEmbedding - Do not Embed fonts in PDF. Set to 1 to stop the
' default process of embedding all fonts in the output PDF. If you are
' using ONLY - any of the standard Windows fonts
' using ONLY - any of the standard 14 Fonts natively supported by the PDF spec
'The 14 Standard Fonts
'All version of Adobe's Acrobat support 14 standard fonts. These fonts are always available
'independent whether they're embedded or not.
'Family name PostScript name Style
'Courier Courier fsNone
'Courier Courier-Bold fsBold
'Courier Courier-Oblique fsItalic
'Courier Courier-BoldOblique fsBold + fsItalic
'Helvetica Helvetica fsNone
'Helvetica Helvetica-Bold fsBold
'Helvetica Helvetica-Oblique fsItalic
'Helvetica Helvetica-BoldOblique fsBold + fsItalic
'Times Times-Roman fsNone
'Times Times-Bold fsBold
'Times Times-Italic fsItalic
'Times Times-BoldItalic fsBold + fsItalic
'Symbol Symbol fsNone, other styles are emulated only
'ZapfDingbats ZapfDingbats fsNone, other styles are emulated only




Dim s As String
Dim blRet As Boolean
' Let's see if the DynaPDF.DLL is available.
blRet = LoadLib()
If blRet = False Then
    ' Cannot find DynaPDF.dll or StrStorage.dll file
    Exit Function
End If

On Error GoTo ERR_CREATSNAP

Dim strPath  As String
Dim strPathandFileName  As String
Dim strEMFUncompressed As String

Dim sOutFile As String
Dim lngRet As Long

' Init our string buffer
strPath = Space(Pathlen)

'Save the ReportName to a local var
mReportName = RptName

' Let's kill any existing Temp SnapShot file
If Len(mUncompressedSnapFile & vbNullString) > 0 Then
    Kill mUncompressedSnapFile
    mUncompressedSnapFile = ""
End If

' If we have been passed the name of a Snapshot file then
' skip the Snapshot creation process below
If Len(SnapshotName & vbNullString) = 0 Then
      
    ' Make sure we were passed a ReportName
    If Len(RptName & vbNullString) = 0 Then
        ' No valid parameters - FAIL AND EXIT!!
        ConvertReportToPDF = ""
        Exit Function
    End If
        
    ' Get the Systems Temp path
    ' Returns Length of path(num characters in path)
    lngRet = GetTempPath(Pathlen, strPath)
    ' Chop off NULLS and trailing "\"
    strPath = Left(strPath, lngRet) & Chr(0)
    
    ' Now need a unique Filename
    ' locked from a previous aborted attemp.
    ' Needs more work!
    strPathandFileName = GetUniqueFilename(strPath, "SNP" & Chr(0), "snp")
    
    ' Export the selected Report to SnapShot format
    DoCmd.OutputTo acOutputReport, RptName, "SnapshotFormat(*.snp)", _
       strPathandFileName
    ' Make sure the process has time to complete
    DoEvents

Else
    strPathandFileName = SnapshotName
 
End If

' Let's decompress into same filename but change type to ".tmp"
'strEMFUncompressed = Mid(strPathandFileName, 1, Len(strPathandFileName) - 3)
'strEMFUncompressed = strEMFUncompressed & "tmp"
Dim sPath As String * 512
lngRet = GetTempPath(512, sPath)

strEMFUncompressed = GetUniqueFilename(sPath, "SNP", "tmp")

lngRet = SetupDecompressOrCopyFile(strPathandFileName, strEMFUncompressed, 0&)

If lngRet <> 0 Then
    Err.raise vbObjectError + 525, "ConvertReportToPDF.SetupDecompressOrCopyFile", _
    "Sorry...cannot Decompress SnapShot File" & vbCrLf & _
    "Please select a different Report to Export"
End If

' Set our uncompressed SnapShot file name var
mUncompressedSnapFile = strEMFUncompressed

' Remember to Cleanup our Temp SnapShot File if we were NOT passed the
' Snapshot file as the optional param
If Len(SnapshotName & vbNullString) = 0 Then
    Kill strPathandFileName
End If


' Do we name output file the same as the input file name
' and simply change the file extension to .PDF or
' do we show the File Save Dialog
If ShowSaveFileDialog = False Then

    ' let's decompress into same filename but change type to ".tmp"
    ' But first let's see if we were passed an output PDF file name
    If Len(OutputPDFname & vbNullString) = 0 Then
        sOutFile = Mid(strPathandFileName, 1, Len(strPathandFileName) - 3)
        sOutFile = sOutFile & "PDF"
    Else
        sOutFile = OutputPDFname
    End If

Else
    ' Call File Save Dialog
    sOutFile = fFileDialog()
    If Len(sOutFile & vbNullString) = 0 Then
        Exit Function
    End If

End If

' Call our function in the StrStorage DLL
' Note the Compression and Password params are not hooked up yet.
blRet = ConvertUncompressedSnapshot(mUncompressedSnapFile, sOutFile, _
CompressionLevel, PasswordOpen, PasswordOwner, PasswordRestrictions, PDFNoFontEmbedding, PDFUnicodeFlags)

If blRet = False Then
Err.raise vbObjectError + 526, "ConvertReportToPDF.ConvertUncompressedSnaphot", _
    "Sorry...damaged SnapShot File" & vbCrLf & _
    "Please select a different Report to Export"
End If

' Do we open new PDF in registered PDF viewer on this system?
If StartPDFViewer = True Then
 ShellExecuteA Application.hWndAccessApp, "open", sOutFile, vbNullString, vbNullString, 1
End If

' Success
ConvertReportToPDF = True


EXIT_CREATESNAP:

' Let's kill any existing Temp SnapShot file
'If Len(mUncompressedSnapFile & vbNullString) > 0 Then
     On Error Resume Next
   Kill mUncompressedSnapFile
    mUncompressedSnapFile = ""
'End If

' If we aready loaded then free the library
If hLibStrStorage <> 0 Then
    hLibStrStorage = FreeLibrary(hLibStrStorage)
End If

If hLibDynaPDF <> 0 Then
    hLibDynaPDF = FreeLibrary(hLibDynaPDF)
End If

Exit Function

ERR_CREATSNAP:
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
mUncompressedSnapFile = ""
ConvertReportToPDF = False
Resume EXIT_CREATESNAP

End Function



Private Function LoadLib() As Boolean
Dim s As String
Dim blRet As Boolean

On Error Resume Next

' *** Please Note ***
' If you are going to process many reports at once then to improve performance you
' should only call LoadLib once.

' May 16/2008
' Always look in the folder where this MDB resides First before checking the System folder.

LoadLib = False

' If we aready loaded then free the library
If hLibDynaPDF <> 0 Then
    hLibDynaPDF = FreeLibrary(hLibDynaPDF)
End If


' Our error string
s = "Sorry...cannot find the DynaPDF.dll file" & vbCrLf
s = s & "Please copy the DynaPDF.dll file into the same folder as this Access MDB or your Windows System32 folder."

' OK Try to load the DLL assuming it is in the same folder as this MDB.
' CurrentDB works with both A97 and A2K or higher
hLibDynaPDF = LoadLibrary(CurrentDBDir() & "DynaPDF.dll")
    
If hLibDynaPDF = 0 Then
    ' OK Try to load the DLL assuming it is in the Window System folder
    hLibDynaPDF = LoadLibrary("DynaPDF.dll")
End If

If hLibDynaPDF = 0 Then
    MsgBox s, vbOKOnly, "MISSING DynaPDF.dll FILE"
    LoadLib = False
    Exit Function
End If



'' ** Commented out for Debugging only - Must be active
'' ***************************************************************************
'
' Load StrStorage.DLL
' If we aready loaded then free the library
If hLibStrStorage <> 0 Then
    hLibStrStorage = FreeLibrary(hLibStrStorage)
End If


' Our error string
s = "Sorry...cannot find the StrStorage.dll file" & vbCrLf
s = s & "Please copy the StrStorage.dll file into the same folder as this Access MDB or your Windows System32 folder."

' OK Try to load the DLL assuming it is in the same folder as this MDB.
' CurrentDB works with both A97 and A2K or higher
hLibStrStorage = LoadLibrary(CurrentDBDir() & "StrStorage.dll")

If hLibStrStorage = 0 Then
    ' OK Try to load the DLL assuming it is in the Window System folder
    hLibStrStorage = LoadLibrary("StrStorage.dll")
End If

If hLibStrStorage = 0 Then
    MsgBox s, vbOKOnly, "MISSING StrStorage.dll FILE"
    LoadLib = False
    Exit Function
End If

' RETURN SUCCESS
LoadLib = True
End Function


'******************** Code Begin ****************
'Code courtesy of
'Terry Kreft & Ken Getz
'
Private Function CurrentDBDir() As String
Dim strDBPath As String
Dim strDBFile As String
    strDBPath = CurrentDb.name
    strDBFile = Dir(strDBPath)
    CurrentDBDir = Left$(strDBPath, Len(strDBPath) - Len(strDBFile))
End Function
'******************** Code End ****************



Private Function GetUniqueFilename(Optional path As String = "", _
Optional Prefix As String = "", _
Optional UseExtension As String = "") _
As String

' originally Posted by Terry Kreft
' to: comp.Databases.ms -Access
' Subject:  Re: Creating Unique filename ??? (Dev code)
' Date: 01/15/2000
' Author: Terry Kreft <terry.kreft@mps.co.uk>

' SL Note: Input strings must be NULL terminated.
' Here it is done by the calling function.

  Dim wUnique As Long
  Dim lpTempFileName As String
  Dim lngRet As Long

  wUnique = 0
  If path = "" Then path = CurDir
  lpTempFileName = String(MaxPath, 0)
  lngRet = GetTempFileName(path, Prefix, _
                            wUnique, lpTempFileName)

  lpTempFileName = Left(lpTempFileName, _
                        InStr(lpTempFileName, Chr(0)) - 1)
  Call Kill(lpTempFileName)
  If Len(UseExtension) > 0 Then
    lpTempFileName = Left(lpTempFileName, Len(lpTempFileName) - 3) & UseExtension
  End If
  GetUniqueFilename = lpTempFileName
End Function


Private Function fFileDialog() As String
' Calls the API File Save Dialog Window
' Returns full path to new File

On Error GoTo Err_fFileDialog

' Call the File Common Dialog Window
Dim clsDialog As Object
Dim strTemp As String
Dim strFname As String

Set clsDialog = New clsCommonDialog

' Fill in our structure
' I'll leave in how to select Gif and Jpeg to
' show you how to build the Filter in case you want
' to use this code in another project.
clsDialog.Filter = "PDF (*.PDF)" & Chr$(0) & "*.PDF" & Chr$(0)
'clsDialog.Filter = clsDialog.Filter & "Gif (*.GIF)" & Chr$(0) & "*.GIF" & Chr$(0)
'clsDialog.Filter = "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)
clsDialog.hDC = 0
clsDialog.MaxFileSize = 256
clsDialog.Max = 256
clsDialog.FileTitle = vbNullString
clsDialog.DialogTitle = "Please Select a path and Enter a Name for the PDF File"
clsDialog.InitDir = vbNullString
clsDialog.DefaultExt = vbNullString

' Display the File Dialog
clsDialog.ShowSave

' See if user clicked Cancel or even selected
' the very same file already selected
strFname = clsDialog.filename
'If Len(strFname & vbNullString) = 0 Then
' Raise the exception
 ' Err.Raise vbObjectError + 513, "clsPrintToFit.fFileDialog", _
  '"Please type in a Name for a New File"
'End If

' Return File Path and Name
fFileDialog = strFname

Exit_fFileDialog:

Err.Clear
Set clsDialog = Nothing
Exit Function

Err_fFileDialog:
fFileDialog = ""
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_fFileDialog

End Function




Public Function fFileDialogSnapshot() As String
' Calls the API File Open Dialog Window
' Returns full path to existing Snapshot File

On Error GoTo Err_fFileDialog

' Call the File Common Dialog Window
Dim clsDialog As Object
Dim strTemp As String
Dim strFname As String

Set clsDialog = New clsCommonDialog

' Fill in our structure
' I'll leave in how to select Gif and Jpeg to
' show you how to build the Filter in case you want
' to use this code in another project.
clsDialog.Filter = "SNAPSHOT (*.SNP)" & Chr$(0) & "*.SNP" & Chr$(0)
'clsDialog.Filter = "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)
clsDialog.hDC = 0
clsDialog.MaxFileSize = 256
clsDialog.Max = 256
clsDialog.FileTitle = vbNullString
clsDialog.DialogTitle = "Please Select a Snapshot File"
clsDialog.InitDir = vbNullString
clsDialog.DefaultExt = vbNullString

' Display the File Dialog
clsDialog.ShowOpen

' See if user clicked Cancel or even selected
' the very same file already selected
strFname = clsDialog.filename
If Len(strFname & vbNullString) = 0 Then
' Do nothing. Add your desired error logic here.
End If

' Return File Path and Name
fFileDialogSnapshot = strFname

Exit_fFileDialog:

Err.Clear
Set clsDialog = Nothing
Exit Function

Err_fFileDialog:
fFileDialogSnapshot = ""
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_fFileDialog

End Function



Public Function fFileDialogSavePDFname() As String
' Calls the API File Open Dialog Window
' Returns full path to existing Snapshot File

On Error GoTo Err_fFileDialog

' Call the File Common Dialog Window
Dim clsDialog As Object
Dim strTemp As String
Dim strFname As String

Set clsDialog = New clsCommonDialog

' Fill in our structure
' I'll leave in how to select Gif and Jpeg to
' show you how to build the Filter in case you want
' to use this code in another project.
clsDialog.Filter = "PDF (*.PDF)" & Chr$(0) & "*.PDF" & Chr$(0)
'clsDialog.Filter = "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)
clsDialog.hDC = 0
clsDialog.MaxFileSize = 256
clsDialog.Max = 256
clsDialog.FileTitle = vbNullString
clsDialog.DialogTitle = "Please Select a name for the PDF File"
clsDialog.InitDir = vbNullString
clsDialog.DefaultExt = vbNullString



' Display the File Dialog
clsDialog.ShowOpen

' See if user clicked Cancel or even selected
' the very same file already selected
strFname = clsDialog.filename
If Len(strFname & vbNullString) = 0 Then
' Do nothing. Add your desired error logic here.
End If

' Return File Path and Name
fFileDialogSavePDFname = strFname

Exit_fFileDialog:

Err.Clear
Set clsDialog = Nothing
Exit Function

Err_fFileDialog:
fFileDialogSavePDFname = ""
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_fFileDialog

End Function





Sub ForeignNameX()

   Dim dbsNorthwind As Database
   Dim relLoop As Relation

   Set dbsNorthwind = CurrentDb() 'OpenDatabase("Northwind.mdb")

   Debug.Print "Relation"
   Debug.Print "        Table - Field"
   Debug.Print "  Primary (One) ";
   Debug.Print ".Table - .Fields(0).Name"
   Debug.Print "  Foreign (Many)  ";
   Debug.Print ".ForeignTable - .Fields(0).ForeignName"

   ' Enumerate the Relations collection of the Northwind
   ' database to report on the property values of
   ' the Relation objects and their Field objects.
   For Each relLoop In dbsNorthwind.Relations
      With relLoop
         Debug.Print
         Debug.Print .name & " Relation"
         Debug.Print "        Table - Field"
         Debug.Print "  Primary (One) ";
         Debug.Print .Table & " - " & .Fields(0).name
         Debug.Print "  Foreign (Many)  ";
         Debug.Print .ForeignTable & " - " & _
            .Fields(0).ForeignName
      End With
   Next relLoop

   dbsNorthwind.Close

End Sub



'Purpose:   Show additional information beside each field in the Print Relationships report.
'Author:    Allen Browne. allen@allenbrowne.com. February 2006.
'Usage:     Set the On Click property of a command button to:
'               =RelReport()
'Method     The Relationships report uses a list box for each table.
'           We open the report, switch to design view, and change the RowSource of each list box,
'           to give more detailed information on each field, by adding the codes below to each field.

' These codes are added to the field names in the Relationships report:

' Field Types:
' ===========
'  A    AutoNumber field (size Long Integer)
'  B    Byte (Number)
'  C    Currency
'  Dbl  Double (Number)
'  Dec  Decimal (Number)
'  Dt   Date/Time
'  Guid Replication ID (Globally Unique IDentifier)
'  Hyp  Hyperlink
'  Int  Integer (Number)
'  L    Long Integer (Number)
'  M    Memo field
'  Ole  OLE Object
'  Sng  Single (Number)
'  T    Text, with number of characters (size)
'  Yn   Yes/No
'  ?    Unknown field type

' Indexes:
' =======
'  P    Primary Key
'  U    Unique Index ('No Duplicates')
'  I    Indexed ('Duplicates Ok')
' Note: Lower case p, u, or i indicates a secondary field in a multi-field index.

' Properties:
' ==========
'  D    Default Value set.
'  R    Required property is Yes
'  V    Validation Rule set.
'  Z    Allow Zero-Length is Yes (Text, Memo and Hyperlink only.)

Public Function RelReport(Optional bSetMarginsAndOrientation As Boolean = True) As Long
'On Error GoTo Err_Handler
    'Purpose:   Main routine. Opens the relationships report with extended field information.
    'Author:    Allen Browne. allen@allenbrowne.com. January 2006.
    'Argument:  bSetMarginsAndOrientation = False to NOT set margins and landscape.
    'Return:    Number of tables adjusted on the Relationships report.
    'Notes:     1. Only tables shown in the Relationships diagram are processed.
    '           2. The table's record count is shown in brackets after the last field.
    '           3. Aliased tables (typically duplicate copies) are not processed.
    '           4. System fields (used for replication) are suppressed.
    '           5. Setting margins and orientation operates only in Access 2002 and later.
    Dim DB As DAO.Database      'This database.
    Dim tdf As DAO.TableDef     'Each table referenced in the Relationships window.
    Dim ctl As Control          'Each control on the report.
    Dim lngKt As Long           'Count of tables processed.
    Dim strReportName As String 'Name of the relationships report
    Dim strMsg As String        'MsgBox message.
    
    'Initialize: Open the Relationships report in design view.
    Set DB = CurrentDb()
    'strReportName = OpenRelReport(strMsg)
    'If strReportName <> vbNullString Then
    
        'Loop through the controls on the report.
        'For Each ctl In Reports(strReportName).Controls
            'If ctl.ControlType = acListBox Then
                'Set the TableDef based on the Caption of the list box's attached label.
                If TdfSetOk(DB, tdf, ctl, strMsg) Then
                    'Change the RowSource to the extended information
                    ctl.RowSource = DescribeFields(tdf)
                    lngKt = lngKt + 1&  'Count the tables processed successfully.
                End If
            'End If
        'Next
        
        'Results
'        If lngKt = 0& Then
'            'Notify the user if the report did not contain the expected controls.
'            strMsg = strMsg & "Diagram of tables not found on report " & strReportName & vbCrLf
'        Else
'            'Preview the report.
'            Reports(strReportName).Section(acFooter).Height = 0&
'            DoCmd.OpenReport strReportName, acViewPreview
'            'Reduce margins and switch to landscape (Access 2002 and later only.)
'            If bSetMarginsAndOrientation Then
'                Call SetMarginsAndOrientation(Reports(strReportName))
'            End If
'        End If
    'End If
    
Exit_Handler:
    'Show any message.
'    If strMsg <> vbNullString Then
'        MsgBox strMsg, vbInformation, "Relationships Report (adjusted)"
'    End If
    'Clean up
    'Set ctl = Nothing
    Set DB = Nothing
    'Return the number of tables processed.
    RelReport = lngKt
    Exit Function

Err_Handler:
    strMsg = strMsg & "RelReport: Error " & Err.Number & ": " & Err.Description & vbCrLf
    Resume Exit_Handler
End Function

Public Function OpenRelReport(strErrMsg As String) As String
On Error GoTo Err_Handler
    'Purpose:   Open the Relationships report.
    'Return:    Name of the report. Zero-length string on failure.
    'Argument:  String to append any error message to.
    Dim iAccessVersion As Integer     'Access version.
    
    iAccessVersion = Int(Val(SysCmd(acSysCmdAccessVer)))
    Select Case iAccessVersion
    Case Is < 9
        strErrMsg = strErrMsg & "Requires Access 2000 or later." & vbCrLf
    Case 9
        RunCommand acCmdRelationships
        SendKeys "%FR", True  'File | Relationships. RunCommand acCmdPrintRelationships is not in A2000.
        RunCommand acCmdDesignView
    Case Is > 9
        RunCommand acCmdRelationships
        RunCommand 483        ' acCmdPrintRelationships
        RunCommand acCmdDesignView
    End Select
    
    'Return the name of the last report opened
    OpenRelReport = Reports(Reports.count - 1&).name

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
    Case 2046&  'Relationships window is already open.
        'A2000 cannot recover, because SendKeys requires focus on the window.
        If iAccessVersion > 9 Then
            Resume Next
        Else
            strErrMsg = strErrMsg & "Close the relationships window, and try again." & vbCrLf
            Resume Exit_Handler
        End If
    Case 2451&, 2191&  'Report not open, or not open in design view.
        strErrMsg = strErrMsg & "The Relationships report must be open in design view." & vbCrLf
        Resume Exit_Handler
    Case Else
        strErrMsg = strErrMsg & "Error " & Err.Number & ": " & Err.Description & vbCrLf
        Resume Exit_Handler
    End Select
End Function

Public Function TdfSetOk(DB As DAO.Database, tdf As DAO.TableDef, ctl As Control, strErrMsg As String) As Boolean
On Error GoTo Err_Handler
    'Purpose:   Set the TableDef passed in, using the name in the Caption in the control's attached label.
    'Return:    True on success. (Fails if the caption is an alias.)
    'Arguments: db = database variable (must already be set).
    '           tdf = the TableDef variable to be set.
    '           ctl = the control that has the name of the table in its attached label.
    '           strMsg = string to append any error messages to.
    Dim strTable As String      'The name of the table.
    
    strTable = ctl.Controls(0).Caption  'Get the name of the table from the attached label's caption.
    Set tdf = DB.TableDefs(strTable)    'Fails if the caption is an alias.
    TdfSetOk = True                     'Return true if it all worked.
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
    Case 3265&  'Item not found in collection. (Table name is an alias.)
        strErrMsg = strErrMsg & "Skipped table " & strTable & vbCrLf
    Case Else
        strErrMsg = strErrMsg & "Error " & Err.Number & ": " & Err.Description & vbCrLf
    End Select
    Resume Exit_Handler
End Function

Public Function DescribeFields(tdf As DAO.TableDef) As String
    'Purpose:   Loop through the fields of the table passed in, to create a string _
                    to use as the RowSource of the list box (Value List type).
    Dim fld As DAO.Field        'Each field of the table.
    Dim strReturn As String     'String to build up and return.
    Const strcSep = ";"         'Separator between items in the list box.
    
    For Each fld In tdf.Fields
        'Skip replication info fields.
        If (fld.Attributes And dbSystemField) = 0& Then
            'strReturn = strReturn & """" & fld.Name & "   "
            strReturn = strReturn & "\le#\FS[8]\FC[0]" & fld.name & " "
            strReturn = strReturn & "\FS[6]\FC[255] - "
            
'\FS[float] // font size
' \FC[ULONG] // font color
            
            'Describe the field type and size.
            Select Case CLng(fld.Type)
                Case dbText
                    strReturn = strReturn & "T" & fld.Size
                    If fld.AllowZeroLength Then
                        strReturn = strReturn & "Z"
                    End If
                Case dbMemo
                    If (fld.Attributes And dbHyperlinkField) <> 0& Then
                        strReturn = strReturn & "Hyp" 'Hyperlink
                    Else
                        strReturn = strReturn & "M"
                    End If
                    If fld.AllowZeroLength Then
                        strReturn = strReturn & "Z"
                    End If
                Case dbLong
                    If (fld.Attributes And dbAutoIncrField) <> 0& Then
                        strReturn = strReturn & "A"   'AutoNumber.
                    Else
                        strReturn = strReturn & "L"
                    End If
                Case dbInteger
                    strReturn = strReturn & "Int"
                Case dbCurrency
                    strReturn = strReturn & "C"
                Case dbDate
                    strReturn = strReturn & "Dt"
                Case dbDouble
                    strReturn = strReturn & "Dbl"
                Case dbSingle
                    strReturn = strReturn & "Sng"
                Case dbByte
                    strReturn = strReturn & "B"
                Case dbDecimal
                    strReturn = strReturn & "Dec"
                Case dbBoolean
                    strReturn = strReturn & "Yn"
                Case dbLongBinary
                    strReturn = strReturn & "Ole"
                Case dbGUID
                    strReturn = strReturn & "Guid"
                Case Else
                    strReturn = strReturn & "?"
            End Select
        
            'Assign codes for the field's crucial properties:
            If fld.Required Then            'Required?
                strReturn = strReturn & "R"
            End If                          'Validation Rule?
            If fld.ValidationRule <> vbNullString Then
                strReturn = strReturn & "V"
            End If                          'Default Value?
            If fld.DefaultValue <> vbNullString Then
                strReturn = strReturn & "D"
            End If
            
            'Indicate if field is indexed.
            strReturn = strReturn & DescribeIndexField(tdf, fld.name) & " " '"""" & strcSep
        End If
    
    strReturn = strReturn & vbCrLf
    Next
    
    DescribeFields = strReturn & "\le#\FS[6]\FC[255]Total Records: " & DCount("*", tdf.name)
    'DescribeFields = strReturn & """     (" & DCount("*", tdf.Name) & ")"""
End Function

Public Function DescribeIndexField(tdf As DAO.TableDef, strField As String) As String
    'Purpose:   Indicate if the field is part of a primary key or unique index.
    'Return:    String containing "P" if primary key, "U" if uniuqe index, "I" if non-unique index.
    '           Lower case letters if secondary field in index. Can have multiple indexes.
    'Arguments: tdf = the TableDef the field belongs to.
    '           strField = name of the field to search the Indexes for.
    Dim ind As DAO.index        'Each index of this table.
    Dim fld As DAO.Field        'Each field of the index
    Dim iCount As Integer
    Dim strReturn As String     'Return string
    
    For Each ind In tdf.Indexes
        iCount = 0
        For Each fld In ind.Fields
            If fld.name = strField Then
                If ind.Primary Then
                    strReturn = strReturn & IIf(iCount = 0, "P", "p")
                ElseIf ind.Unique Then
                    strReturn = strReturn & IIf(iCount = 0, "U", "u")
                Else
                    strReturn = strReturn & IIf(iCount = 0, "I", "i")
                End If
            End If
            iCount = iCount + 1
            
        Next
    Next
    
    DescribeIndexField = strReturn
End Function

Public Function SetMarginsAndOrientation(obj As Object) As Boolean
    'Purpose:   Set half-inch margins, and switch to landscape orientation.
    'Argument:  the report. (Object used, because Report won't compile in early versions.)
    'Return:    True if set.
    'Notes:     1. Applied in Access 2002 and later only.
    '           2. Setting orientation in design view and then opening in preview does not work reliably.
    Const lngcMargin = 720&     'Margin setting in twips (0.5")
    
    'Access 2000 and earlier do not have the Printer object.
    If Int(Val(SysCmd(acSysCmdAccessVer))) >= 10 Then
        With obj.Printer
            .TopMargin = lngcMargin
            .BottomMargin = lngcMargin
            .LeftMargin = lngcMargin
            .RightMargin = lngcMargin
            .Orientation = 2            'acPRORLandscape not available in A2000.
        End With
        
        'Return True if set.
        SetMarginsAndOrientation = True
    End If
End Function



Public Sub GetBlob(rb As RelBlob, rl() As RelWindow, Optional TheUser As String = "", Optional TheMDB As String = "")
' Supply params if using External MDB
' TheMDB must be include full path info
Dim a() As Byte
Dim lTemp As Long
Dim x As Long
'Dim rb As RelBlob
' Module Level instead of private
'Dim rl() As RelWindow
Dim rst As DAO.Recordset
Dim sSQL As String
Dim sSel As String
Dim DB As DAO.Database

' Read the Relationship window BLOB into our array
' Assumes CURRENTUSER is the same user who setup and saved the current Relationship window
' layout for the internal tables. For an External MDB we supply the User!
If Len(TheUser & vbNullString) > 0 Then
    sSel = TheUser
Else
    sSel = CurrentUser
 End If
 
If Len(TheMDB & vbNullString) > 0 Then
    Set DB = OpenDatabase(TheMDB, False, True)
Else
    Set DB = CurrentDb()
 End If
 
 sSQL = "SELECT * FROM MSysObjects WHERE NAME = " & """" & sSel & """"
 Set rst = DB.OpenRecordset(sSQL, dbOpenDynaset, dbReadOnly)

' Get length of BLOB
'lTemp = LenB(rst.Fields("LVExtra"))
lTemp = rst.Fields("LVExtra").FieldSize()

ReDim a(0 To lTemp)
' Copy Blob to our array
a = rst.Fields("LVExtra").GetChunk(0, lTemp)
' Below does not work in A97 so we will use DAO
'a = rst.Fields("LVExtra")
' Free our RecordSet
Set rst = Nothing
DB.Close
Set DB = Nothing

' Fill in our RelBlob header
CopyMem rb, a(0), Len(rb)

' Fill in our TextBox controls
'Me.txtAAAs = rb.AAAs
'Me.txtBlank = rb.Blank
'Me.txtFFFs = rb.FFFs
'Me.txtNumWindows = rb.NumWindows
'Me.txtPadding = rb.Pad
'Me.txtSig = rb.Sig
'Me.txtRelWinX1 = rb.RelWinX1
'Me.txtRelWinX2 = rb.RelWinX2
'Me.txtRelWinY1 = rb.RelWinY1
'Me.txtRelWinY2 = rb.RelWinY2
'Me.txtClientRectY = rb.ClientRectY
'Me.txtClientRectX = rb.ClientRectX

' First 68 Bytes are the Header
' This is followed by (NumWindows + 1) * 284 bytes per record
' Last record seems to be padding
' Let's create an array of our RelWin structures
ReDim rl(0 To rb.NumWindows - 1)
' Fill in our array of structures
For x = 0 To rb.NumWindows - 1
    CopyMem rl(x), a((x * 284) + 68), 284 '(rb.NumWindows + 1) * 128
Next x

End Sub






Public Function RelationsToPDF(ctl As Access.Control) As Boolean
' The Font characteristics of the control passed to this function
' are used for the created PDF document.

Dim rlBlob() As RelWindow
' Copy of RelWindow but with minimal info and no fixed length strings
Dim rl() As RelWindowMin
Dim rlTemp() As RelWindowMin

' The RelationShip window BLOB from the System table
Dim rb As RelBlob

Dim DB As DAO.Database      'This database.
Dim tdf As DAO.TableDef     'Each table referenced in the Relationships window.
Dim tdfForeign As DAO.TableDef
 
Dim SRelTableName As String
Dim SRelFieldName As String
Dim sCodes As String

Dim s As String, sTable As String, sForeign As String
Dim blRet As Boolean
Dim lRet As Long
Dim lTemp As Long

' Current Screen Resolution
Dim Xdpi As Double
Dim Ydpi As Double
Dim lngIC As Long
Dim ConvX As Double
Dim ConvY As Double

Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
Dim X2Max As Long, Y2Max As Long
Dim X1Prev As Long, Y1Prev As Long
X2Max = 0
Y2Max = 0
Dim ctr As Long

' Current Column window width
Dim Width As Long


' Vars to create Font and Measure Text Width and Height
' Structure for DrawText calc
 Dim sRect As RECT
 
 ' Reports Device Context
 Dim hDC As Long
 
 Dim newfont As Long
 ' Handle to our Font Object we created.
 ' We must destroy it before exiting main function

 Dim oldfont As Long
 ' Device COntext's Font we must Select back into the DC
 ' before we exit this function.
 
  ' Logfont struct
 Dim myfont As LOGFONT
 
 ' TextMetric struct
 Dim tm As TEXTMETRIC
 
 ' LineSpacing Amount
 Dim lngLineSpacing As Long
 
 ' Ttemp var
 Dim numLines As Long
 
 ' Temp string var for current printer name
 Dim strName As String
 
 ' Temp vars
 Dim sngTemp1 As Single
 Dim sngTemp2 As Single
 
Dim sText As String
' RelationShip OrdinalPosition Primary table->Field
Dim ReOPp As Integer
' RelationShip OrdinalPosition Foreign table->Field
Dim ReOPf As Integer
Dim fld As DAO.Field

' inner loop counter
Dim i As Integer

Dim rel As Relation


' Let's see if the DynaPDF.DLL is available.
blRet = LoadLib()
If blRet = False Then
    ' Cannot find DynaPDF.dll or StrStorage.dll file
    Exit Function
End If

On Error GoTo ERR_RelationsToPDF


'Initialize: Open the Relationships report in design view.
    Set DB = CurrentDb()
    
sCodes = ""
' Field Types:
' ===========
'  A    AutoNumber field (size Long Integer)
'  B    Byte (Number)
'  C    Currency
'  Dbl  Double (Number)
'  Dec  Decimal (Number)
'  Dt   Date/Time
'  Guid Replication ID (Globally Unique IDentifier)
'  Hyp  Hyperlink
'  Int  Integer (Number)
'  L    Long Integer (Number)
'  M    Memo field
'  Ole  OLE Object
'  Sng  Single (Number)
'  T    Text, with number of characters (size)
'  Yn   Yes/No
'  ?    Unknown field type

' Indexes:
' =======
'  P    Primary Key
'  U    Unique Index ('No Duplicates')
'  I    Indexed ('Duplicates Ok')
' Note: Lower case p, u, or i indicates a secondary field in a multi-field index.

' Properties:
' ==========
'  D    Default Value set.
'  R    Required property is Yes
'  V    Validation Rule set.
'  Z    Allow Zero-Length is Yes (Text, Memo and Hyperlink only.)




' Get current Screen DPI
lngIC = apiCreateIC("DISPLAY", vbNullString, vbNullString, vbNullString)
'If the call to CreateIC didn't fail, then get the Screen X resolution.
If lngIC <> 0 Then
    Xdpi = apiGetDeviceCaps(lngIC, LOGPIXELSX)
    Ydpi = apiGetDeviceCaps(lngIC, LOGPIXELSY)
    'Release the information context.
    apiDeleteDC (lngIC)
Else
    ' Something has gone wrong. Assume an average value.
    Xdpi = 120
    Ydpi = 120
End If


' Create a temp Device Context
' Create our Font and select into the DC
' Get handle to screen Device Context
hDC = apiGetDC(0&)
    
With ctl
     myfont.lfClipPrecision = CLIP_LH_ANGLES
     myfont.lfOutPrecision = OUT_TT_ONLY_PRECIS
     myfont.lfEscapement = 0
     myfont.lfFaceName = .FontName & Chr$(0)
     myfont.lfWeight = .FontWeight
     myfont.lfItalic = .FontItalic
     myfont.lfUnderline = .FontUnderline
     'Must be a negative figure for height or system will return
     'closest match on character cell not glyph
     myfont.lfHeight = (.FontSize / 72) * -Ydpi
     ' Create our temp font
     newfont = apiCreateFontIndirect(myfont)
 End With
 
     If newfont = 0 Then
         Err.raise vbObjectError + 256, "fTextWidthOrHeight", "Cannot Create Font"
     End If

 ' Select the new font into our DC.
 oldfont = apiSelectObject(hDC, newfont)
 
 ' Get TextMetrics. This is required to determine
   ' Text height and the amount of extra spacing between lines.
   lRet = GetTextMetrics(hDC, tm)
 
 ' Our DC is now ready for our calls to:
 ' Calculate our bounding box based on the controls current width
'   lngRet = apiDrawText(hDC, sText, -1, sRect, DT_CALCRECT Or DT_TOP Or _
'   DT_LEFT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_EDITCONTROL Or DT_NOCLIP)
 

' Decode the RelationShip window BLOB
GetBlob rb, rlBlob
' Copy of array of RelWindow structures over to our minimal RelWindow struct
' so we can get rid of unused junk and the fixed length Unicode strings.
ReDim Preserve rl(0 To UBound(rlBlob))



For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        ' rb.ScrollBarXoffset + rb.ScrollBarYoffset will always be either:
        ' 0 - Both Vertical and Horiz ScrollBars are at the Home(0,0) position
        ' a value signifying the offset of the pertinent ScrollBar to be added
        ' to the negative X1,Y1,X2,Y2 coordinates.
        ' We can safely add
        .RelWinX1 = (rlBlob(ctr).RelWinX1) + rb.ScrollBarXoffset
        .RelWinX2 = (rlBlob(ctr).RelWinX2) + rb.ScrollBarXoffset
        .RelWinY1 = (rlBlob(ctr).RelWinY1) + rb.ScrollBarYoffset
        .RelWinY2 = (rlBlob(ctr).RelWinY2) + rb.ScrollBarYoffset
        
'        ' Add a user defined Left Margin
'        Dim LeftMargin As Long
'        LeftMargin = 20
'        .RelWinX1 = .RelWinX1 + LeftMargin
'        .RelWinX2 = .RelWinX2 + LeftMargin

        s = StrConv(rlBlob(ctr).WinName, vbFromUnicode)
        s = Left$(s, InStr(1, s, Chr(0)) - 1)
        .WinName = s
    End With
Next ctr

' We need to perform several modifications to the BLOB data:
'1) Resize the height of each window so that all of the table's fields will be visible.
'   We will have to calculate a new Y1 position after we increate the height of the window.
'
'2) Resize the width of each window so that the Table name and all of the
'   field names will fit. Use a smaller font if the calculated width is larger
'   than our desired max width. Remember, I want to use a fixed width for the
'   columns of our output.
'3)
'
'
' The most difficult issue is to move every window to a column. Basically we want
' to implement a Snap to Grid effect.
' Here is the logic:
' Loop through all windows
' Find the smallest X1 with the smallest Y1
' This becomes our first window
' Start looping again, this time finding the smallest X1 with the smallest Y1
' that is larger than the previous Y1. This logic will ensure we are always working
' down the grid. When we can no longer find any Y1 coords that are larger than previous
' Y1 we are done this column of the grid. We then start over from the top again.
' The logic is further constrained each time in the X direction for each column of
' the grid we are building. X1 must be less than the width of the table at the
' very top of the column we are currently working on. In other words, the starting X1
' position of the next table window below the first one in this column must have a
' starting X1 position less than the X1 + width of the first window in this column.
' If there are two smaller windows under a wide window, and the second window's Y1 meets
' the criteria of being larger than the first small window, we will move this second
' small window directly underneath the first small window. It's the only exception I
' can think of at this pointin developing this logic.

' Ok, we'll need an array and/or a collection to process implement our logic.
' We really only need to store each Table name in final desired column row/order.

' At this point we will not modify the original rl() array.
' Let's try a Collection for now. The key will be the Table Name. We do not need
' to actually store any data as the order of the Key is what is important.
' Basically using the Collection as an odered list.


' First we find the smallest Y1 with the smallest X1.
' This gives us the topmost window in this column
' Next we search for the smallest X1 with a Y1 that is >= to the previous Y1.
' We'll copy our rl() array over to a temp Collection
' so that we can remove entries as we process to
' speed up processing.

' Final Output order of windows
Dim cOut As New Collection
' Temp working Collection
Dim cTmp As New Collection

' Current Column Counter
Dim CurCol As Long
' Need to use/store the array index instead of a single instance of Rel Window structure as VB
' will not accept a structure for the Item param of the Add method of the Collection object.
'Dim r As RelWindowMin
' Copy to temp Collection
For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        cTmp.Add Item:=ctr, Key:=.WinName
    End With
Next ctr

' Non existent seed values
X1Prev = 100000
Y1Prev = 100000
' Find Top and left most window. Smallet X1 and Y1
Dim obj As Variant
Dim sNamePrev As String
' Need to flag when we are at the bottom of a column
'so we can reset seed values.
' No I think we can just keep finding the left most and top most window
' continually until all windows are processed/found.


' SNAP TO GRID
'for i =8 to 80 step 8
' ****************************************************************************************
Dim SpacingInterval As Long

' Add a user defined Left Margin
        Dim LeftMargin As Long
        LeftMargin = 20
        

' Force window to multiple of SpacingInterval value.
' if less than halfway then go backwards to previous multiple.
' if more than or equal to halfway then go ahead to next multiple.
SpacingInterval = 200 ' was 200 sat march 11 at 5:57pm200
'For i = 100 To 200 Step 100
'    SpacingInterval = i '* 25
    For ctr = 0 To rb.NumWindows - 1
        ' Move to multiple of SpacingInterval
        ' Move to 0 if X1 is less than SpacingInterval
        If rl(ctr).RelWinX1 <= SpacingInterval Then
            rl(ctr).RelWinX1 = LeftMargin  '0
        Else
            ' Calculate which column X1 is in.
            lRet = Int(rl(ctr).RelWinX1 / SpacingInterval)
            lTemp = rl(ctr).RelWinX1 - (SpacingInterval * lRet)
            ' Less than half way to next multiple of SpacingInterval
            If lTemp <= SpacingInterval / 2 Then
                ' Move back
                lTemp = -lTemp 'SpacingInterval - lTemp
            Else
                ' More than halfway to next multiple of SpacingInterval
                ' Move forward
                lTemp = SpacingInterval - lTemp
            End If
            ' Update coords
            rl(ctr).RelWinX1 = rl(ctr).RelWinX1 + lTemp
            rl(ctr).RelWinX2 = rl(ctr).RelWinX1 + lTemp
            rl(ctr).Column = Int(rl(ctr).RelWinX1 / SpacingInterval)
        End If
    Next ctr
'Next i

' ****************
' March 11  9:15pm commented out below.
' Its' redundand and alreay done just above.

'' Increase space between SpaceInterval columns
'For ctr = 0 To rb.NumWindows - 1
'    ' Add 300 to each SpacingInterval
'    ' Determine Column #
'    If rl(ctr).RelWinX1 < SpacingInterval Then
'        ' Column = 0
'        lRet = 0
'    Else
'        lRet = Int(rl(ctr).RelWinX1 / SpacingInterval)
'
'    End If
'
'    ' Update Column member
'    rl(ctr).Column = lRet
'    ' Update coords - add min 20 pixels between windows
'    ' ****************************************
'    'comment out below March 11-2006
'    '***********************************************************************
''    lTemp = rl(ctr).RelWinX2 - rl(ctr).RelWinX1
''    rl(ctr).RelWinX1 = rl(ctr).RelWinX1 + (lRet * 20) 'SpacingInterval) '100) 'lTemp
''    rl(ctr).RelWinX2 = rl(ctr).RelWinX1 + lTemp '(lRet * 400)
''
'Next ctr


' Mon - March 6  10:10pm
' commented out

For ctr = 0 To rb.NumWindows - 1

    For Each obj In cTmp

        
        If rl(obj).RelWinX1 = X1Prev Then
        ' Still in same column
            If rl(obj).RelWinY1 < Y1Prev Then
                Y1Prev = rl(obj).RelWinY1
                X1Prev = rl(obj).RelWinX1
                sNamePrev = rl(obj).WinName
                lRet = obj
            End If
        
        Else
            If rl(obj).RelWinX1 < X1Prev Then
        
            'If rl(obj).RelWinY1 = Y1Prev Then
                Y1Prev = rl(obj).RelWinY1
                X1Prev = rl(obj).RelWinX1
                sNamePrev = rl(obj).WinName
                lRet = obj
            
            'ElseIf rl(obj).RelWinY1 <= Y1Prev Then
            
            End If
        
        
        End If
        
    Next obj

    ' Error checking. Processed all windows
    If Len(sNamePrev & vbNullString) = 0 Then Exit For
    ' Update Column member
    
    
    ' Save off this window in our ordered list
    cOut.Add Item:=lRet, Key:=sNamePrev
    ' Remove this item from the temp work collection
    cTmp.Remove sNamePrev
    ' Reset to non existent seed values
    X1Prev = 100000
    Y1Prev = 100000
    sNamePrev = 0

Next ctr

' When we get to here all windows should have been processed
' and our temp work collection should have been emptied.
'

' Mon - March 6  10:10pm
' commented out
'X1 = 0
'Y1 = 0
'
'
' Make a working copy
ReDim rlTemp(0 To UBound(rl))
rlTemp = rl

' What we want to do is copy, in order, to the rl() array, via the Collection Item prop
' from the rlTemp() array. This will put the windows in order from the
' top leftmost to the bottom right most. We need to do this so we can adjust/increase
' the height of each Table windows so that all of the fields will be visible.
ctr = 0
For Each obj In cOut
    With rl(ctr)
        .RelWinX1 = rlTemp(obj).RelWinX1
        .RelWinY1 = rlTemp(obj).RelWinY1
        .RelWinX2 = rlTemp(obj).RelWinX2
        .RelWinY2 = rlTemp(obj).RelWinY2
        .WinName = rlTemp(obj).WinName
        .Column = rlTemp(obj).Column
        ctr = ctr + 1
    End With
Next

Dim MaxDocCharWidth As Long
Dim MaxDocCharHeight As Long
' Width of max documentation characters
' Since we are using a 10 point font to calc width but really
' outputting 8 point with a 10 point leading then we do not
' need any extra char spacing.
 sText = "XXXXg"
With sRect
    .Left = 0
    .Top = 0
    .Bottom = 0
    ' Single line TextWidth
    .Right = 32000
End With

   lRet = apiDrawText(hDC, sText, -1, sRect, DT_CALCRECT Or DT_TOP Or _
            DT_LEFT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_EDITCONTROL Or DT_NOCLIP)

    MaxDocCharWidth = sRect.Right
    ' Allow for 14 pt header and 10 point leading
    MaxDocCharHeight = sRect.Bottom ' * 2
    


' Since the DyanPDF library will automatically wrap text to the next line
' we have to make sure that the Table name, the field names and the extra
' field documenting characters fit one single lines. Otherwise our logic
' to calculate the beginning and ending points of the Join lines will not be accurate.
' There is an issue of overlap though in the X dimension when I increase the width
' of the table window. This is easy to solve in the Y dimension but tougher in the X direction.
' I may have to set a fixed width for all windows to solve this issue.

' X2Max holds widest Table or Field name.
' Loop through all of the table widths and adjust

' Add extra space in width to allow for documenting chars.

' Let's increase the Width of each Table window so that all fields are visible.
' Perhaps we should modify the rl structure to hold max width required to
' ensure the table and field names are visible. No let's use a collection object instead.
' No we will modify as we go - no need to store this value.
X2Max = 0
Y2Max = 0
Dim bHeader As Boolean

For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        ' Call our function to calc height
        SRelTableName = .WinName '(.WinName) 'StrConv(.WinName, vbFromUnicode)
        s = Right$(SRelTableName, 3)
        lRet = InStr(s, "_")
        If lRet = 1 Or lRet = 2 Then
            SRelTableName = Mid$(SRelTableName, 1, Len(SRelTableName) - (4 - lRet))
        End If
        ' DO NOT need to process this clone/copy - just process main table.
        ' No we cannot store the original Table window's Max width as it may not have been
        ' processed at this point.
        'If lRet = 0 Then
        
        
            Set tdf = DB.TableDefs(SRelTableName) '.WinName)
            If Not tdf Is Nothing Then
                'Calc width of Table name and all Field Names
                ' Set width of Table window to max width
                sText = tdf.name
                With sRect
                    .Left = 0
                    .Top = 0
                    .Bottom = 0
                    ' Single line TextWidth
                    .Right = 32000
                End With
                
                   lRet = apiDrawText(hDC, sText, -1, sRect, DT_CALCRECT Or DT_TOP Or _
                            DT_LEFT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_EDITCONTROL Or DT_NOCLIP)
    
                    X2Max = sRect.Right
                    bHeader = True
                    
                For Each fld In tdf.Fields
                   sText = fld.name
                    With sRect
                        .Left = 0
                        .Top = 0
                        .Bottom = 0
                        ' Single line TextWidth
                        .Right = 32000
                    End With
                
                   lRet = apiDrawText(hDC, sText, -1, sRect, DT_CALCRECT Or DT_TOP Or _
                            DT_LEFT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_EDITCONTROL Or DT_NOCLIP)
    
                    If sRect.Right > X2Max Then
                    X2Max = sRect.Right
                    bHeader = False
                    End If
                Next
                
                ' ***********************************************************
                ' Make this a user optional param
                ' Resize to width ALL WINDOWS
                ' Get current width of this window. If it is less than X2Max then adjust.
                'If .RelWinX2 - .RelWinX1 < X2Max + MaxDocCharWidth Then
                ' May 11/2008
                ' If Table window width was sized to fit field name + doc chars then
                ' somehow it was not increased enough in width
                ' Short term solution - set all windows to my calculated width
                ' but COME BACK and figure out why!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    
                 ' May 16/2008
            ' Add logic to calc sFields width and use this calculated value
            ' to determine actual table max width. Remember to take the Table name
            ' into consideration in calculating max width
                    'If bHeader = True Then
                     .RelWinX2 = .RelWinX1 + X2Max + MaxDocCharWidth ' + 16
                    'Else
                     '   .RelWinX2 = .RelWinX1 + X2Max + MaxDocCharWidth
                    'End If
                Set fld = Nothing
                X2Max = 0
    
            End If
        'End If
    
    End With
Next ctr
Set tdf = Nothing








' *****
' Adjust Height of all Relationship Table windows.
' *****

' Let's increase the Width of each Table window so that all fields are visible.
' Perhaps we should modify the rl structure to hold max width required to
' ensure the table and field names are visible. No let's use a collection object instead.
' No we will modify as we go - no need to store this value.
X2Max = 0
Y2Max = 0
For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        ' Call our function to calc height
        SRelTableName = .WinName '(.WinName) 'StrConv(.WinName, vbFromUnicode)
        s = Right$(SRelTableName, 3)
        lRet = InStr(s, "_")
        If lRet = 1 Or lRet = 2 Then
            SRelTableName = Mid$(SRelTableName, 1, Len(SRelTableName) - (4 - lRet))
        End If
        ' DO NOT need to process this clone/copy - just process main table.
        ' No we cannot store the original Table window's Max width as it may not have been
        ' processed at this point.
        'If lRet = 0 Then
        
        ' Build our string starting with Relationship Table window name
        sText = SRelTableName & vbCrLf
        
            Set tdf = DB.TableDefs(SRelTableName) '.WinName)
            If Not tdf Is Nothing Then
                ' Add individual Field names
                                  
                For Each fld In tdf.Fields
                   sText = sText & fld.name & vbCrLf
                Next
                    
                  
                    With sRect
                        .Left = 0
                        .Top = 0
                        .Bottom = 0
                        ' Single line TextWidth
                        .Right = 30000 'rl(ctr).RelWinX2 - rl(ctr).RelWinX1
                    End With
                
                   lRet = apiDrawText(hDC, sText, -1, sRect, DT_CALCRECT Or DT_TOP Or _
                            DT_LEFT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_EDITCONTROL Or DT_NOCLIP)
                Y2Max = sRect.Bottom
                                   
                ' Get current height of this window. If it is less than calc Height then adjust.
                ' We also need to leave room for an extra row to allow for the
                ' Total Recs: line we output
                ' May 11/2008 BUG Fix
                ' If Table window was sized to display all of its fields then
                ' somehow it was being increased too much in height
                ' Short term solution - set all windows to my calculated height
                ' but COME BACK and figufre out why!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                'If .RelWinY2 - .RelWinY1 < Y2Max + MaxDocCharHeight Then
                    .RelWinY2 = .RelWinY1 + Y2Max + MaxDocCharHeight
                'End If
                Set fld = Nothing
    
            End If
        'End If
    ' ***
    '***********************************************
    ' March 11/2006
    ' Add code here to set each window to the start of the column.
    ' Allow user to specify MinColumnSpacing
    Dim MinColumnSpacing As Long
    
    MinColumnSpacing = 40
    ' SpacingInterval contains relative offset
    
    
    
    
    End With
Next ctr
Set tdf = Nothing
Set fld = Nothing









' March 6 -2006 10:18pm
' COMMENTED out below


' *****
' Adjust Starting X1 ANd Y1 of all Relationship Table windows.
' *****

' Let's increase the X1 starting X position of each Table window in order to
' increase the spacing between each table. We do this because overlapping conditions
' are created when we previously increased the width of each Table window.
' To keep this simple, we allow the user to specify a fixed amount for the
' spacing value.
' Since our array of Rel() structures is ordered from top leftmost to
' bottom right most we can basically process the windows in a column by column order.
'
' Because the spacing has to be cumulative per increasing column position, we multiply
' the user's desired spacing value by the current column count(zero indexed).

' Let's increase the Y1 starting Y position of each Table window in order to
' ensure that Table Windows do not overlap. We do this because overlapping conditions
' are created when we previously increased the Height of each Table window in order to
' ensure that all fields in the table window are visible.


Dim ctrCol As Long
'Dim Y1Prev As Long,
Dim Y2Prev As Long
Dim Y2PrevOrig As Long, Y1PrevOrig As Long
Dim VerticalWindowSpacing As Long

VerticalWindowSpacing = 14
Y1Prev = 0
Y2Prev = 0
X2Max = 0
Y2Max = 0
ctrCol = 0

Y2PrevOrig = 0
Y1PrevOrig = 99999999


For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        ' Modify Y1 first
        ' First window in the array is the topmost - leftmost window
        ' Determine if we are still in the current column.
        ' If the Y1 of this window is Greater than the Y1 of the
        ' previous window then we are still in the same column.
        ' Do need to code exception to handle when this current window
        ' is in the next column because even though the this Y1 is greater
        ' than previous Y1, X1 actually places this window in the next column.(I think):-)
        If .RelWinY1 > Y1PrevOrig Then
            ' We're still in the same column
            ' Store Y1
            Y1PrevOrig = .RelWinY1
            ' Are we overlapping the previous window in this column.
            If (.RelWinY1 < Y2Prev + VerticalWindowSpacing) And Y2Prev <> 0 Then
                ' Reposition to avoid overlap - calc resize first
                .RelWinY2 = (.RelWinY2 - .RelWinY1) + Y2Prev + VerticalWindowSpacing
                .RelWinY1 = Y2Prev + VerticalWindowSpacing

'                Y2Prev = .RelWinY2
'                Y1Prev = .RelWinY1
                
            'Else
            
            End If
            Y2Prev = .RelWinY2
                Y1Prev = .RelWinY1
            
        Else
            ' We're in the next column. Do not resize as it is the top most
            ' window in this column. Reset seeds to non existent values.
            ' Next Column
            ctrCol = ctrCol + 1
            Y2Prev = .RelWinY2
                Y1Prev = .RelWinY1
                Y1PrevOrig = .RelWinY1 '0
            ' Since we are at top of column no need to reposition

        End If

    End With
Next ctr




' Set absolute position for start of each column.
' Find Max Width of all windows in each column to calc ColumnWidth
' Storage for column Widths
Dim aColWidths() As Long

Dim lNumColumns As Long

' Get Total number of columns
For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        If lNumColumns < .Column Then lNumColumns = .Column
    End With
Next ctr
    

ReDim aColWidths(0 To lNumColumns)
Dim Gutter As Long
Gutter = 20

' Find largest window width in each column and
' store this value in our column width array.
For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        If (.RelWinX2 - .RelWinX1) > aColWidths(.Column) Then
            aColWidths(.Column) = (.RelWinX2 - .RelWinX1)
        End If
    End With
Next ctr


' Set X1 for every table window to the calc start of the column.
' *****************************
' Here we can set the Left Margin
For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        ' Column starting position =
        ' column widths for all previous columns plus
        ' column spacing value
        lTemp = 0
        For i = 0 To .Column - 1
            lTemp = lTemp + aColWidths(i)
            lTemp = lTemp + Gutter
        Next i
            .RelWinX2 = (.RelWinX2 - .RelWinX1) + lTemp
            .RelWinX1 = IIf(lTemp = 0, LeftMargin, lTemp)

    End With
Next ctr






' Loop through all Relationship Table windows to get
' the largest X2 and Y2 coordinates.
' Modify the starting Y1 coordinate for all Table Windows
' to allow for 1 inch Header section.
' Finally convert Window coords to 72 PPI used by the DynaPDF library
'

X2Max = 0
Y2Max = 0
For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        .RelWinX1 = (.RelWinX1 / Xdpi) * 72
        .RelWinX2 = ((.RelWinX2 / Xdpi) * 72) ' + 16
        .RelWinY1 = ((.RelWinY1 / Ydpi) * 72) '+ 16 ' Space for header section
        .RelWinY2 = ((.RelWinY2 / Ydpi) * 72) '+ 6 ' Space for header section
    End With

    X2 = rl(ctr).RelWinX2
    Y2 = rl(ctr).RelWinY2
    If X2Max < X2 Then X2Max = X2
    If Y2Max < Y2 Then Y2Max = Y2
Next ctr

ctr = 0

' 1) We will have to widen each window to accomodate Allen Browne's
' documentation character symbols.
'
' 2) To make it simpler to create the windows in the PDF document
' I want to make each window the same width.

' In the next release I'll add a param to this function to allow
' the user to specify the desired width.

' So I'll need a function or functions in the StrStorage DLL

Dim sFields As String
Dim sPDF As String

sPDF = "C:\sourcecode\ReportToPDF\Relations.pdf"
' Should calc string width of Allen's Documentation Characters
' instead of using the fixed value of 16 Points.
' We also need to allow space for a Header or Footer
lRet = BeginPDF(sPDF, X2Max + 32, Y2Max + 32)



'GoTo HHH



' The first time through we will just gather the necessary info
' to allow us to draw the Relationship Join lines.
' We will need to store
' Table Name(to index into the Relation object)
' Table Ypos
' Field Name(to index into Relation object)
' Field Pos - 1 to num fields
For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        On Error Resume Next
        SRelTableName = .WinName
        Set tdf = Nothing
        ' We don't have to remove _1(_x) from end of WinName because the Relation object
        ' only stores relations under the original table name - Customers not Customers_1.
        ' We know it is a Clone/Copy of the Table when the Table and ForeightTable props
        ' are the same. We can then examine the Name prop, specifically the last char
        ' to tell what instance of the clone/copy we are working with.
        ' First instance is Customers_1 then Customers_2 etc. But this logic does not
        ' carry over to the Name prop of the Relation object.
        ' Customers_1 = CustomersCustomers
        ' Customers_2 = CustomersCustomers_1
        'etc
'        s = Right$(SRelTableName, 3)
'        lRet = InStr(s, "_")
'        If lRet = 1 Or lRet = 2 Then SRelTableName = Mid$(SRelTableName, 1, Len(SRelTableName) - (4 - lRet))
'
        
        Set tdf = DB.TableDefs(SRelTableName)
        If Not tdf Is Nothing Then
            'Get Field Name + Documenting info
            'sFields = DescribeFields(tdf)
            'lngKt = lngKt + 1&  'Count the tables processed successfully.
            ' See if there are any matching Relation entries.
            ' If there are then store the required information
            ' to allow us to draw the Relationship table/field Lines
            For Each rel In DB.Relations
                If rel.Table = .WinName Then ' Then
                    ' There is a matching relation for this Relationship Window Table.
                    ' We need to find the rel field for this entry and
                    ' store the absolute position of this field in the table->fields collection.
                    ' We cannot draw the line now as the matching rel table may not
                    ' have been drawn yet. Remember, we must draw all of the Relationship lines
                    ' BEFORE we draw the Table windows as the Rel Lines must appear
                    ' behind the Table windows.
                    ' We can use the OrdinalPosition property of the field but it must be in
                    ' the Table object not the Relationship object. OrdinalPosition is a
                    ' zero based prop.
                    ' We then will use the same logic as above to determine the absolute
                    ' position of the matching ForeignName Field for the ForeignTable
                    ' component of this Relationship.
                    ' Do we need to store this information at all? I mean since the
                    ' OrdinalPosition prop is available for Relationship fields we do not
                    ' need to store it. Also since the absolute position of each
                    ' Relationship Table window is known/calculated why can't I simply
                    ' loop through the Relations collection and render the Lines when the
                    ' Table prop of the Relation object has a matching entry in the
                    ' Relationship window BLOB data?
                    
                    
                    ' Draw the Line for this Relationship
                    ' Get the Ordinal Position of the Primary and ForeignTable fields
                    Set fld = rel.Fields(0)
                    
                    ReOPp = tdf.Fields(fld.name).OrdinalPosition
                    
                    ' Check if ForeignTable prop is a Clone/Copy
                    lRet = 0
                    If rel.Table = rel.ForeignTable Then
                        ' Determine which copy(_x) this one is
                        If Len(rel.Table) * 2 = Len(rel.name) Then
                            s = rel.ForeignTable & "_" & 1
                            lRet = 1
                        Else
                            ' Grab last character of Name prop. This logic will
                            ' only support to a max of 9 clones/copies
                            s = Right$(rel.name, 1)
                            s = rel.ForeignTable & "_" & Val(s) + 1
                            lRet = 1
                        End If
                    
                    End If
                    
                    Set tdfForeign = DB.TableDefs(rel.ForeignTable)
                    ReOPf = tdfForeign.Fields(fld.ForeignName).OrdinalPosition + 1
                    
                    ' Calc the start and ending X,Y cordinates for the
                    ' Relationship Line we are going to draw.
                    X1 = .RelWinX1 '(.RelWinX1 / Xdpi) * 72
                    Y1 = .RelWinY1 '(.RelWinY1 / Ydpi) * 72
                    ' Now we need to add an offset to Y1 to bring us down to
                    ' the row containing the relationship field. Since the
                    ' OrdinalPosition index is zero based we don't have to add 1
                    ' to cover the fact that we output a row first containing
                    ' the Table name. 10 pts is the row spacing.
                    Y1 = Y1 + (IIf(ReOPp = 0, 1, ReOPp) * 10)
                    ' Now we need to find X1 and Y1 for the Foreign Table
                    ' Find it in the Rel BLOB data.
                    ' Need to allow logic to determine on which side(left or right)
                    ' we want the Relationship Line to start from.
                    ' If the left edge of the Foreign table window is <= to the
                    ' center of the Primary Table then the Joining line will originate from
                    ' the left side of the Primary table. Otherwise, it will originate
                    ' from the right side of the Primary table
                    If lRet = 0 Then
                        s = rel.ForeignTable
                    End If
                    
                    For i = 0 To rb.NumWindows - 1
                        If rl(i).WinName = s Then 'rel.ForeignTable Then
                        'If Trim(rl(i).WinName) = rel.ForeignTable Then
                            X2 = (rl(i).RelWinX1)
                            Y2 = (rl(i).RelWinY1)
                            Y2 = Y2 + (IIf(ReOPf = 0, 1, ReOPf) * 10)
                            ' Which side of Primary table does the Join line
                            ' originate from left/right.
                            ' Handled in StrStorage DLL by DrawLine function
                            lRet = DrawLine(.RelWinX2 - .RelWinX1, rl(i).RelWinX2 - rl(i).RelWinX1, _
                            X1, Y1, X2, Y2, lRet)
                        End If
                    Next i
                    
                    Set fld = Nothing
                    Set tdfForeign = Nothing
                End If
            Next
        
        End If
        
    End With
    Set tdf = Nothing
Next ctr


HHH:


' Output Header before Table Windows
' Pass 0 in NumFields param to signal this is Header info.
' Pass desired Header info in TableNames param.
' Coordinate params will be used to position Header
' We have modified the starting Y1 coordinate for all Table Windows
' to allow for 1 inch Header section.


'SRelTableName = "RelationShip Report:" & Date & Chr(0) 'vbCrLf
's = CurrentDb().Name & Chr(0)
'lRet = DrawTableWindow(SRelTableName, s, 0, _
'         10, 10, 400, -1)


' Main loop to actually draw each Relationship Table window
' and the Tables component fields.
For ctr = 0 To rb.NumWindows - 1
    With rl(ctr)
        On Error Resume Next
        SRelTableName = .WinName '(.WinName) 'StrConv(.WinName, vbFromUnicode)
        s = Right$(SRelTableName, 3)
        lRet = InStr(s, "_")
        If lRet = 1 Or lRet = 2 Then SRelTableName = Mid$(SRelTableName, 1, Len(SRelTableName) - (4 - lRet))
        Set tdf = DB.TableDefs(SRelTableName)
        If Not tdf Is Nothing Then
            'Get Field Name + Documenting info
            ' May 16/2008
            ' Add logic to calc sFields width and use this calculated value
            ' to determine actual table max width. Remember to take the Table name
            ' into consideration in calculating max width
            sFields = DescribeFields(tdf)
            'lngKt = lngKt + 1&  'Count the tables processed successfully.
        End If
        
       ' lRet = DrawTableWindow(SRelTableName,  sFields, rb.NumWindows, _
        ' (.RelWinX1 / Xdpi) * 72, (.RelWinY1 / Ydpi) * 72, (((.RelWinX2 - .RelWinX1) / Xdpi) * 72) + 12, ((.RelWinY2 - .RelWinY1) / Ydpi) * 72)
    lRet = DrawTableWindow(.WinName, sFields, rb.NumWindows, _
         .RelWinX1, .RelWinY1, (.RelWinX2 - .RelWinX1), (.RelWinY2 - .RelWinY1))
    End With
    Set tdf = Nothing
Next ctr


' Do we open new PDF in registered PDF viewer on this system?
'If StartPDFViewer = True Then
 ShellExecuteA Application.hWndAccessApp, "open", sPDF, vbNullString, vbNullString, 1
'End If

On Error GoTo 0

lRet = EndPDF

RelationsToPDF = True

EXIT_RelationsToPDF:

Set DB = Nothing
Set tdf = Nothing
Set fld = Nothing
Set rel = Nothing

' If we aready loaded then free the library
If hLibStrStorage <> 0 Then
    hLibStrStorage = FreeLibrary(hLibStrStorage)
End If

If hLibDynaPDF <> 0 Then
    hLibDynaPDF = FreeLibrary(hLibDynaPDF)
End If


' Cleanup
   lRet = apiSelectObject(hDC, oldfont)
   ' Delete the Font we created
   apiDeleteObject (newfont)
   
    ' Release the handle to the Screen's DC
    lRet = apiReleaseDC(0&, hDC)
  
Exit Function

ERR_RelationsToPDF:
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number

RelationsToPDF = False
Resume EXIT_RelationsToPDF




End Function

