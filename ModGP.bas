Attribute VB_Name = "ModGP"
Global Curtool As Integer
Global FBcancel As Boolean
Global BWcancel As Boolean
Global RScancel As Boolean
Global CWcancel As Boolean
Global saveCancel As Boolean
Global RulersVis As Boolean
Global finalclose As Boolean
Global Newcancel As Boolean
Global NewHeight As Long
Global NewWidth As Long
Global NewBGcol As Long
Global CurBGindex As Integer
Global curfilter As Integer
Global curfilterlevel As Integer
Global curborder As Integer
Global curborderlevel2 As Integer
Global curborderlevel3 As Integer
Global borderwidth As Integer
Global framewidth As Integer
Global chBGcolor As Long
Global outline As Boolean
Global inline As Boolean
Global AspectRatio As Double
Global NewScaleHeight As Long
Global NewScaleWidth As Long
Global Savepath As String
Global curfile As String
Global ReadLong As Boolean
Global ReadHex As Boolean
Global ReadRgb As Boolean
Global maxcolchose As Integer
Global colslocked As Boolean
Global pastingasnew As Boolean
Global startVSval As Double
Global startHSval As Double
Global freeselection As Boolean
Global dontusePicBU As Boolean
Global Masterpasting As Boolean
Global sfilename As String
Global Cachesize As Long
Public safesavename As String
Private Type POINTAPI
    x As Long
    y As Long
End Type
Public Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function BmpToJpeg Lib "BBJPeg.dll" (ByVal bmpFileName As String, ByVal JpegFilename As String, ByVal Quality As Integer) As Integer
Declare Function TWAIN_AcquireToFilename Lib "EZTW32.DLL" (ByVal hwndApp%, ByVal bmpFileName$) As Integer
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long
Declare Function TWAIN_IsAvailable Lib "EZTW32.DLL" () As Long
Declare Function TWAIN_EasyVersion Lib "EZTW32.DLL" () As Long
Public ret As String
Public Retlen As String
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long
Private Type SHFILEOPSTRUCT
    Hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
    End Type
    Private Const FO_MOVE = &H1
    Private Const FO_COPY = &H2
    Private Const FOF_SILENT = &H4
    Private Const FOF_RENAMEONCOLLISION = &H8
    Private Const FOF_NOCONFIRMATION = &H10
    Private Const FOF_SIMPLEPROGRESS = &H100
    Private Const FOF_ALLOWUNDO = &H40
    Private Const FO_DELETE = &H3
   Private Const FO_RENAME = &H4&
Dim flag As Integer
Dim fred As Integer
Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim FO_FUNC As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const FLOAT = 1, SINK = 0
Public Declare Sub SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Public Const GWL_HWNDPARENT = (-8)
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public ImageCount As Integer
Public PenTip As Integer
Public Shapetype As Integer
Public PenDrawWidth As Integer
Public PenTipWidth As Integer
Global startwidth As Integer
Global singlefactor As Double
Global NoSizeonStart As Boolean


Public Function temppath() As String
    Dim sBuffer As String
    Dim lRet As Long
    sBuffer = String$(255, vbNullChar)
    lRet = GetTempPath(255, sBuffer)
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    temppath = sBuffer
    If Right(temppath, 1) = "\" Then temppath = Left(temppath, Len(temppath) - 1)
End Function
'Used to ensure a unique filename and thus
'avoid overwriting
Public Function SafeSave(path As String) As String
Dim mPath As String, mTemp As String, mFile As String, mExt As String, m As Integer
On Error Resume Next
mPath = Mid$(path, 1, InStrRev(path, "\"))
mname = Mid$(path, InStrRev(path, "\") + 1)
mFile = Left(Mid$(mname, 1, InStrRev(mname, ".")), Len(Mid$(mname, 1, InStrRev(mname, "."))) - 1) 'File only - no extension
If mFile = "" Then mFile = mname
mExt = Mid$(mname, InStrRev(mname, "."))
mTemp = ""
Do
    If Not FileExists(mPath + mFile + mTemp + mExt) Then
        SafeSave = mPath + mFile + mTemp + mExt
        safesavename = mFile + mTemp + mExt
        Exit Do
    End If
    m = m + 1
    mTemp = Right(Str(m), Len(Str(m)) - 1)
Loop
End Function
Function FileExists(ByVal fileName As String) As Integer
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(fileName)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
                End
            End If
    End Select
End Function
Private Function PerformShellAction(sSource As String, sDestination As String) As Long
      On Error Resume Next
    sSource = sSource & Chr$(0) & Chr$(0)
     FOF_FLAGS = BuildBrowseFlags()
     With SHFileOp
        .wFunc = FO_FUNC
        .pFrom = sSource
        .pTo = sDestination
        .fFlags = FOF_FLAGS
    End With
    PerformShellAction = SHFileOperation(SHFileOp)
End Function
Public Sub RenameFile(fileName As String, Target As String)
    On Error Resume Next
    Dim FileStruct As SHFILEOPSTRUCT
    Dim P As Boolean
    Dim x As Long
    Dim strNoConfirm As Integer, strNoConfirmMakeDir As Integer, strRenameOnCollision As Integer
    Dim strSilent As Integer, strSimpleProgress As Integer
         FileStruct.pFrom = fileName
        FileStruct.pTo = Target
        FileStruct.wFunc = FO_RENAME
        x = SHFileOperation(FileStruct)
  End Sub
Private Function BuildBrowseFlags() As Long
     On Error Resume Next
     flag = flag Or FOF_SILENT
        flag = flag Or FOF_NOCONFIRMATION
   If fred = 1 Then flag = flag Or FOF_RENAMEONCOLLISION
 BuildBrowseFlags = flag
End Function
Private Sub ShellDeleteOne(sfile As String)
     On Error Resume Next
   Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim R As Long
    FOF_FLAGS = BuildBrowseFlags()
sfile = sfile & Chr$(0)
With SHFileOp
  .wFunc = FO_DELETE
  .pFrom = sfile
  .fFlags = FOF_FLAGS
End With
R = SHFileOperation(SHFileOp)
End Sub
Public Sub moveme(Source As String, dest As String)
FO_FUNC = 1
Call PerformShellAction(Source, dest)
End Sub
Public Sub CopyMe(Source As String, dest As String)
FO_FUNC = 2
Call PerformShellAction(Source, dest)
End Sub
Public Sub deleteme(path As String)
ShellDeleteOne (path)
End Sub
Public Function FileOnly(ByVal FilePath As String) As String
    FileOnly = Mid$(FilePath, InStrRev(FilePath, "\") + 1)
End Function
Public Function ExtOnly(ByVal FilePath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(FilePath, InStrRev(FilePath, ".") + 1)
If dot = True Then ExtOnly = "." + ExtOnly
End Function
Public Function ChangeExt(ByVal FilePath As String, Optional newext As String) As String
Dim temp As String
temp = Mid$(FilePath, 1, InStrRev(FilePath, "."))
temp = Left(temp, Len(temp) - 1)
If newext <> "" Then newext = "." + newext
ChangeExt = temp + newext
End Function
Public Function PathOnly(ByVal FilePath As String) As String
Dim temp As String
    temp = Mid$(FilePath, 1, InStrRev(FilePath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Public Function labeledit(Destination As String, Length As Integer) As String
Dim y As Integer, m As Integer, temp As String, temp1 As String, temp2 As String, temp3 As String
If Len(Destination) > Length Then
    m = 0
    For y = Len(Destination) To 1 Step -1
        m = m + 1
        If Mid(Destination, y, 1) = "\" Then
            temp2 = Right(Destination, m)
            Exit For
        End If
    Next y
    m = 0
    For y = 4 To Len(Destination)
        m = m + 1
        If Mid(Destination, y, 1) = "\" Then
            temp1 = Left(Destination, m + 3)
            Exit For
        End If
    Next y
    If Len(temp1 + temp2) < Length Then
doagain:
    m = Len(temp1) + 1
    For y = Len(temp1) + 2 To Len(Destination)
        m = m + 1
        If Mid(Destination, y, 1) = "\" Then
            temp = Left(Destination, m)
            Exit For
        End If
    Next y
     If Len(temp + temp2) < Length Then
     temp1 = temp
     GoTo doagain
     Else
     GoTo OKdone
     End If
     Else
     temp1 = Left(Destination, 3)
     End If
OKdone:
        m = Length - Len(temp1 + temp2)
        temp3 = "."
        For y = 1 To m
            temp3 = temp3 + "."
        Next y
    labeledit = temp1 + temp3 + temp2
Else
    labeledit = Destination
End If
End Function
Public Function TrimVoid(Expre)
  On Error Resume Next
  Dim i As Integer
  Dim beg As String
  Dim expr As String
  For i = 1 To Len(Expre)
        beg = Mid(Expre, i, 1)
        If beg Like "[a-zA-Z0-9]" Then expr = expr & beg
    Next
    TrimVoid = expr
End Function
Public Sub WriteINI(fileName As String, Section As String, Key As String, Text As String)
WritePrivateProfileString Section, Key, Text, fileName
End Sub
Public Function ReadINI(fileName As String, Section As String, Key As String)
ret = Space$(255)
Retlen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), fileName)
ret = Left$(ret, Retlen)
ReadINI = ret
End Function



Public Function GetTempFile2(lpTempFilename As String, path As String) As Boolean
    lpTempFilename = String(255, vbNullChar)
    GetTempFile2 = GetTempFilename(path, "bb", 0, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function




Public Function ReadText(path As String) As String
    Dim Line
    Dim temptxt As String
    temptxt = ""
    Open path For Input As #1
    Do While Not EOF(1)
        Input #1, Line
        temptxt = temptxt + Line
    Loop
    Close #1
    ReadText = temptxt
End Function


