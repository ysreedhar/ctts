Attribute VB_Name = "connectionmod"
Dim fs As New FileSystemObject
Dim a As Object
Dim aaa As String
    Dim svr As String
    Dim db As String
    Dim uid As String
    Dim pwd As String
Public ConnectString
Public Cn As New ADODB.Connection
Public ct1 As Integer
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type

' API
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

' Contants
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10 ' Responds with "yes to all" for any dialog box that is displayed.
Private Const FOF_SILENT = &H4

Public Type BROWSEINFOTYPE
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Const WM_USER = &H400

Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const LPTR = (&H0 Or &H40)
Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long



Public Function connect()
Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OpenTextFile(App.Path & "\TIMESHEETSETUP.txt")
    aaa = a.ReadLine
    s = Split(aaa, ";")
    'str = Split(aaa, ";", , vbTextCompare)
    svr = s(0)
    db = s(1)
    uid = s(2)
    pwd = s(3)
    If Cn.State Then Cn.Close
    Cn.ConnectionString = "driver={SQL Server};server=" & svr & ";uid=" & uid & ";pwd=" & pwd & ";database=" & db
    Cn.Open
    Set fs = Nothing
End Function
Public Function validatechk(TempKeyAscii As Integer, Optional TempInt As Integer = 0) As Integer

Select Case TempInt
Case 0: ' FOR ACCEPT ONLY A-Z AND a-z  and  backspace and space AND CONVERT UPPER CASE

If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) Then
 
 validatechk = 0
 Else
validatechk = Asc((UCase(Chr(TempKeyAscii))))
 
 End If

Case 1: ' FOR ACCEPT ONLY A-Z and a-z  and  backspace and  dot and spase and convert upper case
If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) Then
 
 validatechk = 0
 Else
validatechk = Asc((UCase(Chr(TempKeyAscii))))
End If

Case 2: ' FOR ACCEPT ONLY A-Z and a-z  and  backspace and  dot and spase and convert uppercase
If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) And (Not TempKeyAscii Like 44) Then
 
 validatechk = 0
 Else
validatechk = Asc((UCase(Chr(TempKeyAscii))))
End If
Case 3: ' FOR ACCEPT ONLY A-Z and a-z  and  backspace and  convert uppercase
If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not TempKeyAscii Like 8) Then
 
 validatechk = 0
 Else
validatechk = Asc((UCase(Chr(TempKeyAscii))))
 
 End If
 Case 4: ' FOR ACCEPT ONLY 0-9  and  backspace and space and dot and coma
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) And (Not TempKeyAscii Like 44) Then
 
 
 validatechk = 0
 Else
validatechk = TempKeyAscii
 
 End If
Case 5: ' FOR ACCEPT ONLY 0-9  and  backspace
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) Then
 
 
 validatechk = 0
 Else
validatechk = TempKeyAscii
 
 End If
 Case 6: ' FOR ACCEPT ONLY 0-9  and  backspace and minus(-)
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 45) Then
 
 
 validatechk = 0
 Else
validatechk = TempKeyAscii
 
 End If
  Case 7: ' FOR ACCEPT ONLY 0-9  and  backspace AND COLON(:)
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 58) And (Not TempKeyAscii Like 8) Then
 
 
 validatechk = 0
 Else
validatechk = TempKeyAscii
 
 End If
 '59, 58, 34, 39, 44, 60, 46, 62, 47, 61, 43, 63, 123, 125, 92, 124, 96, 126, 33, 37, 94, 38, 95
  Case 8: 'FOR MOMILE NO
If (TempKeyAscii Like 59) Or (TempKeyAscii Like 58) Or (TempKeyAscii Like 34) Or (TempKeyAscii Like 39) _
Or (TempKeyAscii Like 44) Or (TempKeyAscii Like 60) Or (TempKeyAscii Like 46) Or (TempKeyAscii Like 62) _
Or (TempKeyAscii Like 47) Or (TempKeyAscii Like 61) Or (TempKeyAscii Like 43) Or (TempKeyAscii Like 63) _
Or (TempKeyAscii Like 123) Or (TempKeyAscii Like 125) Or (TempKeyAscii Like 92) Or (TempKeyAscii Like 124) _
Or (TempKeyAscii Like 96) Or (TempKeyAscii Like 126) Or (TempKeyAscii Like 33) Or (TempKeyAscii Like 37) _
Or (TempKeyAscii Like 94) Or (TempKeyAscii Like 38) Or (TempKeyAscii Like 95) Then
validatechk = 0
 Else
validatechk = TempKeyAscii
End If

 Case 9: ' FOR ACCEPT ONLY 0-9  and  backspace and minus(-) and opening small brakit(  and closing small brakit )
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 45) And (Not TempKeyAscii Like 40) And (Not TempKeyAscii Like 41) Then
 
 
 validatechk = 0
 Else
validatechk = TempKeyAscii
 
 End If
  Case 10: 'not accept ' "
If (TempKeyAscii Like 34) Or (TempKeyAscii Like 39) Then
validatechk = 0
 Else
validatechk = TempKeyAscii
End If
Case 11: ' FOR ACCEPT ONLY 0-9  and  backspace and space and dot and
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) Then
 
 
 validatechk = 0
 Else
validatechk = TempKeyAscii
 
 End If
Case 12: ' FOR ACCEPT ONLY A-Z AND a-z  and 0-9 and  backspace and AND CONVERT UPPER CASE

If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) Then
 
 validatechk = 0
 Else
validatechk = Asc((UCase(Chr(TempKeyAscii))))
 
 End If


Case 13: ' FOR ACCEPT ONLY A-Z AND a-z  and 0-9 and  backspace and AND SPACE AND DOT AND minus(-) and opening small brakit(  and closing small brakit )and coma CONVERT UPPER CASE

If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) And (Not TempKeyAscii Like 45) And (Not TempKeyAscii Like 40) And (Not TempKeyAscii Like 41) And (Not TempKeyAscii Like 44) Then
 
 validatechk = 0
 Else
validatechk = Asc((UCase(Chr(TempKeyAscii))))
 
 End If
Case 14: ' FOR ACCEPT ONLY A-Z AND a-z  and 0-9 and  backspace and AND SPACE AND DOT AND minus(-) and / and opening small brakit(  and closing small brakit )  and   CONVERT UPPER CASE

If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) And (Not TempKeyAscii Like 45) And (Not TempKeyAscii Like 40) And (Not TempKeyAscii Like 41) And (Not TempKeyAscii Like 47) And (Not TempKeyAscii Like 44) Then
 
 validatechk = 0
 Else
validatechk = Asc((UCase(Chr(TempKeyAscii))))
 
 End If
Case 15: ' FOR ACCEPT ONLY 0-9  and  backspace and minus(-) and /
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 45) And (Not TempKeyAscii Like 47) Then
 
 
 validatechk = 0
 Else
validatechk = TempKeyAscii
 
 End If
Case 16: 'for blood group
If Not TempKeyAscii Like 45 And Not TempKeyAscii Like 43 And Not Chr(TempKeyAscii) Like "[A-Za-z]" And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) Then
validatechk = 0
Else
validatechk = TempKeyAscii
End If
End Select

End Function


