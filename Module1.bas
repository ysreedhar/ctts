Attribute VB_Name = "Module1"
Option Explicit
'**********************************************************************************************



'API types
Private Type USER_INFO
    Name As String
    Comment As String
    UserComment As String
    FullName As String
End Type

Private Type USER_INFO_API
    Name As Long
    Comment As Long
    UserComment As Long
    FullName As Long
End Type

Public UserInfo(0 To 1000) As USER_INFO

'API calls
Private Declare Function NetUserEnum Lib "netapi32" (lpServer As Any, ByVal Level As Long, ByVal Filter As Long, lpBuffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal pBuffer As Long) As Long

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

'API Constants
Private Const NERR_Success As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234&

Private Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Private Const FILTER_NORMAL_ACCOUNT As Long = &H2&
Private Const FILTER_PROXY_ACCOUNT As Long = &H4&
Private Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
Private Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
Private Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&

Private Declare Function GetUserName _
Lib "advapi32.dll" Alias "GetUserNameA" _
                  (ByVal lpBuffer As String, _
                  nSize As Long) As Long
                  
Public Declare Function NetServerEnum _
    Lib "Netapi32.dll" ( _
    vServername As Any, _
    ByVal lLevel As Long, _
    vBufptr As Any, _
    lPrefmaxlen As Long, _
    lEntriesRead As Long, _
    lTotalEntries As Long, _
    vServerType As Any, _
    ByVal sDomain As String, _
    vResumeHandle As Any) _
    As Long

Public Declare Sub RtlMoveMemory _
    Lib "kernel32" ( _
    dest As Any, _
    vSrc As Any, _
    ByVal lSize&)

Public Declare Sub lstrcpyW _
    Lib "kernel32" ( _
    vDest As Any, _
    ByVal sSrc As Any)
    
Declare Sub lstrcpy _
    Lib "kernel32" ( _
    vDest As Any, _
    ByVal vSrc As Any)

Declare Sub lstrcpynW _
    Lib "kernel32" ( _
    ByVal vDest As Any, _
    ByVal vSrc As Any, _
lLength As Long)



Declare Function NetWkstaGetInfo _
    Lib "Netapi32.dll" ( _
    ByVal sServerName$, _
    ByVal lLevel&, _
    vBuffer As Any) _
    As Long

Declare Function NetMessageBufferSend _
    Lib "Netapi32.dll" ( _
    ByVal sServerName$, _
    ByVal sMsgName$, _
    ByVal sFromName$, _
    ByVal sMessageText$, _
    ByVal lBufferLength&) _
    As Long

Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_servername As Long
End Type

Public Type SERVER_INFO_101
    dw_platform_id As Long
    ptr_name As Long
    dw_ver_major As Long
    dw_ver_minor As Long
    dw_type As Long
    ptr_comment As Long
End Type

Type WKSTA_INFO_100
    wki100_platform_id As Long
    wki100_computername As Long
    wki100_langroup As Long
    wki100_ver_major As Long
    wki100_ver_minor As Long
End Type


Public Const SV_TYPE_WORKSTATION = &H1
Public Const SV_TYPE_SERVER = &H2
Public Const SV_TYPE_SQLSERVER = &H4
Public Const SV_TYPE_DOMAIN_CTRL = &H8
Public Const SV_TYPE_DOMAIN_BAKCTRL = &H10
Public Const SV_TYPE_TIMESOURCE = &H20
Public Const SV_TYPE_AFP = &H40
Public Const SV_TYPE_NOVELL = &H80
Public Const SV_TYPE_DOMAIN_MEMBER = &H100
Public Const SV_TYPE_LOCAL_LIST_ONLY = &H40000000
Public Const SV_TYPE_PRINT = &H200
Public Const SV_TYPE_DIALIN = &H400
Public Const SV_TYPE_XENIX_SERVER = &H800
Public Const SV_TYPE_MFPN = &H4000
Public Const SV_TYPE_NT = &H1000
Public Const SV_TYPE_WFW = &H2000
Public Const SV_TYPE_SERVER_NT = &H8000
Public Const SV_TYPE_POTENTIAL_BROWSER = &H10000
Public Const SV_TYPE_BACKUP_BROWSER = &H20000
Public Const SV_TYPE_MASTER_BROWSER = &H40000
Public Const SV_TYPE_DOMAIN_MASTER = &H80000
Public Const SV_TYPE_DOMAIN_ENUM = &H80000000
Public Const SV_TYPE_WINDOWS = &H400000
Public Const SV_TYPE_ALL = &HFFFFFFFF

Public SERVERTYPE  As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public fab As Integer


Public Function FillDomainTree(lType As Long, tvw As TreeView) As Boolean
         
Dim lReturn As Long
Dim Server_Info As Long
Dim lEntries As Long
Dim lTotal As Long
Dim lMax As Long
Dim vResume As Variant
Dim tServer_info_101 As SERVER_INFO_101
Dim sServer As String
Dim sDomain As String
Dim lServerInfo101StructPtr As Long
Dim X As Long, i As Long
Dim bBuffer(512) As Byte
Dim nodex As Node
tvw.Nodes.Clear
Set nodex = tvw.Nodes.Add(, , "R", "Network Domains", "dm")
nodex.Expanded = True

    lReturn = NetServerEnum( _
        ByVal 0&, _
        101, _
        Server_Info, _
        lMax, _
        lEntries, _
        lTotal, _
        ByVal lType, _
        sDomain, _
        vResume)

    If lReturn <> 0 Then
        'StatusBar1.Panels("msg").Text = "Systemmessages here"
        
        Exit Function
    End If

    X = 1
    lServerInfo101StructPtr = Server_Info

    Do While X <= lTotal
        DoEvents
        RtlMoveMemory _
            tServer_info_101, _
            ByVal lServerInfo101StructPtr, _
            Len(tServer_info_101)

        lstrcpyW bBuffer(0), _
            tServer_info_101.ptr_name

      
        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & _
                Chr$(bBuffer(i))
            i = i + 2
            DoEvents
        Loop
       Set nodex = tvw.Nodes.Add("R", tvwChild, sServer, sServer, "dmmac")
       nodex.Expanded = True
       Call AddDomainServers(SERVERTYPE, tvw, sServer)
        DoEvents
        X = X + 1
            sServer = ""
        lServerInfo101StructPtr = _
            lServerInfo101StructPtr + _
            Len(tServer_info_101)

    Loop

    lReturn = NetApiBufferFree(Server_Info)
        
End Function
Private Sub AddDomainServers(lType As Long, tvw As TreeView, Parentkey As String)

Dim lReturn As Long
Dim Server_Info As Long
Dim lEntries As Long
Dim lTotal As Long
Dim lMax As Long
Dim vResume As Variant
Dim tServer_info_101 As SERVER_INFO_101
Dim sServer As String
Dim sDomain As String
Dim lServerInfo101StructPtr As Long
Dim X As Long, i As Long
Dim bBuffer(512) As Byte
Dim nodex As Node
sDomain = StrConv(Parentkey, vbUnicode)

    
    lReturn = NetServerEnum( _
        ByVal 0&, _
        101, _
        Server_Info, _
        lMax, _
        lEntries, _
        lTotal, _
        ByVal lType, _
        sDomain, _
        vResume)

    If lReturn <> 0 Then
        'StatusBar1.Panels("msg").Text = "Systemmessages here"
        Exit Sub
    End If

    X = 1
    lServerInfo101StructPtr = Server_Info

    Do While X <= lTotal
        DoEvents
        RtlMoveMemory _
            tServer_info_101, _
            ByVal lServerInfo101StructPtr, _
            Len(tServer_info_101)

        lstrcpyW bBuffer(0), _
            tServer_info_101.ptr_name


        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & _
                Chr$(bBuffer(i))
            i = i + 2
            DoEvents
        Loop
        Set nodex = tvw.Nodes.Add(Parentkey, tvwChild, "W" + Parentkey + sServer, sServer, "cmac")
       nodex.Expanded = True
        DoEvents
        X = X + 1
            sServer = ""
        lServerInfo101StructPtr = _
            lServerInfo101StructPtr + _
            Len(tServer_info_101)

    Loop

    lReturn = NetApiBufferFree(Server_Info)


End Sub
Public Function GetLocalSystemName()
    Dim lReturnCode As Long
    Dim bBuffer(512) As Byte
    Dim i As Integer
    Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
    Dim lwkstaInfo100StructPtr As Long
    Dim sLocalName As String
    
    lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
 
    lwkstaInfo100StructPtr = lwkstaInfo100
                 
    If lReturnCode = 0 Then
                 
        RtlMoveMemory twkstaInfo100, ByVal _
        lwkstaInfo100StructPtr, Len(twkstaInfo100)
         
        lstrcpyW bBuffer(0), twkstaInfo100.wki100_computername

        i = 0
        Do While bBuffer(i) <> 0
            sLocalName = sLocalName & Chr(bBuffer(i))
            i = i + 2
        Loop
            
        GetLocalSystemName = sLocalName
         
    End If

End Function

Public Function GetDomainName() As String
    
    Dim lReturnCode As Long
    Dim bBuffer(512) As Byte
    Dim i As Integer
    Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
    Dim lwkstaInfo100StructPtr As Long
    Dim sDomainName As String
    
    lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
 
    lwkstaInfo100StructPtr = lwkstaInfo100
                 
    If lReturnCode = 0 Then
                 
        RtlMoveMemory twkstaInfo100, ByVal lwkstaInfo100StructPtr, Len(twkstaInfo100)
         
        lstrcpyW bBuffer(0), twkstaInfo100.wki100_langroup
        
        
        i = 0
        Do While bBuffer(i) <> 0
            sDomainName = sDomainName & Chr(bBuffer(i))
            i = i + 2
        Loop
            
        GetDomainName = sDomainName
         
    End If
        
End Function
Public Function NetSend(msg As String, ToNode As String) As Boolean
Dim lReturnCode As Long
Dim sUnicodeToName As String
Dim sUnicodeFromName As String
Dim sUnicodeMessage As String
Dim lMessageLength As Long
 
    sUnicodeFromName = StrConv(GetLocalSystemName, vbUnicode)
    sUnicodeToName = StrConv(ToNode, vbUnicode)
    sUnicodeMessage = StrConv(msg + vbCrLf & vbCrLf & vbCrLf & "_______________________________________" & vbCrLf & UserName + "(" + GetDomainName + ")", vbUnicode)
    lMessageLength = Len(sUnicodeMessage)

  
    Screen.MousePointer = vbHourglass
    
    

    lReturnCode = NetMessageBufferSend("", _
        sUnicodeToName, _
        sUnicodeFromName, _
        sUnicodeMessage, _
        lMessageLength)
    
   
'    If lReturnCode = 0 Then
'       Form1.StatusBar1.Panels(1).Text = "Your message was sent correct.."
'    Else
'       Form1.StatusBar1.Panels(1).Text = "Cannot Sent the message. Error No  : " + CStr(lReturnCode)
'
'    End If
    
  
   Screen.MousePointer = vbDefault

End Function

  Public Property Get UserName() As Variant
          Dim sBuffer As String
          Dim lSize As Long
          sBuffer = Space$(255)
          lSize = Len(sBuffer)
          Call GetUserName(sBuffer, lSize)
          UserName = Left$(sBuffer, lSize)
          
     End Property



 Private Function PtrToString(lpwString As Long) As String
    'Convert a LPWSTR pointer to a VB string
    Dim Buffer() As Byte
    Dim nLen As Long

    If lpwString Then
        nLen = lstrlenW(lpwString) * 2
        If nLen Then
            ReDim Buffer(0 To (nLen - 1)) As Byte
            CopyMem Buffer(0), ByVal lpwString, nLen
            PtrToString = Buffer
        End If
    End If
End Function

Public Function GetUsers(ByVal ServerName As String) As Long
    Dim lpBuffer As Long
    Dim nRet As Long
    Dim EntriesRead As Long
    Dim TotalEntries As Long
    Dim ResumeHandle As Long
    Dim uUser As USER_INFO_API
    Dim bServer() As Byte
    Dim i As Integer

    If Trim(ServerName) = "" Then
        'Local users
        bServer = vbNullString
    Else
        'Check the syntax of the ServerName string
        If InStr(ServerName, "\\") = 1 Then
            bServer = ServerName & vbNullChar
        Else
            bServer = "\\" & ServerName & vbNullChar
        End If
    End If
    i = 0
    ResumeHandle = 0
    Do
        'Start to enumerate the Users
        If Trim(ServerName) = "" Then
            nRet = NetUserEnum(vbNullString, 10, FILTER_NORMAL_ACCOUNT, lpBuffer, 1, EntriesRead, TotalEntries, ResumeHandle)
        Else
            nRet = NetUserEnum(bServer(0), 10, FILTER_NORMAL_ACCOUNT, lpBuffer, 1, EntriesRead, TotalEntries, ResumeHandle)
        End If
        'Fill the data structure for the User
        If nRet = ERROR_MORE_DATA Then
            CopyMem uUser, ByVal lpBuffer, Len(uUser)
            UserInfo(i).Name = PtrToString(uUser.Name)
            UserInfo(i).Comment = PtrToString(uUser.Comment)
            UserInfo(i).UserComment = PtrToString(uUser.UserComment)
            UserInfo(i).FullName = PtrToString(uUser.FullName)
            i = i + 1
        End If
        If lpBuffer Then
            Call NetApiBufferFree(lpBuffer)
        End If
    Loop While nRet = ERROR_MORE_DATA
    'Return the number of Users
    GetUsers = i
End Function
 




 

