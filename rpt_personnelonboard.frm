VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rpt_personnelonboard 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10695
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Close"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00C0FFC0&
         Caption         =   "View"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   8415
         Begin VB.ComboBox cbo_location 
            Height          =   315
            Left            =   3840
            TabIndex        =   7
            Top             =   120
            Width           =   4335
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   6
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   67436545
            CurrentDate     =   38378
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2880
            TabIndex        =   8
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   975
         Left            =   120
         Top             =   240
         Width           =   8655
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   8145
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   10665
      ExtentX         =   18812
      ExtentY         =   14367
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "rpt_personnelonboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nic As String
Private Sub Check1_Click()
 
Call nocolor
 
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_print_Click()
On Error GoTo XIT
WebBrowser.ExecWB 6, OLECMDEXECOPT_DODEFAULT
XIT:
End Sub

Private Sub cmd_show_Click()
 
'''Load frmBusy
'''frmBusy.Show
'''frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
'Unload frmBusy
 
End Sub

Private Sub Form_Load()
main.lbltitle.Caption = "Personnel On Board List"
DTPicker1.Value = Date
Me.Top = 10
Me.Left = 10
 
WebBrowser.Navigate "About:Blank"
 Dim lc As New ADODB.Recordset
If lc.State Then lc.Close
lc.Open "select DISTINCT(project) from userproject where username='" & main.Label2.Caption & "' ", Cn, 3, 2
While Not lc.EOF
cbo_location.AddItem lc(0)
lc.MoveNext
Wend
lc.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

 

Public Sub nocolor()
Dim fso As New FileSystemObject
   Set fs = fso.CreateTextFile(App.Path & "\rep.html")
   fs.WriteLine " <html> "
   fs.WriteLine "<style>"
   fs.WriteLine "    BODY INPUT"
   fs.WriteLine "    {"
   fs.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
   'fs.WriteLine "      BORDER-BOTTOM: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-LEFT: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-RIGHT: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-TOP: Wheat 1px solid"
   fs.WriteLine "    }"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bolder;"
   fs.WriteLine "        CURSOR:HAND;"
   fs.WriteLine "    }"
   fs.WriteLine "    .TrFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: black;"
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   fs.WriteLine "        CURSOR:HAND;"
   fs.WriteLine "   }"
   fs.WriteLine "</style>"
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"
   
   
 

   
  


   fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
    
   
            
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><b>TL OFFSHORE SDN BHD</td>"
                fs.WriteLine "           <td colspan=5 ><b>PERSONNEL ON BOARD LIST AS OF MID-NIGHT " & Format(DTPicker1.Value, "dd/MMM/yy") & "</td>"
                fs.WriteLine "           <td  colspan=2>Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"
                          
                fs.WriteLine "        </tr>"
    
    
   
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap><font color=white>SNo</td>"
   fs.WriteLine "            <td Nowrap><font color=white>COY</td>"
   fs.WriteLine "            <td Nowrap><font color=white>EmpNo</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Name</td>"
   
   fs.WriteLine "            <td Nowrap><font color=white>Classification</td>"
'   fs.WriteLine "            <td Nowrap><font color=white>NIC</td>"
   fs.WriteLine "            <td ><font color=white>Date onBoard</td>"
   fs.WriteLine "            <td ><font color=white>Days onBoard</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Shift</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Nationality</td>"
   fs.WriteLine "            <td ><font color=white>Room No</td>"
   fs.WriteLine "            <td ><font color=white>Life RaftNo.</td>"
    
   
   fs.WriteLine "        </tr>"
Dim sn As Integer
sn = 1
Dim diff As Double
diff = 0
Dim cnt As Integer
cnt = 0


Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from employee e, onboard o where e.emp_no = o.ob_empno  and o.location ='" & cbo_location.Text & "' order by e.emp_classification", Cn, 3, 2
While Not rs.EOF
diff = 0
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  >" & sn & "</td>"
fs.WriteLine "            <td Nowrap >" & rs!emp_coy & "</td>"
If rs!emp_no = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td Nowrap >" & rs!emp_no & "</td>"
End If
If rs!emp_name = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  Nowrap>" & rs!emp_name & "</td>"
End If
If rs!emp_classification = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td Nowrap >" & rs!emp_classification & "</td>"
End If
 
If rs!ob_dateonboard = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  Nowrap>" & rs!ob_dateonboard & "</td>"
End If
diff = DateDiff("d", rs!ob_dateonboard, DTPicker1.Value)
 
fs.WriteLine "            <td  Nowrap>" & diff & "</td>"
 
If rs!ob_shift = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  Nowrap>" & rs!ob_shift & "</td>"
End If
If rs!emp_nationality = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  Nowrap>" & rs!emp_nationality & "</td>"
End If

If rs!ob_roomno = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  Nowrap>" & rs!ob_roomno & "</td>"
End If

If rs!ob_raftno = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  Nowrap>" & rs!ob_raftno & "</td>"
End If

 
fs.WriteLine "        </tr>"
 sn = sn + 1
 rs.MoveNext
 Wend
  fs.WriteLine " </table>"
  fs.WriteLine "    <table border=0 cellspacing=0 BORDERCOLOR=gray width=95%>"
  fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
 fs.WriteLine "            <td  Nowrap>&nbsp;</td>"
  fs.WriteLine "        </tr>"
 Dim onb As New ADODB.Recordset
 If onb.State Then onb.Close
 onb.Open "select DISTINCT(emp_classification) from employee order by emp_classification", Cn, 3, 2
 While Not onb.EOF
 cnt = 0
 Dim ni As New ADODB.Recordset
 If ni.State Then ni.Close
 ni.Open "select * from employee e, onboard o where e.emp_no = o.ob_empno and e.emp_classification = '" & onb(0) & "'  and o.location='" & cbo_location.Text & "' order by emp_classification", Cn, 3, 2
 While Not ni.EOF
 cnt = cnt + 1
 ni.MoveNext
 Wend
 fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
 fs.WriteLine "            <td  Nowrap>" & onb(0) & "    -    " & cnt & "</td>"
  fs.WriteLine "        </tr>"
 onb.MoveNext
 Wend
  fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

