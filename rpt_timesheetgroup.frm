VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rpt_timesheetgroup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Close"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00C0FFC0&
         Caption         =   "View"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   8655
         Begin VB.ComboBox cbo_location 
            Height          =   315
            Left            =   3960
            TabIndex        =   10
            Top             =   360
            Width           =   4335
         End
         Begin MSComCtl2.DTPicker dtp_from 
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   28311553
            CurrentDate     =   38378
         End
         Begin MSComCtl2.DTPicker dtp_to 
            Height          =   315
            Left            =   2160
            TabIndex        =   3
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   28311553
            CurrentDate     =   38378
         End
         Begin VB.Label Label3 
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
            Left            =   3960
            TabIndex        =   11
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "From"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   120
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "To"
            Height          =   195
            Left            =   2160
            TabIndex        =   4
            Top             =   120
            Width           =   195
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   975
         Left            =   120
         Top             =   240
         Width           =   8895
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   8145
      Left            =   0
      TabIndex        =   9
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "rpt_timesheetgroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim nic As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long
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
main.lbltitle.Caption = "TimeSheet"
dtp_from.Value = Date
dtp_to.Value = Date
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
   
   
 

   
  


   fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=90%>"
    
Dim dta As Integer
Dim dtb As Date
Dim inc As Integer
inc = 0
dta = DateDiff("d", dtp_from.Value, dtp_to.Value)
            
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=2><b>TL OFFSHORE SDN BHD</td>"
                fs.WriteLine "           <td colspan=6 align=center ><b>TimeSheet " & Format(dtp_from.Value, "dd/MMM/yy") & "  -  " & Format(dtp_to.Value, "dd/MMM/yy") & "</td>"
                fs.WriteLine "           <td  colspan=3>Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"
                       
                fs.WriteLine "        </tr>"
    
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td ><font color=white>EmpNo - Name</td>"
                fs.WriteLine "            <td ><font color=white>Classification</td>"
                fs.WriteLine "            <td colspan=9><font color=white>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                 
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td ><font color=white>JOBNo</td>"
                fs.WriteLine "            <td ><font color=white>G-RGhrs</td>"
                fs.WriteLine "            <td ><font color=white>G-OThrs</td>"
                fs.WriteLine "            <td ><font color=white>P-RGhrs</td>"
                fs.WriteLine "            <td ><font color=white>P-OThrs</td>"
                 fs.WriteLine "            <td ><font color=white>R-RGhrs</td>"
                fs.WriteLine "            <td ><font color=white>R-OThrs</td>"
                fs.WriteLine "            <td ><font color=white>T-RGhrs</td>"
                fs.WriteLine "            <td ><font color=white>T-OThrs</td>"
'                fs.WriteLine "            <td ><font color=white>Travel</td>"
'                fs.WriteLine "            <td ><font color=white>Acr-Toff</td>"
                fs.WriteLine "            <td ><font color=white>T-Mhrs</td>"

                fs.WriteLine "        </tr>"
Dim sn As Integer
sn = 1
Dim diff As Double
diff = 0
Dim cnt As Integer
cnt = 0
 Dim xxx As String
Dim ii As Date
Dim jj As Integer
Dim jjj As Integer
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(o.ob_empno),e.emp_name,e.emp_classification  from onboard o , employee e ,timesheet t where o.ob_empno=e.emp_no and t.t_empno=o.ob_empno and e.emp_status='y' and o.location='" & cbo_location.Text & "' order by e.emp_name ", Cn, 3, 2
While Not rs.EOF


            fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
            
            fs.WriteLine "            <td Nowrap >" & rs(0) & "  -  " & rs(1) & "</td>"
            fs.WriteLine "            <td Nowrap >" & rs!emp_classification & "</td>"
            fs.WriteLine "            <td colspan=9><font color=white>&nbsp;</td>"
            
            fs.WriteLine "        </tr>"

 
 
Cn.Execute "delete from jobro"
Dim jn As New ADODB.Recordset
If jn.State Then jn.Close
jn.Open "select DISTINCT(t_r_job),t_r_date  from  timesheet  where t_empno='" & rs(0) & "' and   t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "'  ", Cn, 3, 2
While Not jn.EOF
Dim jro As New ADODB.Recordset
If jro.State Then jro.Close
jro.Open "select * from jobro", Cn, 3, 2
jro.AddNew
jro!job = jn(0)
jro!dte = jn(1)
jro!emp = rs(0)
jro.Update
jn.MoveNext
Wend
Dim jnq As New ADODB.Recordset
If jnq.State Then jnq.Close
jnq.Open "select DISTINCT(t_o_job) ,t_r_date  from  timesheet  where t_empno='" & rs(0) & "' and   t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "'  ", Cn, 3, 2
While Not jnq.EOF
Dim jroq As New ADODB.Recordset
If jroq.State Then jroq.Close
jroq.Open "select * from jobro", Cn, 3, 2
jroq.AddNew
jroq!job = jnq(0)
jroq!dte = jnq(1)
jroq!emp = rs(0)

jroq.Update
jnq.MoveNext
Wend


 
Dim jnw As New ADODB.Recordset
If jnw.State Then jnw.Close
jnw.Open "select DISTINCT(job)  from jobro  where emp ='" & rs(0) & "' and   dte between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "'  ", Cn, 3, 2
While Not jnw.EOF

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap >&nbsp;</td>"
fs.WriteLine "            <td Nowrap >" & jnw(0) & "</td>"
Dim tss As New ADODB.Recordset
If tss.State Then tss.Close
tss.Open "select SUM(t_r_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and t_r_job='" & jnw(0) & "' and daytype='G' ", Cn, 3, 2
If Not tss.EOF Then
If IsNull(tss(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & tss(0) & "</td>"
End If

Else

fs.WriteLine "            <td Nowrap align=center>0</td>"
End If
Dim tsso As New ADODB.Recordset
If tsso.State Then tsso.Close
tsso.Open "select SUM(t_o_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and t_o_job='" & jnw(0) & "' and daytype='G' ", Cn, 3, 2
If Not tsso.EOF Then
If IsNull(tsso(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & tsso(0) & "</td>"
End If

Else

fs.WriteLine "            <td Nowrap align=center>0</td>"
End If

Dim tsss As New ADODB.Recordset
If tsss.State Then tsss.Close
tsss.Open "select SUM(t_r_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and t_r_job='" & jnw(0) & "' and daytype='P' ", Cn, 3, 2
If Not tsss.EOF Then
If IsNull(tsss(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & tsss(0) & "</td>"
End If
 
Else
 
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If
Dim tssso As New ADODB.Recordset
If tssso.State Then tssso.Close
tssso.Open "select SUM(t_o_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and t_o_job='" & jnw(0) & "' and daytype='P' ", Cn, 3, 2
If Not tssso.EOF Then
If IsNull(tssso(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & tssso(0) & "</td>"
End If
 
Else
 
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If
Dim tssss As New ADODB.Recordset
If tssss.State Then tssss.Close
tssss.Open "select SUM(t_r_hrs)  from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and t_r_job='" & jnw(0) & "' and daytype='R' ", Cn, 3, 2
If Not tssss.EOF Then
If IsNull(tssss(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & tssss(0) & "</td>"
End If
 
Else
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If

Dim tsssso As New ADODB.Recordset
If tsssso.State Then tsssso.Close
tsssso.Open "select  SUM(t_o_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and t_o_job='" & jnw(0) & "' and daytype='R' ", Cn, 3, 2
If Not tsssso.EOF Then
If IsNull(tsssso(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & tsssso(0) & "</td>"
End If
 
Else
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If
Dim trhrs As Double
Dim tohrs As Double
trhrs = 0: tohrs = 0
Dim trhrs1 As Double
Dim tohrs1 As Double
trhrs1 = 0: tohrs1 = 0
Dim ts As New ADODB.Recordset
If ts.State Then ts.Close
ts.Open "select SUM(t_r_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and t_r_job='" & jnw(0) & "' ", Cn, 3, 2
If Not ts.EOF Then

If IsNull(ts(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & ts(0) & "</td>"
trhrs1 = ts(0)
End If

Else
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If

Dim tso As New ADODB.Recordset
If tso.State Then tso.Close
tso.Open "select SUM(t_o_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and t_o_job='" & jnw(0) & "' ", Cn, 3, 2
If Not tso.EOF Then

If IsNull(tso(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & tso(0) & "</td>"
tohrs1 = tso(0)
End If

Else
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If


fs.WriteLine "            <td Nowrap align=center>" & CDbl(trhrs1) + CDbl(tohrs1) & "</td>"

fs.WriteLine "        </tr>"
 

'''''''''''''''''''''


jnw.MoveNext
Wend



Dim x1 As Double
Dim x2 As Double
x1 = 0: x2 = 0
Dim rss As New ADODB.Recordset
If rss.State Then rss.Close
rss.Open "select SUM(t_r_hrs),SUM(t_o_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "'  ", Cn, 3, 2
If Not rss.EOF Then
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SubTot</td>"
Dim rss1 As New ADODB.Recordset
If rss1.State Then rss1.Close
rss1.Open "select SUM(t_r_hrs),SUM(t_o_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and daytype='G' ", Cn, 3, 2
If Not rss1.EOF Then
If IsNull(rss1(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & rss1(0) & "</td>"
End If
If IsNull(rss1(1)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & rss1(1) & "</td>"
End If
Else
fs.WriteLine "            <td Nowrap align=center>0</td>"
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If
Dim rss2 As New ADODB.Recordset
If rss2.State Then rss2.Close
rss2.Open "select SUM(t_r_hrs),SUM(t_o_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and daytype='P' ", Cn, 3, 2
If Not rss2.EOF Then
If IsNull(rss2(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & rss2(0) & "</td>"
End If
If IsNull(rss2(1)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & rss2(1) & "</td>"
End If
Else
fs.WriteLine "            <td Nowrap align=center>0</td>"
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If

Dim rss3 As New ADODB.Recordset
If rss3.State Then rss3.Close
rss3.Open "select SUM(t_r_hrs),SUM(t_o_hrs) from timesheet where t_empno='" & rs(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and daytype='R' ", Cn, 3, 2
If Not rss3.EOF Then
If IsNull(rss3(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & rss3(0) & "</td>"
End If
If IsNull(rss3(1)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & rss3(1) & "</td>"
End If
Else
fs.WriteLine "            <td Nowrap align=center>0</td>"
fs.WriteLine "            <td Nowrap align=center>0</td>"
End If
If IsNull(rss(0)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & rss(0) & "</td>"
End If
If IsNull(rss(1)) Then
fs.WriteLine "            <td Nowrap align=center>0</td>"
Else
fs.WriteLine "            <td Nowrap align=center>" & rss(1) & "</td>"
End If
fs.WriteLine "            <td Nowrap align=center><b>" & rss(1) + rss(0) & "</b></td>"
On Error Resume Next
x1 = rss(1) + rss(0)
fs.WriteLine "        </tr>"
 
 
End If


'earnings
Dim tlog As New ADODB.Recordset
If tlog.State Then tlog.Close
tlog.Open "select * from timelog where name='" & rs(1) & "' and Month(mnth)='" & Format(dtp_to.Value, "mm") & "' and year(mnth)='" & Format(dtp_to.Value, "yyyy") & "' ", Cn, 3, 2
Dim tr As Double
Dim act As Double
tr = 0: act = 0
If Not tlog.EOF Then
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=3>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Other Earnings(Hrs)</td>"
fs.WriteLine "            <td Nowrap align=center>Travel</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!travel & "</td>"
fs.WriteLine "            <td Nowrap align=center>Acr TOff</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!actoff & "</td>"
tr = tlog!travel
act = tlog!actoff
fs.WriteLine "            <td Nowrap align=center colspan=2>Total Mhrs</td>"
fs.WriteLine "            <td Nowrap align=right colspan=2><b>" & CDbl(tr) + CDbl(act) + CDbl(x1) & "&nbsp;&nbsp;</b></td></tr>"


'ded

fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=3>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Deduction(RM)</td>"
fs.WriteLine "            <td Nowrap align=center>PhoneCall</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!phonecall & "</td>"
fs.WriteLine "            <td Nowrap align=center>CashAdv</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!cashadvance & "</td>"
fs.WriteLine "            <td Nowrap align=center >BondStr</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!bondstore & "</td>"
fs.WriteLine "            <td Nowrap ><b>T-RM</b></td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!bondstore + tlog!cashadvance + tlog!phonecall & "</td></tr>"

'remarks
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap  >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Remarks</td>"
If (tlog!Notes) = "" Then
fs.WriteLine "            <td Nowrap colspan=10>&nbsp;</td></tr>"
Else
fs.WriteLine "            <td Nowrap colspan=10>" & tlog!Notes & "</td></tr>"
End If
End If




xxx = rs(0)
rs.MoveNext
Wend


fs.WriteLine " </table>"
 
 
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub
 
