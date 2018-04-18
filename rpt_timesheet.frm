VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rpt_timesheet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   10890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10890
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   8655
         Begin VB.ComboBox cbo_emp 
            Height          =   315
            Left            =   720
            TabIndex        =   10
            Top             =   120
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker dtp_from 
            Height          =   375
            Left            =   4680
            TabIndex        =   5
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   67174401
            CurrentDate     =   38378
         End
         Begin MSComCtl2.DTPicker dtp_to 
            Height          =   375
            Left            =   6840
            TabIndex        =   7
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   67174401
            CurrentDate     =   38378
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Emp"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "To"
            Height          =   195
            Left            =   6480
            TabIndex        =   9
            Top             =   120
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "From"
            Height          =   195
            Left            =   4200
            TabIndex        =   8
            Top             =   120
            Width           =   345
         End
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
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Close"
         Height          =   255
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   1455
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
      TabIndex        =   6
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
Attribute VB_Name = "rpt_timesheet"
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
main.lbltitle.Caption = "TimeSheet"
dtp_from.Value = Date
dtp_to.Value = Date
Me.Top = 10
Me.Left = 10
 
WebBrowser.Navigate "About:Blank"
 

Dim emp As New ADODB.Recordset
If emp.State Then emp.Close
emp.Open "select Distinct(e.emp_no),e.emp_name from employee e, onboard b where e.emp_status = 'y' and b.location='" & main.lbllocation & "' order by emp_name", Cn, 3, 2
While Not emp.EOF
 
cbo_emp.AddItem emp(0) & "  -  " & emp(1)
 
emp.MoveNext
Wend
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub
Public Sub nocolor()
On Error Resume Next
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
    
   
            
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3><b>TL OFFSHORE SDN BHD</td>"
                fs.WriteLine "           <td colspan=4 ><b>TimeSheet " & Format(dtp_from.Value, "dd/MMM/yy") & "  -  " & Format(dtp_to.Value, "dd/MMM/yy") & "</td>"
                fs.WriteLine "           <td  colspan=2>Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"
                       
                fs.WriteLine "        </tr>"
    
    
   
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   
   fs.WriteLine "            <td Nowrap><font color=white>Name</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Classification</td>"
   fs.WriteLine "            <td Nowrap><font color=white>BargeNo.</td>"
   fs.WriteLine "            <td ><font color=white>Date onBoard</td>"
   fs.WriteLine "            <td ><font color=white>Days onBoard</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Shift</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Nationality</td>"
   fs.WriteLine "            <td ><font color=white>RoomNo</td>"
   fs.WriteLine "            <td ><font color=white>Life RaftNo.</td>"
    
   
   fs.WriteLine "        </tr>"
Dim sn As Integer
sn = 1
Dim diff As Double
diff = 0
Dim cnt As Integer
cnt = 0
nm = Split(cbo_emp.Text, "  -  ", Len(cbo_emp.Text), vbTextCompare)

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from employee e, timesheet t ,onboard o where e.emp_no = t.t_empno and e.emp_no=o.ob_empno  and t.t_empname='" & nm(1) & "'", Cn, 3, 2
If Not rs.EOF Then
diff = 0
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 
fs.WriteLine "            <td Nowrap >" & rs!emp_name & "</td>"
If rs!emp_classification = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td Nowrap >" & rs!emp_classification & "</td>"
End If
If rs!emp_no = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  Nowrap>" & rs!emp_no & "</td>"
End If
  
If rs!ob_dateonboard = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td >" & rs!ob_dateonboard & "</td>"
End If
diff = DateDiff("d", rs!ob_dateonboard, Date)
 
fs.WriteLine "            <td  >" & diff & "</td>"
 
If rs!ob_shift = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  >" & rs!ob_shift & "</td>"
End If
If rs!emp_nationality = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td >" & rs!emp_nationality & "</td>"
End If

If rs!ob_roomno = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  >" & rs!ob_roomno & "</td>"
End If

If rs!ob_raftno = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  >" & rs!ob_raftno & "</td>"
End If

 
fs.WriteLine "        </tr>"
End If
  
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "        <td colspan=4>"
  fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=100%>"
  fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=3 align=center><b>Regular Details</b></td>"
fs.WriteLine "        </tr>"
   fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
  fs.WriteLine "            <td  >Date</td>"
  fs.WriteLine "            <td  >Hrs</td>"
  fs.WriteLine "            <td  >JobNo</td>"
  fs.WriteLine "        </tr>"
Dim ts As New ADODB.Recordset
If ts.State Then ts.Close
ts.Open "select * from timesheet where t_empname='" & nm(1) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' order by t_r_date", Cn, 3, 2

While Not ts.EOF
  fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
  fs.WriteLine "            <td >" & ts!t_r_date & "  -  " & ts!daytype & "</td>"
  fs.WriteLine "            <td >" & ts!t_r_hrs & "</td>"
  fs.WriteLine "            <td >" & ts!t_r_job & "</td>"
  fs.WriteLine "        </tr>"

ts.MoveNext
Wend

fs.WriteLine " </table>"
fs.WriteLine "        </td>"
fs.WriteLine "        <td colspan=5>"
 fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=100%>"
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=3 align=center><b>OverTime Details</b></td>"
fs.WriteLine "        </tr>"
   fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
  fs.WriteLine "            <td  >Date</td>"
  fs.WriteLine "            <td >Hrs</td>"
  fs.WriteLine "            <td  >JobNo</td>"
  fs.WriteLine "        </tr>"
Dim ts1 As New ADODB.Recordset
If ts1.State Then ts1.Close
ts1.Open "select * from timesheet where t_empname='" & nm(1) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' order by t_r_date", Cn, 3, 2

While Not ts1.EOF
  fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
  fs.WriteLine "            <td >" & ts1!t_r_date & "  -  " & ts1!daytype & "</td>"
  fs.WriteLine "            <td  >" & ts1!t_o_hrs & "</td>"
  fs.WriteLine "            <td  >" & ts1!t_o_job & "</td>"
  fs.WriteLine "        </tr>"

ts1.MoveNext
Wend

fs.WriteLine " </table>"
fs.WriteLine "        </td>"
fs.WriteLine "        </tr>"







Dim x1 As Double
Dim x2 As Double
x1 = 0: x2 = 0
Dim rss As New ADODB.Recordset
If rss.State Then rss.Close
rss.Open "select SUM(t_r_hrs),SUM(t_o_hrs) from timesheet where t_empno='" & nm(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "'  ", Cn, 3, 2
If Not rss.EOF Then
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"

            Dim rss1 As New ADODB.Recordset
            If rss1.State Then rss1.Close
            rss1.Open "select SUM(t_r_hrs),SUM(t_o_hrs) from timesheet where t_empno='" & nm(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and daytype='G' ", Cn, 3, 2
            If Not rss1.EOF Then
            
            fs.WriteLine "            <td Nowrap  >T-GRhrs&nbsp;" & rss1(0) & "</td>"
            
            fs.WriteLine "            <td Nowrap  >T-GOThrs&nbsp;" & rss1(1) & "</td>"
            Else
            fs.WriteLine "            <td Nowrap  >T-GRhrs 0</td>"
            
            fs.WriteLine "            <td Nowrap  >T-GOThrs 0</td>"
            
            End If
                    Dim rss2 As New ADODB.Recordset
                    If rss2.State Then rss2.Close
                    rss2.Open "select SUM(t_r_hrs),SUM(t_o_hrs) from timesheet where t_empno='" & nm(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and daytype='P' ", Cn, 3, 2
                    If Not rss2.EOF Then
                                If IsNull(rss2(0)) Then
                                fs.WriteLine "            <td Nowrap >T-PGhrs 0</td>"
                                Else
                                
                                fs.WriteLine "            <td Nowrap >T-PGhrs&nbsp;" & rss2(0) & "</td>"
                                End If
                                If IsNull(rss2(1)) Then
                                fs.WriteLine "            <td Nowrap >T-POThrs 0</td>"
                                Else
                                
                                fs.WriteLine "            <td Nowrap >T-POThrs&nbsp;" & rss2(1) & "</td>"
                                End If
                    Else
                    fs.WriteLine "            <td Nowrap >T-PGhrs 0</td>"
                    
                    fs.WriteLine "            <td Nowrap >T-POThrs 0</td>"
                    
                    End If

                Dim rss3 As New ADODB.Recordset
                If rss3.State Then rss3.Close
                rss3.Open "select SUM(t_r_hrs),SUM(t_o_hrs) from timesheet where t_empno='" & nm(0) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' and daytype='R' ", Cn, 3, 2
                If Not rss3.EOF Then
                                    If IsNull(rss3(0)) Then
                                    fs.WriteLine "            <td Nowrap >T-RGhrs 0</td>"
                                    Else
                                    
                                    fs.WriteLine "            <td Nowrap >T-RGhrs&nbsp;" & rss3(0) & "</td>"
                                    End If
                                    If IsNull(rss3(1)) Then
                                    fs.WriteLine "            <td Nowrap >T-ROThrs 0</td>"
                                    Else
                                    
                                    fs.WriteLine "            <td Nowrap >T-ROThrs&nbsp;" & rss3(1) & "</td>"
                                    End If
                Else
                fs.WriteLine "            <td Nowrap >T-RGhrs 0</td>"
                
                fs.WriteLine "            <td Nowrap >T-ROThrs 0</td>"
                
                End If


fs.WriteLine "            <td Nowrap >T-Ghrs&nbsp;  " & rss(0) & "</td>"

fs.WriteLine "            <td Nowrap >T-OThrs&nbsp;" & rss(1) & "</td>"
fs.WriteLine "            <td Nowrap align=center>T-Mhrs&nbsp; <b>" & rss(1) + rss(0) & "</b> </td>"
 
x1 = rss(1) + rss(0)
fs.WriteLine "        </tr>"
 
End If







'earnings
Dim tlog As New ADODB.Recordset
If tlog.State Then tlog.Close
tlog.Open "select * from timelog where name='" & nm(1) & "' and Month(mnth)='" & Format(dtp_to.Value, "mm") & "' and year(mnth)='" & Format(dtp_to.Value, "yyyy") & "' ", Cn, 3, 2
Dim tr As Double
Dim act As Double
tr = 0: act = 0
If Not tlog.EOF Then
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Other Earnings(Hrs)</td>"
fs.WriteLine "            <td Nowrap align=center>Travel</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!travel & "</td>"
fs.WriteLine "            <td Nowrap align=center>Acr TOff</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!actoff & "</td>"
tr = tlog!travel
act = tlog!actoff
fs.WriteLine "            <td Nowrap align=center colspan=2>Total Mhrs</td>"
fs.WriteLine "            <td Nowrap align=center><b>" & CDbl(tr) + CDbl(act) + CDbl(x1) & "</b></td></tr>"


'ded

fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Deduction(RM)</td>"
fs.WriteLine "            <td Nowrap align=center>PhoneCall</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!phonecall & "</td>"
fs.WriteLine "            <td Nowrap align=center>CashAdv</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!cashadvance & "</td>"
fs.WriteLine "            <td Nowrap align=center >BondStr</td>"
fs.WriteLine "            <td Nowrap align=center>" & tlog!bondstore & "</td>"
fs.WriteLine "            <td Nowrap ><b>T-RM &nbsp;" & tlog!bondstore + tlog!cashadvance + tlog!phonecall & "</b></td></tr>"


'remarks
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Remarks</td>"
If (tlog!Notes) = "" Then
fs.WriteLine "            <td Nowrap colspan=8>&nbsp;</td></tr>"
Else
fs.WriteLine "            <td Nowrap colspan=8>" & tlog!Notes & "</td></tr>"
End If

End If
fs.WriteLine " </table>"
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub


