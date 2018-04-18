VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frm_login 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15345
   Icon            =   "frm_login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_login.frx":0E42
   ScaleHeight     =   12000
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_location 
      Height          =   315
      Left            =   8520
      TabIndex        =   3
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   2880
   End
   Begin VB.Frame PositionFrame 
      Caption         =   "Position"
      Enabled         =   0   'False
      Height          =   720
      Left            =   720
      TabIndex        =   27
      Top             =   5880
      Visible         =   0   'False
      Width           =   4170
      Begin VB.TextBox CharPosn 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   375
         TabIndex        =   29
         Top             =   255
         Width           =   570
      End
      Begin VB.TextBox CharPosn 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1845
         TabIndex        =   28
         Top             =   255
         Width           =   570
      End
      Begin VB.Label CharPosnLabel 
         Caption         =   "&X:"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   31
         Top             =   300
         Width           =   270
      End
      Begin VB.Label CharPosnLabel 
         Caption         =   "&Y:"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1620
         TabIndex        =   30
         Top             =   300
         Width           =   270
      End
   End
   Begin VB.Frame SpeechOutputFrame 
      Caption         =   "Speech &Output"
      Enabled         =   0   'False
      Height          =   2085
      Left            =   720
      TabIndex        =   21
      Top             =   3750
      Visible         =   0   'False
      Width           =   4170
      Begin VB.CheckBox BalloonStyleOption 
         Caption         =   "A&uto hide"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   420
         TabIndex        =   26
         Top             =   1650
         Width           =   1200
      End
      Begin VB.CheckBox BalloonStyleOption 
         Caption         =   "Auto &pace"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1605
         TabIndex        =   25
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CheckBox BalloonStyleOption 
         Caption         =   "Si&ze to text"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2895
         TabIndex        =   24
         Top             =   1665
         Width           =   1095
      End
      Begin VB.CheckBox BalloonStyleOption 
         Caption         =   "Display &word balloon"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1290
         Width           =   1935
      End
      Begin VB.TextBox SpeakText 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   930
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   255
         Width           =   3900
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   4995
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   4980
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Speak"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   4995
      TabIndex        =   18
      Top             =   3975
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Move"
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   4995
      TabIndex        =   17
      Top             =   6075
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame AnimationFrame 
      Caption         =   "&Animations for"
      Enabled         =   0   'False
      Height          =   2355
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   4155
      Begin VB.ListBox AnimationListBox 
         Enabled         =   0   'False
         Height          =   1620
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   3900
      End
      Begin VB.CheckBox OutputStyleOption 
         Caption         =   "Play sound &effects"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   15
         Top             =   1950
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox OutputStyleOption 
         Caption         =   "Stop &before next action"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1995
         TabIndex        =   14
         Top             =   1950
         Value           =   1  'Checked
         Width           =   1995
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   180
      TabIndex        =   12
      Text            =   "GestureDown"
      Top             =   345
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txt_password 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000006&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   8520
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton cmd_submit 
      Height          =   495
      Left            =   11160
      Picture         =   "frm_login.frx":1C1AE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Click to Login"
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton cmd_cancel 
      Height          =   495
      Left            =   11760
      Picture         =   "frm_login.frx":1C7E0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Click to Clear"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox cbo_userid 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000006&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   8520
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtp_cutdate 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy H:mm:ss"
      Format          =   28377091
      CurrentDate     =   38140
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4260
      Top             =   3105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8520
      TabIndex        =   32
      Top             =   5520
      Width           =   900
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   4860
      Top             =   3105
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   8520
      MouseIcon       =   "frm_login.frx":1CD9E
      Picture         =   "frm_login.frx":1DBE0
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Off PMS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   840
      TabIndex        =   11
      ToolTipText     =   "Turn Off PCIS"
      Top             =   10800
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frm_login.frx":1E189
      Top             =   10680
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CREW TRACKING AND TIMESHEET MANAGEMENT"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   3360
      TabIndex        =   10
      Top             =   720
      Width           =   9450
   End
   Begin VB.Label l1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label l2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label l3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cutt-Off Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3360
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TL OFFSHORE SDN BHD"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   9000
      TabIndex        =   0
      Top             =   3360
      Width           =   4395
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_userid_LostFocus()
Dim lc As New ADODB.Recordset
If lc.State Then lc.Close
lc.Open "select DISTINCT(project) from userproject where username='" & cbo_userid.Text & "' ", Cn, 3, 2
While Not lc.EOF
cbo_location.AddItem lc(0)
lc.MoveNext
Wend
lc.Close
End Sub

Private Sub cmd_cancel_Click()
cbo_userid.SetFocus
cbo_userid.Text = ""
txt_password.Text = ""
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_submit_Click()
main.lbllocation.Caption = frm_login.cbo_location.Text
Dim pwd As New ADODB.Recordset
If pwd.State Then pwd.Close
pwd.Open "select * from userid where a_userid='" & cbo_userid.Text & "' and a_password='" & txt_password.Text & "' ", Cn, 3, 2
If Not pwd.EOF Then
main.Label2.Caption = cbo_userid.Text
'main.Label1.Caption = "User:" & " " & cbo_userid.Text & "  " & "Login Time:" & " " & Format(Time, "HH:MM:SS")

main.Enabled = True
main.Show
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from login", Cn, 3, 2
rs.AddNew
rs!l_userid = cbo_userid.Text
rs!l_intime = Now
rs.Update
Unload frm_login

Else
MsgBox "Enter Correct Password"
cbo_userid.SetFocus
cbo_userid.Text = ""
txt_password.Text = ""

End If



'''Dim rsw As New ADODB.Recordset
'''If rsw.State Then rsw.Close
'''rsw.Open "select * from projectremainder where proj_user='" & main.Label2.Caption & "' and t_date='" & Format(Date, "MM/dd/yyyy") & "' ", Cn, 3, 2
'''If Not rsw.EOF Then
'''Load remainder
'''remainder.Show
'''End If


End Sub

Private Sub Form_Load()
Call connect
Me.Top = 0
Me.Left = 0
Me.Width = 16000
Me.Height = 16000

 
End Sub


Private Sub Label6_Click()
Unload Me
End Sub
