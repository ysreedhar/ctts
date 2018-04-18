VERSION 5.00
Begin VB.Form userid 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Id"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   5895
   End
   Begin VB.TextBox txt_op 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3240
      TabIndex        =   12
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txt_hp 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txt_email 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   5895
   End
   Begin VB.ComboBox cbo_designation 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   5895
   End
   Begin VB.ComboBox cbo_department 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
   End
   Begin VB.TextBox txt_password 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "******************************"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txt_userid 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "Off Phone"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "HandPhone"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "userid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbo_department_Click()
Dim ds As New ADODB.Recordset
If ds.State Then ds.Close
ds.Open "select DISTINCT(dcode) from designation where ddept='" & cbo_department.Text & "'", Cn, 3, 2
While Not ds.EOF
cbo_designation.AddItem ds(0)
ds.MoveNext
Wend
ds.Close
End Sub

Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")

Dim dp As New ADODB.Recordset
If dp.State Then dp.Close
dp.Open "select DISTINCT(dcode) from department", Cn, 3, 2
While Not dp.EOF
cbo_department.AddItem dp(0)
dp.MoveNext
Wend
dp.Close

End Sub

