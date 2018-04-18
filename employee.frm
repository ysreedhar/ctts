VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form employee 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CrewMember Details"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.ComboBox cbo_nationality 
         Height          =   315
         Left            =   3525
         TabIndex        =   7
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txt_name 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   915
         Width           =   5775
      End
      Begin VB.TextBox txt_empno 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cbo_sex 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox cbo_classification 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txt_icno 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cbo_company 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   2280
         Width           =   2775
      End
      Begin VB.ComboBox cbo_chargetype 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   3000
         Width           =   2415
      End
      Begin VB.ComboBox cbo_traveltime 
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txt_notes 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   3600
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker DTP_tdate 
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   28246017
         CurrentDate     =   38733
      End
      Begin MSComCtl2.DTPicker dtp_join 
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   28246017
         CurrentDate     =   38733
      End
      Begin VB.TextBox txt_age 
         Height          =   285
         Left            =   1395
         TabIndex        =   9
         Text            =   "0"
         Top             =   1560
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtp_dob 
         Height          =   315
         Left            =   2160
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   28246017
         CurrentDate     =   38733
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DOB"
         Height          =   195
         Left            =   2160
         TabIndex        =   28
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp No."
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   270
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         Height          =   195
         Left            =   4680
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   195
         Left            =   1395
         TabIndex        =   20
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
         Height          =   195
         Left            =   3525
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Join Date"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   315
         Left            =   3240
         TabIndex        =   17
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICNo."
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Type"
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Travel Time"
         Height          =   195
         Left            =   3960
         TabIndex        =   14
         Top             =   2760
         Width           =   840
      End
   End
End
Attribute VB_Name = "employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
dtp_join.Value = Format(Date, "dd/MM/yyyy")
dtp_dob.Value = Format(Date, "dd/MM/yyyy")
cbo_sex.Text = "M"
cbo_sex.AddItem "M"
cbo_sex.AddItem "F"

Dim nt As New ADODB.Recordset
If nt.State Then nt.Close
nt.Open "select * from nationality order by nt_name", Cn, 3, 2
While Not nt.EOF
cbo_nationality.AddItem nt!nt_name
nt.MoveNext
Wend

Dim jcl As New ADODB.Recordset
If jcl.State Then jcl.Close
jcl.Open "select * from jobclassification order by jcl_name", Cn, 3, 2
While Not jcl.EOF
cbo_classification.AddItem jcl!jcl_name
jcl.MoveNext
Wend

Dim cmp As New ADODB.Recordset
If cmp.State Then cmp.Close
cmp.Open "select * from company order by coy_name", Cn, 3, 2
While Not cmp.EOF
cbo_company.AddItem cmp!coy_name
cmp.MoveNext
Wend
cbo_chargetype.Text = "PROJECT"
cbo_chargetype.AddItem "PROJECT"
cbo_chargetype.AddItem "DOE"
cbo_chargetype.AddItem "OVERHEAD"
cbo_chargetype.AddItem "N/A"


cbo_traveltime.Text = "Applicable"
cbo_traveltime.AddItem "Applicable"
cbo_traveltime.AddItem "Not Applicable"

End Sub

