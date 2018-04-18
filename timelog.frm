VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form timelog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TimeLog"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txt_classification 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   915
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   "0"
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Text            =   "0"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "0"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txt_notes 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   2760
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker DTP_tdate 
         Height          =   315
         Left            =   3840
         TabIndex        =   9
         Top             =   315
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67108865
         CurrentDate     =   38733
      End
      Begin MSComCtl2.DTPicker dtp_tl 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   1515
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   67108867
         CurrentDate     =   38733
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Travel"
         Height          =   195
         Left            =   1680
         TabIndex        =   17
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         Height          =   195
         Left            =   3840
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acr TOff(hrs)"
         Height          =   195
         Left            =   2760
         TabIndex        =   14
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Call"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Advance"
         Height          =   195
         Left            =   1320
         TabIndex        =   12
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bond Store"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   1920
         Width           =   795
      End
   End
End
Attribute VB_Name = "timelog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_name_Click()
Dim em  As New ADODB.Recordset
If em.State Then em.Close
em.Open "select  (emp_classification) from employee  where emp_status = 'y' order by emp_name", Cn, 3, 2
If Not em.EOF Then
txt_classification.Text = em(0)
End If
Dim ts As New ADODB.Recordset
If ts.State Then ts.Close
ts.Open "select * from onboard o , employee e , offboard  ofb where o.ob_empno= e.emp_no and e.emp_no = ofb.empno and e.emp_name ='" & cbo_name & "' and month(o.ob_dateonboard)=" & Format(dtp_tl.Value, "MM") & " and year(o.ob_dateonboard)=" & Format(dtp_tl.Value, "yyyy") & " ", Cn, 3, 2
If Not ts.EOF Then
txt_travel.Text = ts!ob_traveltime + ts!traveltime
Else
txt_travel.Text = 0
End If
End Sub

Private Sub dtp_tl_Click()
Dim ts As New ADODB.Recordset
If ts.State Then ts.Close
ts.Open "select * from onboard o , employee e , ofb.offbwhere o.ob_empno= e.emp_no and e.emp_no = ofb.empno and e.emp_name ='" & cbo_name & "' and month(o.ob_dateonboard)=" & Format(dtp_tl.Value, "MM") & " and year(o.ob_dateonboard)=" & Format(dtp_tl.Value, "yyyy") & "", Cn, 3, 2
If Not ts.EOF Then
txt_travel.Text = ts!ob_traveltime + ts!traveltime
Else
txt_travel.Text = 0
End If
End Sub

Private Sub Form_Load()
DTP_tdate.Value = Date
dtp_tl.Value = Format(Date, "MM/yyyy")
Dim emp As New ADODB.Recordset
If emp.State Then emp.Close
emp.Open "select  Distinct(e.emp_no),e.emp_name from employee e, onboard ob where e.emp_status = 'y' and e.emp_no=ob.ob_empno and ob.location='" & main.lbllocation.Caption & "' order by e.emp_name", Cn, 3, 2
While Not emp.EOF
 
cbo_name.AddItem emp(0) & "  -  " & emp(1)
 
emp.MoveNext
Wend
End Sub
