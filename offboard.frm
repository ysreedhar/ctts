VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form offboard 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OffBoard"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cbo_company 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   3600
         Width           =   5535
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1035
         Width           =   3855
      End
      Begin VB.ComboBox cbo_empno 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txt_traveltime 
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Text            =   "0"
         Top             =   1035
         Width           =   975
      End
      Begin VB.ComboBox cbo_type 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txt_notes 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   2280
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker DTP_tdate 
         Height          =   315
         Left            =   4440
         TabIndex        =   7
         Top             =   435
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67174401
         CurrentDate     =   38733
      End
      Begin MSComCtl2.DTPicker dtp_offboard 
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   1635
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy H:mm:ss"
         Format          =   67174403
         CurrentDate     =   38733
      End
      Begin MSComCtl2.DTPicker dtp_onboard 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1635
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy H:mm:ss"
         Format          =   67174403
         CurrentDate     =   38733
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remarks"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date OffBoard"
         Height          =   195
         Left            =   2160
         TabIndex        =   16
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         Height          =   195
         Left            =   4440
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp No."
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date OnBoard"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Travel Time"
         Height          =   195
         Left            =   4440
         TabIndex        =   11
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "POB Type"
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "offboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_empno_Click()
nm = Split(cbo_empno.Text, "  -  ", Len(cbo_empno.Text), vbTextCompare)
Dim ofb As New ADODB.Recordset
If ofb.State Then ofb.Close
ofb.Open "select emp_classification from employee where emp_no='" & nm(0) & "' ", Cn, 3, 2
If Not ofb.EOF Then

txt_classification.Text = ofb(0)
End If
Dim ep As New ADODB.Recordset
If ep.State Then ep.Close
 ep.Open "select * from employee where emp_no = '" & nm(0) & "' and emp_traveltime='Applicable' ", Cn, 3, 2
 If Not ep.EOF Then
 txt_traveltime.Text = 8
 Else
 txt_traveltime.Text = 0
 End If
 
Dim pob As New ADODB.Recordset
If pob.State Then pob.Close
pob.Open "select ob_dateonboard from onboard where ob_empno='" & nm(0) & "' ", Cn, 3, 2
If Not pob.EOF Then
dtp_onboard.Value = pob(0)
End If
End Sub

Private Sub Form_Load()
Dim emp As New ADODB.Recordset
If emp.State Then emp.Close
emp.Open "select Distinct(e.emp_no),e.emp_name from employee e,onboard ob  where  e.emp_no=ob.ob_empno and  e.emp_status = 'y' and ob.location='" & main.lbllocation.Caption & "' order by e.emp_name", Cn, 3, 2
While Not emp.EOF
 
cbo_empno.AddItem emp(0) & "  -  " & emp(1)
 
emp.MoveNext
Wend

dtp_offboard.Value = Format(Date, "dd/MM/yyyy H:mm:ss")
dtp_onboard.Value = Format(Date, "dd/MM/yyyy H:mm:ss")
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
 
cbo_type.Text = "Contract Completed"
cbo_type.AddItem "Contract Completed"
cbo_type.AddItem "Go for TimeOff"
cbo_type.AddItem "Transfer To Other Barge"
cbo_type.AddItem "On Emergency Leave"

End Sub

