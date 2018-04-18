VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_onboard 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_traveltime 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7080
      TabIndex        =   19
      Text            =   "0"
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   4200
         TabIndex        =   17
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cbo_company 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Charge Type"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Company"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   9240
      TabIndex        =   10
      Top             =   2880
      Width           =   975
      Begin VB.CommandButton cmd_close 
         Caption         =   "Close"
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "View "
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd_process 
         Caption         =   "Process"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ComboBox cbo_shift 
      Height          =   315
      Left            =   7080
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtp_onboard 
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
      Format          =   28180483
      CurrentDate     =   38377
   End
   Begin VB.ListBox List3 
      Height          =   3885
      Left            =   3720
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   5160
      Width           =   3135
   End
   Begin VB.ListBox List2 
      Height          =   3885
      Left            =   3720
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   8160
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Travel Time"
      Height          =   255
      Left            =   7080
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date On Board"
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Raft No."
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No     -     Total       -     Vacant"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee  No - Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frm_onboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbo_company_Click()
List1.Clear
Dim emp1 As New ADODB.Recordset
If emp1.State Then emp1.Close
emp1.Open "select Distinct(e.emp_no),e.emp_name from employee e,company c  where e.emp_coy=c.coy_name  and e.emp_status = 'x' and e.emp_coy = '" & cbo_company.Text & "' and e.emp_chargetype='" & cbo_proj.Text & "' order by e.emp_no", Cn, 3, 2
While Not emp1.EOF
If emp1(1) = "" Then
List1.AddItem emp1(0)
Else
List1.AddItem emp1(0) & "  -  " & emp1(1)
End If
emp1.MoveNext
Wend
End Sub

Private Sub cbo_proj_Click()
List1.Clear
Dim emp2 As New ADODB.Recordset
If emp2.State Then emp2.Close
emp2.Open "select Distinct(e.emp_no),e.emp_name from employee e,company c  where e.emp_coy=c.coy_name  and e.emp_status = 'x' and e.emp_coy = '" & cbo_company.Text & "' and e.emp_chargetype='" & cbo_proj.Text & "' order by e.emp_no", Cn, 3, 2
While Not emp2.EOF
If emp2(1) = "" Then
List1.AddItem emp2(0)
Else
List1.AddItem emp2(0) & "  -  " & emp2(1)
End If
emp2.MoveNext
Wend
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_process_Click()
On Error Resume Next
Dim ob As New ADODB.Recordset
Dim l As Integer
Dim m As Integer
Dim k As Integer
k = 0: m = 0: l = 0
Dim rmn As String
Dim rft As String
For l = 0 To List2.ListCount - 1
If List2.Selected(l) = True Then
rmn = List2.List(l)
End If
Next l
For m = 0 To List3.ListCount - 1
If List3.Selected(m) = True Then
rft = List3.List(m)
End If
Next m

For k = 0 To List1.ListCount - 1
If List1.Selected(k) = True Then
nm = Split(List1.List(k), "  -  ", Len(List1.List(k)), vbTextCompare)
                If ob.State Then ob.Close
                ob.Open "select * from onboard ", Cn, 3, 2
                ob.AddNew
                ob!ob_empno = nm(0)
                ob!ob_dateonboard = dtp_onboard.Value
                ob!ob_shift = cbo_shift.Text
                ob!ob_proj = cbo_proj.Text
                ob!ob_roomno = rmn
                ob!ob_raftno = rft
                ob!ob_traveltime = txt_traveltime.Text
                ob!Location = main.lbllocation.Caption
                ob.Update
                ob.Close
                Cn.Execute "update employee set emp_status = 'y' where emp_no='" & nm(0) & "' "
                

End If
Next k
List1.Clear
Dim emp1 As New ADODB.Recordset
If emp1.State Then emp1.Close
emp1.Open "select Distinct(emp_no),emp_name from employee where emp_status = 'x' order by emp_no", Cn, 3, 2
While Not emp1.EOF
If emp1(1) = "" Then
List1.AddItem emp1(0)
Else
List1.AddItem emp1(0) & "  -  " & emp1(1)
End If
emp1.MoveNext
Wend
MsgBox "POB Processed Successfully"
'
'assad:
'      MsgBox "Employee is Already On Board"
End Sub

Private Sub cmd_view_Click()
rpt_personnelonboard.Show
 
End Sub

Private Sub Form_Load()
On Error Resume Next
dtp_onboard.Value = Date
Me.Top = 5
Me.Left = 5

main.lbltitle.Caption = "OnBoard"

Dim emp As New ADODB.Recordset
If emp.State Then emp.Close
emp.Open "select Distinct(e.emp_no),e.emp_name from employee e,company c  where e.emp_coy=c.coy_name  and e.emp_status = 'x' and e.emp_coy = '" & cbo_company.Text & "' and e.emp_chargetype='" & cbo_proj.Text & "' order by e.emp_no", Cn, 3, 2
While Not emp.EOF
If emp(1) = "" Then
List1.AddItem emp(0)
Else
List1.AddItem emp(0) & "  -  " & emp(1)
End If
emp.MoveNext
Wend



Dim rm As New ADODB.Recordset
If rm.State Then rm.Close
rm.Open "select Distinct(rm_name),rm_capacity  from room order by rm_name", Cn, 3, 2
While Not rm.EOF
Dim rmc As New ADODB.Recordset
 If rmc.State Then rmc.Close
    rmc.Open "select (ob_roomno) from onboard where ob_roomno = " & rm(0), Cn, 3, 2
    If Not rmc.EOF Then
    List2.AddItem rm(0) & "          -          " & rm(1) & "          -          " & (rm(1) - rmc.RecordCount)
    Else
    List2.AddItem rm(0) & "          -          " & rm(1) & "          -          " & rm(1)
    End If
rm.MoveNext
Wend

Dim lr As New ADODB.Recordset
If lr.State Then lr.Close
lr.Open "select Distinct(lr_name),notes from liferaft order by lr_name", Cn, 3, 2
While Not lr.EOF
If lr(1) = "" Then
List3.AddItem lr(0)
Else
List3.AddItem lr(0) & "  -  " & lr(1)
End If
lr.MoveNext
Wend


Dim cy As New ADODB.Recordset
If cy.State Then cy.Close
cy.Open "select DISTINCT(coy_name) from company order by coy_name", Cn, 3, 2
While Not cy.EOF
cbo_company.AddItem cy(0)
cy.MoveNext
Wend

cbo_shift.Text = "DS"
cbo_shift.AddItem "DS"
cbo_shift.AddItem "NS"

cbo_proj.Text = "PROJECT"
cbo_proj.AddItem "PROJECT"
cbo_proj.AddItem "DOE"
 

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub List1_Click()
Dim mt As Integer
Dim mm As Integer
Dim ep As New ADODB.Recordset
For mt = 0 To List1.ListCount - 1
If List1.Selected(mt) = True Then
nn = Split(List1.List(mt), "  -  ", Len(List1.List(mt)), vbTextCompare)
 
 If ep.State Then ep.Close
 ep.Open "select * from employee where emp_no = '" & nn(0) & "' and emp_traveltime='Applicable' ", Cn, 3, 2
 If Not ep.EOF Then
 txt_traveltime.Text = 8
 Else
 txt_traveltime.Text = 0
 End If
Else
txt_traveltime.Text = 0
End If
 
 Next mt
End Sub

Private Sub List2_ItemCheck(Item As Integer)
If List2.SelCount >= 2 Then
MsgBox "Only one Room can Be selected"
 
List2.Selected(Item) = False
 
End If
End Sub

Private Sub List3_ItemCheck(Item As Integer)
If List3.SelCount >= 2 Then
MsgBox "Only one RaftNo can Be selected"
 
List3.Selected(Item) = False
 
End If
End Sub
