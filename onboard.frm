VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form onboard 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crew OnBoard"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   0
      TabIndex        =   15
      Top             =   1200
      Width           =   3015
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear Crew Selection"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txt_search 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select All Crew"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Crew"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   915
      End
   End
   Begin VB.TextBox txt_notes 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   8040
      Width           =   9855
   End
   Begin VB.ListBox List1 
      Height          =   6585
      Left            =   3120
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1320
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   4110
      Left            =   6720
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1320
      Width           =   3255
   End
   Begin VB.ListBox List3 
      Height          =   2085
      Left            =   6720
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   5760
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.ComboBox cbo_shift 
         Height          =   315
         Left            =   9840
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cbo_type 
         Height          =   315
         Left            =   3000
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cbo_company 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtp_onboard 
         Height          =   315
         Left            =   7920
         TabIndex        =   20
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   28311555
         CurrentDate     =   38377
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   255
         Left            =   9840
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date On Board"
         Height          =   255
         Left            =   7920
         TabIndex        =   21
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "POB Type"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Company"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Charge Type"
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Crew No - Name"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No     -     Total       -     Vacant"
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Raft No."
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7680
      Width           =   1095
   End
End
Attribute VB_Name = "onboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flgl2 As Double
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

Private Sub cbo_type_Click()
If cbo_type.Text = "New Crew Member" Then
MsgBox " First Need to enter CrewMember Details"
frm_employee.Show
Exit Sub
End If
End Sub

Private Sub Check1_Click()
Dim lt As Double
If Check1.Value = 1 Then

lt = 0
For lt = 0 To List1.ListCount - 1
List1.Selected(lt) = True
Next lt

Else
lt = 0
For lt = 0 To List1.ListCount - 1
List1.Selected(lt) = False
Next lt
End If
End Sub

Private Sub Check2_Click()
lt = 0
For lt = 0 To List1.ListCount - 1
List1.Selected(lt) = False
Next lt
End Sub

Private Sub Form_Load()
On Error Resume Next
dtp_onboard.Value = Format(Date, "dd/mm/yy H:MM:ss")
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


cbo_type.Text = "New Crew Member"
cbo_type.AddItem "New Crew Member"
cbo_type.AddItem "Return From TimeOff"
cbo_type.AddItem "Transfer From Barge"
cbo_type.AddItem "Return From Emergency Leave"



Dim rm As New ADODB.Recordset
If rm.State Then rm.Close
rm.Open "select Distinct(rm_name),rm_capacity  from room order by rm_name", Cn, 3, 2
While Not rm.EOF
Dim rmc As New ADODB.Recordset
 If rmc.State Then rmc.Close
    rmc.Open "select (ob_roomno) from onboard where ob_roomno = '" & rm(0) & "'", Cn, 3, 2
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

flgl2 = 0
End Sub

Private Sub List1_ItemCheck(Item As Integer)
If flgl2 = 1 Then
Dim ig As Integer
ig = 0
For ig = 0 To List2.ListCount - 1
If List2.Selected(ig) = True Then
    sppl = Split(List2.List(ig), "  -  ", Len(List2.List(ig)), vbTextCompare)
    If List1.SelCount > 1 Then
If List1.SelCount > sppl(2) Then
MsgBox "Sorry! This Room can Accomdate only " & Trim(sppl(2)) & " Members"
List1.Selected(Item) = False
Exit Sub
End If
End If
End If
Next ig



End If
End Sub

Private Sub List2_ItemCheck(Item As Integer)
If List2.SelCount >= 2 Then
MsgBox "Only one Room can be selected"
List2.Selected(Item) = False
End If

spl = Split(List2.List(Item), "  -  ", Len(List2.List(Item)), vbTextCompare)

If spl(2) = 0 Then
List2.Selected(Item) = False
MsgBox "This Room is Fully Occupied"
End If

If List1.SelCount > CDbl(spl(2)) Then
MsgBox "Sorry! This Room can Accomdate only " & Trim(spl(2)) & " Members"
List2.Selected(Item) = False
Exit Sub
End If

If List2.SelCount = 1 Then
flgl2 = 1
End If
End Sub

Private Sub List3_ItemCheck(Item As Integer)
If List3.SelCount >= 2 Then
MsgBox "Only one RaftNo can be selected"
 
List3.Selected(Item) = False
 
End If
End Sub

