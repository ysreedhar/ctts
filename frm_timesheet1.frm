VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_timesheet1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_daytype 
      Height          =   315
      Left            =   5760
      TabIndex        =   24
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   7800
      TabIndex        =   20
      Top             =   960
      Width           =   855
      Begin VB.CommandButton cmd_close 
         Caption         =   "Close"
         Height          =   975
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "Save"
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "Earn/Ded"
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.ComboBox cbo_shift 
      Height          =   315
      Left            =   3720
      TabIndex        =   16
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Notes"
      Height          =   2295
      Left            =   3720
      TabIndex        =   14
      Top             =   5760
      Width           =   3855
      Begin VB.TextBox txt_notes 
         Appearance      =   0  'Flat
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.ComboBox cbo_proj 
      Height          =   315
      Left            =   1620
      TabIndex        =   13
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   7035
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Regular"
      Height          =   2295
      Left            =   3720
      TabIndex        =   5
      Top             =   960
      Width           =   3855
      Begin VB.ListBox List4 
         Height          =   1635
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   480
         Width           =   3000
      End
      Begin VB.TextBox txt_rg 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Job No"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rg Hrs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Over Time"
      Height          =   2295
      Left            =   3720
      TabIndex        =   0
      Top             =   3360
      Width           =   3855
      Begin VB.ListBox List5 
         Height          =   1635
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   480
         Width           =   3000
      End
      Begin VB.TextBox txt_ot 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Job No"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT Hrs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67108865
      CurrentDate     =   38381
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Day Type"
      Height          =   195
      Left            =   5760
      TabIndex        =   25
      Top             =   120
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shift"
      Height          =   195
      Left            =   3720
      TabIndex        =   19
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Proj / DOE"
      Height          =   195
      Left            =   1680
      TabIndex        =   18
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee  No - Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frm_timesheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub cbo_proj_Change()
Call list4func
Call list5func
Call empdisp
End Sub

Private Sub cbo_proj_Click()
Call list4func
Call list5func
Call empdisp
End Sub

Private Sub cbo_shift_Change()
Call empdisp

End Sub

Private Sub cbo_shift_Click()
Call empdisp
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()
Dim p As Integer
Dim q As Integer
Dim r As Integer
p = 0: q = 0: r = 0
Dim sv As New ADODB.Recordset

For p = 0 To List1.ListCount - 1
If List1.Selected(p) = True Then
nm = Split(List1.List(p), "  -  ", Len(List1.List(p)), vbTextCompare)
         q = 0: r = 0
If sv.State Then sv.Close
sv.Open " select * from timesheet where t_empno='" & nm(0) & "' and t_r_date='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "'", Cn, 3, 2
If sv.EOF Then
                                sv.AddNew
                                
                                sv!t_empno = nm(0)
                                sv!t_empname = nm(1)
                                sv!t_r_date = DTPicker1.Value
                                sv!t_r_hrs = txt_rg.Text
                                If cbo_daytype.Text = "GENERAL DAY" Then
                                sv!daytype = "G"
                                ElseIf cbo_daytype.Text = "PUBLIC HOLIDAY" Then
                                sv!daytype = "P"
                                ElseIf cbo_daytype.Text = "REST DAY" Then
                                sv!daytype = "R"
                                End If
                                For q = 0 To List4.ListCount - 1
                                If List4.Selected(q) = True Then
                                sv!t_r_job = List4.List(q)
                                End If
                                Next q
                              
                                sv!t_o_hrs = txt_ot.Text
                                For r = 0 To List5.ListCount - 1
                                If List5.Selected(r) = True Then
                                sv!t_o_job = List5.List(r)
                                End If
                                Next r
                                sv!Notes = txt_notes.Text
                                sv!t_date = Date
                                sv!u_date = Now
                                sv!t_user = main.Label2.Caption
                                sv.Update
     
     ElseIf Not sv.EOF Then
     
     Dim svn As New ADODB.Recordset
     If svn.State Then svn.Close
     svn.Open "select SUM((t_r_hrs)),SUM((t_o_hrs)) from timesheet where t_empno='" & nm(0) & "' and t_r_date='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "'", Cn, 3, 2
     If CDbl(svn(0)) + CDbl(svn(1)) >= 24 Then
     MsgBox "Entry Already Posted for MR. " & nm(1) & ""
     Else
                                sv!t_empno = nm(0)
                                sv!t_empname = nm(1)
                                sv!t_r_date = DTPicker1.Value
                                sv!t_r_hrs = txt_rg.Text
                                If cbo_daytype.Text = "GENERAL DAY" Then
                                sv!daytype = "G"
                                ElseIf cbo_daytype.Text = "PUBLIC HOLIDAY" Then
                                sv!daytype = "P"
                                ElseIf cbo_daytype.Text = "REST DAY" Then
                                sv!daytype = "R"
                                End If
                                For q = 0 To List4.ListCount - 1
                                If List4.Selected(q) = True Then
                                sv!t_r_job = List4.List(q)
                                End If
                                Next q
                              
                                sv!t_o_hrs = txt_ot.Text
                                For r = 0 To List5.ListCount - 1
                                If List5.Selected(r) = True Then
                                sv!t_o_job = List5.List(r)
                                End If
                                Next r
                                sv!Notes = txt_notes.Text
                                sv!t_date = Date
                                sv!u_date = Now
                                sv!t_user = main.Label2.Caption
                                sv.Update
     End If
 
 End If
End If
Next p

MsgBox " Records Processed"
End Sub

Private Sub cmd_view_Click()
 frm_monthendentries.Show
 SetParent frm_monthendentries.HWnd, main.HWnd
End Sub

Private Sub DTPicker1_Change()
If "Sun" = Format(DTPicker1.Value, "ddd") Then
cbo_daytype.Text = "REST DAY"
End If

Dim ph As New ADODB.Recordset
If ph.State Then ph.Close
ph.Open "select ph_date from pholiday where ph_date='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "' ", Cn, 3, 2
If Not ph.EOF Then
cbo_daytype.Text = "PUBLIC HOLIDAY"
ElseIf "Sun" = Format(DTPicker1.Value, "ddd") Then
cbo_daytype.Text = "REST DAY"
Else
cbo_daytype.Text = "GENERAL DAY"
End If
End Sub

Private Sub DTPicker1_Click()
 If "Sun" = Format(DTPicker1.Value, "ddd") Then
cbo_daytype.Text = "REST DAY"
End If
Dim ph1 As New ADODB.Recordset
If ph1.State Then ph1.Close
ph1.Open "select ph_date from pholiday where ph_date='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "' ", Cn, 3, 2
If Not ph1.EOF Then
cbo_daytype.Text = "PUBLIC HOLIDAY"
ElseIf "Sun" = Format(DTPicker1.Value, "ddd") Then
cbo_daytype.Text = "REST DAY"
Else
cbo_daytype.Text = "GENERAL DAY"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DTPicker1.Value = Date
Me.Top = 5
Me.Left = 5

main.lbltitle.Caption = "Daily Attendance"
 
 

cbo_shift.Text = "DS"
cbo_shift.AddItem "DS"
cbo_shift.AddItem "NS"

cbo_proj.Text = "PROJECT"
cbo_proj.AddItem "PROJECT"
cbo_proj.AddItem "DOE"


cbo_daytype.Text = "GENERAL DAY"
cbo_daytype.AddItem "GENERAL DAY"
cbo_daytype.AddItem "PUBLIC HOLIDAY"
cbo_daytype.AddItem "REST DAY"



Dim ph2 As New ADODB.Recordset
If ph2.State Then ph2.Close
ph2.Open "select ph_date from pholiday where ph_date='" & Format(DTPicker1.Value, "mm/dd/yyyy") & "' ", Cn, 3, 2
If Not ph2.EOF Then
cbo_daytype.Text = "PUBLIC HOLIDAY"
ElseIf "Sun" = Format(DTPicker1.Value, "ddd") Then
cbo_daytype.Text = "REST DAY"
Else
cbo_daytype.Text = "GENERAL DAY"
End If
 
 
End Sub
Public Sub list4func()
List4.Clear
Dim jn As New ADODB.Recordset
If jn.State Then jn.Close
jn.Open "select Distinct(jb_name) from jobno where jb_proj='" & cbo_proj.Text & "' order by jb_name ", Cn, 3, 2
While Not jn.EOF
List4.AddItem jn(0)
jn.MoveNext
Wend
End Sub
Public Sub list5func()
List5.Clear
Dim jno As New ADODB.Recordset
If jno.State Then jno.Close
jno.Open "select Distinct(jb_name) from jobno where jb_proj='" & cbo_proj.Text & "' order by jb_name ", Cn, 3, 2
While Not jno.EOF
List5.AddItem jno(0)
jno.MoveNext
Wend
End Sub

Public Sub empdisp()
List1.Clear
Dim emp As New ADODB.Recordset
If emp.State Then emp.Close
emp.Open "select Distinct(e.emp_no),e.emp_name from employee e , onboard o where  e.emp_no=o.ob_empno and e.emp_status = 'y' and o.ob_shift='" & cbo_shift.Text & "' and e.emp_chargetype = '" & cbo_proj.Text & "'  and o.location='" & main.lbllocation.Caption & "' order by e.emp_no", Cn, 3, 2
While Not emp.EOF
If emp(1) = "" Then
List1.AddItem emp(0)
Else
List1.AddItem emp(0) & "  -  " & emp(1)
End If
emp.MoveNext
Wend
End Sub

Private Sub List4_Click()
Dim k As Integer
If List4.SelCount > 1 Then
MsgBox "Cannot Select more then one Job"
k = 0
For k = 0 To List4.ListCount - 1
List4.Selected(k) = False
Next k
End If


End Sub

Private Sub List5_Click()
Dim h As Integer
If List5.SelCount > 1 Then
MsgBox "Cannot Select more then one Job"
h = 0
For h = 0 To List5.ListCount - 1
List5.Selected(h) = False
Next h
End If
End Sub
