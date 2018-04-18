VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_timesheet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   11355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11355
   ScaleWidth      =   15030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_view 
      BackColor       =   &H00FF8080&
      Caption         =   "Close"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmd_save 
      BackColor       =   &H00FF8080&
      Caption         =   "Save"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Over Time"
      Height          =   2295
      Left            =   0
      TabIndex        =   16
      Top             =   5760
      Width           =   3735
      Begin VB.TextBox txt_ot 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.ListBox List5 
         Height          =   1635
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   480
         Width           =   3000
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
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Job No"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Regular"
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   3360
      Width           =   3735
      Begin VB.TextBox txt_rg 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.ListBox List4 
         Height          =   1635
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   480
         Width           =   3000
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
         TabIndex        =   15
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Job No"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3975
      Begin MSComCtl2.DTPicker dtp_from 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   28311553
         CurrentDate     =   38381
      End
      Begin MSComCtl2.DTPicker dtp_to 
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   28311553
         CurrentDate     =   38381
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ListBox List3 
      Height          =   7260
      Left            =   7320
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   7260
      Left            =   3720
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   2085
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Over Time"
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Regular"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee  No - Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frm_timesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rjb As String
Public Y As Integer
Private Sub cmd_save_Click()
Dim p As Integer
Dim q As Integer
Dim r As Integer
p = 0: q = 0: r = 0

For p = 0 To List1.ListCount - 1
If List1.Selected(p) = True Then
nm = Split(List1.List(p), "  -  ", Len(List1.List(p)), vbTextCompare)
                                For q = 0 To List2.ListCount - 1
                                If List2.Selected(q) = True Then
                                nm1 = Split(List2.List(q), "  -  ", Len(List2.List(q)), vbTextCompare)
                                Dim sv As New ADODB.Recordset
                                If sv.State Then sv.Close
                                sv.Open " select * from timesheet where t_empno='" & nm(0) & "' and t_empname='" & nm(1) & "' and t_r_date='" & Format(nm1(0), "mm/dd/yyyy") & "' order by t_id", Cn, 3, 2
                                        If Not sv.EOF Then
                                        sv!t_empno = nm(0)
                                        sv!t_empname = nm(1)
                                        sv!t_r_date = nm1(0)
                                        sv!t_r_hrs = nm1(1)
                                        sv!t_r_job = nm1(2)
                                        sv!u_date = Now
                                        sv!t_user = main.Label2.Caption
                                        sv!t_date = Date
                                        sv.Update
                                        End If
                                End If
                                Next q
      
            q = 0
            For q = 0 To List3.ListCount - 1
            If List3.Selected(q) = True Then
                nm2 = Split(List3.List(q), "  -  ", Len(List3.List(q)), vbTextCompare)
                Dim su As New ADODB.Recordset
                If su.State Then su.Close
                su.Open "select * from timesheet where t_empno='" & nm(0) & "' and t_empname='" & nm(1) & "' and t_r_date='" & Format(nm2(0), "mm/dd/yyyy") & "' order by t_id", Cn, 3, 2
                If Not su.EOF Then
                su!t_o_hrs = nm2(1)
                su!t_o_job = nm2(2)
                
                su.Update
                End If
            
            End If
            Next q
          
End If
Next p

MsgBox " Records Processed"
                

End Sub

Private Sub cmd_view_Click()
 Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
dtp_onboard.Value = Date
dtp_from.Value = Date
dtp_to.Value = Date
Me.Top = 5
Me.Left = 5

main.lbltitle.Caption = "Manhour TimeLog"

Dim emp As New ADODB.Recordset
If emp.State Then emp.Close
emp.Open "select Distinct(e.emp_no),e.emp_name from employee e, onboard b  where e.emp_status = 'y' and b.location= '" & main.lbllocation & "' order by emp_name", Cn, 3, 2
While Not emp.EOF
If emp(1) = "" Then
List1.AddItem emp(0)
Else
List1.AddItem emp(0) & "  -  " & emp(1)
End If
emp.MoveNext
Wend
Y = 1
Call list4func
Call list5func
End Sub

 

Public Sub list4func()
Dim jn As New ADODB.Recordset
If jn.State Then jn.Close
jn.Open "select Distinct(jb_name) from jobno order by jb_name ", Cn, 3, 2
While Not jn.EOF
List4.AddItem jn(0)
jn.MoveNext
Wend
End Sub
Public Sub list5func()
Dim jno As New ADODB.Recordset
If jno.State Then jno.Close
jno.Open "select Distinct(jb_name) from jobno order by jb_name ", Cn, 3, 2
While Not jno.EOF
List5.AddItem jno(0)
jno.MoveNext
Wend
End Sub

Public Sub applyjob()

End Sub

Private Sub List1_Click()
'rglr
 On Error GoTo assad
Dim k As Integer: Dim i As Integer

 'rglr
 
 
If List1.SelCount > 1 Then
MsgBox "Cannot Select more then one Emp"
k = 0
For k = 0 To List1.ListCount - 1
List1.Selected(k) = False
Next k
 List2.Clear
 List3.Clear
End If

Dim f As Integer
f = 0
                For f = 0 To List1.ListCount - 1
                If List1.Selected(f) = True Then
                nm = Split(List1.List(f), "  -  ", Len(List1.List(f)), vbTextCompare)
                End If
                Next f
                
                
                
Dim ts As New ADODB.Recordset
If ts.State Then ts.Close
ts.Open "select * from timesheet where t_empno = '" & nm(0) & "' and t_empname='" & nm(1) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' order by t_r_date ", Cn, 3, 2
f = 0
While Not ts.EOF

List2.List(f) = ts!t_r_date & "  -  " & ts!t_r_hrs & "  -  " & ts!t_r_job
f = f + 1
ts.MoveNext
Wend


'ot
 
Dim tso As New ADODB.Recordset
If tso.State Then tso.Close
tso.Open "select * from timesheet where t_empno = '" & nm(0) & "' and t_empname='" & nm(1) & "' and t_r_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' order by t_r_date ", Cn, 3, 2
g = 0
While Not tso.EOF

List3.List(g) = tso!t_r_date & "  -  " & tso!t_o_hrs & "  -  " & tso!t_o_job
g = g + 1
tso.MoveNext
Wend

k = 0: i = 0
For k = 0 To List3.ListCount - 1
List3.Selected(k) = True
Next k
For i = 0 To List2.ListCount - 1
List2.Selected(i) = True
Next i
Y = 1

assad:

End Sub

Private Sub List2_ItemCheck(Item As Integer)
 
Dim x As String
Dim a As Integer
a = 0

            For a = 0 To List4.ListCount - 1
            If List4.Selected(a) = True Then
            x = List4.List(a)
            End If
            Next a
             
           
               If List2.Selected(Item) = False Then
               List2.Selected(Item) = True
               nm = Split(List2.List(Item), "  -  ", Len(List2.List(Item)), vbTextCompare)
               List2.List(Item) = nm(0) & "  -  " & txt_rg.Text & "  -  " & x
               End If
 
End Sub

Private Sub List3_ItemCheck(Item As Integer)
 
Dim xx As String
Dim aa As Integer
aa = 0

            For aa = 0 To List4.ListCount - 1
            If List5.Selected(aa) = True Then
            xx = List5.List(aa)
            End If
            Next aa
             
           
               If List3.Selected(Item) = False Then
               List3.Selected(Item) = True
               nm = Split(List3.List(Item), "  -  ", Len(List3.List(Item)), vbTextCompare)
               List3.List(Item) = nm(0) & "  -  " & txt_ot.Text & "  -  " & xx
               End If
            
 
End Sub

 
