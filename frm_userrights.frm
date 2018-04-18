VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_userrights 
   BackColor       =   &H00FFFFFF&
   Caption         =   "USER RIGHTS"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   11520
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Location List"
      Height          =   4095
      Left            =   8160
      TabIndex        =   21
      Top             =   960
      Width           =   3255
      Begin VB.ListBox List6 
         Height          =   3660
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Others"
      Height          =   4095
      Left            =   8280
      TabIndex        =   18
      Top             =   5160
      Width           =   3135
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   120
         Width           =   255
      End
      Begin VB.ListBox List5 
         Height          =   3660
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forms"
      Height          =   4095
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   8175
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   15
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   120
         Width           =   255
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transaction Forms"
         Height          =   3735
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   4095
         Begin VB.ListBox List2 
            Height          =   3435
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Master Settings"
         Height          =   3735
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3735
         Begin VB.ListBox List1 
            Height          =   3435
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   240
            Width           =   3375
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reports"
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   5160
      Width           =   8175
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   17
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   120
         Width           =   255
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TransactionReports"
         Height          =   3735
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   4095
         Begin VB.ListBox List4 
            Height          =   3435
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   10
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Master Reports"
         Height          =   3735
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3735
         Begin VB.ListBox List3 
            Height          =   3435
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   240
            Width           =   3375
         End
      End
   End
   Begin VB.ComboBox cbo_user 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   635
      ButtonWidth     =   1402
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "grd"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8200
         ScaleHeight     =   375
         ScaleWidth      =   2295
         TabIndex        =   13
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   58
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_userrights.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DC7E5A&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frm_userrights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double
Dim m As Double
Dim n As Double
Private Sub cbo_user_Click()
On Error Resume Next
                                i = 0
                                For i = 0 To List1.ListCount - 1
                                List1.Selected(i) = False
                                Next i
                                 
                                
                                j = 0
                                For j = 0 To List2.ListCount - 1
                                List2.Selected(j) = False
                                Next j
                                 
                                
                                k = 0
                                For k = 0 To List3.ListCount - 1
                                List3.Selected(k) = False
                                Next k
                                 
                               
                                l = 0
                                For l = 0 To List4.ListCount - 1
                                List4.Selected(l) = False
                                Next l
                                
                                m = 0
                                For m = 0 To List5.ListCount - 1
                                List5.Selected(m) = False
                                Next m
                                
                                n = 0
                                For n = 0 To List6.ListCount - 1
                                List6.Selected(n) = False
                                Next n


Dim Y As Double
Dim d As Double
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select * from userrights where u_name='" & cbo_user.Text & "' ", Cn, 3, 2
If Not rs1.EOF Then
Y = 0
For Y = 0 To List1.ListCount - 1
d = 0
d = Mid(rs1!mforms, Y + 1, 1)
If d = 1 Then
List1.Selected(Y) = True
End If
Next Y


Y = 0
For Y = 0 To List2.ListCount - 1
d = 0
d = Mid(rs1!tforms, Y + 1, 1)
If d = 1 Then
List2.Selected(Y) = True
End If
Next Y


Y = 0
For Y = 0 To List3.ListCount - 1
d = 0
d = Mid(rs1!mreports, Y + 1, 1)
If d = 1 Then
List3.Selected(Y) = True
End If
Next Y



Y = 0
For Y = 0 To List4.ListCount - 1
d = 0
d = Mid(rs1!treports, Y + 1, 1)
If d = 1 Then
List4.Selected(Y) = True
End If
Next Y


Y = 0
For Y = 0 To List5.ListCount - 1
d = 0
d = Mid(rs1!others, Y + 1, 1)
If d = 1 Then
List5.Selected(Y) = True
End If
Next Y

End If
Dim uj As New ADODB.Recordset
If uj.State Then uj.Close
uj.Open "select * from userproject where username='" & cbo_user.Text & "'", Cn, 3, 2
While Not uj.EOF
d = 0
For d = 0 To List6.ListCount - 1
jkl = Split(List6.List(d), "  -  ", Len(List6.List(d)), vbTextCompare)
If jkl(0) = uj!project Then
List6.Selected(d) = True
End If
Next d
uj.MoveNext
Wend
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
i = 0
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next i
Else
i = 0
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next i
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
i = 0
For i = 0 To List2.ListCount - 1
List2.Selected(i) = True
Next i
Else
i = 0
For i = 0 To List2.ListCount - 1
List2.Selected(i) = False
Next i
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
i = 0
For i = 0 To List3.ListCount - 1
List3.Selected(i) = True
Next i
Else
i = 0
For i = 0 To List3.ListCount - 1
List3.Selected(i) = False
Next i
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
i = 0
For i = 0 To List4.ListCount - 1
List4.Selected(i) = True
Next i
Else
i = 0
For i = 0 To List4.ListCount - 1
List4.Selected(i) = False
Next i
End If
End Sub
Private Sub Check5_Click()
If Check5.Value = 1 Then
i = 0
For i = 0 To List5.ListCount - 1
List5.Selected(i) = True
Next i
Else
i = 0
For i = 0 To List5.ListCount - 1
List5.Selected(i) = False
Next i
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
i = 0
For i = 0 To List6.ListCount - 1
List6.Selected(i) = True
Next i
Else
i = 0
For i = 0 To List6.ListCount - 1
List6.Selected(i) = False
Next i
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 5
Me.Left = 5
main.lbltitle.Caption = "USER RIGHTS"
Dim ud As New ADODB.Recordset
If ud.State Then ud.Close
ud.Open "select * from userid order by a_userid", Cn, 3, 2
While Not ud.EOF
cbo_user.AddItem ud!a_userid
ud.MoveNext
Wend
'master forms

 
 List1.AddItem "Company" '6
 List1.AddItem "HrType" '6
 List1.AddItem "Room" '6
 List1.AddItem "LifeRaft" '7
 List1.AddItem "Nationality" '8
 List1.AddItem "Public Holiday" '7
 List1.AddItem "Location" '8
 List1.AddItem "JobClassification" '8
 List1.AddItem "JobNo" '9
 List1.AddItem "Employee" '10
 
 'transaction forms
 
 ' Transaction
 'List2.AddItem "TRANSACTIONS" '1

 
 List2.AddItem "OnBoard" '2
 List2.AddItem "OffBoard" '2
 List2.AddItem "MonthEnd Entries" '3
 List2.AddItem "Daily Attendance" '4
 List2.AddItem "Update Attendance" '4
 List2.AddItem "Excel To TimeSheet" '5
 List2.AddItem "TimeSheet To Excel" '6
 
 
 List4.AddItem "Personnel OnBoard" '2
 List4.AddItem "Personnel OffBoard" '2
 List4.AddItem "Individual TimeSheet" '2
 List4.AddItem "TimeSheet Summary" '3
 
 
 List5.AddItem "Utilities" '2
 List5.AddItem "Backup" '2
 List5.AddItem "Restore" '3
 List5.AddItem "Send Message" '4
 List5.AddItem "Administration" '5
 List5.AddItem "Company Parameter" '6
 List5.AddItem "Create User" '7
 List5.AddItem "User Rights" '7
 List5.AddItem "UserID" '7
 List5.AddItem "Logout" '13
 
 
 
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(bg_name),notes  from  badge   order by bg_name", Cn, 3, 2
While Not pr.EOF
List6.AddItem pr(0) & "  -  " & pr(1)
pr.MoveNext
Wend
pr.Close
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 If Button.Caption = "Save" Then
   If cbo_user.Text = "" Then
   MsgBox "Select User Name"
   Exit Sub
   End If
 
 If rs.State Then rs.Close
 rs.Open "select * from userrights where u_name='" & cbo_user.Text & "' ", Cn, 3, 2
If rs.EOF Then
rs.AddNew
Call rightuser
Else
Cn.Execute "delete from userrights where u_name='" & cbo_user.Text & "' "
rs.AddNew
Call rightuser
End If
  
  Else
  Unload Me
  
  End If
End Sub

Public Sub rightuser()
Dim a(35) As Double
a(35) = 0
Dim b(40) As Double
b(35) = 0
Dim c(40) As Double
c(35) = 0
Dim d(40) As Double
d(35) = 0
Dim f(35) As Double
f(35) = 0
 
rs!u_name = cbo_user.Text
 
Dim ii As Double
    ii = 0
    For ii = 0 To List1.ListCount - 1
    If List1.Selected(ii) = True Then
    a(ii) = 1
    Else
    a(ii) = 0
    End If
    Next ii
 rs!mforms = a(0) & a(1) & a(2) & a(3) & a(4) & a(5) & a(6) & a(7) & a(8) & a(9) & a(10) & a(11) & a(12) & a(13) & a(14) & a(15) & a(16) & a(17) & a(18) & a(19) & a(20) & a(21) & a(22) & a(23) & a(24) & a(25) & a(26) & a(27) & a(28)
 
 
 Dim jj As Double
   jj = 0
    For jj = 0 To List2.ListCount - 1
    If List2.Selected(jj) = True Then
    b(jj) = 1
    Else
    b(jj) = 0
    End If
    Next jj
 rs!tforms = b(0) & b(1) & b(2) & b(3) & b(4) & b(5) & b(6) & b(7) & b(8) & b(9) & b(10) & b(11) & b(12) & b(13) & b(14) & b(15) & b(16) & b(17) & b(18) & b(19) & b(20) & b(21) & b(22) & b(23) & b(24) & b(25) & b(26) & b(27) & b(28) & b(29) & b(30) & b(31) & b(32) & b(33) & b(34) & b(35) & b(36) & b(37)
 
 
  Dim kk As Double
    kk = 0
    For kk = 0 To List3.ListCount - 1
    If List3.Selected(kk) = True Then
    c(kk) = 1
    Else
    c(kk) = 0
    End If
    Next kk
 rs!mreports = c(0) & c(1) & c(2) & c(3) & c(4) & c(5) & c(6) & c(7) & c(8) & c(9) & c(10) & c(11) & c(12) & c(13) & c(14) & c(15) & c(16) & c(17) & c(18) & c(19) & c(20) & c(21) & c(22) & c(23) & c(24) & c(25) & c(26) & c(27) & c(28) & c(29) & c(30) & c(31)
 
 
 
   Dim ll As Double
    ll = 0
    For ll = 0 To List4.ListCount - 1
    If List4.Selected(ll) = True Then
    d(ll) = 1
    Else
    d(ll) = 0
    End If
    Next ll
 rs!treports = d(0) & d(1) & d(2) & d(3) & d(4) & d(5) & d(6) & d(7) & d(8) & d(9) & d(10) & d(11) & d(12) & d(13) & d(14) & d(15) & d(16) & d(17) & d(18) & d(19) & d(20) & d(21) & d(22) & d(23) & d(24) & d(25) & d(26) & d(27) & d(28) & d(29) & d(30) & d(31) & d(32) & d(33) & d(34) & d(35) & d(36) & d(37) & d(38)
 
 
 
 
   Dim mm As Double
    mm = 0
    For mm = 0 To List5.ListCount - 1
    If List5.Selected(mm) = True Then
    f(mm) = 1
    Else
    f(mm) = 0
    End If
    Next mm
 rs!others = f(0) & f(1) & f(2) & f(3) & f(4) & f(5) & f(6) & f(7) & f(8) & f(9) & f(10) & f(11) & f(12) & f(13) & f(14) & f(15) & f(16) & f(17) & f(18)
 
 rs.Update
 Cn.Execute "delete from userproject where username='" & cbo_user.Text & "'"
 Dim q As Double
 q = 0
 For q = 0 To List6.ListCount - 1
 If List6.Selected(q) = True Then
 gh = Split(List6.List(q), "  -  ", Len(List6.List(q)), vbTextCompare)
 Dim rss As New ADODB.Recordset
 If rss.State Then rss.Close
 rss.Open "select * from userproject", Cn, 3, 2
 rss.AddNew
 rss!UserName = cbo_user.Text
 rss!project = gh(0)
 rss.Update
 End If
 Next q
 MsgBox "User Rights for " & cbo_user.Text & " Saved"
                                 
                                i = 0
                                For i = 0 To List1.ListCount - 1
                                List1.Selected(i) = False
                                Next i
                                 
                                
                                j = 0
                                For j = 0 To List2.ListCount - 1
                                List2.Selected(j) = False
                                Next j
                                 
                                
                                k = 0
                                For k = 0 To List3.ListCount - 1
                                List3.Selected(k) = False
                                Next k
                                 
                               
                                l = 0
                                For l = 0 To List4.ListCount - 1
                                List4.Selected(l) = False
                                Next l
                                
                                 m = 0
                                For m = 0 To List5.ListCount - 1
                                List5.Selected(m) = False
                                Next m
                                
                                
                                 n = 0
                                For n = 0 To List6.ListCount - 1
                                List6.Selected(n) = False
                                Next n
                                
End Sub


