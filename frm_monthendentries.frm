VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_monthendentries 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MonthEnd Entries  (Earnings/Deductions)"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   ScaleHeight     =   10140
   ScaleWidth      =   14895
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   14775
      TabIndex        =   0
      Top             =   600
      Width           =   14775
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
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
      Begin VB.PictureBox Picture2 
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
         Left            =   1800
         ScaleHeight     =   375
         ScaleWidth      =   3255
         TabIndex        =   2
         Top             =   0
         Width           =   3255
         Begin MSComCtl2.DTPicker dtp_tl 
            Height          =   315
            Left            =   1800
            TabIndex        =   3
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            Format          =   67174403
            CurrentDate     =   38733
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Month"
            Height          =   195
            Left            =   1200
            TabIndex        =   4
            Top             =   0
            Width           =   450
         End
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Notes"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10080
      TabIndex        =   12
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BondStore"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9120
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cash Adv"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PhoneCall"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acr TOff(hrs)"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Travel"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Classification"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Emp No  -  Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frm_monthendentries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dtp_tl_Change()
Unload vscrollmonthendentries
vscrollmonthendentries.Show
vscrollmonthendentries.Left = 0
vscrollmonthendentries.Top = 0
 
SetParent vscrollmonthendentries.HWnd, frm_monthendentries.Picture1.HWnd

End Sub

Private Sub dtp_tl_Click()
Unload vscrollmonthendentries
vscrollmonthendentries.Show
vscrollmonthendentries.Left = 0
vscrollmonthendentries.Top = 0
 
SetParent vscrollmonthendentries.HWnd, frm_monthendentries.Picture1.HWnd

End Sub

Private Sub Form_Load()
On Error Resume Next
dtp_tl.Value = Format(Date, "MM/yyyy")
Unload vscrollmonthendentries
vscrollmonthendentries.Show
vscrollmonthendentries.Left = 0
vscrollmonthendentries.Top = 0
 
SetParent vscrollmonthendentries.HWnd, frm_monthendentries.Picture1.HWnd

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Caption = "Save" Then

Dim j As Integer
j = 0
For j = 0 To vscrollmonthendentries.txt_name.Count - 1
If vscrollmonthendentries.txt_name(j).Text <> "" Then
sp = Split(vscrollmonthendentries.txt_name(j).Text, "  -  ", Len(vscrollmonthendentries.txt_name(j).Text), vbTextCompare)
     Cn.Execute "delete from timelog where month(mnth) ='" & Format(dtp_tl.Value, "MM") & "' and year(mnth) ='" & Format(dtp_tl.Value, "yyyy") & "' and empno='" & sp(0) & "'"
                Dim sv As New ADODB.Recordset
                If sv.State Then sv.Close
                sv.Open "select * from timelog", Cn, 3, 2
                sv.AddNew
                sv!empno = sp(0)
                sv!Name = sp(1)
                sv!classification = vscrollmonthendentries.txt_classification(j).Text
                sv!mnth = dtp_tl.Value
                sv!travel = vscrollmonthendentries.txt_travel(j).Text
                sv!actoff = vscrollmonthendentries.txt_actoff(j).Text
                sv!phonecall = vscrollmonthendentries.txt_phonecall(j).Text
                sv!cashadvance = vscrollmonthendentries.txt_cashadvance(j).Text
                sv!bondstore = vscrollmonthendentries.txt_bondstore(j).Text
                sv!Notes = vscrollmonthendentries.txt_notes(j).Text
                sv!u_date = Now
                sv!t_user = main.Label2.Caption
                sv!t_date = Now
                sv.Update
                sv.Close

End If
Next j
MsgBox "Saved Successfully"
ElseIf Button.Caption = "Close" Then
Unload Me
End If



End Sub
