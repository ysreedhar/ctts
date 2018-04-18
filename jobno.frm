VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form jobno 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "JobNo."
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txt_jobno 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txt_notes 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   1200
         Width           =   5535
      End
      Begin MSComCtl2.DTPicker DTP_tdate 
         Height          =   315
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   28377089
         CurrentDate     =   38733
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         Height          =   195
         Left            =   4320
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JobNo."
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Proj / DOE"
         Height          =   195
         Left            =   3480
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
   End
End
Attribute VB_Name = "jobno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
cbo_proj.Text = "PROJECT"
cbo_proj.AddItem "PROJECT"
cbo_proj.AddItem "DOE"
End Sub
