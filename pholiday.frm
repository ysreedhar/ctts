VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form pholiday 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Public Holiday"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txt_title 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txt_notes 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   1080
         Width           =   5535
      End
      Begin MSComCtl2.DTPicker DTP_tdate 
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67174401
         CurrentDate     =   38733
      End
      Begin MSComCtl2.DTPicker dtp_pholiday 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67174401
         CurrentDate     =   38733
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remarks"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   195
         Left            =   1680
         TabIndex        =   7
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         Height          =   195
         Left            =   4440
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " Date"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   390
      End
   End
End
Attribute VB_Name = "pholiday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
dtp_pholiday.Value = Format(Date, "dd/MM/yyyy")
End Sub
