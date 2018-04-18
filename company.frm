VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form company 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Company Details"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txt_regno 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txt_contactperson 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   5775
      End
      Begin VB.TextBox txt_name 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txt_address 
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   1560
         Width           =   5415
      End
      Begin VB.TextBox txt_notes 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   2640
         Width           =   5535
      End
      Begin MSComCtl2.DTPicker DTP_tdate 
         Height          =   315
         Left            =   4440
         TabIndex        =   6
         Top             =   315
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67436545
         CurrentDate     =   38733
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         Height          =   195
         Left            =   4440
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact person"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg No."
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   570
      End
   End
End
Attribute VB_Name = "company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
End Sub
