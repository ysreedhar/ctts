VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_exceltoflex 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Import"
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtfilename 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmd_open 
         BackColor       =   &H8000000E&
         Caption         =   "Select File"
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Download Excel OCX"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " Free download of Excel OCX! A powerful ActiveX control for exchanging data between VB and Excel via COM technology "
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   13
      BackColor       =   16777215
      BackColorFixed  =   16744576
      BackColorBkg    =   16777215
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   720
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   8
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frm_exceltoflex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub cmd_open_Click()
On Error Resume Next
 
cdOpen.ShowOpen
    
    If Not vbCancel Then
       txtfilename = cdOpen.FileName
    End If
End Sub

Private Sub cmd_save_Click()
On Error Resume Next
Dim p As Double
p = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from employee ", Cn, 3, 2
For p = 2 To MSFlexGrid1.Rows
   If MSFlexGrid1.TextMatrix(p, 1) <> "" Then
        
        fldata.AddNew
        fldata!emp_icno = MSFlexGrid1.TextMatrix(p, 0)
        fldata!emp_no = MSFlexGrid1.TextMatrix(p, 1)
        fldata!emp_name = MSFlexGrid1.TextMatrix(p, 2)
        fldata!emp_sex = MSFlexGrid1.TextMatrix(p, 3)
        fldata!emp_dob = MSFlexGrid1.TextMatrix(p, 4)
        fldata!emp_age = Round(MSFlexGrid1.TextMatrix(p, 5), 2)
        fldata!emp_nationality = MSFlexGrid1.TextMatrix(p, 6)
        fldata!emp_classification = MSFlexGrid1.TextMatrix(p, 7)
        fldata!emp_joindate = MSFlexGrid1.TextMatrix(p, 8)
        fldata!emp_coy = MSFlexGrid1.TextMatrix(p, 9)
        If MSFlexGrid1.TextMatrix(p, 10) = "" Then
        fldata!emp_chargetype = "-"
        Else
        fldata!emp_chargetype = MSFlexGrid1.TextMatrix(p, 10)
        End If
        If MSFlexGrid1.TextMatrix(p, 11) = "" Then
        fldata!emp_traveltime = "-"
        Else
        fldata!emp_traveltime = MSFlexGrid1.TextMatrix(p, 11)
        End If
        If MSFlexGrid1.TextMatrix(p, 12) = "" Then
        fldata!Notes = "-"
        Else
        fldata!Notes = MSFlexGrid1.TextMatrix(p, 12)
        End If
        fldata.Update
    End If
Next p
MsgBox "Updated"
End Sub

Private Sub Command1_Click()
Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
'Set objWorkbook = objExcel.Workbooks.Open(App.Path & "\test.xls")
Set objWorkbook = objExcel.Workbooks.Open(txtfilename.Text)
Set objWorksheet = objWorkbook.ActiveSheet

With MSFlexGrid1
.Cols = 15
.Rows = 540
For i = 1 To .Rows - 1
    .Row = i
    For n = 0 To .Cols - 1
        .Col = n
        .Text = objWorksheet.Cells(i + 1, n + 1).Value
    Next
Next
End With
MsgBox "Imported"
AppActivate Me.Caption
End Sub

'Private Sub Command2_Click()
'Dim i As Long
'i = ShellExecute(Form1.HWnd, "open", "http://download.com.com/3000-2401-10105891.html?tag=lst-0-9", vbNullString, vbNullString, 1)
'End Sub

 

Private Sub Form_Load()
Me.Top = 10
Me.Left = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objWorkbook = Nothing
Set objExcel = Nothing
End Sub
