VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_flextoexcel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Load "
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtp_from 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   67108865
         CurrentDate     =   38378
      End
      Begin MSComCtl2.DTPicker dtp_to 
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   67108865
         CurrentDate     =   38378
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "To"
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "From"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Download Excel OCX"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Free download of Excel OCX! A powerful ActiveX control for exchanging data between VB and Excel via COM technology "
      Top             =   3720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   13
      BackColor       =   16777215
      BackColorFixed  =   16744576
      BackColorBkg    =   16777215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Copy FlexGrid Contents to Excel"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   8
      Height          =   975
      Left            =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frm_flextoexcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cx As Integer


Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub cmd_load_Click()
Call flex_title
Call flex_data
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
Set objWorkbook = objExcel.Workbooks.Add
AppActivate "FlexGrid To Excel Demo"
For i = 0 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    For n = 0 To 12
        MSFlexGrid1.Col = n
        objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = MSFlexGrid1.Text
    Next
Next
End Sub

'Private Sub Command2_Click()
'Dim i As Long
'i = ShellExecute(Form1.HWnd, "open", "http://download.com.com/3000-2401-10105891.html?tag=lst-0-9", vbNullString, vbNullString, 1)
'End Sub

Private Sub Form_Load()
dtp_from.Value = Date
dtp_to.Value = Date
Me.Top = 10
Me.Left = 10
cx = 0
Dim i As Integer
Dim n As Integer
Me.Caption = "FlexGrid To Excel Demo"
'Populate the FlexGrid with sample data
Call flex_title
Call flex_data
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objWorkbook = Nothing
Set objExcel = Nothing
End Sub


Public Sub flex_title()
On Error Resume Next

   With MSFlexGrid1
        .Row = 0:    .Col = 0
       
        .TextMatrix(0, 0) = "t_id"
        .ColWidth(0) = 1200
        .ColAlignment(0) = 0
        
        .TextMatrix(0, 1) = "t_empno"
        .ColWidth(1) = 1200
        .ColAlignment(1) = 0
       
        .TextMatrix(0, 2) = "t_empname"
        .ColWidth(2) = 1200
        .ColAlignment(2) = 0
        
        .TextMatrix(0, 3) = "t_r_date"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
       
        .TextMatrix(0, 4) = "t_r_hrs"
        .ColWidth(4) = 1200
        .ColAlignment(4) = 0
        
        .TextMatrix(0, 5) = "t_r_job"
        .ColWidth(5) = 1200
        .ColAlignment(5) = 0
       
        .TextMatrix(0, 6) = "t_o_hrs"
        .ColWidth(6) = 1200
        .ColAlignment(6) = 0
        
        .TextMatrix(0, 7) = "t_o_job"
        .ColWidth(7) = 1200
        .ColAlignment(7) = 0
       
        .TextMatrix(0, 8) = "notes"
        .ColWidth(8) = 1200
        .ColAlignment(8) = 0
        
        .TextMatrix(0, 9) = "t_user"
        .ColWidth(9) = 1200
        .ColAlignment(9) = 0
       
        .TextMatrix(0, 10) = "t_date"
        .ColWidth(10) = 1200
        .ColAlignment(10) = 0
        
                .TextMatrix(0, 11) = "u_date"
        .ColWidth(11) = 1200
        .ColAlignment(11) = 0
        
                        .TextMatrix(0, 12) = "daytype"
        .ColWidth(12) = 1200
        .ColAlignment(12) = 0
    End With
End Sub
Public Sub flex_data()
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from timesheet where t_date between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "'  order by t_id", Cn, 3, 2

With MSFlexGrid1
    .Rows = 1
    While Not fldata.EOF
    
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata(1)
        .TextMatrix(.Rows - 1, 2) = fldata(2)
                .TextMatrix(.Rows - 1, 3) = fldata(3)
        .TextMatrix(.Rows - 1, 4) = fldata(4)
        .TextMatrix(.Rows - 1, 5) = fldata(5)
                .TextMatrix(.Rows - 1, 6) = fldata(6)
        .TextMatrix(.Rows - 1, 7) = fldata(7)
        .TextMatrix(.Rows - 1, 8) = fldata(8)
                .TextMatrix(.Rows - 1, 9) = fldata(9)
        .TextMatrix(.Rows - 1, 10) = fldata(10)
        .TextMatrix(.Rows - 1, 11) = fldata(11)
        .TextMatrix(.Rows - 1, 12) = fldata(12)
        fldata.MoveNext
    Wend
End With
End Sub
