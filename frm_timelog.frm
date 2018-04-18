VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_timelog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   17171
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   16744576
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   635
      ButtonWidth     =   1561
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "ar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "grd"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
            Key             =   "hlp"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         TabIndex        =   2
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
            Picture         =   "frm_timelog.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_timelog.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_timelog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_close_Click()

End Sub

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub flex_grid_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbYellow
Next
flex_grid.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture
 


vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbYellow
Next
flex_grid.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture




'--END---------


Unload timelog
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

timelog.cbo_name.Text = Trim(flex_grid.TextMatrix(flex_grid.Row, 1)) & "  -  " & Trim(flex_grid.TextMatrix(flex_grid.Row, 2))
timelog.txt_classification.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
timelog.dtp_tl.Value = flex_grid.TextMatrix(flex_grid.Row, 4)
timelog.txt_travel.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
timelog.txt_actoff.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
timelog.txt_phonecall.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
timelog.txt_cashadvance.Text = flex_grid.TextMatrix(flex_grid.Row, 8)
timelog.txt_bondstore.Text = flex_grid.TextMatrix(flex_grid.Row, 9)
timelog.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 10)
 
timelog.Show
timelog.Top = 3200
timelog.Left = 0
timelog.Height = 4305
timelog.Width = 6300


 
vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "TimeLog"
Call flex_title
Call flex_data
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5
End Sub
Public Sub flex_title()
On Error Resume Next

   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "EmpNo"
        .ColWidth(1) = 1200
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Name"
        .ColWidth(2) = 1200
        .ColAlignment(2) = 0
       
        .TextMatrix(0, 3) = "Classification"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        
        .TextMatrix(0, 4) = "Month"
        .ColWidth(4) = 1200
        .ColAlignment(4) = 0
       
        .TextMatrix(0, 5) = "Travel"
        .ColWidth(5) = 1200
        .ColAlignment(5) = 0
        
         .TextMatrix(0, 6) = "Acr TOffhrs"
        .ColWidth(6) = 1200
        .ColAlignment(6) = 0
       
        .TextMatrix(0, 7) = "Phone Call"
        .ColWidth(7) = 1200
        .ColAlignment(7) = 0
        
        .TextMatrix(0, 8) = "Cash Advance"
        .ColWidth(8) = 1200
        .ColAlignment(8) = 0
       
        .TextMatrix(0, 9) = "BondStore"
        .ColWidth(9) = 1200
        .ColAlignment(9) = 0
        
        .TextMatrix(0, 10) = "Remarks"
        .ColWidth(10) = 4000
        .ColAlignment(10) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload timelog
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
 
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload timelog
timelog.Show
timelog.Top = 3200
timelog.Left = 0
timelog.Height = 4305
timelog.Width = 6300
' to save new record
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
'validate
nm = Split(timelog.cbo_name.Text, "  -  ", Len(timelog.cbo_name.Text), vbTextCompare)
 
 
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from timelog", Cn, 3, 2
sv.AddNew
sv!empno = nm(0)
sv!Name = nm(1)
sv!classification = timelog.txt_classification.Text
sv!mnth = timelog.dtp_tl.Value
sv!travel = timelog.txt_travel.Text
sv!actoff = timelog.txt_actoff.Text
sv!phonecall = timelog.txt_phonecall.Text
sv!cashadvance = timelog.txt_cashadvance.Text
sv!bondstore = timelog.txt_bondstore.Text
sv!Notes = timelog.txt_notes.Text
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv!t_date = timelog.DTP_tdate.Value
sv.Update
sv.Close
MsgBox "New Record Added Succesfully"
Unload timelog
Call flex_data
Call flex_title
Exit Sub
assad:
       
       MsgBox "Duplicate Entries Not Allowed"
'to modify existing record
ElseIf Button.Caption = "Modify" Then
On Error GoTo assad1
'validate

 nm = Split(timelog.cbo_name.Text, "  -  ", Len(timelog.cbo_name.Text), vbTextCompare)
 
Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from timelog where tl_id=" & id1, Cn, 3, 2
If Not md.EOF Then
md!empno = nm(0)
md!Name = nm(1)
md!classification = timelog.txt_classification.Text
md!mnth = timelog.dtp_tl.Value
md!travel = timelog.txt_travel.Text
md!actoff = timelog.txt_actoff.Text
md!phonecall = timelog.txt_phonecall.Text
md!cashadvance = timelog.txt_cashadvance.Text
md!bondstore = timelog.txt_bondstore.Text
md!Notes = timelog.txt_notes.Text
md!u_date = Now
md!t_user = main.Label2.Caption
md!t_date = timelog.DTP_tdate.Value
 
md.Update
md.Close
MsgBox "Selected Record Modified"
End If

Unload timelog
Call flex_data
Call flex_title
Exit Sub
assad1:
       
       MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete" Then
 



Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id2 As Double
id2 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
Cn.Execute "delete from timelog where tl_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload timelog
Call flex_data
Call flex_title
Else
Unload timelog
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload timelog
End If




End Sub

Public Sub flex_data()
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select t.tl_id,t.empno,t.name,t.classification,t.mnth,t.travel,t.actoff,t.phonecall,t.cashadvance,t.bondstore,t.notes from timelog t,onboard b where t.empno=b.ob_empno  and b.location='" & main.lbllocation & "' order by t.mnth,t.name", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = fldata(0)
                    .TextMatrix(.Rows - 1, 1) = fldata(1)
                    .TextMatrix(.Rows - 1, 2) = fldata(2)
                    .TextMatrix(.Rows - 1, 3) = fldata(3)
                    .TextMatrix(.Rows - 1, 4) = Format(fldata(4), "MM/yyyy")
                    .TextMatrix(.Rows - 1, 5) = fldata(5)
                    .TextMatrix(.Rows - 1, 6) = fldata(6)
                    .TextMatrix(.Rows - 1, 7) = fldata(7)
                    .TextMatrix(.Rows - 1, 8) = fldata(8)
                    .TextMatrix(.Rows - 1, 9) = fldata(9)
                    .TextMatrix(.Rows - 1, 10) = fldata(10)
        fldata.MoveNext
    Wend
End With
End Sub





