VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_pob 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   10065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   17171
      _Version        =   393216
      Rows            =   3
      Cols            =   9
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
      Width           =   12030
      _ExtentX        =   21220
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
            Picture         =   "frm_pob.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_pob.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_pob"
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


Unload onboard
onboard.Show
SetParent onboard.HWnd, frm_pob.HWnd
onboard.Top = 200
onboard.Left = 200
onboard.Height = 9120
onboard.Width = 11070

nu = Split((flex_grid.TextMatrix(flex_grid.Row, 1)), "  -  ", Len((flex_grid.TextMatrix(flex_grid.Row, 1))), vbTextCompare)

Dim emp As New ADODB.Recordset
If emp.State Then emp.Close
emp.Open "select Distinct(e.emp_no),e.emp_name,e.emp_coy,e.emp_chargetype from employee  e, company c  where e.emp_coy=c.coy_name and  e.emp_status = 'y'  and e.emp_no = '" & flex_grid.TextMatrix(flex_grid.Row, 1) & "' ", Cn, 3, 2
While Not emp.EOF
If emp(1) = "" Then
onboard.List1.AddItem emp(0)
onboard.cbo_company.Text = emp(2)
onboard.cbo_proj.Text = emp(3)
Else
onboard.List1.AddItem emp(0) & "  -  " & emp(1)
onboard.cbo_company.Text = emp(2)
onboard.cbo_proj.Text = emp(3)
End If
emp.MoveNext
Wend



Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

 Dim m As Integer
 m = 0
 For m = 0 To onboard.List1.ListCount - 1
  nn = Split(onboard.List1.List(m), "  -  ", Len(onboard.List1.List(m)), vbTextCompare)
  If nn(0) = flex_grid.TextMatrix(flex_grid.Row, 1) Then
       onboard.List1.Selected(m) = True
  End If
 Next m
 onboard.cbo_type.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
 onboard.dtp_onboard.Value = flex_grid.TextMatrix(flex_grid.Row, 4)
 onboard.cbo_shift.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
 m = 0
 For m = 0 To onboard.List2.ListCount - 1
  nm = Split(onboard.List2.List(m), "          -          ", Len(onboard.List2.List(m)), vbTextCompare)
  If nm(0) = flex_grid.TextMatrix(flex_grid.Row, 6) Then
       onboard.List2.Selected(m) = True
  End If
 Next m
 m = 0
 For m = 0 To onboard.List3.ListCount - 1
  nmm = Split(onboard.List3.List(m), "          -          ", Len(onboard.List3.List(m)), vbTextCompare)
  If nmm(0) = flex_grid.TextMatrix(flex_grid.Row, 7) Then
       onboard.List3.Selected(m) = True
  End If
 Next m
'onboard.txt_traveltime.Text = flex_grid.TextMatrix(flex_grid.Row, 8)
onboard.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 9)

 
 



 
vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "P.O.B"
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
        
         .TextMatrix(0, 2) = "Emp Name"
        .ColWidth(2) = 4000
        .ColAlignment(2) = 0
        
                .TextMatrix(0, 3) = "POB Type"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        
                .TextMatrix(0, 4) = "Date & Time"
        .ColWidth(4) = 2000
        .ColAlignment(4) = 0
        
                .TextMatrix(0, 5) = "Shift"
        .ColWidth(5) = 800
        .ColAlignment(5) = 0
        
                .TextMatrix(0, 6) = "Room No"
        .ColWidth(6) = 800
        .ColAlignment(6) = 0
        
                .TextMatrix(0, 7) = "LifeRaft"
        .ColWidth(7) = 800
        .ColAlignment(7) = 0
        
                .TextMatrix(0, 8) = "Travel Time"
        .ColWidth(8) = 1000
        .ColAlignment(8) = 0
       
        .TextMatrix(0, 9) = "Notes"
        .ColWidth(9) = 4000
        .ColAlignment(9) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload onboard
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
 
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload onboard
onboard.Show
SetParent onboard.HWnd, frm_pob.HWnd
onboard.Top = 200
onboard.Left = 200
onboard.Height = 9120
onboard.Width = 11070
' to save new record
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
'validate
 
 
On Error Resume Next
Dim ob As New ADODB.Recordset
Dim l As Integer
Dim m As Integer
Dim k As Integer
k = 0: m = 0: l = 0
 
For l = 0 To onboard.List2.ListCount - 1
If onboard.List2.Selected(l) = True Then
rmn = Split(onboard.List2.List(l), "          -          ", Len(onboard.List2.List(l)), vbTextCompare)
End If
Next l
For m = 0 To onboard.List3.ListCount - 1
If onboard.List3.Selected(m) = True Then
rft = Split(onboard.List3.List(m), "          -          ", Len(onboard.List3.List(m)), vbTextCompare)
End If
Next m

For k = 0 To onboard.List1.ListCount - 1
If onboard.List1.Selected(k) = True Then
nm = Split(onboard.List1.List(k), "  -  ", Len(onboard.List1.List(k)), vbTextCompare)
                If ob.State Then ob.Close
                ob.Open "select * from onboard ", Cn, 3, 2
                ob.AddNew
                ob!ob_empno = nm(0)
                ob!ob_empname = nm(1)
                ob!ob_type = onboard.cbo_type.Text
                ob!ob_dateonboard = onboard.dtp_onboard.Value
                ob!ob_shift = onboard.cbo_shift.Text
                ob!ob_roomno = rmn(0)
                ob!ob_raftno = rft(0)
                
                Dim mt As Integer

 If ep.State Then ep.Close
 ep.Open "select * from employee where emp_no = '" & nm(0) & "' and emp_traveltime='Applicable' ", Cn, 3, 2
 If Not ep.EOF Then
 ob!ob_traveltime = 8
 Else
 ob!ob_traveltime = 0
 End If
                
                ob!ob_notes = onboard.txt_notes.Text
                ob!Location = main.lbllocation.Caption
                ob.Update
                ob.Close
                Cn.Execute "update employee set emp_status = 'y' where emp_no='" & nm(0) & "' "
                

End If
Next k
List1.Clear
Dim emp1 As New ADODB.Recordset
If emp1.State Then emp1.Close
emp1.Open "select Distinct(emp_no),emp_name from employee where emp_status = 'x' order by emp_no", Cn, 3, 2
While Not emp1.EOF
If emp1(1) = "" Then
List1.AddItem emp1(0)
Else
List1.AddItem emp1(0) & "  -  " & emp1(1)
End If
emp1.MoveNext
Wend
MsgBox "POB Processed Successfully"
'
'assad:
'      MsgBox "Employee is Already On Board"Unload onboard
Unload onboard
Call flex_data
Call flex_title
Exit Sub
assad:
       
       MsgBox "Duplicate Entries Not Allowed"
'to modify existing record
ElseIf Button.Caption = "Modify" Then
On Error GoTo assad1
'validate
l = 0
For l = 0 To onboard.List2.ListCount - 1
If onboard.List2.Selected(l) = True Then
rmn = Split(onboard.List2.List(l), "          -          ", Len(onboard.List2.List(l)), vbTextCompare)
End If
Next l
m = 0
For m = 0 To onboard.List3.ListCount - 1
If onboard.List3.Selected(m) = True Then
rft = Split(onboard.List3.List(m), "          -          ", Len(onboard.List3.List(m)), vbTextCompare)
End If
Next m
Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)

ns = Split(onboard.List1.List(k), "  -  ", Len(onboard.List1.List(k)), vbTextCompare)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from onboard where ob_id=" & id1, Cn, 3, 2
If Not md.EOF Then
                 
                md!ob_type = onboard.cbo_type.Text
                md!ob_shift = onboard.cbo_shift.Text
                md!ob_roomno = rmn(0)
                md!ob_raftno = rft(0)
                
                Dim ep1 As ADODB.Recordset
                           

 If ep1.State Then ep1.Close
 ep1.Open "select * from employee where emp_no = '" & ns(0) & "' and emp_traveltime='Applicable' ", Cn, 3, 2
 If Not ep1.EOF Then
 md!ob_traveltime = 8
 Else
 md!ob_traveltime = 0
 End If
                 
                md!ob_notes = onboard.txt_notes.Text
                md!Location = main.lbllocation.Caption
                md.Update
                md.Close
                Cn.Execute "update employee set emp_status = 'y' where emp_no='" & ns(0) & "' "
 
 
MsgBox "Selected Record Modified"
End If

Unload onboard
Call flex_data
Call flex_title
Exit Sub
assad1:
       
       MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete" Then
'''Dim dlk As New ADODB.Recordset
'''If dlk.State Then dlk.Close
'''dlk.Open "select * from employee where emp_classification ='" & flex_grid.TextMatrix(flex_grid.Row, 2) & "'", Cn, 3, 2
'''If Not dlk.EOF Then
'''MsgBox "Cannot Delete This Record"
'''Exit Sub
'''End If
'''


Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
            If dlt = vbYes Then
            Dim id2 As Double
            id2 = 0
            
                    If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
                    
                    id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
                    Cn.Execute "delete from onboard where ob_id=" & id2
                    MsgBox "Selected onboard Has Been Deleted"
                    Unload onboard
                    Call flex_data
                    Call flex_title
                    Else
                    Unload onboard
                    End If
            
 
ElseIf Button.Caption = "Close" Then
Unload Me
Unload onboard
End If
 
End Sub

Public Sub flex_data()
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select ob_id,ob_empno,ob_empname,ob_type,ob_dateonboard,ob_shift,ob_roomno,ob_raftno,ob_chargetype,ob_traveltime,ob_notes from onboard  where location='" & main.lbllocation.Caption & "' order by ob_dateonboard", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = fldata(0)
                        .TextMatrix(.Rows - 1, 1) = fldata(1)
                        .TextMatrix(.Rows - 1, 2) = fldata(2)
                        .TextMatrix(.Rows - 1, 3) = fldata(3)
                        .TextMatrix(.Rows - 1, 4) = fldata(4)
                        .TextMatrix(.Rows - 1, 5) = fldata(5)
                        nm = Split(fldata(6), "  -  ", Len(fldata(6)), vbTextCompare)
                        .TextMatrix(.Rows - 1, 6) = nm(0)
                        mm = Split(fldata(7), "  -  ", Len(fldata(7)), vbTextCompare)
                        .TextMatrix(.Rows - 1, 7) = mm(0)
                        .TextMatrix(.Rows - 1, 8) = fldata(9)
                        .TextMatrix(.Rows - 1, 9) = fldata(10)
 
        
    
        fldata.MoveNext
    Wend
End With
End Sub




