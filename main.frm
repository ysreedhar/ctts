VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "POB"
   ClientHeight    =   8925
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14970
   Icon            =   "main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   2100
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picLeftPane 
      Align           =   3  'Align Left
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   8235
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      Begin VB.PictureBox Splitter 
         Height          =   8040
         Left            =   2865
         ScaleHeight     =   8040
         ScaleWidth      =   15
         TabIndex        =   1
         Top             =   -120
         Visible         =   0   'False
         Width           =   15
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   9180
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   16193
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   882
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "imlTree"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList51 
      Left            =   3720
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":105C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1376
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":17C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2116
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2430
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":274A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3308
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3622
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":393C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":18AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1ED48
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":24FE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":252FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":25456
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":25770
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":25BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":25EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":261F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":26648
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":26962
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":26DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2720E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":27528
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":27842
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":27B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":27E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":28190
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":285E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":28A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":28D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29068
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29382
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2969C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29AEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTree 
      Left            =   4155
      Top             =   3315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2A280
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2A6D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2AB26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2AF78
            Key             =   "employee"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2B3CC
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2B820
            Key             =   "open"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2BC74
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2C550
            Key             =   "report"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2C9A4
            Key             =   "shipper"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2D280
            Key             =   "group"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2D6D4
            Key             =   "supplier"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2DB28
            Key             =   "taxonomy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":302DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":305F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3090E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":30C28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   8595
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   " TL   OFFSHORE  SDN  BHD"
            TextSave        =   " TL   OFFSHORE  SDN  BHD"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "20/07/2006"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "12:30 AM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   15435
         TabIndex        =   5
         Top             =   0
         Width           =   15500
         Begin VB.Label lbltitle 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   10920
            TabIndex        =   9
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label lbllocation 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4320
            TabIndex        =   8
            Top             =   0
            Width           =   6375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   8640
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   11055
         End
      End
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   19800
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
            Picture         =   "main.frx":30F42
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":31054
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":314A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":318F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":31D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3219C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":38436
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":38750
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":38A6A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":39004
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3959E
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":39B38
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3A0D2
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3A1E4
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3A726
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3ACC0
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3B25A
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BB34
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BC46
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BD58
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BE6A
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BF7C
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3C08E
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3C1A0
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3C73A
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3CCD4
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3D26E
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3D808
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3D91A
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3DA2C
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3DFC6
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3E0D8
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3E1EA
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3E784
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3E896
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3EE30
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3F3CA
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3F4DC
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3FA76
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40010
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":405AA
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":406BC
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40C56
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40D68
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40E7A
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40F8C
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4109E
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":411B0
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4174A
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4185C
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4196E
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":41F08
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":424A2
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":42A3C
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":42FD6
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":43570
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":43B0A
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":440A4
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   4560
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   78
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":441B6
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":442C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4471A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":44B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":44FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":45410
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4B6AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4B9C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4BCDE
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4C278
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4C812
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4CDAC
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4D346
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4D458
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4D99A
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4DF34
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4E4CE
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4EDA8
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4EEBA
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4EFCC
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F0DE
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F1F0
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F302
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F414
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F9AE
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4FF48
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":504E2
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":50A7C
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":50B8E
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":50CA0
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5123A
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5134C
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5145E
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":519F8
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":51B0A
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":520A4
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5263E
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":52750
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":52CEA
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":53284
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5381E
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":53930
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":53ECA
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":53FDC
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":540EE
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54200
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54312
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54424
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":549BE
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54AD0
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54BE2
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5517C
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":55716
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":55CB0
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5624A
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":567E4
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":56D7E
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":57318
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5742A
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":613EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":72804
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":82069
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":824BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":82911
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":82D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":83214
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":836DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":83C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":8410C
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":846C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":84B0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":8502C
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":98F5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":ACFB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":B09BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":B5173
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":B56B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":BCF80
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim f As String

Private Sub MDIForm_Load()
On Error Resume Next
 
Call connect

fab = 0
 

Dim rst As New ADODB.Recordset
If rst.State Then rst.Close
rst.Open "select * from userrights where u_name='" & frm_login.cbo_userid.Text & "' ", Cn, 3, 2
If Not rst.EOF Then
a = rst!mforms
b = rst!tforms
c = rst!mreports
d = rst!treports
f = rst!others
End If
 
 Call tree
 
 'Call userinvisible
 
 

End Sub

Private Function tree()

TreeView1.Nodes.Add , , "l", "POB", 4

'Company master
 
TreeView1.Nodes.Add "l", tvwChild, "CMPmaster", UCase("MASTERS"), 5, 6
If Mid(a, 1, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "com", ("Company"), 5, 6
End If
If Mid(a, 2, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "hrt", ("HrType"), 5, 6
End If
If Mid(a, 3, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "roo", ("Room"), 5, 6
End If
If Mid(a, 4, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "lif", ("LifeRaft"), 5, 6
End If
If Mid(a, 5, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "nat", ("Nationality"), 5, 6
End If
If Mid(a, 6, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "pub", ("Public Holiday"), 5, 6
End If
If Mid(a, 7, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "bar", ("Location"), 5, 6
End If
If Mid(a, 8, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "jobc", ("JobClassification"), 5, 6
End If
If Mid(a, 9, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "jobn", ("JobNo"), 5, 6
End If
If Mid(a, 10, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "emp", ("Employee"), 5, 6
End If
 
' Transaction

TreeView1.Nodes.Add "l", tvwChild, "Tranx", UCase("TRANSACTIONS"), 5, 6
If Mid(b, 1, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "onb", ("OnBoard"), 5, 6
End If
If Mid(b, 2, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "ofb", ("OffBoard"), 5, 6
End If
If Mid(b, 3, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "tim", ("MonthEnd Entries"), 5, 6
End If
If Mid(b, 4, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "datd", ("Daily Attendance"), 5, 6
End If
If Mid(b, 5, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "uatd", ("Update Attendance"), 5, 6
End If

If Mid(b, 6, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "extots", ("Excel To TimeSheet"), 5, 6
End If
If Mid(b, 7, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "tstoex", ("TimeSheet To Excel"), 5, 6
End If
 

 
 

TreeView1.Nodes.Add "l", tvwChild, "Rtranx", UCase("Reports"), 5, 6
If Mid(d, 1, 1) = 1 Then
TreeView1.Nodes.Add "Rtranx", tvwChild, "rponb", ("Personnel OnBoard"), 5, 6
End If
If Mid(d, 2, 1) = 1 Then
TreeView1.Nodes.Add "Rtranx", tvwChild, "rpofb", ("Personnel OffBoard"), 5, 6
End If
If Mid(d, 3, 1) = 1 Then
TreeView1.Nodes.Add "Rtranx", tvwChild, "indts", ("Individual TimeSheet"), 5, 6
End If
If Mid(d, 4, 1) = 1 Then
TreeView1.Nodes.Add "Rtranx", tvwChild, "sumts", ("TimeSheet Summary"), 5, 6
End If

 
 

 
TreeView1.Nodes.Add "l", tvwChild, "othe", UCase("Others"), 5, 6
' utilities
If Mid(f, 1, 1) = 1 Then
TreeView1.Nodes.Add "othe", tvwChild, "util", UCase("Utilities"), 5, 6
If Mid(f, 2, 1) = 1 Then
TreeView1.Nodes.Add "util", tvwChild, "backup1", ("Backup"), 5, 6
End If
If Mid(f, 3, 1) = 1 Then
TreeView1.Nodes.Add "util", tvwChild, "Restore1", ("Restore"), 5, 6
End If
If Mid(f, 4, 1) = 1 Then
TreeView1.Nodes.Add "util", tvwChild, "msg", ("Send Message"), 5, 6
End If
End If

' Administration
If Mid(f, 5, 1) = 1 Then
TreeView1.Nodes.Add "othe", tvwChild, "admin", UCase("Administration"), 5, 6
If Mid(f, 6, 1) = 1 Then
TreeView1.Nodes.Add "admin", tvwChild, "cmppara", ("Company Parameter"), 5, 6
End If
If Mid(f, 7, 1) = 1 Then
TreeView1.Nodes.Add "admin", tvwChild, "Chngpswd", ("Create Password"), 5, 6
End If
If Mid(f, 8, 1) = 1 Then
TreeView1.Nodes.Add "admin", tvwChild, "usr", ("User Rights"), 5, 6
End If
If Mid(f, 9, 1) = 1 Then
TreeView1.Nodes.Add "admin", tvwChild, "uid", ("UserID"), 5, 6
End If
End If
' Help
If Mid(f, 10, 1) = 1 Then
TreeView1.Nodes.Add "othe", tvwChild, "hlp", UCase("Help"), 5, 6
If Mid(f, 11, 1) = 1 Then
TreeView1.Nodes.Add "hlp", tvwChild, "Dataflow", ("Data Flow"), 5, 6
End If
If Mid(f, 12, 1) = 1 Then
TreeView1.Nodes.Add "hlp", tvwChild, "frmhelp", ("Form Help"), 5, 6
End If
End If

TreeView1.Nodes.Add "l", tvwChild, "logo", UCase("logout"), 5, 6



'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''





End Function

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Dim rss As New ADODB.Recordset
If rss.State Then rss.Close
rss.Open "select * from login where l_userid='" & Label2.Caption & "' and (l_intime)='" & DTP_login.Value & "'", Cn, 5, 6
If Not rss.EOF Then
rss!l_outtime = Now
rss.Update
End If
End Sub




Private Sub mnu_BackUp_Click()
Cn.Execute "BACKUP DATABASE PCMS to disk='c:\timesheet" & Format(Date, "dd-MMM-yyyy") & ".bak'"
 
    If Err.Number = 0 Then MsgBox "Backup Succeded", vbInformation
End Sub


Private Sub mnuutilbackup_Click()
 
 
    Cn.Execute "BACKUP DATABASE PCMS to disk='D:\timesheet" & Format(Date, "dd-MM") & ".bak'"
 
    If Err.Number = 0 Then MsgBox "Backup Succeded", vbInformation
End Sub



Private Sub TreeView1_DblClick()
If TreeView1.SelectedItem.Key = "bar" Then frm_badge.Show
If TreeView1.SelectedItem.Key = "com" Then frm_company.Show
If TreeView1.SelectedItem.Key = "emp" Then frm_employee.Show
If TreeView1.SelectedItem.Key = "extots" Then frm_exceltoflex.Show
If TreeView1.SelectedItem.Key = "tstoex" Then frm_flextoexcel.Show
If TreeView1.SelectedItem.Key = "hrt" Then frm_hrtype.Show
If TreeView1.SelectedItem.Key = "jobc" Then frm_jobclassification.Show

If TreeView1.SelectedItem.Key = "jobn" Then frm_jobno.Show

If TreeView1.SelectedItem.Key = "lif" Then frm_liferaft.Show
If TreeView1.SelectedItem.Key = "nat" Then frm_nationality.Show
If TreeView1.SelectedItem.Key = "ofb" Then frm_offboard.Show
If TreeView1.SelectedItem.Key = "onb" Then frm_pob.Show
If TreeView1.SelectedItem.Key = "pub" Then frm_pholiday.Show
If TreeView1.SelectedItem.Key = "roo" Then frm_room.Show


If TreeView1.SelectedItem.Key = "tim" Then
frm_monthendentries.Show
SetParent frm_monthendentries.HWnd, main.HWnd
End If
If TreeView1.SelectedItem.Key = "uatd" Then frm_timesheet.Show
If TreeView1.SelectedItem.Key = "datd" Then frm_timesheet1.Show


If TreeView1.SelectedItem.Key = "usr" Then frm_userrights.Show
If TreeView1.SelectedItem.Key = "uid" Then frm_userid.Show
'If TreeView1.SelectedItem.Key = "pty" Then frm_pertype.Show

'reports
If TreeView1.SelectedItem.Key = "rponb" Then rpt_personnelonboard.Show
If TreeView1.SelectedItem.Key = "rpofb" Then rpt_offboard.Show
If TreeView1.SelectedItem.Key = "indts" Then rpt_timesheet.Show
If TreeView1.SelectedItem.Key = "sumts" Then rpt_timesheetgroup.Show



If TreeView1.SelectedItem.Key = "User" Then frm_userrights.Show

If TreeView1.SelectedItem.Key = "backup1" Then
Cn.Execute "BACKUP DATABASE PCMS to disk='D:\PCMS" & Format(Date, "dd-MMM-yyyy") & ".bak'"
 
    If Err.Number = 0 Then MsgBox "Backup Succeded", vbInformation
End If

If TreeView1.SelectedItem.Key = "pdiary" Then frm_projectremainder.Show
If TreeView1.SelectedItem.Key = "ujc" Then updatejobcharge.Show

If TreeView1.SelectedItem.Key = "logo" Then
Unload Me
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
 




