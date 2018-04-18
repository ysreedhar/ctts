VERSION 5.00
Begin VB.Form vscrollmonthendentries 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   8595
      Left            =   14400
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000006&
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   34
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   281
         Top             =   8400
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   34
         Left            =   9120
         TabIndex        =   280
         Text            =   "0"
         Top             =   8400
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   34
         Left            =   8040
         TabIndex        =   279
         Text            =   "0"
         Top             =   8400
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   34
         Left            =   7080
         TabIndex        =   278
         Text            =   "0"
         Top             =   8400
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   34
         Left            =   6120
         TabIndex        =   277
         Text            =   "0"
         Top             =   8400
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   34
         Left            =   0
         TabIndex        =   276
         Top             =   8400
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   34
         Left            =   5160
         TabIndex        =   275
         Text            =   "0"
         Top             =   8400
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   34
         Left            =   3480
         TabIndex        =   274
         Top             =   8400
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   33
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   273
         Top             =   8160
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   33
         Left            =   9120
         TabIndex        =   272
         Text            =   "0"
         Top             =   8160
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   33
         Left            =   8040
         TabIndex        =   271
         Text            =   "0"
         Top             =   8160
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   33
         Left            =   7080
         TabIndex        =   270
         Text            =   "0"
         Top             =   8160
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   33
         Left            =   6120
         TabIndex        =   269
         Text            =   "0"
         Top             =   8160
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   33
         Left            =   0
         TabIndex        =   268
         Top             =   8160
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   33
         Left            =   5160
         TabIndex        =   267
         Text            =   "0"
         Top             =   8160
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   33
         Left            =   3480
         TabIndex        =   266
         Top             =   8160
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   32
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   265
         Top             =   7920
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   32
         Left            =   9120
         TabIndex        =   264
         Text            =   "0"
         Top             =   7920
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   32
         Left            =   8040
         TabIndex        =   263
         Text            =   "0"
         Top             =   7920
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   32
         Left            =   7080
         TabIndex        =   262
         Text            =   "0"
         Top             =   7920
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   32
         Left            =   6120
         TabIndex        =   261
         Text            =   "0"
         Top             =   7920
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   32
         Left            =   0
         TabIndex        =   260
         Top             =   7920
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   32
         Left            =   5160
         TabIndex        =   259
         Text            =   "0"
         Top             =   7920
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   32
         Left            =   3480
         TabIndex        =   258
         Top             =   7920
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   31
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   257
         Top             =   7680
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   31
         Left            =   9120
         TabIndex        =   256
         Text            =   "0"
         Top             =   7680
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   31
         Left            =   8040
         TabIndex        =   255
         Text            =   "0"
         Top             =   7680
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   31
         Left            =   7080
         TabIndex        =   254
         Text            =   "0"
         Top             =   7680
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   31
         Left            =   6120
         TabIndex        =   253
         Text            =   "0"
         Top             =   7680
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   31
         Left            =   0
         TabIndex        =   252
         Top             =   7680
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   31
         Left            =   5160
         TabIndex        =   251
         Text            =   "0"
         Top             =   7680
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   31
         Left            =   3480
         TabIndex        =   250
         Top             =   7680
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   30
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   249
         Top             =   7440
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   30
         Left            =   9120
         TabIndex        =   248
         Text            =   "0"
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   30
         Left            =   8040
         TabIndex        =   247
         Text            =   "0"
         Top             =   7440
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   30
         Left            =   7080
         TabIndex        =   246
         Text            =   "0"
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   30
         Left            =   6120
         TabIndex        =   245
         Text            =   "0"
         Top             =   7440
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   30
         Left            =   0
         TabIndex        =   244
         Top             =   7440
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   30
         Left            =   5160
         TabIndex        =   243
         Text            =   "0"
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   30
         Left            =   3480
         TabIndex        =   242
         Top             =   7440
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   29
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   241
         Top             =   7200
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   29
         Left            =   9120
         TabIndex        =   240
         Text            =   "0"
         Top             =   7200
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   29
         Left            =   8040
         TabIndex        =   239
         Text            =   "0"
         Top             =   7200
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   29
         Left            =   7080
         TabIndex        =   238
         Text            =   "0"
         Top             =   7200
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   29
         Left            =   6120
         TabIndex        =   237
         Text            =   "0"
         Top             =   7200
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   29
         Left            =   0
         TabIndex        =   236
         Top             =   7200
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   29
         Left            =   5160
         TabIndex        =   235
         Text            =   "0"
         Top             =   7200
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   29
         Left            =   3480
         TabIndex        =   234
         Top             =   7200
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   28
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   233
         Top             =   6960
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   28
         Left            =   9120
         TabIndex        =   232
         Text            =   "0"
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   28
         Left            =   8040
         TabIndex        =   231
         Text            =   "0"
         Top             =   6960
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   28
         Left            =   7080
         TabIndex        =   230
         Text            =   "0"
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   28
         Left            =   6120
         TabIndex        =   229
         Text            =   "0"
         Top             =   6960
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   28
         Left            =   0
         TabIndex        =   228
         Top             =   6960
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   28
         Left            =   5160
         TabIndex        =   227
         Text            =   "0"
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   28
         Left            =   3480
         TabIndex        =   226
         Top             =   6960
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   27
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   225
         Top             =   6720
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   27
         Left            =   9120
         TabIndex        =   224
         Text            =   "0"
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   27
         Left            =   8040
         TabIndex        =   223
         Text            =   "0"
         Top             =   6720
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   27
         Left            =   7080
         TabIndex        =   222
         Text            =   "0"
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   27
         Left            =   6120
         TabIndex        =   221
         Text            =   "0"
         Top             =   6720
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   27
         Left            =   0
         TabIndex        =   220
         Top             =   6720
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   27
         Left            =   5160
         TabIndex        =   219
         Text            =   "0"
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   27
         Left            =   3480
         TabIndex        =   218
         Top             =   6720
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   26
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   217
         Top             =   6480
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   26
         Left            =   9120
         TabIndex        =   216
         Text            =   "0"
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   26
         Left            =   8040
         TabIndex        =   215
         Text            =   "0"
         Top             =   6480
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   26
         Left            =   7080
         TabIndex        =   214
         Text            =   "0"
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   26
         Left            =   6120
         TabIndex        =   213
         Text            =   "0"
         Top             =   6480
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   26
         Left            =   0
         TabIndex        =   212
         Top             =   6480
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   26
         Left            =   5160
         TabIndex        =   211
         Text            =   "0"
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   26
         Left            =   3480
         TabIndex        =   210
         Top             =   6480
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   25
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   209
         Top             =   6240
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   25
         Left            =   9120
         TabIndex        =   208
         Text            =   "0"
         Top             =   6240
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   25
         Left            =   8040
         TabIndex        =   207
         Text            =   "0"
         Top             =   6240
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   25
         Left            =   7080
         TabIndex        =   206
         Text            =   "0"
         Top             =   6240
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   25
         Left            =   6120
         TabIndex        =   205
         Text            =   "0"
         Top             =   6240
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   25
         Left            =   0
         TabIndex        =   204
         Top             =   6240
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   25
         Left            =   5160
         TabIndex        =   203
         Text            =   "0"
         Top             =   6240
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   25
         Left            =   3480
         TabIndex        =   202
         Top             =   6240
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   24
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   201
         Top             =   6000
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   24
         Left            =   9120
         TabIndex        =   200
         Text            =   "0"
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   24
         Left            =   8040
         TabIndex        =   199
         Text            =   "0"
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   24
         Left            =   7080
         TabIndex        =   198
         Text            =   "0"
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   24
         Left            =   6120
         TabIndex        =   197
         Text            =   "0"
         Top             =   6000
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   24
         Left            =   0
         TabIndex        =   196
         Top             =   6000
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   24
         Left            =   5160
         TabIndex        =   195
         Text            =   "0"
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   24
         Left            =   3480
         TabIndex        =   194
         Top             =   6000
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   23
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   193
         Top             =   5760
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   23
         Left            =   9120
         TabIndex        =   192
         Text            =   "0"
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   23
         Left            =   8040
         TabIndex        =   191
         Text            =   "0"
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   23
         Left            =   7080
         TabIndex        =   190
         Text            =   "0"
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   23
         Left            =   6120
         TabIndex        =   189
         Text            =   "0"
         Top             =   5760
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   23
         Left            =   0
         TabIndex        =   188
         Top             =   5760
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   23
         Left            =   5160
         TabIndex        =   187
         Text            =   "0"
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   23
         Left            =   3480
         TabIndex        =   186
         Top             =   5760
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   22
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   185
         Top             =   5520
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   22
         Left            =   9120
         TabIndex        =   184
         Text            =   "0"
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   22
         Left            =   8040
         TabIndex        =   183
         Text            =   "0"
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   22
         Left            =   7080
         TabIndex        =   182
         Text            =   "0"
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   22
         Left            =   6120
         TabIndex        =   181
         Text            =   "0"
         Top             =   5520
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   22
         Left            =   0
         TabIndex        =   180
         Top             =   5520
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   22
         Left            =   5160
         TabIndex        =   179
         Text            =   "0"
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   22
         Left            =   3480
         TabIndex        =   178
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   21
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   177
         Top             =   5280
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   21
         Left            =   9120
         TabIndex        =   176
         Text            =   "0"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   21
         Left            =   8040
         TabIndex        =   175
         Text            =   "0"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   21
         Left            =   7080
         TabIndex        =   174
         Text            =   "0"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   21
         Left            =   6120
         TabIndex        =   173
         Text            =   "0"
         Top             =   5280
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   21
         Left            =   0
         TabIndex        =   172
         Top             =   5280
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   21
         Left            =   5160
         TabIndex        =   171
         Text            =   "0"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   21
         Left            =   3480
         TabIndex        =   170
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   20
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   169
         Top             =   5040
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   20
         Left            =   9120
         TabIndex        =   168
         Text            =   "0"
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   20
         Left            =   8040
         TabIndex        =   167
         Text            =   "0"
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   20
         Left            =   7080
         TabIndex        =   166
         Text            =   "0"
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   20
         Left            =   6120
         TabIndex        =   165
         Text            =   "0"
         Top             =   5040
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   20
         Left            =   0
         TabIndex        =   164
         Top             =   5040
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   20
         Left            =   5160
         TabIndex        =   163
         Text            =   "0"
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   20
         Left            =   3480
         TabIndex        =   162
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   19
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   161
         Top             =   4800
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   19
         Left            =   9120
         TabIndex        =   160
         Text            =   "0"
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   19
         Left            =   8040
         TabIndex        =   159
         Text            =   "0"
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   19
         Left            =   7080
         TabIndex        =   158
         Text            =   "0"
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   19
         Left            =   6120
         TabIndex        =   157
         Text            =   "0"
         Top             =   4800
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   19
         Left            =   0
         TabIndex        =   156
         Top             =   4800
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   19
         Left            =   5160
         TabIndex        =   155
         Text            =   "0"
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   19
         Left            =   3480
         TabIndex        =   154
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   18
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   153
         Top             =   4560
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   18
         Left            =   9120
         TabIndex        =   152
         Text            =   "0"
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   18
         Left            =   8040
         TabIndex        =   151
         Text            =   "0"
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   18
         Left            =   7080
         TabIndex        =   150
         Text            =   "0"
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   18
         Left            =   6120
         TabIndex        =   149
         Text            =   "0"
         Top             =   4560
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   18
         Left            =   0
         TabIndex        =   148
         Top             =   4560
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   18
         Left            =   5160
         TabIndex        =   147
         Text            =   "0"
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   18
         Left            =   3480
         TabIndex        =   146
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   17
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   145
         Top             =   4320
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   17
         Left            =   9120
         TabIndex        =   144
         Text            =   "0"
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   17
         Left            =   8040
         TabIndex        =   143
         Text            =   "0"
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   17
         Left            =   7080
         TabIndex        =   142
         Text            =   "0"
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   17
         Left            =   6120
         TabIndex        =   141
         Text            =   "0"
         Top             =   4320
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   17
         Left            =   0
         TabIndex        =   140
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   17
         Left            =   5160
         TabIndex        =   139
         Text            =   "0"
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   17
         Left            =   3480
         TabIndex        =   138
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   16
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   137
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   16
         Left            =   9120
         TabIndex        =   136
         Text            =   "0"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   16
         Left            =   8040
         TabIndex        =   135
         Text            =   "0"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   16
         Left            =   7080
         TabIndex        =   134
         Text            =   "0"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   16
         Left            =   6120
         TabIndex        =   133
         Text            =   "0"
         Top             =   4080
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   16
         Left            =   0
         TabIndex        =   132
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   16
         Left            =   5160
         TabIndex        =   131
         Text            =   "0"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   16
         Left            =   3480
         TabIndex        =   130
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   15
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   129
         Top             =   3840
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   15
         Left            =   9120
         TabIndex        =   128
         Text            =   "0"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   15
         Left            =   8040
         TabIndex        =   127
         Text            =   "0"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   15
         Left            =   7080
         TabIndex        =   126
         Text            =   "0"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   15
         Left            =   6120
         TabIndex        =   125
         Text            =   "0"
         Top             =   3840
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   15
         Left            =   0
         TabIndex        =   124
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   15
         Left            =   5160
         TabIndex        =   123
         Text            =   "0"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   15
         Left            =   3480
         TabIndex        =   122
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   14
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   121
         Top             =   3600
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   14
         Left            =   9120
         TabIndex        =   120
         Text            =   "0"
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   14
         Left            =   8040
         TabIndex        =   119
         Text            =   "0"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   14
         Left            =   7080
         TabIndex        =   118
         Text            =   "0"
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   14
         Left            =   6120
         TabIndex        =   117
         Text            =   "0"
         Top             =   3600
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   14
         Left            =   0
         TabIndex        =   116
         Top             =   3600
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   14
         Left            =   5160
         TabIndex        =   115
         Text            =   "0"
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   14
         Left            =   3480
         TabIndex        =   114
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   13
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   113
         Top             =   3360
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   13
         Left            =   9120
         TabIndex        =   112
         Text            =   "0"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   13
         Left            =   8040
         TabIndex        =   111
         Text            =   "0"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   13
         Left            =   7080
         TabIndex        =   110
         Text            =   "0"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   13
         Left            =   6120
         TabIndex        =   109
         Text            =   "0"
         Top             =   3360
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   13
         Left            =   0
         TabIndex        =   108
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   13
         Left            =   5160
         TabIndex        =   107
         Text            =   "0"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   13
         Left            =   3480
         TabIndex        =   106
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   12
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   12
         Left            =   9120
         TabIndex        =   104
         Text            =   "0"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   12
         Left            =   8040
         TabIndex        =   103
         Text            =   "0"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   12
         Left            =   7080
         TabIndex        =   102
         Text            =   "0"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   12
         Left            =   6120
         TabIndex        =   101
         Text            =   "0"
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   12
         Left            =   0
         TabIndex        =   100
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   12
         Left            =   5160
         TabIndex        =   99
         Text            =   "0"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   12
         Left            =   3480
         TabIndex        =   98
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   11
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   97
         Top             =   2880
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   11
         Left            =   9120
         TabIndex        =   96
         Text            =   "0"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   11
         Left            =   8040
         TabIndex        =   95
         Text            =   "0"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   11
         Left            =   7080
         TabIndex        =   94
         Text            =   "0"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   11
         Left            =   6120
         TabIndex        =   93
         Text            =   "0"
         Top             =   2880
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   11
         Left            =   0
         TabIndex        =   92
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   11
         Left            =   5160
         TabIndex        =   91
         Text            =   "0"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   11
         Left            =   3480
         TabIndex        =   90
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   10
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   89
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   10
         Left            =   9120
         TabIndex        =   88
         Text            =   "0"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   10
         Left            =   8040
         TabIndex        =   87
         Text            =   "0"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   10
         Left            =   7080
         TabIndex        =   86
         Text            =   "0"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   10
         Left            =   6120
         TabIndex        =   85
         Text            =   "0"
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   10
         Left            =   0
         TabIndex        =   84
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   10
         Left            =   5160
         TabIndex        =   83
         Text            =   "0"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   10
         Left            =   3480
         TabIndex        =   82
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   9
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   81
         Top             =   2400
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   9
         Left            =   9120
         TabIndex        =   80
         Text            =   "0"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   9
         Left            =   8040
         TabIndex        =   79
         Text            =   "0"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   9
         Left            =   7080
         TabIndex        =   78
         Text            =   "0"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   9
         Left            =   6120
         TabIndex        =   77
         Text            =   "0"
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   9
         Left            =   0
         TabIndex        =   76
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   9
         Left            =   5160
         TabIndex        =   75
         Text            =   "0"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   9
         Left            =   3480
         TabIndex        =   74
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   8
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   73
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   8
         Left            =   9120
         TabIndex        =   72
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   8
         Left            =   8040
         TabIndex        =   71
         Text            =   "0"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   8
         Left            =   7080
         TabIndex        =   70
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   8
         Left            =   6120
         TabIndex        =   69
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   8
         Left            =   0
         TabIndex        =   68
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   8
         Left            =   5160
         TabIndex        =   67
         Text            =   "0"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   8
         Left            =   3480
         TabIndex        =   66
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   7
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   65
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   7
         Left            =   9120
         TabIndex        =   64
         Text            =   "0"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   7
         Left            =   8040
         TabIndex        =   63
         Text            =   "0"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   7
         Left            =   7080
         TabIndex        =   62
         Text            =   "0"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   7
         Left            =   6120
         TabIndex        =   61
         Text            =   "0"
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   7
         Left            =   0
         TabIndex        =   60
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   7
         Left            =   5160
         TabIndex        =   59
         Text            =   "0"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   7
         Left            =   3480
         TabIndex        =   58
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   6
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   6
         Left            =   9120
         TabIndex        =   56
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   6
         Left            =   8040
         TabIndex        =   55
         Text            =   "0"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   6
         Left            =   7080
         TabIndex        =   54
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   6
         Left            =   6120
         TabIndex        =   53
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   6
         Left            =   0
         TabIndex        =   52
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   6
         Left            =   5160
         TabIndex        =   51
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   6
         Left            =   3480
         TabIndex        =   50
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   5
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   5
         Left            =   9120
         TabIndex        =   48
         Text            =   "0"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   5
         Left            =   8040
         TabIndex        =   47
         Text            =   "0"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   5
         Left            =   7080
         TabIndex        =   46
         Text            =   "0"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   5
         Left            =   6120
         TabIndex        =   45
         Text            =   "0"
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   5
         Left            =   0
         TabIndex        =   44
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   5
         Left            =   5160
         TabIndex        =   43
         Text            =   "0"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   5
         Left            =   3480
         TabIndex        =   42
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   4
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   4
         Left            =   9120
         TabIndex        =   40
         Text            =   "0"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   4
         Left            =   8040
         TabIndex        =   39
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   4
         Left            =   7080
         TabIndex        =   38
         Text            =   "0"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   4
         Left            =   6120
         TabIndex        =   37
         Text            =   "0"
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   4
         Left            =   0
         TabIndex        =   36
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   4
         Left            =   5160
         TabIndex        =   35
         Text            =   "0"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   34
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   3
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   3
         Left            =   9120
         TabIndex        =   32
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   3
         Left            =   8040
         TabIndex        =   31
         Text            =   "0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   3
         Left            =   7080
         TabIndex        =   30
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   3
         Left            =   6120
         TabIndex        =   29
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   3
         Left            =   0
         TabIndex        =   28
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   3
         Left            =   5160
         TabIndex        =   27
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   3
         Left            =   3480
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   2
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   2
         Left            =   9120
         TabIndex        =   24
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   2
         Left            =   8040
         TabIndex        =   23
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   2
         Left            =   7080
         TabIndex        =   22
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   2
         Left            =   6120
         TabIndex        =   21
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   20
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   19
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   1
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   420
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   1
         Left            =   9120
         TabIndex        =   16
         Text            =   "0"
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   1
         Left            =   8040
         TabIndex        =   15
         Text            =   "0"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   1
         Left            =   7080
         TabIndex        =   14
         Text            =   "0"
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   1
         Left            =   6120
         TabIndex        =   13
         Text            =   "0"
         Top             =   420
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   420
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   1
         Left            =   5160
         TabIndex        =   11
         Text            =   "0"
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   10
         Top             =   420
         Width           =   1695
      End
      Begin VB.TextBox txt_notes 
         Height          =   285
         Index           =   0
         Left            =   10080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   120
         Width           =   4335
      End
      Begin VB.TextBox txt_bondstore 
         Height          =   285
         Index           =   0
         Left            =   9120
         TabIndex        =   8
         Text            =   "0"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txt_cashadvance 
         Height          =   285
         Index           =   0
         Left            =   8040
         TabIndex        =   7
         Text            =   "0"
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txt_phonecall 
         Height          =   285
         Index           =   0
         Left            =   7080
         TabIndex        =   6
         Text            =   "0"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txt_actoff 
         Height          =   285
         Index           =   0
         Left            =   6120
         TabIndex        =   5
         Text            =   "0"
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cbo_name 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   3495
      End
      Begin VB.TextBox txt_travel 
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   3
         Text            =   "0"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txt_classification 
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "vscrollmonthendentries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyClassificationArray As Object
Private MyNameArray As Object
Private MyTravelArray As Object
Private MyTOffArray As Object
Private MyPhoneCallArray As Object
Private MyCashAdvanceArray As Object
Private MyBondStoreArray As Object
Private MyRemarksArray As Object
Private Sub Form_Load()
''''scroll

Dim VPos As Integer
 
  'Change the following numbers to the Full height and width of your Form
  intFullHeight = 12000 'Maximized the Form and note the Figures
   
  'This is the how much of your Form is displayed
  intDisplayHeight = Me.Height
   

  With VScroll1
    '.Height = Me.ScaleHeight
    .Min = 0
    .Max = intFullHeight - intDisplayHeight
    .SmallChange = Screen.TwipsPerPixelX * 10
    .LargeChange = .SmallChange
  End With
    
'scroll
End Sub
Sub ScrollForm(Direction As Byte, NewVal As Integer)
  
  Dim CTL As Control
  Static hOldVal As Integer
  Static vOldVal As Integer
  Dim hMoveDiff As Integer 'Diff in the horizontal controls movements
  Dim vMoveDiff As Integer 'Diff in the vertical controls Movements
  
  Select Case Direction
    
  Case 0 'Scroll Vertically
  
    'Check The Direction of the Vertical Scroll & Extract Value Diff
    If NewVal > vOldVal Then 'Scrolled From Top to Bottom
      'Controls MUST move to the TOP, therefore TOP value Decreases
      vMoveDiff = -(NewVal - vOldVal)
      
            '''''''''''''''''
'        pView.Height = pView.Height - 400
        Frame1.Height = Frame1.Height + 400
'        vscrollform.Height = vscrollform.Height + 400
''''''''''''''''
    Else 'Scrolled From Bottom to Top
      'Controls MUST move to the Bottom, therefore TOP value Increases
      vMoveDiff = (vOldVal - NewVal)
      
      '''''''''''''''''
'        pView.Height = pView.Height - 400
        Frame1.Height = Frame1.Height - 400
'        Me.Height = Me.Height - 400
''''''''''''''''
      
      
      
    End If
  
    For Each CTL In Me.Controls
      'Make sure it's not a ScrollBar
      If Not (TypeOf CTL Is VScrollBar) Then
        'If it's a Line then
        If TypeOf CTL Is Line Then
          CTL.Y1 = CTL.Y1 + vMoveDiff '+ VPos - VScroll1.Value
          CTL.Y2 = CTL.Y2 + vMoveDiff '+ VPos - VScroll1.Value
        Else
          CTL.Top = CTL.Top + vMoveDiff '+ VPos - VScroll1.Value
        End If
      End If
    Next
    
      vOldVal = NewVal 'Reset vOldVal to reflect New Pos of ScrollBar
    
     
  End Select

End Sub

Private Sub VScroll1_Change()
  
  ScrollForm 0, VScroll1.Value
'''
With addname
.Top = cbo_name(MyClassificationArray.ubound - 1).Top + txt_training(MyClassificationArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With
      
With addclassification
.Top = txt_classification(MyClassificationArray.ubound - 1).Top + txt_classification(MyClassificationArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addtravel
.Top = txt_travel(MyTravelArray.ubound - 1).Top + txt_travel(MyTravelArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addtoff
.Top = txt_actoff(MyTOffArray.ubound - 1).Top + txt_actoff(MyTOffArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addphonecall
.Top = txt_phonecall(MyPhoneCallArray.ubound - 1).Top + txt_phonecall(MyPhoneCallArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addcashadvance
.Top = txt_cashadvance(MyCashAdvanceArray.ubound - 1).Top + txt_cashadvance(MyCashAdvanceArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addbondstore
.Top = txt_bondstore(MyBondStoreArray.ubound - 1).Top + txt_bondstore(MyBondStoreArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With
With addremarks
.Top = txt_notes(MyRemarksArray.ubound - 1).Top + txt_notes(MyRemarksArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With


'pView.Height = pView.Height + 400
Frame1.Height = Frame1.Height + 400
'Me.Height = Frame1.Height + 400
''''''

End Sub

Private Sub VScroll1_Scroll()
  
 ScrollForm 0, VScroll1.Value
'''
With addname
.Top = cbo_name(MyClassificationArray.ubound - 1).Top + txt_training(MyClassificationArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With
      
With addclassification
.Top = txt_classification(MyClassificationArray.ubound - 1).Top + txt_classification(MyClassificationArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addtravel
.Top = txt_travel(MyTravelArray.ubound - 1).Top + txt_travel(MyTravelArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addtoff
.Top = txt_actoff(MyTOffArray.ubound - 1).Top + txt_actoff(MyTOffArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addphonecall
.Top = txt_phonecall(MyPhoneCallArray.ubound - 1).Top + txt_phonecall(MyPhoneCallArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addcashadvance
.Top = txt_cashadvance(MyCashAdvanceArray.ubound - 1).Top + txt_cashadvance(MyCashAdvanceArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With

With addbondstore
.Top = txt_bondstore(MyBondStoreArray.ubound - 1).Top + txt_bondstore(MyBondStoreArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With
With addremarks
.Top = txt_notes(MyRemarksArray.ubound - 1).Top + txt_notes(MyRemarksArray.ubound - 1).Height + 100
         .Visible = True
         .Text = ""
.SetFocus
End With


'pView.Height = pView.Height + 400
Frame1.Height = Frame1.Height + 400
'Me.Height = Frame1.Height + 400
''''''

End Sub

Private Sub Form_Initialize()
     Set MyNameArray = Me.Controls("cbo_name")
     Set MyClassificationArray = Me.Controls("txt_classification")
     Set MyTravelArray = Me.Controls("txt_travel")
     Set MyTOffArray = Me.Controls("txt_actoff")
     Set MyPhoneCallArray = Me.Controls("txt_phonecall")
     Set MyCashAdvanceArray = Me.Controls("txt_cashadvance")
     Set MyBondStoreArray = Me.Controls("txt_bondstore")
     Set MyRemarksArray = Me.Controls("txt_notes")
     End Sub
Public Function addname() As ComboBox
   Dim m As Integer
   m = MyNameArray.ubound + 1
   Load MyNameArray(m)
   Set addname = MyNameArray(m)
End Function
Public Function addclassification() As TextBox
   Dim m As Integer
   m = MyClassificationArray.ubound + 1
   Load MyClassificationArray(m)
   Set addclassification = MyClassificationArray(m)
End Function

Public Function addtravel() As TextBox
   Dim m As Integer
   m = MyTravelArray.ubound + 1
   Load MyTravelArray(m)
   Set addtravel = MyTravelArray(m)
End Function
Public Function addtoff() As TextBox
   Dim m As Integer
   m = MyTOffArray.ubound + 1
   Load MyTOffArray(m)
   Set addtoff = MyTOffArray(m)
End Function
Public Function addphonecall() As TextBox
   Dim m As Integer
   m = MyPhoneCallArray.ubound + 1
   Load MyPhoneCallArray(m)
   Set addphonecall = MyPhoneCallArray(m)
End Function
Public Function addcashadvance() As TextBox
   Dim m As Integer
   m = MyCashAdvanceArray.ubound + 1
   Load MyCashAdvanceArray(m)
   Set addcashadvance = MyCashAdvanceArray(m)
End Function
Public Function addbondstore() As TextBox
   Dim m As Integer
   m = MyBondStoreArray.ubound + 1
   Load MyBondStoreArray(m)
   Set addbondstore = MyBondStoreArray(m)
End Function

