VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmVerCalendario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Calendario"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   Icon            =   "FormVCalend.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   461
      Text            =   "Text1"
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   375
      Left            =   9000
      TabIndex        =   460
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3855
      Left            =   6840
      TabIndex        =   109
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5883
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   6840
      TabIndex        =   457
      Top             =   5160
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   240
      TabIndex        =   462
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "HORARIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   459
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Festivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   458
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   240
      X2              =   2040
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   2520
      X2              =   4320
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   4800
      X2              =   6600
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   240
      X2              =   2040
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   2520
      X2              =   4320
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   4800
      X2              =   6600
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   240
      X2              =   2040
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   2520
      X2              =   4320
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   4800
      X2              =   6600
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   4800
      X2              =   6600
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   447
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2520
      X2              =   4320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   446
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   2040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   36
      Left            =   5160
      OLEDropMode     =   1  'Manual
      TabIndex        =   445
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   35
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   444
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   443
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   600
      TabIndex        =   442
      Top             =   4200
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   360
      TabIndex        =   441
      Top             =   4200
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   1800
      TabIndex        =   440
      Top             =   3960
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   1560
      TabIndex        =   439
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   1320
      TabIndex        =   438
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   1080
      TabIndex        =   437
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   840
      TabIndex        =   436
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   600
      TabIndex        =   435
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   360
      TabIndex        =   434
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   1800
      TabIndex        =   433
      Top             =   3720
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   1560
      TabIndex        =   432
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   1320
      TabIndex        =   431
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   1080
      TabIndex        =   430
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   840
      TabIndex        =   429
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   600
      TabIndex        =   428
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   360
      TabIndex        =   427
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   1800
      TabIndex        =   426
      Top             =   3480
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   1560
      TabIndex        =   425
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   1320
      TabIndex        =   424
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   1080
      TabIndex        =   423
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   840
      TabIndex        =   422
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   600
      TabIndex        =   421
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   360
      TabIndex        =   420
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   1800
      TabIndex        =   419
      Top             =   3240
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   1560
      TabIndex        =   418
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   1320
      TabIndex        =   417
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   1080
      TabIndex        =   416
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   840
      TabIndex        =   415
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   600
      TabIndex        =   414
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   360
      TabIndex        =   413
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   1800
      TabIndex        =   412
      Top             =   3000
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   1560
      TabIndex        =   411
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   1320
      TabIndex        =   410
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   1080
      TabIndex        =   409
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   840
      TabIndex        =   408
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   407
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   406
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   600
      TabIndex        =   405
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   360
      TabIndex        =   404
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   1800
      TabIndex        =   403
      Top             =   8040
      Width           =   210
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   1560
      TabIndex        =   402
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   1320
      TabIndex        =   401
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   1080
      TabIndex        =   400
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   840
      TabIndex        =   399
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   600
      TabIndex        =   398
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   360
      TabIndex        =   397
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   1800
      TabIndex        =   396
      Top             =   7800
      Width           =   210
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   1560
      TabIndex        =   395
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   1320
      TabIndex        =   394
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   1080
      TabIndex        =   393
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   840
      TabIndex        =   392
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   600
      TabIndex        =   391
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   360
      TabIndex        =   390
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   1800
      TabIndex        =   389
      Top             =   7560
      Width           =   210
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   1560
      TabIndex        =   388
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   1320
      TabIndex        =   387
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   1080
      TabIndex        =   386
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   840
      TabIndex        =   385
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   600
      TabIndex        =   384
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   360
      TabIndex        =   383
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   1800
      TabIndex        =   382
      Top             =   7320
      Width           =   210
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   1560
      TabIndex        =   381
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   1320
      TabIndex        =   380
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   1080
      TabIndex        =   379
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   840
      TabIndex        =   378
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   600
      TabIndex        =   377
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   360
      TabIndex        =   376
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   1800
      TabIndex        =   375
      Top             =   7080
      Width           =   210
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   1560
      TabIndex        =   374
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   1320
      TabIndex        =   373
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   1080
      TabIndex        =   372
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   840
      TabIndex        =   371
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   370
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   369
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   5160
      TabIndex        =   368
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   4920
      TabIndex        =   367
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   6360
      TabIndex        =   366
      Top             =   8040
      Width           =   210
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   6120
      TabIndex        =   365
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   5880
      TabIndex        =   364
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   5640
      TabIndex        =   363
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   5400
      TabIndex        =   362
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   5160
      TabIndex        =   361
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   4920
      TabIndex        =   360
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   6360
      TabIndex        =   359
      Top             =   7800
      Width           =   210
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   6120
      TabIndex        =   358
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   5880
      TabIndex        =   357
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   5640
      TabIndex        =   356
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   5400
      TabIndex        =   355
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   5160
      TabIndex        =   354
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   4920
      TabIndex        =   353
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   6360
      TabIndex        =   352
      Top             =   7560
      Width           =   210
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   6120
      TabIndex        =   351
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   5880
      TabIndex        =   350
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   5640
      TabIndex        =   349
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   5400
      TabIndex        =   348
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   5160
      TabIndex        =   347
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   4920
      TabIndex        =   346
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   6360
      TabIndex        =   345
      Top             =   7320
      Width           =   210
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   6120
      TabIndex        =   344
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   5880
      TabIndex        =   343
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   5640
      TabIndex        =   342
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   5400
      TabIndex        =   341
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   5160
      TabIndex        =   340
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   4920
      TabIndex        =   339
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   6360
      TabIndex        =   338
      Top             =   7080
      Width           =   210
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   6120
      TabIndex        =   337
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   5880
      TabIndex        =   336
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   5640
      TabIndex        =   335
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   5400
      TabIndex        =   334
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   5160
      TabIndex        =   333
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   4920
      TabIndex        =   332
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   2880
      TabIndex        =   331
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   2640
      TabIndex        =   330
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   4080
      TabIndex        =   329
      Top             =   8040
      Width           =   210
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   3840
      TabIndex        =   328
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   3600
      TabIndex        =   327
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   3360
      TabIndex        =   326
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   3120
      TabIndex        =   325
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   2880
      TabIndex        =   324
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   2640
      TabIndex        =   323
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   4080
      TabIndex        =   322
      Top             =   7800
      Width           =   210
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   3840
      TabIndex        =   321
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   3600
      TabIndex        =   320
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   3360
      TabIndex        =   319
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   3120
      TabIndex        =   318
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   2880
      TabIndex        =   317
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   2640
      TabIndex        =   316
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   4080
      TabIndex        =   315
      Top             =   7560
      Width           =   210
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   3840
      TabIndex        =   314
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   3600
      TabIndex        =   313
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   3360
      TabIndex        =   312
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   3120
      TabIndex        =   311
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   2880
      TabIndex        =   310
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   2640
      TabIndex        =   309
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   4080
      TabIndex        =   308
      Top             =   7320
      Width           =   210
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   3840
      TabIndex        =   307
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   3600
      TabIndex        =   306
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   3360
      TabIndex        =   305
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   3120
      TabIndex        =   304
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   2880
      TabIndex        =   303
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   2640
      TabIndex        =   302
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   4080
      TabIndex        =   301
      Top             =   7080
      Width           =   210
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   3840
      TabIndex        =   300
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   3600
      TabIndex        =   299
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   3360
      TabIndex        =   298
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   3120
      TabIndex        =   297
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   2880
      TabIndex        =   296
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   2640
      TabIndex        =   295
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   5160
      TabIndex        =   294
      Top             =   6240
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   4920
      TabIndex        =   293
      Top             =   6240
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   6360
      TabIndex        =   292
      Top             =   6000
      Width           =   210
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   6120
      TabIndex        =   291
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   5880
      TabIndex        =   290
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   5640
      TabIndex        =   289
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   5400
      TabIndex        =   288
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   5160
      TabIndex        =   287
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   4920
      TabIndex        =   286
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   6360
      TabIndex        =   285
      Top             =   5760
      Width           =   210
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   6120
      TabIndex        =   284
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   5880
      TabIndex        =   283
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   5640
      TabIndex        =   282
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   5400
      TabIndex        =   281
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   5160
      TabIndex        =   280
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   4920
      TabIndex        =   279
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   6360
      TabIndex        =   278
      Top             =   5520
      Width           =   210
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   6120
      TabIndex        =   277
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   5880
      TabIndex        =   276
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   5640
      TabIndex        =   275
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   5400
      TabIndex        =   274
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   5160
      TabIndex        =   273
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   4920
      TabIndex        =   272
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   6360
      TabIndex        =   271
      Top             =   5280
      Width           =   210
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   6120
      TabIndex        =   270
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   5880
      TabIndex        =   269
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   5640
      TabIndex        =   268
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   5400
      TabIndex        =   267
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   5160
      TabIndex        =   266
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   4920
      TabIndex        =   265
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   6360
      TabIndex        =   264
      Top             =   5040
      Width           =   210
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   6120
      TabIndex        =   263
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   5880
      TabIndex        =   262
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   5640
      TabIndex        =   261
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   5400
      TabIndex        =   260
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   5160
      TabIndex        =   259
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   4920
      TabIndex        =   258
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   2880
      TabIndex        =   257
      Top             =   6240
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   2640
      TabIndex        =   256
      Top             =   6240
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   4080
      TabIndex        =   255
      Top             =   6000
      Width           =   210
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   3840
      TabIndex        =   254
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   3600
      TabIndex        =   253
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   3360
      TabIndex        =   252
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   3120
      TabIndex        =   251
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   2880
      TabIndex        =   250
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   2640
      TabIndex        =   249
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   4080
      TabIndex        =   248
      Top             =   5760
      Width           =   210
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   3840
      TabIndex        =   247
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   3600
      TabIndex        =   246
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   3360
      TabIndex        =   245
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   3120
      TabIndex        =   244
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   2880
      TabIndex        =   243
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   2640
      TabIndex        =   242
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   4080
      TabIndex        =   241
      Top             =   5520
      Width           =   210
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   3840
      TabIndex        =   240
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   3600
      TabIndex        =   239
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   3360
      TabIndex        =   238
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   3120
      TabIndex        =   237
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   2880
      TabIndex        =   236
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   2640
      TabIndex        =   235
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   4080
      TabIndex        =   234
      Top             =   5280
      Width           =   210
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   3840
      TabIndex        =   233
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   3600
      TabIndex        =   232
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   3360
      TabIndex        =   231
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   3120
      TabIndex        =   230
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   2880
      TabIndex        =   229
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   2640
      TabIndex        =   228
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   4080
      TabIndex        =   227
      Top             =   5040
      Width           =   210
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   3840
      TabIndex        =   226
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   3600
      TabIndex        =   225
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   3360
      TabIndex        =   224
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   3120
      TabIndex        =   223
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   2880
      TabIndex        =   222
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   2640
      TabIndex        =   221
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   600
      TabIndex        =   220
      Top             =   6240
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   360
      TabIndex        =   219
      Top             =   6240
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   1800
      TabIndex        =   218
      Top             =   6000
      Width           =   210
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   1560
      TabIndex        =   217
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   1320
      TabIndex        =   216
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   1080
      TabIndex        =   215
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   840
      TabIndex        =   214
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   600
      TabIndex        =   213
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   360
      TabIndex        =   212
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   1800
      TabIndex        =   211
      Top             =   5760
      Width           =   210
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   1560
      TabIndex        =   210
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   1320
      TabIndex        =   209
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   1080
      TabIndex        =   208
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   840
      TabIndex        =   207
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   600
      TabIndex        =   206
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   360
      TabIndex        =   205
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   1800
      TabIndex        =   204
      Top             =   5520
      Width           =   210
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   1560
      TabIndex        =   203
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   1320
      TabIndex        =   202
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   1080
      TabIndex        =   201
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   840
      TabIndex        =   200
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   600
      TabIndex        =   199
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   360
      TabIndex        =   198
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   1800
      TabIndex        =   197
      Top             =   5280
      Width           =   210
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   1560
      TabIndex        =   196
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   1320
      TabIndex        =   195
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   1080
      TabIndex        =   194
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   840
      TabIndex        =   193
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   600
      TabIndex        =   192
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   360
      TabIndex        =   191
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   1800
      TabIndex        =   190
      Top             =   5040
      Width           =   210
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   1560
      TabIndex        =   189
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   1320
      TabIndex        =   188
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   1080
      TabIndex        =   187
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   840
      TabIndex        =   186
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   185
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   184
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   5160
      TabIndex        =   183
      Top             =   4200
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   4920
      TabIndex        =   182
      Top             =   4200
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   6360
      TabIndex        =   181
      Top             =   3960
      Width           =   210
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   6120
      TabIndex        =   180
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   5880
      TabIndex        =   179
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   5640
      TabIndex        =   178
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   5400
      TabIndex        =   177
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   5160
      TabIndex        =   176
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   4920
      TabIndex        =   175
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   6360
      TabIndex        =   174
      Top             =   3720
      Width           =   210
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   6120
      TabIndex        =   173
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   5880
      TabIndex        =   172
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   5640
      TabIndex        =   171
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   5400
      TabIndex        =   170
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   5160
      TabIndex        =   169
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   4920
      TabIndex        =   168
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   6360
      TabIndex        =   167
      Top             =   3480
      Width           =   210
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   6120
      TabIndex        =   166
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   5880
      TabIndex        =   165
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   5640
      TabIndex        =   164
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   5400
      TabIndex        =   163
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   5160
      TabIndex        =   162
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   4920
      TabIndex        =   161
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   6360
      TabIndex        =   160
      Top             =   3240
      Width           =   210
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   6120
      TabIndex        =   159
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   5880
      TabIndex        =   158
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   5640
      TabIndex        =   157
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   5400
      TabIndex        =   156
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   5160
      TabIndex        =   155
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   4920
      TabIndex        =   154
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   6360
      TabIndex        =   153
      Top             =   3000
      Width           =   210
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   6120
      TabIndex        =   152
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   5880
      TabIndex        =   151
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   5640
      TabIndex        =   150
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   5400
      TabIndex        =   149
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   5160
      TabIndex        =   148
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   4920
      TabIndex        =   147
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   2880
      TabIndex        =   146
      Top             =   4200
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   2640
      TabIndex        =   145
      Top             =   4200
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   4080
      TabIndex        =   144
      Top             =   3960
      Width           =   210
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   3840
      TabIndex        =   143
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   3600
      TabIndex        =   142
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   3360
      TabIndex        =   141
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   3120
      TabIndex        =   140
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   2880
      TabIndex        =   139
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   2640
      TabIndex        =   138
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   4080
      TabIndex        =   137
      Top             =   3720
      Width           =   210
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   3840
      TabIndex        =   136
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   3600
      TabIndex        =   135
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   3360
      TabIndex        =   134
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   3120
      TabIndex        =   133
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   2880
      TabIndex        =   132
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   2640
      TabIndex        =   131
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   4080
      TabIndex        =   130
      Top             =   3480
      Width           =   210
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   3840
      TabIndex        =   129
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   3600
      TabIndex        =   128
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   3360
      TabIndex        =   127
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   3120
      TabIndex        =   126
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   2880
      TabIndex        =   125
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   2640
      TabIndex        =   124
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   4080
      TabIndex        =   123
      Top             =   3240
      Width           =   210
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   3840
      TabIndex        =   122
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   3600
      TabIndex        =   121
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   3360
      TabIndex        =   120
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   3120
      TabIndex        =   119
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   2880
      TabIndex        =   118
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   2640
      TabIndex        =   117
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   4080
      TabIndex        =   116
      Top             =   3000
      Width           =   210
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   3840
      TabIndex        =   115
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   3600
      TabIndex        =   114
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   3360
      TabIndex        =   113
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   3120
      TabIndex        =   112
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   2880
      TabIndex        =   111
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   2640
      TabIndex        =   110
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   0
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   108
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   1
      Left            =   5160
      OLEDropMode     =   1  'Manual
      TabIndex        =   107
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   2
      Left            =   5400
      OLEDropMode     =   1  'Manual
      TabIndex        =   106
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   3
      Left            =   5640
      OLEDropMode     =   1  'Manual
      TabIndex        =   105
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   4
      Left            =   5880
      OLEDropMode     =   1  'Manual
      TabIndex        =   104
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   5
      Left            =   6120
      OLEDropMode     =   1  'Manual
      TabIndex        =   103
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   6360
      OLEDropMode     =   1  'Manual
      TabIndex        =   102
      Top             =   1080
      Width           =   210
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   7
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   101
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   8
      Left            =   5160
      OLEDropMode     =   1  'Manual
      TabIndex        =   100
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   9
      Left            =   5400
      OLEDropMode     =   1  'Manual
      TabIndex        =   99
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   10
      Left            =   5640
      OLEDropMode     =   1  'Manual
      TabIndex        =   98
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   11
      Left            =   5880
      OLEDropMode     =   1  'Manual
      TabIndex        =   97
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   12
      Left            =   6120
      OLEDropMode     =   1  'Manual
      TabIndex        =   96
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   6360
      OLEDropMode     =   1  'Manual
      TabIndex        =   95
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   14
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   94
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   15
      Left            =   5160
      OLEDropMode     =   1  'Manual
      TabIndex        =   93
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   16
      Left            =   5400
      OLEDropMode     =   1  'Manual
      TabIndex        =   92
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   17
      Left            =   5640
      OLEDropMode     =   1  'Manual
      TabIndex        =   91
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   18
      Left            =   5880
      OLEDropMode     =   1  'Manual
      TabIndex        =   90
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   19
      Left            =   6120
      OLEDropMode     =   1  'Manual
      TabIndex        =   89
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   6360
      OLEDropMode     =   1  'Manual
      TabIndex        =   88
      Top             =   1560
      Width           =   210
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   21
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   87
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   22
      Left            =   5160
      OLEDropMode     =   1  'Manual
      TabIndex        =   86
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   23
      Left            =   5400
      OLEDropMode     =   1  'Manual
      TabIndex        =   85
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   24
      Left            =   5640
      OLEDropMode     =   1  'Manual
      TabIndex        =   84
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   25
      Left            =   5880
      OLEDropMode     =   1  'Manual
      TabIndex        =   83
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   26
      Left            =   6120
      OLEDropMode     =   1  'Manual
      TabIndex        =   82
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   6360
      OLEDropMode     =   1  'Manual
      TabIndex        =   81
      Top             =   1800
      Width           =   210
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   28
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   80
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   29
      Left            =   5160
      OLEDropMode     =   1  'Manual
      TabIndex        =   79
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   30
      Left            =   5400
      OLEDropMode     =   1  'Manual
      TabIndex        =   78
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   31
      Left            =   5640
      OLEDropMode     =   1  'Manual
      TabIndex        =   77
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   32
      Left            =   5880
      OLEDropMode     =   1  'Manual
      TabIndex        =   76
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   33
      Left            =   6120
      OLEDropMode     =   1  'Manual
      TabIndex        =   75
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   6360
      OLEDropMode     =   1  'Manual
      TabIndex        =   74
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   73
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   72
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   71
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   70
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   69
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   68
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   67
      Top             =   1080
      Width           =   210
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   66
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   65
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   64
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   63
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   62
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   61
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   60
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   59
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   58
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   57
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   56
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   55
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   54
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   53
      Top             =   1560
      Width           =   210
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   52
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   51
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   50
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   49
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   48
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   47
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   46
      Top             =   1800
      Width           =   210
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   45
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   44
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   43
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   42
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   41
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   40
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   39
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   38
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   37
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   36
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   35
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   34
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   33
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   32
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   31
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   30
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   29
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   28
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   27
      Top             =   1800
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   25
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   24
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   23
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   22
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   21
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   20
      Top             =   1560
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   19
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   18
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   17
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   15
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   12
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   1080
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   450
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   449
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   448
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4800
      TabIndex        =   456
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2520
      TabIndex        =   455
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   454
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   451
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   452
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label LabelMes 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   453
      Top             =   4680
      Width           =   1815
   End
End
Attribute VB_Name = "frmVerCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public CodigoTrab As Integer
Public Texto As String
Public idCal As Integer

Public FeIni As Date
Public FeFin As Date


Dim IT As ListItem
Dim PrimeraVez As Boolean
Dim Desplazamiento As Integer 'Desplazamiento para cargar temporads partidas



Private Sub LabelDia(mes As Integer, Indice As Integer, Texto As String)
    Select Case mes
    Case 1
        Label1(Indice).Caption = Texto
    Case 2
        Label2(Indice).Caption = Texto
    Case 3
        Label3(Indice).Caption = Texto
    Case 4
        Label4(Indice).Caption = Texto
    Case 5
        Label5(Indice).Caption = Texto
    Case 6
        Label6(Indice).Caption = Texto
    Case 7
        Label7(Indice).Caption = Texto
    Case 8
        Label8(Indice).Caption = Texto
    Case 9
        Label9(Indice).Caption = Texto
    Case 10
        Label10(Indice).Caption = Texto
    Case 11
        Label11(Indice).Caption = Texto
    Case 12
        Label12(Indice).Caption = Texto
    End Select
End Sub


Private Sub PintaDia(mes As Integer, Indice As Integer, IdHorario As Integer)
Dim Color
Dim vMEs As Integer
    If IdHorario > 15 Then
        Color = QBColor(0)
    Else
        Color = QBColor(IdHorario)
    End If
    
    If Desplazamiento = 0 Then
        vMEs = mes
    Else
        vMEs = mes + Desplazamiento
        If vMEs < 1 Then vMEs = 12 + vMEs
    End If
    'El primer dia
    Indice = LabelMes(vMEs - 1).Tag + Indice - 1

    Select Case vMEs
    Case 1
        Label1(Indice).ForeColor = Color
    Case 2
        Label2(Indice).ForeColor = Color
    Case 3
        Label3(Indice).ForeColor = Color
    Case 4
        Label4(Indice).ForeColor = Color
    Case 5
        Label5(Indice).ForeColor = Color
    Case 6
        Label6(Indice).ForeColor = Color
    Case 7
        Label7(Indice).ForeColor = Color
    Case 8
        Label8(Indice).ForeColor = Color
    Case 9
        Label9(Indice).ForeColor = Color
    Case 10
        Label10(Indice).ForeColor = Color
    Case 11
        Label11(Indice).ForeColor = Color
    Case 12
        Label12(Indice).ForeColor = Color
    End Select
End Sub



Private Sub PintaFestivo(mes As Integer, Indice As Integer)
Dim vMEs As Integer
    
    If Desplazamiento = 0 Then
        vMEs = mes
    Else
        vMEs = mes + Desplazamiento
        If vMEs < 1 Then vMEs = 12 + vMEs
    End If
    'El primer dia
    Indice = LabelMes(vMEs - 1).Tag + Indice - 1
    
    
    
    Select Case vMEs
    Case 1
        Label1(Indice).ForeColor = vbRed
    Case 2
        Label2(Indice).ForeColor = vbRed
    Case 3
        Label3(Indice).ForeColor = vbRed
    Case 4
        Label4(Indice).ForeColor = vbRed
    Case 5
        Label5(Indice).ForeColor = vbRed
    Case 6
        Label6(Indice).ForeColor = vbRed
    Case 7
        Label7(Indice).ForeColor = vbRed
    Case 8
        Label8(Indice).ForeColor = vbRed
    Case 9
        Label9(Indice).ForeColor = vbRed
    Case 10
        Label10(Indice).ForeColor = vbRed
    Case 11
        Label11(Indice).ForeColor = vbRed
    Case 12
        Label12(Indice).ForeColor = vbRed
    End Select
End Sub

Private Sub CargarElCalendario()


    If Year(FeIni) = Year(FeFin) Then
        
    
    
    Else
        'TEmporadas
    
    End If

End Sub

Private Sub CargaCalendario(MesIni As Integer, MesFin As Integer, Anyo As Integer)
Dim Dias As Integer
Dim PrimerDia As Integer
Dim i As Integer
Dim J As Integer
Dim vLabel As Integer
Dim AnyadeAnyo As String

    If MesFin - MesIni <> 11 Then
        AnyadeAnyo = "  " & Anyo
    Else
        AnyadeAnyo = ""
    End If
    
    
    For J = MesIni To MesFin
        vLabel = J + Desplazamiento
        If vLabel < 1 Then vLabel = 12 + vLabel
        
    
        Me.LabelMes(vLabel - 1).Caption = UCase(Format("01/" & J & "/" & Anyo, "mmm")) & AnyadeAnyo
        
        'cARGAMOS EL MES
        '----------------------
        Dias = DiasMes(J, Anyo)
        PrimerDia = Format("01/" & J & "/" & Anyo, "w", vbMonday)
        
        Me.LabelMes(vLabel - 1).Tag = PrimerDia - 1 'Guardo el desplazamiento aqui
        'Vaciamos los primeros textos
        If PrimerDia > 1 Then
            For i = 0 To PrimerDia - 2
                LabelDia vLabel, i, ""
            Next i
        End If
        
        'EL dia
        PrimerDia = PrimerDia - 1
        For i = 0 To Dias - 1
            LabelDia vLabel, i + PrimerDia, CStr(i + 1)
        Next i
        
        'El resto a ''
        i = i + PrimerDia
        While i <= 36
                LabelDia vLabel, i, ""
                i = i + 1
        Wend
       
    Next J
End Sub




Private Sub CargaTrabajador()
    miSQL = "Select fecha,idhorario from calendariot where fecha>='" & Format(vEmpresa.FechaInicio, FormatoFecha)
    miSQL = miSQL & "' and fecha<='" & Format(vEmpresa.FechaFin, FormatoFecha)
    miSQL = miSQL & "' and idtrabajador =" & CodigoTrab
    Set miRs = New ADODB.Recordset
    miRs.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRs.EOF
        PintaDia Month(miRs!Fecha), Day(miRs!Fecha), miRs!IdHorario
        miRs.MoveNext
    Wend
    miRs.Close
    Set miRs = Nothing
    PonDiasFestivos
End Sub

Private Sub CargaDatosCalendario()
    miSQL = "Select fecha,idhorario from calendariol where fecha>='" & Format(vEmpresa.FechaInicio, FormatoFecha)
    miSQL = miSQL & "' and fecha<='" & Format(vEmpresa.FechaFin, FormatoFecha)
    miSQL = miSQL & "' and idcal =" & idCal
    Set miRs = New ADODB.Recordset
    miRs.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRs.EOF
        PintaDia Month(miRs!Fecha), Day(miRs!Fecha), miRs!IdHorario
        miRs.MoveNext
    Wend
    miRs.Close
    Set miRs = Nothing
    PonDiasFestivos
End Sub




Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        DoEvents
        If CodigoTrab > 0 Then
            CargaTrabajador
        Else
            CargaDatosCalendario
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    miSQL = "Select idhorario,nomhorario from horarios order by idhorario"
    Set miRs = New ADODB.Recordset
    miRs.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRs.EOF
        Set IT = ListView2.ListItems.Add
        IT.Text = miRs!NomHorario
        IT.Tag = miRs!IdHorario
        IT.Bold = True
        If miRs!IdHorario < 15 Then
            IT.ForeColor = QBColor(miRs!IdHorario)
        Else
            IT.ForeColor = QBColor(0)
        End If
        'IT.SmallIcon = 6
        miRs.MoveNext
    Wend
    miRs.Close
    Set miRs = Nothing
    Set ListView2.SelectedItem = Nothing
    
'Public EsCalendario As Boolean
'Public Codigo As Integer
'Public Texto As String
    
    If CodigoTrab > 0 Then
        Label14.Caption = "Trabajador"
    Else
        Label14.Caption = "Calendario"
    End If
    Text1.Text = Texto
    
    
    If Year(FeIni) = Year(FeFin) Then
        Desplazamiento = 0
        CargaCalendario 1, 12, Year(vEmpresa.FechaInicio)
    Else
        Desplazamiento = 1 - Month(FeIni)
        CargaCalendario Month(FeIni), 12, Year(FeIni)
        CargaCalendario 1, Month(FeFin), Year(FeFin)
    End If
    
    
    
    CargaFestivos  'Los carga en el list
    PonDiasFestivos 'los pinta de rojo
End Sub



Private Sub CargaFestivos()
    miSQL = "Select fecha,descripcion from calendariof where idcal=" & idCal & " AND fecha>='" & Format(vEmpresa.FechaInicio, FormatoFecha)
    miSQL = miSQL & "' and fecha<='" & Format(vEmpresa.FechaFin, FormatoFecha)
    miSQL = miSQL & "' order by fecha"
    Set miRs = New ADODB.Recordset
    miRs.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRs.EOF
        Set IT = ListView1.ListItems.Add
        IT.Text = Format(miRs!Fecha, "dd/mm/yyyy")
        
        IT.SubItems(1) = miRs!descripcion
        'IT.SmallIcon = 6
        miRs.MoveNext
    Wend
    miRs.Close
    Set miRs = Nothing
    Set ListView1.SelectedItem = Nothing
    
End Sub

Private Sub PonDiasFestivos()
Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        PintaFestivo Month(CDate(ListView1.ListItems(i).Text)), Day(CDate(ListView1.ListItems(i).Text))
    Next i
End Sub




