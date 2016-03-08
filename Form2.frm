VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "frmVerCalendario"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12330
   LinkTopic       =   "Form2"
   ScaleHeight     =   8670
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8520
      TabIndex        =   42
      Top             =   7320
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5775
      Left            =   9960
      TabIndex        =   127
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   10186
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Label Label4 
      Caption         =   "MES"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   130
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "MES"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   129
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "MES"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   128
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   0
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   126
      Top             =   600
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
      TabIndex        =   125
      Top             =   600
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
      TabIndex        =   124
      Top             =   600
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
      TabIndex        =   123
      Top             =   600
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
      TabIndex        =   122
      Top             =   600
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
      TabIndex        =   121
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   6
      Left            =   6360
      OLEDropMode     =   1  'Manual
      TabIndex        =   120
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   7
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   119
      Top             =   840
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
      TabIndex        =   118
      Top             =   840
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
      TabIndex        =   117
      Top             =   840
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
      TabIndex        =   116
      Top             =   840
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
      TabIndex        =   115
      Top             =   840
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
      TabIndex        =   114
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   13
      Left            =   6360
      OLEDropMode     =   1  'Manual
      TabIndex        =   113
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   14
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   112
      Top             =   1080
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
      TabIndex        =   111
      Top             =   1080
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
      TabIndex        =   110
      Top             =   1080
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
      TabIndex        =   109
      Top             =   1080
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
      TabIndex        =   108
      Top             =   1080
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
      TabIndex        =   107
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   20
      Left            =   6360
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
      Index           =   21
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   105
      Top             =   1320
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
      TabIndex        =   104
      Top             =   1320
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
      TabIndex        =   103
      Top             =   1320
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
      TabIndex        =   102
      Top             =   1320
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
      TabIndex        =   101
      Top             =   1320
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
      TabIndex        =   100
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   27
      Left            =   6360
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
      Index           =   28
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   98
      Top             =   1560
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
      TabIndex        =   97
      Top             =   1560
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
      TabIndex        =   96
      Top             =   1560
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
      TabIndex        =   95
      Top             =   1560
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
      TabIndex        =   94
      Top             =   1560
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
      TabIndex        =   93
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   34
      Left            =   6360
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
      Index           =   35
      Left            =   4920
      OLEDropMode     =   1  'Manual
      TabIndex        =   91
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   36
      Left            =   5160
      OLEDropMode     =   1  'Manual
      TabIndex        =   90
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   37
      Left            =   5400
      OLEDropMode     =   1  'Manual
      TabIndex        =   89
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   38
      Left            =   5640
      OLEDropMode     =   1  'Manual
      TabIndex        =   88
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      DragMode        =   1  'Automatic
      Height          =   180
      Index           =   39
      Left            =   5880
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
      Index           =   40
      Left            =   6120
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
      Index           =   41
      Left            =   6360
      OLEDropMode     =   1  'Manual
      TabIndex        =   85
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   0
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   84
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   83
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   2
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   82
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   3
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   81
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   4
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   80
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   79
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   6
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   78
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   7
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   77
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   8
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   76
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   9
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   75
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   74
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   11
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   73
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   72
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   13
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   71
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   14
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   70
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   15
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   69
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   16
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   68
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   17
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   67
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   18
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   66
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   65
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   20
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   64
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   21
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   63
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   22
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   62
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   23
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   61
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   24
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   60
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   25
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   59
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   58
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   27
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   57
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   28
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   56
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   29
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   55
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   30
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   54
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   31
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   53
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   52
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   51
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   34
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   50
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   35
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   49
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   36
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   48
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   37
      Left            =   3120
      OLEDropMode     =   1  'Manual
      TabIndex        =   47
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   38
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   46
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   39
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   45
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   40
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   44
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   41
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   43
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   41
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   41
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   40
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   40
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   39
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   39
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   38
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   38
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   37
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   37
      Top             =   1800
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
      Top             =   1800
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
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   34
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   34
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   33
      Top             =   1560
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
      Top             =   1560
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
      Top             =   1560
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
      Top             =   1560
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
      Top             =   1560
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
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   27
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   27
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   26
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      Top             =   1320
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
      Top             =   1320
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
      Top             =   1320
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
      Top             =   1320
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
      Top             =   1320
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
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   20
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   20
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   19
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   19
      Top             =   1080
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
      Top             =   1080
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
      Top             =   1080
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
      Top             =   1080
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
      Top             =   1080
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
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   13
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   12
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   12
      Top             =   840
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
      Top             =   840
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
      Top             =   840
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
      Top             =   840
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
      Top             =   840
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
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   6
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   180
      Index           =   5
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   600
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
      Top             =   600
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
      Top             =   600
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
      Top             =   600
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
      Top             =   600
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
      Top             =   600
      Width           =   180
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer


    For i = 1 To 31
        Me.Label1(i - 1).Caption = i
        Me.Label2(i - 1).Caption = i
        Me.Label3(i - 1).Caption = i
    Next i
    For i = 32 To Label1.Count
        Me.Label1(i - 1).Caption = ""
        Me.Label2(i - 1).Caption = ""
        Me.Label3(i - 1).Caption = ""
    Next i
End Sub




Private Sub Form_Load()
miSQL = "Select idhorario,nomhorario from horarios order by idhorario"
    Set miRs = New ADODB.Recordset
    miRs.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRs.EOF
        Set IT = ListView2.ListItems.Add
        IT.Text = miRs!nomhorario
        IT.Tag = miRs!idhorario
        IT.Bold = True
        If miRs!idhorario < 15 Then
            IT.ForeColor = QBColor(miRs!idhorario)
        Else
            IT.ForeColor = QBColor(0)
        End If
        'IT.SmallIcon = 6
        miRs.MoveNext
    Wend
    miRs.Close
    Set miRs = Nothing
End Sub
