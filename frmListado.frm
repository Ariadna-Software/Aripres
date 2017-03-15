VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14880
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameA3 
      Height          =   2895
      Left            =   120
      TabIndex        =   298
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkA3 
         Caption         =   "Excel dias NO trabajados"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   301
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkA3 
         Caption         =   "Fichero integración"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   300
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton cmdGenNominaA3 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   302
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   4440
         TabIndex        =   303
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   20
         Left            =   2760
         ScrollBars      =   1  'Horizontal
         TabIndex        =   299
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   20
         Left            =   2400
         Picture         =   "frmListado.frx":6852
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Generar datos mes A3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   13
         Left            =   840
         TabIndex        =   305
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   31
         Left            =   1680
         TabIndex        =   304
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame FrameRelojesAuxiliares 
      Height          =   4575
      Left            =   120
      TabIndex        =   280
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkSinProcesar 
         Caption         =   "Agrupa por trabajador"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   286
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   21
         Left            =   1920
         TabIndex        =   285
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2880
         TabIndex        =   295
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   284
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2880
         TabIndex        =   292
         Top             =   2040
         Width           =   3375
      End
      Begin VB.CommandButton cmdRelojesAuxiliares 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   287
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   19
         Left            =   4680
         ScrollBars      =   1  'Horizontal
         TabIndex        =   283
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   18
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   282
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   4920
         TabIndex        =   288
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   60
         Left            =   960
         TabIndex        =   296
         Top             =   2400
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   21
         Left            =   1560
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   30
         Left            =   240
         TabIndex        =   294
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   960
         TabIndex        =   293
         Top             =   2040
         Width           =   540
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   20
         Left            =   1560
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   58
         Left            =   3720
         TabIndex        =   291
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   19
         Left            =   4320
         Picture         =   "frmListado.frx":68DD
         ToolTipText     =   "Buscar fecha"
         Top             =   1177
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   57
         Left            =   960
         TabIndex        =   290
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   18
         Left            =   1560
         Picture         =   "frmListado.frx":6968
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   29
         Left            =   240
         TabIndex        =   289
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   12
         Left            =   720
         TabIndex        =   281
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrameTrabajadores 
      Height          =   5535
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkTarjeta 
         Caption         =   "Tarjeta"
         Height          =   255
         Left            =   1440
         TabIndex        =   297
         Top             =   4680
         Width           =   975
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   240
         TabIndex        =   97
         Top             =   3960
         Width           =   2535
         Begin VB.OptionButton optOrdenTraba 
            Caption         =   "Codigo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   99
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optOrdenTraba 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   98
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame FrameTapaSecc 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   240
         TabIndex        =   96
         Top             =   2160
         Width           =   6135
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   44
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   45
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   95
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   94
         Top             =   3000
         Width           =   3375
      End
      Begin VB.CheckBox chkFoto 
         Caption         =   "Foto"
         Height          =   255
         Left            =   360
         TabIndex        =   90
         Top             =   4680
         Width           =   975
      End
      Begin VB.CheckBox chkSeccion 
         Caption         =   "Sección"
         Height          =   255
         Left            =   360
         TabIndex        =   89
         Top             =   3600
         Width           =   975
      End
      Begin VB.OptionButton optListTrab 
         Caption         =   "Datos extendidos"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   55
         Top             =   3600
         Width           =   1575
      End
      Begin VB.OptionButton optListTrab 
         Caption         =   "Datos básicos"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   54
         Top             =   3600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdTraba 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   46
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   8
         Left            =   5040
         TabIndex        =   49
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   48
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   47
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   43
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   42
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   0
         Left            =   1920
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   1
         Left            =   1920
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   20
         Left            =   1320
         TabIndex        =   93
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   19
         Left            =   1320
         TabIndex        =   92
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Seccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   91
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   5
         Left            =   1920
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   4
         Left            =   1920
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   1440
         TabIndex        =   53
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   1440
         TabIndex        =   52
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   51
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado trabajadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   50
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameActual 
      Height          =   5895
      Left            =   120
      TabIndex        =   137
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   147
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2880
         TabIndex        =   277
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   146
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2880
         TabIndex        =   274
         Top             =   3120
         Width           =   3375
      End
      Begin VB.CheckBox chkSinProcesar 
         Caption         =   "Agrupa por trabajador"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   150
         Top             =   4560
         Width           =   2175
      End
      Begin VB.OptionButton optActual 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   149
         Top             =   4080
         Width           =   855
      End
      Begin VB.OptionButton optActual 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   148
         Top             =   4080
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   5040
         TabIndex        =   152
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdActual 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   151
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2880
         TabIndex        =   154
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2880
         TabIndex        =   153
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   145
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   144
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   9
         Left            =   4320
         ScrollBars      =   1  'Horizontal
         TabIndex        =   141
         Text            =   "Text1"
         Top             =   1035
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   8
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   138
         Text            =   "Text1"
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Ordenado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   279
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   11
         Left            =   1560
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   960
         TabIndex        =   278
         Top             =   3480
         Width           =   465
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   10
         Left            =   1560
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   55
         Left            =   960
         TabIndex        =   276
         Top             =   3120
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   27
         Left            =   240
         TabIndex        =   275
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   11
         Left            =   1560
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   10
         Left            =   1560
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   30
         Left            =   960
         TabIndex        =   157
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   29
         Left            =   960
         TabIndex        =   156
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   155
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Marcajes sin procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   7
         Left            =   360
         TabIndex        =   143
         Top             =   240
         Width           =   5415
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   3960
         Picture         =   "frmListado.frx":69F3
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   3360
         TabIndex        =   142
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   140
         Top             =   840
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1560
         Picture         =   "frmListado.frx":6A7E
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   27
         Left            =   960
         TabIndex        =   139
         Top             =   1080
         Width           =   465
      End
   End
   Begin VB.Frame FrameCostesTrabajador 
      Height          =   6975
      Left            =   120
      TabIndex        =   245
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2760
         TabIndex        =   272
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   9
         Left            =   1800
         TabIndex        =   251
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2760
         TabIndex        =   269
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   8
         Left            =   1800
         TabIndex        =   250
         Top             =   2760
         Width           =   855
      End
      Begin VB.OptionButton optHorasPorecesadas 
         Caption         =   "Trabajador - Horas"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   268
         Top             =   5160
         Width           =   2055
      End
      Begin VB.OptionButton optHorasPorecesadas 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   267
         Top             =   5160
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox chkDesglosaDias 
         Caption         =   "Agrupa por empresa"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   252
         Top             =   5760
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   735
         Left            =   1080
         Style           =   1  'Checkbox
         TabIndex        =   265
         Top             =   4080
         Width           =   5055
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   19
         Left            =   1800
         TabIndex        =   249
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2760
         TabIndex        =   263
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox chkDesglosaDias 
         Caption         =   "Desglosa trabajador"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   253
         Top             =   5760
         Width           =   2055
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   18
         Left            =   1800
         TabIndex        =   248
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2760
         TabIndex        =   260
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   4920
         TabIndex        =   255
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdHorasProcesadas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   254
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   17
         Left            =   4920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   247
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   16
         Left            =   1800
         ScrollBars      =   1  'Horizontal
         TabIndex        =   246
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   54
         Left            =   840
         TabIndex        =   273
         Top             =   3120
         Width           =   465
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   9
         Left            =   1440
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   271
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   53
         Left            =   840
         TabIndex        =   270
         Top             =   2760
         Width           =   465
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   8
         Left            =   1440
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   266
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   52
         Left            =   840
         TabIndex        =   264
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   19
         Left            =   1440
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   262
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   51
         Left            =   840
         TabIndex        =   261
         Top             =   1800
         Width           =   540
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   18
         Left            =   1440
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   17
         Left            =   4560
         Picture         =   "frmListado.frx":6B09
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   50
         Left            =   4080
         TabIndex        =   259
         Top             =   960
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Horas procesadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   11
         Left            =   480
         TabIndex        =   258
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   840
         TabIndex        =   257
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   256
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   16
         Left            =   1440
         Picture         =   "frmListado.frx":6B94
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   360
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCostes 
      Height          =   5175
      Left            =   120
      TabIndex        =   216
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   244
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   223
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2880
         TabIndex        =   242
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   6
         Left            =   1920
         TabIndex        =   222
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   239
         Top             =   3000
         Width           =   3375
      End
      Begin VB.CommandButton cmdCostes 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   228
         Top             =   4560
         Width           =   1215
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   225
         Top             =   3960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   224
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2880
         TabIndex        =   237
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   17
         Left            =   1920
         TabIndex        =   221
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2880
         TabIndex        =   234
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   16
         Left            =   1920
         TabIndex        =   220
         Top             =   1920
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   3240
         TabIndex        =   233
         Top             =   3840
         Width           =   3015
         Begin VB.OptionButton optTrab 
            Caption         =   "Codigo"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   226
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTrab 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   227
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   14
         Left            =   2160
         ScrollBars      =   1  'Horizontal
         TabIndex        =   218
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   15
         Left            =   4560
         ScrollBars      =   1  'Horizontal
         TabIndex        =   219
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   5040
         TabIndex        =   229
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   7
         Left            =   1560
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   960
         TabIndex        =   243
         Top             =   3360
         Width           =   420
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   6
         Left            =   1560
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   960
         TabIndex        =   241
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   240
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   17
         Left            =   1560
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   960
         TabIndex        =   238
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   16
         Left            =   1560
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   45
         Left            =   960
         TabIndex        =   236
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   235
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   14
         Left            =   1800
         Picture         =   "frmListado.frx":6C1F
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   44
         Left            =   3720
         TabIndex        =   232
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   231
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   15
         Left            =   4200
         Picture         =   "frmListado.frx":6CAA
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   43
         Left            =   1200
         TabIndex        =   230
         Top             =   960
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Costes diarios / trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   10
         Left            =   480
         TabIndex        =   217
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame frHorascombinado 
      Height          =   5655
      Left            =   120
      TabIndex        =   158
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkHorasCombinadas 
         Caption         =   "Ajustar calendario"
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   211
         Top             =   4080
         Width           =   1815
      End
      Begin VB.CheckBox chkHorasCombinadas 
         Caption         =   "Horas decimal"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   210
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CheckBox chkHorasCombinadas 
         Caption         =   "Agrupar por Fecha"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   209
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   179
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   178
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   174
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   173
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdHorasCombinadas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   177
         Top             =   5160
         Width           =   1215
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   176
         Top             =   4560
         Width           =   975
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   175
         Top             =   4560
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   168
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   12
         Left            =   1920
         TabIndex        =   167
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2880
         TabIndex        =   166
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2880
         TabIndex        =   165
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   11
         Left            =   4320
         ScrollBars      =   1  'Horizontal
         TabIndex        =   161
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   10
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   160
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   4920
         TabIndex        =   159
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblCombinado 
         Height          =   255
         Left            =   240
         TabIndex        =   212
         Top             =   5040
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   182
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   36
         Left            =   960
         TabIndex        =   181
         Top             =   3120
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   35
         Left            =   960
         TabIndex        =   180
         Top             =   3600
         Width           =   420
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   3
         Left            =   1560
         Top             =   3600
         Width           =   255
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   2
         Left            =   1560
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado horas combinado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   8
         Left            =   480
         TabIndex        =   172
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   171
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   960
         TabIndex        =   170
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   960
         TabIndex        =   169
         Top             =   2400
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   13
         Left            =   1560
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   12
         Left            =   1560
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   32
         Left            =   960
         TabIndex        =   164
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   3960
         Picture         =   "frmListado.frx":6D35
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   163
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   31
         Left            =   3480
         TabIndex        =   162
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   10
         Left            =   1560
         Picture         =   "frmListado.frx":6DC0
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame frPresenciareal 
      Height          =   4695
      Left            =   120
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkPresReal 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1920
         TabIndex        =   208
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdPResenciaReal 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   66
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   67
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   5
         Left            =   4680
         ScrollBars      =   1  'Horizontal
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   1275
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   1275
         Width           =   1215
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2880
         TabIndex        =   69
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   68
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   65
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   6
         Left            =   1920
         TabIndex        =   64
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton optNomTra 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   61
         Top             =   3390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optNomTra 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   60
         Top             =   3390
         Width           =   975
      End
      Begin VB.Label Label2 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   77
         Top             =   4080
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   76
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   3840
         TabIndex        =   75
         Top             =   1320
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   4320
         Picture         =   "frmListado.frx":6E4B
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   74
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1680
         Picture         =   "frmListado.frx":6ED6
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   7
         Left            =   1560
         Top             =   2760
         Width           =   255
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   6
         Left            =   1560
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   73
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   72
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   71
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Presencia Real"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   3
         Left            =   600
         TabIndex        =   70
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrMarxTrab 
      Height          =   4815
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.OptionButton optNomTra 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   58
         Top             =   3390
         Width           =   975
      End
      Begin VB.OptionButton optNomTra 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   57
         Top             =   3390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Por fecha"
         Height          =   255
         Left            =   1920
         TabIndex        =   56
         Top             =   3360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   28
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   29
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   36
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   35
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1275
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   2
         Left            =   4680
         ScrollBars      =   1  'Horizontal
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1275
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   31
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdMarcajeTrabajador 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Marcajes por trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   40
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   39
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   960
         TabIndex        =   38
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   37
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   3
         Left            =   1560
         Top             =   2280
         Width           =   255
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   2
         Left            =   1560
         Top             =   2760
         Width           =   255
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1680
         Picture         =   "frmListado.frx":6F61
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   1080
         TabIndex        =   34
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   4320
         Picture         =   "frmListado.frx":6FEC
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   3840
         TabIndex        =   33
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Frameincres 
      Height          =   5655
      Left            =   120
      TabIndex        =   112
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CheckBox chkInci 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   215
         Top             =   4050
         Width           =   1095
      End
      Begin VB.CheckBox chkInci 
         Caption         =   "Mostrar detalle"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   214
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CheckBox chkInci 
         Caption         =   "Decimal"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   213
         Top             =   4440
         Width           =   1095
      End
      Begin VB.OptionButton optInci 
         Caption         =   "Nombre trab."
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   136
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton optInci 
         Caption         =   "Codigo trab."
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   135
         Top             =   4080
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdInci 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   134
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtDInci 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   132
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox txtInci 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   131
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtDInci 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   128
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtInci 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   127
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   119
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   118
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2880
         TabIndex        =   117
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2880
         TabIndex        =   116
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   7
         Left            =   5040
         TabIndex        =   115
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   6
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   11
         Left            =   5040
         TabIndex        =   113
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Image imgInci 
         Height          =   255
         Index           =   4
         Left            =   1560
         Top             =   3615
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   133
         Top             =   3285
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   130
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Image imgInci 
         Height          =   255
         Index           =   3
         Left            =   1560
         Top             =   3255
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   960
         TabIndex        =   129
         Top             =   3645
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Incidencia resumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   6
         Left            =   240
         TabIndex        =   126
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   125
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   960
         TabIndex        =   124
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   123
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   9
         Left            =   1560
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   8
         Left            =   1560
         Top             =   2160
         Width           =   255
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   4680
         Picture         =   "frmListado.frx":7077
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   960
         TabIndex        =   122
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1560
         Picture         =   "frmListado.frx":7102
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   4200
         TabIndex        =   121
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   120
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame frDiasTrabajados 
      Height          =   4815
      Left            =   120
      TabIndex        =   183
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   13
         Left            =   5160
         ScrollBars      =   1  'Horizontal
         TabIndex        =   186
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   12
         Left            =   2040
         TabIndex        =   185
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.CommandButton cmdDiasTrabajados 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   192
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox ChkDiasTrab 
         Caption         =   "Agrupar por seccion"
         Height          =   255
         Left            =   1320
         TabIndex        =   191
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   14
         Left            =   5160
         TabIndex        =   193
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   200
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   199
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   190
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   189
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   15
         Left            =   2160
         TabIndex        =   188
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   3000
         TabIndex        =   197
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   14
         Left            =   2160
         TabIndex        =   187
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   3000
         TabIndex        =   184
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   13
         Left            =   4800
         Picture         =   "frmListado.frx":718D
         ToolTipText     =   "Buscar fecha"
         Top             =   1177
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   42
         Left            =   4320
         TabIndex        =   207
         Top             =   1200
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   12
         Left            =   1680
         Picture         =   "frmListado.frx":7218
         ToolTipText     =   "Buscar fecha"
         Top             =   1177
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   41
         Left            =   1080
         TabIndex        =   206
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   205
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblDiasTrabajados 
         Height          =   255
         Left            =   240
         TabIndex        =   204
         Top             =   4320
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Seccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   203
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   40
         Left            =   1320
         TabIndex        =   202
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   39
         Left            =   1320
         TabIndex        =   201
         Top             =   3480
         Width           =   420
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   5
         Left            =   1920
         Top             =   3480
         Width           =   255
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   4
         Left            =   1920
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   38
         Left            =   1320
         TabIndex        =   198
         Top             =   2400
         Width           =   420
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   15
         Left            =   1920
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Dias trabajados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   9
         Left            =   600
         TabIndex        =   196
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   195
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   1320
         TabIndex        =   194
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   14
         Left            =   1920
         Top             =   2040
         Width           =   255
      End
   End
   Begin VB.Frame FrameIncidenciaGenerada 
      Height          =   2175
      Left            =   120
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdGeneraInci 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   84
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   87
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtHoraD 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   83
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   82
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtInci 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   80
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDInci 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   79
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "INCIDENCIAS GENERADAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   88
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hora Dec."
         Height          =   195
         Index           =   18
         Left            =   1800
         TabIndex        =   86
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hora"
         Height          =   195
         Index           =   17
         Left            =   360
         TabIndex        =   85
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Incide."
         Height          =   195
         Index           =   16
         Left            =   360
         TabIndex        =   81
         Top             =   720
         Width           =   465
      End
      Begin VB.Image imgInci 
         Height          =   255
         Index           =   2
         Left            =   960
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame FrameRevision 
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdRevisarMarcajes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtDInci 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   24
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtInci 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtDInci 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   23
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtInci 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   4
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   15
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   14
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox ChkIncorr 
         Caption         =   "Incorrectos"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CheckBox chkCorrec 
         Caption         =   "Correctos"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   9
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1035
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Image imgInci 
         Height          =   255
         Index           =   1
         Left            =   1560
         Top             =   3480
         Width           =   255
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   1
         Left            =   1560
         Top             =   2280
         Width           =   255
      End
      Begin VB.Image imgInci 
         Height          =   255
         Index           =   0
         Left            =   1560
         Top             =   3000
         Width           =   255
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   0
         Left            =   1560
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   22
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   21
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   19
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   18
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "REVISION MARCAJES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   3720
         TabIndex        =   12
         Top             =   1080
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   4200
         Picture         =   "frmListado.frx":72A3
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   11
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1560
         Picture         =   "frmListado.frx":732E
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
   End
   Begin VB.Frame FrCopiaHorario 
      Height          =   5295
      Left            =   120
      TabIndex        =   100
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdCopiaHorario 
         Caption         =   "Copiar"
         Height          =   375
         Left            =   5400
         TabIndex        =   111
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CheckBox chkTempoActual 
         Caption         =   "Temporada actual"
         Height          =   255
         Left            =   4800
         TabIndex        =   110
         Top             =   1560
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.TextBox txtCalendarioDestino 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox chkMas1año 
         Caption         =   "Incrementa 1 año"
         Height          =   255
         Left            =   4800
         TabIndex        =   107
         Top             =   1920
         Width           =   2535
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   240
         TabIndex        =   105
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5953
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.TextBox txtCalendD 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtCalen 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   10
         Left            =   6720
         TabIndex        =   102
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Image imgCalen 
         Height          =   255
         Index           =   0
         Left            =   1080
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Destino"
         Height          =   255
         Left            =   4800
         TabIndex        =   109
         Top             =   840
         Width           =   1695
      End
      Begin VB.Image imgcheckall 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmListado.frx":73B9
         ToolTipText     =   "Seleccionar todos"
         Top             =   4900
         Width           =   240
      End
      Begin VB.Image imgcheckall 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmListado.frx":7503
         ToolTipText     =   "Quitar seleccion"
         Top             =   4900
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Calendario"
         Height          =   255
         Left            =   240
         TabIndex        =   106
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Copiar  dias  festivos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   5
         Left            =   1200
         TabIndex        =   101
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0  .-  Revision marcajes
    '1  .-  Listado Marcajes  por trabajador
    '2  .-  Presemcia Real
    '3  .-  Ins/mod incidencia GEnerada
    
    
    '8  .-  Listado trabajadores
        
    '10 .- Copia horarios
    '11 .- Listado incidencias resumen(o final)
    '12 .- Marcas sin procesar
    '13 .- Listado horas combinado
    '14 .- Dias trabajados
    
    '15 .- Listado incicencias generadas
    
    '16 .- Costes diario trabajador
    
    '17.- Horas procesadas (En alzira desde la tabla
    
    '18.- Relojes auxiliares
    '19.- Relojes axiliares
    
    '20.- Nominas A3
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Dim Cad As String
Dim i  As Integer
Dim NumPa As Integer
Dim CadPa As String

Dim vSQL As String

'DESDE HASTA
'Campo:
    '0. Fecha
    '1. Trabajador
    '2. Seccion
    
Private Function DesdeHastaSelect(Txt As Byte, Indice As Integer, CampoTabla As String, LLevaAndOR As String)
Dim C As String

    Select Case Txt
    Case 0
        C = Me.txtFec(Indice).Text
        If C <> "" Then C = "'" & Format(C, FormatoFecha) & "'"
    Case 1
        C = Me.txtTrab(Indice).Text
    Case 2
        C = Me.txtSecc(Indice).Text
    Case 3
        C = Me.txtInci(Indice).Text
    Case Else
        C = ""
    End Select
    If C = "" Then
        DesdeHastaSelect = ""
        Exit Function
    End If
    
    DesdeHastaSelect = " " & LLevaAndOR & " " & CampoTabla & C
    
End Function


Private Function TieneDatos(ByRef C As String) As Boolean
On Error GoTo ET
    Set miRsAux = New ADODB.Recordset
    TieneDatos = False
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneDatos = True
        End If
    End If
    miRsAux.Close
ET:
    If Err.Number <> 0 Then MuestraError Err.Number, "Tiene datos" & vbCrLf & C
    Set miRsAux = Nothing

End Function


Private Sub chkA3_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub chkCorrec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerAccion
End Sub

Private Sub chkCorrec_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub


Private Sub chkDesglosaDias_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub chkFoto_Click()
    Me.chkTarjeta.Value = IIf(Me.chkFoto.Value = 1, 0, 1)
End Sub

Private Sub ChkIncorr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerAccion
End Sub

Private Sub ChkIncorr_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub chkSeccion_Click()
    Me.FrameTapaSecc.Visible = chkSeccion.Value = 0
End Sub

Private Sub chkSinProcesar_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub chkTarjeta_Click()
    
    Me.chkFoto.Value = IIf(Me.chkTarjeta.Value = 1, 0, 1)
End Sub

Private Sub cmdActual_Click()
Dim B As Boolean
        
    Screen.MousePointer = vbHourglass
    B = ImprimirTicajeActual
        
    Screen.MousePointer = vbDefault
    If B Then
        Cad = ""
        
        If txtSecc(10).Text <> "" Then Cad = Cad & "    desde " & txtSecc(10).Text & " - " & txtDSecc(10).Text
        If txtSecc(11).Text <> "" Then Cad = Cad & "    hasta " & txtSecc(11).Text & " - " & txtDSecc(11).Text
        If Cad <> "" Then Cad = "Sección: " & Trim(Cad)
        CadPa = ""
        If txtFec(8).Text <> "" Then CadPa = "Desde " & txtFec(8).Text
        If txtFec(9).Text <> "" Then CadPa = CadPa & "  Hasta " & txtFec(9).Text
        CadPa = Trim(CadPa & "   " & Cad)
        
        Cad = ""
        If txtTrab(10).Text <> "" Then Cad = "Desde " & txtTrab(10).Text & " - " & txtDT(10).Text
        If txtTrab(11).Text <> "" Then Cad = Cad & "    hasta " & txtTrab(11).Text & " - " & txtDT(11).Text
        Cad = Trim(Cad)
        If Cad <> "" Then
            If CadPa <> "" Then Cad = """ + chr(13) + """ & Cad
        End If
        CadPa = CadPa & Cad
        
        
        If optActual(0).Value Then
            Cad = "pOrden= {tmpcombinada.idtrabajador}|"
        Else
            Cad = "pOrden= {trabajadores.nomtrabajador}|"
        End If
        
        Cad = Cad & "DesdeHasta= """ & CadPa & """|"
        
        CadPa = "0"
        If vEmpresa.QueEmpresa = 0 Then CadPa = "1"
        Cad = Cad & "EsTeinsa= " & CadPa & "|"
        
        
        If Me.chkSinProcesar(0).Value = 1 Then
            frmImprimir.Opcion = 60
        Else
            frmImprimir.Opcion = 32
        End If
        frmImprimir.FormulaSeleccion = "{tmpcombinada.codusu} = " & vUsu.Codigo
        frmImprimir.OtrosParametros = Cad
        frmImprimir.NumeroParametros = 3

        frmImprimir.Show vbModal
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Ningún registro con esos valores", vbExclamation
    End If
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 3 Then CadenaDesdeOtroForm = ""   'Para que no refresque datos en el form de donde viene
    If Index = 12 Then CadenaDesdeOtroForm = ""   'Para que no refresque datos en el form de donde viene
    Unload Me
End Sub

Private Sub cmdCopiaHorario_Click()
Dim F As Date
    If txtCalen(0).Text = "" Then
        MsgBox "Seleccione calendario origen", vbExclamation
        Exit Sub
    End If
    
    
    Cad = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then Cad = Cad & "1"
    Next i
    
    If Cad = "" Then
        MsgBox "No tiene festivos para copiar", vbExclamation
        Exit Sub
    Else
        Cad = "Desea copiar " & Len(Cad) & " dia(s) sobre el calendario: " & Me.txtCalendarioDestino & "?"
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    'OK
    'A copiar
    Set miRsAux = New ADODB.Recordset
    CadPa = "INSERT INTO calendariof (idcal,fecha,descripcion) VALUES (" & Me.txtCalendarioDestino.Tag & ","
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
              F = CDate(ListView1.ListItems(i))
              If Me.chkMas1año.Value = 1 Then F = DateAdd("yyyy", 1, F)
              Cad = ListView1.ListItems(i).SubItems(1)
              NombreSQL Cad
              Cad = CadPa & "'" & Format(F, FormatoFecha) & "','" & Cad & "')"
              If Not ExisteElFestivo(F) Then EjecutaSQL Cad
        End If
    Next i
    Set miRsAux = Nothing
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Function ExisteElFestivo(Fec As Date) As Boolean
    ExisteElFestivo = False
    miRsAux.Open "Select fecha from calendariof where idcal=" & Me.txtCalendarioDestino.Tag & " AND fecha ='" & Format(Fec, FormatoFecha) & "'", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then ExisteElFestivo = True
    End If
    miRsAux.Close
End Function

Private Sub cmdCostes_Click()

    If vEmpresa.QueEmpresa = 2 Then
        GenerarImpresionimportesCostesDesdejornadassemanalesalz
    Else
        GenerarImpresionimportesCostesDesdeMarcajes
    End If
End Sub

Private Sub cmdDiasTrabajados_Click()

    'Hacemos las comprobaciones y demas
    If txtFec(12).Text = "" Or txtFec(13).Text = "" Then
        MsgBox "Ponga fecha desde/hasta", vbExclamation
        Exit Sub
    End If
    
    
    'Generamos los datos
    If GenerarDiasTrabajados Then
        
        Cad = ""
        If txtFec(12).Text <> "" Then Cad = "DESDE " & txtFec(12).Text
        If txtFec(13).Text <> "" Then Cad = Cad & "     HASTA " & txtFec(13).Text
        Cad = "LasFEchas= ""      " & Trim(Cad) & """|"
        CadPa = Cad
        Cad = ""
        If txtTrab(14).Text <> "" Then Cad = "Desde " & txtTrab(14).Text & " - " & txtDT(14).Text
        If txtTrab(15).Text <> "" Then Cad = Cad & "    hasta " & txtTrab(15).Text & " - " & txtDT(15).Text
        If txtSecc(4).Text <> "" Then Cad = Cad & "    desde " & txtSecc(4).Text & " - " & txtDSecc(4).Text
        If txtSecc(5).Text <> "" Then Cad = Cad & "    hasta " & txtSecc(5).Text & " - " & txtDSecc(5).Text
        Cad = "FechaIni= """ & Trim(Cad) & """|"
        CadPa = CadPa & Cad
        NumPa = 2
        With frmImprimir
            .FormulaSeleccion = "{tmpdatosmes.codusu}=" & vUsu.Codigo
            .OtrosParametros = CadPa
            .NumeroParametros = NumPa
            .Opcion = 40
            .Show vbModal
        End With
    End If

End Sub



Private Sub cmdExcel_Click(Index As Integer)
    
    If Index = 0 Then

    
    
            'Generamos los datos
            If GenerarImpresionimportesCostesDesdejornadassemanalesAlzira Then
                'Ya tenemos en tmpmarcajes. Ahora generaremos el xls
                GeneraExcel
                    
        
        
            End If
       End If
End Sub

Private Sub cmdGeneraInci_Click()
    If txtHora(0).Text = "" Or txtHoraD(0).Text = "" Or txtInci(2).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    If Me.FrameIncidenciaGenerada.Tag = "" Then
        'Insertar incidencia manual
        Cad = CStr(DameIncidenciaGenerada())
        Cad = "INSERT INTO incidenciasgeneradas (Id, EntradaMarcaje, Incidencia, horas) VALUES (" & Cad
        Cad = Cad & "," & NumRegElim & "," & txtInci(2).Text
        Cad = Cad & "," & TransformaComasPuntos(txtHoraD(0).Text) & ")"
    Else
        Cad = "UPDATE incidenciasgeneradas SET horas =" & TransformaComasPuntos(txtHoraD(0).Text)
        Cad = Cad & " ,incidencia = " & txtInci(2).Text
        Cad = Cad & " WHERE Id = " & FrameIncidenciaGenerada.Tag
        
    End If
    If EjecutaSQL(Cad) Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
    
End Sub

Private Function DameIncidenciaGenerada() As Long
Dim L As Long
Dim Fin As Boolean

    Cad = "Select id from incidenciasgeneradas order by id"
    L = 1
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Fin
        If Val(miRsAux.Fields(0)) - L > 0 Then
            DameIncidenciaGenerada = L
            Fin = True
        Else
            L = L + 1
            miRsAux.MoveNext
            If miRsAux.EOF Then
                DameIncidenciaGenerada = L
                Fin = True
            End If
        End If
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    'INSERT INTO incidenciasgeneradas (Id, EntradaMarcaje, Incidencia, horas) VALUES (1, 24274, 99, 1.5)

End Function


Private Sub cmdGenNominaA3_Click()
    If txtFec(20).Text = "" Then Exit Sub
    If chkA3(0).Value = 0 And chkA3(1).Value = 0 Then
        MsgBox "Seleccione alguna opcion de exportacion", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    If generarDatosNominas Then
    
        If Me.chkA3(0).Value Then
    
            GeneraNominaA3 CDate(Me.txtFec(20).Text)
        
            If vEmpresa.QueEmpresa = 5 Then
                'COOPIC lo copiamos en c:\Ariadna\enlaces
                CopiarFicheroAEnlaces
            End If
            
        
        End If
        If Me.chkA3(1).Value Then
            
        
            Screen.MousePointer = vbHourglass
            
            'Lanzamos el programa de EXCEL
            If Dir(App.Path & "\gestoriaCoopic.exe", vbArchive) = "" Then
                MsgBox "No existe programa enlace EXCEL", vbCritical
            Else
                vSQL = App.Path & "\gestoriaCoopic.exe"
                Lanza_EXE_Y_Espera vSQL
            End If
            Screen.MousePointer = vbDefault
        End If
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CopiarFicheroAEnlaces()
    On Error Resume Next
    
    FileCopy App.Path & "\nominaA3.txt", "C:\Ariadna\enlaceA3\ALMA" & Format(Now, "yymmdd_hhnn") & ".txt"
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
End Sub

Private Sub cmdHorasCombinadas_Click()
    If Not Comprobarfechas(10, 11) Then Exit Sub
    HacerListadoCombinado
End Sub

Private Sub cmdHorasProcesadas_Click()
Dim F1 As Date
    NumPa = 0
    CadPa = ""
    vSQL = ""
    Cad = ""
    
    
    
    F1 = "01/01/2001"
    If txtFec(16).Text <> "" Then F1 = CDate(txtFec(16).Text)
    NumPa = NumPa + 1
    CadPa = CadPa & "FechaIni= """ & Format(F1, "dd/mm/yyyy") & """|"
    vSQL = "({jornadassemanalesalz.fecha} >= Date(" & Year(F1) & "," & Month(F1) & "," & Day(F1) & ")"
    Cad = "Desde " & F1
    
    F1 = Now
    If txtFec(17).Text <> "" Then F1 = CDate(txtFec(17).Text)
    NumPa = NumPa + 1
    CadPa = CadPa & "FechaFin= """ & Format(F1, "dd/mm/yyyy") & """|"
    vSQL = vSQL & " AND {jornadassemanalesalz.fecha} <= Date(" & Year(F1) & "," & Month(F1) & "," & Day(F1) & "))"
    Cad = Cad & " hasta " & F1
  
        
    
    'Trabajador
    If txtTrab(18).Text <> "" Then
        Cad = Cad & "   Desde " & txtTrab(18).Text & " " & txtDT(18).Text
    End If
    vSQL = vSQL & DesdeHastaSelect(1, 18, "{jornadassemanalesalz.idtrabajador} >=", " AND ")
    
    If txtTrab(19).Text <> "" Then
        Cad = Cad & "hasta " & txtTrab(19).Text & " " & txtDT(19).Text
    End If
    vSQL = vSQL & DesdeHastaSelect(1, 19, "{jornadassemanalesalz.idtrabajador} <=", " AND ")
    
    'Seccion
    If txtSecc(8).Text <> "" Then
        Cad = Cad & "   Desde " & txtSecc(8).Text & " " & txtDSecc(8).Text
    End If
    vSQL = vSQL & DesdeHastaSelect(2, 8, "{trabajadores.Seccion} >=", " AND ")
    
    
    If txtSecc(9).Text <> "" Then
        Cad = Cad & "hasta " & txtSecc(9).Text & " " & txtDSecc(9).Text
    End If
    vSQL = vSQL & DesdeHastaSelect(2, 9, "{trabajadores.Seccion} <=", " AND ")
    
    
    
    
    'Para el resumen Hoja trabajador , para los desde hasta solo PINTO en el rpt la coopertaiva o no
    If Me.optHorasPorecesadas(1).Value Then Cad = ""
    
    
    'Desde /hasta empresa
    'Si estan las dos no pongo nada
    If vEmpresa.QueEmpresa = 2 Then
        If List1.Selected(0) Xor List1.Selected(1) Then
            If List1.Selected(0) Then
                Cad = Trim(Cad & "      " & List1.List(0))
                vSQL = vSQL & " AND {jornadassemanalesalz.ParaEmpresa}=0 "
            Else
                Cad = Trim(Cad & "      " & List1.List(1))
                vSQL = vSQL & " AND {jornadassemanalesalz.ParaEmpresa}=1 "
            End If
        End If
    End If
    
    
    
    CadPa = CadPa & "Intervalo= """ & Cad & """|"
    CadPa = CadPa & "DetallaTr= " & Abs(chkDesglosaDias(0).Value) & "|"
    NumPa = NumPa + 2
    
    
    
    If Me.optHorasPorecesadas(0).Value Then
        Cad = "AlzHorasProcesadasFecha"
        If Me.chkDesglosaDias(1).Value = 1 Then Cad = Cad & "Emp"
        Cad = Cad & ".rpt"
    
    Else
        'Desglose trabajador
        If vEmpresa.QueEmpresa = 5 Then
            Cad = "picHorasTrabajador.rpt"
        Else
            Cad = "AlzHorasTrabajador.rpt"
        End If
    End If
    
    With frmImprimir
        .FormulaSeleccion = vSQL
        .OtrosParametros = CadPa
        .NumeroParametros = NumPa
        .NombreRPT100 = Cad
        .Opcion = 65
        .Show vbModal
    End With
    
End Sub

Private Sub cmdInci_Click()
Dim l1 As String
Dim L2 As String

'lanzaremos el informe
    NumPa = 0
    CadPa = ""
    vSQL = ""
    l1 = ""
    L2 = ""
    If txtTrab(8).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "DTra= " & txtTrab(8).Text & "|"
        L2 = L2 & "desde " & txtTrab(8).Text & "-" & txtDT(8).Text & "   "
    End If
    vSQL = vSQL & DesdeHastaSelect(1, 8, "marcajes.idtrabajador>=", " AND ")
    If txtTrab(9).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "HTra= " & txtTrab(9).Text & "|"
        L2 = L2 & "hasta " & txtTrab(9).Text & "-" & txtDT(9).Text
    End If
    vSQL = vSQL & DesdeHastaSelect(1, 9, "marcajes.idtrabajador<=", " AND ")
    
    
    
    If txtFec(6).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "FechaIni= """ & txtFec(6).Text & """|"
        l1 = l1 & "desde " & txtFec(6).Text & "   "
    End If
    vSQL = vSQL & DesdeHastaSelect(0, 6, "Fecha>=", " AND ")
    
    If txtFec(7).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "FechaFin= """ & txtFec(7).Text & """|"
        l1 = l1 & "hasta " & txtFec(7).Text & "   "
    End If
    vSQL = vSQL & DesdeHastaSelect(0, 7, "Fecha<=", " AND ")
    
    
    If Opcion = 11 Then
        If txtInci(3).Text <> "" Then
            NumPa = NumPa + 1
            CadPa = CadPa & "inciini= " & txtInci(3).Text & "|"
            l1 = l1 & "desde " & txtInci(3).Text & "-" & txtDSecc(3).Text
        End If
        vSQL = vSQL & DesdeHastaSelect(3, 3, "incfinal>=", " AND ")
        
        If txtInci(4).Text <> "" Then
            NumPa = NumPa + 1
            CadPa = CadPa & "incifin= " & txtInci(4).Text & "|"
            l1 = l1 & "hasta " & txtInci(4).Text & "-" & txtDSecc(4).Text
        End If
        vSQL = vSQL & DesdeHastaSelect(3, 4, "incfinal<=", " AND ")
        
        Cad = "SELECT count(*) from marcajes"
        If vSQL <> "" Then Cad = Cad & " WHERE " & Mid(vSQL, 6)
    Else
        
        'Incidencia generada
        If txtInci(3).Text <> "" Then
            NumPa = NumPa + 1
            CadPa = CadPa & "inciini= " & txtInci(3).Text & "|"
            l1 = l1 & "desde " & txtInci(3).Text & "-" & txtDSecc(3).Text
        End If
        vSQL = vSQL & DesdeHastaSelect(3, 3, "incidencia>=", " AND ")
        
        If txtInci(4).Text <> "" Then
            NumPa = NumPa + 1
            CadPa = CadPa & "incifin= " & txtInci(4).Text & "|"
            l1 = l1 & "hasta " & txtInci(4).Text & "-" & txtDSecc(4).Text
        End If
        vSQL = vSQL & DesdeHastaSelect(3, 4, "incidencia<=", " AND ")
        
        Cad = "SELECT count(*) from marcajes,incidenciasgeneradas where incidenciasgeneradas.entradamarcaje=marcajes.entrada"
        If vSQL <> "" Then Cad = Cad & " AND " & Mid(vSQL, 6)
        
        
    End If
    
    If Not TieneDatos(Cad) Then
        MsgBox "Ningún registro con esos valores", vbExclamation
        Exit Sub
    End If
    l1 = Trim(l1)
    L2 = Trim(L2)
    CadPa = CadPa & "Linea1= """ & l1 & """|"
    CadPa = CadPa & "Linea2= """ & L2 & """|"
    NumPa = NumPa + 2
    
    'DEcimal  / sexagesimal
    vSQL = 0
    If chkInci(0).Value = 1 Then vSQL = 1
    CadPa = CadPa & "EnDecimal= " & vSQL & "|"
    
    
    
    If Opcion = 11 Then
        'INCIDENCIA RESUMEN
        If Me.chkInci(2).Value = 1 Then
            'Agrupada por trabajador
            i = 54
            
            
            'Detallar o no
            vSQL = 0
            If chkInci(1).Value = 1 Then vSQL = 1
            CadPa = CadPa & "Detallar= " & vSQL & "|"
            NumPa = NumPa + 1
        Else
            i = 33
        End If
        
        If Me.optInci(1).Value Then i = i + 1
    
    Else
        'Incidencia generada
        If Me.chkInci(2).Value = 1 Then
            'Agrupada por trabajador
            i = 52
        Else
            i = 50
        End If
        
        If Me.optInci(1).Value Then i = i + 1
        
        
        
        vSQL = 0
        If chkInci(1).Value = 1 Then vSQL = 1
        CadPa = CadPa & "Detallar= " & vSQL & "|"
        
        NumPa = NumPa + 1
    End If
    
    With frmImprimir
        .FormulaSeleccion = ""
        .OtrosParametros = CadPa
        .NumeroParametros = NumPa
        .Opcion = i
        .Show vbModal
    End With
        
End Sub

Private Sub cmdMarcajeTrabajador_Click()
    
    'lanzaremos el informe
    NumPa = 0
    CadPa = ""
    vSQL = ""
    
    If txtTrab(3).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "DTra= " & txtTrab(3).Text & "|"
    End If
    vSQL = vSQL & DesdeHastaSelect(1, 3, "Marcajes.idtrabajador>=", " AND ")
    If txtTrab(2).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "HTra= " & txtTrab(2).Text & "|"
    End If
    vSQL = vSQL & DesdeHastaSelect(1, 2, "Marcajes.idtrabajador <= ", " AND ")
    
    If txtFec(3).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "FechaIni= """ & txtFec(3).Text & """|"
    End If
    vSQL = vSQL & DesdeHastaSelect(0, 3, "Marcajes.Fecha>=", " AND ")
    If txtFec(2).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "FechaFin= """ & txtFec(2).Text & """|"
    End If
    vSQL = vSQL & DesdeHastaSelect(0, 2, "Marcajes.Fecha<=", " AND ")
    
    
    Cad = "SELECT count(*) from marcajes "
    If vSQL <> "" Then Cad = Cad & " WHERE " & Mid(vSQL, 6)  'le quito el primer AND
    If Not TieneDatos(Cad) Then
        MsgBox "Ningún registro con esos valores", vbExclamation
        Exit Sub
    End If
    If Me.Check1.Value = 0 Then
        i = 2
    Else
        i = 0
    End If
    If Me.optNomTra(1).Value Then i = i + 1 'Por nombre en lugar de por codigo
    With frmImprimir
        .FormulaSeleccion = ""
        .OtrosParametros = CadPa
        .NumeroParametros = NumPa
        .Opcion = i
        .Show vbModal
    End With
    
    
End Sub

Private Sub cmdPResenciaReal_Click()
    
    If Not Comprobarfechas(4, 5) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    If HacerPresReal Then
        'Imprimiremos
            CadPa = ""
            If txtFec(4).Text <> "" Then CadPa = txtFec(4).Text
            If txtFec(5).Text <> "" Then
                If CadPa <> "" Then CadPa = CadPa & "  "
                CadPa = CadPa & " hasta " & txtFec(5).Text
            End If
            Cad = ""
            If txtTrab(6).Text <> "" Then
                Cad = txtTrab(6).Text & " (" & txtDT(6).Text & ")"
            End If
            If txtTrab(7).Text <> "" Then
                If CadPa <> "" Then CadPa = CadPa & "  "
                Cad = Cad & " hasta " & txtTrab(7).Text & " (" & txtDT(7).Text & ")"
            End If
            Cad = Trim(CadPa & " " & Cad)
            CadPa = "FechaFin= """ & Cad & """|"
            NumPa = 1
            'Agrupado por
            If Me.chkPresReal.Value = 1 Then
                'POR FECHA
                i = 6
            Else
                'Por trabajador
                i = 8
            End If
            'Segundo orden. Codigo o nombre
            If optNomTra(3).Value Then i = i + 1
           
            With frmImprimir
                .FormulaSeleccion = "{tmppresencia.codusu} = " & vUsu.Codigo
                .OtrosParametros = CadPa
                .NumeroParametros = NumPa
                .Opcion = i
                .Show vbModal
            End With
    End If
  
    Screen.MousePointer = vbDefault
End Sub

Private Function HacerPresReal() As Boolean
Dim N As Long
Dim m As Long
Dim Inci As String
Dim anyo As Integer

    On Error GoTo EHacerPresReal
    HacerPresReal = False

    Cad = "Delete from tmppresencia where codusu =" & vUsu.Codigo
    conn.Execute Cad
    
    'HACemos el select
    Cad = "select entradamarcajes.* , nomtrabajador,nominci "
    Cad = Cad & " From entradamarcajes, trabajadores, incidencias"
    Cad = Cad & " Where entradamarcajes.idTrabajador = trabajadores.idTrabajador"
    Cad = Cad & " and entradamarcajes.idinci =incidencias.idinci"
    If txtFec(4).Text <> "" Then Cad = Cad & " and fecha >= '" & Format(txtFec(4).Text, FormatoFecha) & "'"
    If txtFec(5).Text <> "" Then Cad = Cad & " and fecha <= '" & Format(txtFec(5).Text, FormatoFecha) & "'"
    
    If txtTrab(6).Text <> "" Then Cad = Cad & " and entradamarcajes.idtrabajador >= " & txtTrab(6).Text
    If txtTrab(7).Text <> "" Then Cad = Cad & " and entradamarcajes.idtrabajador <= " & txtTrab(7).Text
    Cad = Cad & " order by idmarcaje,horareal"
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    N = 0
    m = 0
    Inci = ""
    CadPa = "INSERT INTO tmppresencia (codusu,Id,idtra, NomTrabajador"
    CadPa = CadPa & ",Fecha ,semana, H1, h2, H3, H4, H5, H6, H7, H8, Incidencias) VALUES (" & vUsu.Codigo & ","
    If Not miRsAux.EOF Then anyo = Year(miRsAux!Fecha)
    While Not miRsAux.EOF
        If miRsAux!idMarcaje <> m Then
            If m <> 0 Then
                'INSERTAMOS EL MARCAJE
                InsertaHoraReal i, Inci
            End If
           ' Label2.Caption = miRsAux!idTrabajador & " " & Format(miRsAux!Fecha, "ddmmyy")
           ' Label2.Refresh
            Inci = ""
            N = N + 1
            Cad = miRsAux!nomtrabajador
            NombreSQL Cad
            Cad = N & "," & miRsAux!idTrabajador & ",'" & Cad & "','" & Format(miRsAux!Fecha, FormatoFecha) & "'"
            'Semana
            i = (Year(miRsAux!Fecha) - anyo) * 100 + Format(miRsAux!Fecha, "ww", vbMonday)
            Cad = Cad & "," & i
            i = 0
            m = miRsAux!idMarcaje
        End If
        
        'Pongo la hora
        If miRsAux!IdInci <> 0 Then
            If Inci <> "" Then
                Inci = "Mas de una incidencia"
            Else
                
                Inci = miRsAux!NomInci
            End If
        End If
        
        i = i + 1
        If i <= 8 Then Cad = Cad & ",'" & Format(miRsAux!HoraReal, "hh:mm:ss") & "'"
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Esta pendiente el ultimo marcaje
    If m > 0 Then InsertaHoraReal i, Inci
    
   ' Label2.Caption = "Finalizando proceso"
    If N > 0 Then
        HacerPresReal = True
    Else
        MsgBox "Ningún registro con estos valores", vbExclamation
    End If
    Exit Function
EHacerPresReal:
    MuestraError Err.Number, Err.Description
End Function


Private Sub InsertaHoraReal(NumHoras As Integer, ByRef Incidencia As String)

    'CAD        ---> TIene los values a insertar
    'CadPara    ---> Tiene el INSERT INTO ......
    NumHoras = NumHoras + 1
    While NumHoras <= 8
        Cad = Cad & ",NULL"
        NumHoras = NumHoras + 1
    Wend
    If Incidencia <> "" Then
        NombreSQL Incidencia
        Incidencia = "'" & Incidencia & "'"
    Else
        Incidencia = "NULL"
    End If
    
    Cad = CadPa & Cad & "," & Incidencia & ")"
    conn.Execute Cad
    
    
    
End Sub


Private Sub cmdRelojesAuxiliares_Click()
  Dim B As Boolean
        
             
    Screen.MousePointer = vbHourglass
    B = ImprimirTicajeActualRelojesAuxiliares
        
    Screen.MousePointer = vbDefault
    If B Then
        Cad = ""
        
   '     If txtSecc(10).Text <> "" Then Cad = Cad & "    desde " & txtSecc(10).Text & " - " & txtDSecc(10).Text
   '     If txtSecc(11).Text <> "" Then Cad = Cad & "    hasta " & txtSecc(11).Text & " - " & txtDSecc(11).Text
   '     If Cad <> "" Then Cad = "Sección: " & Trim(Cad)
        CadPa = ""
        If txtFec(18).Text <> "" Then CadPa = "Desde " & txtFec(18).Text
        If txtFec(19).Text <> "" Then CadPa = CadPa & "  Hasta " & txtFec(19).Text
        CadPa = Trim(CadPa & "   " & Cad)
        
        Cad = ""
        If txtTrab(20).Text <> "" Then Cad = "Desde " & txtTrab(20).Text & " - " & txtDT(20).Text
        If txtTrab(21).Text <> "" Then Cad = Cad & "    hasta " & txtTrab(21).Text & " - " & txtDT(21).Text
        Cad = Trim(Cad)
        If Cad <> "" Then
            If CadPa <> "" Then Cad = """ + chr(13) + """ & Cad
        End If
        CadPa = CadPa & Cad
        
        
        If optActual(0).Value Then
            Cad = "pOrden= {tmpcombinada.idtrabajador}|"
        Else
            Cad = "pOrden= {trabajadores.nomtrabajador}|"
        End If
        
        Cad = Cad & "DesdeHasta= """ & CadPa & """|"
        
        CadPa = "0"
        If vEmpresa.QueEmpresa = 0 Then CadPa = "1"
        Cad = Cad & "EsTeinsa= " & CadPa & "|"
        
        
      With frmImprimir
            .FormulaSeleccion = "{tmpcombinada.codusu} = " & vUsu.Codigo
            If Opcion = 18 Then
                .NombreRPT100 = IIf(chkSinProcesar(1).Value = 0, "marcactualAux.rpt", "marcactualAuxT.rpt")
            Else
                .NombreRPT100 = IIf(chkSinProcesar(1).Value = 0, "marcactualAuxResum.rpt", "marcactualAuxTResum.rpt")
            End If
            .Titulo100 = "Relojes auxiliares"
            .OtrosParametros = Cad
            .Opcion = 100
            .NumeroParametros = 3
            .Show vbModal
      End With
       
    Else
        MsgBox "Ningún registro con esos valores", vbExclamation
    End If
        
        
        
        
End Sub

Private Sub cmdRevisarMarcajes_Click()


    If Not Comprobarfechas(0, 1) Then Exit Sub
    
    If txtTrab(0).Text <> "" And txtTrab(1).Text <> "" Then
        If Val(txtTrab(0).Text) > Val(txtTrab(1).Text) Then
            MsgBox "Error en desde hasta TRABAJADOR", vbExclamation
            Exit Sub
        End If
    End If
    If txtInci(0).Text <> "" And txtInci(1).Text <> "" Then
        If Val(txtInci(0).Text) > Val(txtInci(1).Text) Then
            MsgBox "Error en desde hasta INCIDENCIA", vbExclamation
            Exit Sub
        End If
    End If
    
    '-----------------------------------
    'Guardaremos en CadenaDesdeOtroForm
    'los datos para que los lea luego
    'el formulario de revision de marcajes
    
    'Fecha inico-fin
    CadenaDesdeOtroForm = txtFec(0).Text & "|" & txtFec(1).Text & "|"
    'Trabajador incio -fin
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & txtTrab(0).Text & "|" & txtTrab(1).Text & "|"
    'Incidencia
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & txtInci(0).Text & "|" & txtInci(1).Text & "|"
    i = Me.ChkIncorr.Value * 2
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.chkCorrec.Value + i & "|"
    Unload Me
End Sub


Private Function Comprobarfechas(Desde As Integer, Hasta As Integer) As Boolean
    Comprobarfechas = False
    
    If Me.txtFec(Desde).Text <> "" Then
        If Not EsFechaOK(txtFec(Desde)) Then Exit Function
    End If
    If Me.txtFec(Hasta).Text <> "" Then
        If Not EsFechaOK(txtFec(Hasta)) Then Exit Function
    End If
    If Me.txtFec(Desde).Text = "" Or txtFec(Hasta).Text = "" Then
        Comprobarfechas = True
        Exit Function
    End If
    
    If CDate(Me.txtFec(Desde).Text) > CDate(txtFec(Hasta).Text) Then
        MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
        Exit Function
    End If
    Comprobarfechas = True
End Function


Private Sub cmdTraba_Click()
Dim d As String
Dim formu As String

  'lanzaremos el informe
    NumPa = 0
    CadPa = ""
    vSQL = ""
    d = ""
    formu = ""
    If txtTrab(4).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "DTra= " & txtTrab(4).Text & "|"
        d = "Desde " & txtTrab(4).Text & " " & txtDT(4).Text
    End If
    vSQL = vSQL & DesdeHastaSelect(1, 4, "idtrabajador>=", " AND ")
    If txtTrab(5).Text <> "" Then
        NumPa = NumPa + 1
        CadPa = CadPa & "HTra= " & txtTrab(5).Text & "|"
        d = d & "   hasta " & txtTrab(5).Text & " " & txtDT(5).Text
    End If
    vSQL = vSQL & DesdeHastaSelect(1, 5, "idtrabajador<=", " AND ")
    
    
    If chkSeccion.Value = 1 Then
        If txtSecc(0).Text <> "" Then
            
            d = d & "   Desde " & txtSecc(0).Text & " " & txtDSecc(0).Text
            formu = "{secciones.idseccion}>=" & txtSecc(0).Text
        End If
        vSQL = vSQL & DesdeHastaSelect(2, 0, "trabajadores.seccion >=", " AND ")
                    
        If txtSecc(1).Text <> "" Then
            d = d & "   hasta " & txtSecc(1).Text & " " & txtDSecc(1).Text
            If formu <> "" Then formu = formu & " AND "
            formu = formu & "{secciones.idseccion}<=" & txtSecc(1).Text
        End If
        vSQL = vSQL & DesdeHastaSelect(2, 1, "trabajadores.seccion <=", " AND ")
    
    End If
    
    
    CadPa = CadPa & "CampoSeleccion= """ & Trim(d) & """|"
    NumPa = NumPa + 1
    
    i = 0
    i = i + (Me.chkFoto.Value * 4)
    If Me.chkSeccion.Value = 1 Then
        i = i + (Me.chkSeccion.Value * 8)
    Else
        i = i + (Me.chkSeccion.Value * 8)
       '''''''' If optListTrab(1).Value Then i = i + 2  'Solo para los que no hay seccion salen extendidos o basicos
    End If
    If optListTrab(1).Value Then i = i + 2
    If optOrdenTraba(1).Value Then i = i + 1
    
    
    
    d = "SELECT count(*) from trabajadores"
    If vSQL <> "" Then d = d & " WHERE " & Mid(vSQL, 6)
    If Not TieneDatos(d) Then
        MsgBox "Ningun registro con esos valores", vbExclamation
        Exit Sub
    End If
    
    i = i + 10 'Pq empiezan los listado en el 10
    
    
    With frmImprimir
        If Me.chkTarjeta.Value = 1 Then
            i = 100
            .NombreRPT100 = "trabtarjeta.rpt"
            .Titulo100 = "Tarjetas"
            .ConSubreport100 = False
        End If
        
        
        .FormulaSeleccion = formu
        .OtrosParametros = CadPa
        .NumeroParametros = NumPa
        .Opcion = i
        .Show vbModal
    End With

        
    
    
End Sub



Private Sub Command1_Click()
    
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim IndiceCancelar As Integer

    Me.Icon = frmMain.Icon

    Me.FrameRevision.Visible = False
    FrameIncidenciaGenerada.Visible = False
    FrameTrabajadores.Visible = False
    Me.FrCopiaHorario.Visible = False
    Frameincres.Visible = False
    FrameActual.Visible = False
    frHorascombinado.Visible = False
    FrameCostes.Visible = False
    FrameCostesTrabajador.Visible = False
    FrameRelojesAuxiliares.Visible = False
    Me.frDiasTrabajados.Visible = False
    FrameA3.Visible = False
    Limpiar Me
    IndiceCancelar = Opcion
    
    Select Case Opcion
    Case 0
        'Cargo las imagenes
        For H = 0 To 1
            imgTra(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
            imgInci(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next H
        'Para revision marcajes
        H = Me.FrameRevision.Height
        W = Me.FrameRevision.Width
        FrameRevision.Visible = True
        Caption = "Revisión marcajes"
        txtFec(1).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
    Case 1
        'Listado de marcajes por trabjaodr
        For H = 2 To 3
            imgTra(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next H
        H = Me.FrMarxTrab.Height
        W = Me.FrMarxTrab.Width
        
        FrMarxTrab.Visible = True
        Caption = "Listado marcajes"
        txtFec(2).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
    
    Case 2
        'Listado de PRESENCIA REAL
        For H = 4 To 5
            imgTra(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next H
        H = Me.frPresenciareal.Height
        W = Me.frPresenciareal.Width
        
        frPresenciareal.Visible = True
        Caption = "Listado presencia"
        txtFec(5).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
    
    
    Case 3
        'INSERTAR MODIFICAR INCIDENCIAS GENERADAS
        H = Me.FrameIncidenciaGenerada.Height
        W = Me.FrameIncidenciaGenerada.Width
        
        FrameIncidenciaGenerada.Visible = True
        Caption = "Inciden. generadas"
        imgInci(2).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        
        If CadenaDesdeOtroForm <> "" Then
            txtDInci(2).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            txtInci(2).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
            txtHora(0).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            txtHoraD(0).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
            FrameIncidenciaGenerada.Tag = RecuperaValor(CadenaDesdeOtroForm, 5)
        Else
            FrameIncidenciaGenerada.Tag = ""
        End If


    
    Case 8
        
        'INSERTAR MODIFICAR INCIDENCIAS GENERADAS
        H = Me.FrameTrabajadores.Height
        W = Me.FrameTrabajadores.Width
        
        FrameTrabajadores.Visible = True
        Caption = "Trabajadores"
        
        imgTra(4).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        imgTra(5).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        
    Case 10
        
        'INSERTAR MODIFICAR INCIDENCIAS GENERADAS
        H = Me.FrCopiaHorario.Height
        W = Me.FrCopiaHorario.Width
        
        FrCopiaHorario.Visible = True
        Caption = "Calendario"
        imgCalen(0).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        txtCalendarioDestino.Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        txtCalendarioDestino.Tag = RecuperaValor(CadenaDesdeOtroForm, 2)
    Case 11, 15
    
        ' I N C I D E N C I A S
        '   11: Resumen
        '   15: Generadas
        For H = 3 To 4
            imgTra(H + 5).Picture = frmPpal.imgListImages16.ListImages(3).Picture
            imgInci(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next H
        
        
        
        Frameincres.Visible = True
        H = Me.Frameincres.Height
        W = Me.Frameincres.Width
        lblTitulo(6).Caption = "Incidencias "
        If Opcion = 11 Then
            lblTitulo(6).Caption = lblTitulo(6).Caption & "resumen"
        Else
            lblTitulo(6).Caption = lblTitulo(6).Caption & "generadas"
        End If
        
        Frameincres.Visible = True
        Caption = "Incidencias"
        imgCalen(0).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        
        
    Case 12
        ' I N C I D E N C I A S
        
        imgTra(10).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        imgTra(11).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        FrameActual.Visible = True
        H = Me.FrameActual.Height
        W = Me.FrameActual.Width
        
        
        txtFec(8).Text = CadenaDesdeOtroForm
        txtFec(9).Text = CadenaDesdeOtroForm
        If CadenaDesdeOtroForm <> "" Then PonerFocoBtn Me.cmdActual
        CadenaDesdeOtroForm = ""
        Caption = "Actual"
        
        
    Case 13
        Caption = "Horas combinado"
        Me.frHorascombinado.Visible = True
        For H = 2 To 3
            imgTra(10 + H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
            imgSecc(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next
        frHorascombinado.Visible = True
        H = Me.frHorascombinado.Height
        W = Me.frHorascombinado.Width
        
        
    Case 14
        Caption = "Dias trabajados"
         
        Me.frDiasTrabajados.Visible = True
        For H = 4 To 5
            imgTra(10 + H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
            imgSecc(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next
        frDiasTrabajados.Visible = True
        H = Me.frDiasTrabajados.Height
        W = frDiasTrabajados.Width
        txtFec(12).Text = Format(DateAdd("m", -1, Now), "dd/mm/yyyy")
        txtFec(13).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
    
    Case 16
         For H = 16 To 17
            imgTra(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
            imgSecc(H - 10).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next H
        FrameCostes.Visible = True
        H = Me.FrameCostes.Height
        W = FrameCostes.Width
        txtFec(14).Text = "01/" & Format(Now, "mm/yyyy")
        
        
    Case 17
        
        For H = 18 To 19
            imgTra(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
            'La seccion son el index 8,9
            imgSecc(H - 10).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next H
        FrameCostesTrabajador.Visible = True
        H = Me.FrameCostesTrabajador.Height
        W = FrameCostesTrabajador.Width
        If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = Format(DateAdd("d", -7, Now), "dd/mm/yyyy")
            
        txtFec(16).Text = CadenaDesdeOtroForm
        CadenaDesdeOtroForm = ""
        Caption = "Listado"
        
        
        'Igual deberia venir de una BDs
        If vEmpresa.QueEmpresa = 2 Then
            List1.AddItem "Fruxeresa"
            List1.Selected(0) = True
            List1.AddItem "Cooperativa"
            List1.Selected(1) = True
        Else
            List1.AddItem "Cooperativa"
            List1.Selected(0) = True
        End If

    
    Case 18, 19
        Caption = "Relojes auxiliares"
        
        If Opcion = 18 Then
            lblTitulo(12).Caption = "Listado marcajes relol aux"
        Else
            lblTitulo(12).Caption = "Listado resumen relojes auxiliares "
        End If
        
        FrameRelojesAuxiliares.Visible = True
        H = Me.FrameRelojesAuxiliares.Height
        W = FrameRelojesAuxiliares.Width
        txtFec(18).Text = CadenaDesdeOtroForm
        txtFec(19).Text = CadenaDesdeOtroForm
    
    
    Case 20
        FrameA3.Visible = True
        H = Me.FrameA3.Height
        W = FrameA3.Width
        vSQL = DevuelveDesdeBD("max(fecha)", "nominas", "1", "1")
        If vSQL = "" Then vSQL = Now
        Caption = "Exportación"
        txtFec(20).Text = Format(vSQL, "dd/mm/yyyy")

    End Select
    
    Me.Height = H + 500
    Me.Width = W + 300
    
    If Opcion = 15 Then
        i = 11  'COmo es lo mismo para la opcion 15 que la 11..
    ElseIf Opcion = 19 Then
        i = 18
    Else
        i = Opcion
    End If
    Me.cmdCancelar(i).Cancel = True



    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    
    'EN imgTra(0).Tag  tengo que opcion  ha sido (trabajadores, incidencias...
    ' En imgTra(0).Tag  tendre que INDEX dentro del img
    
    Select Case imgTra(0).Tag
    Case 0
        'TRABAJADORES
        txtTrab(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 1)
        txtDT(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 2)
        
    Case 1
        'INCIDENCIAS
        txtInci(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 1)
        txtDInci(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 2)
    Case 2
        'CALENDARIOS
        txtCalen(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 1)
        txtCalendD(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 2)
        CargarFestivos
    Case 3
        txtSecc(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 1)
        txtDSecc(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 2)
        
    End Select

End Sub

Private Sub frmc_Selec(vFecha As Date)
    txtFec(CInt(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCalen_Click(Index As Integer)
    imgTra(0).Tag = 2 'Para que el devuelve grid sepa que es CALENDARIOS
    imgInci(0).Tag = Index 'Dentro de calendarios, que INDEX
    Cad = "Codigo|idcal|N||15·"
    Cad = Cad & "Descripción|Descripcion|T||85·"
    Set frmB = New frmBuscaGrid
    frmB.vTabla = "Calendario"
    frmB.vCampos = Cad
    frmB.vDevuelve = "0|1|"
    frmB.vSelElem = 0
    frmB.vTitulo = "CALENDARIO"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub imgcheckall_Click(Index As Integer)
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = (Index = 0)
    Next i
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim Obj As Object

    Set frmC = New frmCal
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    
    Set Obj = imgFec(Index).Container
    
    While imgFec(Index).Parent.Name <> Obj.Name
        esq = esq + Obj.Left
        dalt = dalt + Obj.Top
        Set Obj = Obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
    
    
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtFec(Index).Text <> "" Then frmC.NovaData = txtFec(Index).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtFec(CByte(imgFec(0).Tag)) '<===
    ' ********************************************
End Sub


Private Sub imgInci_Click(Index As Integer)
    imgTra(0).Tag = 1 'Para que el devuelve grid sepa que es INCIDENCIAS
    imgInci(0).Tag = Index 'Dentro de trabajadores, que INDEX
    Cad = "Codigo|idInci|N||15·"
    Cad = Cad & "Nombre|nominci|T||60·"
    Set frmB = New frmBuscaGrid
    frmB.vTabla = "Incidencias"
    frmB.vCampos = Cad
    frmB.vDevuelve = "0|1|"
    frmB.vSelElem = 0
    frmB.vTitulo = "INCIDENCIAS"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub imgSecc_Click(Index As Integer)
    imgTra(0).Tag = 3 'Para que el devuelve grid sepa que es SECCION
    imgInci(0).Tag = Index 'Dentro de trabajadores, que INDEX
    Cad = "Codigo|idseccion|N||15·"
    Cad = Cad & "Descripcion|nombre|T||75·"
    Set frmB = New frmBuscaGrid
    frmB.vTabla = "Secciones"
    frmB.vCampos = Cad
    frmB.vDevuelve = "0|1|"
    frmB.vSelElem = 0
    frmB.vTitulo = "Secciones"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub imgTra_Click(Index As Integer)
    imgTra(0).Tag = 0 'Para que el devuelve grid sepa que es TRABAJADORES
    imgInci(0).Tag = Index 'Dentro de trabajadores, que INDEX
    Cad = "Codigo|idTrabajador|N||15·"
    Cad = Cad & "Nombre|nomtrabajador|T||60·"
    Cad = Cad & "Tarjeta|numtarjeta|T||20·"
    Set frmB = New frmBuscaGrid
    frmB.vTabla = "Trabajadores"
    frmB.vCampos = Cad
    frmB.vDevuelve = "0|1|"
    frmB.vSelElem = 0
    frmB.vTitulo = "TRABAJADORES"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub



Private Sub optActual_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub optHorasPorecesadas_Click(Index As Integer)
    chkDesglosaDias(0).Visible = Index = 0
    chkDesglosaDias(1).Visible = Index = 0
End Sub

Private Sub optTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtCalen_GotFocus(Index As Integer)
    txtCalen(Index).SelStart = 0
    txtCalen(Index).SelLength = Len(txtCalen(Index).Text)
End Sub

Private Sub txtCalen_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtCalen_LostFocus(Index As Integer)
    txtInci(Index).Text = Trim(txtInci(Index).Text)
    If txtCalen(Index).Text = "" Then
        Me.txtCalendD(Index).Text = ""

    Else
        If Not EsEntero(txtCalen(Index).Text) Then
            MsgBox "Número incorrecto: " & txtCalen(Index).Text, vbExclamation
            txtCalen(Index).Text = ""
            txtCalendD(Index).Text = ""
            PonerFoco txtCalen(Index)
        Else
            Cad = DevuelveDesdeBD("descripcion", "calendario", "idcal", txtCalen(Index).Text, "N")
            
            txtCalendD(Index).Text = Cad
        End If
    End If
            
    DoEvents
    Screen.MousePointer = vbHourglass
    CargarFestivos
    Screen.MousePointer = vbDefault

End Sub

Private Sub txtFec_GotFocus(Index As Integer)
    txtFec(Index).SelStart = 0
    txtFec(Index).SelLength = Len(txtFec(Index).Text)
End Sub

Private Sub txtFec_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerAccion
End Sub

Private Sub txtFec_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtFec_LostFocus(Index As Integer)
    txtFec(Index).Text = Trim(txtFec(Index).Text)
    If txtFec(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(txtFec(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFec(Index).Text, vbExclamation
        txtFec(Index).Text = ""
        PonerFoco txtFec(Index)
    End If
End Sub


Private Sub KeyPress(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me 'ESC
    End If
End Sub


Private Sub HacerAccion()
    Select Case Opcion
    Case 0
        
            cmdRevisarMarcajes_Click
        
    Case 1
    
    
    End Select
End Sub






Private Sub txtHora_GotFocus(Index As Integer)
    ConseguirFoco txtHora(Index), 3
    
End Sub

Private Sub txtHora_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtHora_LostFocus(Index As Integer)
        txtHora(Index).Text = Trim(txtHora(Index).Text)
        If txtHora(Index).Text = "" Then
            Cad = ""
        Else
            txtHora(Index).Text = TransformaPunto2Puntos(txtHora(Index).Text)
            txtHora(Index).Text = Format(txtHora(Index).Text, "hh:mm:ss")
            If IsDate(txtHora(Index).Text) Then
                'Cambiamos las horas del campo decimal
                Cad = Format(DevuelveValorHora(CDate(txtHora(Index).Text)), "0.00")
            Else
                Cad = ""
                txtHora(Index).Text = ""
            End If
            
        End If
        txtHoraD(Index).Text = Cad
End Sub



Private Sub txtHoraD_GotFocus(Index As Integer)
    ConseguirFoco txtHoraD(Index), 3
End Sub

Private Sub txtHoraD_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtHoraD_LostFocus(Index As Integer)
    txtHoraD(Index).Text = Trim(txtHoraD(Index).Text)
    If txtHoraD(Index).Text = "" Then
        Cad = ""
    Else
        txtHoraD(Index).Text = TransformaPuntosComas(txtHoraD(Index).Text)
        If Not IsNumeric(txtHoraD(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            Cad = ""
        Else
            Cad = DevuelveHora(CSng(txtHoraD(Index).Text))
        End If
    End If
    txtHora(Index).Text = Cad
        
End Sub

Private Sub txtInci_GotFocus(Index As Integer)
    ConseguirFoco txtInci(Index), 3
End Sub

Private Sub txtInci_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerAccion
End Sub

Private Sub txtInci_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbBusca1 Or KeyAscii = vbBusca1 Then
        KeyAscii = 0
        KEYBusqueda 1, Index   '1.- INCIDENCIAS
    Else
        KeyPress KeyAscii
    End If
End Sub

Private Sub txtInci_LostFocus(Index As Integer)
    txtInci(Index).Text = Trim(txtInci(Index).Text)
    If txtInci(Index).Text = "" Then
        Me.txtInci(Index).Text = ""
        Exit Sub
    End If
    
    If Not EsEntero(txtInci(Index).Text) Then
        MsgBox "Número incorrecto: " & txtInci(Index).Text, vbExclamation
        txtInci(Index).Text = ""
        txtDInci(Index).Text = ""
        PonerFoco txtInci(Index)
    Else
        Cad = DevuelveDesdeBD("nominci", "incidencias", "idinci", txtInci(Index).Text, "N")
        If Cad = "" Then Cad = "NO EXISTE"
        txtDInci(Index).Text = Cad
    End If
End Sub




Private Sub txtSecc_GotFocus(Index As Integer)
    ConseguirFoco txtSecc(Index), 3
End Sub

Private Sub txtSecc_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtSecc_LostFocus(Index As Integer)
    txtSecc(Index).Text = Trim(txtSecc(Index).Text)
    If txtSecc(Index).Text = "" Then
        Me.txtDSecc(Index).Text = ""
        Exit Sub
    End If
    
    If Not EsEntero(txtSecc(Index).Text) Then
        MsgBox "Número incorrecto: " & txtSecc(Index).Text, vbExclamation
        txtSecc(Index).Text = ""
        txtDSecc(Index).Text = ""
        PonerFoco txtSecc(Index)
    Else
        Cad = DevuelveDesdeBD("nombre", "secciones", "idseccion", txtSecc(Index).Text, "N")
        If Cad = "" Then Cad = "NO EXISTE"
        txtDSecc(Index).Text = Cad
    End If
End Sub

Private Sub txtTrab_GotFocus(Index As Integer)
    ConseguirFoco txtTrab(Index), 3
End Sub

Private Sub txtTrab_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerAccion
End Sub

Private Sub txtTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbBusca1 Or KeyAscii = vbBusca1 Then
        KeyAscii = 0
        KEYBusqueda 0, Index   '0.- Trbajadores
    Else
        KeyPress KeyAscii
    End If
End Sub

Private Sub txtTrab_LostFocus(Index As Integer)
    txtTrab(Index).Text = Trim(txtTrab(Index).Text)
    If txtTrab(Index).Text = "" Then
        Me.txtDT(Index).Text = ""
        Exit Sub
    End If
    
    If Not EsEntero(txtTrab(Index).Text) Then
        MsgBox "Número incorrecto: " & txtTrab(Index).Text, vbExclamation
        txtTrab(Index).Text = ""
        txtDT(Index).Text = ""
        PonerFoco txtTrab(Index)
    Else
        Cad = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idtrabajador", txtTrab(Index).Text, "N")
        If Cad = "" Then Cad = "NO EXISTE"
        txtDT(Index).Text = Cad
    End If
End Sub




Private Sub KEYBusqueda(Opcion As Byte, Indice As Integer)
   
    Select Case Opcion
    Case 0
        'TRABAJADORES
        imgTra_Click Indice
    Case 1
        imgInci_Click Indice
    
    End Select
    
End Sub

Private Sub CargarFestivos()
Dim IT As ListItem
    ListView1.ListItems.Clear
    If txtCalen(0).Text <> "" Then
        Set miRsAux = New ADODB.Recordset
        Cad = "Select * from calendariof where idcal =" & txtCalen(0).Text
        If Me.chkTempoActual.Value = 1 Then
            Cad = Cad & " and fecha >= '" & Format(vEmpresa.FechaInicio, FormatoFecha)
            Cad = Cad & "' and fecha <= '" & Format(vEmpresa.FechaFin, FormatoFecha) & "'"
        End If
        Cad = Cad & " ORDER BY fecha"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Set IT = ListView1.ListItems.Add()
            IT.Text = Format(miRsAux!Fecha, "dd/mm/yyyy")
            IT.SubItems(1) = miRsAux!descripcion
            IT.Checked = True
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
End Sub


Private Function ImprimirTicajeActual() As Boolean
Dim SQL As String
Dim F As Date
Dim T As Long
Dim vHora As Integer


Dim PuedeQuitarParadas As Boolean
Dim Entrada As Boolean
Dim FueraIntervalo_ As Byte  'Sera 0 o 24, dependera
Dim vH As CHorarios
Dim Minutos As Integer
Dim HI As Date
Dim HF As Date
Dim HIAustada As Date
Dim difer As Currency
Dim Horas  As Currency
Dim Ajustadas As Currency

Dim QuitoMeriendaAlmuerzo As Currency
Dim QuitoMeriAlm As Byte '0 No he quitado nada     1. Ya he quitado almuerzo    2. Quito la merienda

    SQL = "Delete from tmpCombinada where codusu = " & vUsu.Codigo
    conn.Execute SQL
    Set vH = New CHorarios
    '''Sql = "Select entradafichajes.*,nomtrabajador from entradafichajes,trabajadores where entradafichajes.idtrabajador =trabajadores.idtrabajador "
    SQL = "select entradafichajes.idtrabajador,fecha,hour(hora) lahora,minute(hora) minutos,second(hora) segundos "
    SQL = SQL & ",Control from entradafichajes inner join trabajadores t on t.idtrabajador=entradafichajes.idtrabajador"
    Cad = ""
    If Me.txtFec(8).Text <> "" Then Cad = Cad & " AND fecha >='" & Format(txtFec(8).Text, FormatoFecha) & "'"
    If Me.txtFec(9).Text <> "" Then Cad = Cad & " AND fecha <='" & Format(txtFec(9).Text, FormatoFecha) & "'"
    If Me.txtTrab(10).Text <> "" Then Cad = Cad & " AND entradafichajes.idtrabajador >= " & txtTrab(10).Text
    If Me.txtTrab(11).Text <> "" Then Cad = Cad & " AND entradafichajes.idtrabajador <= " & txtTrab(11).Text
    
    'Abril 2014
    If Me.txtSecc(10).Text <> "" Then Cad = Cad & " AND t.seccion >= " & txtSecc(10).Text
    If Me.txtSecc(11).Text <> "" Then Cad = Cad & " AND t.seccion <= " & txtSecc(11).Text
    
    
    
    If Cad <> "" Then Cad = " WHERE " & Mid(Cad, 5)
    Cad = Cad & " ORDER BY fecha,idtrabajador,hora"
        
    Set miRsAux = New ADODB.Recordset
    SQL = SQL & Cad
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    F = CDate("01/01/1911")
    T = 0
    While Not miRsAux.EOF
        If miRsAux!Fecha <> F Then
            If T > 0 Then T = 9999999
        End If
        If T <> miRsAux!idTrabajador Then
            
            If T > 0 Then InsertaProcesarActual vHora, Horas, Ajustadas, QuitoMeriendaAlmuerzo, Entrada = True
                
            
            'Empezamos con el SQL
            F = miRsAux!Fecha
            T = miRsAux!idTrabajador
            
     
            vHora = 0
            QuitoMeriendaAlmuerzo = 0   'Currency de cuanto he quitado
            QuitoMeriAlm = 0  '0 no he quitado nada   1. El almuezro    2 La merienda
            Entrada = True
            Ajustadas = 0
            Horas = 0
            PuedeQuitarParadas = False
            
            'Si el trabjado no tiene el tipo de control 2 entonces NI miramos si quita paradas
            If vEmpresa.QueEmpresa = 2 Then
                If miRsAux!Control = 2 Then PuedeQuitarParadas = True
            End If
            
            If PuedeQuitarParadas Then
                'Veamos el horario para el trabajador, dia
                Cad = "calendariol.idcal=trabajadores.idcal and fecha=" & DBSet(F, "F") & " and idtrabajador"
                CadPa = "trabajadores.idcal"
                Cad = DevuelveDesdeBD("idhorario", "calendariol,trabajadores", Cad, CStr(T), "N", CadPa)
                If Val(Cad) = 0 Then Err.Raise 513, , "Error obteniendo horario trabajador: " & miRsAux!idTrabajador
                
                If Val(Cad) <> vH.IdHorario Then
                    If vH.Leer(CInt(Cad), F, CInt(CadPa)) = 1 Then Err.Raise 513, , "Error obteniendo horario nº: " & Cad
                End If
                
                'Si puede quitar paradas, y el horario lo tiene:
                Minutos = 0
                
                    If vH.Rectificar > 0 Then
                      If vH.Rectificar = vbRecESCuarto Then
                        Minutos = 15
                      Else
                          Minutos = 30   'Entradas salidas cada media hora
                      End If
                    End If
                 
                If vH.DtoMer = 0 And vH.DtoAlm = 0 Then PuedeQuitarParadas = False
  
                 
            End If
            vSQL = "INSERT INTO tmpCombinada(codusu,idTrabajador,Fecha,HT,HE,HR,idinci,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15,H16) VALUES (" & vUsu.Codigo & ","
            vSQL = vSQL & miRsAux!idTrabajador & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "',"
            Cad = ""
        End If
        
        
        If vHora < 16 Then   'solo ionserto 16
                
                
                
            If miRsAux!LaHora >= 0 And miRsAux!LaHora <= 23 Then
                i = miRsAux!LaHora
                FueraIntervalo_ = 0
            Else
                FueraIntervalo_ = 24
                If miRsAux!LaHora < 0 Then Stop  'De momento NO deberia entrar aqui
                i = miRsAux!LaHora - FueraIntervalo_
            End If
            
            CadPa = Format(i, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
            Cad = Cad & ",'" & CadPa & "'"
            
            
            
            
            If Not Entrada Then
                HF = Format(i, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                difer = DateDiff("n", HI, HF)
                If FueraIntervalo_ > 0 Then difer = difer + 1440
                
                Horas = Horas + difer
        
                'Ajustada
                If Minutos > 0 Then
                    HF = HoraRectificada(HF, vEmpresa.AjusteSalida, Minutos)
                    difer = DateDiff("n", HIAustada, HF)
                    If FueraIntervalo_ > 0 Then difer = difer + 1440
                End If
                Ajustadas = Ajustadas + difer
                    
            
            
            
            Else
                HI = Format(i, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                If Minutos > 0 Then
                    HIAustada = HoraRectificada(HI, vEmpresa.AjusteSalida, Minutos)
                Else
                    HIAustada = HI
                End If
              
            
            End If
            
            
 
                        
                        
            
            
            If PuedeQuitarParadas Then
                If vH.DtoAlm > 0 And FueraIntervalo_ = 0 Then
                    If QuitoMeriAlm = 0 Then
                        'Compruebo si el ticaje es menor que la hora del almuerzo
                        If HIAustada < vH.HoraDtoAlm Then
                            QuitoMeriAlm = 1
                            QuitoMeriendaAlmuerzo = vH.DtoAlm
                        End If
                    End If
                End If
                If vH.DtoMer > 0 Then
                    If QuitoMeriAlm < 2 Then
                        
                        
                        If HIAustada > vH.HoraDtoMer Then
                            QuitoMeriAlm = 2
                            QuitoMeriendaAlmuerzo = QuitoMeriendaAlmuerzo + vH.DtoMer
                        End If
                        
                    End If
                End If
            End If
            
            Entrada = Not Entrada
            

        End If
        vHora = vHora + 1
        miRsAux.MoveNext
   Wend
   miRsAux.Close
   Set miRsAux = Nothing
   Set vH = Nothing
   If T > 0 Then
        InsertaProcesarActual vHora, Horas, Ajustadas, QuitoMeriendaAlmuerzo, Entrada = True
        ImprimirTicajeActual = True
   Else
        ImprimirTicajeActual = False
   End If
End Function



Private Sub InsertaProcesarActual(NTicajes As Integer, HSinajustar As Currency, HAjustadas As Currency, Paradas As Currency, Correcto As Boolean)
Dim J As Integer

        'HT,HE,HR,idinci
        If Not Correcto Then
            HSinajustar = 0
            HAjustadas = 0
            Paradas = 0
        Else
            
            
            HSinajustar = Round(HSinajustar / 60, 2)
            
            HAjustadas = Round(HAjustadas / 60, 2)
            If Paradas <> 0 Then HAjustadas = HAjustadas - Paradas
            
        End If
        J = NTicajes
        While J < 16
            Cad = Cad & ",NULL"
            J = J + 1
        Wend
        
        Cad = DBSet(Paradas, "N") & "," & Abs(Correcto) & Cad
        
        Cad = DBSet(HSinajustar, "N") & "," & DBSet(HAjustadas, "N") & "," & Cad
        
    
        
        vSQL = vSQL & Cad & ")"
        EjecutaSQL vSQL
    
    
End Sub

Private Sub InsertaProcesarActualAuxiliar(NTicajes As Integer, HSinajustar As Currency, HAjustadas As Currency, Correcto As Boolean)
Dim J As Integer

        'HT,HE,HR,idinci
        If Not Correcto Then
            HSinajustar = 0
            HAjustadas = 0

        Else
            
            'Guardamos minutos
            'HSinajustar = Round(HSinajustar / 60, 2)
            
            'HAjustadas = Round(HAjustadas / 60, 2)

            
        End If
        J = NTicajes
        While J < 16
            Cad = Cad & ",NULL"
            J = J + 1
        Wend
        
        Cad = DBSet(0, "N") & "," & Abs(Correcto) & Cad
        
        Cad = DBSet(HSinajustar, "N") & "," & DBSet(HAjustadas, "N") & "," & Cad
        
    
        
        vSQL = vSQL & Cad & ")"
        EjecutaSQL vSQL
    
    
End Sub

'-------------------------------------------------------------------------------
'LISTADO COMBINADO
Private Sub HacerListadoCombinado()

    On Error GoTo EHacerListadoCombinado
    Screen.MousePointer = vbHourglass
    lblCombinado.Caption = "Comienzo proceso"
    lblCombinado.Refresh
    If CargaDatosCombinados Then
        
        espera 0.5
        i = 35
        If chkHorasCombinadas(0).Value = 1 Then i = 37
        'Ordenado por nombre
        If optTrab(1).Value Then i = i + 1
        
        
        
        'Parametros
        NumPa = 0
        CadPa = ""
        vSQL = ""
        If txtFec(10).Text <> "" Then vSQL = vSQL & "desde " & txtFec(10).Text & "  "
        If txtFec(11).Text <> "" Then vSQL = vSQL & "hasta " & txtFec(11).Text & "  "
        
        If txtSecc(2).Text <> "" Then vSQL = vSQL & "desde " & txtSecc(2).Text & "-" & Me.txtDSecc(2).Text & "  "
        If txtSecc(3).Text <> "" Then vSQL = vSQL & "hasta " & txtSecc(2).Text & "-" & Me.txtDSecc(3).Text & "  "
        vSQL = Trim(vSQL)
        CadPa = "FechaFin= """ & vSQL & """|"
        
        vSQL = ""
        If txtTrab(12).Text <> "" Then vSQL = vSQL & "desde " & txtTrab(12).Text & "-" & Me.txtDT(12).Text & "  "
        If txtTrab(13).Text <> "" Then vSQL = vSQL & "hasta " & txtTrab(13).Text & "-" & Me.txtDT(13).Text & "  "
        vSQL = Trim(vSQL)
        CadPa = CadPa & "FechaIni= """ & vSQL & """|"
        
        
        CadPa = CadPa & "EnDecimal= " & Abs(Me.chkHorasCombinadas(1).Value) & "|"
        
        NumPa = 3
        
        
        
        
        
        
        
        
        With frmImprimir
            .FormulaSeleccion = "{tmpcombinada.codusu} = " & vUsu.Codigo
            .OtrosParametros = CadPa
            .Opcion = i
            .NumeroParametros = NumPa
            .Show vbModal
        End With
    End If
    lblCombinado.Caption = ""
    Screen.MousePointer = vbDefault
    Exit Sub
EHacerListadoCombinado:
    MuestraError Err.Number, "Hacer Listado Combinado" & vbCrLf & Err.Description
End Sub

'Esta funcion modifica la tabla para mostrar el informe por lineas
Private Function CargaDatosCombinados() As Boolean
Dim RsBase As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim C As Integer
Dim CadenaSQL As String
Dim Cad As String
Dim Fecha As Date
Dim NH As Currency
Dim vH As CHorarios
Dim LeerHorario As Boolean
Dim Hora As Integer

On Error GoTo ErrSQL
CargaDatosCombinados = False



'--------------------------------------------------------------
'Montamos el sql
'Fechas
CadenaSQL = DesdeHastaSelect(0, 10, "Marcajes.fecha >= ", "AND")
CadenaSQL = CadenaSQL & DesdeHastaSelect(0, 11, "Marcajes.fecha <= ", "AND")
'D/H trabajador
CadenaSQL = CadenaSQL & DesdeHastaSelect(1, 12, "Marcajes.idtrabajador >= ", "AND")
CadenaSQL = CadenaSQL & DesdeHastaSelect(1, 13, "Marcajes.idtrabajador <= ", "AND")
'D/H Seccion
CadenaSQL = CadenaSQL & DesdeHastaSelect(2, 2, "Trabajadores.Seccion >= ", "AND")
CadenaSQL = CadenaSQL & DesdeHastaSelect(2, 3, "Trabajadores.Seccion <= ", "AND")


'Devolvemos la cadena
'Ahora recorremos los textos para hallar la subconsulta
Cad = "SELECT Marcajes.entrada,Marcajes.idTrabajador,Marcajes.Fecha,Marcajes.HorasTrabajadas,Marcajes.HorasIncid,ExcesoDefecto,Marcajes.idhorario,incfinal,Trabajadores.idcal"
Cad = Cad & " FROM Trabajadores,Marcajes,Incidencias,Secciones"
Cad = Cad & " WHERE  Trabajadores.IdTrabajador = Marcajes.idTrabajador"
Cad = Cad & " AND Trabajadores.Seccion = Secciones.Idseccion"
Cad = Cad & " AND Incidencias.idInci = Marcajes.IncFinal"
Cad = Cad & " AND Marcajes.correcto = 1"

'unimos la cadena sql
Cad = Cad & CadenaSQL

Cad = Cad & " ORDER BY idhorario,fecha,idcal"
Set vH = New CHorarios
vH.IdHorario = -1

Set RsBase = New ADODB.Recordset
lblCombinado.Caption = "Obteniendo conjunto registros"
lblCombinado.Refresh
RsBase.Open Cad, conn, , , adCmdText
If RsBase.EOF Then
    MsgBox "Ningun registro con esos valores", vbExclamation
    Set RsBase = Nothing
    Exit Function
End If
'Borramos los registros anteriores
conn.Execute "Delete  from tmpCombinada where codusu = " & vUsu.Codigo

'Empezamos para insertar
Set RT = New ADODB.Recordset
RT.CursorType = adOpenKeyset
RT.LockType = adLockOptimistic
RT.Open "Select * from tmpCombinada", conn, , , adCmdText

Set RS = New ADODB.Recordset
DoEvents
While Not RsBase.EOF

   

    lblCombinado.Caption = vH.IdHorario & " - " & RsBase!Fecha
    lblCombinado.Refresh
    'El horario
    If vH.IdHorario <> RsBase!IdHorario Then
        LeerHorario = True
    Else
        If vH.Fecha <> RsBase!Fecha Then
            LeerHorario = True
        Else
            'Veremos si el calendario del trabajador con el del horario que habia son el mismo
            If vH.idCal <> RsBase!idCal Then
                LeerHorario = False
                vH.CambioCalendario RsBase!idCal
            Else
                LeerHorario = False
            End If
        End If
    End If
    If LeerHorario Then vH.Leer CInt(RsBase!IdHorario), RsBase!Fecha, RsBase!idCal
    
    
    
    
'    If RsBase!Entrada = 105711 Then Stop
'    If RsBase!idTrabajador = 50049 Then Stop

    RT.AddNew
    Cad = "Select IdInci,hour(hora) LaHora,minute(hora) minutos,second(hora) segundos from EntradaMarcajes WHERE IdMarcaje=" & RsBase!Entrada
    Cad = Cad & " ORDER BY Hora"
    RS.Open Cad, conn, , , adCmdText
    RT!idTrabajador = RsBase!idTrabajador
    RT!Fecha = RsBase!Fecha
    RT!codusu = vUsu.Codigo
    'Trbajadas
    NH = CCur(RsBase!HorasTrabajadas)
    RT!HE = 0
    RT!hr = 0
    RT!IdInci = RsBase!IncFinal
    If chkHorasCombinadas(2).Value = 1 Then
        
        '-----------------
        'Ajuste calendario
        
        If Not vH.EsDiaFestivo Then
            'Si ha trabajado mas que las que pone el horario
            'Le pondre que tienes paradas. En otro caso 0
            
            If NH > vH.TotalHoras Then
                RT!HE = 0
                RT!hr = NH - vH.TotalHoras
                RT!IdInci = 0
            Else
                If RsBase!IncFinal > 0 Then
                    RT!hr = CCur(RsBase!HorasIncid)
                End If
            End If
        Else
            'EStamos ajustando y es festivo
            RT!HE = CCur(RsBase!HorasIncid)
        End If
    Else
        If RsBase!IncFinal = 0 Then
            RT!HE = 0
            RT!hr = 0
        Else
            If RsBase!ExcesoDefecto = 0 Then
                RT!hr = CCur(RsBase!HorasIncid)
            Else
                RT!HE = CCur(RsBase!HorasIncid)
            End If
        End If
    End If
    RT!HT = NH
    'Las horas
    
    
   
    C = 1
    While Not RS.EOF
        If RS!LaHora < 0 Then
            Hora = RS!LaHora + 24
        ElseIf RS!LaHora > 23 Then
            Hora = RS!LaHora - 24
        Else
            Hora = RS!LaHora
        End If
        Fecha = Format(Hora, "00") & ":" & Format(RS!Minutos, "00") & ":" & Format(RS!segundos, "00")
        If C < 9 Then
            Select Case C
            Case 1
                RT!H1 = Fecha
            Case 2
                RT!h2 = Fecha
            Case 3
                RT!H3 = Fecha
            Case 4
                RT!h4 = Fecha
            Case 5
                RT!h5 = Fecha
            Case 6
                RT!h6 = Fecha
            Case 7
                RT!h7 = Fecha
            Case 8
                RT!h8 = Fecha
            End Select
        Else
            Select Case C
            Case 9
                RT!H9 = Fecha
            Case 10
                RT!H10 = Fecha
            Case 11
                RT!H11 = Fecha
            Case 12
                RT!H12 = Fecha
            Case 13
                RT!H13 = Fecha
            Case 14
                RT!H14 = Fecha
            Case 15
                RT!H15 = Fecha
            Case 16
                RT!h16 = Fecha
            End Select
        End If
        RS.MoveNext
        C = C + 1
    Wend
    RT.Update
    RS.Close
    RsBase.MoveNext
Wend
RT.Close
RsBase.Close
Set RS = Nothing
Set RT = Nothing
Set RsBase = Nothing
CargaDatosCombinados = True
lblCombinado.Caption = ""
Exit Function
ErrSQL:
    MsgBox "Error: " & Err.Description, vbExclamation
End Function




'--------------------------------------------------------------
'------------------------------------------------------------
'
'   DIAS TRABAJADOS
'
'   Informe sobre los dias trabjados, horas totales, dias de vacaciones
'   Porcentajes etc
'
Private Function GenerarDiasTrabajados() As Boolean
Dim Trab As Long
Dim HT As Currency
Dim HExc As Currency
Dim HRet As Currency
Dim Dias As Integer

On Error GoTo EGenerarDiasTrabajados
    GenerarDiasTrabajados = False
    
    vSQL = ""
    'Desde hasta trabajador
    vSQL = vSQL & DesdeHastaSelect(1, 14, "marcajes.idtrabajador>=", " AND ")
    vSQL = vSQL & DesdeHastaSelect(1, 15, "marcajes.idtrabajador<=", " AND ")

    'D/H fecha
    vSQL = vSQL & DesdeHastaSelect(0, 12, "marcajes.Fecha>=", " AND ")
    vSQL = vSQL & DesdeHastaSelect(0, 13, "marcajes.Fecha<=", " AND ")


    'Seccion
    vSQL = vSQL & DesdeHastaSelect(2, 4, "trabajadores.seccion >=", " AND ")
    vSQL = vSQL & DesdeHastaSelect(2, 5, "trabajadores.seccion <=", " AND ")


    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select count(*) from marcajes,trabajadores,incidencias where marcajes.idtrabajador = trabajadores.idtrabajador"
    Cad = Cad & " and marcajes.incfinal =incidencias.idinci "
    Cad = Cad & vSQL
    Cad = Cad & " AND correcto = 0"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then Cad = ""
    End If
    miRsAux.Close
    
    'Muestro el mensaje de que hay marcajes incorrectos
    If Cad = "" Then
        MsgBox "Hay marcajes INCORRECTOS con esos parametros", vbExclamation
    End If
    
    '
    lblDiasTrabajados.Caption = "Obteniendo registros"
    lblDiasTrabajados.Refresh
    miSQL = "DELETE FROM tmpdatosmes where codusu =" & vUsu.Codigo
    conn.Execute miSQL
    miSQL = "DELETE FROM tmpDiasTrabajInci where codusu =" & vUsu.Codigo
    conn.Execute miSQL
    miSQL = "INSERT INTO tmpdatosmes (Mes, codusu,Trabajador, DiasTrabajados,  HorasT, HorasE,HorasN)"
    miSQL = miSQL & " VALUES (0," & vUsu.Codigo & ","
    
    Cad = "Select marcajes.idtrabajador,incfinal,excesodefecto,count(*) as dias,sum(horastrabajadas) as ht,sum(horasincid) as hi"
    Cad = Cad & " from marcajes,trabajadores,incidencias where marcajes.idtrabajador = trabajadores.idtrabajador and marcajes.incfinal =incidencias.idinci"
    Cad = Cad & vSQL
    Cad = Cad & " group by marcajes.idtrabajador,incfinal,excesodefecto order by marcajes.idtrabajador"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Trab = -1
    
    While Not miRsAux.EOF
        'Ahora
        If Trab <> miRsAux!idTrabajador Then
            If Trab > 0 Then
                'INSERTAMOS EN BD   DiasTrabajados,  HorasT, HorasE,HorasN
                Cad = miSQL & Trab & "," & Dias & "," & TransformaComasPuntos(CStr(HT))
                Cad = Cad & "," & TransformaComasPuntos(CStr(HExc)) & "," & TransformaComasPuntos(CStr(HRet)) & ")"
                conn.Execute Cad
            End If
            Trab = miRsAux!idTrabajador
            lblDiasTrabajados.Caption = Trab
            lblDiasTrabajados.Refresh
            Dias = 0
            HT = 0
            HExc = 0
            HRet = 0
        End If
            
        Dias = Dias + miRsAux!Dias
        HT = HT + miRsAux!HT
        If miRsAux!IncFinal = 0 Then
            'NO HAGO NADA
        Else
            If miRsAux!ExcesoDefecto = 1 Then
                'Horas extra
                HExc = HExc + miRsAux!HI
            Else
                HRet = HRet + miRsAux!HI
            End If
        End If
        
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'El ultimo
    If Trab > 0 Then
        'INSERTAMOS EN BD   DiasTrabajados,  HorasT, HorasE,HorasN
        Cad = miSQL & Trab & "," & Dias & "," & TransformaComasPuntos(CStr(HT))
        Cad = Cad & "," & TransformaComasPuntos(CStr(HExc)) & "," & TransformaComasPuntos(CStr(HRet)) & ")"
        conn.Execute Cad
    Else
        MsgBox "Ningun dato con esos parametros", vbExclamation
        Set miRsAux = Nothing
        Exit Function
    End If
    
    '-------------------------------------
    lblDiasTrabajados.Caption = "Incidencias generadas"
    lblDiasTrabajados.Refresh
                'NO TOCAR vsql
    Cad = "INSERT INTO tmpdiastrabajinci (idtrabajador, incidencia, horas,dias, codusu) "
    Cad = Cad & " select marcajes.idtrabajador,incidencia,sum(horas),count(*)," & vUsu.Codigo
    Cad = Cad & " from marcajes,trabajadores,incidencias,incidenciasgeneradas where"
    Cad = Cad & " marcajes.idtrabajador = trabajadores.idtrabajador and "
    Cad = Cad & " incidenciasgeneradas.entradamarcaje=marcajes.entrada  and"
    Cad = Cad & " incidencias.idinci = incidenciasgeneradas.Incidencia"
    
    Cad = Cad & vSQL
    Cad = Cad & " group by 1,2"
    conn.Execute Cad

    '---------------------------------------
    lblDiasTrabajados.Caption = "Vacaciones"
    lblDiasTrabajados.Refresh

    If vEmpresa.TodosLosDias Then
        'Llevo control de vacaciones
        '
        Cad = "select idtrabajador,count(*) from calendariot where tipodia=2"
        vSQL = DesdeHastaSelect(0, 12, "Fecha>=", " AND ")
        vSQL = vSQL & DesdeHastaSelect(0, 13, "Fecha<=", " AND ")
        Cad = Cad & vSQL
'        cad = cad & " and idtrabajador in (select idtrabajador from tmpdatosmes where codusu=" & vUsu.Codigo & ")"
        Cad = Cad & " group by idtrabajador"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = "UPDATE tmpDatosMes SET Anticipos=" & miRsAux.Fields(1) & " where codusu =" & vUsu.Codigo & " AND Trabajador =" & miRsAux.Fields(0)
            conn.Execute Cad
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        Cad = "Select idtrabajador,count(*) From marcajes where  incfinal= " & vEmpresa.IncVacaciones
        Cad = Cad & vSQL & " group by idtrabajador"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = "UPDATE tmpDatosMes SET Extras=" & miRsAux.Fields(1) & " where codusu =" & vUsu.Codigo & " AND Trabajador =" & miRsAux.Fields(0)
            conn.Execute Cad
            miRsAux.MoveNext
        Wend
        miRsAux.Close
            
        
    End If
    
    
    
    '---------------------------------
    Set miRsAux = Nothing
    lblDiasTrabajados.Caption = ""
    GenerarDiasTrabajados = True
    Exit Function
EGenerarDiasTrabajados:
    lblDiasTrabajados.Caption = ""
    MuestraError Err.Number, Err.Description
End Function



Private Sub GenerarImpresionimportesCostesDesdeMarcajes()
Dim F1 As Date
Dim F2 As Date
Dim SQL As String
Dim i As Long
Dim RS As Recordset
Dim Horas As Currency
Dim h2 As Currency

    On Error GoTo EGenerarImpresionimportesCostesDesdeMarcajes
        
    If txtFec(14).Text <> "" Then
        SQL = txtFec(14).Text
    Else
        SQL = Format("01/01/2003", "dd/mm/yyyy")
    End If
    F1 = CDate(SQL)
    
    If txtFec(15).Text <> "" Then
        SQL = txtFec(15).Text
    Else
        SQL = Format(Now, "dd/mm/yyyy")
    End If
    F2 = CDate(SQL)
        
    CadPa = ""
        
    If ComprobarMarcajesCorrectos(F1, F2, False) <> 0 Then
        CadPa = CadPa & vbCrLf & "Existen marcajes incorrectos entre las fechas"
        
    End If
    
        
    If CadPa <> "" Then
        SQL = CadPa & vbCrLf & "Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    CadPa = ""

    conn.Execute "DELETE FROm tmpMarcajes WHERE codusu =" & vUsu.Codigo
    


    
    SQL = "Select marcajes.*,excesodefecto,IRPFempresa from marcajes,incidencias,trabajadores where"
    SQL = SQL & " marcajes.idTrabajador = trabajadores.idTrabajador"
    SQL = SQL & " AND marcajes.incfinal = incidencias.idinci"
    SQL = SQL & " AND fecha >= '" & Format(F1, FormatoFecha)
    SQL = SQL & "'  AND fecha <='" & Format(F2, FormatoFecha) & "' "
    
    'Seccion
    If Me.txtSecc(6).Text <> "" Then SQL = SQL & " AND trabajadores.Seccion >= " & Me.txtSecc(6).Text
    If Me.txtSecc(7).Text <> "" Then SQL = SQL & " AND trabajadores.Seccion <= " & Me.txtSecc(7).Text
    
    
    'Trabajdo desde
    If Me.txtTrab(16).Text <> "" Then SQL = SQL & " AND marcajes.idTrabajador >=" & txtTrab(16).Text
    If Me.txtTrab(17).Text <> "" Then SQL = SQL & " AND marcajes.idTrabajador <=" & txtTrab(17).Text
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not RS.EOF
        SQL = "INSERT INTO tmpMarcajes(codusu,incFinal,entrada,idTrabajador,Fecha,HorasTrabajadas,HorasIncid) VALUES (" & vUsu.Codigo & ",0," & i & ","
        SQL = SQL & RS!idTrabajador & ",'"
        SQL = SQL & Format(RS!Fecha, FormatoFecha) & "',"
        If RS!ExcesoDefecto Then
            Horas = RS!HorasTrabajadas - RS!HorasIncid
            h2 = RS!HorasIncid
        Else
            Horas = RS!HorasTrabajadas
            h2 = 0
        End If
        SQL = SQL & TransformaComasPuntos(CStr(Horas)) & ","
        SQL = SQL & TransformaComasPuntos(CStr(h2)) & ")"
        conn.Execute SQL
        'Sig
        RS.MoveNext
        i = i + 1
    Wend
    RS.Close
    
    If i = 0 Then
        MsgBox "no se ha generado ningun dato con esos valores", vbExclamation
        Exit Sub
    End If
    
    
    
    'Mostramos el informe
     'Ponemos cadena
    SQL = ""
    
    If Me.txtFec(14).Text <> "" Then SQL = SQL & "   Desde : " & txtFec(14).Text
    If Me.txtFec(15).Text <> "" Then SQL = SQL & "   Hasta : " & txtFec(15).Text
    If txtTrab(16).Text <> "" Then SQL = SQL & "   Desde : " & txtTrab(16).Text
    If txtTrab(17).Text <> "" Then SQL = SQL & "   Hasta : " & txtTrab(17).Text
    If Me.txtSecc(6).Text <> "" Or Me.txtSecc(7).Text <> "" Then
        SQL = SQL & "   Seccion: "
        If Me.txtSecc(6).Text <> "" Then SQL = SQL & " desde " & Me.txtSecc(6).Text
        If Me.txtSecc(7).Text <> "" Then SQL = SQL & " hasta " & Me.txtSecc(7).Text
    End If
    SQL = Trim(SQL)
    If SQL <> "" Then SQL = "Intervalo= """ & SQL & """|"
    Me.Tag = SQL
    
    
    
    SQL = Me.Tag & SQL
    Me.Tag = ""
    
    If optTrab(5).Value Then
        i = 61
    Else
        i = 63
    End If
    
    If Me.optTrab(3).Value Then i = i + 1
        
    frmImprimir.Opcion = i
    frmImprimir.FormulaSeleccion = "{tmpMarcajes.codusu} = " & vUsu.Codigo
    frmImprimir.OtrosParametros = SQL
    frmImprimir.NumeroParametros = 2
    frmImprimir.Show vbModal
    
    Exit Sub
EGenerarImpresionimportesCostesDesdeMarcajes:
    MuestraError Err.Number, Err.Description
End Sub


'ALZIRA
'Septiembre 2014
' Los costes saldran ahora de la tabla jornadassemanalesalz
'Donde llevamos una entrada por cada trabajador,dia y tipo de hora (0. Normal   1-Estructural  2-Extra
'

Private Sub GenerarImpresionimportesCostesDesdejornadassemanalesalz()
Dim SQL As String
    If Not GenerarImpresionimportesCostesDesdejornadassemanalesAlzira Then Exit Sub

     'Mostramos el informe
     'Ponemos cadena
    SQL = ""
    
    If Me.txtFec(14).Text <> "" Then SQL = SQL & "   Desde : " & txtFec(14).Text
    If Me.txtFec(15).Text <> "" Then SQL = SQL & "   Hasta : " & txtFec(15).Text
    If txtTrab(16).Text <> "" Then SQL = SQL & "   Desde : " & txtTrab(16).Text
    If txtTrab(17).Text <> "" Then SQL = SQL & "   Hasta : " & txtTrab(17).Text
    If Me.txtSecc(6).Text <> "" Or Me.txtSecc(7).Text <> "" Then
        SQL = SQL & "   Seccion: "
        If Me.txtSecc(6).Text <> "" Then SQL = SQL & " desde " & Me.txtSecc(6).Text
        If Me.txtSecc(7).Text <> "" Then SQL = SQL & " hasta " & Me.txtSecc(7).Text
    End If
    SQL = Trim(SQL)
    If SQL <> "" Then SQL = "Intervalo= """ & SQL & """|"
    Me.Tag = SQL
    
    

    SQL = Me.Tag & SQL
    Me.Tag = ""
    
    If optTrab(5).Value Then
        i = 61
    Else
        i = 63
    End If
    
    If Me.optTrab(3).Value Then i = i + 1
        
    frmImprimir.Opcion = i
    frmImprimir.FormulaSeleccion = "{tmpMarcajes.codusu} = " & vUsu.Codigo
    frmImprimir.OtrosParametros = SQL
    frmImprimir.NumeroParametros = 2
    
    frmImprimir.Show vbModal
End Sub

Private Function GenerarImpresionimportesCostesDesdejornadassemanalesAlzira() As Boolean
Dim F1 As Date
Dim F2 As Date
Dim SQL As String
Dim i As Long
Dim RS As Recordset
Dim Horas As Currency
Dim h2 As Currency
Dim SeccionesNormales As String
Dim SeccionesAjustesHoras As String


    On Error GoTo EgenerarImpresionimportesCostesAlzira
    GenerarImpresionimportesCostesDesdejornadassemanalesAlzira = False
    If txtFec(14).Text <> "" Then
        SQL = txtFec(14).Text
    Else
        SQL = Format("01/01/2003", "dd/mm/yyyy")
    End If
    F1 = CDate(SQL)
    
    If txtFec(15).Text <> "" Then
        SQL = txtFec(15).Text
    Else
        SQL = Format(Now, "dd/mm/yyyy")
    End If
    F2 = CDate(SQL)
        
    CadPa = ""
        
    If ComprobarMarcajesCorrectos(F1, F2, False) <> 0 Then
        CadPa = CadPa & vbCrLf & "Existen marcajes incorrectos entre las fechas"
        
    End If
    
    
    
    'Veremos que secciones tienen un proceso normal de horas y cuales tienen un poroceso de ajuste de estructurales
    SeccionesNormales = ""
    SeccionesAjustesHoras = ""
    Set RS = New ADODB.Recordset
    SQL = " SELECT idseccion,Nominas FROM secciones WHERE 1=1 "
    If Me.txtSecc(6).Text <> "" Then SQL = SQL & " AND idSeccion >= " & Me.txtSecc(6).Text
    If Me.txtSecc(7).Text <> "" Then SQL = SQL & " AND idSeccion <= " & Me.txtSecc(7).Text
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        If RS!Nominas = 1 Then
            SeccionesAjustesHoras = SeccionesAjustesHoras & ", " & RS!IdSeccion
        Else
            SeccionesNormales = SeccionesNormales & ", " & RS!IdSeccion
        End If
        RS.MoveNext
    Wend
    RS.Close
    If SeccionesNormales <> "" Then SeccionesNormales = Trim(Mid(SeccionesNormales, 2))
    If SeccionesAjustesHoras <> "" Then SeccionesAjustesHoras = Trim(Mid(SeccionesAjustesHoras, 2))
        
        
        
    'Para las que ajustan horas, vere si todos los dias a procesar son los
    'Alzira, veremos si todos los dias entre el intervalo estan procesados
    If SeccionesAjustesHoras <> "" Then
    
            SQL = "Select count(distinct(fecha)) from marcajes,incidencias,trabajadores where"
            SQL = SQL & " marcajes.idTrabajador = trabajadores.idTrabajador"
            SQL = SQL & " AND marcajes.incfinal = incidencias.idinci AND "
            SQL = SQL & WHERE_CostesDiarios(F1, F2)
            SQL = SQL & " AND seccion IN (" & SeccionesAjustesHoras & ")"
            i = 0
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                i = DBLet(RS.Fields(0), "N")
            End If
            RS.Close
            'I= Dias en marcajes
        
            SQL = "Select count(distinct(fecha)) from jornadassemanalesalz,trabajadores"
            SQL = SQL & " where jornadassemanalesalz.idTrabajador = trabajadores.idTrabajador AND"
            SQL = SQL & WHERE_CostesDiarios(F1, F2)
            SQL = SQL & " AND seccion IN (" & SeccionesAjustesHoras & ")"
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = "0"
            If Not RS.EOF Then
                SQL = DBLet(RS.Fields(0), "N")
            End If
            RS.Close
            
            If Val(SQL) <> i Then
                'Dias en marcajes para los trabajadores Seccion: SeccionesAjustesHoras distinto a los procesados
                SQL = "Secc. ajuste horas. Dias procesados: " & SQL
                SQL = vbCrLf & "Dias marcajes: " & i & "    " & SQL
                CadPa = CadPa & SQL
            End If
     End If
        
        
        
        
        
        
        
        
        
    If CadPa <> "" Then
        Set RS = Nothing
        SQL = CadPa & vbCrLf & "Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    CadPa = ""
    Set RS = New ADODB.Recordset
    
    conn.Execute "DELETE FROm tmpMarcajes WHERE codusu =" & vUsu.Codigo
    
    'Las horas vendran, para los que ajustan nominas, desde dos sitios,
    'Desde para las secciones
    '  SeccionesAjustesHoras y SeccionesNormales
    i = 0
    If SeccionesNormales <> "" Then
        SQL = "Select marcajes.*,excesodefecto from marcajes,incidencias,trabajadores where"
        SQL = SQL & " marcajes.idTrabajador = trabajadores.idTrabajador"
        SQL = SQL & " AND marcajes.incfinal = incidencias.idinci AND "
        SQL = SQL & WHERE_CostesDiarios(F1, F2)
        SQL = SQL & " AND seccion IN (" & SeccionesNormales & ")"
        
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF
            SQL = "INSERT INTO tmpMarcajes(codusu,incFinal,entrada,idTrabajador,Fecha,HorasTrabajadas,HorasIncid,HorasExt) VALUES (" & vUsu.Codigo & ",0," & i & ","
            SQL = SQL & RS!idTrabajador & ",'"
            SQL = SQL & Format(RS!Fecha, FormatoFecha) & "',"
            If RS!ExcesoDefecto Then
                Horas = RS!HorasTrabajadas - RS!HorasIncid
                h2 = RS!HorasIncid
            Else
                Horas = RS!HorasTrabajadas
                h2 = 0
            End If
            SQL = SQL & TransformaComasPuntos(CStr(Horas)) & ","
            SQL = SQL & TransformaComasPuntos(CStr(h2)) & ",0)"
            conn.Execute SQL
            'Sig
            RS.MoveNext
            i = i + 1
        Wend
        RS.Close
        
    End If
        
        
        
    If SeccionesAjustesHoras <> "" Then

        SQL = "Select jornadassemanalesalz.idtrabajador,fecha,sum(if(tipohoras=0,horastrabajadas,0)) HNor,"
        SQL = SQL & " sum(if(tipohoras=1,horastrabajadas,0)) HEstr,"
        SQL = SQL & " sum(if(tipohoras=2,horastrabajadas,0)) HExtr "
        SQL = SQL & " from jornadassemanalesalz,trabajadores where jornadassemanalesalz.idTrabajador = trabajadores.idTrabajador AND "
        SQL = SQL & WHERE_CostesDiarios(F1, F2)
        SQL = SQL & " AND seccion IN (" & SeccionesAjustesHoras & ") GROUP BY 1,2 "
 

        
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF
            SQL = "INSERT INTO tmpMarcajes(codusu,incFinal,entrada,idTrabajador,Fecha,HorasTrabajadas,HorasIncid,HorasExt) VALUES (" & vUsu.Codigo & ",0," & i & ","
            SQL = SQL & RS!idTrabajador & ",'"
            SQL = SQL & Format(RS!Fecha, FormatoFecha) & "',"
            
            SQL = SQL & TransformaComasPuntos(CStr(RS!HNor)) & ","
            SQL = SQL & TransformaComasPuntos(CStr(RS!HEstr)) & ","
            SQL = SQL & TransformaComasPuntos(CStr(RS!HExtr)) & ")"
            conn.Execute SQL
            'Sig
            RS.MoveNext
            i = i + 1
        Wend
        RS.Close
        
    End If
    
    Set RS = Nothing
        
    If i = 0 Then
        MsgBox "no se ha generado ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    
    GenerarImpresionimportesCostesDesdejornadassemanalesAlzira = True
   
    
    Exit Function
EgenerarImpresionimportesCostesAlzira:
    MuestraError Err.Number, Err.Description
End Function


Private Function WHERE_CostesDiarios(F1 As Date, F2 As Date) As String
    WHERE_CostesDiarios = " fecha >= '" & Format(F1, FormatoFecha)
    WHERE_CostesDiarios = WHERE_CostesDiarios & "'  AND fecha <='" & Format(F2, FormatoFecha) & "' "
    
    'Seccion
    If Me.txtSecc(6).Text <> "" Then WHERE_CostesDiarios = WHERE_CostesDiarios & " AND trabajadores.Seccion >= " & Me.txtSecc(6).Text
    If Me.txtSecc(7).Text <> "" Then WHERE_CostesDiarios = WHERE_CostesDiarios & " AND trabajadores.Seccion <= " & Me.txtSecc(7).Text
    
    
    'Trabajdo desde
    If Me.txtTrab(16).Text <> "" Then WHERE_CostesDiarios = WHERE_CostesDiarios & " AND trabajadores.idTrabajador >=" & txtTrab(16).Text
    If Me.txtTrab(17).Text <> "" Then WHERE_CostesDiarios = WHERE_CostesDiarios & " AND trabajadores.idTrabajador <=" & txtTrab(17).Text
End Function


Private Function GeneraExcel() As Boolean
Dim Importe As Currency
Dim Acum As Currency
Dim Anti As Currency


    On Error GoTo eGeneraExcel
    GeneraExcel = False
    i = -1
    Cad = App.Path & "\tmpxls.csv"
    If Dir(Cad, vbArchive) <> "" Then Kill Cad
    
   
    
    CadPa = " SELECT nombre,`tmpMarcajes`.`idTrabajador`,nomtrabajador, fecha,`tmpMarcajes`.`HorasTrabajadas`,"
    CadPa = CadPa & " `tmpMarcajes`.`HorasIncid`,`tmpMarcajes`.`HorasExt`,  `Categorias`.`Importe1`, `Categorias`.`Importe2`,"
    CadPa = CadPa & " `Categorias`.`Importe3`, `Trabajadores`.`PorcAntiguedad`,  `Trabajadores`.`PorcIRPF`, `Trabajadores`.`PorcSS`,`Trabajadores`.`IRPFempresa`"
    CadPa = CadPa & "  FROM   ((`tmpmarcajes` `tmpMarcajes`"
    CadPa = CadPa & " INNER JOIN `trabajadores` `Trabajadores` ON `tmpMarcajes`.`idTrabajador`=`Trabajadores`.`IdTrabajador`)"
    CadPa = CadPa & " INNER JOIN `categorias` `Categorias` ON `Trabajadores`.`idCategoria`=`Categorias`.`IdCategoria`)"
    CadPa = CadPa & " INNER JOIN `secciones` `secciones` ON `Trabajadores`.`Seccion`=`secciones`.`IdSeccion`"
    CadPa = CadPa & " Where codusu = " & vUsu.Codigo
    CadPa = CadPa & " ORDER BY 1,"
    
    If optTrab(5).Value Then
        CadPa = CadPa & "fecha,`tmpMarcajes`.`idTrabajador`"
    Else
        CadPa = CadPa & "`tmpMarcajes`.`idTrabajador`, fecha"
    End If
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open CadPa, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
     i = FreeFile
    Open Cad For Output As #i
    
    
    'Primera linea, encabezados
    CadPa = "Seccion;"
    If optTrab(5).Value Then
        'Por fecha
        CadPa = CadPa & "Fecha;Trabajador;Nombre"
    Else
        'por trabajador
        CadPa = CadPa & "Trabajador;Nombre;Fecha"
    End If
    CadPa = CadPa & ";HT;Imp1;Anti1;Tot1;"
    CadPa = CadPa & "HEst;Imp2;Anti2;Tot2;"
    CadPa = CadPa & "HEXT;Imp3;Anti3;Tot3;"
    CadPa = CadPa & "Total H;Tot Imp;"
    CadPa = CadPa & "IRPF;SS;Liquido;SSEmpresa;Coste"
    Print #1, CadPa
    CadPa = ""
    
    'El csv llevara
    'Seccion
    '   Trabajador -->  Trabajador fecha
    '   Fecha      ..>  Fecha trabajador
    'HN, importe1* hn         Hest , importe2* hestr    , hext  importe3* hext   ,
    While Not miRsAux.EOF
        CadPa = EncomillarCampo(miRsAux!Nombre) & ";"
        
        If optTrab(5).Value Then
            'Por fecha
            CadPa = CadPa & Format(miRsAux!Fecha, "dd/mm/yyyy") & ";" & miRsAux!idTrabajador & ";" & EncomillarCampo(miRsAux!nomtrabajador)
        Else
            'por trabajador
            CadPa = CadPa & miRsAux!idTrabajador & ";" & EncomillarCampo(miRsAux!nomtrabajador) & ";" & Format(miRsAux!Fecha, "dd/mm/yyyy")
        End If
    
        'HN, importe
        Importe = Round(miRsAux!HorasTrabajadas * miRsAux!Importe1, 2)
        Anti = Round(Importe * DBLet(miRsAux!PorcAntiguedad, "N"), 2)
        CadPa = CadPa & ";" & miRsAux!HorasTrabajadas & ";" & CStr(Importe) & ";" & CStr(Anti) & ";" & CStr(Importe + Anti)
        Acum = Importe + Anti
        'Estructurales
        Importe = Round(miRsAux!HorasIncid * miRsAux!Importe2, 2)
        Anti = Round(Importe * DBLet(miRsAux!PorcAntiguedad, "N"), 2)
        CadPa = CadPa & ";" & miRsAux!HorasIncid & ";" & CStr(Importe) & ";" & CStr(Anti) & ";" & CStr(Importe + Anti)
        Acum = Acum + Importe + Anti
        'Extra
        Importe = Round(miRsAux!HorasExt * miRsAux!Importe3, 2)
        Anti = Round(Importe * DBLet(miRsAux!PorcAntiguedad, "N"), 2)
        CadPa = CadPa & ";" & miRsAux!HorasExt & ";" & CStr(Importe) & ";" & CStr(Anti) & ";" & CStr(Importe + Anti)
        Acum = Acum + Importe + Anti
        
        
        'Sumatorio de Horas e importe
        Anti = miRsAux!HorasTrabajadas + miRsAux!HorasIncid + miRsAux!HorasExt
        CadPa = CadPa & ";" & CStr(Anti) & ";" & CStr(Acum)
        'IRPF
        Anti = Round((Acum * DBLet(miRsAux!PorcIRPF, "N")) / 100, 2) 'lo guardo
        CadPa = CadPa & ";" & CStr(Anti)
        
        
        'SS
        Importe = Round((Acum * DBLet(miRsAux!PorcSS, "N") / 100), 2) 'lo guardo
        CadPa = CadPa & ";" & CStr(Anti)
        
        'Liquido empresa
        Importe = Acum - Anti - Importe
        CadPa = CadPa & ";" & CStr(Importe)
        
        
        'SS a cargo de la empresa
        Anti = Round((Acum * DBLet(miRsAux!IRPFempresa, "N")) / 100, 2)
        CadPa = CadPa & ";" & CStr(Anti)
        
        'Coste empresa
        'Total importe(antes del liquido) + coste IRPF empresa
        Importe = Acum + Anti
        CadPa = CadPa & ";" & CStr(Importe)
        
        Print #i, CadPa
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Close #i
    i = 0 'Para que no intenten volver a cerrarlo
    
    If CadPa = "" Then
        MsgBox "ningun dato generado", vbExclamation
    Else
        cd1.FileName = ""
        cd1.Filter = "*.csv|*.csv"
        cd1.DefaultExt = "csv"
        cd1.ShowSave
        If cd1.FileName <> "" Then
            CadPa = ""
            If Dir(cd1.FileName, vbArchive) <> "" Then
                If MsgBox("El fichero: " & cd1.FileName & " YA existe, ¿Sobreescribir?", vbQuestion + vbYesNoCancel) <> vbYes Then CadPa = "NO"
            End If
            If CadPa = "" Then
                FileCopy Cad, cd1.FileName
                MsgBox "Se ha generado correctamente el fichero: " & cd1.FileName, vbInformation
            End If
        End If
            
    End If
    
    
eGeneraExcel:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    If i > 0 Then Close #i


End Function



Private Function EncomillarCampo(Campo As String) As String
    EncomillarCampo = Replace(Campo, """", "''")
    EncomillarCampo = """" & EncomillarCampo & """"
End Function





'****************************************************************************
' Alzira. Relojes auxiliares
Private Function ImprimirTicajeActualRelojesAuxiliares() As Boolean
Dim SQL As String
Dim F As Date
Dim T As Long
Dim vHora As Integer



Dim Entrada As Boolean
Dim FueraIntervalo_ As Byte  'Sera 0 o 24, dependera
Dim vH As CHorarios
Dim Minutos As Integer
Dim HI As Date
Dim HF As Date
Dim HIAustada As Date
Dim difer As Currency
Dim Horas  As Currency
Dim Ajustadas As Currency
Dim Min As Integer
Dim Seg As Integer

    SQL = "Delete from tmpCombinada where codusu = " & vUsu.Codigo
    conn.Execute SQL
    Set vH = New CHorarios
    '''Sql = "Select entradafichajes.*,nomtrabajador from entradafichajes,trabajadores where entradafichajes.idtrabajador =trabajadores.idtrabajador "
    SQL = "select entradafichajauxliares.idtrabajador,fecha,hour(hora) lahora,minute(hora) minutos,second(hora) segundos "
    SQL = SQL & ",Control from entradafichajauxliares inner join trabajadores t on t.idtrabajador=entradafichajauxliares.idtrabajador"
    Cad = ""
    If Me.txtFec(18).Text <> "" Then Cad = Cad & " AND fecha >='" & Format(txtFec(18).Text, FormatoFecha) & "'"
    If Me.txtFec(19).Text <> "" Then Cad = Cad & " AND fecha <='" & Format(txtFec(19).Text, FormatoFecha) & "'"
    If Me.txtTrab(20).Text <> "" Then Cad = Cad & " AND entradafichajauxliares.idtrabajador >= " & txtTrab(20).Text
    If Me.txtTrab(21).Text <> "" Then Cad = Cad & " AND entradafichajauxliares.idtrabajador <= " & txtTrab(21).Text
   
    
    
    If Cad <> "" Then Cad = " WHERE " & Mid(Cad, 5)
    Cad = Cad & " ORDER BY fecha,idtrabajador,hora"
        
    Set miRsAux = New ADODB.Recordset
    SQL = SQL & Cad
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    F = CDate("01/01/1911")
    T = 0
    While Not miRsAux.EOF
        If miRsAux!Fecha <> F Then
            If T > 0 Then T = 9999999
        End If
        If T <> miRsAux!idTrabajador Then
            
            If T > 0 Then InsertaProcesarActualAuxiliar vHora, Horas, Ajustadas, Entrada = True
                
            
            'Empezamos con el SQL
            F = miRsAux!Fecha
            T = miRsAux!idTrabajador
            
     
            vHora = 0
           
            Entrada = True
            Ajustadas = 0
            Horas = 0
           
           
            vSQL = "INSERT INTO tmpCombinada(codusu,idTrabajador,Fecha,HT,HE,HR,idinci,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15,H16) VALUES (" & vUsu.Codigo & ","
            vSQL = vSQL & miRsAux!idTrabajador & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "',"
            Cad = ""
        End If
        
        
        If vHora < 16 Then   'solo ionserto 16
                
                
                
            If miRsAux!LaHora >= 0 And miRsAux!LaHora <= 23 Then
                i = miRsAux!LaHora
                FueraIntervalo_ = 0
            Else
                FueraIntervalo_ = 24
                If miRsAux!LaHora < 0 Then Stop  'De momento NO deberia entrar aqui
                i = miRsAux!LaHora - FueraIntervalo_
            End If
            
            CadPa = Format(i, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
            Cad = Cad & ",'" & CadPa & "'"
            
            
            
            
            If Not Entrada Then
                HF = Format(i, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                difer = DateDiff("s", HI, HF)
                If FueraIntervalo_ > 0 Then difer = difer + 86400  'segundos
                Min = difer \ 60
                Seg = difer - (CCur(Min) * 60)
                difer = Round((Seg / 60), 2) + Min
                'Lo paso a decimal
                             
                
                Horas = Horas + difer
        
                'Ajustada
                If Minutos > 0 Then
                    HF = HoraRectificada(HF, vEmpresa.AjusteSalida, Minutos)
                    difer = DateDiff("n", HIAustada, HF)
                    If FueraIntervalo_ > 0 Then difer = difer + 86400
                    
                       Min = difer \ 60
                        Seg = difer - (Min * 60)
                        difer = Round((Seg / 60), 2) + Min
                        'Lo paso a decimal
                            
                    
                    
                    
                End If
                Ajustadas = Ajustadas + difer
                    
            
            
            
            Else
                HI = Format(i, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                If Minutos > 0 Then
                    HIAustada = HoraRectificada(HI, vEmpresa.AjusteSalida, Minutos)
                Else
                    HIAustada = HI
                End If
              
            
            End If
            
            Entrada = Not Entrada
            

        End If
        vHora = vHora + 1
        miRsAux.MoveNext
   Wend
   miRsAux.Close
   Set miRsAux = Nothing
   Set vH = Nothing
   If T > 0 Then
        InsertaProcesarActualAuxiliar vHora, Horas, Ajustadas, Entrada = True
        ImprimirTicajeActualRelojesAuxiliares = True
   Else
        ImprimirTicajeActualRelojesAuxiliares = False
   End If
End Function


'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'
'       Generar datos para exportacion Nominas A3
'
'
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Private Function generarDatosNominas() As Boolean
Dim DiasTrabajadosPorMes As String  'Sera un vector que llevevara N: NO S: Si    F: Festivo
Dim VectorDiasTrab As String  'Sera la copia d este y rellenado
Dim DiasDelMes As Integer  '28,29,30,31
Dim DiasOficiales As Integer  'los que marque el calendairo
Dim diasTrabajados As Integer
Dim H As Currency
Dim Importe As Currency
Dim FF As Date
Dim vH As CHorarios
Dim DiasFestivos As String
Dim VectorDiasFestivos As String
Dim RS As ADODB.Recordset
Dim k As Integer
Dim J As Integer

    On Error GoTo egenerarDatosNominas
    generarDatosNominas = False
    Set miRsAux = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    'tmppagosmes   idTrabajador,Nombre,IRPF,SS,importe1,importe2
    Cad = "DELETE FROM tmppagosmes"
    conn.Execute Cad
    
    FF = CDate(txtFec(20).Text)
    'De momento cojo el IDCAL1
    'FALTA###
    DiasDelMes = DiasMes(Month(FF), Year(FF))
    
    H = CalculaHorasHorarioALZConVector(1, DiasOficiales, CDate("01" & Format(FF, "/mm/yyyy")), CDate(DiasDelMes & Format(FF, "/mm/yyyy")), DiasTrabajadosPorMes)
    
    
    Set vH = New CHorarios
    DiasFestivos = vH.LeerDiasFestivos(1, CDate("01" & Format(FF, "/mm/yyyy")), CDate(DiasDelMes & Format(FF, "/mm/yyyy")))
    'Añadimos los domingos
    For i = 1 To DiasDelMes
        If Format(CDate(i & "/" & Format(FF, "mm/yyyy")), "w") = 1 Then DiasFestivos = DiasFestivos & CDate(i & "/" & Format(FF, "mm/yyyy")) & "|"
    Next
    'Lo transformo en un array de enteros
    vSQL = DiasFestivos
    Cad = ""
    VectorDiasFestivos = ""
    While vSQL <> ""
        i = InStr(1, vSQL, "|")
        Cad = Mid(vSQL, 1, i - 1)
        vSQL = Mid(vSQL, i + 1)
        VectorDiasFestivos = VectorDiasFestivos & ", " & Day(CDate(Cad))
        
    Wend
    VectorDiasFestivos = Mid(VectorDiasFestivos, 2)
    
    Set vH = Nothing
    
    Cad = "select * from nominas where month(fecha)=" & Month(FF) & " and year(fecha)=" & Year(FF)
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    vSQL = ""
    While Not miRsAux.EOF
        i = i + 1
        VectorDiasTrab = CStr(DiasTrabajadosPorMes)  'Lo copio
       
        
       ' If miRsAux!idTrabajador = 20468 Then Stop
        
        
        'Veremos si ha trabajado algun dia festivo fesivos.. FESTIVO
        'Eso implicara que
        diasTrabajados = miRsAux!Dias
'        'Vere si ha trabajado algun dia festivo
'        cad = "select fecha from marcajes where day(fecha) IN (" & VectorDiasFestivos & ")  and  idtrabajador = " & miRsAux!idTrabajador
'        cad = cad & " AND month(fecha)=" & Month(FF) & " and year(fecha)=" & Year(FF)
'        RS.Open cad, conn, adOpenKeyset, adCmdText
'        If Not RS.EOF Then
'            'De momento no hacemos nada
'
'        End If
'
        'Calculamos dias
        'Vector: FNNNNFFFNNNNNFFNNNNNFFNNNNNFFNN   -> Son los dias del 1 al diasmes que son festivos. Iremos cambiando las N por S los dias que haya trabajado
        'Trabaja los dias oficiales.
        'No hacemos NADA, reemplazamoslas N de
        If diasTrabajados = DiasOficiales Then
            'Perfecto. Todas las N son S
            VectorDiasTrab = Replace(VectorDiasTrab, "N", "S")
            
        
        Else
            'Empieza la fiesta. Ha trabajado menos dias
            
            'Veremos cuales ha trabajado, SEGURO entre sin finde ni festivos
            Cad = "select fecha from marcajes where festivo=0 AND weekday(fecha)<5 and  idtrabajador = " & miRsAux!idTrabajador
            Cad = Cad & " AND month(fecha)=" & Month(FF) & " and year(fecha)=" & Year(FF)
            RS.Open Cad, conn, adOpenKeyset, adCmdText
            k = miRsAux!Dias
            While Not RS.EOF
                If InStr(1, DiasFestivos, Format(RS!Fecha, "dd/mm/yyyy")) = 0 Then
                    J = Day(RS!Fecha)
                    VectorDiasTrab = Mid(VectorDiasTrab, 1, J - 1) & "S" & Mid(VectorDiasTrab, J + 1)
                    k = k - 1
                Else
                    Stop
                End If
                
                RS.MoveNext
            Wend
            RS.Close
            
            'Comprobemos que no esta de baja
            If k > 0 Then
                For J = 1 To DiasDelMes
                    If Mid(VectorDiasTrab, J, 1) = "N" Then
                        'ESTE ES EL QUE COMPENSAMOS
                        VectorDiasTrab = Mid(VectorDiasTrab, 1, J - 1) & "S" & Mid(VectorDiasTrab, J + 1)
                        k = k - 1
                        If k = 0 Then Exit For
                    End If
                Next
            Else
               ' If diasTrabajados <> DiasOficiales Then Stop
            End If
            
            If k > 0 Then
                MsgBox "Mal,. NO ha compensado todos los dias", vbExclamation
                
            End If
        End If
        
        
        'DAtos para insertar en tmp
        '----------------------------------
        
        
        
        'Coopic.
        H = 0
        If miRsAux!hp > 0 Then H = miRsAux!hp * miRsAux!preciohe
        If miRsAux!HC > 0 Then H = H + miRsAux!HC * miRsAux!preciohc
        H = Round(H, 2)
        
        Importe = miRsAux!HN * miRsAux!preciohn
        Cad = ", (" & miRsAux!idTrabajador & ",'" & VectorDiasTrab & "'," & miRsAux!Dias & ",'',"
        Cad = Cad & DBSet(Importe, "N") & "," & DBSet(H, "N") & ")"
        vSQL = vSQL & Cad
        If (i Mod 10) = 0 Then
            vSQL = Mid(vSQL, 2)
            vSQL = "INSERT INTO tmppagosmes(idTrabajador,Nombre,IRPF,SS,importe1,importe2) VALUES " & vSQL
            conn.Execute vSQL
            DoEvents
            vSQL = ""
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If vSQL <> "" Then
        vSQL = Mid(vSQL, 2)
        vSQL = "INSERT INTO tmppagosmes(idTrabajador,Nombre,IRPF,SS,importe1,importe2) VALUES " & vSQL
        conn.Execute vSQL
    End If
    
    If i > 0 Then generarDatosNominas = True
    
egenerarDatosNominas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set RS = Nothing
End Function


