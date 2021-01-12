VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14880
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameActual 
      Height          =   6975
      Left            =   2520
      TabIndex        =   139
      Top             =   0
      Width           =   6495
      Begin VB.Frame FrameActual2 
         Caption         =   "Frame1"
         Height          =   495
         Left            =   1320
         TabIndex        =   374
         Top             =   5160
         Width           =   5055
         Begin VB.OptionButton optActual 
            Caption         =   "Area"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   377
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton optActual 
            Caption         =   "Trabajador"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   376
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optActual 
            Caption         =   "Seccion"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   375
            Top             =   120
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.ListBox List3 
         Height          =   735
         Left            =   960
         Style           =   1  'Checkbox
         TabIndex        =   371
         Top             =   3960
         Width           =   5055
      End
      Begin VB.CheckBox chkSinProcesar 
         Caption         =   "Adapta horario"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   339
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CheckBox chkSinProcesar 
         Caption         =   "Agrupa por seccion"
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   309
         Top             =   4800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   149
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2880
         TabIndex        =   280
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   148
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2880
         TabIndex        =   277
         Top             =   3120
         Width           =   3375
      End
      Begin VB.CheckBox chkSinProcesar 
         Caption         =   "Agrupa por trabajador"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   152
         Top             =   5880
         Width           =   2175
      End
      Begin VB.OptionButton optActual 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   151
         Top             =   4800
         Width           =   855
      End
      Begin VB.OptionButton optActual 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   150
         Top             =   4800
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   5040
         TabIndex        =   154
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdActual 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   153
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2880
         TabIndex        =   156
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2880
         TabIndex        =   155
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   147
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   146
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   9
         Left            =   4320
         ScrollBars      =   1  'Horizontal
         TabIndex        =   143
         Text            =   "Text1"
         Top             =   1035
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   8
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Agrupacion "
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
         Index           =   44
         Left            =   240
         TabIndex        =   373
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Area"
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
         Index           =   43
         Left            =   240
         TabIndex        =   372
         Top             =   3960
         Width           =   1335
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
         TabIndex        =   282
         Top             =   4800
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
         TabIndex        =   281
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
         TabIndex        =   279
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
         TabIndex        =   278
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
         TabIndex        =   159
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   29
         Left            =   960
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   145
         Top             =   240
         Width           =   5415
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   3960
         Picture         =   "frmListado.frx":6852
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   3360
         TabIndex        =   144
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
         TabIndex        =   142
         Top             =   840
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1560
         Picture         =   "frmListado.frx":68DD
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   27
         Left            =   960
         TabIndex        =   141
         Top             =   1080
         Width           =   465
      End
   End
   Begin VB.Frame FrameCostesTrabajador 
      Height          =   8415
      Left            =   120
      TabIndex        =   248
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ListBox List2 
         Height          =   735
         Left            =   960
         Style           =   1  'Checkbox
         TabIndex        =   370
         Top             =   3720
         Width           =   5055
      End
      Begin VB.OptionButton optHorasPorecesadas 
         Caption         =   "Asesoria"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   338
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CheckBox chkResumeTrabajador 
         Caption         =   "Resumen horas x traba(€)"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   337
         Top             =   7200
         Width           =   2655
      End
      Begin VB.CheckBox chkResumeTrabajador 
         Caption         =   "Por trabajador"
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   311
         Top             =   6840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkResumeTrabajador 
         Caption         =   "Resumen entrega "
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   310
         Top             =   6840
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2760
         TabIndex        =   275
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   9
         Left            =   1800
         TabIndex        =   254
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2760
         TabIndex        =   272
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   8
         Left            =   1800
         TabIndex        =   253
         Top             =   2760
         Width           =   855
      End
      Begin VB.OptionButton optHorasPorecesadas 
         Caption         =   "Trabajador - Horas"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   271
         Top             =   6240
         Width           =   2055
      End
      Begin VB.OptionButton optHorasPorecesadas 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   270
         Top             =   6240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox chkDesglosaDias 
         Caption         =   "Agrupa por empresa"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   255
         Top             =   5760
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   735
         ItemData        =   "frmListado.frx":6968
         Left            =   960
         List            =   "frmListado.frx":696A
         Style           =   1  'Checkbox
         TabIndex        =   268
         Top             =   4920
         Width           =   5055
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   19
         Left            =   1800
         TabIndex        =   252
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2760
         TabIndex        =   266
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox chkDesglosaDias 
         Caption         =   "Desglosa trabajador"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   256
         Top             =   5760
         Width           =   2055
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   18
         Left            =   1800
         TabIndex        =   251
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2760
         TabIndex        =   263
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   5040
         TabIndex        =   258
         Top             =   7680
         Width           =   1215
      End
      Begin VB.CommandButton cmdHorasProcesadas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   257
         Top             =   7680
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   17
         Left            =   4920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   250
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   16
         Left            =   1800
         ScrollBars      =   1  'Horizontal
         TabIndex        =   249
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Area"
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
         Index           =   42
         Left            =   120
         TabIndex        =   369
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   54
         Left            =   840
         TabIndex        =   276
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
         TabIndex        =   274
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   53
         Left            =   840
         TabIndex        =   273
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
         TabIndex        =   269
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   52
         Left            =   840
         TabIndex        =   267
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
         TabIndex        =   265
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   51
         Left            =   840
         TabIndex        =   264
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
         Picture         =   "frmListado.frx":696C
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   50
         Left            =   4080
         TabIndex        =   262
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
         TabIndex        =   261
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   840
         TabIndex        =   260
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
         TabIndex        =   259
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   16
         Left            =   1440
         Picture         =   "frmListado.frx":69F7
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame FrameTrabajadores 
      Height          =   5535
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.OptionButton optListTrab 
         Caption         =   "Seguridad social"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   56
         Top             =   3960
         Width           =   1575
      End
      Begin VB.ComboBox cboBaja 
         Height          =   315
         ItemData        =   "frmListado.frx":6A82
         Left            =   1440
         List            =   "frmListado.frx":6A8F
         Style           =   2  'Dropdown List
         TabIndex        =   340
         Top             =   3360
         Width           =   2895
      End
      Begin VB.CheckBox chkTarjeta 
         Caption         =   "Tarjeta"
         Height          =   255
         Left            =   1440
         TabIndex        =   60
         Top             =   4920
         Width           =   975
      End
      Begin VB.Frame FrameOrdenListTra 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   240
         TabIndex        =   101
         Top             =   4200
         Width           =   3495
         Begin VB.OptionButton optOrdenTraba 
            Caption         =   "Codigo"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   57
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optOrdenTraba 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   58
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ordenacion"
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
            Index           =   40
            Left            =   0
            TabIndex        =   367
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FrameTapaSecc 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   999
         Left            =   240
         TabIndex        =   100
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
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   99
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   98
         Top             =   2880
         Width           =   3375
      End
      Begin VB.CheckBox chkFoto 
         Caption         =   "Foto"
         Height          =   255
         Left            =   360
         TabIndex        =   59
         Top             =   4920
         Width           =   975
      End
      Begin VB.CheckBox chkSeccion 
         Caption         =   "Sección"
         Height          =   255
         Left            =   4920
         TabIndex        =   94
         Top             =   3360
         Width           =   975
      End
      Begin VB.OptionButton optListTrab 
         Caption         =   "Extendidos"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   55
         Top             =   3960
         Width           =   1575
      End
      Begin VB.OptionButton optListTrab 
         Caption         =   "Básicos"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   54
         Top             =   3960
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdTraba 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   46
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   8
         Left            =   5040
         TabIndex        =   49
         Top             =   4920
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
      Begin VB.Label Label1 
         Caption         =   "Tipo listado"
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
         Index           =   41
         Left            =   240
         TabIndex        =   368
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Situacion"
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
         Index           =   36
         Left            =   240
         TabIndex        =   341
         Top             =   3360
         Width           =   1335
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
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   20
         Left            =   1320
         TabIndex        =   97
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   19
         Left            =   1320
         TabIndex        =   96
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
         TabIndex        =   95
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
   Begin VB.Frame FrameDatosHco 
      Height          =   4095
      Left            =   5400
      TabIndex        =   349
      Top             =   1440
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton optOrdenTraba 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   366
         Top             =   3000
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optOrdenTraba 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   365
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdLisHcoMarcajes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   354
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   25
         Left            =   4200
         ScrollBars      =   1  'Horizontal
         TabIndex        =   353
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   24
         Left            =   1320
         ScrollBars      =   1  'Horizontal
         TabIndex        =   352
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   25
         Left            =   1320
         TabIndex        =   351
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   2160
         TabIndex        =   360
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   24
         Left            =   1320
         TabIndex        =   350
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   2160
         TabIndex        =   357
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   24
         Left            =   4320
         TabIndex        =   355
         Top             =   3600
         Width           =   1215
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
         Index           =   39
         Left            =   120
         TabIndex        =   364
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   25
         Left            =   3840
         Picture         =   "frmListado.frx":6AA4
         ToolTipText     =   "Buscar fecha"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   71
         Left            =   3360
         TabIndex        =   363
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   600
         TabIndex        =   362
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   24
         Left            =   1080
         Picture         =   "frmListado.frx":6B2F
         ToolTipText     =   "Buscar fecha"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   69
         Left            =   600
         TabIndex        =   361
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   25
         Left            =   1080
         Top             =   1560
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
         Index           =   38
         Left            =   120
         TabIndex        =   359
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   68
         Left            =   600
         TabIndex        =   358
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   24
         Left            =   1080
         Top             =   1200
         Width           =   255
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
         Index           =   16
         Left            =   360
         TabIndex        =   356
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame FrameTraspasoHCO 
      Height          =   3135
      Left            =   6600
      TabIndex        =   342
      Top             =   1800
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   23
         Left            =   3840
         TabIndex        =   348
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdTraspasar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2400
         TabIndex        =   347
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   23
         Left            =   2280
         ScrollBars      =   1  'Horizontal
         TabIndex        =   344
         Text            =   "Text1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Igual o anterior a la fecha introducida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   67
         Left            =   240
         TabIndex        =   346
         Top             =   1680
         Width           =   3240
      End
      Begin VB.Label Label1 
         Caption         =   "FECHA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   37
         Left            =   1080
         TabIndex        =   345
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   23
         Left            =   1800
         Picture         =   "frmListado.frx":6BBA
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Traspasar datos a histórico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   343
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame frHorascombinado 
      Height          =   6135
      Left            =   120
      TabIndex        =   160
      Top             =   0
      Width           =   6495
      Begin VB.ComboBox cboReloj 
         Height          =   315
         ItemData        =   "frmListado.frx":6C45
         Left            =   1560
         List            =   "frmListado.frx":6C52
         Style           =   2  'Dropdown List
         TabIndex        =   177
         Top             =   4080
         Width           =   1815
      End
      Begin VB.CheckBox chkHorasCombinadas 
         Caption         =   "Ajustar calendario"
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   214
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CheckBox chkHorasCombinadas 
         Caption         =   "Horas decimal"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   213
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CheckBox chkHorasCombinadas 
         Caption         =   "Agrupar por Fecha"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   212
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   182
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   181
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   176
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   175
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdHorasCombinadas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   180
         Top             =   5520
         Width           =   1215
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   179
         Top             =   5040
         Width           =   975
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   178
         Top             =   5040
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   170
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   12
         Left            =   1920
         TabIndex        =   169
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2880
         TabIndex        =   168
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2880
         TabIndex        =   167
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   11
         Left            =   4320
         ScrollBars      =   1  'Horizontal
         TabIndex        =   163
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   10
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   162
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   4920
         TabIndex        =   161
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Relojes"
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
         Index           =   35
         Left            =   240
         TabIndex        =   336
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblCombinado 
         Height          =   255
         Left            =   240
         TabIndex        =   215
         Top             =   5400
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
         TabIndex        =   185
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   36
         Left            =   960
         TabIndex        =   184
         Top             =   3120
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   35
         Left            =   960
         TabIndex        =   183
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
         TabIndex        =   174
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
         TabIndex        =   173
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   960
         TabIndex        =   172
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   960
         TabIndex        =   171
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
         TabIndex        =   166
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   3960
         Picture         =   "frmListado.frx":6C72
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
         TabIndex        =   165
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   31
         Left            =   3480
         TabIndex        =   164
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   10
         Left            =   1560
         Picture         =   "frmListado.frx":6CFD
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameA3 
      Height          =   3375
      Left            =   120
      TabIndex        =   300
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.ComboBox cboCentroTrabajo 
         Height          =   315
         ItemData        =   "frmListado.frx":6D88
         Left            =   1920
         List            =   "frmListado.frx":6D8A
         TabIndex        =   304
         Tag             =   "Centro trabajo trabajo|N|S|||trabajadores|idCentroA3|||"
         Text            =   "Combo1"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CheckBox chkA3 
         Caption         =   "Excel dias NO trabajados"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   303
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkA3 
         Caption         =   "Fichero integración"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   302
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton cmdGenNominaA3 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   305
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   4440
         TabIndex        =   306
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   20
         Left            =   2760
         ScrollBars      =   1  'Horizontal
         TabIndex        =   301
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblCentrotrabajo 
         Caption         =   "Centro trabajo"
         Height          =   255
         Left            =   600
         TabIndex        =   335
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   20
         Left            =   2400
         Picture         =   "frmListado.frx":6D8C
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
         TabIndex        =   308
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
         TabIndex        =   307
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame FrameListadoNominas 
      Height          =   5055
      Left            =   5640
      TabIndex        =   320
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton cmdNominas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   318
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   21
         Left            =   5040
         TabIndex        =   319
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   317
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2880
         TabIndex        =   333
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   12
         Left            =   1920
         TabIndex        =   316
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2880
         TabIndex        =   330
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   2880
         TabIndex        =   328
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   23
         Left            =   1920
         TabIndex        =   315
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   2880
         TabIndex        =   325
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   22
         Left            =   1920
         TabIndex        =   314
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   22
         Left            =   5040
         ScrollBars      =   1  'Horizontal
         TabIndex        =   313
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   21
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   312
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   13
         Left            =   1560
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   66
         Left            =   960
         TabIndex        =   334
         Top             =   3480
         Width           =   465
      End
      Begin VB.Image imgSecc 
         Height          =   255
         Index           =   12
         Left            =   1560
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   65
         Left            =   960
         TabIndex        =   332
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
         Index           =   34
         Left            =   240
         TabIndex        =   331
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   23
         Left            =   1560
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   64
         Left            =   960
         TabIndex        =   329
         Top             =   2400
         Width           =   465
      End
      Begin VB.Image imgTra 
         Height          =   255
         Index           =   22
         Left            =   1560
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   63
         Left            =   960
         TabIndex        =   327
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
         Index           =   33
         Left            =   240
         TabIndex        =   326
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   62
         Left            =   4200
         TabIndex        =   324
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   22
         Left            =   4680
         Picture         =   "frmListado.frx":6E17
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   21
         Left            =   1560
         Picture         =   "frmListado.frx":6EA2
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
         Index           =   32
         Left            =   240
         TabIndex        =   323
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   61
         Left            =   960
         TabIndex        =   322
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Listado nóminas generada"
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
         Index           =   14
         Left            =   1080
         TabIndex        =   321
         Top             =   240
         Width           =   5415
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
      TabIndex        =   219
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   247
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   226
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2880
         TabIndex        =   245
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   6
         Left            =   1920
         TabIndex        =   225
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   242
         Top             =   3000
         Width           =   3375
      End
      Begin VB.CommandButton cmdCostes 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   231
         Top             =   4560
         Width           =   1215
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   228
         Top             =   3960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   227
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2880
         TabIndex        =   240
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   17
         Left            =   1920
         TabIndex        =   224
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2880
         TabIndex        =   237
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   16
         Left            =   1920
         TabIndex        =   223
         Top             =   1920
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   3240
         TabIndex        =   236
         Top             =   3840
         Width           =   3015
         Begin VB.OptionButton optTrab 
            Caption         =   "Codigo"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   229
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTrab 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   230
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   14
         Left            =   2160
         ScrollBars      =   1  'Horizontal
         TabIndex        =   221
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   15
         Left            =   4560
         ScrollBars      =   1  'Horizontal
         TabIndex        =   222
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   5040
         TabIndex        =   232
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
         TabIndex        =   246
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
         TabIndex        =   244
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
         TabIndex        =   243
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
         TabIndex        =   241
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
         TabIndex        =   239
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
         TabIndex        =   238
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   14
         Left            =   1800
         Picture         =   "frmListado.frx":6F2D
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   44
         Left            =   3720
         TabIndex        =   235
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
         TabIndex        =   234
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   15
         Left            =   4200
         Picture         =   "frmListado.frx":6FB8
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   43
         Left            =   1200
         TabIndex        =   233
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
         TabIndex        =   220
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame frPresenciareal 
      Height          =   4695
      Left            =   120
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkPresReal 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1920
         TabIndex        =   211
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdPResenciaReal 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   71
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   72
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   5
         Left            =   4680
         ScrollBars      =   1  'Horizontal
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   1275
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   67
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
         TabIndex        =   74
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   73
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   70
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   6
         Left            =   1920
         TabIndex        =   69
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton optNomTra 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   66
         Top             =   3390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optNomTra 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   65
         Top             =   3390
         Width           =   975
      End
      Begin VB.Label Label2 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   82
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
         TabIndex        =   81
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   3840
         TabIndex        =   80
         Top             =   1320
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   4320
         Picture         =   "frmListado.frx":7043
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   79
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1680
         Picture         =   "frmListado.frx":70CE
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
         TabIndex        =   78
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   63
         Top             =   3390
         Width           =   975
      End
      Begin VB.OptionButton optNomTra 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   62
         Top             =   3390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Por fecha"
         Height          =   255
         Left            =   1920
         TabIndex        =   61
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
         Picture         =   "frmListado.frx":7159
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
         Picture         =   "frmListado.frx":71E4
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
      TabIndex        =   114
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CheckBox chkInci 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   218
         Top             =   4050
         Width           =   1095
      End
      Begin VB.CheckBox chkInci 
         Caption         =   "Mostrar detalle"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   217
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CheckBox chkInci 
         Caption         =   "Decimal"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   216
         Top             =   4440
         Width           =   1095
      End
      Begin VB.OptionButton optInci 
         Caption         =   "Nombre trab."
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   138
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton optInci 
         Caption         =   "Codigo trab."
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   137
         Top             =   4080
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdInci 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   136
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtDInci 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   134
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox txtInci 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   133
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtDInci 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   130
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtInci 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   129
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   121
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   120
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2880
         TabIndex        =   119
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2880
         TabIndex        =   118
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   7
         Left            =   5040
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   6
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   116
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   11
         Left            =   5040
         TabIndex        =   115
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
         TabIndex        =   135
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   128
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
         TabIndex        =   127
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   960
         TabIndex        =   126
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   125
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
         Picture         =   "frmListado.frx":726F
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   960
         TabIndex        =   124
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1560
         Picture         =   "frmListado.frx":72FA
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   4200
         TabIndex        =   123
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
         TabIndex        =   122
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame frDiasTrabajados 
      Height          =   4815
      Left            =   120
      TabIndex        =   186
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   13
         Left            =   5160
         ScrollBars      =   1  'Horizontal
         TabIndex        =   189
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   12
         Left            =   2040
         TabIndex        =   188
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.CommandButton cmdDiasTrabajados 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   195
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox ChkDiasTrab 
         Caption         =   "Agrupar por seccion"
         Height          =   255
         Left            =   1320
         TabIndex        =   194
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   14
         Left            =   5160
         TabIndex        =   196
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   203
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtDSecc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   202
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   193
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtSecc 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   192
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   15
         Left            =   2160
         TabIndex        =   191
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   3000
         TabIndex        =   200
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   14
         Left            =   2160
         TabIndex        =   190
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   3000
         TabIndex        =   187
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   13
         Left            =   4800
         Picture         =   "frmListado.frx":7385
         ToolTipText     =   "Buscar fecha"
         Top             =   1177
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   42
         Left            =   4320
         TabIndex        =   210
         Top             =   1200
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   12
         Left            =   1680
         Picture         =   "frmListado.frx":7410
         ToolTipText     =   "Buscar fecha"
         Top             =   1177
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   41
         Left            =   1080
         TabIndex        =   209
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
         TabIndex        =   208
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblDiasTrabajados 
         Height          =   255
         Left            =   240
         TabIndex        =   207
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
         TabIndex        =   206
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   40
         Left            =   1320
         TabIndex        =   205
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   39
         Left            =   1320
         TabIndex        =   204
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
         TabIndex        =   201
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
         TabIndex        =   199
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
         TabIndex        =   198
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   1320
         TabIndex        =   197
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
      TabIndex        =   83
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdGeneraInci 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   89
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   92
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtHoraD 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   88
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   87
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtInci 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   85
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDInci 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   84
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
         TabIndex        =   93
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hora Dec."
         Height          =   195
         Index           =   18
         Left            =   1800
         TabIndex        =   91
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hora"
         Height          =   195
         Index           =   17
         Left            =   360
         TabIndex        =   90
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Caption         =   "Incide."
         Height          =   195
         Index           =   16
         Left            =   360
         TabIndex        =   86
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
         Picture         =   "frmListado.frx":749B
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
         Picture         =   "frmListado.frx":7526
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
   End
   Begin VB.Frame FrCopiaHorario 
      Height          =   5295
      Left            =   120
      TabIndex        =   102
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdCopiaHorario 
         Caption         =   "Copiar"
         Height          =   375
         Left            =   5400
         TabIndex        =   113
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CheckBox chkTempoActual 
         Caption         =   "Temporada actual"
         Height          =   255
         Left            =   4800
         TabIndex        =   112
         Top             =   1560
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.TextBox txtCalendarioDestino 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox chkMas1año 
         Caption         =   "Incrementa 1 año"
         Height          =   255
         Left            =   4800
         TabIndex        =   109
         Top             =   1920
         Width           =   2535
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   240
         TabIndex        =   107
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
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtCalen 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Index           =   10
         Left            =   6720
         TabIndex        =   104
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
         Index           =   0
         Left            =   4800
         TabIndex        =   111
         Top             =   840
         Width           =   1695
      End
      Begin VB.Image imgcheckall 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmListado.frx":75B1
         ToolTipText     =   "Seleccionar todos"
         Top             =   4900
         Width           =   240
      End
      Begin VB.Image imgcheckall 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmListado.frx":76FB
         ToolTipText     =   "Quitar seleccion"
         Top             =   4900
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Calendario"
         Height          =   255
         Left            =   240
         TabIndex        =   108
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
         TabIndex        =   103
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameRelojesAuxiliares 
      Height          =   4575
      Left            =   120
      TabIndex        =   283
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkSinProcesar 
         Caption         =   "Agrupa por trabajador"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   289
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   21
         Left            =   1920
         TabIndex        =   288
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2880
         TabIndex        =   298
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtTrab 
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   287
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtDT 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2880
         TabIndex        =   295
         Top             =   2040
         Width           =   3375
      End
      Begin VB.CommandButton cmdRelojesAuxiliares 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   290
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   19
         Left            =   4680
         ScrollBars      =   1  'Horizontal
         TabIndex        =   286
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   18
         Left            =   1920
         ScrollBars      =   1  'Horizontal
         TabIndex        =   285
         Text            =   "Text1"
         Top             =   1155
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   4920
         TabIndex        =   291
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   60
         Left            =   960
         TabIndex        =   299
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
         TabIndex        =   297
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   960
         TabIndex        =   296
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
         TabIndex        =   294
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   19
         Left            =   4320
         Picture         =   "frmListado.frx":7845
         ToolTipText     =   "Buscar fecha"
         Top             =   1177
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   57
         Left            =   960
         TabIndex        =   293
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   18
         Left            =   1560
         Picture         =   "frmListado.frx":78D0
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
         TabIndex        =   292
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
         TabIndex        =   284
         Top             =   360
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
    
    '21  Listado nomminas
    '22  Horas por relojes
    '23  Marcajes a HCO
    '24  Impresion marcajes hco
    
Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Dim Cad As String
Dim I  As Integer
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


Private Sub cboCentroTrabajo_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub cboReloj_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

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
  '  Me.chkTarjeta.Value = IIf(Me.chkFoto.Value = 1, 0, 1)
End Sub

Private Sub ChkIncorr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerAccion
End Sub

Private Sub ChkIncorr_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub chkResumeTrabajador_Click(Index As Integer)
    If Index = 0 Then
        chkResumeTrabajador(1).Value = 0
        
    Else
        chkResumeTrabajador(0).Value = 0
    End If
End Sub

Private Sub chkSeccion_Click()
    Me.FrameTapaSecc.Visible = chkSeccion.Value = 0
End Sub

Private Sub chkSinProcesar_Click(Index As Integer)
    
    If Index = 3 And chkSinProcesar(3).Value = 1 Then
        chkSinProcesar(0).Value = 0
        chkSinProcesar(2).Value = 0
    End If
End Sub

Private Sub chkSinProcesar_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub chkTarjeta_Click()
    
   ' Me.chkFoto.Value = IIf(Me.chkTarjeta.Value = 1, 0, 1)
    
End Sub

Private Sub cmdActual_Click()
Dim B As Boolean
        
    Screen.MousePointer = vbHourglass
    
    If chkSinProcesar(3).Value = 1 Then
        chkSinProcesar(0).Value = 0
        chkSinProcesar(2).Value = 0
    End If
    
    
    
    
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
        
        Cad = ""
        If Me.optActual(4).Value Then
            For I = 0 To List3.ListCount - 1
                If List3.Selected(I) Then Cad = List3.List(I)
            Next
        End If
        If Cad <> "" Then
            Cad = "    AREA : " & Cad
            CadPa = CadPa & Cad
        End If
        
        
        
        
        If optActual(0).Value Then
            Cad = "pOrden= {tmpcombinada.idtrabajador}|"
        Else
            Cad = "pOrden= {trabajadores.nomtrabajador}|"
        End If
        
        Cad = Cad & "DesdeHasta= """ & CadPa & """|"
        
        CadPa = "0"
        If vEmpresa.QueEmpresa = 0 Then CadPa = "1"
        Cad = Cad & "EsTeinsa= " & CadPa & "|"
        
        Dim PorSeccion As Boolean
        
        PorSeccion = False
        If Me.optActual(2).Value Then PorSeccion = True
        
        'If chkSinProcesar(2).Value = 0 Then
        If Not PorSeccion Then
            If Me.chkSinProcesar(3).Value = 1 Then
                'Adapta horario
                frmImprimir.Opcion = 73
            Else
                If Me.chkSinProcesar(0).Value = 1 Then
                    frmImprimir.Opcion = 74
                    'por trabajador
                    If optActual(0).Value Then Cad = Cad & "AgrupadoPor= {tmpcombinada.idtrabajador}|"
                    
                Else
                    frmImprimir.Opcion = 67
                    
                End If
            End If
        Else
            'Agrupa seccion
            
            '     por trabajador
            If Me.chkSinProcesar(0).Value = 1 Then
                frmImprimir.Opcion = 60
            Else
                frmImprimir.Opcion = 32
            End If
        End If
        frmImprimir.FormulaSeleccion = "{tmpcombinada.codusu} = " & vUsu.Codigo
        frmImprimir.OtrosParametros = Cad
        frmImprimir.NumeroParametros = 3

        frmImprimir.Show vbModal
        Screen.MousePointer = vbDefault
    Else
        If Cad <> "" Then MsgBox "Ningún registro con esos valores", vbExclamation
    End If
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 3 Then CadenaDesdeOtroForm = ""   'Para que no refresque datos en el form de donde viene
    If Index = 12 Then CadenaDesdeOtroForm = ""   'Para que no refresque datos en el form de donde viene
    If Index = 23 Then CadenaDesdeOtroForm = ""   'Para que no refresque datos en el form de donde viene
    Unload Me
End Sub

Private Sub cmdCopiaHorario_Click()
Dim F As Date
    If txtCalen(0).Text = "" Then
        MsgBox "Seleccione calendario origen", vbExclamation
        Exit Sub
    End If
    
    
    Cad = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then Cad = Cad & "1"
    Next I
    
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
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
              F = CDate(ListView1.ListItems(I))
              If Me.chkMas1año.Value = 1 Then F = DateAdd("yyyy", 1, F)
              Cad = ListView1.ListItems(I).SubItems(1)
              NombreSQL Cad
              Cad = CadPa & "'" & Format(F, FormatoFecha) & "','" & Cad & "')"
              If Not ExisteElFestivo(F) Then EjecutaSQL Cad
        End If
    Next I
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
    
    
    
    
    If vEmpresa.QueEmpresa = 4 Then
        Cad = ""
        If Dir(App.Path & "\A3_Importes.exe", vbArchive) = "" Then Cad = Cad & vbCrLf & "-Fichero enlace importes A3"
        If Dir(App.Path & "\A3_inactivos.exe", vbArchive) = "" Then Cad = Cad & vbCrLf & "-Fichero enlace inactivos A3"
        If Cad <> "" Then
            MsgBox "Falta configurar. " & vbCrLf & Cad, vbExclamation
            Exit Sub
        End If
    ElseIf vEmpresa.CompensaHorasNominaMES Then
        Cad = ""
        If Dir(App.Path & "\A3_Importes.exe", vbArchive) = "" Then Cad = Cad & vbCrLf & "-Fichero enlace importes A3"
        If Dir(App.Path & "\A3_Generico.exe", vbArchive) = "" Then Cad = Cad & vbCrLf & "-Fichero enlace inactivos A3"
        If Cad <> "" Then
            MsgBox "Falta configurar. " & vbCrLf & Cad, vbExclamation
            Exit Sub
        End If
    End If
    
    If vEmpresa.TieneCentrosA3 Then
        If Me.cboCentroTrabajo.ListIndex < 0 Then
            MsgBox "Seleccione centro de trabajo ", vbExclamation
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    If generarDatosNominas Then
    
        If Me.chkA3(0).Value Then
            
            GeneraNominaA3 CDate(Me.txtFec(20).Text)
            
            'If vEmpresa.QueEmpresa = 4 Then
            If vEmpresa.QueEmpresa = 4 Then
                vSQL = App.Path & "\A3_Importes.exe"
                Lanza_EXE_Y_Espera vSQL
            
            Else
                If vEmpresa.CompensaHorasNominaMES Then
                    vSQL = App.Path & "\A3_Importes.exe"
                    Lanza_EXE_Y_Espera vSQL
                Else
                    If vEmpresa.QueEmpresa = 5 Then
                        'COOPIC lo copiamos en c:\Ariadna\enlaces
                        CopiarFicheroAEnlaces
                    End If
                End If
            End If
        End If
        If Me.chkA3(1).Value Then
            
        
            Screen.MousePointer = vbHourglass
            
            'Lanzamos el programa de EXCEL
            If vEmpresa.QueEmpresa = 4 Then
                vSQL = App.Path & "\A3_inactivos.exe"
                Lanza_EXE_Y_Espera vSQL
                
            Else
                If vEmpresa.CompensaHorasNominaMES Then
                    vSQL = App.Path & "\A3_Generico.exe"
                    Lanza_EXE_Y_Espera vSQL
                    
                Else
                    'COOPIC
                    If Dir(App.Path & "\gestoriaCoopic.exe", vbArchive) = "" Then
                        MsgBox "No existe programa enlace EXCEL", vbCritical
                    Else
                        vSQL = App.Path & "\gestoriaCoopic.exe"
                        Lanza_EXE_Y_Espera vSQL
                    End If
                End If
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
    
    Screen.MousePointer = vbHourglass
    If Opcion = 13 Then   'enero 2020   Ponia =14
        HacerListadoCombinado
    Else
        HacerListadoPorTipoReloj
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdHorasProcesadas_Click()
Dim F1 As Date
Dim Informe As String
Dim I As Integer
Dim info As Integer
Dim Aux As String

    NumPa = 0
    CadPa = ""
    vSQL = ""
    Cad = ""
    
    'Opcion asesoria. LAs fechas obligadas
    If Me.optHorasPorecesadas(2).Value Then
        If txtFec(16).Text = "" Or txtFec(17).Text = "" Then
            MsgBox "Fechas obligadas para el listado de datos asesoria", vbExclamation
            Exit Sub
        End If
    End If
        
        
    
    
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
    If Me.optHorasPorecesadas(1).Value Then
        If chkResumeTrabajador(0).Value = 1 Then Cad = ""
    End If
    
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
    
    
    'Label1(42):   Llevaremos el texto para elmensaje y el texto para el desdehasta
    Aux = ""
    Label1(42).Tag = ""
    info = 0
    For I = 0 To List2.ListCount - 1
        If List2.Selected(I) Then
            Aux = Aux & ", " & List2.ItemData(I)
            Label1(42).Tag = Label1(42).Tag & "- " & Mid(List2.List(I), 1, 15)
            info = info + 1
        End If
    Next
    If info = 0 Then
        MsgBox "Seleccione algun valor para " & Label1(42).Caption, vbExclamation
        Exit Sub
    End If
    
    
    If info < List2.ListCount Then
        'Significa que ha cogido menos
        'para el de/has rpt
        Label1(42).Tag = Mid(Label1(42).Tag, 2)
        Cad = Trim(Cad & "     " & Label1(42).Caption & ": " & Label1(42).Tag)
        Label1(42).Tag = ""
        
        'para el SQL
        Aux = Mid(Aux, 2)
        vSQL = vSQL & " AND {jornadassemanalesalz.codarea} in [ " & Aux & "]"
    End If
    
    
    
    
    CadPa = CadPa & "Intervalo= """ & Cad & """|"
    CadPa = CadPa & "DetallaTr= " & Abs(chkDesglosaDias(0).Value) & "|"
    NumPa = NumPa + 2
    
    
    info = -1
    If Me.optHorasPorecesadas(0).Value Then
        Cad = "AlzHorasProcesadasFecha"
        If Me.chkDesglosaDias(1).Value = 1 Then Cad = Cad & "Emp"
        Cad = Cad & ".rpt"
    
    ElseIf Me.optHorasPorecesadas(2).Value Then
        info = 2
        Cad = "resHorasTrabajaNominaAse.rpt"
    
    Else
        'Desglose trabajador
        If chkResumeTrabajador(2).Value = 1 Then
            Cad = "resHorasTrabajaNomina.rpt"
        Else
            If chkResumeTrabajador(1).Value = 1 Then
                Cad = "resHorasTrabaja.rpt"
            Else
                info = 1
                If vEmpresa.QueEmpresa = 4 Or vEmpresa.QueEmpresa = 5 Then
                    Cad = "picHorasTrabajador.rpt"
                Else
                    Cad = "AlzHorasTrabajador.rpt"
                End If
                
            End If
        End If
    End If
    
    If info > 0 Then
        'Informe personalizable
        Informe = DevuelveDesdeBD("informe", "scryst", "codigo", CStr(info))
        If Informe <> "" Then Cad = Informe
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
            I = 54
            
            
            'Detallar o no
            vSQL = 0
            If chkInci(1).Value = 1 Then vSQL = 1
            CadPa = CadPa & "Detallar= " & vSQL & "|"
            NumPa = NumPa + 1
        Else
            I = 33
        End If
        
        If Me.optInci(1).Value Then I = I + 1
    
    Else
        'Incidencia generada
        If Me.chkInci(2).Value = 1 Then
            'Agrupada por trabajador
            I = 52
        Else
            I = 50
        End If
        
        If Me.optInci(1).Value Then I = I + 1
        
        
        
        vSQL = 0
        If chkInci(1).Value = 1 Then vSQL = 1
        CadPa = CadPa & "Detallar= " & vSQL & "|"
        
        NumPa = NumPa + 1
    End If
    
    With frmImprimir
        .FormulaSeleccion = ""
        .OtrosParametros = CadPa
        .NumeroParametros = NumPa
        .Opcion = I
        .Show vbModal
    End With
        
End Sub

Private Sub cmdLisHcoMarcajes_Click()
    Dim B As Boolean
    Screen.MousePointer = vbHourglass
    B = ImprimirDatosHco
        
    Screen.MousePointer = vbDefault
    If B Then
        Cad = ""
        
   '     If txtSecc(10).Text <> "" Then Cad = Cad & "    desde " & txtSecc(10).Text & " - " & txtDSecc(10).Text
   '     If txtSecc(11).Text <> "" Then Cad = Cad & "    hasta " & txtSecc(11).Text & " - " & txtDSecc(11).Text
   '     If Cad <> "" Then Cad = "Sección: " & Trim(Cad)
        CadPa = ""
        If txtFec(24).Text <> "" Then CadPa = "Desde " & txtFec(24).Text
        If txtFec(25).Text <> "" Then CadPa = CadPa & "  Hasta " & txtFec(25).Text
        CadPa = Trim(CadPa & "   " & Cad)
        
        Cad = ""
        If txtTrab(24).Text <> "" Then Cad = "Desde " & txtTrab(24).Text & " - " & txtDT(24).Text
        If txtTrab(25).Text <> "" Then Cad = Cad & "    hasta " & txtTrab(25).Text & " - " & txtDT(25).Text
        Cad = Trim(Cad)
        If Cad <> "" Then
            If CadPa <> "" Then Cad = """ + chr(13) + """ & Cad
        End If
        CadPa = CadPa & Cad
        
        
        
        Cad = "pOrden= {tmppresencia.idtra}|"
        
        
        Cad = Cad & "DesdeHasta= """ & CadPa & """|"
        
        CadPa = "0"
        If vEmpresa.QueEmpresa = 0 Then CadPa = "1"
        Cad = Cad & "EsTeinsa= " & CadPa & "|"
        
        
      With frmImprimir
            .FormulaSeleccion = "{tmppresencia.codusu} = " & vUsu.Codigo
            
            If optOrdenTraba(3).Value Then
                .NombreRPT100 = "marcajeHCOF.rpt"
            Else
                .NombreRPT100 = "marcajeHCOT.rpt"
            End If
            .Titulo100 = "Historico marcajes"
            .OtrosParametros = Cad
            .Opcion = 100
            .NumeroParametros = 3
            .Show vbModal
      End With
    
   End If
    
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
        I = 2
    Else
        I = 0
    End If
    If Me.optNomTra(1).Value Then I = I + 1 'Por nombre en lugar de por codigo
    With frmImprimir
        .FormulaSeleccion = ""
        .OtrosParametros = CadPa
        .NumeroParametros = NumPa
        .Opcion = I
        .Show vbModal
    End With
    
    
End Sub

Private Sub cmdNominas_Click()
Dim dH As String
Dim Aux As String

    If Not Comprobarfechas(21, 22) Then Exit Sub
    dH = ""
    vSQL = ""
    Aux = ""
    If txtFec(21).Text <> "" Then
        Aux = Aux & "desde " & txtFec(21).Text
        vSQL = vSQL & " AND {nominas.fecha}>=date(" & Format(txtFec(21).Text, "yyyy,mm,dd)")
    End If
    If txtFec(22).Text <> "" Then
        Aux = Aux & " hasta " & txtFec(22).Text
        vSQL = vSQL & " AND {nominas.fecha}<=date(" & Format(txtFec(22).Text, "yyyy,mm,dd)")
    End If
    If Aux <> "" Then dH = "Fecha: " & Trim(Aux)
    
    Aux = ""
    If txtTrab(22).Text <> "" Then
        Aux = Aux & "desde " & txtTrab(22).Text & " " & Me.txtDT(22).Text
        vSQL = vSQL & " AND {nominas.idtrabajador}>=" & txtTrab(22).Text
    End If
    If txtTrab(23).Text <> "" Then
        Aux = Aux & " hasta " & txtTrab(23).Text & " " & Me.txtDT(23).Text
        vSQL = vSQL & " AND {nominas.idtrabajador}<=" & txtTrab(23).Text
    End If
    If Aux <> "" Then dH = Trim(dH & "      Trabajador: " & Trim(Aux))
    
    Aux = ""
    If txtSecc(12).Text <> "" Then
        Aux = Aux & "desde " & txtSecc(12).Text & " " & Me.txtDSecc(12).Text
        vSQL = vSQL & " AND {trabajadores.seccion}>=" & txtSecc(12).Text
    End If
    If txtSecc(13).Text <> "" Then
        Aux = Aux & " hasta " & txtSecc(13).Text & " " & Me.txtDSecc(13).Text
        vSQL = vSQL & " AND {trabajadores.seccion}<=" & txtSecc(13).Text
    End If
    If Aux <> "" Then dH = Trim(dH & "    Seccion: " & Trim(Aux))
    
    dH = "DH= """ & dH & """|"
    vSQL = Trim(vSQL)
    If UCase(Mid(vSQL, 1, 3)) = "AND" Then vSQL = Mid(vSQL, 4)
    
    
    I = 3
    If vEmpresa.CompensaHorasNominaMES Then I = 4
    Cad = DevuelveDesdeBD("informe", "scryst", "codigo", CStr(I))
    If Cad = "" Then
        MsgBox "Avisie soporte tecnico, Falta configurar datos en scryst"
        Cad = "rNomina.rpt"
    End If
    
    With frmImprimir
            .FormulaSeleccion = vSQL
            
            .NombreRPT100 = Cad
            .Titulo100 = "Listado nominas"
            .OtrosParametros = dH
            .Opcion = 100
            .NumeroParametros = 3
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
                I = 6
            Else
                'Por trabajador
                I = 8
            End If
            'Segundo orden. Codigo o nombre
            If optNomTra(3).Value Then I = I + 1
           
            With frmImprimir
                .FormulaSeleccion = "{tmppresencia.codusu} = " & vUsu.Codigo
                .OtrosParametros = CadPa
                .NumeroParametros = NumPa
                .Opcion = I
                .Show vbModal
            End With
    End If
  
    Screen.MousePointer = vbDefault
End Sub

Private Function HacerPresReal() As Boolean
Dim N As Long
Dim m As Long
Dim Inci As String
Dim Anyo As Integer
Dim Hora As String

    On Error GoTo EHacerPresReal
    HacerPresReal = False

    Cad = "Delete from tmppresencia where codusu =" & vUsu.Codigo
    conn.Execute Cad
    
    'HACemos el select
    Cad = "select entradamarcajes.* , nomtrabajador,nominci "
    Cad = Cad & " ,hour(horareal) lahora,minute(horareal) minutos,second(horareal) segundos "
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
    If Not miRsAux.EOF Then Anyo = Year(miRsAux!Fecha)
    While Not miRsAux.EOF
        If miRsAux!idmarcaje <> m Then
            If m <> 0 Then
                'INSERTAMOS EL MARCAJE
                InsertaHoraReal I, Inci
            End If
           ' Label2.Caption = miRsAux!idTrabajador & " " & Format(miRsAux!Fecha, "ddmmyy")
           ' Label2.Refresh
            Inci = ""
            N = N + 1
            Cad = miRsAux!nomtrabajador
            NombreSQL Cad
            Cad = N & "," & miRsAux!idTrabajador & ",'" & Cad & "','" & Format(miRsAux!Fecha, FormatoFecha) & "'"
            'Semana
            I = (Year(miRsAux!Fecha) - Anyo) * 100 + Format(miRsAux!Fecha, "ww", vbMonday)
            Cad = Cad & "," & I
            I = 0
            m = miRsAux!idmarcaje
        End If
        
        'Pongo la hora
        If miRsAux!IdInci <> 0 Then
            If Inci <> "" Then
                Inci = "Mas de una incidencia"
            Else
                
                Inci = miRsAux!NomInci
            End If
        End If
        
        I = I + 1
        If I <= 8 Then
        
            If miRsAux!LaHora > 23 Then
                Hora = miRsAux!LaHora - 24
            Else
                Hora = miRsAux!LaHora
            End If
            Hora = Format(Val(Hora), "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
        
            Cad = Cad & ",'" & Hora & "'"
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Esta pendiente el ultimo marcaje
    If m > 0 Then InsertaHoraReal I, Inci
    
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
    I = Me.ChkIncorr.Value * 2
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.chkCorrec.Value + I & "|"
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

    If chkFoto.Value = 1 And Me.chkTarjeta.Value = 1 Then
        MsgBox "Seleccione una o ninguna de las opciones ( FOTO - TARJETA )", vbExclamation
        Exit Sub
    End If
    
    
    If Me.optListTrab(2).Value Then
        'Listado extendido o SS no llevan foto ni tarjeta
        If chkFoto.Value = 1 Or Me.chkTarjeta.Value = 1 Then
            MsgBox "Listado datos Seguridad Social no llevan  FOTO - TARJETA ", vbExclamation
            Exit Sub
        End If
    End If
    
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
    
    
    If Me.cboBaja.ListIndex > 0 Then
         vSQL = vSQL & " AND "
        
        If Me.cboBaja.ListIndex = 1 Then
            d = d & "    ACTIVOS"
        Else
            d = d & "    DE BAJA"
            vSQL = vSQL & " NOT "
            formu = " NOT "
        End If
        vSQL = vSQL & " FecBaja is  null"
        formu = formu & " isnull({trabajadores.FecBaja})"
    End If
    
    If chkSeccion.Value = 1 Then
        If txtSecc(0).Text <> "" Then
            If formu <> "" Then formu = formu & " AND "
            d = d & "   Desde " & txtSecc(0).Text & " " & txtDSecc(0).Text
            formu = formu & "{secciones.idseccion}>=" & txtSecc(0).Text
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
    
    I = 0
    I = I + (Me.chkFoto.Value * 4)
    If Me.chkSeccion.Value = 1 Then
        I = I + (Me.chkSeccion.Value * 8)
    Else
        I = I + (Me.chkSeccion.Value * 8)
       '''''''' If optListTrab(1).Value Then i = i + 2  'Solo para los que no hay seccion salen extendidos o basicos
    End If
    If optListTrab(1).Value Then I = I + 2
    If optOrdenTraba(1).Value Then I = I + 1
    
  
    
    d = "SELECT count(*) from trabajadores"
    If vSQL <> "" Then d = d & " WHERE " & Mid(vSQL, 6)
    If Not TieneDatos(d) Then
        MsgBox "Ningun registro con esos valores", vbExclamation
        Exit Sub
    End If
    
    I = I + 10 'Pq empiezan los listado en el 10
    
    If optListTrab(2).Value Then
        'Datos Seguridad SOCIAL
        If Me.chkSeccion.Value = 1 Then
            d = "trabsecsipsin"
        Else
            d = "trabsipsin"
        End If
        
        d = d & IIf(optOrdenTraba(1).Value, "nom", "cod") & ".rpt"
        
    End If
    
    
    With frmImprimir
        If Me.chkTarjeta.Value = 1 Then
            I = 100
            .NombreRPT100 = "trabtarjeta.rpt"
            .Titulo100 = "Tarjetas"
            .ConSubreport100 = False
            
            
        Else
            If optListTrab(2).Value Then
                I = 100
                .NombreRPT100 = d
                .Titulo100 = "Datos seguridad social"
                .ConSubreport100 = False
            End If
        End If
        
        
        .FormulaSeleccion = formu
        .OtrosParametros = CadPa
        .NumeroParametros = NumPa
        .Opcion = I
        .Show vbModal
    End With

        
    
    
End Sub



Private Sub Command1_Click()
    
End Sub

Private Sub cmdTraspasar_Click()
Dim F As Date
    txtFec(23).Text = Trim(txtFec(23).Text)

    If Me.txtFec(23).Text = "" Then Exit Sub
    
    If txtFec(23).Text >= CDate(vEmpresa.FechaInicio) Then
        MsgBox "Fecha debe ser menor a " & vEmpresa.FechaInicio, vbExclamation
        Exit Sub
    End If
    
    
    
    Cad = DevuelveDesdeBD("min(fecha)", "entradamarcajes", "1", "1")
    If Cad = "" Then
        MsgBox "Ningun dato a traspasar", vbExclamation
        Exit Sub
    End If
    F = CDate(Cad)
    
    
    
    Cad = DevuelveDesdeBD("MAX(fecha)", "entradamarcajeshco", "1", "1")
    If Cad <> "" Then
        If CDate(Cad) > F Then
            Cad = vbCrLf & vbCrLf & "max fecha: " & Cad & "   Min traspaso: " & F
            MsgBox "Ya hay datos entre esas fechas. Consulte soporte tecnico" & Cad, vbExclamation
            Exit Sub
        End If
    End If
        
    
    
    
    
    
    
    
    Cad = "Traspasar a histórico fecha menor o igual a " & txtFec(23).Text
    Cad = InputBox(Cad, "Contraseña")
    Cad = Trim(Cad)
    If Cad = "" Then Exit Sub
    
    If UCase(Cad) <> "ARIADNA" Then
        MsgBox "Password incorrecto", vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    
    CadenaDesdeOtroForm = txtFec(23).Text
    Unload Me
    
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
    FrameListadoNominas.Visible = False
    FrameTraspasoHCO.Visible = False
    
    FrameDatosHco.Visible = False
    
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
        
        If CadenaDesdeOtroForm <> "" Then
            Cad = DevuelveDesdeBD("max(fecha)", "marcajeshco", "1", "1")
            I = 1
            
        Else
            Cad = ""
            I = 0
        End If
        If Cad = "" Then Cad = DateAdd("d", -1, Now)
        txtFec(1).Text = Format(Cad, "dd/mm/yyyy")
        If I = 1 Then txtFec(0).Text = txtFec(1).Text
        CadenaDesdeOtroForm = ""
        
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
        txtFec(4).Text = "01/" & Format(DateAdd("d", -1, Now), "mm/yyyy")
    
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
        
        'IListado trabajadores
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
        '
        
        imgTra(10).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        imgTra(11).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        FrameActual.Top = 0
        FrameActual.Left = 120
        FrameActual.Visible = True
        FrameActual2.BorderStyle = 0
        H = Me.FrameActual.Height
        W = Me.FrameActual.Width
        
        
        Cad = "select descripcion descripcion ,codarea id from areas order by 1"
        CargaListBoxDesdeTabla Cad, List3, False, True
        optActual(4).Visible = List3.ListCount > 1
        
        txtFec(8).Text = CadenaDesdeOtroForm
        txtFec(9).Text = CadenaDesdeOtroForm
        If CadenaDesdeOtroForm <> "" Then PonerFocoBtn Me.cmdActual
        CadenaDesdeOtroForm = ""
        Caption = "Actual"
        
        chkSinProcesar(2).Value = IIf(vEmpresa.QueEmpresa = 2, 1, 0)
        
    Case 13, 22
    
        If Opcion = 13 Then
            lblTitulo(8).Caption = "Listado horas combinado"
        Else
            lblTitulo(8).Caption = "Listado horas x reloj"
            cboReloj.ListIndex = 0
        End If
    
    
        Caption = "Horas combinado"
        Me.frHorascombinado.Visible = True
        For H = 2 To 3
            imgTra(10 + H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
            imgSecc(H).Picture = frmPpal.imgListImages16.ListImages(3).Picture
        Next
        frHorascombinado.Visible = True
        H = Me.frHorascombinado.Height
        W = Me.frHorascombinado.Width
        txtFec(11).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
        txtFec(10).Text = "01/" & Format(DateAdd("d", -1, Now), "mm/yyyy")
        
        
        chkHorasCombinadas(2).Visible = Opcion <> 22
        chkHorasCombinadas(1).Visible = Opcion <> 22
        Label1(35).Visible = Opcion = 22
        cboReloj.Visible = Opcion = 22

        
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
        If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = Format(DateAdd("d", -7, Now), "dd/mm/yyyy")
            
        H = Me.FrameCostesTrabajador.Height
        W = FrameCostesTrabajador.Width
            
            
        txtFec(16).Text = CadenaDesdeOtroForm
        
        Caption = "Listado"
        
        Cad = "select NomSubEmpre descripcion ,idSubEmr id from areasubempresa order by 1"
        CargaListBoxDesdeTabla Cad, List1, False, True
        
        Cad = "select descripcion descripcion ,codarea id from areas order by 1"
        CargaListBoxDesdeTabla Cad, List2, False, True
        
        CadenaDesdeOtroForm = ""
        optHorasPorecesadas_Click 0
    
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

        Me.lblCentrotrabajo.Visible = vEmpresa.TieneCentrosA3
        Me.cboCentroTrabajo.Visible = vEmpresa.TieneCentrosA3
        CargaCentroTrabajo
    Case 21
        FrameListadoNominas.Top = 60
        FrameListadoNominas.Left = 90
        FrameListadoNominas.Visible = True
        H = Me.FrameListadoNominas.Height
        W = FrameListadoNominas.Width
        vSQL = DevuelveDesdeBD("max(fecha)", "nominas", "1", "1")
        If vSQL = "" Then vSQL = Now
        Caption = "Exportación"
        txtFec(21).Text = Format(vSQL, "dd/mm/yyyy")
        txtFec(22).Text = txtFec(21).Text
    
    Case 23
        
        FrameTraspasoHCO.Top = 60
        FrameTraspasoHCO.Left = 90
        FrameTraspasoHCO.Visible = True
        H = Me.FrameTraspasoHCO.Height
        W = FrameTraspasoHCO.Width
    
    Case 24
        
        FrameDatosHco.Top = 60
        FrameDatosHco.Left = 90
        FrameDatosHco.Visible = True
        H = Me.FrameDatosHco.Height
        W = FrameDatosHco.Width
        
    End Select
    
    Me.Height = H + 500
    Me.Width = W + 300
    
    If Opcion = 15 Then
        I = 11  'COmo es lo mismo para la opcion 15 que la 11..
    ElseIf Opcion = 19 Then
        I = 18
    ElseIf Opcion = 22 Then
        I = 13
    Else
        I = Opcion
    End If
    Me.cmdCancelar(I).Cancel = True



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
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = (Index = 0)
    Next I
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim Obj As Object

    Set frmc = New frmCal
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    
    Set Obj = imgFec(Index).Container
    
    While imgFec(Index).Parent.Name <> Obj.Name
        esq = esq + Obj.Left
        dalt = dalt + Obj.Top
        Set Obj = Obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
    
    
    frmc.Left = esq + imgFec(Index).Parent.Left + 30
    frmc.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtFec(Index).Text <> "" Then frmc.NovaData = txtFec(Index).Text
    ' ********************************************

    frmc.Show vbModal
    Set frmc = Nothing
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
    
    chkResumeTrabajador(0).Visible = Index = 1
    chkResumeTrabajador(1).Visible = Index = 1
    chkResumeTrabajador(2).Visible = Index = 1
    'If Index = 1 Then
    '    chkResumeTrabajador(0).Value = 1
        
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
Dim B As Boolean
    
    
    If Me.chkSinProcesar(3).Value = 1 Then
        B = ImprimirTicajeActualAdaptado
    Else
        B = ImprimirTicajeActualNormal
    End If
    
    ImprimirTicajeActual = B
End Function

Private Function ImprimirTicajeActualNormal() As Boolean
Dim SQL As String
Dim F As Date
Dim T As Long
Dim vHora As Integer


Dim PuedeQuitarParadas As Boolean
Dim Entrada As Boolean
Dim FueraIntervalo2 As Integer     '0 NORMAL    1. Hora negativa     2. Hora mayor=24
Dim vH As CHorarios
Dim Minutos As Integer
Dim HI As Date
Dim HF As Date
Dim HIAustada As Date
Dim Difer As Currency
Dim Horas  As Currency
Dim Ajustadas As Currency

Dim QuitoMeriendaAlmuerzo As Currency
Dim QuitoMeriAlm As Byte '0 No he quitado nada     1. Ya he quitado almuerzo    2. Quito la merienda
Dim AuxArea As String



    On Error GoTo eImprimirTicajeActual
    ImprimirTicajeActualNormal = False

    SQL = "Delete from tmpCombinada where codusu = " & vUsu.Codigo
    conn.Execute SQL
    Set vH = New CHorarios
    '''Sql = "Select entradafichajes.*,nomtrabajador from entradafichajes,trabajadores where entradafichajes.idtrabajador =trabajadores.idtrabajador "
    SQL = "select entradafichajes.idtrabajador,fecha,hour(hora) lahora,minute(hora) minutos,second(hora) segundos ,Control, "
    If vEmpresa.AcabalgadoDiaInicio Then
        SQL = SQL & " 0 "   'Acabalgdo dia inicio no hay hpras negativas
    Else
        SQL = SQL & "if(hora<0,1,0) "
    End If
    SQL = SQL & " HorasNegativas , coalesce(area,0)"
    SQL = SQL & " from entradafichajes inner join trabajadores t on t.idtrabajador=entradafichajes.idtrabajador"
    SQL = SQL & " LEFT JOIN terminales ON entradafichajes.reloj = terminales.id "
    Cad = ""
    If Me.txtFec(8).Text <> "" Then Cad = Cad & " AND fecha >='" & Format(txtFec(8).Text, FormatoFecha) & "'"
    If Me.txtFec(9).Text <> "" Then Cad = Cad & " AND fecha <='" & Format(txtFec(9).Text, FormatoFecha) & "'"
    If Me.txtTrab(10).Text <> "" Then Cad = Cad & " AND entradafichajes.idtrabajador >= " & txtTrab(10).Text
    If Me.txtTrab(11).Text <> "" Then Cad = Cad & " AND entradafichajes.idtrabajador <= " & txtTrab(11).Text
    
    'Abril 2014
    If Me.txtSecc(10).Text <> "" Then Cad = Cad & " AND t.seccion >= " & txtSecc(10).Text
    If Me.txtSecc(11).Text <> "" Then Cad = Cad & " AND t.seccion <= " & txtSecc(11).Text
    
    
    'Noviembre 2020
    'AREAS-ALMACENES
    AuxArea = ""
    NumRegElim = 0
    For T = 0 To List3.ListCount - 1
        If List3.Selected(T) Then
            NumRegElim = NumRegElim + 1
            AuxArea = AuxArea & ", " & List3.ItemData(T)
        End If
    Next
    If NumRegElim = 0 Then
        MsgBox "Seleccione un AREA-ALMACEN", vbExclamation
        Cad = ""
        Exit Function
    End If
    
    
    If optActual(4).Value Then
        'AREA A area
        If NumRegElim <> 1 Then
            MsgBox "Selección una area únicamente", vbExclamation
            Cad = ""
            Exit Function
        End If
    End If
    
    
    If NumRegElim < List3.ListCount Then
        If Not optActual(4).Value Then
            MsgBox "Selección de areas sólo disponible para listado agrupado por areas", vbExclamation
            Cad = ""
            Exit Function
        End If
        Cad = Cad & " AND terminales.area IN (" & Mid(AuxArea, 2) & ")"
    End If
    
    
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
            

            If miRsAux!Control = 2 Then PuedeQuitarParadas = True
  
            
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
                
    
            If miRsAux!HorasNegativas = 1 Then
                FueraIntervalo2 = 1
                I = 24 - miRsAux!LaHora
            Else
                If miRsAux!LaHora >= 0 And miRsAux!LaHora <= 23 Then
                    I = miRsAux!LaHora
                    FueraIntervalo2 = 0
                Else
                    FueraIntervalo2 = 2
                    I = miRsAux!LaHora - 24
                    
                    
                    
                    
                End If
            End If
            CadPa = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
            Cad = Cad & ",'" & CadPa & "'"
            
            
            
            
            If Not Entrada Then
                HF = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                If HI > HF Then
                    'Hora acabalgada . Ha entrado el dia de antes
                    Difer = DateDiff("n", HF, HI)
                    Difer = 1440 - Difer
                Else
                    'Lo que habia hasta ahora
                    Difer = DateDiff("n", HI, HF)
                End If
                
                If FueraIntervalo2 <> 0 Then
                    If FueraIntervalo2 = 1 Then
                        Difer = Difer + 1440
                    Else
                        '
                    End If
                    
                    
                    
                End If
                Horas = Horas + Difer
        
                'Ajustada
                If Minutos > 0 Then
                    HF = HoraRectificada(HF, vEmpresa.AjusteSalida, Minutos)
                    Difer = DateDiff("n", HIAustada, HF)
                    If HIAustada > HF Then
                        'Dia antes entra
                        Difer = Difer + 1440
                    Else
                        If FueraIntervalo2 <> 0 Then Difer = Difer + 1440
                    End If
                End If
                Ajustadas = Ajustadas + Difer
                    
            
            
            
            Else
                HI = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                If Minutos > 0 Then
                    HIAustada = HoraRectificada(HI, vEmpresa.AjusteEntrada, Minutos)
                Else
                    HIAustada = HI
                End If
              
            
            End If
            
            
 
                        
                        
            
            
            If PuedeQuitarParadas Then
                If vH.DtoAlm > 0 And FueraIntervalo2 <> 1 Then
                    If QuitoMeriAlm = 0 Then
                        'Compruebo si el ticaje es menor que la hora del almuerzo
                        If HIAustada < vH.HoraDtoAlm Then
                            QuitoMeriAlm = 1
                            QuitoMeriendaAlmuerzo = vH.DtoAlm
                        End If
                    End If
                End If
                If vH.DtoMer > 0 And FueraIntervalo2 <> 1 Then
                    If QuitoMeriAlm < 2 Then
                        I = 0
                        If FueraIntervalo2 = 2 Then
                            'Seguro que quito la merienda
                            I = 1
                        Else
                            If Not Entrada Then
                                If HF > vH.HoraDtoMer Then I = 1
                            End If
                        End If
                        If I = 1 Then
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
        ImprimirTicajeActualNormal = True
        
   End If
   
   Exit Function
eImprimirTicajeActual:
    MuestraError Err.Number, Err.Description
   
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
        I = 35
        If chkHorasCombinadas(0).Value = 1 Then I = 37
        'Ordenado por nombre
        If optTrab(1).Value Then I = I + 1
        
        
        
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
            .Opcion = I
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
    
    
    
    
'    If RsBase!Entrada = 105711 Then St op
'    If RsBase!idTrabajador = 50049 Then S top

    RT.AddNew
    Cad = "Select IdInci,hour(hora) LaHora,minute(hora) minutos,second(hora) segundos, if(hora<'0:00:00',1,0) esNegativo,Hora from EntradaMarcajes WHERE IdMarcaje=" & RsBase!Entrada
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
        If RS!esnegativo = 1 Then
            
            Cad = Horas_Quitar24(RS!Hora, True)
            Fecha = CDate(Cad)
        Else
            If RS!LaHora > 23 Then
                Hora = RS!LaHora - 24
            Else
                Hora = RS!LaHora
            End If
            Fecha = Format(Hora, "00") & ":" & Format(RS!Minutos, "00") & ":" & Format(RS!segundos, "00")
        End If
        If C < 9 Then
            Select Case C
            Case 1
                RT!H1 = Fecha
            Case 2
                RT!h2 = Fecha
            Case 3
                RT!H3 = Fecha
            Case 4
                RT!H4 = Fecha
            Case 5
                RT!H5 = Fecha
            Case 6
                RT!H6 = Fecha
            Case 7
                RT!H7 = Fecha
            Case 8
                RT!H8 = Fecha
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
Dim I As Long
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
    I = 0
    While Not RS.EOF
        SQL = "INSERT INTO tmpMarcajes(codusu,incFinal,entrada,idTrabajador,Fecha,HorasTrabajadas,HorasIncid) VALUES (" & vUsu.Codigo & ",0," & I & ","
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
        I = I + 1
    Wend
    RS.Close
    
    If I = 0 Then
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
        I = 61
    Else
        I = 63
    End If
    
    If Me.optTrab(3).Value Then I = I + 1
        
    frmImprimir.Opcion = I
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
        I = 61
    Else
        I = 63
    End If
    
    If Me.optTrab(3).Value Then I = I + 1
        
    frmImprimir.Opcion = I
    frmImprimir.FormulaSeleccion = "{tmpMarcajes.codusu} = " & vUsu.Codigo
    frmImprimir.OtrosParametros = SQL
    frmImprimir.NumeroParametros = 2
    
    frmImprimir.Show vbModal
End Sub

Private Function GenerarImpresionimportesCostesDesdejornadassemanalesAlzira() As Boolean
Dim F1 As Date
Dim F2 As Date
Dim SQL As String
Dim I As Long
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
            I = 0
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                I = DBLet(RS.Fields(0), "N")
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
            
            If Val(SQL) <> I Then
                'Dias en marcajes para los trabajadores Seccion: SeccionesAjustesHoras distinto a los procesados
                SQL = "Secc. ajuste horas. Dias procesados: " & SQL
                SQL = vbCrLf & "Dias marcajes: " & I & "    " & SQL
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
    I = 0
    If SeccionesNormales <> "" Then
        SQL = "Select marcajes.*,excesodefecto from marcajes,incidencias,trabajadores where"
        SQL = SQL & " marcajes.idTrabajador = trabajadores.idTrabajador"
        SQL = SQL & " AND marcajes.incfinal = incidencias.idinci AND "
        SQL = SQL & WHERE_CostesDiarios(F1, F2)
        SQL = SQL & " AND seccion IN (" & SeccionesNormales & ")"
        
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF
            SQL = "INSERT INTO tmpMarcajes(codusu,incFinal,entrada,idTrabajador,Fecha,HorasTrabajadas,HorasIncid,HorasExt) VALUES (" & vUsu.Codigo & ",0," & I & ","
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
            I = I + 1
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
            SQL = "INSERT INTO tmpMarcajes(codusu,incFinal,entrada,idTrabajador,Fecha,HorasTrabajadas,HorasIncid,HorasExt) VALUES (" & vUsu.Codigo & ",0," & I & ","
            SQL = SQL & RS!idTrabajador & ",'"
            SQL = SQL & Format(RS!Fecha, FormatoFecha) & "',"
            
            SQL = SQL & TransformaComasPuntos(CStr(RS!HNor)) & ","
            SQL = SQL & TransformaComasPuntos(CStr(RS!HEstr)) & ","
            SQL = SQL & TransformaComasPuntos(CStr(RS!HExtr)) & ")"
            conn.Execute SQL
            'Sig
            RS.MoveNext
            I = I + 1
        Wend
        RS.Close
        
    End If
    
    Set RS = Nothing
        
    If I = 0 Then
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
    I = -1
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
    
     I = FreeFile
    Open Cad For Output As #I
    
    
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
        
        Print #I, CadPa
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Close #I
    I = 0 'Para que no intenten volver a cerrarlo
    
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
    If I > 0 Then Close #I


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
Dim Difer As Currency
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
                I = miRsAux!LaHora
                FueraIntervalo_ = 0
            Else
                FueraIntervalo_ = 24
                If miRsAux!LaHora < 0 Then Debug.Print "Stop"  'De momento NO deberia entrar aqui
                I = miRsAux!LaHora - FueraIntervalo_
            End If
            
            CadPa = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
            Cad = Cad & ",'" & CadPa & "'"
            
            
            
            
            If Not Entrada Then
                HF = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                Difer = DateDiff("s", HI, HF)
                If FueraIntervalo_ > 0 Then Difer = Difer + 86400  'segundos
                Min = Difer \ 60
                Seg = Difer - (CCur(Min) * 60)
                Difer = Round((Seg / 60), 2) + Min
                'Lo paso a decimal
                             
                
                Horas = Horas + Difer
        
                'Ajustada
                If Minutos > 0 Then
                    HF = HoraRectificada(HF, vEmpresa.AjusteSalida, Minutos)
                    Difer = DateDiff("n", HIAustada, HF)
                    If FueraIntervalo_ > 0 Then Difer = Difer + 86400
                    
                       Min = Difer \ 60
                        Seg = Difer - (Min * 60)
                        Difer = Round((Seg / 60), 2) + Min
                        'Lo paso a decimal
                            
                    
                    
                    
                End If
                Ajustadas = Ajustadas + Difer
                    
            
            
            
            Else
                HI = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
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
Dim Dias_Oficiales As Integer  'los que marque el calendairo
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
Dim EstaDeBaja As String
Dim L As Integer
Dim L2 As Integer
Dim FI As Date
Dim F As Date
Dim Impor3 As Currency
Dim Impor4 As Currency
Dim InicioMes As Date
Dim FinMes As Date
Dim B As Boolean
Dim AlgunLaborableACero As Boolean

Dim PosiblesErrores As String


    On Error GoTo egenerarDatosNominas
    generarDatosNominas = False
    Set miRsAux = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    'tmppagosmes   idTrabajador,Nombre,IRPF,SS,importe1,importe2
    Cad = "DELETE FROM tmppagosmes"
    conn.Execute Cad
    
    Cad = "DELETE FROM tmpmarcajes"
    conn.Execute Cad
    
    FF = CDate(txtFec(20).Text)
    
    Cad = "INSERT INTO tmpmarcajes(Entrada,fecha) VALUES (-1," & DBSet(FF, "F") & ")"
    conn.Execute Cad
    
    
    'De momento cojo el IDCAL1
    'FALTA###
    DiasDelMes = DiasMes(Month(FF), Year(FF))
    
    InicioMes = "01/" & Format(CDate(txtFec(20).Text), "mm/yyyy")
    FinMes = DiasDelMes & Format(CDate(txtFec(20).Text), "/mm/yyyy")
    
    H = CalculaHorasHorarioALZConVector(1, Dias_Oficiales, CDate("01" & Format(FF, "/mm/yyyy")), CDate(DiasDelMes & Format(FF, "/mm/yyyy")), DiasTrabajadosPorMes)
    
    
    Set vH = New CHorarios
    DiasFestivos = vH.LeerDiasFestivos(1, CDate("01" & Format(FF, "/mm/yyyy")), CDate(DiasDelMes & Format(FF, "/mm/yyyy")))
    'Añadimos los domingos
    For I = 1 To DiasDelMes
        If Format(CDate(I & "/" & Format(FF, "mm/yyyy")), "w") = 1 Then DiasFestivos = DiasFestivos & CDate(I & "/" & Format(FF, "mm/yyyy")) & "|"
    Next
    'Lo transformo en un array de enteros
    vSQL = DiasFestivos
    Cad = ""
    VectorDiasFestivos = ""
    While vSQL <> ""
        I = InStr(1, vSQL, "|")
        Cad = Mid(vSQL, 1, I - 1)
        vSQL = Mid(vSQL, I + 1)
        VectorDiasFestivos = VectorDiasFestivos & ", " & Day(CDate(Cad))
        
    Wend
    VectorDiasFestivos = Mid(VectorDiasFestivos, 2)
    
    Set vH = Nothing
    
    Cad = "select nominas.* from nominas"
    Cad = Cad & " where month(fecha)=" & Month(FF) & " and year(fecha)=" & Year(FF)
    If Me.cboCentroTrabajo.ListIndex >= 0 Then Cad = Cad & " and CentroA3 = " & Me.cboCentroTrabajo.ItemData(Me.cboCentroTrabajo.ListIndex)
      
      
    PosiblesErrores = ""
    
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    vSQL = ""
    While Not miRsAux.EOF
        I = I + 1
        VectorDiasTrab = CStr(DiasTrabajadosPorMes)  'Lo copio
       
        
        'If miRsAux!idTrabajador = 101 Then Stop
        'If miRsAux!idTrabajador = 146 Then St op
        
        'Veremos si ha trabajado algun dia festivo fesivos.. FESTIVO
        'Eso implicara que
        diasTrabajados = miRsAux!Dias
        
        'Calculamos dias
        'Vector: FNNNNFFFNNNNNFFNNNNNFFNNNNNFFNN   -> Son los dias del 1 al diasmes que son festivos. Iremos cambiando las N por S los dias que haya trabajado
        'Trabaja los dias oficiales.
        'No hacemos NADA, reemplazamoslas N de
        
        B = False
        If vEmpresa.QueEmpresa <> vbCatadau Then
            If diasTrabajados = Dias_Oficiales Then B = True
        End If
        If B Then
            'Perfecto. Todas las N son S
            VectorDiasTrab = Replace(VectorDiasTrab, "N", "S")
            
        
        Else
            'Empieza la fiesta. Ha trabajado menos dias
            
            Cad = "select  FecBaja,FecAlta  from trabajadores where idtrabajador=" & miRsAux!idTrabajador
            RS.Open Cad, conn, adOpenKeyset, adCmdText
            If Not RS.EOF Then
                If Not IsNull(RS!FecAlta) Then
                    F = RS!FecAlta
                    If F >= InicioMes And F <= FinMes Then
                        'St op
                        For I = 1 To Day(F) - 1   '-1:  El dia del alta TRABAJAN
                            VectorDiasTrab = Mid(VectorDiasTrab, 1, I - 1) & "B" & Mid(VectorDiasTrab, I + 1)
                        Next
                    End If
                End If
                
                If Not IsNull(RS!FecBaja) Then
                    F = RS!FecBaja
                    If F >= InicioMes And F <= FinMes Then
                        'St op
                        For I = Day(F) + 1 To Len(VectorDiasTrab)
                            VectorDiasTrab = Mid(VectorDiasTrab, 1, I - 1) & "B" & Mid(VectorDiasTrab, I + 1)
                        Next
                    End If
                End If
                
            End If
            RS.Close
            
            
            
            'Dias que esta de baja. NO deberia haber fichaje
            FI = "01/" & Format(FF, "mm/yyyy")
            EstaDeBaja = "|"
            Cad = "select * from bajas where idtrab=" & miRsAux!idTrabajador & " and fechabaja<=" & DBSet(FF, "F")
            Cad = Cad & " and (fechaalta is null or fechaalta>" & DBSet(FI, "F") & ")"
            RS.Open Cad, conn, adOpenKeyset, adCmdText
            While Not RS.EOF
                If RS!FechaBaja < FI Then
                    L = 1
                    
                Else
                    L = Day(RS!FechaBaja)
                End If
                If IsNull(RS!fechaalta) Then
                    k = Day(FF) + 1
                Else
                    k = Day(RS!fechaalta)
                End If
                
                For L2 = L To k
                    EstaDeBaja = EstaDeBaja & L2 & "|"
                Next
                RS.MoveNext
            Wend
            RS.Close
                
            
            
            If vEmpresa.QueEmpresa <> vbCatadau Then
                'Menos para catadau
                
                'Los leemos de marcajes
                'Veremos cuales ha trabajado, SEGURO entre sin finde ni festivos
                Cad = " festivo=0 AND weekday(fecha)<5 and"
                Cad = Cad & " idtrabajador = " & miRsAux!idTrabajador
                Cad = "select fecha , 1 laborable from marcajes  WHERE " & Cad
                Cad = Cad & " AND month(fecha)=" & Month(FF) & " and year(fecha)=" & Year(FF)
                
            Else
                'Catadau. Dia que viene, dia que cotiza
                'Los leemos de jopranadassemanakes
                
                
                Cad = " idtrabajador = " & miRsAux!idTrabajador
                Cad = "select fecha,sum(laborable) laborable from jornadassemanalesalz  WHERE " & Cad
                Cad = Cad & " AND month(fecha)=" & Month(FF) & " and year(fecha)=" & Year(FF) & " GROUP BY fecha"
            End If

            
            RS.Open Cad, conn, adOpenKeyset, adCmdText
            k = miRsAux!Dias
            AlgunLaborableACero = False
            While Not RS.EOF
            
                If vEmpresa.QueEmpresa = 4 Then
                    'catadAU..     dIA QUE VIENE, da igula lo que sea.   X
                    J = Day(RS!Fecha)
                    VectorDiasTrab = Mid(VectorDiasTrab, 1, J - 1) & "S" & Mid(VectorDiasTrab, J + 1)
                
                
                
                Else
                    If InStr(1, DiasFestivos, Format(RS!Fecha, "dd/mm/yyyy")) = 0 Then
                        If RS!Laborable > 0 Then
                            J = Day(RS!Fecha)
                            VectorDiasTrab = Mid(VectorDiasTrab, 1, J - 1) & "S" & Mid(VectorDiasTrab, J + 1)
                            k = k - 1
                      
                        Else
                            'No esta maracado como labroiarble. Suele ser sabado, que ya ha venido miercoles a trabajar
                               ' Stop
                            AlgunLaborableACero = True
                        End If
                   
                    End If
                End If
                RS.MoveNext
            Wend
            RS.Close
            
            
                        
            
            
            If vEmpresa.QueEmpresa = vbCatadau Then
            
                'Da lo mismo. En el proceso de arriba hemos marcado CADA dia como trabajado. Venga media hora o venga 24
                k = 0
            
            
                'En KATADAU, los que falten dias a cotizar(k>0) entonces si tiene algun sabado marcaado como "no laborable", se lo vuelvo a marcar
                If k > 0 And AlgunLaborableACero Then
                    
                    RS.Open Cad, conn, adOpenKeyset, adCmdText
                    While Not RS.EOF
                        If k > 0 Then
                            If InStr(1, DiasFestivos, Format(RS!Fecha, "dd/mm/yyyy")) = 0 Then
                                If RS!Laborable = 0 Then
                                    J = Day(RS!Fecha)
                                    
                                    If Mid(VectorDiasTrab, J, 1) = "N" Then
                                    
                                        VectorDiasTrab = Mid(VectorDiasTrab, 1, J - 1) & "S" & Mid(VectorDiasTrab, J + 1)
                                        k = k - 1
                                    End If
                                End If
                            End If
                        End If
                        RS.MoveNext
                    Wend
                    RS.Close
            
                End If
                'En catadau, NO compensan dias ni nada
                If k > 0 Then
                    
                    'Vamos a ver si hay alguno incorrecto
                    Cad = ""
                    For J = 1 To Len(VectorDiasTrab)
                        If Mid(VectorDiasTrab, J, 1) = "S" Then Cad = Cad & "X"
                    Next J
                    If Len(Cad) <> miRsAux!Dias Then
                        PosiblesErrores = PosiblesErrores & Mid(miRsAux!Dias & "    ", 1, 8) & Mid(Len(Cad) & "    ", 1, 8)
                        Cad = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idtrabajador", miRsAux!idTrabajador)
                        PosiblesErrores = PosiblesErrores & Cad & "  (" & miRsAux!idTrabajador & ")" & vbCrLf
                    End If
                End If
            End If
            
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
               ' If diasTrabajados <> DiasOficiales Then St op
            End If
            
            If k > 0 Then
                Cad = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idtrabajador", miRsAux!idTrabajador)
                Cad = miRsAux!idTrabajador & " " & Cad & vbCrLf
                Cad = "NO ha compensado todos los dias: " & vbCrLf & Cad
                If vEmpresa.QueEmpresa <> 4 Then MsgBox Cad, vbExclamation
                
            End If
        End If
        
        
        If Len(EstaDeBaja) > 1 Then
            'Hay bajas
            'Quitamos el primer pipe
            EstaDeBaja = Mid(EstaDeBaja, 2)
            
            While EstaDeBaja <> ""
                L2 = InStr(1, EstaDeBaja, "|")
                If L2 = 0 Then
                    EstaDeBaja = ""
                Else
                    Cad = Mid(EstaDeBaja, 1, L2 - 1)
                    EstaDeBaja = Mid(EstaDeBaja, L2 + 1)
                    L = Val(Cad)
                    Cad = Mid(VectorDiasTrab, L, 1)
                    If Cad = "S" Then MsgBox "Trabaja dia de baja: " & L & vbCrLf & "  Trabajador: " & miRsAux!idTrabajador, vbExclamation
                    
                    VectorDiasTrab = Mid(VectorDiasTrab, 1, L - 1) & "B" & Mid(VectorDiasTrab, L + 1)
                
                End If
            
            
            Wend
            
        End If
        
        'DAtos para insertar en tmp
        '----------------------------------
        
        
        
        
        If vEmpresa.QueEmpresa = vbCatadau Then
            'CATA
            
                
            'Noramles-> Bruto
            
            Importe = DBLet(miRsAux!Bruto, "N")
            'PLUS  Plus"
            H = miRsAux!plus
            'Son EXTRAS
             H = DBLet(miRsAux!HE, "N")
            
            'Neto -> estrcuturales ,ImporEstruc 'Imp est',
            Impor3 = miRsAux!ImporEstruc
            
            Impor4 = miRsAux!LlevaPlus  'Son el plus de las nominas
        Else
            'Coopic.
            H = 0
            If miRsAux!hp > 0 Then H = miRsAux!hp * miRsAux!preciohe
            If miRsAux!HC > 0 Then H = H + miRsAux!HC * miRsAux!preciohc
            H = Round(H, 2)
            
            Importe = miRsAux!HN * miRsAux!preciohn
            Impor3 = 0  'estrcuturales en CATA
            Impor4 = 0 ' PlusHoras cata
        End If
        
        
        Cad = ", (" & miRsAux!idTrabajador & ",'" & VectorDiasTrab & "'," & miRsAux!Dias & "," & DiasDelMes & ","
        Cad = Cad & DBSet(Importe, "N") & "," & DBSet(H, "N") & "," & DBSet(Impor3, "N") & "," & DBSet(Impor4, "N") & ")"
        vSQL = vSQL & Cad
        If (I Mod 10) = 0 Then
            vSQL = Mid(vSQL, 2)
            vSQL = "INSERT INTO tmppagosmes(idTrabajador,Nombre,IRPF,SS,importe1,importe2,Neto,Bruto) VALUES " & vSQL
            conn.Execute vSQL
            DoEvents
            vSQL = ""
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If vSQL <> "" Then
        vSQL = Mid(vSQL, 2)
        vSQL = "INSERT INTO tmppagosmes(idTrabajador,Nombre,IRPF,SS,importe1,importe2,Neto,bruto) VALUES " & vSQL
        conn.Execute vSQL
    End If
    
    
    If vUsu.Codigo = 0 Then
        If PosiblesErrores <> "" Then
            PosiblesErrores = "Dias        X's       Trbajador" & vbCrLf & String(30, "=") & vbCrLf & PosiblesErrores
            MsgBox PosiblesErrores, vbExclamation
        End If
    End If
    
    If I > 0 Then generarDatosNominas = True
    
egenerarDatosNominas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set RS = Nothing
End Function



Private Sub CargaCentroTrabajo()
    cboCentroTrabajo.Clear
    If Not vEmpresa.TieneCentrosA3 Then Exit Sub
    
    
    Cad = "SELECT * FROM CentrosA3  ORDER BY idCentro"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not miRsAux.EOF
        cboCentroTrabajo.AddItem miRsAux!nomcentro
        cboCentroTrabajo.ItemData(cboCentroTrabajo.NewIndex) = miRsAux!idCentro
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
            
End Sub






'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'
'       Listado horas por relojes
'
'
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Private Sub HacerListadoPorTipoReloj()

    On Error GoTo EHacerListadoCombinado
    Screen.MousePointer = vbHourglass
    lblCombinado.Caption = "Comienzo proceso"
    lblCombinado.Refresh
    If CargaDatosTipoReloj Then
        
        espera 0.25
        

        
        
        If chkHorasCombinadas(0).Value = 1 Then
            'FECHA
            If optTrab(0).Value Then
                I = 68
            Else
                I = 71
            End If
        Else
            
            If optTrab(0).Value Then
                'Ordenado por nombre
                I = 70
            Else
                'Ordenado por codigo
                I = 69
            End If
        End If
        
        
        
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
        
        Cad = "Todos los relojes"
        If Me.cboReloj.ListIndex = 1 Then Cad = "Solo principal"
        If Me.cboReloj.ListIndex = 2 Then Cad = "Biostar"
        vSQL = Trim(Cad & "    " & vSQL)
        CadPa = CadPa & "FechaIni= """ & vSQL & """|"
        
        
        CadPa = CadPa & "EnDecimal= " & Abs(Me.chkHorasCombinadas(1).Value) & "|"
        
        NumPa = 3
        
        
        
        
        
        
        
        
        With frmImprimir
            .FormulaSeleccion = "{tmppresencia.codusu} = " & vUsu.Codigo
            .OtrosParametros = CadPa
            .Opcion = I
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

Private Function CargaDatosTipoReloj() As Boolean
Dim RsBase As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim CadenaInsert As String
Dim CadenaSQL As String
Dim Cad As String
Dim Hora As Integer
Dim HoraF As String
Dim idReloj As Integer



Dim CadenaInsertReloj1 As String
Dim CadenaInsertReloj2 As String
Dim Num1 As Integer  'Si es la primera , las segunda.. para el reloj 1
Dim Num2 As Integer
Dim H1 As Currency
Dim h2 As Currency
Dim Entrada As Boolean
Dim HoraAnt As Date
Dim HorasParada As Currency
Dim AnteriorA_Las_9 As Boolean
Dim YaDescontado As Boolean
Dim Aux As String

Dim AlmuerzoEnReloj As Byte
On Error GoTo ErrSQL
    CargaDatosTipoReloj = False
    
    
    
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
    Cad = "SELECT Marcajes.entrada,Marcajes.idTrabajador,Marcajes.Fecha,Marcajes.HorasTrabajadas,Marcajes.HorasIncid,ExcesoDefecto,"
    Cad = Cad & "Marcajes.idhorario,incfinal,Trabajadores.idcal,secciones.Nombre,Trabajadores.NomTrabajador,HorasDto"
    Cad = Cad & " FROM Trabajadores,Marcajes,Incidencias,Secciones"
    Cad = Cad & " WHERE  Trabajadores.IdTrabajador = Marcajes.idTrabajador"
    Cad = Cad & " AND Trabajadores.Seccion = Secciones.Idseccion"
    Cad = Cad & " AND Incidencias.idInci = Marcajes.IncFinal"
    Cad = Cad & " AND Marcajes.correcto = 1"
    
    'unimos la cadena sql
    Cad = Cad & CadenaSQL
    
    Cad = Cad & " ORDER BY idhorario,fecha,idcal"
    Cad = Cad & " , " & IIf(optTrab(0).Value, "Trabajadores.IdTrabajador", "Trabajadores.NomTrabajador")
    
    Set RsBase = New ADODB.Recordset
    lblCombinado.Caption = "Obteniendo conjunto registros"
    lblCombinado.Refresh
    RsBase.Open Cad, conn, , , adCmdText
    If RsBase.EOF Then
        MsgBox "Ningun registro con esos valores", vbExclamation
        Set RsBase = Nothing
        Exit Function
    End If
    'Borramos los registros anteriores   tmpinformehorasmes
    conn.Execute "Delete  from tmppresencia where codusu = " & vUsu.Codigo
    
    'Empezamos para insertar
    
    Set RS = New ADODB.Recordset
    CadenaInsert = ""
    DoEvents
    NumRegElim = 0
    While Not RsBase.EOF
    
        'tmppresencia(Id,NomTrabajador,NomEmpresa,Fecha,H1,H2,H3,H4,H5,H6,H7,H8,Incidencias,Seccion,idtra,codusu,semana)
    
        lblCombinado.Caption = RsBase!Fecha & " " & RsBase!idTrabajador
        lblCombinado.Refresh
        
        Cad = "Select IdInci,hour(hora) LaHora,minute(hora) minutos,second(hora) segundos, if(hora<'0:00:00',1,0) esNegativo,Hora,reloj from EntradaMarcajes WHERE IdMarcaje=" & RsBase!Entrada
        Cad = Cad & " ORDER BY Horareal,reloj"
        RS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
       
       
        
       
        
     
        CadenaInsertReloj1 = ""
        CadenaInsertReloj2 = ""
        Num1 = 0
        Num2 = 0
        H1 = 0
        h2 = 0
        HorasParada = RsBase!HorasDto
        AnteriorA_Las_9 = False
        YaDescontado = False
        AlmuerzoEnReloj = 101
        Entrada = True
        CadenaSQL = ""  'si hay alguna incidcencia al listado
        While Not RS.EOF
            'Como pintaremos la hora
            If RS!esnegativo = 1 Then
                Cad = Horas_Quitar24(RS!Hora, True)
                HoraF = CDate(Cad)
            Else
                If RS!LaHora > 23 Then
                    Hora = RS!LaHora - 24
                Else
                    Hora = RS!LaHora
                End If
                HoraF = Format(Hora, "00") & ":" & Format(RS!Minutos, "00") & ":" & Format(RS!segundos, "00")
            End If
            
            
            If Entrada Then
                'Veremos si va al reloj uno , u 2
                idReloj = RS!Reloj
                HoraAnt = HoraF
                Entrada = False
                If HoraF < CDate("9:00:00") Then AnteriorA_Las_9 = True
            Else
                If idReloj <> RS!Reloj Then
                    
                    'ha fichado en relojes distintos la E/S
                    CadenaSQL = "Dos relojes distintos E/S"
                End If
                
                
                If idReloj = 1 Then
                    Num2 = Num2 + 1
                    If Num2 > 4 Then MsgBox "Error. Mas de 4 E/S por reloj /dia"
                    CadenaInsertReloj2 = CadenaInsertReloj2 & ",'" & HoraAnt & "','" & HoraF & "'"
                    h2 = h2 + DateDiff("n", HoraAnt, HoraF)
                    If HorasParada > 0 And AnteriorA_Las_9 And Not YaDescontado Then
                        HorasParada = 0
                        YaDescontado = True
                        AlmuerzoEnReloj = 2
                    End If
                Else
                    Num1 = Num1 + 1
                    If Num1 > 4 Then MsgBox "Error. Mas de 4 E/S por reloj /dia"
                    CadenaInsertReloj1 = CadenaInsertReloj1 & ",'" & HoraAnt & "','" & HoraF & "'"
                    H1 = H1 + DateDiff("n", HoraAnt, HoraF)
                    
                    If HorasParada > 0 And AnteriorA_Las_9 And Not YaDescontado Then
                        HorasParada = 0
                        YaDescontado = True
                        AlmuerzoEnReloj = 1
                    End If
                    
                End If
                
                
                AnteriorA_Las_9 = False
                Entrada = True
            End If
            
            RS.MoveNext
            
            
        Wend
        
        RS.Close
        
        
        
        
        'Insertamos en la tabla temporal
        'tmppresencia(Id,NomTrabajador,Fecha,Incidencias,Seccion,idtra,codusu,semana,NomEmpresa,H1,H2,H3,H4,H5,H6,H7,H8)
        '   noempre: horas   semana:reloj  icidncias : txt si pasa alog   semana:numero reloj
        Cad = "," & DBSet(RsBase!nomtrabajador, "T") & "," & DBSet(RsBase!Fecha, "F") & "," & DBSet(CadenaSQL, "T", "S") & ","
        Cad = Cad & DBSet(RsBase!Nombre, "T") & "," & RsBase!idTrabajador & "," & vUsu.Codigo & ","
        
        
        'Insertamos reloj 1

        If Me.cboReloj.ListIndex <> 2 Then  'ha pedido todos o reloj 1
            If CadenaInsertReloj1 <> "" Then
                HorasParada = 0
                If AlmuerzoEnReloj = 1 Then
                    'Almuezo en el primero
                    HorasParada = RsBase!HorasDto
                    H1 = H1 - (RsBase!HorasDto * 60)
                    AlmuerzoEnReloj = 127
                End If
            
            
                I = H1 \ 60
                H1 = H1 - (60 * I)
                H1 = Round((H1 / 60), 2)
                H1 = I + H1

                For I = Num1 + 1 To 4
                    CadenaInsertReloj1 = CadenaInsertReloj1 & ",NULL,NULL"
                Next I
                Aux = Right(Space(10) & CStr(H1), 10) & Right(Space(10) & CStr(HorasParada), 10)
                
                CadenaSQL = Cad & "1,'" & Aux & "'" & CadenaInsertReloj1
                NumRegElim = NumRegElim + 1
                CadenaInsert = CadenaInsert & ", (" & NumRegElim & CadenaSQL & ")"
            End If
        End If
        
        'Reloj2
        If Me.cboReloj.ListIndex <> 1 Then  'ha pedido todos o reloj 2
            If CadenaInsertReloj2 <> "" Then
                HorasParada = 0
                If AlmuerzoEnReloj = 2 Then
                    'Almuezo en el primero
                    HorasParada = RsBase!HorasDto
                    h2 = h2 - (RsBase!HorasDto * 60)
                    AlmuerzoEnReloj = 127
                End If
            
            
                I = h2 \ 60
                h2 = h2 - (60 * I)
                h2 = Round((h2 / 60), 2)
                h2 = I + h2
                For I = Num2 + 1 To 4
                    CadenaInsertReloj2 = CadenaInsertReloj2 & ",NULL,NULL"
                Next I
                
                Aux = Right(Space(10) & CStr(h2), 10) & Right(Space(10) & CStr(HorasParada), 10)
                
                CadenaSQL = Cad & "2,'" & Aux & "'" & CadenaInsertReloj2
                NumRegElim = NumRegElim + 1
                CadenaInsert = CadenaInsert & ", (" & NumRegElim & CadenaSQL & ")"
            End If
        End If
        

        
        
        RsBase.MoveNext
        
        If Len(CadenaInsert) > 10000 Then
            Cad = "INSERT INTO tmppresencia(Id,NomTrabajador,Fecha,Incidencias,Seccion,idtra,codusu,semana,NomEmpresa,H1,H2,H3,H4,H5,H6,H7,H8) VALUES "
            CadenaInsert = Mid(CadenaInsert, 2)
            conn.Execute Cad & CadenaInsert
            CadenaInsert = ""
        End If
        
    Wend
    If CadenaInsert <> "" Then
        Cad = "INSERT INTO tmppresencia(Id,NomTrabajador,Fecha,Incidencias,Seccion,idtra,codusu,semana,NomEmpresa,H1,H2,H3,H4,H5,H6,H7,H8) VALUES "
        CadenaInsert = Mid(CadenaInsert, 2)
        conn.Execute Cad & CadenaInsert
    End If
    RsBase.Close
    Set RS = Nothing
    Set RsBase = Nothing
    If NumRegElim > 0 Then
        CargaDatosTipoReloj = True
    Else
        MsgBox "Ningun dato generado ", vbExclamation
    End If
    lblCombinado.Caption = ""
    Exit Function
ErrSQL:
    MsgBox "Error: " & Err.Description, vbExclamation
End Function









'Adaptacion del ticaje NORMAL al horario
Private Function ImprimirTicajeActualAdaptado() As Boolean
Dim SQL As String
Dim F As Date
Dim T As Long
Dim vHora As Integer


Dim PuedeQuitarParadas As Boolean
Dim Entrada As Boolean
'De momento no trabajao con esto
Dim FueraIntervalo As Integer     '0 NORMAL    1. Hora negativa     2. Hora mayor=24

Dim vH As CHorarios
Dim Minutos As Integer
Dim HI As Date
Dim HF As Date
Dim HIAustada As Date
Dim Difer As Currency
Dim Horas  As Currency
Dim Ajustadas As Currency

Dim QuitoMeriendaAlmuerzo As Currency
Dim QuitoMeriAlm As Byte '0 No he quitado nada     1. Ya he quitado almuerzo    2. Quito la merienda
Dim EsLaUltima As Boolean
Dim AuxiliarResta As Integer
Dim MaximoMinutosDias As Integer
Dim Restamoas As Boolean
Dim DiaSem As Integer



    On Error GoTo eImprimirTicaje2

    ImprimirTicajeActualAdaptado = False
    
    SQL = "Delete from tmpCombinada where codusu = " & vUsu.Codigo
    conn.Execute SQL
    Set vH = New CHorarios
    '''Sql = "Select entradafichajes.*,nomtrabajador from entradafichajes,trabajadores where entradafichajes.idtrabajador =trabajadores.idtrabajador "
    SQL = "select entradafichajes.idtrabajador,fecha,hour(hora) lahora,minute(hora) minutos,second(hora) segundos ,Control, "
    If vEmpresa.AcabalgadoDiaInicio Then
        SQL = SQL & " 0 "   'Acabalgdo dia inicio no hay hpras negativas
    Else
        SQL = SQL & "if(hora<0,1,0) "
    End If
    SQL = SQL & " HorasNegativas from entradafichajes inner join trabajadores t on t.idtrabajador=entradafichajes.idtrabajador"
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
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
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
            EsLaUltima = False
    
            
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
                
'                    If vH.Rectificar > 0 Then
'                      If vH.Rectificar = vbRecESCuarto Then
'                        Minutos = 15
'                      Else
'                          Minutos = 30   'Entradas salidas cada media hora
'                      End If
'                    End If
                 
                If vH.DtoMer = 0 And vH.DtoAlm = 0 Then PuedeQuitarParadas = False
  
                 
            
            vSQL = "INSERT INTO tmpCombinada(codusu,idTrabajador,Fecha,HT,HE,HR,idinci,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15,H16) VALUES (" & vUsu.Codigo & ","
            vSQL = vSQL & miRsAux!idTrabajador & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "',"
            Cad = ""
        End If
        
        
        If vHora < 16 Then   'solo ionserto 16
                
    
           ' If miRsAux!HorasNegativas = 1 Then
           '     FueraIntervalo2 = 1
           '     i = 24 - miRsAux!LaHora
           ' Else
           '     If miRsAux!LaHora >= 0 And miRsAux!LaHora <= 23 Then
           '         i = miRsAux!LaHora
           '         FueraIntervalo2 = 0
           '     Else
           '         FueraIntervalo2 = 2
           '         i = miRsAux!LaHora - 24
           '
           '     End If
           ' End If
           I = miRsAux!LaHora
            CadPa = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
            
            If Entrada Then HIAustada = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
            
            
            
            If PuedeQuitarParadas Then
                'If vH.DtoAlm > 0 And FueraIntervalo2 <> 1 Then
                If vH.DtoAlm > 0 Then
                    If QuitoMeriAlm = 0 Then
                        'Compruebo si el ticaje es menor que la hora del almuerzo
                        If HIAustada < vH.HoraDtoAlm Then
                            QuitoMeriAlm = 1
                            QuitoMeriendaAlmuerzo = vH.DtoAlm
                        End If
                    End If
                End If
                'If vH.DtoMer > 0 And FueraIntervalo2 <> 1 Then
                If vH.DtoMer > 0 Then
                    If QuitoMeriAlm < 2 Then
                        I = 0
                        'If FueraIntervalo2 = 2 Then
                        '    'Seguro que quito la merienda
                        '    i = 1
                        'Else
                            If Not Entrada Then
                                If HF > vH.HoraDtoMer Then I = 1
                            End If
                        'End If
                        If I = 1 Then
                            QuitoMeriAlm = 2
                            QuitoMeriendaAlmuerzo = QuitoMeriendaAlmuerzo + vH.DtoMer
                        End If
                        
                    End If
                End If
            End If
            
            
            
            
            
            
            If Not Entrada Then
            
                miRsAux.MoveNext
                
                If miRsAux.EOF Then
                     EsLaUltima = True
                Else
                    If miRsAux!idTrabajador <> T Then
                        EsLaUltima = True
                    Else
                        If miRsAux!Fecha <> F Then EsLaUltima = True
                    End If
                End If

                miRsAux.MovePrevious  'volvemos al sitio
                
                HF = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
               ' If HI > HF Then
               '     'Hora acabalgada . Ha entrado el dia de antes
               '     Difer = DateDiff("n", HF, HI)
               '     Difer = 1440 - Difer
               ' Else
                    'Lo que habia hasta ahora
                    Difer = DateDiff("n", HI, HF)
               ' End If
                
              
               
        
               
                Ajustadas = Ajustadas + Difer
                    
                If EsLaUltima Then
                
                    MaximoMinutosDias = vH.TotalHoras * 60
                    
                    MaximoMinutosDias = MaximoMinutosDias + (QuitoMeriendaAlmuerzo * 60)   'maximo minutos es Horas * 8 +as paradas
                    
                    If vH.TotalHoras = 0 Then MaximoMinutosDias = 32000
                    If Ajustadas > MaximoMinutosDias Then
                        
                        MaximoMinutosDias = Ajustadas - MaximoMinutosDias - 1
                        'Para que no seam exactos
                        'En vez de random, por si acaso reimprime muchas veces el mismo listado, que siempre salga lo mismo
                        
                        
                        Restamoas = False
                        DiaSem = Weekday(F)
                        DiaSem = DiaSem Mod 4
                        Select Case DiaSem
                        Case 1
                            AuxiliarResta = (miRsAux!idTrabajador Mod 8) + 1
                            If AuxiliarResta = 1 Or AuxiliarResta = 5 Then Restamoas = True
                            
                            If (AuxiliarResta Mod 3) = 0 Then
                                AuxiliarResta = (AuxiliarResta + 1) Mod 4
                            ElseIf (AuxiliarResta Mod 3) = 1 Then
                                AuxiliarResta = (AuxiliarResta + 1) Mod 5
                            Else
                                AuxiliarResta = (AuxiliarResta + 1) Mod 3
                            End If
                        Case 2
                            AuxiliarResta = (miRsAux!idTrabajador Mod 8) + 1
                            If AuxiliarResta = 2 Or AuxiliarResta = 6 Then Restamoas = True
                            
                            If (AuxiliarResta Mod 3) = 0 Then
                                AuxiliarResta = (AuxiliarResta + 1) Mod 1
                                ElseIf (AuxiliarResta Mod 3) = 1 Then
                                AuxiliarResta = (AuxiliarResta + 1) Mod 2
                            Else
                                AuxiliarResta = (AuxiliarResta + 1) Mod 3
                            End If
                            AuxiliarResta = AuxiliarResta + 1
                        
                        
                        Case 3
                            AuxiliarResta = (miRsAux!idTrabajador Mod 8) + 1
                            If AuxiliarResta = 3 Or AuxiliarResta = 7 Then Restamoas = True
                            
                            If (AuxiliarResta Mod 3) = 0 Then
                                AuxiliarResta = (AuxiliarResta + 1) Mod 2
                            ElseIf (AuxiliarResta Mod 3) = 2 Then
                                AuxiliarResta = (AuxiliarResta + 1) Mod 3
                            Else
                                AuxiliarResta = (AuxiliarResta + 1) Mod 1
                            End If
                            AuxiliarResta = AuxiliarResta + 1
                        
                        
                        
                        Case Else
                            AuxiliarResta = (miRsAux!idTrabajador Mod 8) + 1
                            If AuxiliarResta = 0 Or AuxiliarResta = 8 Then Restamoas = True
                            
                            If (AuxiliarResta Mod 2) = 0 Then
                                AuxiliarResta = (miRsAux!idTrabajador + 1) Mod 3
                            Else
                                AuxiliarResta = (miRsAux!idTrabajador + 1) Mod 5
                            End If
                        
                        End Select
                        
                        
                        
                        If Restamoas Then AuxiliarResta = -1 * AuxiliarResta
                            
                        
                        MaximoMinutosDias = MaximoMinutosDias - AuxiliarResta
                        
                        
                        
                        HF = DateAdd("n", -MaximoMinutosDias, HF)
                        
                    Else
                        MaximoMinutosDias = 0
                    End If
                    Difer = Difer - MaximoMinutosDias
                    
                End If
                CadPa = Format(HF, "hh:nn:ss")
                Horas = Horas + Difer
                Cad = Cad & ",'" & CadPa & "'"
            
            
            Else
                Cad = Cad & ",'" & CadPa & "'"
                HI = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                'If Minutos > 0 Then
                '    HIAustada = HoraRectificada(HI, vEmpresa.AjusteEntrada, Minutos)
                'Else
                '    HIAustada = HI
                'End If
              
                HIAustada = HI
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
        ImprimirTicajeActualAdaptado = True
   
   End If
   
   Exit Function
eImprimirTicaje2:
    MuestraError Err.Number, Err.Description
   
End Function














'--------------------------------
Private Function ImprimirDatosHco() As Boolean
Dim SQL As String
Dim F As Date
Dim T As Long
Dim vHora As Integer



Dim FueraIntervalo2 As Integer     '0 NORMAL    1. Hora negativa     2. Hora mayor=24
Dim vH As CHorarios
Dim Minutos As Integer
Dim HI As Date
Dim HF As Date
Dim HIAustada As Date
Dim Difer As Currency
Dim Horas  As Currency


    On Error GoTo eImprimirDatosHco
    ImprimirDatosHco = False

    SQL = "Delete from tmppresencia where codusu = " & vUsu.Codigo
    conn.Execute SQL

    '''Sql = "Select entradafichajes.*,nomtrabajador from entradafichajes,trabajadores where entradafichajes.idtrabajador =trabajadores.idtrabajador "
    
    Cad = "select marcajeshco.idtrabajador,marcajeshco.fecha,hour(hora) lahora,minute(hora) minutos,second(hora) segundos , "
    If vEmpresa.AcabalgadoDiaInicio Then
        Cad = Cad & " 0 "  'Acabalgdo dia inicio no hay hpras negativas
    Else
        Cad = Cad & "if(hora<0,1,0) "
    End If
    Cad = Cad & " HorasNegativas,marcajeshco.* "
    
    Cad = Cad & " from marcajeshco,entradamarcajeshco where marcajeshco.entrada =entradamarcajeshco.idmarcaje"
    If Me.txtFec(24).Text <> "" Then Cad = Cad & " AND marcajeshco.fecha >='" & Format(txtFec(24).Text, FormatoFecha) & "'"
    If Me.txtFec(25).Text <> "" Then Cad = Cad & " AND marcajeshco.fecha <='" & Format(txtFec(25).Text, FormatoFecha) & "'"
    If Me.txtTrab(24).Text <> "" Then Cad = Cad & " AND marcajeshco.idtrabajador >= " & txtTrab(24).Text
    If Me.txtTrab(25).Text <> "" Then Cad = Cad & " AND marcajeshco.idtrabajador <= " & txtTrab(25).Text
    
    
    
    
    Cad = Cad & " ORDER BY marcajeshco.fecha,marcajeshco.idtrabajador,hora"
        
    Set miRsAux = New ADODB.Recordset
   
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    F = CDate("01/01/1911")
    T = 0
    NumRegElim = 0
    vSQL = ""
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        If miRsAux!Fecha <> F Then
            If T > 0 Then T = 9999999
        End If
        If T <> miRsAux!idTrabajador Then
            
            If T > 0 Then InsertaProcesarHco vHora, False
                
            'tmppresencia(Id,NomTrabajador,NomEmpresa,Fecha,H1,H2,H3,H4,H5,H6,H7,H8,Incidencias,Seccion,idtra,codusu,semana)
            'Empezamos con el SQL
            F = miRsAux!Fecha
            T = miRsAux!idTrabajador
            
     
            
          '  vSql = "INSERT INTO tmpCombinada(codusu,idTrabajador,Fecha,HT,HE,HR,idinci,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15,H16) VALUES (" & vUsu.Codigo & ","
          '  vSql = vSql & miRsAux!idTrabajador & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "',"
            
           
            
            Cad = "(" & DBSet(miRsAux!nomtrabajador2, "T") & "," & DBSet(IIf(miRsAux!IncFinal = 0, "", miRsAux!nominci2), "T") & ",'" & miRsAux!HorasTrabajadas & "',''"
            
            Cad = Cad & "," & miRsAux!idTrabajador & "," & vUsu.Codigo & "," & DBSet(miRsAux!Fecha, "F") & "," & NumRegElim
            vHora = 0
        End If
        
        
        If vHora < 8 Then   'solo ionserto 16
                
    
            If miRsAux!HorasNegativas = 1 Then
                FueraIntervalo2 = 1
                I = 24 - miRsAux!LaHora
            Else
                If miRsAux!LaHora >= 0 And miRsAux!LaHora <= 23 Then
                    I = miRsAux!LaHora
                    FueraIntervalo2 = 0
                Else
                    FueraIntervalo2 = 2
                    I = miRsAux!LaHora - 24
                    
                    
                    
                    
                End If
            End If
            CadPa = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
            Cad = Cad & ",'" & CadPa & "'"
            
            
            

        End If
        vHora = vHora + 1
        miRsAux.MoveNext
   Wend
   miRsAux.Close
   Set miRsAux = Nothing
   Set vH = Nothing
   If T > 0 Then InsertaProcesarHco vHora, False
        
   If NumRegElim > 0 Then ImprimirDatosHco = True
        
   
   
   Exit Function
eImprimirDatosHco:
    MuestraError Err.Number, Err.Description
   
End Function


Private Sub InsertaProcesarHco(NTicajes As Integer, Insertar As Boolean)
Dim J As Integer

       
        
        J = NTicajes
        While J < 8
            Cad = Cad & ",NULL"
            J = J + 1
        Wend
        
        Cad = Cad & "," & Abs(NTicajes > 8) & ")"
                
        
        
        vSQL = vSQL & ", " & Cad
        Cad = ""
        
        
        If Not Insertar Then
            If Len(vSQL) > 10000 Then Insertar = True
        End If
        
        If Insertar Then
            vSQL = Mid(vSQL, 2)
            Cad = "INSERT tmppresencia(NomTrabajador,Incidencias,NomEmpresa,Seccion,idtra,codusu,Fecha,id,H1,H2,H3,H4,H5,H6,H7,H8,semana) VALUES " & vSQL
            conn.Execute Cad
            vSQL = ""
        End If
        Cad = ""
End Sub

