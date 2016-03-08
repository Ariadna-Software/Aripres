VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmpleado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empleados"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "frmEmpleado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame Frame7 
      Caption         =   "Observaciones"
      ForeColor       =   &H00972E0B&
      Height          =   1335
      Left            =   240
      TabIndex        =   86
      Top             =   5895
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   975
         Index           =   14
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Tag             =   "Observaciones|T|S|||empleado|observac|||"
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Datos Laborales"
      ForeColor       =   &H00972E0B&
      Height          =   3495
      Left            =   5880
      TabIndex        =   75
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   2760
         TabIndex        =   89
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   17
         Tag             =   "Código empresa|N|N|0|999|empleado|codempss|000||"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   31
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   26
         Tag             =   "Categoria profesional|T|S|||empleado|catprofe|||"
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   25
         Tag             =   "Tipo empleado|N|N|0|99|empleado|tipemple|00||"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   30
         Left            =   2760
         TabIndex        =   84
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "Fecha fin contrato actual|F|S|||empleado|fechafin|dd/mm/yyyy||"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "Fecha inicio contrato actual|F|S|||empleado|fechactu|dd/mm/yyyy||"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Fecha primer contrato|F|S|||empleado|fechprim|dd/mm/yyyy||"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   21
         Tag             =   "Modalidad contrato|T|S|||empleado|modcontr|||"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   20
         Tag             =   "Tipo contrato|T|S|||empleado|tipcontr|||"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   19
         Tag             =   "Tipo de nómina|N|S|0|99|empleado|tipnomin|00||"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   2760
         TabIndex        =   77
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Código Seguridad Social|T|S|||empleado|codsegso|||"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1560
         MouseIcon       =   "frmEmpleado.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar empresa"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa alta"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   90
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria profesional"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   88
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo empleado"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   85
         Top             =   2760
         Width           =   1275
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1560
         MouseIcon       =   "frmEmpleado.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo de empleado"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "(a)"
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   83
         Top             =   2400
         Width           =   195
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   29
         Left            =   3720
         Picture         =   "frmEmpleado.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Contr. actual (de)"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   82
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   28
         Left            =   1560
         Picture         =   "frmEmpleado.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Primer contrato"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   81
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   27
         Left            =   1560
         Picture         =   "frmEmpleado.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Modalidad contrato"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   80
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo contrato"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   79
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de nómina"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   78
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1560
         MouseIcon       =   "frmEmpleado.frx":0451
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo de nómina"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Cód. Seguridad Social"
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Datos Identificación"
      ForeColor       =   &H00972E0B&
      Height          =   1160
      Left            =   240
      TabIndex        =   60
      Top             =   4710
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "frmEmpleado.frx":05A3
         Left            =   3840
         List            =   "frmEmpleado.frx":05AD
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "Permiso|N|N|1|2|empleado|permiweb|0||"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   13
         Left            =   1100
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   14
         Tag             =   "Contraseña|T|N|||empleado|passwweb|||"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1100
         MaxLength       =   20
         TabIndex        =   13
         Tag             =   "Usuario|T|N|||empleado|loginweb|||"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Calidad"
         Height          =   255
         Left            =   3840
         TabIndex        =   63
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Contraseña"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos Bancarios"
      ForeColor       =   &H00972E0B&
      Height          =   2520
      Left            =   5880
      TabIndex        =   47
      Top             =   4710
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   32
         Left            =   2120
         MaxLength       =   2
         TabIndex        =   29
         Tag             =   "Código IBAN Dígito de Control|T|S|||empleado|ibandctl|||"
         ToolTipText     =   "Código IBAN Dígito de Control"
         Top             =   674
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         ItemData        =   "frmEmpleado.frx":05C3
         Left            =   1420
         List            =   "frmEmpleado.frx":05C5
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Tag             =   "Código IBAN pais|N|N|0|9999|empleado|codnacio|0000||"
         ToolTipText     =   "Código IBAN pais"
         Top             =   674
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   36
         Tag             =   "Código Postal banco|T|S|||empleado|postbanc|||"
         Top             =   2120
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1420
         MaxLength       =   6
         TabIndex        =   35
         Tag             =   "Población banco|N|S|0|999999|empleado|pobbanco|000000||"
         Top             =   1766
         Width           =   750
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2280
         TabIndex        =   72
         Top             =   1766
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   34
         Tag             =   "Domicilio banco|T|S|||empleado|domibanc|||"
         Top             =   1412
         Width           =   4215
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   1080
         TabIndex        =   69
         Top             =   1058
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   2580
         MaxLength       =   4
         TabIndex        =   30
         Tag             =   "Entidad|N|S|0|9999|empleado|codbanco|0000||"
         Text            =   "99"
         ToolTipText     =   "Entidad"
         Top             =   674
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   3160
         MaxLength       =   4
         TabIndex        =   31
         Tag             =   "Oficina|N|S|0|9999|empleado|codsucur|0000||"
         Text            =   "9999"
         ToolTipText     =   "Oficina"
         Top             =   674
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   3740
         MaxLength       =   2
         TabIndex        =   32
         Tag             =   "Digito Control|T|S|0|99|empleado|digcontr|00||"
         Text            =   "99"
         ToolTipText     =   "Dígito de Control"
         Top             =   674
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "Número de Cuenta|T|S|||empleado|ctabanco|0000000000||"
         Text            =   "999999999"
         ToolTipText     =   "Número de Cuenta"
         Top             =   674
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         ItemData        =   "frmEmpleado.frx":05C7
         Left            =   3840
         List            =   "frmEmpleado.frx":05D1
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "Cobro|H|N|1|2|empleado|formcobr|0||"
         Top             =   2120
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1420
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "Cuenta Contable|T|S|||empleado|ctaemple|||"
         Top             =   320
         Width           =   1095
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2595
         TabIndex        =   64
         Top             =   320
         Width           =   2700
      End
      Begin VB.Label Label13 
         Caption         =   "C.P."
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   2120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   73
         Top             =   1766
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1080
         MouseIcon       =   "frmEmpleado.frx":05E7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar población"
         Top             =   1766
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   71
         Top             =   1412
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   70
         Top             =   1058
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "C.Bancaria"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   68
         Top             =   674
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1080
         MouseIcon       =   "frmEmpleado.frx":0739
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta bancaria"
         Top             =   674
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Cobro"
         Height          =   255
         Left            =   3240
         TabIndex        =   67
         Top             =   2120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "C.Contable"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   65
         Top             =   320
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1080
         MouseIcon       =   "frmEmpleado.frx":088B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   320
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos Básicos"
      ForeColor       =   &H00972E0B&
      Height          =   3495
      Left            =   240
      TabIndex        =   46
      Top             =   1200
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "frmEmpleado.frx":09DD
         Left            =   1080
         List            =   "frmEmpleado.frx":09E7
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Situación|N|N|1|2|empleado|situacio|0||"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "E-mail|T|S|||empleado|mailempl|||"
         Top             =   2520
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   10
         Tag             =   "Móvil|T|S|||empleado|tfnmovil|||"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Teléfono|T|S|||empleado|telefono|||"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "NIF|T|S|||empleado|nifemple|||"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Código de Población|N|S|0|999999|empleado|codpobla|000000||"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   53
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "Código Postal|T|S|||empleado|codposta|||"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Domicilio|T|S|||empleado|domemple|||"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "Código de Agencia|N|N|0|999|empleado|codagenc|000||"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   50
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "Código de Empresa|N|N|0|999|empleado|codempre|000||"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   48
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label11 
         Caption         =   "Situación"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "E-mail"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   59
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Móvil"
         Height          =   255
         Left            =   3000
         TabIndex        =   58
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "NIF"
         Height          =   255
         Left            =   3000
         TabIndex        =   56
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1080
         MouseIcon       =   "frmEmpleado.frx":09FD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar población"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label28 
         Caption         =   "C.P."
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   52
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1080
         MouseIcon       =   "frmEmpleado.frx":0B4F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar agencia"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         MouseIcon       =   "frmEmpleado.frx":0CA1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar empresa"
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Index           =   0
      Left            =   240
      TabIndex        =   42
      Top             =   480
      Width           =   11295
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   600
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Código de Empleado|N|N|1|9999|empleado|codemple|0000|S|"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "Nombre|T|N|||empleado|nomemple|||"
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Apellidos|T|N|||empleado|apeemple|||"
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre "
         Height          =   255
         Left            =   6720
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellidos "
         Height          =   255
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cód."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   40
      Top             =   7200
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10380
      TabIndex        =   39
      Top             =   7320
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   38
      Top             =   7320
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4080
      Top             =   7320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10380
      TabIndex        =   87
      Top             =   7320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   92
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: CÈSAR +-+-
' +-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els TAGs
' 3. Posar els MAXLENGTHs
' 4. Posar els TABINDEXs
'
' +-+-+-+-+- CODIFICACIÓ +-+-+-+-+-
' 1. Definir variables per a cridar a atres formularis
' 2. En Form_Load() canviar el nom de la taula i la clau primaria de l'ORDER BY
' 3. En PonerModo Revisar lo que bloquejem, el nom de la clau primaria,
'    les imagens de buscar i les de dates
' 4. En PonerLongCampos() posar només els camps numérics
' 5. En PonerModoOpcionesMenu(Modo) comentar o descomentar depenent
'    de si n'hi ha menú desplegable o no
' 6. (SI N'HI HA BUSCAR DATA) En imgFec_Click(Index As Integer) i en
'    frmC_Selec(vFecha As Date), canviar l'index de imgFec pel 1r index de les
'    imagens de buscar data
' 7. (SI N'HI HAN CAMPS DE BUSCAR CODIS) En imgBuscar_Click(Index As Integer)
'    codificar tots els camps. Per a cada camp fer la funció, per eixemple,
'    frmPob_DatoSeleccionado(CadenaSeleccion As String)
' 8. En MandaBusquedaPrevia(CadB As String) arreglar-ho per a vore lo que es desije
' 9. En BotonAnyadir() canviar el nom de la taula i el nom de la clau primaria
' 10. En BotonEliminar() canviar els noms, els formats i el DELETE
' 11. Si alguna atra taula apunta a la actual, en SePuedeEliminar()
'     canviar els parametres de la funció; sino, comentar-ho tot
' 12. (SI N'HI HAN CAMPS DE BUSCAR CODIS) En PonerCampos() configurar els camps
'     de les descripcions
' 13. En DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
' 14. En Eliminar() canviar el nom de la clau primaria
' 15. (SI N'HI HAN COMBOS) En CargaCombo() configurar els distints Combos
'
'
' 30. Si el nom del camp que te la clau primaria NO es Text1(0), canviar-ho en:
'     BotonBuscar(), HacerBusqueda(), BotonAnyadir(), BotonModificar(), CmdCancelar_Click()
' *******************************************************************************

Option Explicit

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)

' ****** Definir variables per a cridar a atres formularis *********
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmEmp As frmEmpresas
Attribute frmEmp.VB_VarHelpID = -1
Private WithEvents frmAge As frmAgencias2
Attribute frmAge.VB_VarHelpID = -1
Private WithEvents frmPob As frmPoblacio
Attribute frmPob.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmBan As frmBancsofi
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmTiN As frmTiponomi
Attribute frmTiN.VB_VarHelpID = -1
Private WithEvents frmTiE As frmTiposemp
Attribute frmTiE.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
' *****************************************************************

Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'***Variables comuns a tots els formularis*****
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
Dim Indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos
Dim cad_meua As String
Dim RS As ADODB.Recordset
Dim CadB As String


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    'situarnos en el registro que acabamos de insertar
                    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE codemple=" & Text1(0).Text & Ordenacion
                    PosicionarData
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

' *** adrede per ad este manteniment ***
Private Sub Combo1_Click(Index As Integer)
    Text1_LostFocus (16) 'el camp que te el codi del banc
End Sub
'***************************************+

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True

    ' ICONITOS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Todos
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
      
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
      
    LimpiarCampos   'Limpia los campos TextBox
    
    ' *** canviar el nom de la taula i la clau primaria de l'ORDER BY ***
    NombreTabla = "empleado"
    Ordenacion = " ORDER BY codemple"
    ' **********************************************************
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codemple=-1"
    Data1.Refresh
          
    For i = 0 To Combo1.Count - 1
        CargaCombo i
    Next i
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow 'codclien
    End If
End Sub


Private Sub LimpiarCampos()
Dim i As Integer
    
    On Error Resume Next

    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    
    For i = 0 To Me.Combo1.Count - 1
        Me.Combo1(i).ListIndex = -1
    Next i
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualiza Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    'PonerIndicador lblIndicador, Modo, ModoLineas
    PonerIndicador lblIndicador, Modo
       
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = (Modo = 2)
    Else
        cmdRegresar.Visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = b
    cmdAceptar.Visible = b
        
    'Bloquea los campos Text1 si no estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    
    ' **************************************************************************************************
    ' *** Revisar lo que bloquejem, el nom de la clau primaria, les imagens de buscar i les de dates ***
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    PosicionarCombo Combo1(2), 724
'    For i = 0 To Combo1(2).ListCount - 1
'        If Combo1(2).ItemData(i) = 724 Then
'            Combo1(2).ListIndex = i
'            Exit For
'        End If
'    Next i
    
    If Modo = 4 Then _
        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    
    BloquearImgBuscar Me, Modo
    BloquearImgFec Me, 27, Modo
    BloquearImgFec Me, 28, Modo
    BloquearImgFec Me, 29, Modo
    ' **************************************************************************************************
                
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los Text1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
'Activa/desact. las Opciones de Menu y Toolbar según permisos de usuario
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Activa/desact. las Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
   
' ******** comentar o descomentar depenent de si n'hi ha menú desplegable o no ****
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    'Insertar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
   
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = b Or (Modo = 0)
    Toolbar1.Buttons(12).Enabled = b
' ********************************************************************************
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        ' *** canviar o llevar el WHERE ***
        'CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object
       
    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend

    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    ' es desplega baix i cap a la dreta
    'frmC.Left = esq + imgFec(Index).Parent.Left + 30
    'frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left - frmC.Width + imgFec(Index).Width + 40
    frmC.Top = dalt + imgFec(Index).Parent.Top - frmC.Height + menu - 25
    
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(27).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If Text1(Index).Text <> "" Then frmC.NovaData = Text1(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco Text1(CByte(imgFec(27).Tag))
    ' **************************************************************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CByte(imgFec(27).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
'    PonerFoco txtAux(CByte(imgFec(27).Tag))
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    'Screen.MousePointer = vbHourglass
    TerminaBloquear
    Select Case Index
        Case 0, 8 'empresa
            If Index = 8 Then
                Indice = 33
            Else
                Indice = 3
            End If
            Set frmEmp = New frmEmpresas
            frmEmp.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
            frmEmp.CodigoActual = Text1(Indice).Text
            frmEmp.Show vbModal
            Set frmEmp = Nothing
            PonerFoco Text1(Indice)
            
        Case 1 'agencia
            Indice = 4
            Set frmAge = New frmAgencias2
            frmAge.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
            frmAge.DeConsulta = True
            frmAge.Empresa = Text1(3).Text
            frmAge.CodigoActual = Text1(4).Text
            frmAge.Show vbModal
            Set frmAge = Nothing
            PonerFoco Text1(Indice)
            
        Case 2, 5 'población i población banco
            If Index = 2 Then
                Indice = 6
            ElseIf Index = 5 Then
                Indice = 21
            End If
            Set frmPob = New frmPoblacio
            frmPob.DatosADevolverBusqueda = "0|1|2|3|4|5|"
            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
            frmPob.CodigoActual = Text1(Indice).Text
            frmPob.Show vbModal
            Set frmPob = Nothing
            PonerFoco Text1(Indice)
            
        Case 3 'Cuenta Contable
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(Text1(15).Text) Then Text1(15).Text = ""
            frmCtas.CodigoActual = Text1(15).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(15)
            
        Case 4 'Cuenta Bancaria
            Set frmBan = New frmBancsofi
            frmBan.DatosADevolverBusqueda = "4|1|3|"
            frmBan.CodigoActual = Text1(16).Text
            frmBan.NuevoPais = Me.Combo1(2).ItemData(Combo1(2).ListIndex)
            frmBan.Show vbModal
            Set frmBan = Nothing
            PonerFoco Text1(16)
            
        Case 6 'tipo de nómina
            Set frmTiN = New frmTiponomi
            frmTiN.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(Text1(24).Text) Then Text1(24).Text = ""
            frmTiN.CodigoActual = Text1(24).Text
            frmTiN.Show vbModal
            Set frmTiN = Nothing
            PonerFoco Text1(24)
            
        Case 7 'tipo de empleado
            Set frmTiE = New frmTiposemp
            frmTiE.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(Text1(30).Text) Then Text1(30).Text = ""
            frmTiE.CodigoActual = Text1(30).Text
            frmTiE.Show vbModal
            Set frmTiE = Nothing
            PonerFoco Text1(30)
    End Select

    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmPob_DatoSeleccionado(CadenaSeleccion As String)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codpobla
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'despobla
    Text1(Indice + 1).Text = RecuperaValor(CadenaSeleccion, 3) 'codposta
    If (Indice = 6) And (Text1(9).Text = "") Then _
        Text1(9).Text = RecuperaValor(CadenaSeleccion, 6) 'prefix
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    Text1(15).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(15).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmEmp_DatoSeleccionado(CadenaSeleccion As String)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAge_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(4)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
Dim valor As String
'Dim i As Integer

    valor = RecuperaValor(CadenaSeleccion, 1)
    
    PosicionarCombo Combo1(2), Val(valor)
'    For i = 0 To Combo1(2).ListCount - 1
'        If Combo1(2).ItemData(i) = Val(valor) Then
'            Combo1(2).ListIndex = i
'            Exit For
'        End If
'    Next i
    
    Text1(16).Text = RecuperaValor(CadenaSeleccion, 2)
    FormateaCampo Text1(16)
    Text2(16).Text = RecuperaValor(CadenaSeleccion, 3)
    If Text2(16).Text = "" Then
        Text2(16).Text = DevuelveDesdeBDNew(cPTours, "bancsofi", "nombanco", "codnacio", valor, "N", , "codbanco", Text1(16).Text, "N")
    End If
End Sub

Private Sub frmTiN_DatoSeleccionado(CadenaSeleccion As String)
    Text1(24).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(24)
    Text2(24).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTiE_DatoSeleccionado(CadenaSeleccion As String)
    Text1(30).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(30)
    Text2(30).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Buscar
           mnBuscar_Click
        Case 4  'Todos
            mnVerTodos_Click
        Case 7  'Nuevo
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 12 'Imprimir
        '    AbrirListado (10)
            'BotonImprimir
            printNou
'            MsgBox "Falta fer el llista d'empleats"
        Case 13    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
' *************** Si la clau primaria no es Text1(0), canviar-ho ***************
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' ******** Si la clau primaria no es Text1(0), canviar-ho ***********
        PonerFoco Text1(0)
        ' *******************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim cad As String
        'Llamamos a al form
        ' **************** arreglar-ho per a vore lo que es desije ****************
        cad = ""
        cad = cad & ParaGrid(Text1(0), 10, "Cód.")
        cad = cad & ParaGrid(Text1(1), 50, "Apellidos")
        cad = cad & ParaGrid(Text1(2), 40, "Nombre")
        If cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Empleados"
            frmB.vSelElem = 0

            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
            End If
        End If
        ' *************************************************************************
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim j As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    i = 0
    Do
        j = i + 1
        i = InStr(j, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, j, i - j)
            j = Val(Aux)
            cad = cad & Text1(j).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonVerTodos()
    CadB = ""
    LimpiarCampos 'Limpia los Text1
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
    LimpiarCampos 'Vacía los TextBox
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    ' ******* Canviar el nom de la taula, el nom de la clau primaria, i el
    ' nom del camp que te la clau primaria si no es Text1(0) *************
    Text1(0).Text = SugerirCodigoSiguienteStr("empleado", "codemple")
    FormateaCampo Text1(0)
    
    'empresa
    Text1(3).Text = vSesion.Empresa
    Text2(3).Text = PonerNombreDeCod(Text1(3), "empresas", "nomempre", "codempre", "N")
    
    'empresa alta
    Text1(33).Text = vSesion.Empresa
    Text2(33).Text = PonerNombreDeCod(Text1(33), "empresas", "nomempre", "codempre", "N")
    
    PonerFoco Text1(0)
    ' ********************************************************************
End Sub


Private Sub BotonModificar()
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    ' *** Canviar el nom del camp que te la clau primaria si no es Text1(0) ***
'    BloquearTxt Text1(0), True
    PonerFoco Text1(1)
    ' *************************************************************************
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    'Comprobamos si se puede eliminar
'    If Not SePuedeEliminar Then Exit Sub

    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub

    ' *************** canviar els noms, els formats i el DELETE ****************                  "
    cad = cad & "¿Seguro que desea eliminar el Empleado?"
    cad = cad & vbCrLf & "  Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "  Apellidos: " & Data1.Recordset.Fields(2)
    cad = cad & vbCrLf & "  Nombre: " & Data1.Recordset.Fields(1)
    
    ' **************************************************************************
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Proveedor", Err.Description
End Sub


Private Sub PonerCampos()

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
       
    ' ************* configurar els camps de les descripcions *************
    Text2(3).Text = PonerNombreDeCod(Text1(3), "empresas", "nomempre", "codempre", "N")
    Text2(33).Text = PonerNombreDeCod(Text1(33), "empresas", "nomempre", "codempre", "N")
    Text2(4).Text = PonerNombreDeCod(Text1(4), "agencias", "desagenc", "codagenc", "N")
    Text2(4).Text = DevuelveDesdeBDNew(cPTours, "agencias", "desagenc", "codempre", Text1(3).Text, "N", , "codagenc", Text1(4).Text, "N")
    Text2(6).Text = PonerNombreDeCod(Text1(6), "poblacio", "despobla", "codpobla", "N")
    Text2(15).Text = PonerNombreCuenta(Text1(15))
    If Combo1(2).ListIndex <> -1 Then
        Text2(16).Text = DevuelveDesdeBDNew(cPTours, "bancsofi", "nombanco", "codnacio", Combo1(2).ItemData(Combo1(2).ListIndex), "N", , "codbanco", Text1(16).Text, "N")
    End If
    Text2(21).Text = PonerNombreDeCod(Text1(21), "poblacio", "despobla", "codpobla", "N")
    Text2(24).Text = PonerNombreDeCod(Text1(24), "tiponomi", "desnomin", "tipnomin", "N")
    Text2(30).Text = PonerNombreDeCod(Text1(30), "tiposemp", "desemple", "tipemple", "N")
    ' *******************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = PonerContRegistros(Me.Data1)
End Sub


Private Sub cmdCancelar_Click()

    ' *** canviar el nom del camp que te la clau primaria si no es Text1(0) ***
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            If Data1.Recordset.EOF Then
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
            PonerFoco Text1(0)

        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select
    ' *************************************************************************
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me)
    If Not b Then Exit Function
    
    ' ******************** canviar els arguments de la funcio i el mensage ****************
    If (Modo = 3) Then 'Insertar
        'comprobar si ya existe el campo de clave primaria
        If ExisteCP(Text1(0)) Then b = False
        
'         Datos = DevuelveDesdeBD("codemple", "empleado", "codemple", Text1(0).Text, "N")
'         If Datos <> "" Then
'            MsgBox "Ya existe el Código de Empleado: " & Text1(0).Text, vbExclamation
'            DatosOk = False
'            PonerFoco Text1(0)
'            Exit Function
'         End If
    End If
    ' *************************************************************************************
         
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per la clua primaria ***
    cad = "(codemple=" & Text1(0).Text & ")"
    ' ***************************************
    
    If SituarData(Data1, cad, Indicador) Then
       PonerModo 2
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

'    Conn.BeginTrans
    ' ***** canviar el nom de la clau primaria *******
    vWhere = " WHERE codemple=" & Data1.Recordset!CodEmple
    ' ************************************************
              
    Conn.Execute "Delete from " & NombreTabla & vWhere
               
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
'        Conn.RollbackTrans
        Eliminar = False
    Else
'        Conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Not Text1(Index).MultiLine Then ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ***************** configurar els camps de buscar codis *****************
    Select Case Index
        Case 0, 17, 18, 19
            PonerFormatoEntero Text1(Index)

        Case 8 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
            
        Case 9, 10 'telèfons, fax i mòbils
            PosarFormatTelefon Text1(Index)
            
        Case 12, 13 'loginweb i passwweb
            Text1(Index).Text = LCase(Text1(Index).Text)
            If Index = 12 Then 'login
                If Not ComprobarLoginEmp(Text1(Index).Text) Then PonerFoco Text1(Index)
            End If
            
        Case 6, 21 'poblacion
            Nuevo = False
            If Index = 6 Then
                PonerDatosPoblacion Text1(Index), Text2(Index), Text1(Index + 1), , , Nuevo, Text1(9)
            ElseIf Index = 21 Then
                PonerDatosPoblacion Text1(Index), Text2(Index), Text1(Index + 1), , , Nuevo
            End If
            If Nuevo Then
                Indice = Index
                Set frmPob = New frmPoblacio
                frmPob.DatosADevolverBusqueda = "0|1|2|3|4|5|"
                frmPob.NuevoCodigo = Text1(Index).Text
                Text1(Index).Text = ""
                TerminaBloquear
                frmPob.Show vbModal
                Set frmPob = Nothing
                If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
            End If

        Case 3, 33 'empresa
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "empresas", "nomempre", "codempre", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Empresa: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmEmp = New frmEmpresas
                        frmEmp.DatosADevolverBusqueda = "0|1|"
                        frmEmp.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmEmp.Show vbModal
                        Set frmEmp = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
            If Index = 3 And Text1(4).Text <> "" Then Text1_LostFocus (4)

        Case 4 'agencia
            If PonerFormatoEntero(Text1(Index)) Then
                cadMen = Text1(3).Text 'empresa
                If (cadMen = "" Or Not IsNumeric(cadMen)) Then
                    Text1(4).Text = ""
                    Text2(4).Text = ""
                    Exit Sub
                End If

                Text2(Index).Text = DevuelveDesdeBDNew(cPTours, "agencias", "desagenc", "codempre", cadMen, "N", , "codagenc", Text1(Index).Text, "N")
                FormateaCampo Text1(Index)
                If Text2(Index).Text = "" And Text1(Index) <> "" Then
                    cadMen = "No existe la Agencia: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "para la Empresa: " & Text1(3).Text & "  " & Text2(3).Text & vbCrLf
                    MsgBox cadMen, vbExclamation
'                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmAge = New frmAgencias
'                        frmAge.DatosADevolverBusqueda = "0|1|"
'                        frmAge.NuevoCodigo = text1(Index).Text
'                        text1(Index).Text = ""
'                        TerminaBloquear
'                        frmAge.Show vbModal
'                        Set frmAge = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        text1(Index).Text = ""
'                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 15 'Cuenta Contable
            If Text1(Index).Text = "" Then
                Text2(Index).Text = ""
                Exit Sub
            End If
            If Modo = 3 And ContieneCaracterBusqueda(Text1(Index).Text) Then Exit Sub     'Busquedas
            Text2(Index).Text = PonerNombreCuenta(Text1(Index))
                
        Case 16 'Cuenta Bancaria
            If PonerFormatoEntero(Text1(Index)) Then
                If Text1(Index).Text = "" Then Exit Sub
                Text2(Index).Text = DevuelveDesdeBDNew(cPTours, "bancsofi", "nombanco", "codnacio", Combo1(2).ItemData(Combo1(2).ListIndex), "N", , "codbanco", Text1(Index).Text, "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Banco: " & Text1(Index).Text & "  "
                    cadMen = cadMen & "para el pais: " & Combo1(2).List(Combo1(2).ListIndex) & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmBan = New frmBancsofi
                        frmBan.DatosADevolverBusqueda = "0|1|"
                        frmBan.NuevoCodigo = Text1(Index).Text
                        frmBan.NuevoPais = Combo1(2).ItemData(Combo1(2).ListIndex)
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmBan.Show vbModal
                        Set frmBan = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If

        Case 24 'Tipo de Nómina
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiponomi", "desnomin")
                If Text2(Index).Text = "" And Text1(Index) <> "" Then
                    cadMen = "No existe el Tipo de Nómina: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTiN = New frmTiponomi
                        frmTiN.DatosADevolverBusqueda = "0|1|"
                        frmTiN.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTiN.Show vbModal
                        Set frmTiN = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 30 'Tipo de Empleado
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiposemp", "desemple")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Empleado: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTiE = New frmTiposemp
                        frmTiE.DatosADevolverBusqueda = "0|1|"
                        frmTiE.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTiE.Show vbModal
                        Set frmTiE = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 27, 28, 29 'dates
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
    End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(0).BackColor = vbYellow Then Combo1(0).BackColor = vbWhite
End Sub

Private Sub text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text1(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 1 Or Modo = 3 Or Modo = 4 Then
                Select Case Index
                    Case 3: KEYBusqueda KeyAscii, 0 'empresa
                    Case 4: KEYBusqueda KeyAscii, 1 'agencia
                    Case 6: KEYBusqueda KeyAscii, 2 'poblacion
                    Case 33: KEYBusqueda KeyAscii, 8 'empresa alta
                    Case 24: KEYBusqueda KeyAscii, 6 'tipo de nomina
                    Case 30: KEYBusqueda KeyAscii, 7 'tipo de empleado
                    Case 15: KEYBusqueda KeyAscii, 3 'cuenta contable
                    Case 16: KEYBusqueda KeyAscii, 4 'banco oficial
                    Case 21: KEYBusqueda KeyAscii, 5 'poblacion
                    
                    Case 27: KEYFecha KeyAscii, 27
                    Case 28: KEYFecha KeyAscii, 28
                    Case 29: KEYFecha KeyAscii, 29
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not Text1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub
Private Sub CargaCombo(Index As Integer)
Dim i, F, j As Integer

    Combo1(Index).Clear

    ' ******* configurar els distints Combos **********
    Select Case Index
        Case 0 'situación
            Combo1(Index).AddItem "Activo"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 1
        
            Combo1(Index).AddItem "Inactivo"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 2
            
        Case 1 'permiso Gestión de Calidad
            Combo1(Index).AddItem "Normal"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 1
        
            Combo1(Index).AddItem "Administrador"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 2
            
        Case 2 'ibanpais
            cad_meua = "SELECT * FROM naciones WHERE ibanpais <> """" ORDER BY ibanpais"
            Set RS = New ADODB.Recordset
            RS.Open cad_meua, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS.EOF
                Combo1(Index).AddItem RS!ibanPais
                Combo1(Index).ItemData(Combo1(Index).NewIndex) = RS!codNacio
                RS.MoveNext
            Wend
            
        Case 3 'forma de cobro
            Combo1(Index).AddItem "Talón"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 1
        
            Combo1(Index).AddItem "Transferencia"
            Combo1(Index).ItemData(Combo1(Index).NewIndex) = 2
    End Select
    
    If Index <> 2 Then 'excepte per a ibanpais sempre seleccione la 1ª opció per defecte
        Combo1(Index).ListIndex = 0
    Else
        'per defecte seleccione ES
        PosicionarCombo Combo1(Index), 724
'        For j = 0 To Combo1(Index).ListCount - 1
'            If Combo1(Index).ItemData(j) = 724 Then
'                Combo1(Index).ListIndex = j
'                Exit For
'            End If
'        Next j
    End If
    ' **************************************************************
End Sub


Private Sub BotonImprimir()
'Dim cadParam As String
'Dim cadFormula As String
'
'    'Añadir el parametro de Empresa
'    cadParam = "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
'    'selecciona solo los empleados de esa empresa
'    cadFormula = "{empleado.codempre}=" & vEmpresa.codEmpre
'    With frmImprimir
'        .FormulaSeleccion = cadFormula
'        .OtrosParametros = cadParam
'        .NumeroParametros = 1 'Solo parametro de la empresa
'        .SoloImprimir = False
'        .Opcion = 13 'Opcionlistado
'        .Show vbModal
'    End With
End Sub


Private Sub printNou()
    
    With frmImprimir2
        .cadTabla2 = "empleado"
        .Informe2 = "rEmpleados.rpt"
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomEmpre & "'|pOrden={empleado.apeemple}|"
        .NumeroParametros2 = 2
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        
        .Show vbModal
    End With
End Sub



Private Function ComprobarLoginEmp(cadLogin As String) As Boolean
Dim SQL As String
    
    On Error GoTo ECompLogin
    
    SQL = DevuelveDesdeBDNew(cPTours, "empleado", "codemple", "loginweb", cadLogin, "T")
    
    If SQL <> "" Then
        If CInt(SQL) <> CInt(Text1(0).Text) Then
            'existe ya un usuario con ese valor
            SQL = "Ya existe un empleado con el login " & cadLogin & vbCrLf
            MsgBox SQL, vbExclamation
            ComprobarLoginEmp = False
        Else
            ComprobarLoginEmp = True
        End If
    Else
        ComprobarLoginEmp = True
    End If
    
    Exit Function
    
ECompLogin:
    MuestraError Err.Number, "", Err.Description
End Function
