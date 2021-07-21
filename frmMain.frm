VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00CCD1D0&
   Caption         =   "A R I P R E S     4 "
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   13575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Picture         =   "frmMain.frx":FC8A
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1440
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A7CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4102C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":415C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E68A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F564
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5043E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":509D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5723A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":599EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A6C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FEB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6193A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":67EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69CA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":70504
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5520
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmMain.frx":76D66
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:20"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Trabajadores"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calendarios"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Horarios"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Revisar marcajes"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Revisar busq."
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Operaciones TCP3"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar marcaje actual"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Procesar marcajes"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Presencia"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Horas trabajadas"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Traer datos maquinas"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   2460
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A328
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A386
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A3E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A442
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A4A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A4FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A55C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A5BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A618
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A676
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A6D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A732
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A790
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A7EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A84C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A8AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A908
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A966
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A9C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AA22
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AA80
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList11 
      Left            =   840
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C7E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7CB02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":831F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":83510
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":843EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":852C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B55E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B878
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9149A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":92174
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":941F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":94510
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9482A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AAC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B99E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9C878
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList22 
      Left            =   9600
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9CB92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DC24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9ECB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9FD48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A0DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A1E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A2EFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A3F90
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA7F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AB884
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AC916
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B3178
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnDatos 
      Caption         =   "&Datos básicos"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnHorarios 
         Caption         =   "&Horarios"
      End
      Begin VB.Menu mnCalendarios 
         Caption         =   "&Calendarios"
      End
      Begin VB.Menu mnIncidencias 
         Caption         =   "&Incidencias"
      End
      Begin VB.Menu mnsecciones 
         Caption         =   "&Secciones"
      End
      Begin VB.Menu mnCategorias 
         Caption         =   "&Categorias"
      End
      Begin VB.Menu mnTrabajadores 
         Caption         =   "&Trabajadores"
      End
      Begin VB.Menu mnTareas 
         Caption         =   "Tareas"
      End
      Begin VB.Menu mnZonas 
         Caption         =   "Areas-Zonas"
      End
      Begin VB.Menu mnTerminales 
         Caption         =   "Terminales"
      End
      Begin VB.Menu mnVacaciones 
         Caption         =   "Vacaciones"
      End
      Begin VB.Menu mnbarr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSelecImpresora 
         Caption         =   "Seleccionar impresora"
      End
      Begin VB.Menu mn_barra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnEmpresas 
         Caption         =   "&Empresa"
      End
      Begin VB.Menu mnBancos 
         Caption         =   "Bancos"
      End
      Begin VB.Menu mnMantenUsuarios 
         Caption         =   "Mantenimiento de Usuarios"
      End
      Begin VB.Menu mnBarra29 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnOperaciones 
      Caption         =   "&Operaciones"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnVerMarcajes 
         Caption         =   "&Marcajes"
      End
      Begin VB.Menu mnHistoricoMarcajes 
         Caption         =   "Histórico de marcajes"
      End
      Begin VB.Menu mnbarra2_10 
         Caption         =   "-"
      End
      Begin VB.Menu mnResumenMarcajes 
         Caption         =   "Resumen marcajes"
      End
      Begin VB.Menu mnTicajeActual 
         Caption         =   "Consultar marcaje actual"
      End
      Begin VB.Menu mnbarra2_5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnCambioHorarioMenu 
         Caption         =   "Cambios horario"
         Visible         =   0   'False
         Begin VB.Menu mnCabioHorario 
            Caption         =   "Masivo"
         End
         Begin VB.Menu mnCambioHorarioAjuste 
            Caption         =   "Ajustes"
         End
      End
      Begin VB.Menu mnbarra2_8 
         Caption         =   "-"
      End
      Begin VB.Menu mnRelojesAuxiliares 
         Caption         =   "Marcajes relojes auxiliares"
      End
      Begin VB.Menu mnbarra2_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnProcesoPrevioALZ 
         Caption         =   "Proceso previo importación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnProcesar 
         Caption         =   "Procesar marcajes"
      End
      Begin VB.Menu mnImportar 
         Caption         =   "Importar &fichero de datos"
      End
      Begin VB.Menu mnbarra2_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnLecturaReloj 
         Caption         =   "Leer TCP3"
         Index           =   0
      End
      Begin VB.Menu mnLecturaReloj 
         Caption         =   "Relojes Kreta"
         Index           =   1
      End
      Begin VB.Menu mnLecturaReloj 
         Caption         =   "Fingkey Access"
         Index           =   2
      End
      Begin VB.Menu mnLecturaReloj 
         Caption         =   "Relojes ZKTeco"
         Index           =   3
      End
      Begin VB.Menu mnLecturaReloj 
         Caption         =   "Biostar1  / Xpass"
         Index           =   4
      End
      Begin VB.Menu mnLecturaReloj2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnLecturaReloj2 
         Caption         =   "Lectura "
         Index           =   1
      End
   End
   Begin VB.Menu mnLaboral 
      Caption         =   "&Laboral"
      Begin VB.Menu mnLaboral1 
         Caption         =   "Horas"
         Index           =   0
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "Horas procesadas"
            Index           =   0
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "Proceso cálculo horas"
            Index           =   1
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "Ver datos mes trabajado "
            Index           =   2
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "Bajas"
            Index           =   4
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "-"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "libre"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "Libre"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "libre"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnLaboralHoras1 
            Caption         =   "Deshacer proceso diario/semanal"
            Index           =   10
         End
      End
      Begin VB.Menu mnLaboral1 
         Caption         =   "Nominas"
         Index           =   1
         Begin VB.Menu mnLaboralNominas 
            Caption         =   "Historico nominas"
            Index           =   0
         End
         Begin VB.Menu mnLaboralNominas 
            Caption         =   "Generacion desde horas"
            Index           =   1
         End
         Begin VB.Menu mnLaboralNominas 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnLaboralNominas 
            Caption         =   "Exportación "
            Index           =   3
         End
      End
      Begin VB.Menu mnLaboral1 
         Caption         =   "Anticipos"
         Index           =   2
         Begin VB.Menu mnLaboralAnticipos 
            Caption         =   "Mantenimiento anticipos"
            Index           =   0
         End
         Begin VB.Menu mnLaboralAnticipos 
            Caption         =   "Generacion desde horas"
            Index           =   1
         End
         Begin VB.Menu mnLaboralAnticipos 
            Caption         =   "Generar pagos banco"
            Index           =   2
         End
         Begin VB.Menu mnLaboralAnticipos 
            Caption         =   "3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnLaboralAnticipos 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnLaboralAnticipos 
            Caption         =   "Mantenimientos BIC/SWIFT"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnProduccion 
      Caption         =   "&Producción"
      Begin VB.Menu mnProduccion1 
         Caption         =   "Datos reloj"
         Index           =   0
      End
      Begin VB.Menu mnProduccion1 
         Caption         =   "Tarea actual"
         Index           =   1
      End
      Begin VB.Menu mnProduccion1 
         Caption         =   "Ver ticajes/tareas"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnProduccion1 
         Caption         =   "Insertar ticajes manual"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnProduccion1 
         Caption         =   "Traer datos maquina"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnProduccion1 
         Caption         =   "libre"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnProduccion1 
         Caption         =   "-"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnProduccion1 
         Caption         =   "Eliminar datos para recalcular"
         Index           =   7
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnGeneraInformes 
      Caption         =   "&Informes"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnInformesDatosBasicos 
         Caption         =   "Informes datos basicos"
         Visible         =   0   'False
         Begin VB.Menu mnListTrabajadores 
            Caption         =   "Trabajadores"
         End
         Begin VB.Menu mnListadoHorarios 
            Caption         =   "Horarios"
         End
         Begin VB.Menu mnListadoSecciones 
            Caption         =   "Secciones"
         End
      End
      Begin VB.Menu mnbarra103 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnListadoPendienteProcesar 
         Caption         =   "Pendiente procesar"
      End
      Begin VB.Menu mnbarra17 
         Caption         =   "-"
      End
      Begin VB.Menu mnPresencia 
         Caption         =   "&Marcajes real"
      End
      Begin VB.Menu mnCombinado 
         Caption         =   "&Presencia"
      End
      Begin VB.Menu mnListHorTrab 
         Caption         =   "Listado horas trabajadas"
         Begin VB.Menu mnResumenMensual 
            Caption         =   "Resumen mensual"
            Visible         =   0   'False
         End
         Begin VB.Menu mnImportes 
            Caption         =   "Horas con importes"
         End
         Begin VB.Menu mnListadoHorasJornadas 
            Caption         =   "Horas Jornadas"
            Visible         =   0   'False
         End
         Begin VB.Menu mnHorasxReloj 
            Caption         =   "Horas x Reloj"
         End
      End
      Begin VB.Menu mnInformaesCominiados 
         Caption         =   "Combinados Nom."
         Visible         =   0   'False
         Begin VB.Menu mnResumenHorasNomin 
            Caption         =   "Horas totales"
         End
         Begin VB.Menu mnResumenCuartilla 
            Caption         =   "Resumen cuartilla"
         End
         Begin VB.Menu mnBarra20 
            Caption         =   "-"
         End
         Begin VB.Menu mnNominasBolsa 
            Caption         =   "Nominas/Bolsa"
         End
      End
      Begin VB.Menu mnDiasTrabajados 
         Caption         =   "Dias trabajados"
      End
      Begin VB.Menu mnIncidPpal 
         Caption         =   "&Incidencias"
         Begin VB.Menu mnIncResumen 
            Caption         =   "&Incidencias RESUMEN"
         End
         Begin VB.Menu mnGeneradas 
            Caption         =   "Incidencias &Generadas"
         End
      End
      Begin VB.Menu mnBarraProd1 
         Caption         =   "-"
      End
      Begin VB.Menu mnInformesproduccion 
         Caption         =   "Produccion"
      End
      Begin VB.Menu mnbarra4_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnGenerarCodigoBarras 
         Caption         =   "Generar codigo de barras"
         Begin VB.Menu mnEANTrabajadores 
            Caption         =   "Trabajadores"
         End
         Begin VB.Menu mnEANTareas 
            Caption         =   "Tareas"
         End
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnRobotics 
         Caption         =   "Operaciones &Robotics"
      End
      Begin VB.Menu mnOperacionesTCP3 
         Caption         =   "Operaciones &TCP-3"
      End
      Begin VB.Menu mnBarra19 
         Caption         =   "-"
      End
      Begin VB.Menu mnCopiaSeg 
         Caption         =   "Copia seguridad local"
      End
      Begin VB.Menu mnUsuariosActivos 
         Caption         =   "Usuarios activos"
      End
   End
   Begin VB.Menu mnAcerca 
      Caption         =   "Acerca de ..."
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnAcercaDef 
         Caption         =   "Control de Presencia y Gestión Laboral"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Private FechaRevision As Date
Private PrimeraVez As Boolean

'---------------------------------------------------------------------------------------------
Private LlevaRelojesAuxiliares As Boolean   'Si existe la tabla es que lleva relojes auxiliares


Private Sub MDIForm_Activate()
   
    
    If PrimeraVez Then
        PrimeraVez = False
        If Not vEmpresa Is Nothing Then
            Me.Tag = "Aripres   ver. " & App.Major & "." & App.Minor & "." & Format(App.Revision, "00") & "   -   " & vEmpresa.NomEmpresa & "   -   " & vUsu.Nombre
            Screen.MousePointer = vbHourglass
            Caption = ".................  Leyendo datos BD ............................"
            DoEvents
            'Comprobamos si la fecha de hoy menos la del ultimo horario asignado
            '   a los trabajadores es menor que 15 dias y muestro mensaje
            
            If vEmpresa.TodosLosDias Then
                CadenaDesdeOtroForm = DiferenciaDias
                If CadenaDesdeOtroForm <> "" Then
                    CadenaDesdeOtroForm = vbCrLf & vbCrLf & vbCrLf & "Los trabajadores no tienen horario asignado "
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "de aqui a " & NumRegElim & " dia"
                    If NumRegElim <> 1 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "s. " & vbCrLf & vbCrLf
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Debería asignarselo."
                    MsgBox CadenaDesdeOtroForm, vbExclamation
                    CadenaDesdeOtroForm = ""
                End If
                NumRegElim = 0
            End If
            
            If vEmpresa.Reloj = vbTCP3 Then
    '            If mConfig.ComprobarHoraReloj Then
    '                frmTCP3.Comprobar = True
    '                frmTCP3.Show vbModal
    '            End If
            End If
            
            
            If vEmpresa.SolicitudVacaciones Then
                CadenaDesdeOtroForm = " fecha >= " & DBSet(vEmpresa.FechaInicio, "F") & " AND situacion"
                CadenaDesdeOtroForm = DevuelveDesdeBD("count(*)", "trabajadoresvacaciones", CadenaDesdeOtroForm, "0")
                If Val(CadenaDesdeOtroForm) Then
                    MsgBox "Tiene dias de vacaciones pendientes de aprobar", vbInformation
                End If
            End If
            CadenaDesdeOtroForm = ""
            Caption = Me.Tag
            Me.Tag = ""
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub MDIForm_Load()
Dim B As Boolean

On Error Resume Next
     PrimeraVez = True
   
    Me.Left = 9
    Me.Top = 0
    Me.Width = 12000
    Me.Height = 9000
    
    If Toolbar1.ImageList Is Nothing Then Set Toolbar1.ImageList = ImageList22
    
    If vEmpresa Is Nothing Then
        Toolbar1.Visible = False
        mnGeneraInformes.Visible = False
        mnLaboral.Visible = False
        mnProduccion.Visible = False
        mnOperaciones.Visible = False
        
        'Tampoco podrá..
        Me.mnTrabajadores.Enabled = False
        Me.mnCalendarios.Enabled = False
        Me.mnCategorias.Enabled = False
        Exit Sub
    End If
    'Ponemos los dibujitos
    Toolbar1.Buttons(1).Image = 1   'Trabajadores
    Toolbar1.Buttons(2).Image = 3  'Horario
    Toolbar1.Buttons(3).Image = 4   'Horario
    Toolbar1.Buttons(5).Image = 2  'Revisar
  '  Toolbar1.Buttons(6).Image = 2  'revisar
    Toolbar1.Buttons(8).Image = 12  'TCP3
    Toolbar1.Buttons(9).Image = 5  'procesar
    Toolbar1.Buttons(10).Image = 6
           
             
    Toolbar1.Buttons(12).Image = 8  'Presecia
    Toolbar1.Buttons(13).Image = 7  'Resumen
    Toolbar1.Buttons(14).Image = 7 'Dias trabajados
    Toolbar1.Buttons(15).Image = 11 'Maquina
    
    Toolbar1.Buttons(17).Image = 9  'SAlir
    
    '
    
    'B = (vUsu.Nivel < 2) And vEmpresa.reloj = vbTCP3
    B = vEmpresa.Reloj = vbTCP3
    mnOperacionesTCP3.Enabled = B
    Toolbar1.Buttons(8).Visible = B
    
    Me.Toolbar1.Buttons(15).Visible = (vEmpresa.Reloj = vbKimaldi)
    
    
   ' Toolbar1.Buttons(9).Visible = mConfig.Ariadna
  '  mnTraspasar.Enabled = mConfig.Ariadna
    
    'Si es reloj KIMALDI
  '  B = (vUsu.Nivel < 2) And mConfig.Kimaldi
'    Me.mnProducción.Visible = mConfig.Kimaldi
'    Me.Toolbar1.Buttons(15).Visible = B
'    Me.mnBarraProd1.Visible = mConfig.Kimaldi
'    Me.mnInformesproduccion.Visible = mConfig.Kimaldi
'    mnBarraProd1.Visible = mConfig.Kimaldi
'    mnbarra4_3.Visible = mConfig.Kimaldi
'    Me.mnGenerar.Visible = mConfig.Kimaldi
'    Me.mnEliminarDatosKimaldi.Enabled = B
    
     
    'Si no lleva laboral
    mnLaboral.Visible = vEmpresa.laboral
    mnProduccion.Visible = vEmpresa.produccion
    mnbarra4_3.Visible = vEmpresa.produccion
    mnBarraProd1.Visible = vEmpresa.produccion
    mnGenerarCodigoBarras.Visible = vEmpresa.produccion
    Me.mnInformesproduccion.Visible = vEmpresa.produccion
    mnInformaesCominiados.Visible = vEmpresa.laboral
    mnTareas.Visible = vEmpresa.produccion
    
    
    'Cambiado 20 Octubre 2004
    'No dejamos visible importar ficherito
    '-----------------------------------------
    ' mnbarra2_1.Visible = Not mConfig.Kimaldi
    If vEmpresa.Reloj = vbKimaldi Then
        mnImportar.Caption = "Generar entradas presencia"
    Else
        mnImportar.Caption = "Importar &fichero de datos"
    End If
        
    mnLecturaReloj(0).Visible = vEmpresa.Reloj = vbTCP3
    mnLecturaReloj(1).Visible = vEmpresa.Reloj = vbKimaldi
    mnLecturaReloj(2).Visible = vEmpresa.Reloj = vbFingKey
    mnLecturaReloj(3).Visible = vEmpresa.Reloj = vbZKTeco
    mnLecturaReloj(4).Visible = vEmpresa.Reloj = vbBiostarXpass
   
    mnLecturaReloj2(0).Visible = False
    mnLecturaReloj2(1).Visible = False
    If vEmpresa.Reloj2 > 0 Then
        mnLecturaReloj2(1).Visible = True
        If vEmpresa.Reloj2 = 1 Then mnLecturaReloj2(1).Caption = "Leer reloj Biostar"
    
    End If
   
   
   
    Me.StatusBar1.Panels(2).Text = "Empresa: " & vEmpresa.NomEmpresa & "            Usuario:    " & vUsu.Nombre
    

    
    
    'Para los l esten visibles aplicamos el nivel usuario
    'Nivel 0 y 1. ADministrador.
    B = vUsu.Nivel < 2
    Me.mnMantenUsuarios.Enabled = B
    
    
    mnRobotics.Enabled = (vEmpresa.Reloj = vbRobotics) And B
    
    
    'Importar fichero de datos.
    'True para administradores y...
    '   - Reloj:   tcp3,alzira, robotics
    '       NO:     CATADAU
    mnImportar.Enabled = B And (vEmpresa.Reloj <> vbKimaldi)
        
    
    
    mnInformaesCominiados.Visible = False 'NO estan desarrollados
    
    'mnListHorTrab.Visible = vEmpresa.QueEmpresa <> 5
    mnImportes.Visible = vEmpresa.QueEmpresa <> 5
    Me.mnHorasxReloj.Visible = vEmpresa.Reloj2 > 0

    'Relojes auxiliares
    mnbarra2_8.Visible = False
    mnRelojesAuxiliares.Visible = False
    
    If vEmpresa.QueEmpresa = 2 Then
        TieneRelojesAuxiliares
    End If
    
    mnProcesoPrevioALZ.Visible = vEmpresa.HorarioNocturno2
    
    
    'SI lleva dos relojes
    mnHorasxReloj.Visible = vEmpresa.Reloj2 > 0
    
    'Aqui asociamos los botones de la tool con el menu
    Toolbar1.Buttons(13).Visible = mnCombinado.Visible
    Toolbar1.Buttons(13).Visible = mnPresencia.Visible

   

    'Deshacer semana
    mnLaboralHoras1(9).Visible = vEmpresa.QueEmpresa = 2 Or vEmpresa.QueEmpresa = 4
    mnLaboralHoras1(10).Visible = vEmpresa.QueEmpresa = 2 Or vEmpresa.QueEmpresa = 4


   'SolicitudVacaciones
   mnVacaciones.Visible = vEmpresa.SolicitudVacaciones
    
    
    If vEmpresa.Reloj = vbKimaldi Then
        'Alzira  entra aqui
    
        ' La bD esta en el ODBC driver de MDB y se llama accGestorHuella
        AbrirBaseDatos
    End If



End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    conn.Close
    Set conn = Nothing
    CerrarConexionesKreta
    End
End Sub

Private Sub TieneRelojesAuxiliares()
    On Error Resume Next
    conn.Execute "Select * from entradafichajAuxliares where Secuencia = -1"
    'Si no da error es que la tabla existe3
    If Err.Number = 0 Then
        mnbarra2_8.Visible = True
        mnRelojesAuxiliares.Visible = True
    Else
        Err.Clear
    End If
    
End Sub






Private Sub mnAcercaDef_Click()
    frmAbout.Show vbModal
End Sub






Private Sub mnBancos_Click()
    frmBancos.Show vbModal
End Sub

Private Sub mnCalendarios_Click()
    frmCalendario.Show vbModal
End Sub

Private Sub mnCategorias_Click()
    frmCategoria.Show vbModal
End Sub





Private Sub mnCombinado_Click()
    AbrirListado 13
End Sub

Private Sub mnCopiaSeg_Click()
    frmBackUP.Show vbModal
End Sub


Private Sub mnDiasTrabajados_Click()
    AbrirListado 14
End Sub



Private Sub mnEmpresas_Click()
    frmEmpresa.Show vbModal
End Sub








Private Sub mnGeneradas_Click()
    AbrirListado 15
End Sub

Private Sub mnHistoricoMarcajes_Click()
    frmRevisionHCO.Show vbModal
    
End Sub

Private Sub mnHorarios_Click()
    frmHorario.Show vbModal
End Sub


Private Sub mnHorasxReloj_Click()
    frmListado.Opcion = 22
    frmListado.Show vbModal
End Sub

Private Sub mnImportar_Click()
    frmTraspaso.Opcion = 0
    frmTraspaso.Show vbModal
    If vEmpresa.QueEmpresa = 5 Then frmRelojBiostar2.Show vbModal
End Sub


Private Sub mnImportes_Click()
    frmListado.Opcion = 16
    frmListado.Show vbModal
End Sub

Private Sub mnIncidencias_Click()
    frmIncidencias.Show vbModal
End Sub

Private Sub mnIncResumen_Click()
    AbrirListado 11
End Sub

Private Sub mnLaboralAnticipos_Click(Index As Integer)
    Select Case Index
    Case 0
        frmListadoAnticipos.Show vbModal
    Case 1
         frmGeneraAnti.Show vbModal
    Case 2
        frmPagosBanco2.Opcion = 0
        frmPagosBanco2.Show vbModal
    
    Case 5
        frmbic.Show vbModal
    End Select
End Sub

Private Sub mnLaboralHoras1_Click(Index As Integer)
    Select Case Index
    Case 0
         frmHorasProcesadas2.Show vbModal
    
    Case 1
        'GEneracion de HORAS
        Do
            CadenaDesdeOtroForm = ""
            FrmVarios.Opcion = 4
            FrmVarios.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                frmCalcularHorasSemana.TodasSecciones = RecuperaValor(CadenaDesdeOtroForm, 3) = 1
                frmCalcularHorasSemana.Show vbModal
            End If
        Loop Until CadenaDesdeOtroForm = ""
        
    Case 2
        CadenaDesdeOtroForm = ""
        FrmVarios.Opcion = 5
        FrmVarios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            frmDatosMesAlz.Mes = CadenaDesdeOtroForm
            frmDatosMesAlz.Show vbModal
        End If
    
    
    Case 4
        frmBajas.Show vbModal
    
    Case 6
       
    Case 10
        'deshaceProcesoJorandasSemanas
        frmDeshacerProcesoSemanal.Show vbModal
    End Select
End Sub

Private Sub mnLaboralNominas_Click(Index As Integer)
Dim B As Boolean

    Select Case Index
    Case 0
        frmNominas.Show vbModal
    Case 1
        B = False
        If vEmpresa.CompensaHorasNominaMES Then
            B = True
        Else
            If vEmpresa.QueEmpresa = 4 Then B = True
        End If
        'If vEmpresa.CompensaHorasNominaMES Then
        If B Then
            frmCalculoHorasMesConEstrc.Show vbModal
        Else
            frmCalculoHorasMes.Show vbModal
        End If
    Case 3
        'exportar
        frmListado.Opcion = 20
        frmListado.Show vbModal
    End Select
End Sub

Private Sub mnLecturaReloj_Click(Index As Integer)
    Select Case Index
    Case 0
        'tcp3
        frmTCP3.Show vbModal
    Case 1
        'Kreta
        frmKreta3.Show vbModal
    Case 2
        frmNitGen.Show vbModal
        
    Case 3
        frmRelojZKTeco.Show vbModal
        
    Case 4
        frmRelojXpass.Show vbModal
    End Select
End Sub


'Segundos relojes
Private Sub mnLecturaReloj2_Click(Index As Integer)
    If Index = 1 Then
        frmRelojBiostar2.Show vbModal
    End If
End Sub

Private Sub mnListadoHorarios_Click()
    ImprimeBasicos 30, True
End Sub

Private Sub ImprimeBasicos(NumInforme As Integer, TienSubinformes As Boolean)
    With frmImprimir
        .NumeroParametros = 1
        .FormulaSeleccion = ""
        .OtrosParametros = "pEmpresa= '" & vEmpresa.NomEmpresa & "'|"
        .Opcion = NumInforme
        .Show vbModal
    End With
End Sub

Private Sub mnListadoPendienteProcesar_Click()
    CadenaDesdeOtroForm = ""
    AbrirListado 12
End Sub

Private Sub mnListadoSecciones_Click()
    ImprimeBasicos 31, False
End Sub

Private Sub mnListTrabajadores_Click()
    AbrirListado 8
End Sub

Private Sub AbrirListado(vOpcion As Integer)

    frmListado.Opcion = vOpcion
    frmListado.Show vbModal
End Sub


Private Sub mnMantenUsuarios_Click()

    'Si tiene valor significa que la BD NO, repito NO,
    'esta en el mismo server que aripres
    
    If vEmpresa.Server <> "" Then
        conn.Close
        'Abrir otra conexion
        If Not AbrirConnParaUsuarios() Then
            AbrirConexion
            frmEmpresa.Show vbModal
            End
        End If
    End If
    
    
    frmMantenusu.Show vbModal
    
    
    If vEmpresa.Server <> "" Then
        conn.Close
        'Abrir otra conexion
        AbrirConexion
    End If
    
End Sub

Private Sub mnOperacionesTCP3_Click()
    frmTCP3.Show vbModal
End Sub




Private Sub mnPresencia_Click()
            frmListado.Opcion = 2
            frmListado.Show vbModal
End Sub

Private Sub mnProcesar_Click()
    HacerToolBar 10
End Sub

Private Sub mnProcesoPrevioALZ_Click()
Dim Aux As String

    If Not vEmpresa.AcabaJornadaDiaSiguiente Then Exit Sub

    'Hacer PROCESO horas acabalgadas
    Screen.MousePointer = vbHourglass
    Aux = Me.StatusBar1.Panels(2).Text
    Me.StatusBar1.Panels(2).Text = "Leyendo datos"
    Me.StatusBar1.Refresh
     
    CadenaDesdeOtroForm = DevuelveDesdeBD("min(fecha)", "entradafichajes", "1", "1")
    If CadenaDesdeOtroForm <> "" Then
        frmPrevioProcesar.Modificar = False
        frmPrevioProcesar.Fecha = CDate(CadenaDesdeOtroForm)
        frmPrevioProcesar.Show vbModal
    End If
    
    
    Me.StatusBar1.Panels(2).Text = Aux
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnProduccion1_Click(Index As Integer)
    Select Case Index
    Case 0
        If vEmpresa.Reloj = vbKimaldi Then
            frmDatosKimaldi.Show vbModal
        End If
    Case 1
        frmTareaActual.QueFecha = Now
        frmTareaActual.Opcion = 0
        frmTareaActual.Show vbModal
    End Select
End Sub

Private Sub mnRelojesAuxiliares_Click()
    frmTareaActuaRelojAuxiliar.Show vbModal
End Sub

'Private Sub mnResumen_Click()
'     frmListado.Opcion = 1
'        frmListado.Show vbModal
'End Sub

Private Sub mnResumenMarcajes_Click()
    frmMarcajesPantalla.QuieroVerDatos = ""
    frmMarcajesPantalla.Show vbModal
End Sub



Private Sub mnRobotics_Click()
    Screen.MousePointer = vbHourglass
    LanzaRobotics
    CadenaDesdeOtroForm = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnSalir_Click()
    HacerToolBar 17
End Sub

Private Sub mnsecciones_Click()
    frmSeccion.Show vbModal
End Sub

'
'
'
'Private Sub mnOperacionesTCP3_Click()
'    'Utilizaremos esta variable global para saber si hay que importar
'    'un nuevo ficehero de datos
'    MostrarErrores = False
'    frmTCP3.Comprobar = False
'    frmTCP3.Show vbModal
'    If MostrarErrores Then
'        'Hay que importar
'        Screen.MousePointer = vbHourglass
'        frmTraspaso.Opcion = 1  'PARA SABER QUE VENIMOS DESDE TCP3
'        frmTraspaso.Show vbModal
'        Screen.MousePointer = vbDefault
'    End If
'End Sub










Private Sub mnSelecImpresora_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    cd1.DialogTitle = "SELECCIONA LA IMPRESORA"
    cd1.ShowPrinter
    Screen.MousePointer = vbDefault
End Sub




Private Sub mnTareas_Click()
    frmTarea.Show vbModal
End Sub

Private Sub mnTerminales_Click()
    frmTerminales.Show vbModal
End Sub

Private Sub mnTicajeActual_Click()
    frmTareaActual.QueFecha = Now
    frmTareaActual.Opcion = 1
    frmTareaActual.Show vbModal
End Sub



Private Sub mnTrabajadores_Click()
    frmTrabajadores.Show vbModal
End Sub






Private Sub mnUsuariosActivos_Click()
    NoHaceNada
End Sub

Private Sub mnVacaciones_Click()
    frmColVacaciones.Show vbModal
End Sub

Private Sub mnVerMarcajes_Click()
    CadenaDesdeOtroForm = ""
    frmRevision.MostrarUnosDatos = 0
    frmRevision.Show vbModal
End Sub

Private Sub mnZonas_Click()
    frmZonas.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 
    HacerToolBar Button.Index
End Sub

Private Sub HacerToolBar(Index As Integer)

    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 1
    
            'Trabajadores
             frmTrabajadores.Show vbModal
    Case 2
            'Calendario
            frmCalendario.Show vbModal
            
            
    Case 3
            mnHorarios_Click
            
    Case 5
            
            CadenaDesdeOtroForm = ""
            frmRevision.MostrarUnosDatos = 0
            frmRevision.Show vbModal
    Case 8
            'TCP3
            mnOperacionesTCP3_Click
    Case 9
          
        'marcaje actual
            mnTicajeActual_Click
            
    Case 10
            
            'procesar marcaje
            frmProcesarEntradasMarcajes.Show vbModal
            
            
    Case 12
        'listado de maracjes. PRESENCIA
        mnPresencia_Click

                
            
                
    Case 13
       
        
        mnCombinado_Click
       
       
       
       
       
    Case 15
        mnLecturaReloj_Click 1
    Case 17
        Unload Me
    End Select
    Screen.MousePointer = vbDefault
End Sub








Private Function DiferenciaDias() As String
    On Error GoTo EDiferenciaDias
    
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 1000
    miRsAux.Open "Select max(fecha) from calendariot", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            NumRegElim = Abs(DateDiff("d", Now, miRsAux.Fields(0)))
            If NumRegElim < 15 Then DiferenciaDias = NumRegElim
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If NumRegElim = 1000 Then DiferenciaDias = ""
    
    
    Exit Function
EDiferenciaDias:
    MuestraError Err.Number, Err.Description, "Diferencia Dias inicio"
    DiferenciaDias = ""
End Function


Private Sub NoHaceNada()
     MsgBox "Opción no disponible temporalmente" & vbCrLf & vbCrLf, vbExclamation
End Sub


'
'Private Sub PoneMenusDelEditor()
'Dim T As Control
'Dim SQL As String
'Dim C As String
'
'    On Error GoTo ELeerEditorMenus
'
'    SQL = "Select * from usu arios.appmenususuario where aplicacion='conta' and codusu = " & Val(Right(CStr(vUsu.Codigo), 3))
'    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    SQL = ""
'
'    While Not miRsAux.EOF
'        If Not IsNull(miRsAux.Fields(3)) Then
'            SQL = SQL & miRsAux.Fields(3) & "·"
'        End If
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
'
'    If SQL <> "" Then
'        SQL = "·" & SQL
'        For Each T In Me.Controls
'            If TypeOf T Is menu Then
'               ' C = DevuelveCadenaMenu(T)
'                C = "·" & C & "·"
'                If InStr(1, SQL, C) > 0 Then T.Visible = False
'
'            End If
'        Next
'    End If
'ELeerEditorMenus:
'    Set miRsAux = Nothing
'    If Err.Number <> 0 Then Err.Clear
'End Sub



Private Sub LanzaRobotics()
    On Error GoTo ELanzaRobotics
    CadenaDesdeOtroForm = DevuelveDesdeBD("configreloj", "empresas", "idempresa", 1, "N")
    If CadenaDesdeOtroForm = "" Then Exit Sub
    
    If Dir(CadenaDesdeOtroForm, vbArchive) = "" Then
        MsgBox "No existe " & CadenaDesdeOtroForm, vbExclamation
        Exit Sub
    End If
    
    
    Shell CadenaDesdeOtroForm, vbNormalFocus
    
    
    Exit Sub
ELanzaRobotics:
    MuestraError Err.Number, Err.Description
End Sub










'***************************************
'***************************************************************************************************************
'***************************************************************************************************************
'***************************************************************************************************************
'
'
'   Proceso nocturno. Horas acabalagadas
'
'
'***************************************************************************************************************
'***************************************************************************************************************
'***************************************************************************************************************
'  Ira dia a dia, desde el ultimo dia procesado hasta el siguiente que pueda procesar
'  Vera si en el dia N hay algun marcaje superior a las HoraParametros
'  Si. Ese marcae puede ser el incial de los que ficharan al dia seguiente
'   o puede ser el ultimo del dia actual
'    ¿Como lo sabremos?  Pq no habra ningun fichaje entre [ HoraParametros -6   y  horaparametros]
'    es decir, por ejemplo, Dia 1. si Horaparametros son las 22:30, si no tiene ningun marcaje entre
'    las 16:30 y las 22:30 significara que ESE es el primero de los que vendran al dia siguiente
'    Es decir, cuando procesemos el dia 2, los marcajes anteriores a HoraParametros -6  seran del mismo dia que el 1
'    Si diaPrimeraFichada= dia trabajado
'       -updaearemos los del dia siguiente a hh:mm + 24:00 y dia=dia -1
'    Si no
'       -updaearemos el primero a hh:mm  - 24:00    y dia=dia +1
'
Private Sub ProcesoHorasAcabalgadas()
Dim primeraFechaProcesar As Date
Dim Cad As String
Dim RegistrosTratar As Collection
Dim FE As Date
Dim Hora As Date
Dim H1 As Date
Dim I As Long

Dim QueDia As Integer
Dim DiaTraba As Collection

Dim UltimaDiaHoraTraidoMaquina2 As String
Dim UltimoDiaProcesado As Date
Dim FechaParaActualizarEnParametros As Date
Dim RS As ADODB.Recordset

Dim DiasATratar As Collection
Dim J As Integer

    On Error GoTo eProcesoHorasAcabalgadas
    
    Me.StatusBar1.Panels(2).Text = "Horas acabalgadas"
    Me.StatusBar1.Refresh
    
    
    
    'Ver si hay que entrar en el proceso
    
    
    
    
    Set RS = New ADODB.Recordset
    
    Cad = DevuelveDesdeBD("AcabalUltimoDiaProcesado", "empresas", "1", "1")
    If Cad = "" Then
        'No esta grababado todavia. voy a ver la primera fecha de entradafichajes
        'select min(fecha) from entradafichajes
        Cad = DevuelveDesdeBD("min(fecha)", "entradafichajes", "1", "1")
        If Cad = "" Then Cad = "02/01/1900"
        Cad = DateAdd("d", -1, CDate(Cad))  'Para que el primer dia trabajado se el primero de entradafichakes
    End If
    primeraFechaProcesar = CDate(Cad)
    UltimoDiaProcesado = primeraFechaProcesar
    
    'Voy a ver ultimo dia -hora que hemos traido desde la maquina
    Cad = "Select fecha , concat(horareal,'') h1 from entradafichajes ORDER BY 1 desc,2 desc"
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        Err.Raise 513, , "Entrada fichejes vacia"
    Else
        Cad = "23:59:59"
        If HoraFueraInterval(RS!H1) = 0 Then Cad = Format(RS!H1, "hh:nn:ss")
        Cad = Format(RS!Fecha, "dd/mm/yyyy") & " " & Cad
    End If
    RS.Close
    UltimaDiaHoraTraidoMaquina2 = Cad
    
    
    
    'Vamos a ver los dias a tratar
    Set DiasATratar = New Collection
    
    I = Round(vEmpresa.MaxRetraso * 60, 0)
    Cad = DateAdd("n", -I, vEmpresa.AcabalgadoHora)   'para ver si puedo procesar el dia
    If CDate(Format(UltimaDiaHoraTraidoMaquina2, "hh:nn:ss")) >= CDate(Cad) Then
        Cad = ""
    Else
        Cad = " AND fecha <" & DBSet(UltimaDiaHoraTraidoMaquina2, "F")
    End If
    
    
    
    Cad = "Select distinct fecha from entradafichajes where fecha> " & DBSet(primeraFechaProcesar, "F") & Cad
    Cad = Cad & "  order by 1"
    
 
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        DiasATratar.Add CStr(Format(RS.Fields(0), "dd/mm/yyyy"))
        RS.MoveNext
    Wend
    RS.Close
    
    
    If DiasATratar.Count = 0 Then GoTo eProcesoHorasAcabalgadas
    
    
    'Vamos a estudiar cada dia
    For QueDia = 1 To DiasATratar.Count
            'Voy a ver que dias tienen fichajes superiror a las 22:30 (parametros)
            ' Y luego estudiare esos dias
            I = Round(vEmpresa.MaxRetraso * 60, 0)
            Cad = DateAdd("n", -I, vEmpresa.AcabalgadoHora)   'para ver si puedo procesar el dia
            Hora = Cad
            Cad = "fecha = " & DBSet(DiasATratar.Item(QueDia), "F") & " AND hora > " & DBSet(Hora, "H") & " and hora <= '23:59:59'"
            'Select  from entradafichajes where fecha> '2001-01-10' AND hora > '22:00:00' and hora <= '23:59:59' ORDER BY fecha,idtrabajador
            Cad = "Select distinct idtrabajador from entradafichajes where " & Cad & " ORDER BY idtrabajador"
            RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Set RegistrosTratar = New Collection
            I = -1
            While Not RS.EOF
                Cad = RS!idTrabajador
                RegistrosTratar.Add Cad
                RS.MoveNext
            Wend
            RS.Close
            
            'Para que haga los nothing
            If RegistrosTratar.Count > 0 Then
                    Me.StatusBar1.Panels(2).Text = "Dia " & DiasATratar.Item(QueDia)
                    Me.StatusBar1.Refresh
    

                    'Procesar el dia
                    ' Para ello vere fichajes mas alla de las 22:30 del dia D
                    ' vere si el trabajador para ese dia NO tiene ninguna entrada entre las 22:30 y las 14:30, es decir,
                    ' si realmente la 22:30 es la entrada del dia
                    '
                    ' SACARE UN FRM previo, donde podra decir si el trabajador ese dia es acabalglado  o no
                    
                    '
                    conn.Execute "Delete from tmpnotrabajo"
                    espera 0.5
                    Cad = ""
                    For I = 1 To RegistrosTratar.Count
                        Cad = Cad & ", (" & RegistrosTratar(I) & ")"
                    Next
                    Cad = Mid(Cad, 2)
                    Cad = "INSERT INTO tmpnotrabajo(idTra) VALUES " & Cad
                    conn.Execute Cad
                    espera 0.5
                    
                    CadenaDesdeOtroForm = ""
                    frmAcabalgados.Fecha = CDate(DiasATratar.Item(QueDia))
                    frmAcabalgados.Show vbModal
                    
                    'Ha pulsado aceptar
                    If CadenaDesdeOtroForm = "" Then
                        'Cancelado proceso
                         GoTo eProcesoHorasAcabalgadas
                    Else
                        Me.StatusBar1.Panels(2).Text = "Leyendo registros"
                        Me.StatusBar1.Refresh

                        DoEvents
                        Screen.MousePointer = vbHourglass
                    
                        Cad = "Select * from tmpnotrabajo ORDER by idtra"
                        Set DiaTraba = New Collection
                        
                        RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        While Not RS.EOF
                            DiaTraba.Add CStr(RS!idTRa)
                            RS.MoveNext
                        Wend
                        RS.Close
                        
                        
                        'Para cada trabajador
                        For J = 1 To DiaTraba.Count
                        
                         
                            'Debug.Print DiaTraba(J)
                        
                            Me.StatusBar1.Panels(2).Text = "Trabajador: " & DiaTraba(J)
                            Me.StatusBar1.Refresh
                            
                            'Abrimos un frm para que nos diga si son acabalgados o normales y se l
                            
                            
                        
                            
                            FE = CDate(DiasATratar.Item(QueDia))
                            
                            If vEmpresa.AcabalgadoDiaInicio Then
                                'La primera fichada marca el dia de inicio
                                '       -updaearemos los del dia siguiente a hh:mm + 24:00 y dia=dia -1
                                Cad = "fecha = " & DBSet(FE, "F")
                                Cad = Cad & ",hora = ADDTIME(hora , '24:00:00' ) "
                                Cad = Cad & ",horareal = ADDTIME(horareal , '24:00:00' ) "
                                Cad = Cad & " WHERE fecha = " & DBSet(DateAdd("d", 1, FE), "F")
                                Cad = Cad & " AND hora < " & DBSet(Hora, "H")
                            Else
                                'La fichada es la primera del dia siguiente
                                '       -updaearemos el primero a hh:mm  - 24:00    y dia=dia +1
                                Cad = "fecha = " & DBSet(DateAdd("d", 1, FE), "F")
                                Cad = Cad & ",hora = ADDTIME(hora , '-24:00:00' ) "
                                Cad = Cad & ",horareal = ADDTIME(horareal , '-24:00:00' ) "
                                Cad = Cad & " WHERE fecha = " & DBSet(FE, "F")
                                Cad = Cad & " AND hora >= " & DBSet(Hora, "H")
                                                  
                            End If
                            Cad = Cad & " AND idtrabajador = " & DiaTraba(J)
                            Cad = "UPDATE entradafichajes set " & Cad
                            conn.Execute Cad
                            espera 0.1
                                
                        Next J
                                
                        Cad = "UPDATE empresas set AcabalUltimoDiaProcesado = " & DBSet(FE, "F")
                        conn.Execute Cad
                        Me.StatusBar1.Panels(2).Text = "Actualizando ......."
                        Me.StatusBar1.Refresh
                        espera 1.5
                        
                    End If
            Else
                'Hemos pulsado tratar dia pero no hay trabajadores para mirar
                Cad = "UPDATE empresas set AcabalUltimoDiaProcesado = " & DBSet(DiasATratar.Item(QueDia), "F")
                conn.Execute Cad
            End If
            
            Set RegistrosTratar = Nothing
    Next QueDia
eProcesoHorasAcabalgadas:
    
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set RS = Nothing
    Set RegistrosTratar = Nothing
    Set DiaTraba = Nothing
    Set DiasATratar = Nothing
    
End Sub






