VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRevisionHCO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historico marcajes"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   Icon            =   "frmRevisionHCO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRevisionHCO.frx":6852
   ScaleHeight     =   7470
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTapaImg 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8640
      TabIndex        =   48
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame FrameEspecial2 
      Enabled         =   0   'False
      Height          =   1035
      Left            =   120
      TabIndex        =   44
      Top             =   2520
      Width           =   11055
      Begin VB.TextBox txtHorario 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   50
         Text            =   "Text2"
         Top             =   360
         Width           =   3795
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Nocturno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Festivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   45
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Calendario"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.TextBox TextHt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   24
      Top             =   6960
      Width           =   915
   End
   Begin VB.TextBox TextHt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox TextHt 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   2340
      TabIndex        =   21
      Top             =   6960
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1845
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   11055
      Begin VB.TextBox txtDec 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   6840
         TabIndex        =   6
         Tag             =   "Paradas|N|N|0||marcajes|HorasDto|0.00||"
         Text            =   "Text1"
         Top             =   1320
         Width           =   825
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   615
         Left            =   9000
         TabIndex        =   41
         Top             =   360
         Width           =   1935
         Begin VB.CheckBox chkCorrecto 
            Caption         =   "Correcto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   40
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   39
         Top             =   900
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   6840
         TabIndex        =   5
         Tag             =   "HI|N|N|||marcajes|HorasInci|0.00||"
         Text            =   "Text1"
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   6840
         TabIndex        =   4
         Tag             =   "HT|N|N|0||marcajes|HorasTrabajadas|0.00||"
         Text            =   "Text1"
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtDec 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtDec 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Tag             =   "Entrada|N|N|0||marcajes|entrada||S|"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "Trabajador|N|N|0||marcajes|idtrabajador|||"
         Top             =   900
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Incidencia|N|N|0||marcajes|incfinal|||"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Fecha|F|N|||marcajes|fecha|||"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   47
         Tag             =   "idhorario|N|N|0||marcajes|idhorario|||"
         Top             =   1440
         Width           =   150
      End
      Begin VB.Label Label1 
         Caption         =   "Paradas"
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   43
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Image imgZoom 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   960
         MouseIcon       =   "frmRevisionHCO.frx":6DDC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar incidencia"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   960
         MouseIcon       =   "frmRevisionHCO.frx":6F2E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   3480
         Picture         =   "frmRevisionHCO.frx":7080
         ToolTipText     =   "Buscar fecha"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Horas incidencia"
         Height          =   255
         Index           =   12
         Left            =   5280
         TabIndex        =   38
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Horas Trabajadas"
         Height          =   195
         Left            =   5400
         TabIndex        =   37
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Sexagesimal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   7800
         TabIndex        =   36
         Top             =   120
         Width           =   1110
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Decimal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6900
         TabIndex        =   35
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencia"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Entrada"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8700
      TabIndex        =   8
      Top             =   6900
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   6900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9960
      TabIndex        =   12
      Top             =   6900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   4560
      TabIndex        =   10
      Top             =   6720
      Width           =   3945
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
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
         Left            =   45
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4440
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   14
         Top             =   120
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   120
      TabIndex        =   20
      Top             =   3900
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hora"
         Object.Width           =   2082
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Incidencia Manual"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ValorEnBD"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2415
      Left            =   6180
      TabIndex        =   23
      Top             =   3900
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Inidencia automática"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Horas"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Decimal"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Incidencia"
         Object.Width           =   0
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   510
      Left            =   1080
      Top             =   5760
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   900
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
      ConnectStringType=   3
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
   Begin MSComctlLib.ListView ListView3 
      Height          =   2415
      Left            =   3900
      TabIndex        =   25
      Top             =   3900
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hora"
         Object.Width           =   2258
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   882
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "HISTORICO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   3360
      TabIndex        =   52
      Top             =   6960
      Width           =   1035
   End
   Begin VB.Image imgInciGen 
      Height          =   240
      Index           =   2
      Left            =   9360
      Picture         =   "frmRevisionHCO.frx":710B
      ToolTipText     =   "Eliminar"
      Top             =   3600
      Width           =   240
   End
   Begin VB.Image imgInciGen 
      Height          =   240
      Index           =   1
      Left            =   9000
      Picture         =   "frmRevisionHCO.frx":7695
      ToolTipText     =   "Modificar"
      Top             =   3600
      Width           =   240
   End
   Begin VB.Image imgInciGen 
      Height          =   240
      Index           =   0
      Left            =   8640
      Picture         =   "frmRevisionHCO.frx":8097
      ToolTipText     =   "Insertar"
      Top             =   3600
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Horas Trabajadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   32
      Top             =   6420
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Nº de marc."
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   31
      Top             =   6720
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "MARCAJES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   3660
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "INCIDENCIAS GENERADAS"
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
      Left            =   6180
      TabIndex        =   29
      Top             =   3660
      Width           =   2595
   End
   Begin VB.Label Label9 
      Caption         =   "Decimal"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   6720
      Width           =   675
   End
   Begin VB.Label Label10 
      Caption         =   "Sexagesimal"
      Height          =   195
      Left            =   1200
      TabIndex        =   27
      Top             =   6720
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3120
      Y1              =   6660
      Y2              =   6660
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "REAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3900
      TabIndex        =   26
      Top             =   3660
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4560
      Picture         =   "frmRevisionHCO.frx":8621
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnbarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnOrdenar 
      Caption         =   "Ordenar"
      Begin VB.Menu mnFecha 
         Caption         =   "Fecha,trabajador"
      End
      Begin VB.Menu mnTrabajador 
         Caption         =   "Trabajador,fecha"
      End
   End
   Begin VB.Menu mnOperaciones 
      Caption         =   "Operaciones"
      Begin VB.Menu mnTraspasar 
         Caption         =   "Traspasar datos a historico"
      End
   End
End
Attribute VB_Name = "frmRevisionHCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: LAURA    +-+-
' +-+- Fecha: 28/02/06 +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+

Option Explicit


Public MostrarUnosDatos As Long

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)



Private FI As Date
Private FF As Date
Private DTra As Long
Private HTra As Long
Private DInci As Integer
Private HInci As Integer
Private CorrectosIncorrectos As Byte  '0.- Ambos  1.- Correctos  2.-Incorrectos

Private WithEvents frmHoras As frmHorasMarcajes
Attribute frmHoras.VB_VarHelpID = -1
Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1

' *** per a cridar ad atres formularis ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

' *************************************




Private CadenaConsulta As String
Private CadB As String

Dim Modo As Byte
'-------------- MODOS ---------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'------------------------------------------------

Dim Ordenacion As String

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
Dim Indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos
Dim SQL As String
Dim PrimeraVez As Boolean




Private Sub PonerModo(vModo)
Dim B As Boolean
Dim NumReg As Byte

    On Error GoTo EPonerModo
    
    Modo = vModo
    If Modo = 2 Then
        lblIndicador.Caption = PonerContRegistros(Me.adodc1)
    Else
       ' cmdCorrecto.Visible = False
        PonerIndicador lblIndicador, Modo
    End If
    
    
    B = (Modo = 2)
    If MostrarUnosDatos Then B = False
    '=======================================
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Me.adodc1.Recordset.EOF Then
        If adodc1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    '---------------------------------------------
    B = B And (NumReg = 2)
    imgInciGen(0).Enabled = B
    imgInciGen(1).Enabled = B
    imgInciGen(2).Enabled = B
    FrameTapaImg.Visible = Not B
     
     
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
        

    BloquearText1 Me, Modo
    If Modo = 4 Then
        BloquearTxt Text1(3), True
        BloquearTxt Text1(1), True
    End If
    Me.Text1(0).Enabled = Modo < 2
    
    B = Modo > 2
    imgFec(3).Visible = B
    Me.imgZoom(0).Visible = B
    Me.imgZoom(1).Visible = B
   ' Me.imgZoom(2).Visible = B
    Frame3.Enabled = B
        
    
    'BloquearImgBuscar Me, Modo
    'BloquearCmb Combo1(0), (Modo <> 1 And Modo <> 3 And Modo <> 4)

'    BloquearImgBuscar Me, Modo
    ' ********************************************************
    
    
    'Si es regresar
'    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    PonerLongCampos 'Pone el Maxlength de los campos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner modo.", Err.Description
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean


    If MostrarUnosDatos > 0 Then
        'HE VENIDO DESDE OTRO FORM. SOLO quiero ver los datos
        B = False
    Else
        B = (Modo = 2) Or Modo = 0
    End If
    

    
    'Busqueda
    Toolbar1.Buttons(2).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Ordenar
    Me.mnOrdenar.Enabled = B
    'Insertar
    'Toolbar1.Buttons(6).Enabled = B 'And Not DeConsulta
    'Me.mnNuevo.Enabled = B 'And Not DeConsulta
    
    B = (Modo = 2 And adodc1.Recordset.RecordCount > 0) 'And Not DeConsulta
    If MostrarUnosDatos > 0 Then B = False
    'Modificar
    'Toolbar1.Buttons(7).Enabled = B
    'Me.mnModificar.Enabled = B
    'Eliminar
    'Toolbar1.Buttons(8).Enabled = B
    'Me.mnEliminar.Enabled = B
    
    'REVISION COMPLETA
    'mnRevisionmultiple.Enabled = B
    
    'Imprimir
    Toolbar1.Buttons(11).Enabled = True
    
    
End Sub


Private Sub BotonAnyadir()
Dim NumF As String
    
    LimpiarCampos 'Vacía los TextBox
    CadB = ""
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
     '******************** canviar taula i camp **************************
   ' If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
   '     NumF = NuevoCodigo
   ' Else
        NumF = SugerirCodigoSiguienteStr("marcajes", "entrada")
   ' End If
    
    ' ******* Canviar el nom de la taula, el nom de la clau primaria, i el
    ' nom del camp que te la clau primaria si no es Text1(0) *************
    Text1(0).Text = NumF
    FormateaCampo Text1(0)
    chkCorrecto.Value = 0
    
    
    'PosicionarCombo Me.Combo1(0), 724
    
    'PosarDescripcions
    PonerFoco Text1(3)
    ' ********************************************************************
End Sub


Private Sub BotonVerTodos()
    CadB = ""
    LimpiarCampos 'Limpia los Text1
    CadenaDesdeOtroForm = ""
    SeparaValores
    CargaDatos
    PonerCampos
    PonerModo 2
    
'
'    HTra = 100000
'    If chkVistaPrevia.Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
'        CadenaConsulta = "Select * from " & NomTabla & Ordenacion
'        PonerCadenaBusqueda
'    End If
End Sub




Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String

        'Llamamos a al form
        ' **************** arreglar-ho per a vore lo que es desije ****************
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 10, "Cód.")
        Cad = Cad & ParaGrid(Text1(1), 26, "Nombre")
        Cad = Cad & ParaGrid(Text1(2), 32, "1º Apellido")
        Cad = Cad & ParaGrid(Text1(3), 32, "2º Apellido")
        
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            'frmB.vTabla = NomTabla
            frmB.vSQL = CadB

            '###A mano
            frmB.vDevuelve = "0|1|2|3|"
            frmB.vTitulo = "Guias viaje"
            frmB.vSelElem = 0

            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            
            PonerFoco Text1(1)
            
        End If
        ' *************************************************************************
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    SQL = "Select * from marcajes"
    If Me.mnTrabajador.Checked Then
        SQL = SQL & " ORDER BY idtrabajador,fecha"
    Else
        SQL = SQL & " ORDER BY fecha,idtrabajador"
    End If
        
    CargaDatos
    
    If adodc1.Recordset.RecordCount <= 0 Then
    
            MsgBox "No hay ningún registro en la tabla ", vbInformation
'            Screen.MousePointer = vbDefault
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub



Private Sub BotonBuscar()


    CadenaDesdeOtroForm = "hco"
    frmListado.Opcion = 0
    frmListado.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        CadB = ""
        SeparaValores
        CargaDatos
        CadenaDesdeOtroForm = ""
        
        If Me.adodc1.Recordset.EOF Then
            LimpiarCampos
            PonerModo 0
        Else
            PonerCampos
            PonerModo 2
        End If
        Screen.MousePointer = vbDefault
    End If
   
End Sub


Private Sub BotonModificar()
    
    PonerModo 4
   
    'Como es modificar
    ' *** primer control que no siga clau primaria ***
    PonerFoco Text1(2)
    ' ************************************************
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    On Error GoTo EEliminar
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Adodc2.Recordset.EOF Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el marcaje?"
    SQL = SQL & vbCrLf & "Código: " & Adodc2.Recordset!Entrada & "     -    " & Format(Adodc2.Recordset!Fecha, "dd/mm/yyyy")
    SQL = SQL & vbCrLf & "Nombre: " & Adodc2.Recordset!idTrabajador & " - " & Me.Adodc2.Recordset!nomtrabajador
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        SQL = "DELETE FROM incidenciasgeneradas where entradamarcaje=" & adodc1.Recordset!Entrada
        conn.Execute SQL
        
        SQL = "DELETE FROM entradamarcajes where idmarcaje=" & adodc1.Recordset!Entrada
        conn.Execute SQL
        
        SQL = "DELETE FROM marcajes where entrada=" & adodc1.Recordset!Entrada
        conn.Execute SQL
        
        If SituarDataTrasEliminar(adodc1, NumRegElim) Then
            EnlazaAdo True
            PonerCampos
        Else
            EnlazaAdo False
            LimpiarCampos
            PonerModo 0
        End If
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub cmdAceptar_Click()
Dim vM As CMarcajes

    Select Case Modo
         Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3, 4 'INSERTAR
       
    End Select
End Sub


Private Sub cmdCancelar_Click()
    On Error Resume Next

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            If Me.adodc1.Recordset.EOF Then
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

    If Err.Number <> 0 Then Err.Clear
End Sub






Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        'Si vienen datos o no
        'cargaremos o no
        'Qu carge a vacio
       ' DoEvents
       ' If CadenaDesdeOtroForm = "" Then
            CadB = " false "
       ' Else
       '     CadB = CadenaDesdeOtroForm
       ' End If
        
        CargaDatos
        'If Not adodc1.Recordset.EOF Then
        '    PonerCampos
        '    PonerModo 2
        'Else
            
            PonerModo 0
           ' BotonBuscar
       ' End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
    ' ICONITOS DE LA BARRA
    btnPrimero = 15 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        '14 y 15 separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

    'cargar IMAGES de busqueda
    Me.imgZoom(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgZoom(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
   ' Me.imgZoom(2).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    'Las imagenes
    Me.ListView1.SmallIcons = frmPpal.ImageListReloj
    Me.ListView2.SmallIcons = frmPpal.ImageListReloj
    Me.ListView3.SmallIcons = frmPpal.ImageListReloj

    'cargar IMAGE de mail
 '   Me.imgMail(0).Picture = frmPpal.imgListImages16.ListImages(2).Picture

    LimpiarCampos   'Limpia los campos TextBox
    'Vemos como esta guardado el valor del check
   ' chkVistaPrevia.Value = CheckValueLeer(Name)

    Me.Check3.Visible = vEmpresa.HorarioNocturno2

    CargaCombo (0)

    'FrameEspecial2.Visible = vEmpresa.TodosLosDias
    

    If MostrarUnosDatos <> 0 Then CadenaDesdeOtroForm = " entrada =" & MostrarUnosDatos
    
    TratarOrdenacion True
    SeparaValores
    LimpiarCampos
    PrimeraVez = True
    CadB = ""

End Sub



Private Sub TratarOrdenacion(Leer As Boolean)
Dim Traba As Boolean
Dim I As Integer

    If Leer Then
        
        Traba = Dir(App.Path & "\Ordtra.dat", vbArchive) = ""
        Me.mnTrabajador.Checked = Traba
        Me.mnFecha.Checked = Not Traba
    
    Else
        Traba = Me.mnTrabajador.Checked
        If Traba Then
            If Dir(App.Path & "\Ordtra.dat", vbArchive) = "" Then
                I = FreeFile
                Open App.Path & "\Ordtra.dat" For Output As #I
                Print #I, Now
                Close #I
                
            End If
        Else
            'Fecha
            If Dir(App.Path & "\Ordtra.dat", vbArchive) <> "" Then Kill App.Path & "\Ordtra.dat"
            
        End If
    End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
 '   CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    TratarOrdenacion False
End Sub




Private Sub frmB_Selecionado(CadenaDevuelta As String)
    
    Select Case Me.imgZoom(0).Tag
    Case 0
        'Trabajador
        Me.Text1(1).Text = RecuperaValor(CadenaDevuelta, 1)
        Me.Text2.Text = RecuperaValor(CadenaDevuelta, 2)
    Case 1
        'Incidencias
        Me.Text1(2).Text = RecuperaValor(CadenaDevuelta, 1)
        Me.Text3.Text = RecuperaValor(CadenaDevuelta, 2)
        
        
    Case 2
        Text1(7).Text = RecuperaValor(CadenaDevuelta, 1)
        txtHorario(0).Text = RecuperaValor(CadenaDevuelta, 2)
        
    End Select
    'If CadenaDevuelta <> "" Then
    '
    '    Screen.MousePointer = vbHourglass
    '    'Sabemos que campos son los que nos devuelve
    '    'Creamos una cadena consulta y ponemos los datos
    '    CadB = ""
    '    Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
    '    CadB = Aux
    '    '   Como la clave principal es unica, con poner el sql apuntando
    '    '   al valor devuelto sobre la clave ppal es suficiente
    '    ' *** canviar o llevar el WHERE ***
    '    'CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
    '    ' **********************************
    '    'PonerCadenaBusqueda
    '    'Screen.MousePointer = vbDefault
    'End If
End Sub









Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim Obj As Object
    If Modo = 4 Then Exit Sub
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
    imgFec(3).Tag = Index '<===

    If Text1(Index).Text <> "" Then frmc.NovaData = Text1(Index).Text


    frmc.Show vbModal
    Set frmc = Nothing

    PonerFoco Text1(CByte(imgFec(3).Tag)) '<===

End Sub

Private Sub imgInciGen_Click(Index As Integer)
Dim Carga As Boolean
    If Index > 0 Then
        If ListView2.ListItems.Count = 0 Then Exit Sub
        If ListView2.SelectedItem Is Nothing Then
            MsgBox "Seleccione una incidencia", vbExclamation
            Exit Sub
        End If
    End If
    Carga = False
    If Index < 2 Then
        If Index = 0 Then
            CadenaDesdeOtroForm = ""
        Else
            With ListView2.SelectedItem
                CadenaDesdeOtroForm = .Text & "|"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & .SubItems(1) & "|"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & .SubItems(2) & "|"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & .SubItems(3) & "|"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & .Tag & "|"
            End With
        End If
        NumRegElim = Val(Text1(0).Text)
        frmListado.Opcion = 3
        frmListado.Show vbModal
        If CadenaDesdeOtroForm <> "" Then Carga = True
            
    Else
        'Eliminar.
        With ListView2.SelectedItem
            CadenaDesdeOtroForm = "Desea eliminar la incidencia generada: " & vbCrLf & vbCrLf
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & .SubItems(3) & " - " & .Text & vbCrLf
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Horas: " & .SubItems(1) & " (" & .SubItems(2) & ")" & vbCrLf
            If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbYes Then CadenaDesdeOtroForm = ""
            If CadenaDesdeOtroForm = "" Then
                CadenaDesdeOtroForm = "Delete from incidenciasgeneradas where id =" & .Tag
                If EjecutaSQL(CadenaDesdeOtroForm) Then Carga = True
            End If
            CadenaDesdeOtroForm = ""
        End With
    End If
    
    'Volver a cargar ls incidencias generadas
    If Carga Then
        ListView2.ListItems.Clear
        CargaDatosMarcajes True
    End If
End Sub

Private Sub imgZoom_Click(Index As Integer)
Dim Cad As String
Dim LeerHorario As Boolean

    If Index = 2 Then
        If Me.Text1(3).Text = "" Then
            MsgBox "Ponga primero la fecha", vbExclamation
            Exit Sub
        End If
    End If

    Set frmB = New frmBuscaGrid
    imgZoom(0).Tag = Index
    Select Case Index
    Case 0
        'TRABAJADORES
            If Text1(1).Locked Then
                Set frmB = Nothing
                Exit Sub
            End If
    
            'Llamamos a al form
            ' **************** arreglar-ho per a vore lo que es desije ****************
            'Cod Diag.|idDiag|N|Formato|10·
            Cad = "Codigo|idtrabajador|N||20·"
            Cad = Cad & "Nombre|nomtrabajador|T||60·"
            Cad = Cad & "Tarjeta|numtarjeta|N||20·"
            frmB.vCampos = Cad
            frmB.vTabla = "trabajadores"
            frmB.vSQL = ""
            
            '###A mano
            frmB.vTitulo = "Trabajadores"
            
        
    Case 1
            'INCIDENCIAS
            
            'Cod Diag.|idDiag|N|Formato|10·
            Cad = "Codigo|idinci|N||20·"
            Cad = Cad & "Descripcion|nominci|T||70·"
            frmB.vCampos = Cad
            frmB.vTabla = "incidencias"
            frmB.vSQL = ""
            
            '###A mano
            
            frmB.vTitulo = "Incidencias"
    
    
    
    Case 2
            'HORARIO
    
            Cad = "Codigo|idhorario|N||20·"
            Cad = Cad & "Descripcion|NomHorario|T||70·"
            frmB.vCampos = Cad
            frmB.vTabla = "Horarios"
            frmB.vSQL = ""
            
            '###A mano
            
            frmB.vTitulo = "Horarios"
    
    
    
    End Select
    frmB.vDevuelve = "0|1|"
    frmB.vSelElem = 0
    frmB.Show vbModal
    Set frmB = Nothing

    If Index > 1 Then
'        If Text1(7).Text = "" Then Exit Sub
'        LeerHorario = False
'        If Format(vH.Fecha, "dd/mm/yyyy") <> Text1(3).Text Then
'            LeerHorario = True
'        Else
'            If vH.IdHorario <> Text1(7).Text Then LeerHorario = True
'        End If
'        If LeerHorario Then
'            LeerHorario = vH.Leer(CInt(Text1(7).Text), Text1(3).Text, 0) = 0
'            PonerHorario LeerHorario
'        End If
    End If
End Sub

Private Sub ListView3_DblClick()
    If Modo <> 2 Then Exit Sub
    If ListView3.SelectedItem Is Nothing Then Exit Sub
    If ListView3.SelectedItem.SubItems(2) <> "*" Then Exit Sub
    AbrirGeolocalizacion
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub



Private Sub mnEliminar_Click()
    BotonEliminar
End Sub



Private Sub mnFecha_Click()
    If Not mnFecha.Checked Then
        mnFecha.Checked = True
        mnTrabajador.Checked = False
        CheckOrden
    End If
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub

    'Preparar para modificar
    '-----------------------
    'If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
    BotonModificar
End Sub





Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub CheckOrden()

    
   ' If InStr(1, adodc1.Recordset.Source) > 0 Then
        CargaDatos
        If Not adodc1.Recordset.EOF Then PonerCampos
   ' Else
   '     BotonVerTodos
   ' End If

End Sub

Private Sub mnTrabajador_Click()
    CheckOrden
End Sub

Private Sub mnTraspasar_Click()
    CadenaDesdeOtroForm = ""
    frmListado.Opcion = 23
    frmListado.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        TraspasarDatosHco
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    Indice = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 6: KEYBusqueda KeyAscii, 0 'poblacion
                Case 18: KEYBusqueda KeyAscii, 1 'banco
            End Select
        End If
    Else
        KeyPress KeyAscii
    End If
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim s As Single
Dim BuscarHorario As Boolean
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    Text1(Index).Text = Trim(Text1(Index).Text)
    BuscarHorario = False
    Select Case Index
    
        Case 1 'codigo trab
            If Text1(1).Text = "" Then
                SQL = ""
                BuscarHorario = True
            Else
                If Not PonerFormatoEntero(Text1(1)) Then
                    Text1(1).Text = ""
                    SQL = ""
                    BuscarHorario = True
                Else
                    SQL = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idtrabajador", Text1(1).Text, "N")
                    If SQL = "" Then
                        MsgBox "No existe el trabajador: " & Text1(1).Text, vbExclamation
                        Text1(1).Text = ""
                        PonerFoco Text1(1)
                    End If
                    If Text1(3).Text <> "" Then BuscarHorario = True
                End If
            End If
            Text2.Text = SQL
        
        Case 2  'incdi
            If Text1(2).Text = "" Then
                Text3.Text = ""
                Exit Sub
            End If
            
            If PonerFormatoEntero(Text1(2)) Then
                SQL = DevuelveDesdeBD("nominci", "incidencias", "idinci", Text1(2).Text, "N")
                If SQL = "" Then
                    MsgBox "No existe la incidencia: " & Text1(2).Text, vbExclamation
                    Text1(2).Text = ""
                    PonerFoco Text1(2)
                End If
            Else
                Text1(2).Text = ""
                SQL = ""
            End If
            Text3.Text = SQL
        
        Case 3
            'Fecha
            If Text1(3).Text = "" Then Exit Sub
            If Not EsFechaOK(Text1(3)) Then
                MsgBox "Fecha incorrecta: " & Text1(3).Text, vbExclamation
                Text1(3).Text = ""
                PonerFoco Text1(3)
            End If
            If Text1(1).Text <> "" Then BuscarHorario = True
        Case 4, 5, 6
        
            If Text1(Index).Text = "" Then
                txtDec(Index - 4).Text = ""
                Exit Sub
            End If
        
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Campo numérico: " & Text1(Index).Text, vbExclamation
                
                Text1(Index).Text = "0"
                PonerFoco Text1(Index)
            End If
                
            
            Text1(Index).Text = TransformaPuntosComas(Text1(Index).Text)
            s = CSng(Text1(Index).Text)
            txtDec(Index - 4).Text = Format(DevuelveHora(s), "hh:mm")

            If Modo = 4 Then
                
        
            
            End If
  
    End Select
    
    If BuscarHorario Then BuscarHorarioMarcaje
    
    
End Sub


Private Sub BuscarHorarioMarcaje()
Dim B As Boolean
'    If Modo = 3 Then
'            If Text1(1).Text = "" Or Text1(3).Text = "" Then
'                'PonerHorario False
'            Else
'                Set miRsAux = New ADODB.Recordset
'                SQL = "Select * from calendariothco where idtrabajador = " & Text1(1).Text
'                SQL = SQL & " AND Fecha ='" & Format(Text1(3).Text, FormatoFecha) & "'"
'                miRsAux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'                If Not miRsAux.EOF Then
'                    If Text1(7).Text <> "" Then
'                        If Val(Text1(7).Text) <> miRsAux!IdHorario Then Text1(7).Text = ""
'                    End If
'                    If Text1(7).Text = "" Then
' '                       B = vH.Leer(miRsAux!IdHorario, Text1(3).Text, 0) = 0
'                        If B Then Text1(7).Text = miRsAux!IdHorario
'                        PonerHorario B
'                    End If
'                End If
'                miRsAux.Close
'                Set miRsAux = Nothing
'            End If
'    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
                BotonBuscar
        Case 3
                BotonVerTodos
        Case 6
                BotonAnyadir
        Case 7
                mnModificar_Click
        Case 8
                BotonEliminar
        Case 11 'Imprimir
                'AbrirListado (2)  'OpcionListado=2 Formas de pago
                printNou
        Case 13 'Salir
                mnSalir_Click
                
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
   
    
    DesplazamientoData adodc1, Index
    If adodc1.Recordset.EOF Then
        LimpiarCampos
    Else
        EnlazaAdo True
        PonerCampos
    End If
    
End Sub



Private Sub PonerCampos()
Dim RS As ADODB.Recordset
Dim LeerHorario As Boolean
    
    If adodc1.Recordset.EOF Then Exit Sub

    
    
    'Campos
    With Adodc2.Recordset
        Text1(0).Text = !Entrada
        Text1(1).Text = !idTrabajador
        Text1(2).Text = !IncFinal
        Text1(3).Text = !Fecha
        Text1(4).Text = !HorasTrabajadas
        Text1(5).Text = !HorasIncid
        Text1(6).Text = !HorasDto
        If !Correcto Then
            chkCorrecto.Value = 1
        Else
            chkCorrecto.Value = 0
        End If
        Text2.Text = !nomtrabajador2
        Text3.Text = !nominci2
        txtHorario(0).Text = !nomcalendar
        
        txtDec(0).Text = Format(DevuelveHora(!HorasTrabajadas), "hh:mm")
        txtDec(1).Text = Format(DevuelveHora(!HorasIncid), "hh:mm")
        txtDec(2).Text = Format(DevuelveHora(!HorasDto), "hh:mm")
        
        If !Festivo = 0 Then
            Me.Check1.Value = 0
        Else
            Me.Check1.Value = 1
        End If
        
        If vEmpresa.HorarioNocturno2 Then
            If Val(!Nocturno) = 0 Then
                Me.Check3.Value = 0
            Else
                Me.Check3.Value = 1
            End If
        End If
    End With
    
    
    'Ponemos los marcajes, ticajes etc etc
    '-------------------------------------
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    CargaDatosMarcajes False
    CalculaHoras
    
    LeerHorario = False
'    If vH.IdHorario <> Val(Adodc2.Recordset!IdHorario) Then
'        LeerHorario = True
'    Else
'        If vH.Fecha <> CDate(Adodc2.Recordset!Fecha) Then LeerHorario = True
'    End If
    
    If LeerHorario Then
        lblIndicador.Caption = "Leyendo horario ..."
        Me.Refresh
        DoEvents
        
      '  Indice = vH.Leer(Adodc2.Recordset!IdHorario, Adodc2.Recordset!Fecha, Adodc2.Recordset!idCal)
        PonerHorario Indice = 0
    Else
        lblIndicador.Caption = ""
    End If
'    Text1(7).Text = vH.IdHorario
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = PonerContRegistros(Me.adodc1)
End Sub

Private Sub PonerHorario(OK As Boolean)


'
'    If OK Then
''        txtHorario(0).Text = vH.NomHorario
''        txtHorario(1).Text = vH.HoraE1
''        txtHorario(2).Text = vH.HoraS1
''        txtHorario(3).Text = vH.HoraE2
''        txtHorario(4).Text = vH.HoraS2
''        txtHorario(5).Text = vH.TotalHoras
''        txtHorario(6).Text = vH.NumTikadas
'    Else
'        txtHorario(0).Text = "Error leyendo horario"
'        txtHorario(1).Text = ""
'        txtHorario(2).Text = ""
'        txtHorario(3).Text = ""
'        txtHorario(4).Text = ""
'        txtHorario(5).Text = ""
'        txtHorario(6).Text = ""
'    End If
End Sub



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function

Private Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then
        If Modo = 0 Or Modo = 2 Then Unload Me 'ESC
    End If
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
'    imgBuscar_Click (indice)
End Sub


' ********** SI N'HI HAN COMBOS *****************************


Private Sub CargaCombo(Index As Integer)
'Dim SQL As String
'Dim RS As ADODB.Recordset
'
'    Combo1(Index).Clear
'
'    Select Case Index
'        Case 0 'IBAN PAIS BANCOS
'            SQL = "SELECT * FROM naciones WHERE ibanpais <> """" ORDER BY ibanpais"
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'            While Not RS.EOF
'                Combo1(Index).AddItem RS!ibanPais
'                Combo1(Index).ItemData(Combo1(Index).NewIndex) = RS!codNacio
'                RS.MoveNext
'            Wend
'            RS.Close
'            Set RS = Nothing
'    End Select
End Sub



Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        'CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' ******** Si la clau primaria no es Text1(0), canviar-ho ***********
        PonerFoco Text1(1)
        ' *******************************************************************
    End If
End Sub



Private Sub LimpiarCampos()

    On Error Resume Next

    Limpiar Me
   ' Me.Combo1(0).ListIndex = -1
    
    ' ****************************************************
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    If Err.Number <> 0 Then Err.Clear
End Sub

' ***** SI N'HI HAN BOTONS I CAMPS DE BUSCAR EN ATRES FORMULARIS ********



Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codguiav=" & Text1(0).Text & ")"
    If SituarData(Me.adodc1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub



Private Sub printNou()
    
'    With frmImprimir2
'        .cadTabla2 = "guiaviaj"
'        .Informe2 = "rGuiasViaje.rpt"
'        If CadB <> "" Then
'            .cadRegSelec = SQL2SF(CadB)
'        Else
'            .cadRegSelec = ""
'        End If
'        .cadRegActua = POS2SF(adodc1, Me)
'        .cadTodosReg = ""
'        .OtrosParametros2 = "pEmpresa=" & DBSet(vEmpresa.nomEmpre, "T") & "|" '& "'|pOrden={forpagos.desforpa}|"
'        .NumeroParametros2 = 1
'        .MostrarTree2 = False
'        .InfConta2 = False
'        .ConSubInforme2 = False
'
'        .Show vbModal
'    End With


    frmListado.Opcion = 24
    frmListado.Show vbModal

End Sub

Private Sub SeparaValores()
Dim Cad As String


    Cad = RecuperaValor(CadenaDesdeOtroForm, 1)
    If Cad = "" Then Cad = "01/01/1900"
    FI = CDate(Cad)
    
    Cad = RecuperaValor(CadenaDesdeOtroForm, 2)
    If Cad = "" Then Cad = "01/01/2900"
    FF = CDate(Cad)
    
    Cad = RecuperaValor(CadenaDesdeOtroForm, 3)
    If Cad = "" Then Cad = "-1"
    DTra = Val(Cad)
    
    Cad = RecuperaValor(CadenaDesdeOtroForm, 4)
    If Cad = "" Then Cad = "100000000"
    HTra = Val(Cad)

    Cad = RecuperaValor(CadenaDesdeOtroForm, 5)
    If Cad = "" Then Cad = "-1"
    DInci = Val(Cad)
    
    Cad = RecuperaValor(CadenaDesdeOtroForm, 6)
    If Cad = "" Then Cad = "32200"
    HInci = Val(Cad)



    'Correctos o incorrectos
    Cad = RecuperaValor(CadenaDesdeOtroForm, 7)
    If Val(Cad) > 2 Then Cad = "0"
    CorrectosIncorrectos = CByte(Val(Cad))
    
End Sub


Private Sub MontaSQL()
    SQL = "Select entrada from marcajeshco where"
    'LA select
    
    'Public FI As Date
    'Public FF As Date
    'Public DTra As Integer
    'Public HTra As Integer
    'Public CorrectosIncorrectos As Byte  '0.- Ambos  1.- Correctos  2.-Incorrectos
    SQL = SQL & " fecha >='" & Format(FI, FormatoFecha) & "'"
    SQL = SQL & " AND fecha <='" & Format(FF, FormatoFecha) & "'"
    SQL = SQL & " AND idtrabajador >= " & DTra
    SQL = SQL & " AND idtrabajador <= " & HTra
    
    SQL = SQL & " AND incfinal >= " & DInci
    SQL = SQL & " AND incfinal <= " & HInci
    
    
    If CorrectosIncorrectos = 1 Then
        SQL = SQL & " AND correcto = 1"
    Else
        If CorrectosIncorrectos = 2 Then SQL = SQL & " AND correcto = 0"
    End If
    
    If CadB <> "" Then SQL = SQL & " AND " & CadB
    
    
    
    'Ordenacion
    If Me.mnTrabajador.Checked Then
        SQL = SQL & " ORDER BY idtrabajador,fecha"
    Else
        SQL = SQL & " ORDER BY fecha,idtrabajador"
    End If
    
End Sub


Private Sub CargaDatos()
    
    MontaSQL
    Me.adodc1.ConnectionString = conn
    Me.adodc1.RecordSource = SQL
    Me.adodc1.Refresh
    EnlazaAdo Not adodc1.Recordset.EOF

End Sub


Private Sub EnlazaAdo(Si As Boolean)
    
    If Si Then
        SQL = "select marcajeshco.* "
        SQL = SQL & " FROM marcajeshco "
        SQL = SQL & " WHERE entrada = " & adodc1.Recordset!Entrada
    Else
        SQL = "Select * from marcajeshco where false"
    End If
    
    Me.Adodc2.ConnectionString = conn
    Me.Adodc2.RecordSource = SQL
    Me.Adodc2.Refresh
    
    
End Sub


Private Sub CargaDatosMarcajes(SoloIncidenciasGeneradas As Boolean)
Dim RS As ADODB.Recordset
Dim IT As ListItem
Dim I As Integer
Dim Cad As String
Dim FueraIntervaloHoras As Byte   '0.No  1<0    2>=24
    Set RS = New ADODB.Recordset
    If Not SoloIncidenciasGeneradas Then
        SQL = "select hour(hora) lahora,minute(hora) minutos,second(hora) segundos "
        SQL = SQL & ",hour(horareal) lahorar,minute(horareal) minutosr,second(horareal) segundosr"
        SQL = SQL & " ,entradamarcajes.idInci ,nominci,if(hora<'0:00:00',1,0) Negativa,hora horabd,horareal,reloj "
        SQL = SQL & ",latitud,longitud"
        SQL = SQL & " from entradamarcajeshco entradamarcajes where  "
        SQL = SQL & " idmarcaje=" & adodc1.Recordset!Entrada & " ORDER BY horareal,reloj"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
        
            FueraIntervaloHoras = 0
            If RS!Negativa Then
               SQL = Horas_Quitar24(RS!horabd, True)
               
            Else
                If RS!LaHora <= 23 Then
                    I = RS!LaHora
                Else
                    If RS!LaHora > 23 Then I = -24
                    I = RS!LaHora + I
                End If
                SQL = Format(I, "00") & ":" & Format(RS!Minutos, "00") & ":" & Format(RS!segundos, "00")
            
            End If
            
            Set IT = ListView1.ListItems.Add(, , SQL)
            If RS!IdInci > 0 Then IT.SubItems(1) = RS!NomInci
            IT.SubItems(2) = RS!LaHora & ":" & Format(RS!Minutos, "00") & ":" & Format(RS!segundos, "00")
            IT.Tag = RS!Negativa
            
            'Hora real
            '-----------------------
            If RS!Negativa = 1 Then
                SQL = Horas_Quitar24(RS!HoraReal, True)
            Else
            If RS!LaHora <= 23 Then
                I = RS!lahorar
            Else
                I = -24
                I = RS!lahorar + I
            End If
            SQL = Format(I, "00") & ":" & Format(RS!Minutosr, "00") & ":" & Format(RS!Segundosr, "00")
            End If
            Set IT = ListView3.ListItems.Add(, , SQL)
            
            If vEmpresa.Reloj2 > 0 Then
                If DBLet(RS!Reloj, "N") > 0 Then
                    IT.ToolTipText = "Biostar2"
                    IT.SmallIcon = 4
                End If
            End If
            'COORDENADAS
            IT.SubItems(1) = " "
            If Not IsNull(RS!Longitud) And Not IsNull(RS!latitud) Then
                IT.SubItems(1) = TransformaComasPuntos(CStr(RS!latitud)) & "," & TransformaComasPuntos(CStr(RS!Longitud))
                IT.SubItems(2) = "*"
                IT.ToolTipText = Trim(IT.ToolTipText & "    Geolocalizacion")
            End If
            
            
            
            RS.MoveNext
        Wend
        RS.Close
    End If
        
        
    'Las incidencias generadas
    SQL = "Select incidenciasgeneradashco.horas,nominci,id from incidenciasgeneradashco where entradamarcaje ="
    SQL = SQL & Adodc2.Recordset!Entrada & " ORDER BY id"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = ListView2.ListItems.Add(, , RS!NomInci)
        IT.SubItems(1) = DevuelveHora(RS!Horas)
        IT.SubItems(2) = Format(RS!Horas, "0.00")
        IT.SubItems(3) = RS!Id
        IT.Tag = RS!Id
        RS.MoveNext
    Wend
    RS.Close
    
End Sub


Private Sub CalculaHoras()
Dim g As Integer
Dim I As Integer
Dim Horas As Single
Dim v As Single
Dim FInter As Byte
Dim T1 As Single
Dim T2 As Single
g = ListView1.ListItems.Count
If g = 0 Then
    TextHt(0).Text = 0
    TextHt(1).Text = 0
    TextHt(2).Text = 0
    Exit Sub
End If
TextHt(2).Text = g
g = g \ 2
Horas = 0



For I = 1 To g
    
    If ListView1.ListItems((I * 2)).Tag = 1 Then
        FInter = 1
    Else
        FInter = HoraFueraInterval(ListView1.ListItems((I * 2)).SubItems(2))
    End If
    T1 = DevuelveValorHora3(FInter, ListView1.ListItems((I * 2)).SubItems(2))
    
    If ListView1.ListItems((I * 2) - 1).Tag = 1 Then
        FInter = 1
    Else
        FInter = HoraFueraInterval(ListView1.ListItems((I * 2) - 1).SubItems(2))
    End If
    T2 = DevuelveValorHora3(FInter, ListView1.ListItems((I * 2) - 1).SubItems(2))
    v = T1 - T2
    
    'v = DevuelveValorHora(CDate(ListView1.ListItems((I * 2))) - CDate(ListView1.ListItems((I * 2) - 1)))
    Horas = Horas + v
Next I
TextHt(0).Text = Round(Horas, 2)
TextHt(1).Text = DevuelveHora(Horas)

End Sub




Private Sub AbrirGeolocalizacion()
Dim Cad As String
    Cad = "https://www.google.com/maps/?q=" & ListView3.SelectedItem.SubItems(1)
    LanzaVisorMimeDocumento Me.Hwnd, Cad
    
End Sub




Private Sub TraspasarDatosHco()
Dim Cad As String
Dim OK  As Boolean
    
    lblIndicador.Caption = "Marcajes"
    lblIndicador.Refresh

    OK = True

    Cad = "INSERT IGNORE INTO marcajeshco(Entrada , idTrabajador, Fecha, Correcto, IncFinal, HorasTrabajadas, HorasIncid, IdHorario,"
    Cad = Cad & " HorasDto, Festivo, Baja, Nocturno, nomtrabajador2, idcal2 ,nominci2 , nomcalendar   )"

    Cad = Cad & " select Entrada,marcajes.idTrabajador,Fecha,Correcto,IncFinal,HorasTrabajadas,HorasIncid,idHorario,HorasDto,Festivo,"
    Cad = Cad & " Baja,Nocturno,nomtrabajador,trabajadores.idcal,NomInci,calendario.descripcion"
    Cad = Cad & " from  marcajes,incidencias , trabajadores  ,calendario"
    Cad = Cad & " Where marcajes.idTrabajador = trabajadores.idTrabajador And marcajes.IncFinal = incidencias.IdInci"
    Cad = Cad & " and trabajadores.idcal=calendario.idcal"
    Cad = Cad & " and fecha<= " & DBSet(CadenaDesdeOtroForm, "F")
    If Not EjecutaSQL(Cad) Then OK = False


    If OK Then
        DoEvents
        lblIndicador.Caption = "datos horas"
        lblIndicador.Refresh

        Cad = "INSERT IGNORE INTO entradamarcajeshco(Secuencia,idTrabajador,idMarcaje,Fecha,Hora,idInci,HoraReal,Reloj,latitud,longitud,ssid,imei,observaciones,appinfo,nominci)"
        Cad = Cad & " select Secuencia,idTrabajador,idMarcaje,Fecha,Hora,entradamarcajes.idInci,HoraReal,Reloj,"
        Cad = Cad & " latitud,longitud,ssid,imei,observaciones,appinfo,coalesce(nominci ,'')"
        Cad = Cad & " from entradamarcajes left join incidencias on entradamarcajes.idinci=incidencias.idinci where"
        Cad = Cad & " idmarcaje in ( select entrada from marcajes where fecha<=" & DBSet(CadenaDesdeOtroForm, "F") & ")"
        If Not EjecutaSQL(Cad) Then OK = False
    
    End If
    
    If OK Then
        DoEvents
        lblIndicador.Caption = "Incidencias generadas"
        lblIndicador.Refresh

        Cad = "INSERT IGNORE INTO incidenciasgeneradashco( Id,EntradaMarcaje,Incidencia,horas,nominci )"
        Cad = Cad & " select  Id,EntradaMarcaje,Incidencia,horas,coalesce(nominci ,'')"
        Cad = Cad & " from incidenciasgeneradas left join incidencias on incidenciasgeneradas.Incidencia=incidencias.idinci where"
        Cad = Cad & " EntradaMarcaje in ( select entrada from marcajes where fecha<=" & DBSet(CadenaDesdeOtroForm, "F") & ")"
        If Not EjecutaSQL(Cad) Then OK = False

    End If
    
    
    If OK Then
         DoEvents
        lblIndicador.Caption = "Borrando"
        lblIndicador.Refresh

    
        'Los deletes
        Cad = " DELETE from incidenciasgeneradas  where"
        Cad = Cad & " EntradaMarcaje in ( select entrada from marcajes where fecha<=" & DBSet(CadenaDesdeOtroForm, "F") & ")"
        If Not EjecutaSQL(Cad) Then OK = False
        Cad = " DELETE from entradamarcajes  where"
        Cad = Cad & " idmarcaje in ( select entrada from marcajes where fecha<=" & DBSet(CadenaDesdeOtroForm, "F") & ")"
        If Not EjecutaSQL(Cad) Then OK = False
        Cad = "DELETE FROM marcajes where fecha<=" & DBSet(CadenaDesdeOtroForm, "F")
        If Not EjecutaSQL(Cad) Then OK = False
        
        Cad = "DELETE FROM calendariof where fecha<=" & DBSet(CadenaDesdeOtroForm, "F")
        EjecutaSQL Cad
        Cad = "DELETE FROM calendariol where fecha<=" & DBSet(CadenaDesdeOtroForm, "F")
        EjecutaSQL Cad
        Cad = "DELETE FROM calendariot where fecha<=" & DBSet(CadenaDesdeOtroForm, "F")
        EjecutaSQL Cad
        Cad = "DELETE FROM festivos where fecha<=" & DBSet(CadenaDesdeOtroForm, "F")
        EjecutaSQL Cad
        
        

    End If
    
    
    If OK Then
        MsgBox "Proceso finalizado correctamente", vbInformation
    Else
        Cad = "ERROR : " & lblIndicador.Caption & vbCrLf & Cad & vbCrLf & "Avise soporte tecnico"
        MsgBox Cad, vbExclamation
    End If
    lblIndicador.Caption = ""
End Sub
