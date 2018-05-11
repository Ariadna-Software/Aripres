VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHorario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Horarios"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   Icon            =   "frmHorario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   645
      Index           =   0
      Left            =   240
      TabIndex        =   56
      Top             =   480
      Width           =   11295
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   600
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "id|N|N|1|9999|horarios|IdHorario|0000|S|"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   8760
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "Horas|N|N|||horarios|totalhoras|||"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Descripcion|T|N|||horarios|nomhorario|||"
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Horas"
         Height          =   255
         Index           =   0
         Left            =   7800
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   58
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cód."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   255
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   54
      Top             =   6480
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
         TabIndex        =   55
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10260
      TabIndex        =   53
      Top             =   6720
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9000
      TabIndex        =   52
      Top             =   6720
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3240
      Top             =   6600
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
      Left            =   10260
      TabIndex        =   60
      Top             =   6720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
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
         TabIndex        =   62
         Top             =   120
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   240
      TabIndex        =   63
      Top             =   1320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9022
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Horario semanal"
      TabPicture(0)   =   "frmHorario.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paradas"
      TabPicture(1)   =   "frmHorario.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "Label8"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Rectificación"
      TabPicture(2)   =   "frmHorario.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux0"
      Tab(2).Control(1)=   "Combo1(0)"
      Tab(2).Control(2)=   "Label5"
      Tab(2).ControlCount=   3
      Begin VB.Frame FrameAux0 
         Height          =   4575
         Left            =   -71040
         TabIndex        =   105
         Top             =   240
         Width           =   7095
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   112
            Tag             =   "id|N|N|||modificarfichajes|idhorario||S|"
            Text            =   "Nº exped"
            Top             =   4080
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   240
            MaxLength       =   8
            TabIndex        =   109
            Tag             =   "inicio|H|N|||modificarfichajes|inicio|hh:mm:ss|S|"
            Text            =   "Nº exped"
            Top             =   4080
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   1080
            MaxLength       =   8
            TabIndex        =   110
            Tag             =   "fin|H|N|||modificarfichajes|fin|hh:mm:ss||"
            Text            =   "li"
            Top             =   4080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   111
            Tag             =   "fin|H|N|||modificarfichajes|modificada|hh:mm:ss||"
            Text            =   "nombre"
            Top             =   4080
            Visible         =   0   'False
            Width           =   795
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   0
            Left            =   4680
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
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
            Caption         =   "AdoAux(0)"
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
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   240
            TabIndex        =   106
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
            Begin VB.CheckBox Check2 
               Caption         =   "Vista previa"
               Height          =   195
               Index           =   1
               Left            =   8400
               TabIndex        =   107
               Top             =   120
               Visible         =   0   'False
               Width           =   1215
            End
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmHorario.frx":0060
            Height          =   3855
            Index           =   0
            Left            =   1440
            TabIndex        =   108
            Top             =   480
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               AllowFocus      =   0   'False
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "frmHorario.frx":0078
         Left            =   -74760
         List            =   "frmHorario.frx":008E
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Tag             =   "Rect|N|N|||horarios|Rectificar|||"
         Top             =   840
         Width           =   3375
      End
      Begin VB.Frame Frame3 
         Height          =   4455
         Left            =   120
         TabIndex        =   80
         Top             =   480
         Width           =   10815
         Begin VB.CheckBox CheckF 
            Caption         =   "Festivo"
            Height          =   195
            Index           =   6
            Left            =   1080
            TabIndex        =   45
            Top             =   3960
            Width           =   915
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   33
            Left            =   3000
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   3900
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   34
            Left            =   4200
            TabIndex        =   48
            Text            =   "34"
            Top             =   3900
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   35
            Left            =   6240
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   3900
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   36
            Left            =   7500
            TabIndex        =   50
            Text            =   "40"
            Top             =   3900
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   10020
            TabIndex        =   51
            Text            =   "3"
            Top             =   3900
            Width           =   500
         End
         Begin VB.CheckBox CheckF 
            Caption         =   "Festivo"
            Height          =   195
            Index           =   5
            Left            =   1080
            TabIndex        =   38
            Top             =   3480
            Width           =   915
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   28
            Left            =   3000
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   29
            Left            =   4200
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   30
            Left            =   6240
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   31
            Left            =   7500
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   10020
            TabIndex        =   44
            Text            =   "T"
            Top             =   3420
            Width           =   500
         End
         Begin VB.CheckBox CheckF 
            Caption         =   "Festivo"
            Height          =   195
            Index           =   4
            Left            =   1080
            TabIndex        =   31
            Top             =   3000
            Width           =   915
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   23
            Left            =   3000
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   24
            Left            =   4200
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   25
            Left            =   6240
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   26
            Left            =   7500
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   27
            Left            =   10020
            TabIndex        =   37
            Text            =   "T"
            Top             =   2940
            Width           =   500
         End
         Begin VB.CheckBox CheckF 
            Caption         =   "Festivo"
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   24
            Top             =   2520
            Width           =   915
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   18
            Left            =   3000
            TabIndex        =   26
            Text            =   "18"
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   19
            Left            =   4200
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   20
            Left            =   6240
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   21
            Left            =   7500
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   10020
            TabIndex        =   30
            Text            =   "t"
            Top             =   2460
            Width           =   500
         End
         Begin VB.CheckBox CheckF 
            Caption         =   "Festivo"
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   17
            Top             =   2040
            Width           =   915
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   13
            Left            =   3000
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   14
            Left            =   4200
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   15
            Left            =   6240
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   16
            Left            =   7500
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   10020
            TabIndex        =   23
            Text            =   "T"
            Top             =   1980
            Width           =   500
         End
         Begin VB.CheckBox CheckF 
            Caption         =   "Festivo"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   10
            Top             =   1560
            Width           =   915
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   8
            Left            =   3000
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   1500
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   9
            Left            =   4200
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1500
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   10
            Left            =   6240
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1500
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   11
            Left            =   7500
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1500
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   10020
            TabIndex        =   16
            Text            =   "T"
            Top             =   1500
            Width           =   500
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   10020
            TabIndex        =   9
            Text            =   "T"
            Top             =   1020
            Width           =   500
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   6
            Left            =   7500
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1020
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   5
            Left            =   6240
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1020
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   4
            Left            =   4200
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1020
            Width           =   855
         End
         Begin VB.CheckBox CheckF 
            Caption         =   "Festivo"
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   3
            Top             =   1080
            Width           =   915
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   2280
            TabIndex        =   11
            Text            =   "T"
            Top             =   1500
            Width           =   500
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   2280
            TabIndex        =   18
            Text            =   "T"
            Top             =   1980
            Width           =   500
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   2280
            TabIndex        =   25
            Text            =   "t"
            Top             =   2460
            Width           =   500
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   2280
            TabIndex        =   32
            Text            =   "T"
            Top             =   2940
            Width           =   500
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2280
            TabIndex        =   39
            Text            =   "T"
            Top             =   3420
            Width           =   500
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   2280
            TabIndex        =   46
            Text            =   "3"
            Top             =   3900
            Width           =   500
         End
         Begin VB.TextBox txtaux 
            Height          =   285
            Index           =   3
            Left            =   3000
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1020
            Width           =   855
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   2280
            TabIndex        =   4
            Text            =   "T"
            Top             =   1020
            Width           =   500
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C00000&
            BorderWidth     =   3
            Index           =   3
            X1              =   9840
            X2              =   10680
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C00000&
            BorderWidth     =   3
            Index           =   2
            X1              =   6120
            X2              =   8520
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C00000&
            BorderWidth     =   3
            Index           =   1
            X1              =   2280
            X2              =   5160
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C00000&
            BorderWidth     =   3
            Index           =   0
            X1              =   120
            X2              =   1920
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line1 
            Index           =   15
            X1              =   8580
            X2              =   9900
            Y1              =   4020
            Y2              =   4020
         End
         Begin VB.Line Line1 
            Index           =   14
            X1              =   8580
            X2              =   9900
            Y1              =   3540
            Y2              =   3540
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   8580
            X2              =   9900
            Y1              =   3060
            Y2              =   3060
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   5100
            X2              =   6100
            Y1              =   4020
            Y2              =   4020
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   5100
            X2              =   6100
            Y1              =   3540
            Y2              =   3540
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   5100
            X2              =   6100
            Y1              =   3060
            Y2              =   3060
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   5100
            X2              =   6100
            Y1              =   2580
            Y2              =   2580
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   8580
            X2              =   9900
            Y1              =   2580
            Y2              =   2580
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   5100
            X2              =   6100
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   8580
            X2              =   9900
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   5100
            X2              =   6100
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   8580
            X2              =   9900
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   5100
            X2              =   6100
            Y1              =   1140
            Y2              =   1140
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   8580
            X2              =   9900
            Y1              =   1140
            Y2              =   1140
         End
         Begin VB.Label Label4 
            Caption         =   "Horas/Día"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   4
            Left            =   9840
            TabIndex        =   96
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "DIA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   95
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label4 
            Caption         =   "Salida"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   3
            Left            =   7620
            TabIndex        =   94
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label4 
            Caption         =   "Entrada"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   2
            Left            =   6300
            TabIndex        =   93
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label4 
            Caption         =   "Salida"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   1
            Left            =   4320
            TabIndex        =   92
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label4 
            Caption         =   "Entrada"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   5
            Left            =   3240
            TabIndex        =   91
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label2 
            Caption         =   "HORAS 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   6840
            TabIndex        =   90
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Domingo"
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
            Left            =   120
            TabIndex        =   89
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Sábado"
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
            Left            =   120
            TabIndex        =   88
            Top             =   3480
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Viernes"
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
            Left            =   120
            TabIndex        =   87
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Jueves"
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
            Left            =   120
            TabIndex        =   86
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Miércoles"
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
            Left            =   120
            TabIndex        =   85
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Martes"
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
            Left            =   120
            TabIndex        =   84
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Lunes"
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
            Left            =   120
            TabIndex        =   83
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "HORAS 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   3660
            TabIndex        =   82
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label4 
            Caption         =   "Dias/Nómina"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   6
            Left            =   2280
            TabIndex        =   81
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4155
         Index           =   1
         Left            =   -74880
         TabIndex        =   65
         Top             =   480
         Width           =   10935
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   102
            Text            =   "Text2"
            Top             =   2865
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   3000
            TabIndex        =   101
            Text            =   "Text2"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   79
            Tag             =   "Hora dto merienda|H|S|||horarios|horadtomer|hh:mm||"
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   78
            Tag             =   "Descuento merienda|N|S|||horarios|dtomer|0.00||"
            Top             =   2865
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   77
            Tag             =   "Horas|H|S|||horarios|horadtoalm|hh:mm||"
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   76
            Tag             =   "Dto alm|N|S|||horarios|dtoalm|0.00||"
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Sexagesimal"
            Height          =   255
            Index           =   7
            Left            =   3000
            TabIndex        =   100
            Top             =   2400
            Width           =   1035
         End
         Begin VB.Label Label11 
            Caption         =   "Decimal"
            Height          =   255
            Index           =   6
            Left            =   1920
            TabIndex        =   99
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Sexagesimal"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   98
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label Label11 
            Caption         =   "Decimal"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   97
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Almuerzo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   75
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Merienda"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   270
            Index           =   1
            Left            =   180
            TabIndex        =   74
            Top             =   1920
            Width           =   1050
         End
         Begin VB.Label Label10 
            Caption         =   "Descuento "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   73
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label10 
            Caption         =   "Hora "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   72
            Top             =   1440
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Hora a partir de la cual NO se contabilizará el almuerzo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4680
            TabIndex        =   71
            Top             =   1560
            Width           =   5715
         End
         Begin VB.Label Label11 
            Caption         =   "Descuento en minutos que se descontarán por el almuerzo. Cero(0) es no descuento almuerzo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   4680
            TabIndex        =   70
            Top             =   720
            Width           =   5835
         End
         Begin VB.Label Label10 
            Caption         =   "Descuento "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   69
            Top             =   2880
            Width           =   1155
         End
         Begin VB.Label Label10 
            Caption         =   "Hora "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   68
            Top             =   3360
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Hora a partir de la cual se contabilizará la merienda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   4680
            TabIndex        =   67
            Top             =   3460
            Width           =   5715
         End
         Begin VB.Label Label11 
            Caption         =   "Descuento en minutos que se descontarán por la merienda. Cero(0) es no descuento almuerzo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   5
            Left            =   4680
            TabIndex        =   66
            Top             =   2520
            Width           =   5835
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            Index           =   0
            X1              =   1440
            X2              =   10800
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Index           =   1
            X1              =   1440
            X2              =   10740
            Y1              =   2040
            Y2              =   2040
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Rectificacion de los marcajes"
         Height          =   255
         Left            =   -74760
         TabIndex        =   104
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   -70200
         TabIndex        =   64
         Top             =   4320
         Width           =   3555
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
Attribute VB_Name = "frmHorario"
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

Dim ModoLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas
Dim NumTabMto As Integer

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
Dim Indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos

Dim RS As ADODB.Recordset
Dim CadB As String
Dim i As Integer

Dim HanCambiadoSubHorarios As Boolean

Private Sub CheckF_Click(Index As Integer)
Dim v As Integer
    If CheckF(Index).Value = 1 Then
        i = (Index * 5) + 3
        For v = i To i + 4
            txtAux(v).Text = ""
        Next v
        'Recalculamos la horas
        v = (5 * Index) + 3
        If v > 0 Then CalculaHorasDia v, 0, False
    End If
End Sub

Private Sub CheckF_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            'SI NO LLEVA LABORAL tiene oculto los campos de dia nomina
            If Not vEmpresa.laboral Then
                For i = 38 To 44
                    txtAux(i).Text = "1"
                Next i
            End If
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    SubHorariosAbd Val(Text1(0).Text)
                    'situarnos en el registro que acabamos de insertar
                    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE idhorario=" & Text1(0).Text & Ordenacion
                    PosicionarData
                    
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    If HanCambiadoSubHorarios Then SubHorariosAbd Data1.Recordset!IdHorario
                    PosicionarData
                End If
            End If
            
        Case 5 'LLINIES
            Select Case ModoLineas
                Case 1 'afegir llinia
                    InsertarLinea
                Case 2 'modificar llinies
                    ModificarLinea
                    If ModoLineas = 0 Then PosicionarData
            End Select

    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

' *** adrede per ad este manteniment ***
Private Sub Combo1_Click(Index As Integer)
   
End Sub
'***************************************+

'Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '  CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
End Sub

Private Sub Form_Load()


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
      
    'ICONITOS DE LOS GRIDS DE LINEAS
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            '.ImageList = frmPpal.imgListComun_VELL
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
      
    Me.Label4(6).Visible = vEmpresa.laboral
    For i = 38 To 44
        txtAux(i).Visible = vEmpresa.laboral
    Next i
      
    LimpiarCampos   'Limpia los campos TextBox
    
    ' *** canviar el nom de la taula i la clau primaria de l'ORDER BY ***
    NombreTabla = "horarios"
    Ordenacion = " ORDER BY idhorario"
    ' **********************************************************
        
    'Vemos como esta guardado el valor del check
    'chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where idhorario=-1"
    Data1.Refresh
          
    CargaGrid 0, False
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow 'codclien
    End If
End Sub


Private Sub LimpiarCampos()

    
    On Error Resume Next

    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    For i = 0 To CheckF.Count - 1
        CheckF(i).Value = 0
    Next i
    Me.Combo1(0).ListIndex = -1
    
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim NumReg As Byte
Dim B As Boolean

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
    B = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
  

    
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
    Frame3.Enabled = Modo > 2
        
        
    'Bloquear los campos interiores
    B = (Modo = 3 Or Modo = 4) Or (Modo = 5 And ModoLineas = 0)
    BloquearRestoCampos B
    'Bloquea los campos Text1 si no estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    
    ' **************************************************************************************************
    ' *** Revisar lo que bloquejem, el nom de la clau primaria, les imagens de buscar i les de dates ***
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    'PosicionarCombo Combo1(2), 724
'    For i = 0 To Combo1(2).ListCount - 1
'        If Combo1(2).ItemData(i) = 724 Then
'            Combo1(2).ListIndex = i
'            Exit For
'        End If
'    Next i
    
    
      
    
    
    If Modo = 4 Then _
        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    
    'BloquearImgBuscar Me, Modo
    'BloquearImgFec Me, 27, Modo
    'BloquearImgFec Me, 28, Modo
    'BloquearImgFec Me, 29, Modo
    ' **************************************************************************************************
                
                
    B = (Modo = 4) Or (Modo = 2)
    For i = 0 To DataGridAux.Count - 1
        DataGridAux(i).Enabled = B
    Next i
               
    If (Modo < 2) Or (Modo = 3) Then
        For i = 0 To DataGridAux.Count - 1
            CargaGrid i, False
        Next i
    End If
                
                
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
Dim B As Boolean
Dim Baux As Boolean
   
' ******** comentar o descomentar depenent de si n'hi ha menú desplegable o no ****
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    B = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(4).Enabled = B
    Me.mnVerTodos.Enabled = B
    'Insertar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
   
    'Modificar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(9).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = b Or (Modo = 0)
    Toolbar1.Buttons(12).Enabled = B
    ' ********************************************************************************
    B = (Modo = 4 Or Modo = 2)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = B
        If B Then Baux = (B And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = Baux
        ToolAux(i).Buttons(3).Enabled = Baux
    Next i


    

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





Private Sub Text2_GotFocus(Index As Integer)
    ConseguirFoco Text2(Index), Modo
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Modo < 3 Then Exit Sub
    
    i = InStr(1, Text2(Index).Text, ".")
    If i > 0 Then Text2(Index).Text = Mid(Text2(Index).Text, 1, i - 1) & ":" & Mid(Text2(Index).Text, i + 1)
    If Right(Text2(Index), 1) = ":" Then Text2(Index).Text = Text2(Index).Text & "00"
    
    If IsDate(Text2(Index).Text) Then Text2(Index).Text = Format(Text2(Index).Text, "hh:mm")

    
    
    i = 3
    If Index = 1 Then i = 5
    If Not IsDate(Text2(Index).Text) Then
        Text2(0).Text = ""
        Text1(i).Text = ""
    Else
        Text1(i).Text = DevuelveValorHora(CDate(Text2(Index).Text))
    End If
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
'            TerminaBloquear
            BotonAnyadirLinea Index
        Case 2
'            TerminaBloquear
            BotonModificarLinea Index
        Case 3
            TerminaBloquear
            BotonEliminarLinea Index
            If Modo = 4 Then
                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            End If
        Case 6 'Imprimir
 '           BotonImprimirLinea Index
    End Select

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
        cad = cad & ParaGrid(Text1(1), 50, "Descripcion")
        cad = cad & ParaGrid(Text1(2), 40, "Horas")
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
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & Text1(J).Text & "|"
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
    Text1(0).Text = SugerirCodigoSiguienteStr("horarios", "idhorario")
    FormateaCampo Text1(0)
    
    
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
    HanCambiadoSubHorarios = False
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    'Comprobamos si se puede eliminar
    If Not SePuedeEliminar Then Exit Sub

    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub

    ' *************** canviar els noms, els formats i el DELETE ****************                  "
    cad = cad & "¿Seguro que desea eliminar el horario?"
    cad = cad & vbCrLf & "  Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "  Descripción: " & Data1.Recordset.Fields(1)
    
    
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
    'Text2(30).Text = PonerNombreDeCod(Text1(30), "tiposemp", "desemple", "tipemple", "N")
    ' *******************************************************************
    PonerSubHorarios
    
    i = Data1.Recordset!Rectificar
    PosicionarCombo Combo1(0), i

    
    'Los datagrid
    For i = 0 To DataGridAux.Count - 1
        CargaGrid i, True
        'Poner Formato campos de la Lineas
'        If Not AdoAux(i).Recordset.EOF Then _
'            PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i

    
    
    
    Modo = 3
    Text1_LostFocus 3 'Dto alm
    Text1_LostFocus 5 'dto mer
    Modo = 2
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = PonerContRegistros(Me.Data1)
    PonerModoOpcionesMenu (Modo)
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
            
            
            
        Case 5
            Select Case ModoLineas
                Case 1 'afegir llinia
                    ModoLineas = 0
                    DataGridAux(NumTabMto).AllowAddNew = False
                    SituarTab (NumTabMto)
                    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    'If DataGridAux(NumTabMto).Enabled Then DataGridAux(NumTabMto).SetFocus
                    DataGridAux(NumTabMto).Enabled = True
                    DataGridAux(NumTabMto).SetFocus

                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llinies
                    ModoLineas = 0
                    SituarTab (NumTabMto)
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                '        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 1 es el nº de llinia
                '       AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                    End If
'                    Select Case NumTabMto
'                        Case 0 'Cuentas bancarias
'                            BloquearTxt txtAux(11), True
'                            BloquearTxt txtAux(12), True
'                        Case 1 'departamentos
'                            For i = 21 To 24
'                                BloquearTxt txtAux(i), True
'                            Next i
'                    End Select
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select

            TerminaBloquear
            PosicionarData
        
    End Select
    ' *************************************************************************
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim J As Integer
Dim valor As Integer
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    ' ******************** canviar els arguments de la funcio i el mensage ****************
    If (Modo = 3) Then 'Insertar
        'comprobar si ya existe el campo de clave primaria
        If ExisteCP(Text1(0)) Then B = False
        
'         Datos = DevuelveDesdeBD("codemple", "empleado", "codemple", Text1(0).Text, "N")
'         If Datos <> "" Then
'            MsgBox "Ya existe el Código de Empleado: " & Text1(0).Text, vbExclamation
'            DatosOk = False
'            PonerFoco Text1(0)
'            Exit Function
'         End If
    End If
    ' *************************************************************************************
         
         
    If Not B Then Exit Function
    
    
    
    'Comrpobaciones horas etc etc
    valor = 0
    For i = 0 To 6
        If CheckF(i).Value = 0 Then
            J = (i * 5) + 3
            'Entrada 1
            If Not FechaOk(txtAux(J).Text) Then
                valor = 1
                Exit For
            End If
            'Salida 1
            If Not FechaOk(txtAux(J + 1).Text) Then
                valor = 2
                Exit For
            End If
            'Entrada 2
            If Not FechaOk(txtAux(J + 2).Text) Then
                valor = 3
                Exit For
            End If
            'Salida 2
            If Not FechaOk(txtAux(J + 3).Text) Then
                valor = 4
                Exit For
            End If
            
            
            If txtAux(J + 4).Text = "" Then
                valor = 5
            Else
                If Not IsNumeric(txtAux(J + 4).Text) Then valor = 6
            End If
            If valor > 0 Then Exit For
                
            
        End If
    Next i
    
    'Error
    If valor > 0 Then
        If valor < 5 Then
            miSQL = "La hora de "
            If (valor Mod 2) = 1 Then
                miSQL = miSQL & "entrada"
            Else
                miSQL = miSQL & "salida"
            End If
            If valor > 2 Then
                miSQL = miSQL & " por la tarde"
            Else
                miSQL = miSQL & " por la mañana"
            End If
        Else
            miSQL = "La cantidad de horas/dias"
        End If
        'La semana del 1 al 7 de mayo de 2006 es lunes a domingo
        If i < 7 Then miSQL = miSQL & " del " & Format(i + 1 & "/05/2006", "dddd")
        miSQL = miSQL & " es incorrecta"
        MsgBox miSQL, vbExclamation
        miSQL = ""
        Exit Function
    End If
    'Comprobamos los datos de los dtos
    'ALMUERZO
    If Text1(3).Text = "" Then Text1(3).Text = 0
    If Text1(3).Text = "0" Or Text1(3).Text = "0,00" Then
        Text1(4).Text = ""
        Text1(4).Text = ""
        
        Else
            'Comprobamos su valor
            If Not IsNumeric(Text1(3).Text) Then
                MsgBox "El descuento de empleados por el almuerzo debe de ser numérico.", vbExclamation
                Exit Function
            End If
            'Llegados a este punto tiene dto. Comprobaremos la hora
            If Text1(4).Text = "" Then
                MsgBox "Ponga una hora almuerzo.", vbExclamation
                Exit Function
            End If
            If Not IsDate(Text1(4).Text) Then
                MsgBox "Hora almuerzo incorrecta. Formato fecha incorrecto.", vbExclamation
                Exit Function
            End If
    End If
    
    'MERIENDA
    If Text1(5).Text = "" Then Text1(5).Text = 0
    If Text1(5).Text = "0" Or Text1(5).Text = "0,00" Then
        Text1(6).Text = ""
        Text1(6).Text = ""
        
        Else
            'Comprobamos su valor
            If Not IsNumeric(Text1(5).Text) Then
                MsgBox "El descuento de empleados por la merienda debe de ser numérico.", vbExclamation
                Exit Function
            End If
            'Llegados a este punto tiene dto. Comprobaremos la hora
            If Text1(6).Text = "" Then
                MsgBox "Ponga una hora merienda.", vbExclamation
                Exit Function
            End If
            If Not IsDate(Text1(6).Text) Then
                MsgBox "Hora merienda incorrecta. Formato fecha incorrecto.", vbExclamation
                Exit Function
            End If
    End If
    
    
    For i = 38 To 44
        txtAux(i).Text = Trim(txtAux(i).Text)
        J = i - 38
        valor = Abs(CheckF(J).Value)
        If txtAux(i).Text <> "" Then
            'Tiene puesto datos en el text
            If valor = 0 Then
                If Not IsNumeric(txtAux(i).Text) Then
                    MsgBox "Dias / Nomina   debe ser numérico.", vbExclamation
                    Exit Function
                End If
                If Val(txtAux(i).Text) > 1 Then
                    MsgBox "Valor maximo Dia/nomina es 1", vbExclamation
                    Exit Function
                End If
        
            End If
        Else
            If vEmpresa.laboral Then
                If valor = 0 Then
                    MsgBox "Campo Dia/nomina requerido", vbExclamation
                    Exit Function
                End If
            End If
        End If
    Next i
        
    
    

         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per la clua primaria ***
    cad = "(idhorario=" & Val(Text1(0).Text) & ")"
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
    vWhere = " WHERE idhorario=" & Data1.Recordset!IdHorario
    ' ************************************************
    
    'Borramos los modificarfichajes
    conn.Execute "DELETE FROM modificarfichajes" & vWhere
    conn.Execute "DELETE FROM subhorarios" & vWhere
    conn.Execute "DELETE FROM " & NombreTabla & vWhere
    
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


Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ***************** configurar els camps de buscar codis *****************
    Select Case Index
        Case 2
            PonerFormatoEntero Text1(Index)


        Case 3, 5
            PonerFormatoDecimal Text1(Index), 4
            i = 0
            If Index = 5 Then i = 1
            If Text1(Index).Text = "" Then
                Text2(i).Text = ""
            Else
                Text2(i).Text = DevuelveHora(CSng(Text1(Index).Text))
            End If
        Case 4, 6
            i = InStr(1, Text1(Index).Text, ".")
            If i > 0 Then Text1(Index).Text = Mid(Text1(Index).Text, 1, i - 1) & ":" & Mid(Text1(Index).Text, i + 1)
            If Right(Text1(Index), 1) = ":" Then Text1(Index).Text = Text1(Index).Text & "00"
    
            If IsDate(Text1(Index).Text) Then Text1(Index).Text = Format(Text1(Index).Text, "hh:mm")

'        Case 8 'NIF
'            Text1(Index).Text = UCase(Text1(Index).Text)
'            ValidarNIF Text1(Index).Text
'
'        Case 9, 10 'telèfons, fax i mòbils
'            PosarFormatTelefon Text1(Index)
'
'        Case 12, 13 'loginweb i passwweb
'            Text1(Index).Text = LCase(Text1(Index).Text)
'            If Index = 12 Then 'login
'                If Not ComprobarLoginEmp(Text1(Index).Text) Then PonerFoco Text1(Index)
'            End If
'
'        Case 6, 21 'poblacion
'            Nuevo = False
'            If Index = 6 Then
'                PonerDatosPoblacion Text1(Index), Text2(Index), Text1(Index + 1), , , Nuevo, Text1(9)
'            ElseIf Index = 21 Then
'                PonerDatosPoblacion Text1(Index), Text2(Index), Text1(Index + 1), , , Nuevo
'            End If
'            If Nuevo Then
'                Indice = Index
'                Set frmPob = New frmPoblacio
'                frmPob.DatosADevolverBusqueda = "0|1|2|3|4|5|"
'                frmPob.NuevoCodigo = Text1(Index).Text
'                Text1(Index).Text = ""
'                TerminaBloquear
'                frmPob.Show vbModal
'                Set frmPob = Nothing
'                If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'            End If
'
'        Case 3, 33 'empresa
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "empresas", "nomempre", "codempre", "N")
'                If Text2(Index).Text = "" Then
'                    cadMen = "No existe la Empresa: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmEmp = New frmEmpresas
'                        frmEmp.DatosADevolverBusqueda = "0|1|"
'                        frmEmp.NuevoCodigo = Text1(Index).Text
'                        Text1(Index).Text = ""
'                        TerminaBloquear
'                        frmEmp.Show vbModal
'                        Set frmEmp = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        Text1(Index).Text = ""
'                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'            If Index = 3 And Text1(4).Text <> "" Then Text1_LostFocus (4)
'
'        Case 4 'agencia
'            If PonerFormatoEntero(Text1(Index)) Then
'                cadMen = Text1(3).Text 'empresa
'                If (cadMen = "" Or Not IsNumeric(cadMen)) Then
'                    Text1(4).Text = ""
'                    Text2(4).Text = ""
'                    Exit Sub
'                End If
'
'                Text2(Index).Text = DevuelveDesdeBDNew(cPTours, "agencias", "desagenc", "codempre", cadMen, "N", , "codagenc", Text1(Index).Text, "N")
'                FormateaCampo Text1(Index)
'                If Text2(Index).Text = "" And Text1(Index) <> "" Then
'                    cadMen = "No existe la Agencia: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "para la Empresa: " & Text1(3).Text & "  " & Text2(3).Text & vbCrLf
'                    MsgBox cadMen, vbExclamation
''                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
''                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
''                        Set frmAge = New frmAgencias
''                        frmAge.DatosADevolverBusqueda = "0|1|"
''                        frmAge.NuevoCodigo = text1(Index).Text
''                        text1(Index).Text = ""
''                        TerminaBloquear
''                        frmAge.Show vbModal
''                        Set frmAge = Nothing
''                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
''                    Else
''                        text1(Index).Text = ""
''                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'        Case 15 'Cuenta Contable
'            If Text1(Index).Text = "" Then
'                Text2(Index).Text = ""
'                Exit Sub
'            End If
'            If Modo = 3 And ContieneCaracterBusqueda(Text1(Index).Text) Then Exit Sub     'Busquedas
'            Text2(Index).Text = PonerNombreCuenta(Text1(Index))
'
'        Case 16 'Cuenta Bancaria
'            If PonerFormatoEntero(Text1(Index)) Then
'                If Text1(Index).Text = "" Then Exit Sub
'                Text2(Index).Text = DevuelveDesdeBDNew(cPTours, "bancsofi", "nombanco", "codnacio", Combo1(2).ItemData(Combo1(2).ListIndex), "N", , "codbanco", Text1(Index).Text, "N")
'                If Text2(Index).Text = "" Then
'                    cadMen = "No existe el Banco: " & Text1(Index).Text & "  "
'                    cadMen = cadMen & "para el pais: " & Combo1(2).List(Combo1(2).ListIndex) & vbCrLf
'                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmBan = New frmBancsofi
'                        frmBan.DatosADevolverBusqueda = "0|1|"
'                        frmBan.NuevoCodigo = Text1(Index).Text
'                        frmBan.NuevoPais = Combo1(2).ItemData(Combo1(2).ListIndex)
'                        Text1(Index).Text = ""
'                        TerminaBloquear
'                        frmBan.Show vbModal
'                        Set frmBan = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        Text1(Index).Text = ""
'                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'        Case 24 'Tipo de Nómina
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiponomi", "desnomin")
'                If Text2(Index).Text = "" And Text1(Index) <> "" Then
'                    cadMen = "No existe el Tipo de Nómina: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmTiN = New frmTiponomi
'                        frmTiN.DatosADevolverBusqueda = "0|1|"
'                        frmTiN.NuevoCodigo = Text1(Index).Text
'                        Text1(Index).Text = ""
'                        TerminaBloquear
'                        frmTiN.Show vbModal
'                        Set frmTiN = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        Text1(Index).Text = ""
'                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
'        Case 30 'Tipo de Empleado
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiposemp", "desemple")
'                If Text2(Index).Text = "" Then
'                    cadMen = "No existe el Tipo de Empleado: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmTiE = New frmTiposemp
'                        frmTiE.DatosADevolverBusqueda = "0|1|"
'                        frmTiE.NuevoCodigo = Text1(Index).Text
'                        Text1(Index).Text = ""
'                        TerminaBloquear
'                        frmTiE.Show vbModal
'                        Set frmTiE = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        Text1(Index).Text = ""
'                    End If
'                    PonerFoco Text1(Index)
'                End If
'            Else
'                Text2(Index).Text = ""
'            End If
'
        Case 27, 28, 29 'dates
         '   If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
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
            KeyPress KeyAscii
        End If
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not Text1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
  '  imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    'imgFec_Click (Indice)
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
        .cadTabla2 = "horarios"
        .Informe2 = "rHorarios.rpt"
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa= '" & vEmpresa.NomEmpresa & "'|pOrden={horarios.idhorario}|"
        .NumeroParametros2 = 2
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = True
        
        .Show vbModal
    End With

End Sub




Private Function SePuedeEliminar() As Boolean
    
    SePuedeEliminar = True
End Function


Private Sub PonerSubHorarios()
Dim cad As String
Dim valor As Integer
Dim i As Integer
    cad = "Select * From SubHorarios Where IdHorario=" & Data1.Recordset.Fields(0)
    cad = cad & " ORDER BY DiaSemana"
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, , , adCmdText
    valor = 0
    While Not RS.EOF
        valor = RS.Fields!DiaSemana - 1
        
        i = (valor * 5) + 3
        txtAux(38 + valor).Text = ""
        If RS!Festivo = 0 Then
            CheckF(valor).Value = 0
            txtAux(i).Text = Format(DBLet(RS.Fields!HEntrada1), "hh:mm")
            txtAux(i + 1).Text = Format(DBLet(RS.Fields!HSalida1), "hh:mm")
            txtAux(i + 2).Text = Format(DBLet(RS.Fields!hentrada2), "hh:mm")
            txtAux(i + 3).Text = Format(DBLet(RS.Fields!HSalida2), "hh:mm")
            txtAux(i + 4).Text = DBLet(RS.Fields!HorasDia)
        
            'Dias nomina
            If Not IsNull(RS.Fields!DiaNomina) Then
                If RS.Fields!DiaNomina <> 0 Then
                    If RS.Fields!DiaNomina <> Int(RS.Fields!DiaNomina) Then
                        txtAux(38 + valor).Text = Format(RS.Fields!DiaNomina, "0.00")
                    Else
                        txtAux(38 + valor).Text = Int(RS.Fields!DiaNomina)
                    End If
                End If
            End If
        
        Else
            CheckF(valor).Value = 1
            txtAux(i).Text = ""
            txtAux(i + 1).Text = ""
            txtAux(i + 2).Text = ""
            txtAux(i + 3).Text = ""
            txtAux(i + 4).Text = ""
            

        End If
        valor = valor + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub

Private Sub txtaux_Change(Index As Integer)
    HanCambiadoSubHorarios = True
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim SoloSemanales As Boolean
Dim valor As Integer

If Modo > 2 Then
    Select Case Index
            Case 3 To 7
                'Lunes
                valor = 3
            Case 8 To 12
                'Martes
                valor = 8
            Case 13 To 17
                'Miercoles
                valor = 13
            Case 18 To 22
                'Jueves
                valor = 18
            Case 23 To 27
                'Viernes
                valor = 23
            Case 28 To 32
                'Sabado
                valor = 28
            Case 33 To 37
                'Domingo
                valor = 33
                
            Case 38 To 44
                'Si tiene valor
                txtAux(Index).Text = Trim(txtAux(Index).Text)
                If txtAux(Index).Text <> "" Then
                    If Not IsNumeric(txtAux(Index).Text) Then
                        MsgBox "Campo numérico", vbExclamation
                        txtAux(Index).Text = ""
                    End If
                   Exit Sub
                End If
            Case Else
                
                valor = 0
            End Select
                
                
            If txtAux(Index).Text <> "" Then
                'Cambiamos puntos por dos puntos
                If Index <> (valor + 4) Then
                    
                    i = InStr(1, txtAux(Index).Text, ".")
                    If i > 0 Then txtAux(Index).Text = Mid(txtAux(Index).Text, 1, i - 1) & ":" & Mid(txtAux(Index).Text, i + 1)
                    If Right(txtAux(Index), 1) = ":" Then txtAux(Index).Text = txtAux(Index).Text & "00"
            
                    If IsDate(txtAux(Index).Text) Then txtAux(Index).Text = Format(txtAux(Index).Text, "hh:mm")
                End If
                
            End If
            i = (Index - 3) Mod 5
            If Index = (valor + 4) Then
                If Not IsNumeric(txtAux(Index).Text) Then
                    txtAux(Index).Text = ""
                Else
                    txtAux(Index).Text = TransformaPuntosComas(txtAux(Index).Text)
                End If
            End If
            'SoloSemanales = True
            If valor > 0 Then
                'Si el que pierde el foco es el campo hora NO recalculo
                If ((Index - 2) Mod 5) <> 0 Then
                'If txtaux(Valor + 4).Text = "" Then SoloSemanales = False
                'CalculaHorasDia Valor, I, SoloSemanales
                    CalculaHorasDia valor, i, False
                End If
            End If
   
    End If
End Sub



Private Sub CalculaHorasDia(v As Integer, vInd As Integer, SemanalesSolo As Boolean)
Dim k As Integer
Dim T1 As Single
Dim T2 As Single
Dim v1(3) As Single


If Not SemanalesSolo Then
    For k = 0 To 3
        If txtAux(v + k).Text <> "" Then
            If IsDate(txtAux(v + k).Text) Then
                  v1(k) = DevuelveValorHora(txtAux(v + k).Text)
            Else
                v1(k) = 0
            End If
        Else
            v1(k) = 0
        End If
    Next k
    
    
    'Ya tenemos en cada tag
    If v1(0) = 0 And v1(1) = 0 Then
        T1 = 0
    Else
        T1 = v1(1) - v1(0)
    End If
    If v1(3) = 0 And v1(2) = 0 Then
        T2 = 0
    Else
        T2 = v1(3) - v1(2)
    End If
    'Las horas totales del dia son la suma de ambas
    txtAux(v + 4).Text = T1 + T2

End If
    
'Recalcularemos las horas totales semanales

T2 = 0
For i = 0 To 6
    If txtAux(7 + (5 * i)).Text <> "" Then
        T1 = CSng(txtAux(7 + (5 * i)).Text)
        T2 = T2 + T1
    End If
Next i
If T2 > 0 Then
    Text1(2).Text = T2
    Else
    Text1(2).Text = ""
End If
End Sub


Private Function FechaOk(Texto As String) As Boolean
FechaOk = True
If Texto <> "" Then FechaOk = IsDate(Texto)
End Function



Private Function SubHorariosAbd(idHora As Integer)
Dim J As Integer
Dim k As Integer
    miSQL = "Delete  from SubHorarios where idHorario=" & idHora
    conn.Execute miSQL
    
    'INSERT INTO subhorarios (IdHorario, DiaSemana, Festivo, HEntrada1,
    'HSalida1, HEntrada2, HSalida2, N_Tikadas, HorasDia, DiaNomina) VALUES (
    
    For i = 0 To 6
        miSQL = idHora & "," & i + 1
        If CheckF(i).Value = 1 Then
            'Es festivo
            miSQL = miSQL & ",1,NULL,NULL,NULL,NULL,0,0,0)"
        
        Else
            miSQL = miSQL & ",0"
            J = (i * 5) + 3
            'Introducimos los subhorarios para cada dia
            For k = 0 To 3
                miSQL = miSQL & DevuelveFecha(txtAux(J + k).Text)
            Next k
            
            
            k = 0
            For J = ((i * 5) + 3) To ((i * 5) + 3 + 3)
                If txtAux(J).Text <> "" Then k = k + 1
            Next J
            miSQL = miSQL & "," & k & ","
            
            'Horas al dia
            J = ((i * 5) + 3 + 4)
            miSQL = miSQL & TransformaComasPuntos(txtAux(J).Text) & ","
            
            
            
            
            J = i + 38
            If txtAux(J).Text = "" Then txtAux(J).Text = 0
            miSQL = miSQL & TransformaComasPuntos(txtAux(J).Text) & ")"
            
        End If
        miSQL = "INSERT INTO subhorarios (IdHorario, DiaSemana, Festivo, HEntrada1,HSalida1, HEntrada2, HSalida2, N_Tikadas, HorasDia, DiaNomina) VALUES (" & miSQL
        Debug.Print miSQL
        conn.Execute miSQL
    
    Next i
    

End Function



Private Function DevuelveFecha(Texto As String) As String
If Texto = "" Then
    DevuelveFecha = ",Null"
    Else
        DevuelveFecha = ",'" & Texto & ":00'"
End If
End Function







Private Sub InsertarLinea()
'Inserta registro en las tablas de Lineas: provbanc, provdpto
Dim nomFrame As String
Dim B As Boolean
On Error Resume Next

'    Select Case NumTabMto
'        Case 0: nomFrame = "FrameAux0" 'viajeros
'        Case 1: nomFrame = "FrameAux1" 'Departamentos
'        Case 2: nomFrame = "FrameAux2" 'Productos
'    End Select
    
    nomFrame = "FrameAux0"
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
'            If NumTabMto = 0 Then
'                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtAux(2).Text)
'                End If
'            End If
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            CargaGrid NumTabMto, True
            If B Then BotonAnyadirLinea 0
            SituarTab (NumTabMto)
            PonerFoco txtAux2(1)
            
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registro en las tablas de Lineas: provbanc, provdpto
Dim nomFrame As String
Dim v As Date
On Error GoTo EModificarLin

    'Select Case SSTab1.Tab
    '    Case 0: nomFrame = "FrameAux0" 'cuentas Bancarias
    '
    'End Select
    
    nomFrame = "FrameAux0"
    
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
'            If NumTabMto = 0 Then
'                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtAux(2).Text)
'                End If
'            End If
            v = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
            SituarTab (NumTabMto)
'            SSTab1.Tab = 1
'            SSTab2.Tab = NumTabMto
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'per a que es quede en modificar
'            PonerModo 4
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " = '" & Format(v, "hh:mm:ss'"))

            LLamaLineas NumTabMto, 0
        End If
    End If
EModificarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Linea", Err.Description
End Sub



Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim B As Boolean
On Error GoTo EDatosOKLlin

    DatosOkLlin = False
        
    B = CompForm2(Me, 2, nomFrame) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
   Select Case NumTabMto
   Case 0
        If CDate(txtAux2(1).Text) > CDate(txtAux2(2).Text) Then
            MsgBox "Incio mayor que fin", vbExclamation
            Exit Function
        End If
   
        If CDate(txtAux2(3).Text) > CDate(txtAux2(2).Text) Then
            MsgBox "Hora final mayor que la modificada", vbExclamation
            Exit Function
        End If
   
        'Nos vamos a la BD
        '---------------
        If ModoLineas = 1 Then
            miSQL = "Select count(*) from modificarfichajes where idhorario = " & Data1.Recordset!IdHorario
            miSQL = miSQL & " AND ((inicio <='" & Format(txtAux2(1).Text, "hh:mm:ss") & "' AND fin >= '"
            miSQL = miSQL & Format(txtAux2(1).Text, "hh:mm:ss") & "') OR "
            miSQL = miSQL & " ( inicio <='" & Format(txtAux2(2).Text, "hh:mm:ss") & "' AND fin >= '"
            miSQL = miSQL & Format(txtAux2(2).Text, "hh:mm:ss") & "'))  "
            Set miRs = New ADODB.Recordset
            miRs.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            i = 0
            If Not miRs.EOF Then
                i = DBLet(miRs.Fields(0), "N")
            End If
            miRs.Close
            Set miRs = Nothing
            If i > 0 Then
                MsgBox "El intervalo esta comprendido entre otros", vbExclamation
                Exit Function
            End If
        End If
   End Select
         
    DatosOkLlin = B
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim B As Boolean
Dim tots As String
On Error GoTo ECarga

      
    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 290
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If
    
    'DataGridAux(Index).Enabled = b
    PrimeraVez = False
    
    Select Case Index
        Case 0 'Rectificacion
            'si es visible|control|tipo campo|nombre campo|ancho control|formato campo|
'            tots = "N||||0|;S|txtAux(1)|T|NºLinea|800|;" 'numexped,numlinea
            tots = "N||||0|;" 'idhorario
            tots = tots & "S|txtAux2(1)|T|Inicio|1000|;S|txtAux2(2)|T|Fin|1000|;" 'nombre, apellido
            tots = tots & "S|txtAux2(3)|T|Modificada|1000|;"
            
            arregla tots, DataGridAux(Index), Me
'            DataGridAux(Index).Columns(4).Alignment = dbgCenter
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub



Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    'If ModificaLineas = 2 Then Exit Sub
    ModoLineas = 1 'Ponemos Modo Añadir Linea
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cabecera
        cmdAceptar_Click
        'No se ha insertado la cabecera
        If ModoLineas = 0 Then Exit Sub
'        'si la cabecera no esta insertada salir
'        If DevuelveDesdeBD("codprove", "proveedo", "codprove", text1(0).Text, "N") = "" Then
'            Exit Sub
'        End If
    End If
    
    NumTabMto = Index
    PonerModo 5
'    If b Then BloquearText1 Me, 4 'Si viene de Insertar Cabecera no bloquear los Text1
    
'    BloquearTxt Text1(0), True
    
    'Obtener el numero de linea ha insertar
'    Select Case Index
'        Case 0: vTabla = "expinvia"
'        Case 1: vTabla = "provdpto"
'        Case 2: vTabla = "provprod"
'    End Select
'    vWhere = " codprove=" & text1(0).Text & " AND codempre=" & codEmpre
    vWhere = ObtenerWhereCab(False)
    'NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

    'Situamos el grid al final
    AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
    anc = DataGridAux(Index).Top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
    End If
    
    LLamaLineas Index, ModoLineas, anc
    
    Select Case Index
'        Case 0 'cuentas
'            txtAux(0).Text = Text1(0).Text 'codprove
'            txtAux(1).Text = codEmpre 'codempre
'            txtAux(2).Text = NumF 'numlinea
'            For i = 3 To 7
'                txtAux(i).Text = ""
'            Next i
'            txtAux2(0).Text = ""
'            CargaCombo Index 'per a carregar els valors per defecte dels combo
'
'             'valor por defecto del cmbAux(0). per defecte seleccione España (724)
'            For i = 0 To cmbAux(0).ListCount - 1
'                If cmbAux(0).ItemData(i) = 724 Then
'                    cmbAux(0).ListIndex = i
'                    Exit For
'                End If
'            Next i
'
'            BloquearTxt txtAux(11), False
'            BloquearTxt txtAux(12), False
''            PonerFoco txtAux(3)
'            Me.cmbAux(0).SetFocus
'
'        Case 1 'departamentos
'            txtAux(13).Text = Text1(0).Text 'codprove
'            txtAux(14).Text = codEmpre 'codempre
'            txtAux(15).Text = NumF 'numlinea
'            For i = 16 To 17
'                txtAux(i).Text = ""
'            Next i
'            CargaCombo Index 'per a carregar els valors per defecte dels combo
'
'            For i = 21 To 24
'                BloquearTxt txtAux(i), False
'                txtAux(i).Text = ""
'            Next i
'            txtAux2(22).Text = ""
'            PonerFoco txtAux(16)
            
        Case 0 'Viajeros
            
            For i = 1 To 3
                txtAux2(i).Text = ""
            Next i
            txtAux2(0).Text = Data1.Recordset!IdHorario
            BloquearTxt txtAux2(1), False
            PonerFoco txtAux2(1)
    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    ModoLineas = 2 'Modificar llínia
    
    If Modo = 4 Then 'Modificar Cabecera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
    
    NumTabMto = Index
    PonerModo 5
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub

'    Me.lblIndicador.Caption = "MODIFICAR LINEA"
    
    If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
        i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
        DataGridAux(Index).Scroll 0, i
        DataGridAux(Index).Refresh
    End If
      
    anc = DataGridAux(Index).Top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
    End If

    Select Case Index
'        Case 0 'cuentas bancarias
'            For j = 0 To 2
'                txtAux(j).Text = DataGridAux(Index).Columns(j).Text
'            Next j
'
'            SelComboBool AdoAux(Index).Recordset!codNacio, cmbAux(0)
'
'            For j = 5 To 7
'                txtAux(j).Text = DataGridAux(Index).Columns(j).Text
'            Next j
'
'            SelComboBool AdoAux(Index).Recordset!ctactiva, cmbAux(9)
'            SelComboBool AdoAux(Index).Recordset!ctaprpal, cmbAux(10)
'            txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'            txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
'            BloquearTxt txtAux(11), False
'            BloquearTxt txtAux(12), False
            
'        Case 1 'departamentos
'            For j = 13 To 17
'                txtAux(j).Text = DataGridAux(Index).Columns(j - 13).Text
'            Next j
'
'            SelComboBool AdoAux(Index).Recordset!facturac, cmbAux(18)
'            SelComboBool AdoAux(Index).Recordset!document, cmbAux(19)
'            SelComboBool AdoAux(Index).Recordset!princpal, cmbAux(20)
'
'            For i = 21 To 24
'                BloquearTxt txtAux(i), False
'            Next i
            
        Case 0 'viajeros
            For J = 0 To 3
                txtAux2(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
          
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    Select Case Index
'        Case 0 'cuentas bancarias
'            PonerFoco txtAux(3)
'        Case 1 'departamentos
'            PonerFoco txtAux(16)
        Case 0 'Viajeros
            PonerFoco txtAux2(3)
    End Select
End Sub



Private Sub SituarTab(numTab As Integer)
On Error Resume Next
'    If numTab = 0 Or numTab = 1 Then
'        SSTab1.Tab = 0
'        SSTab2.Tab = NumTabMto
'    ElseIf numTab = 2 Then
        SSTab1.Tab = 2
'        SSTab3.Tab = 0
'    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim JJ As Integer
Dim B As Boolean

    On Error GoTo ELLamaLin

    DeseleccionaGrid DataGridAux(Index)
    'PonerModo xModo + 1
    
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    Select Case Index
        Case 0 'Viajeros
            For JJ = 1 To 3
                txtAux2(JJ).Top = alto
                txtAux2(JJ).Visible = B
            Next JJ
           
           
    End Select
    
    If xModo = 2 Then BloquearTxt txtAux2(1), True
    
ELLamaLin:
    Err.Clear
End Sub



Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    Select Case Index
'        Case 0 'CUENTAS BANCARIAS
'        '    If Data1.Recordset.EOF Then enlaza = False
'            tabla = "provbanc"
'            'SQL = "SELECT * FROM " & tabla
'            SQL = "SELECT codprove,codempre,numlinea,provbanc.codnacio,naciones.ibanpais,provbanc.codbanco,codsucur,digcontr,ctabanco,ibandctl,ctactiva,If(ctactiva=1,""Si"",""No""),ctaprpal,If(ctaprpal=1,""Si"",""No""),direccio,observac,nombanco "
'            SQL = SQL & " FROM " & tabla & " INNER JOIN bancsofi ON " & tabla & ".codnacio=bancsofi.codnacio AND " & tabla & ".codbanco=bancsofi.codbanco INNER JOIN naciones ON " & tabla & ".codnacio=naciones.codnacio "
'            If enlaza Then
'                SQL = SQL & ObtenerWhereCab(True)
'            Else
'                SQL = SQL & " WHERE codprove = -1"
'            End If
'            SQL = SQL & " ORDER BY " & tabla & ".numlinea "
            
'        Case 1 'DEPARTAMENTOS
'        '    If Data1.Recordset.EOF Then enlaza = False
'            SQL = "SELECT provdpto.codprove,provdpto.codempre,provdpto.numlinea,provdpto.nomdepto,provdpto.contacto,provdpto.direccio,provdpto.codpobla,poblacio.despobla,provdpto.codposta,provdpto.facturac,If(provdpto.facturac=1,""Si"",""No""),provdpto.document,If(provdpto.document=1,""Si"",""No""),provdpto.princpal,If(provdpto.princpal=1,""Si"",""No""),provdpto.observac  FROM provdpto, poblacio"
'            If enlaza Then
'                SQL = SQL & " WHERE provdpto.codprove=" & Text1(0).Text & " AND provdpto.codempre= " & codEmpre
'            Else
'                SQL = SQL & " WHERE provdpto.codprove = -1"
'            End If
'            SQL = SQL & " AND provdpto.codpobla = poblacio.codpobla"
'            SQL = SQL & " ORDER BY provdpto.numlinea"

        Case 0 'Viajeros del expediente
            Sql = "Select idhorario,inicio,fin,modificada FROM modificarfichajes"
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE idhorario = -1"
            End If
            Sql = Sql & " ORDER BY inicio "
    End Select
    MontaSQLCarga = Sql
End Function


Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    vWhere = ""
    If conW Then vWhere = " WHERE "
'    vWhere = vWhere & " codempre=" & vEmpresa.codEmpre & " AND numexped=" & Val(Text1(0).Text)
    vWhere = vWhere & " idhorario=" & Val(Text1(0).Text)
    ObtenerWhereCab = vWhere
End Function

Private Sub txtAux2_GotFocus(Index As Integer)
    ConseguirFoco txtAux2(Index), 3
End Sub

Private Sub txtAux2_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtAux2_LostFocus(Index As Integer)
    i = 0
    txtAux2(Index).Text = Trim(txtAux2(Index).Text)
    If txtAux2(Index).Text <> "" Then
    
    
        i = InStr(1, txtAux2(Index), ".")
        If i > 0 Then
            miSQL = Mid(txtAux2(Index).Text, i + 1)
            If miSQL = "" Then miSQL = "00"
            txtAux2(Index).Text = Mid(txtAux2(Index).Text, 1, i - 1) & ":" & miSQL & ":00"
        End If
        If IsDate(txtAux2(Index).Text) Then
            txtAux2(Index) = Format(txtAux2(Index), "hh:mm:ss")
            i = 0
        Else
            
            MsgBox "Campo hora incorrecto:" & txtAux2(Index).Text, vbExclamation
            txtAux2(Index).Text = ""
            PonerFoco txtAux2(Index)
            i = 1
        End If
        
    
    End If
    If i = 0 Then
        If Index = 3 Then PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Sub BloquearRestoCampos(Si As Boolean)
Dim J As Integer
    On Error Resume Next
    Frame3.Enabled = Si
    Text2(0).Enabled = Si
    Text2(1).Enabled = Si
    Err.Clear
        
End Sub




Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim VEliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    VEliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'cltebanc
            Sql = "¿Seguro que desea eliminar la rectificación?" & vbCrLf
            For i = 1 To 3
            
                Sql = Sql & DataGridAux(0).Columns(i).Caption & " : "
                Sql = Sql & Space(30 - Len(DataGridAux(0).Columns(i).Caption)) & DataGridAux(0).Columns(i).Text & vbCrLf
            Next i
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                VEliminar = True
                Sql = "DELETE FROM modificarfichajes"
                Sql = Sql & vWhere & " AND inicio= '" & Format(AdoAux(Index).Recordset!Inicio, "hh:mm:ss") & "'"
            End If
            
    End Select

    If VEliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If

        ' ***************************************
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto)
        ' ************************
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


