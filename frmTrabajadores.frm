VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrabajadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empleados"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "frmTrabajadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin MSAdodcLib.Adodc AdodcImg 
      Height          =   495
      Left            =   5520
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   240
      TabIndex        =   34
      Top             =   1320
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11668
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmTrabajadores.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4(9)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4(12)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4(13)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label4(23)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4(21)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4(22)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "imgFec(11)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "imgEMAIL"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "imgFec(28)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label4(20)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(17)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(5)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(7)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(8)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Combo1(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Combo1(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "FrameImagen"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(9)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Option1(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Option1(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(11)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(12)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(18)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Combo1(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(28)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "FrameTapaCtaBancaria"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "FrameHuella"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).ControlCount=   40
      TabCaption(1)   =   "Horario"
      TabPicture(1)   =   "frmTrabajadores.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).Control(1)=   "Text1(26)"
      Tab(1).Control(2)=   "ListView1"
      Tab(1).Control(3)=   "ListView2"
      Tab(1).Control(4)=   "imgVacas"
      Tab(1).Control(5)=   "imgCalendariP(0)"
      Tab(1).Control(6)=   "Label2(2)"
      Tab(1).Control(7)=   "ImgModifHora"
      Tab(1).Control(8)=   "imgZoom(0)"
      Tab(1).Control(9)=   "Label2(1)"
      Tab(1).Control(10)=   "Label2(0)"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Laboral"
      TabPicture(2)   =   "frmTrabajadores.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Framelaboral"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameHuella 
         Height          =   3675
         Left            =   8760
         TabIndex        =   91
         Top             =   2280
         Width           =   2415
         Begin VB.CheckBox Check2 
            Caption         =   "Sin huella"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Tag             =   "Sin huella|N|S|||Trabajadores|NoCapturarHuella|||"
            Top             =   150
            Width           =   1095
         End
         Begin VB.CommandButton cmdHuella 
            Caption         =   "Capturar"
            Height          =   435
            Left            =   480
            TabIndex        =   92
            ToolTipText     =   "Capturar huella"
            Top             =   2520
            Width           =   1455
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   480
            TabIndex        =   93
            ToolTipText     =   "Calidad imagen"
            Top             =   3240
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   100
            SelStart        =   70
            TickStyle       =   3
            Value           =   70
         End
         Begin VB.Image imgHuella 
            Height          =   1815
            Left            =   360
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   135
            Index           =   1
            Left            =   360
            TabIndex        =   97
            Top             =   3240
            Width           =   60
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   135
            Index           =   0
            Left            =   1920
            TabIndex        =   96
            Top             =   3240
            Width           =   240
         End
         Begin VB.Label lblInfCodigo 
            Alignment       =   2  'Center
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   855
            Left            =   360
            TabIndex        =   95
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Frame FrameTapaCtaBancaria 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   5520
         TabIndex        =   85
         Top             =   960
         Width           =   5655
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   29
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "IBAN|T|S|||trabajadores|iban|||"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   18
            Tag             =   "Tarjeta|T|S|||trabajadores|entidad|0000||"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   2880
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "Tarjeta|T|S|||trabajadores|oficina|0000||"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   15
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "Tarjeta|T|S|||trabajadores|controlcta|00||"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   16
            Left            =   4200
            MaxLength       =   10
            TabIndex        =   21
            Tag             =   "Tarjeta|T|S|||trabajadores|cuenta|0000000000||"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "IBAN"
            Height          =   195
            Index           =   29
            Left            =   1440
            TabIndex        =   100
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label4 
            Caption         =   "Datos banco"
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
            Index           =   14
            Left            =   120
            TabIndex        =   90
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Entidad"
            Height          =   195
            Index           =   15
            Left            =   2160
            TabIndex        =   89
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label4 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   16
            Left            =   2880
            TabIndex        =   88
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "CC"
            Height          =   255
            Index           =   17
            Left            =   3720
            TabIndex        =   87
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   18
            Left            =   4200
            TabIndex        =   86
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "F.Alta|F|S|||trabajadores|fecbaja|||"
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -66840
         TabIndex        =   78
         Text            =   "Text2"
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   -67920
         MaxLength       =   4
         TabIndex        =   77
         Tag             =   "Calendario|N|N|||trabajadores|idCal|||"
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Framelaboral 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   54
         Top             =   360
         Width           =   11175
         Begin VB.Frame FrameTapaBolsaHoras 
            Height          =   1095
            Left            =   120
            TabIndex        =   102
            Top             =   4320
            Width           =   5775
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   240
            MaxLength       =   5
            TabIndex        =   60
            Tag             =   "IRPF cargo empresa|N|S|||trabajadores|IRPFempresa|0.00||"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Frame FrameContadoresHoras 
            Caption         =   "Contador horas"
            Height          =   3855
            Left            =   4560
            TabIndex        =   98
            Top             =   360
            Width           =   5895
            Begin MSComctlLib.ListView ListView3 
               Height          =   3135
               Left            =   240
               TabIndex        =   99
               Top             =   360
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   5530
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Tipo"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Cooperativa"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Horas"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   56
            Tag             =   "DNI|T|S|||trabajadores|nummat|||"
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   27
            Left            =   240
            MaxLength       =   50
            TabIndex        =   55
            Tag             =   "F.Alta|F|S|||trabajadores|fecalta|||"
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   63
            Tag             =   "Tipo trabajo|N|S|||trabajadores|tipocontrato|||"
            Text            =   "Combo1"
            Top             =   3480
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   25
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   66
            Tag             =   "bolsa horas|N|S|||trabajadores|bolsabruto|||"
            Top             =   4800
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   2880
            MaxLength       =   4
            TabIndex        =   65
            Tag             =   "bolsa neto|N|S|||trabajadores|bolsaneto|||"
            Top             =   4800
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   64
            Tag             =   "bolsa horas|N|S|||trabajadores|bolsahoras|||"
            Top             =   4800
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   22
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   61
            Tag             =   "Cod. Asesoria|N|S|||trabajadores|idAsesoria|||"
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Pago bancario"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   62
            Tag             =   "P|N|S|||trabajadores|pagobancario|||"
            Top             =   2400
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   2880
            MaxLength       =   5
            TabIndex        =   59
            Tag             =   "IRPF|N|N|||trabajadores|porcIRPF|0.00||"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   58
            Tag             =   "Seguridad social|N|N|||trabajadores|porcSS|0.00||"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   240
            MaxLength       =   5
            TabIndex        =   57
            Tag             =   "Tarjeta|N|N|||trabajadores|porcantiguedad|0.00||"
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "% IRPF empresa"
            Height          =   255
            Index           =   30
            Left            =   240
            TabIndex        =   101
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Matricula"
            Height          =   255
            Index           =   11
            Left            =   1800
            TabIndex        =   83
            Top             =   240
            Width           =   1455
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   27
            Left            =   840
            Picture         =   "frmTrabajadores.frx":0060
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "F. Alta"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   79
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo trabajo"
            Height          =   255
            Index           =   28
            Left            =   240
            TabIndex        =   75
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "BRUTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   240
            Index           =   3
            Left            =   4680
            TabIndex        =   74
            Top             =   4440
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "NETO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   2
            Left            =   3120
            TabIndex        =   73
            Top             =   4440
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "HORAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   1
            Left            =   1440
            TabIndex        =   72
            Top             =   4440
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "BOLSA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   71
            Top             =   4800
            Width           =   750
         End
         Begin VB.Label Label4 
            Caption         =   "ID asesoria"
            Height          =   255
            Index           =   27
            Left            =   1680
            TabIndex        =   70
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "% IRPF"
            Height          =   255
            Index           =   26
            Left            =   2880
            TabIndex        =   69
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "% SS"
            Height          =   255
            Index           =   25
            Left            =   1560
            TabIndex        =   68
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "% Antiguedad"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   67
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   8040
         TabIndex        =   16
         Tag             =   "s|N|N|||trabajadores|control|||"
         Text            =   "Combo1"
         Top             =   660
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   18
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Tarjeta|T|S|||trabajadores|numtarjeta|||"
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "SS|T|S|||trabajadores|numSS|||"
         Top             =   3060
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   480
         MaxLength       =   50
         TabIndex        =   13
         Tag             =   "F.Alta|F|S|||trabajadores|antiguedad|||"
         Top             =   5880
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   " Mujer"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   23
         Top             =   5880
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hombre"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   22
         Top             =   5880
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   480
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "DNI|T|S|||trabajadores|numDNI|||"
         Top             =   3060
         Width           =   1935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   45
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inicio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fin"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Horario"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Frame FrameImagen 
         Height          =   3735
         Left            =   5400
         TabIndex        =   44
         Top             =   2280
         Width           =   3135
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   240
            TabIndex        =   80
            Top             =   240
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Ver"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar "
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar "
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "WebCam "
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Guardar a disco"
               EndProperty
            EndProperty
         End
         Begin VB.Image ImageTra 
            Height          =   2175
            Left            =   240
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   2535
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   12
         Tag             =   "Categoria|N|N|||trabajadores|idcategoria|||"
         Text            =   "Combo1"
         Top             =   5100
         Width           =   4695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Tag             =   "Seccion|N|N|||trabajadores|seccion|||"
         Text            =   "Combo1"
         Top             =   4380
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   480
         MaxLength       =   100
         TabIndex        =   10
         Tag             =   "Nombre|T|S|||trabajadores|email|||"
         Top             =   3660
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "D|T|S|||trabajadores|movtrabajador|||"
         Top             =   2460
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   480
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "D|T|S|||trabajadores|teltrabajador|||"
         Top             =   2460
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "D|T|S|||trabajadores|provtrabajador|||"
         Top             =   1860
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   480
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "D|T|S|||trabajadores|codpostrabajador|||"
         Top             =   1860
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   480
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "D|T|S|||trabajadores|pobtrabajador|||"
         Top             =   1260
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   480
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "Direccion|T|S|||trabajadores|domtrabajador|||"
         Top             =   660
         Width           =   4695
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3975
         Left            =   -67920
         TabIndex        =   81
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inicio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fin"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   49
         TabStop         =   0   'False
         Tag             =   "Tarjeta|N|S|||trabajadores|sexo|||"
         Top             =   690
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "F. Baja"
         Height          =   255
         Index           =   20
         Left            =   1920
         TabIndex        =   84
         Top             =   5640
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   28
         Left            =   2640
         Picture         =   "frmTrabajadores.frx":00EB
         ToolTipText     =   "Buscar fecha"
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image imgVacas 
         Height          =   240
         Left            =   -66960
         Picture         =   "frmTrabajadores.frx":0176
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgEMAIL 
         Height          =   255
         Left            =   1080
         ToolTipText     =   "Enviar mail"
         Top             =   3400
         Width           =   255
      End
      Begin VB.Image imgCalendariP 
         Height          =   240
         Index           =   0
         Left            =   -67080
         MouseIcon       =   "frmTrabajadores.frx":0B78
         MousePointer    =   4  'Icon
         ToolTipText     =   "Asignar calendario"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Vacaciones"
         Height          =   195
         Index           =   2
         Left            =   -67920
         TabIndex        =   82
         Top             =   1320
         Width           =   975
      End
      Begin VB.Image ImgModifHora 
         Height          =   240
         Left            =   -72960
         Picture         =   "frmTrabajadores.frx":0CCA
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   -73320
         MouseIcon       =   "frmTrabajadores.frx":16CC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Ver horarios"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   1560
         Picture         =   "frmTrabajadores.frx":181E
         ToolTipText     =   "Buscar fecha"
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Calendario"
         Height          =   195
         Index           =   1
         Left            =   -67920
         TabIndex        =   76
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Control"
         Height          =   255
         Index           =   22
         Left            =   8040
         TabIndex        =   53
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Tarjeta"
         Height          =   255
         Index           =   21
         Left            =   6840
         TabIndex        =   52
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Datos Reloj"
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
         Index           =   23
         Left            =   5640
         TabIndex        =   51
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Horarios asignados"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   50
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Nº S.S."
         Height          =   255
         Index           =   13
         Left            =   2640
         TabIndex        =   48
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "F. Antiguedad"
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   47
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "NIF"
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   46
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Sección"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   43
         Top             =   4140
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Categoría"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   42
         Top             =   4860
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "e-mail"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   41
         Top             =   3420
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Movil"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   40
         Top             =   2220
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   39
         Top             =   2220
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   38
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "C.P."
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   37
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Población"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   36
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Dirección"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   35
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Width           =   11295
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Tag             =   "Código de Empleado|N|N|1|99999|trabajadores|idTrabajador|0000|S|"
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||trabajadores|nomtrabajador|||"
         Top             =   210
         Width           =   6015
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre "
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cód."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   8040
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
         TabIndex        =   27
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10380
      TabIndex        =   25
      Top             =   8160
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   24
      Top             =   8160
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3240
      Top             =   8160
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
      TabIndex        =   31
      Top             =   8160
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   32
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
         TabIndex        =   33
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
Attribute VB_Name = "frmTrabajadores"
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

Private Const ExeWebCam = "aricam.exe"


Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)

' ****** Definir variables per a cridar a atres formularis *********
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1
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

Private WithEvents frmS As frmCalendario
Attribute frmS.VB_VarHelpID = -1

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
Dim indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos
Dim cad_meua As String
Dim CadB As String


Dim ErrorAcecionedoOCX As Byte


Private Sub cmdAceptar_Click()
Dim CambioCalendario As Boolean

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    'Si es nuevo entonces le copio los horarios
                    CopiandoHorarios
                    If vEmpresa.imgtrabaj Then TratarImagen
                    'situarnos en el registro que acabamos de insertar
                    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE idTrabajador=" & Text1(0).Text & Ordenacion
                    PosicionarData
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                CambioCalendario = False
                If Val(Data1.Recordset!idCal) <> Val(Text1(26).Text) Then
                    
                    CadenaDesdeOtroForm = ""
                    cad_meua = Text1(1).Text & "|"
                    
                    FrmVarios.Opcion = 2
                    FrmVarios.Parametros = cad_meua
                    FrmVarios.Show vbModal
                    If CadenaDesdeOtroForm = "" Then Exit Sub
                    cad_meua = CadenaDesdeOtroForm
                    CambioCalendario = True
                Else
                    cad_meua = ""
                End If
                If ModificaDesdeFormulario(Me) Then
                    If vEmpresa.imgtrabaj Then TratarImagen
                    If CambioCalendario Then HacerCambioCalendario
                    
                    'De momento alzira
                    ActualizarHuellaEnBSGestorHuella
                    
                    
                    
                    
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



Private Sub ActualizarHuellaEnBSGestorHuella()
Dim usu As UsuarioHuella

    'Alzira
    If vEmpresa.QueEmpresa <> 2 Then Exit Sub
    
    If Me.Check2.Value = 1 And Val(Me.Data1.Recordset!NoCapturarHuella) = 0 Then
        'Antes capturaba huella y ahora NO
        Set usu = New UsuarioHuella
        If usu.Leer(Text1(18).Text) Then
             usu.FIR = ""
             usu.Guardar
        Else
           MsgBox "No se ha encontrado usuario en tarminales KRETA", vbExclamation
        End If
    
    
        'usu.CapturaHuella Byt
       
        'If Dir(vEmpresa.DirHuellas & "\" & usu.CodUsuario & ".jpg") <> "" Then
        '    imgHuella.Picture = LoadPicture(vEmpresa.DirHuellas & "\" & usu.CodUsuario & ".jpg")
        '    imgHuella.Visible = True
        'Else
        '    imgHuella.Visible = False
        'End If
        
        
        Check2.Value = 0
        DoEvents
        Set usu = Nothing
    End If
    
End Sub

Private Sub HacerCambioCalendario()
Dim F As Date
Dim FESTIVOS As String
Dim i As Integer
    On Error GoTo EHacerCambioCalendario
    
    F = CDate(cad_meua)
    
    
    cad_meua = "Select * from calendariot where idtrabajador =" & Text1(0).Text
    cad_meua = cad_meua & " AND fecha >='" & Format(F, FormatoFecha) & "' AND Fecha <='"
    cad_meua = cad_meua & Format(vEmpresa.FechaFin, FormatoFecha) & "'"
    cad_meua = cad_meua & " AND tipodia=2"   'DOS: Festivos del trabajador
    FESTIVOS = ""
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad_meua, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        FESTIVOS = FESTIVOS & Format(miRsAux!Fecha, FormatoFecha) & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Borramos los horarios a partir de la fecha
    cad_meua = "delete from calendariot where idtrabajador=" & Text1(0).Text
    cad_meua = cad_meua & " AND fecha >='" & Format(F, FormatoFecha) & "' AND Fecha <= '"
    cad_meua = cad_meua & Format(vEmpresa.FechaFin, FormatoFecha) & "'"
    conn.Execute cad_meua
    
    'Por si acaso
    conn.Execute "commit"
    DoEvents
    
    'Creamos los nuevos horarios
    cad_meua = "INSERT INTO calendariot(idtrabajador,fecha,idhorario,tipodia) SELECT " & Text1(0).Text
    cad_meua = cad_meua & " ,fecha,idhorario,0 from calendariol where idcal =" & Text1(26).Text
    cad_meua = cad_meua & " AND fecha >='" & Format(F, FormatoFecha) & "' AND Fecha <= '"
    cad_meua = cad_meua & Format(vEmpresa.FechaFin, FormatoFecha) & "'"
    conn.Execute cad_meua
    
    
    'Updateo los dias festivos del calendario
    cad_meua = "SELECT * FROM calendariof where idcal=" & Val(Text1(26).Text)
    cad_meua = cad_meua & " AND fecha >='" & Format(F, FormatoFecha) & "' AND Fecha <='"
    cad_meua = cad_meua & Format(vEmpresa.FechaFin, FormatoFecha) & "'"
    miRsAux.Open cad_meua, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    While Not miRsAux.EOF
        cad_meua = "UPDATE calendarioT set tipodia=1 where fecha='" & Format(miRsAux!Fecha, FormatoFecha) & "' AND idtrabajador=" & Val(Text1(0).Text)
        conn.Execute cad_meua
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    'Updateo los festivos
    While FESTIVOS <> ""
        cad_meua = RecuperaValor(FESTIVOS, 1)
        cad_meua = "UPDATE calendarioT set tipodia=2 where fecha='" & cad_meua & "' AND idtrabajador=" & Text1(0).Text
        conn.Execute cad_meua
        i = InStr(1, FESTIVOS, "|")
        FESTIVOS = Mid(FESTIVOS, i + 1)
    Wend
    
    Exit Sub
EHacerCambioCalendario:
    MuestraError Err.Number, "Cambio calendario"
    Set miRsAux = Nothing
End Sub



Private Sub CopiandoHorarios()
Dim C As String
Dim i As Integer


    On Error GoTo ECopiandoHorarios
    'Borramos los anteriorres ? si, no, qui lo sa
    
    'Insertamos los horarios
    Set miRsAux = New ADODB.Recordset
    C = "select calendariol.*,calendariof.fecha as f2 from calendariol left join calendariof on"
    C = C & " calendariol.idcal=calendariof.idcal  and calendariol.fecha=calendariof.fecha"
    C = C & " Where calendariol.idCal = " & Text1(26).Text
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not miRsAux.EOF
        If i > 100 Then
            EjecutaSQL C & ";"
            i = 0
        End If
        If i = 0 Then
            C = "INSERT INTO calendariot (idtrabajador, fecha, idhorario, TipoDia) VALUES "
        Else
            C = C & ","
        End If
        
        C = C & "(" & Text1(0).Text & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "'," & miRsAux!IdHorario & ","
        If IsNull(miRsAux!F2) Then
            C = C & "0)"
        Else
            C = C & "1)"
        End If
        
        miRsAux.MoveNext
        i = i + 1
    Wend
    miRsAux.Close
    If i > 0 Then EjecutaSQL C & ";"
            
    Exit Sub
ECopiandoHorarios:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub cmdHuella_Click()

    'Traido de GESLAB
    
    
    Dim Byt As Byte

    If Modo <> 2 Then
        If Modo > 2 Then
            MsgBox "Esta editando el usuario", vbExclamation
        Else
            MsgBox "Debe primero dar de alta el trabajador, antes de capturar su huella", vbExclamation
        End If
        Exit Sub
    End If
    
    If Val(Me.Slider1.Value) > 100 Or Val(Me.Slider1.Value) = 0 Then
        MsgBox "No deberia haber entrado aqui. Error valor slider", vbExclamation
        Exit Sub
    End If
    Byt = CByte(Me.Slider1.Value)
    
    Dim usu As UsuarioHuella
    Set usu = New UsuarioHuella
    
    If usu.Leer(Text1(18).Text) Then
    
    Else
        usu.CodUsuario = Text1(18)
        usu.GesLabID = Text1(0)
        usu.Mensaje = Left(Text1(1) & String(20, " "), 20)
    End If
    
    
    usu.CapturaHuella Byt
    If usu.FIR <> "" Then
        usu.Guardar
        If Dir(vEmpresa.DirHuellas & "\" & usu.CodUsuario & ".jpg") <> "" Then
            imgHuella.Picture = LoadPicture(vEmpresa.DirHuellas & "\" & usu.CodUsuario & ".jpg")
            imgHuella.Visible = True
        Else
            imgHuella.Visible = False
        End If
        
        
        conn.Execute "UPDATE trabajadores SET NoCapturarHuella =0 WHERE idtrabajador=" & Text1(0).Text
        Check2.Value = 0
        DoEvents
    End If
    Set usu = Nothing
    
    


End Sub

' *** adrede per ad este manteniment ***
Private Sub Combo1_Click(Index As Integer)
    Text1_LostFocus (16) 'el camp que te el codi del banc
End Sub
'***************************************+

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
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
      
    If vEmpresa.imgtrabaj Then
        With Me.ToolAux(0)
          .HotImageList = frmPpal.imgListComun_OM16
          .DisabledImageList = frmPpal.imgListComun_BN16
          .ImageList = frmPpal.imgListComun16
          
          .Buttons(1).Image = 1   'Insertar
          .Buttons(2).Image = 4   'Modificar
          .Buttons(3).Image = 5   'Borrar
          
          .Buttons(4).Image = 18
          
          .Buttons(5).Image = 15    'Borrar
          
          

          
          .Buttons(4).Visible = (Dir(App.Path & "\" & ExeWebCam & "") <> "")
        End With
        Me.FrameImagen.Visible = True
        
        'El adod de las imagenes
        AdodcImg.LockType = adLockPessimistic
        AdodcImg.CursorType = adOpenKeyset
        AdodcImg.ConnectionString = conn.ConnectionString
        
        
        
        
    Else
        Me.FrameImagen.Visible = False
    End If
    
    Me.SSTab1.TabVisible(2) = vEmpresa.laboral
    FrameTapaCtaBancaria.Visible = vEmpresa.laboral
    
    FrameTapaBolsaHoras.Visible = False
    FrameContadoresHoras.Visible = False
    If vEmpresa.QueEmpresa = 2 Or vEmpresa.QueEmpresa = 5 Then  'ALZIRA  COOPIC
        FrameContadoresHoras.Visible = True
        FrameTapaBolsaHoras.Visible = True
    End If
    
    ' *** canviar el nom de la taula i la clau primaria de l'ORDER BY ***
    NombreTabla = "trabajadores"
    Ordenacion = " ORDER BY idTrabajador"
    ' **********************************************************
        
    'Vemos como esta guardado el valor del check
    'ckVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where idtrabajador=-1"
    Data1.Refresh
          
    For i = 0 To Combo1.Count - 1
        CargaCombo i
    Next i
    
    imgZoom(0).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    imgEMAIL.Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Me.imgCalendariP(0).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    
    LimpiarCampos   'Limpia los campos TextBox
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow 'codclien
    End If
    
    
    FrameHuella.Visible = vEmpresa.QueEmpresa = 2 'Alzira
    
End Sub


Private Sub LimpiarCampos()
Dim i As Integer
    
    On Error Resume Next

    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    Check1(0).Value = 0
    For i = 0 To Me.Combo1.Count - 1
        Me.Combo1(i).ListIndex = -1
    Next i
    If vEmpresa.imgtrabaj Then
        ImageTra.Picture = LoadPicture("")
        ImageTra.Tag = ""
    End If
    
    Me.ListView1.ListItems.Clear
    Me.ListView2.ListItems.Clear
    lblInfCodigo.Caption = ""
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Integer, NumReg As Byte
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
    imgEMAIL.Visible = B
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg


    'Solo cuando mod>2
    If vEmpresa.imgtrabaj Then
        'modo 2 e insertar...
        ToolAux(0).Buttons(1).Enabled = Modo >= 2
        ToolAux(0).Buttons(5).Enabled = Modo >= 2
        'Solo modificando o insertantod
        ToolAux(0).Buttons(2).Enabled = Modo > 2
        ToolAux(0).Buttons(3).Enabled = Modo > 2
        
        If ToolAux(0).Buttons(4).Visible Then ToolAux(0).Buttons(4).Enabled = Modo > 2
        
    End If
    
    

    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
        
    'Bloquea los campos Text1 si no estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    
    ' **************************************************************************************************
    ' *** Revisar lo que bloquejem, el nom de la clau primaria, les imagens de buscar i les de dates ***
    BloquearText1 Me, Modo
    BloquearCheck1 Me, Modo
    BloquearCombo Me, Modo
    BloquearImgZoom Me, Modo, 0
    
    
    Me.Option1(0).Enabled = B
    Me.Option1(1).Enabled = B
'    PosicionarCombo Combo1(2), 724
'    For i = 0 To Combo1(2).ListCount - 1
'        If Combo1(2).ItemData(i) = 724 Then
'            Combo1(2).ListIndex = i
'            Exit For
'        End If
'    Next i
    ImgModifHora.Visible = Modo = 4
    imgVacas.Visible = Modo = 4
    Me.imgCalendariP(0).Visible = Modo > 2
    If Modo = 4 Then _
        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    
 '   BloquearImgBuscar Me, Modo
    BloquearImgFec Me, 11, Modo
    BloquearImgFec Me, 27, Modo
    BloquearImgFec Me, 28, Modo
 
    ' **************************************************************************************************
                
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    
    
    
    
        
    
    If vEmpresa.QueEmpresa = 2 Then
        Me.cmdHuella.Visible = Modo = 2
        Me.Slider1.Visible = Me.cmdHuella.Visible
        Me.Label8(0).Visible = Me.cmdHuella.Visible
        Me.Label8(1).Visible = Me.cmdHuella.Visible
        Me.Check2.Enabled = Modo = 1 Or Modo > 2
    End If
    
    
    
    
    
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
    'Toolbar1.Buttons(12).Enabled = B
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

'Private Sub imgFec_Click(Index As Integer)
'    Dim esq As Long
'    Dim dalt As Long
'    Dim menu As Long
'    Dim obj As Object
'
'    Set frmC = New frmCal
'
'    esq = imgFec(Index).Left
'    dalt = imgFec(Index).Top
'
'    Set obj = imgFec(Index).Container
'
'    While imgFec(Index).Parent.Name <> obj.Name
'        esq = esq + obj.Left
'        dalt = dalt + obj.Top
'        Set obj = obj.Container
'    Wend
'
'    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
'
'    ' es desplega baix i cap a la dreta
'    'frmC.Left = esq + imgFec(Index).Parent.Left + 30
'    'frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
'
'    ' es desplega dalt i cap a la esquerra
'    frmC.Left = esq + imgFec(Index).Parent.Left - frmC.Width + imgFec(Index).Width + 40
'    frmC.Top = dalt + imgFec(Index).Parent.Top - frmC.Height + menu - 25
'
'    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
'    imgFec(27).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
'    If Text1(Index).Text <> "" Then frmC.NovaData = Text1(Index).Text
'
'    frmC.Show vbModal
'    Set frmC = Nothing
'    PonerFoco Text1(CByte(imgFec(27).Tag))
    ' **************************************************************************
'End Sub

'Private Sub frmC_Selec(vFecha As Date)
'    Text1(CByte(imgFec(27).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
''    PonerFoco txtAux(CByte(imgFec(27).Tag))
'End Sub

Private Sub imgBuscar_Click(Index As Integer)
    'Screen.MousePointer = vbHourglass
'    TerminaBloquear
'    Select Case Index
'        Case 0, 8 'empresa
'            If Index = 8 Then
'                Indice = 33
'            Else
'                Indice = 3
'            End If
'            Set frmEmp = New frmEmpresas
'            frmEmp.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
'            frmEmp.CodigoActual = Text1(Indice).Text
'            frmEmp.Show vbModal
'            Set frmEmp = Nothing
'            PonerFoco Text1(Indice)
'
'        Case 1 'agencia
'            Indice = 4
'            Set frmAge = New frmAgencias2
'            frmAge.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
'            frmAge.DeConsulta = True
'            frmAge.Empresa = Text1(3).Text
'            frmAge.CodigoActual = Text1(4).Text
'            frmAge.Show vbModal
'            Set frmAge = Nothing
'            PonerFoco Text1(Indice)
'
'        Case 2, 5 'población i población banco
'            If Index = 2 Then
'                Indice = 6
'            ElseIf Index = 5 Then
'                Indice = 21
'            End If
'            Set frmPob = New frmPoblacio
'            frmPob.DatosADevolverBusqueda = "0|1|2|3|4|5|"
'            If Not IsNumeric(Text1(Indice).Text) Then Text1(Indice).Text = ""
'            frmPob.CodigoActual = Text1(Indice).Text
'            frmPob.Show vbModal
'            Set frmPob = Nothing
'            PonerFoco Text1(Indice)
'
'        Case 3 'Cuenta Contable
'            Set frmCtas = New frmCtasConta
'            frmCtas.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(15).Text) Then Text1(15).Text = ""
'            frmCtas.CodigoActual = Text1(15).Text
'            frmCtas.Show vbModal
'            Set frmCtas = Nothing
'            PonerFoco Text1(15)
'
'        Case 4 'Cuenta Bancaria
'            Set frmBan = New frmBancsofi
'            frmBan.DatosADevolverBusqueda = "4|1|3|"
'            frmBan.CodigoActual = Text1(16).Text
'            frmBan.NuevoPais = Me.Combo1(2).ItemData(Combo1(2).ListIndex)
'            frmBan.Show vbModal
'            Set frmBan = Nothing
'            PonerFoco Text1(16)
'
'        Case 6 'tipo de nómina
'            Set frmTiN = New frmTiponomi
'            frmTiN.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(24).Text) Then Text1(24).Text = ""
'            frmTiN.CodigoActual = Text1(24).Text
'            frmTiN.Show vbModal
'            Set frmTiN = Nothing
'            PonerFoco Text1(24)
'
'        Case 7 'tipo de empleado
'            Set frmTiE = New frmTiposemp
'            frmTiE.DatosADevolverBusqueda = "0|1|"
'            If Not IsNumeric(Text1(30).Text) Then Text1(30).Text = ""
'            frmTiE.CodigoActual = Text1(30).Text
'            frmTiE.Show vbModal
'            Set frmTiE = Nothing
'            PonerFoco Text1(30)
'    End Select

    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub frmc_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(CByte(imgFec(11).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub ImageTra_DblClick()
    MostrarImg
End Sub

Private Sub imgCalendariP_Click(Index As Integer)
    Set frmS = New frmCalendario
    frmS.DatosADevolverBusqueda = "0|1|"
    frmS.Show vbModal
    Set frmS = Nothing
End Sub

Private Sub imgEMAIL_Click()
    If Text1(8).Text = "" Then Exit Sub
    LanzaMailGnral Text1(8).Text
End Sub

Private Sub ImgModifHora_Click()
Dim F As Date

    miSQL = Text1(1).Text
    F = DateAdd("m", -1, Now)
    F = CDate("01/" & Month(F) & "/" & Year(F))
    miSQL = miSQL & "|" & Format(F, "dd/mm/yyyy") & "|"
    F = DateAdd("m", 1, Now)
    F = CDate(DiasMes(Month(F), Year(F)) & "/" & Month(F) & "/" & Year(F))
    miSQL = miSQL & Format(F, "dd/mm/yyyy") & "|"
    
    'Parametros
    ' Nombre | Fec ini | Fec Fin | Codigo trabjador
    miSQL = miSQL & Text1(0).Text & "|"
    
    FrmVarios.Opcion = 0
    FrmVarios.Parametros = miSQL
    FrmVarios.Show vbModal
    CargaCalendario
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim Obj As Object

    Set frmc = New frmCal
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    ' *** adrede ***
    'If Index <> 49 Then
    '    esq = imgFec(Index).Left
    '    dalt = imgFec(Index).Top
    'Else
    '    esq = btnFec(Index).Left
    '    dalt = btnFec(Index).Top
    'End If
    ' *************

'    Set obj = imgFec(Index).Container
'
'    While imgFec(Index).Parent.Name <> obj.Name
'        esq = esq + obj.Left
'        dalt = dalt + obj.Top
'        Set obj = obj.Container
'    Wend
    
    ' *** adrede ***
  '  If Index <> 49 Then
        Set Obj = imgFec(Index).Container

        While imgFec(Index).Parent.Name <> Obj.Name
            esq = esq + Obj.Left
            dalt = dalt + Obj.Top
            Set Obj = Obj.Container
        Wend
'    Else
'        Set obj = btnFec(Index).Container
'
'        While btnFec(Index).Parent.Name <> obj.Name
'            esq = esq + obj.Left
'            dalt = dalt + obj.Top
'            Set obj = obj.Container
'        Wend
'
'    End If
    ' *************
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    ' *** adredre ***
'    If Index <> 49 Then 'dreta i baix
        frmc.Left = esq + imgFec(Index).Parent.Left + 30
        frmc.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
'    Else 'esquerra i dalt
'        frmC.Left = esq + btnFec(Index).Parent.Left - frmC.Width + btnFec(Index).Width + 40
'        frmC.Top = dalt + btnFec(Index).Parent.Top - frmC.Height + menu - 25
'    End If
    ' ***************

    imgFec(11).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(Index).Text <> "" Then frmc.NovaData = Text1(Index).Text
    ' ********************************************

    frmc.Show vbModal
    Set frmc = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(11).Tag)) '<===
    ' ********************************************
End Sub



Private Sub imgVacas_Click()
    FrmVarios.Parametros = Text1(1).Text & "|" & Text1(0).Text & "|"
    FrmVarios.Opcion = 3
    FrmVarios.Show vbModal
    CargaCalendario
End Sub

Private Sub imgZoom_Click(Index As Integer)
    frmVerCalendario.CodigoTrab = Val(Text1(0).Text)
    frmVerCalendario.FeIni = vEmpresa.FechaInicio
    frmVerCalendario.FeFin = vEmpresa.FechaFin
    frmVerCalendario.idCal = Val(Text1(26).Text)
    frmVerCalendario.Texto = Text1(1).Text
    frmVerCalendario.Show vbModal
End Sub

Private Sub Label4_Click(Index As Integer)
    MsgBox "Que envio un mail eeeee"
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


Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    Select Case Index
    Case 0
        '---------------------------------
        'Tool de la imagen del trabajadpr
        If Button.Index <> 2 And Button.Index <> 4 Then
            If ImageTra.Picture = 0 Then
                MsgBox "No hay imagen disponible", vbExclamation
                Exit Sub
            End If
        End If
        Select Case Button.Index
        Case 1
            MostrarImg
        Case 2
            CargarImagen
        Case 3
            Me.ImageTra.Picture = LoadPicture("")
            Me.ImageTra.Tag = "OK"
            
        Case 4
            'WEB CAM
            ObtenerDesdeWebCam
            
        Case 5
            GuardarADisco
        End Select 'Del tool del imagen trabjador
    End Select
End Sub


Private Sub MostrarImg()

    If Modo <> 2 Then Exit Sub
    
    If ImageTra.Picture Is Nothing Then Exit Sub
    If ImageTra.Picture = 0 Then Exit Sub
    
    
    On Error GoTo ECargandoImagenGrande
    
    With FrmVarios
        
        .Opcion = 1
        .imgt.Stretch = False
        
        .imgt.Picture = ImageTra.Picture
        If .imgt.Width > 5500 Or .imgt.Height > 4800 Then
            If .imgt.Width > .imgt.Height Then
                .imgt.Height = CInt((5500 * .imgt.Height) / .imgt.Width)
                .imgt.Width = 5500
            Else
                .imgt.Width = CInt((4800 * .imgt.Width) / .imgt.Height)
                .imgt.Height = 4800
            End If
            .imgt.Stretch = True
        End If
        .Label2.Caption = Text1(1).Text
        .Show vbModal
    End With
    Exit Sub
ECargandoImagenGrande:
    MuestraError Err.Number
End Sub


Private Sub CargarImagen()
    On Error GoTo EC
    
    
    frmPpal.cd1.ShowOpen
    If frmPpal.cd1.FileName <> "" Then
        If FileLen(frmPpal.cd1.FileName) > 512000 Then
            MsgBox "Tamaño de imagen excede de lo permitido. Max 500Kb", vbExclamation
        Else
            Me.ImageTra.Picture = LoadPicture(frmPpal.cd1.FileName)
            ImageTra.Tag = "OK"
        End If
    End If
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
End Sub



Private Sub GuardarADisco()
    On Error GoTo EC
    
    
    frmPpal.cd1.ShowSave
    If frmPpal.cd1.FileName <> "" Then
        If Dir(frmPpal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo  " & frmPpal.cd1.FileName & " ya existe. ¿Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        SavePicture ImageTra.Picture, frmPpal.cd1.FileName
        MsgBox "Proceso finalizado. Se ha creado con exito el archivo: " & vbCrLf & frmPpal.cd1.FileName, vbInformation
    End If
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
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
        Option1(0).Value = False
        Option1(1).Value = False
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
Dim Cad As String
        'Llamamos a al form
        ' **************** arreglar-ho per a vore lo que es desije ****************
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 20, "Cód.")
        Cad = Cad & ParaGrid(Text1(1), 60, "Nombre")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
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
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
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
    
    For kCampo = 19 To 21
        Text1(kCampo).Text = "0,00"
    Next kCampo
    
    
    If vEmpresa.laboral Then Text1(30).Text = Format(vEmpresa.IRPF_, FormatoPorcen)
        
    
    ' ******* Canviar el nom de la taula, el nom de la clau primaria, i el
    ' nom del camp que te la clau primaria si no es Text1(0) *************
    Text1(0).Text = SugerirCodigoSiguienteStr("Trabajadores", "idTrabajador")
    FormateaCampo Text1(0)
    
    kCampo = 0
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
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    'Comprobamos si se puede eliminar
'    If Not SePuedeEliminar Then Exit Sub

    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub

    ' *************** canviar els noms, els formats i el DELETE ****************                  "
    Cad = Cad & "¿Seguro que desea eliminar el Trabajador?"
    Cad = Cad & vbCrLf & "  Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "  Nombre: " & Data1.Recordset.Fields(3)
    
    ' **************************************************************************
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
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
    Screen.MousePointer = vbHourglass
    PonerCamposForma Me, Data1
       
    Option1(1).Value = (Data1.Recordset!sexo = 1)
    Option1(0).Value = Not Option1(1).Value
    
    
    Text2.Text = DevuelveDesdeBD("descripcion", "calendario", "idcal", Text1(26).Text, "N")

    
    'Cargamos el calendario laboral
    'Del trabajador
    CargaCalendario
    

    
    '-- Esto permanece para saber donde estamos
    kCampo = 0
    
    Me.Refresh
    DoEvents
    If vEmpresa.imgtrabaj Then CargaImagen
    
    'Si tienen que cargar la carga, si no pues res
    
    If vEmpresa.QueEmpresa = 2 Then
        lblInfCodigo.Visible = Me.Check2.Value = 1
        Me.lblInfCodigo.Visible = lblInfCodigo.Visible
        If Me.Check2.Value = 1 Then
            imgHuella.Visible = False
            
            CadB = Trim(Text1(18).Text)
            If CadB <> "" Then CadB = "&H" & CadB
            PintaCodigoTrabajadorSinHuella CadB
        Else
            ImagenHuella
        End If
        lblInfCodigo.Visible = Me.Check2.Value = 1
   End If
        
    If vEmpresa.QueEmpresa = 2 Or vEmpresa.QueEmpresa = 5 Then CargaDatosBolsaHoras
    
    


    
    
    Screen.MousePointer = vbDefault
    
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
Dim B As Boolean
Dim i As Integer

    On Error GoTo EDatosOK

    'Considereaciones antes de insertar o modificar
    If vEmpresa.laboral Then
        'Los campos txtson necesarios
        If Trim(Text1(30).Text) = "" Then
            MsgBox "IRPF a cargo de la empresa es obligatorio", vbExclamation
            If Modo = 4 Then
                If Not IsNull(Data1.Recordset!IRPFempresa) Then Text1(30).Text = Format(Data1.Recordset!IRPFempresa, FormatoPorcen)
            End If
            
            If Trim(Text1(30).Text) = "" Then Text1(30).Text = Format(vEmpresa.IRPF_, FormatoPorcen)
            PonerFoco Text1(30)
            Exit Function
        End If
    End If
    Text1(17).Text = Abs(Option1(1).Value)
    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    ' ******************** canviar els arguments de la funcio i el mensage ****************
    If (Modo = 3) Then 'Insertar
        'comprobar si ya existe el campo de clave primaria
        If ExisteCP(Text1(0)) Then B = False
        
'         Datos = DevuelveDesdeBD("idTrabajador", "empleado", "idTrabajador", Text1(0).Text, "N")
'         If Datos <> "" Then
'            MsgBox "Ya existe el Código de Empleado: " & Text1(0).Text, vbExclamation
'            DatosOk = False
'            PonerFoco Text1(0)
'            Exit Function
'         End If
    End If
    ' *************************************************************************************
         
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per la clua primaria ***
    Cad = "(idTrabajador=" & Text1(0).Text & ")"
    ' ***************************************
    
    If SituarData(Data1, Cad, Indicador) Then
       PonerModo 2
       PonerCampos
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
    vWhere = " WHERE idTrabajador=" & Data1.Recordset!idTrabajador
    ' ************************************************
    conn.Execute "Delete from calendariot " & vWhere
    conn.Execute "Delete from " & NombreTabla & vWhere
               
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar" & vbCrLf & Err.Description
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

Private Sub PonTipoControl(valor As String)
Dim i As Integer
    For i = 0 To Combo1(2).ListCount - 1
        If Combo1(2).ItemData(i) = valor Then
            Combo1(2).ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean








    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    ' ***************** configurar els camps de buscar codis *****************
    Select Case Index
        Case 0, 17, 19
            PonerFormatoEntero Text1(Index)


        Case 8 'NIF
            'Text1(Index).Text = UCase(Text1(Index).Text)
            'ValidarNIF Text1(Index).Text


        Case 11, 27, 28
            If Text1(Index).Text = "" Then Exit Sub
            If Not EsFechaOK(Text1(Index)) Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
        
        Case 13 To 16
        
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoEntero Text1(Index)
            
            
                        'Cuenta bancaria
            If Index < 15 Then
                indice = 4
            Else
                If Index = 15 Then
                    indice = 2
                Else
                    indice = 10
                End If
            End If
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Cuenta banco debe ser numérico: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
                Exit Sub
            Else
                'Formateamos
                cadMen = String(CLng(indice), "0") & Text1(Index).Text
                Text1(Index).Text = Right(cadMen, indice)
                
            End If
            
            cadMen = ""
            For indice = 13 To 16
                cadMen = cadMen & Text1(indice).Text
            Next
            
            If Len(cadMen) = 20 And Index = 16 Then 'solo cuando pierde el foco la cuentaban
                'OK. Calculamos el IBAN
                
                
                If Text1(29).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cadMen, cadMen) Then Text1(29).Text = "ES" & cadMen
                Else
                    cad_meua = CStr(Mid(Text1(29).Text, 1, 2))
                    If DevuelveIBAN2(CStr(cad_meua), cadMen, cadMen) Then
                        If Mid(Text1(29).Text, 3) <> cadMen Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & cad_meua & cadMen & "]", vbExclamation
                            'Text1(49).Text = "ES" & SQL
                        End If
                    End If
                End If
            End If
            

            
            
            
            'ESTE TROZO NO SE SI DEBERIA ESTAR
            If Index = 16 Then
                If Modo = 3 Then
                    If Combo1(0).ListIndex < 0 Then Exit Sub
                    'Vemos si tiene puesto el calendario
                    cad_meua = "controlempleados"
                    If Text1(26).Text = "" Then
                        'Ofertamos el de la seccion
                        cadMen = Combo1(0).ItemData(Combo1(0).ListIndex)
                        cadMen = DevuelveDesdeBD("idcal", "secciones", "idseccion", cadMen, "N", cad_meua)
                        If Not IsNumeric(cadMen) Then cadMen = ""
                        If cadMen <> "" Then
                            Text1(26).Text = cadMen
                            If Combo1(2).ListIndex < 0 Then
                                'NO TIENE SELECCIONADO EL TIPO
                                 PonTipoControl cad_meua
                            End If
                            Text1_LostFocus 26
                        End If
                    Else
                        If Combo1(2).ListIndex < 0 Then
                            cadMen = Combo1(0).ItemData(Combo1(0).ListIndex)
                            cadMen = DevuelveDesdeBD("controlempleados", "secciones", "idseccion", cadMen, "N")
                            PonTipoControl cadMen
                        End If
                    End If
                    
                End If
            End If
        Case 26
            If Text1(Index).Text = "" Then
                Text2.Text = ""
                Exit Sub
            End If
            Nuevo = False 'Para poner el foco sobre el otra vez
            If PonerFormatoEntero(Text1(Index)) Then
                cadMen = DevuelveDesdeBD("descripcion", "calendario", "idcal", Text1(Index).Text, "N")
                If cadMen = "" Then
                    MsgBox "Calendario no existe", vbExclamation
                    Nuevo = True
                End If
            Else
                Nuevo = True
                cadMen = ""
            End If
            Text2.Text = cadMen
            If Nuevo Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
'        Case 9, 10 'telèfons, fax i mòbils
'            PosarFormatTelefon Text1(Index)

'        Case 12, 13 'loginweb i passwweb
'            Text1(Index).Text = LCase(Text1(Index).Text)
'            If Index = 12 Then 'login
'                If Not ComprobarLoginEmp(Text1(Index).Text) Then PonerFoco Text1(Index)
'            End If
'
'        Case 6, 21 'poblacion
'            Nuevo = False
'            If Index = 6 Then
''                PonerDatosPoblacion Text1(Index), Text2(Index), Text1(Index + 1), , , Nuevo, Text1(9)
'            ElseIf Index = 21 Then
''                PonerDatosPoblacion Text1(Index), Text2(Index), Text1(Index + 1), , , Nuevo
'            End If
'            If Nuevo Then
''                Indice = Index
''                Set frmPob = New frmPoblacio
''                frmPob.DatosADevolverBusqueda = "0|1|2|3|4|5|"
''                frmPob.NuevoCodigo = Text1(Index).Text
''                Text1(Index).Text = ""
''                TerminaBloquear
''                frmPob.Show vbModal
''                Set frmPob = Nothing
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
'        Case 27, 28, 29 'dates
'            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
    End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
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

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
   ' imgFec_Click (Indice)
End Sub
Private Sub CargaCombo(Index As Integer)


    Combo1(Index).Clear

    ' ******* configurar els distints Combos **********
    Select Case Index
        Case 0  'Seccion
            cad_meua = "SELECT * FROM secciones ORDER BY idseccion"
            Set miRs = New ADODB.Recordset
            miRs.Open cad_meua, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

            While Not miRs.EOF
                Combo1(Index).AddItem miRs!Nombre
                Combo1(Index).ItemData(Combo1(Index).NewIndex) = miRs!IdSeccion
                miRs.MoveNext
            Wend

            miRs.Close
            
        Case 1 'forma de cobro
            
            cad_meua = "SELECT * FROM categorias ORDER BY idcategoria"
            Set miRs = New ADODB.Recordset
            miRs.Open cad_meua, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

            While Not miRs.EOF
                Combo1(Index).AddItem miRs!nomcategoria
                Combo1(Index).ItemData(Combo1(Index).NewIndex) = miRs!idCategoria
                miRs.MoveNext
            Wend

            miRs.Close
            
        Case 2
            
            cad_meua = "SELECT * FROM stipocontrol ORDER BY tipocontrol"
            Set miRs = New ADODB.Recordset
            miRs.Open cad_meua, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

            While Not miRs.EOF
                Combo1(Index).AddItem miRs!desccontrol
                Combo1(Index).ItemData(Combo1(Index).NewIndex) = miRs!TipoControl
                miRs.MoveNext
            Wend

            miRs.Close
            
        Case 3
            If Not vEmpresa.laboral Then Exit Sub
            
            cad_meua = "SELECT * FROM tipocontrato ORDER BY idcontrato"
            Set miRs = New ADODB.Recordset
            miRs.Open cad_meua, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

            While Not miRs.EOF
                Combo1(Index).AddItem miRs!desccontrato
                Combo1(Index).ItemData(Combo1(Index).NewIndex) = miRs!idcontrato
                miRs.MoveNext
            Wend

            miRs.Close
    End Select
    
'    If Index <> 2 Then 'excepte per a ibanpais sempre seleccione la 1ª opció per defecte
'        Combo1(Index).ListIndex = 0
'    Else
'        'per defecte seleccione ES
'        PosicionarCombo Combo1(Index), 724
'        For j = 0 To Combo1(Index).ListCount - 1
''            If Combo1(Index).ItemData(j) = 724 Then
''                Combo1(Index).ListIndex = j
''                Exit For
''            End If
''        Next j
'    End If
'    ' **************************************************************
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
        'Este NO es printnou.
        'Llamamos a listado y punto-pelota
        frmListado.Opcion = 8
        frmListado.Show vbModal
End Sub


Private Sub InsertaFestivo(F1 As String, F2 As Date)
    Dim IT As ListItem
    Set IT = ListView2.ListItems.Add()
    IT.Text = F1
    F2 = DateAdd("d", -1, F2)
    If F1 <> F2 Then IT.SubItems(1) = Format(F2, "dd/mm/yyyy")
    
End Sub

Private Sub InsertaItem(H As Integer, FI As Date, FF As Date)
Dim IT As ListItem
    Set IT = ListView1.ListItems.Add()
    IT.Text = Format(FI, "dd/mm/yyyy")
    FF = DateAdd("d", -1, FF)
    If FI <> FF Then
        IT.SubItems(1) = Format(FF, "dd/mm/yyyy")
    Else
        IT.SubItems(1) = ""
    End If
    miSQL = DevuelveDesdeBD("nomhorario", "horarios", "idhorario", CStr(kCampo), "N")
    IT.SubItems(2) = miSQL
 
    If Now >= FI Then
        If Format(Now, "dd/mm/yyyy") <= FF Then
            IT.EnsureVisible
            IT.Selected = True
            IT.Bold = True
            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(1).ForeColor = vbBlue
            IT.ListSubItems(2).ForeColor = vbBlue
            IT.ForeColor = vbBlue
            Set ListView1.SelectedItem = IT
        End If
    End If
End Sub

Private Sub CargaCalendario()
Dim F As Date
Dim F2 As Date
Dim vFe As String
Dim IniVacas As Date
Dim FinVacas As Date


    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    Set miRs = New ADODB.Recordset
    'F = DateAdd("d", -14, Now)
    'F2 = DateAdd("m", 1, Now)
    F = vEmpresa.FechaInicio
    F2 = vEmpresa.FechaFin
    miSQL = "Select * from calendarioT where fecha >= '" & Format(F, FormatoFecha) & "'"
    miSQL = miSQL & " and fecha <= '" & Format(F2, FormatoFecha) & "' and idtrabajador =" & Data1.Recordset!idTrabajador
    miSQL = miSQL & " order by fecha"
    miRs.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    kCampo = 0
    vFe = ""
    F = vEmpresa.FechaInicio
    While Not miRs.EOF
        If kCampo <> miRs!IdHorario Then
            If kCampo > 0 Then InsertaItem kCampo, F, miRs!Fecha
            F = miRs!Fecha
            kCampo = miRs!IdHorario
        End If
        If vFe = "" Then
            If miRs!tipodia > 0 Then
                If vFe = "" Then vFe = miRs!Fecha
                
            End If
        Else
            If miRs!tipodia = 0 Then
                'INSERTAR INICIO FIN VACIONES
                InsertaFestivo vFe, miRs!Fecha
                vFe = ""
            End If
        End If
        
        F2 = miRs!Fecha
        miRs.MoveNext
    Wend
    If kCampo > 0 Then
        If F <> F2 Then InsertaItem kCampo, F, F2
    End If
    
    'Ha iniciado las vaciones y no las ha acabado a 31 diciembre
    If vFe <> "" Then InsertaFestivo vFe, vEmpresa.FechaFin
        
    
    
    miRs.Close
    
    
    
    
    Set miRs = Nothing
End Sub


Private Sub CargaImagen()


On Error GoTo EEnlazaImagen
     
    
    AdodcImg.RecordSource = "Select * from timagen where idtrabajador =" & Data1.Recordset!idTrabajador
    AdodcImg.Refresh
    If Not AdodcImg.Recordset.EOF Then
        LeerBinary AdodcImg.Recordset.Fields(1), Me.ImageTra
    Else
        Me.ImageTra.Picture = LoadPicture("")
    End If
    Exit Sub
EEnlazaImagen:
    MuestraError Err.Number, "Mostrar imagen." & Err.Description
    Me.ImageTra.Picture = LoadPicture("")
End Sub


Private Sub TratarImagen()
On Error GoTo Etrata
    'NO SE HA TOCADO NADA
    If ImageTra.Tag = "" Then Exit Sub
    
    If ImageTra.Picture.Width > 0 Then
        
        'Tiene imagen
        If AdodcImg.Recordset.EOF Then
            'NUEVA
            AdodcImg.Recordset.AddNew
            If Modo = 3 Then
                AdodcImg.Recordset.Fields(0) = Text1(0).Text
            Else
                AdodcImg.Recordset.Fields(0) = Data1.Recordset.Fields!idTrabajador
            End If
        End If
        GuardarBinary AdodcImg.Recordset.Fields(1), Me.ImageTra
        AdodcImg.Recordset.Update
        espera 0.4
    Else
        conn.Execute "DELETE FROM timagen where idtrabajador=" & Data1.Recordset.Fields!idTrabajador
    End If
    
    Exit Sub
Etrata:
    MuestraError Err.Number, "Guardando imagen en BD " & Err.Description
    AdodcImg.Recordset.CancelUpdate
End Sub





Private Sub PintaCodigoTrabajadorSinHuella(ByRef C As String)
    On Error Resume Next
    If C = "" Then
        lblInfCodigo.Caption = ""
    Else
        C = CStr(CLng(C))
        If Val(C) > 9999 Then
            lblInfCodigo.Caption = "> max"
        Else
            lblInfCodigo.Caption = Format(C, "0000")
        End If
    End If
   
    If Err.Number <> 0 Then
        Err.Clear
        lblInfCodigo.Caption = "N/D"
    End If
End Sub



Private Sub ImagenHuella()
Dim Husu As UsuarioHuella
Dim CreadoOBj As Boolean
    On Error GoTo EIm
    
    
    'YA se que da error
    If ErrorAcecionedoOCX = 2 Then Exit Sub
    
    
    Set Husu = New UsuarioHuella
    CreadoOBj = True
    If Husu.Leer(Trim(Text1(18))) Then
        
        If Dir(vEmpresa.DirHuellas & "\" & Husu.CodUsuario & ".jpg") <> "" Then
            
            imgHuella.Picture = LoadPicture(vEmpresa.DirHuellas & "\" & Husu.CodUsuario & ".jpg")
            imgHuella.Visible = True
        Else
            
            imgHuella.Visible = False
        End If
    Else
        
        imgHuella.Visible = False
    End If


EIm:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargando huella:" & ErrorAcecionedoOCX
        If ErrorAcecionedoOCX = 0 Then
            If CreadoOBj Then
                ErrorAcecionedoOCX = 1
            Else
                ErrorAcecionedoOCX = 2
            End If
        End If
    End If
    Set Husu = Nothing
End Sub




Private Sub ObtenerDesdeWebCam()
Dim C As String


    On Error GoTo EC
    'Abro el programa
    'Llamandolo con mi nombre de archivo como parametro
    C = App.Path & "\pr1.jpg"
    If Dir(C, vbArchive) <> "" Then Kill C
    
    C = """" & App.Path & "\aricam.exe" & """ """ & C & """"""
        
    Me.Caption = "Leyendo Webcam"
    
    DoEvents
    
    Lanza_EXE_Y_Espera C
    
    C = App.Path & "\pr1.jpg"
    If Dir(C, vbArchive) <> "" Then
        ImageTra.Picture = LoadPicture(C)
        ImageTra.Tag = "OK"
    End If
    
    Me.Caption = "Trabajadores"
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub CargaDatosBolsaHoras()
Dim IT
    Set miRs = New ADODB.Recordset
    ListView3.ListItems.Clear
    If vEmpresa.QueEmpresa = 2 Then
        'ALZIRA
        miSQL = "select Desctipohora,if(paraempresa=0,'Si','')"
    Else
        'COOPIC
        miSQL = "select Desctipohora,if(paraempresa=1,'Si','')"
    End If
    miSQL = miSQL & " Lacope,horasbolsa from trabajadoresbolsahoras,"
    miSQL = miSQL & " tiposhora where tiposhora.TipoHora = trabajadoresbolsahoras.tipohora"
    miSQL = miSQL & " AND idtrabajador =" & Data1.Recordset!idTrabajador
    miSQL = miSQL & " order by Lacope desc,Desctipohora"
    miRs.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRs.EOF
        Set IT = ListView3.ListItems.Add()
        IT.Text = miRs!DescTipoHora
        IT.SubItems(1) = miRs!Lacope
        IT.SubItems(2) = Format(miRs!horasbolsa, "0.00")
        miRs.MoveNext
    Wend
    miRs.Close
End Sub
