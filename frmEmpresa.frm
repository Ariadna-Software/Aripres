VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresa"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   10275
   Icon            =   "frmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   54
      Top             =   1080
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmEmpresa.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(24)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(27)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(28)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(10)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(11)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(23)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(32)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(33)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Check1(8)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Configuracion"
      TabPicture(1)   =   "frmEmpresa.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(34)"
      Tab(1).Control(1)=   "Text1(30)"
      Tab(1).Control(2)=   "Text1(29)"
      Tab(1).Control(3)=   "Check2"
      Tab(1).Control(4)=   "Check3"
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(6)=   "Text1(25)"
      Tab(1).Control(7)=   "Check1(6)"
      Tab(1).Control(8)=   "Check4"
      Tab(1).Control(9)=   "Combo2"
      Tab(1).Control(10)=   "Label3(3)"
      Tab(1).Control(11)=   "Label3(9)"
      Tab(1).Control(12)=   "Label3(8)"
      Tab(1).Control(13)=   "Label3(4)"
      Tab(1).Control(14)=   "Label1(23)"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Presencia"
      TabPicture(2)   =   "frmEmpresa.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text3(0)"
      Tab(2).Control(1)=   "Text1(31)"
      Tab(2).Control(2)=   "FrameTapaCortesia"
      Tab(2).Control(3)=   "Text3(1)"
      Tab(2).Control(4)=   "Text1(12)"
      Tab(2).Control(5)=   "Text1(13)"
      Tab(2).Control(6)=   "Text1(14)"
      Tab(2).Control(7)=   "Text1(15)"
      Tab(2).Control(8)=   "Text1(16)"
      Tab(2).Control(9)=   "Text2(0)"
      Tab(2).Control(10)=   "Text2(1)"
      Tab(2).Control(11)=   "Text2(2)"
      Tab(2).Control(12)=   "Text2(3)"
      Tab(2).Control(13)=   "Text2(4)"
      Tab(2).Control(14)=   "Text1(17)"
      Tab(2).Control(15)=   "Text1(18)"
      Tab(2).Control(16)=   "Combo1"
      Tab(2).Control(17)=   "Text1(19)"
      Tab(2).Control(18)=   "Text1(20)"
      Tab(2).Control(19)=   "Text1(21)"
      Tab(2).Control(20)=   "Text2(5)"
      Tab(2).Control(21)=   "Text1(24)"
      Tab(2).Control(22)=   "Label1(16)"
      Tab(2).Control(23)=   "Label1(26)"
      Tab(2).Control(24)=   "Label1(11)"
      Tab(2).Control(25)=   "Label3(0)"
      Tab(2).Control(26)=   "Label1(12)"
      Tab(2).Control(27)=   "Label1(13)"
      Tab(2).Control(28)=   "Label1(14)"
      Tab(2).Control(29)=   "Label1(15)"
      Tab(2).Control(30)=   "Label1(17)"
      Tab(2).Control(31)=   "Label1(18)"
      Tab(2).Control(32)=   "Label3(1)"
      Tab(2).Control(33)=   "Label1(19)"
      Tab(2).Control(34)=   "Label1(20)"
      Tab(2).Control(35)=   "Label1(21)"
      Tab(2).Control(36)=   "Label1(22)"
      Tab(2).Control(37)=   "imgBuscar(0)"
      Tab(2).Control(38)=   "imgBuscar(1)"
      Tab(2).Control(39)=   "imgBuscar(2)"
      Tab(2).Control(40)=   "imgBuscar(3)"
      Tab(2).Control(41)=   "imgBuscar(4)"
      Tab(2).Control(42)=   "Line1"
      Tab(2).Control(43)=   "imgBuscar(5)"
      Tab(2).Control(44)=   "Label1(25)"
      Tab(2).ControlCount=   45
      TabCaption(3)   =   "Laboral"
      TabPicture(3)   =   "frmEmpresa.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameHorasAcabalgadas"
      Tab(3).Control(1)=   "Check1(5)"
      Tab(3).Control(2)=   "Check1(7)"
      Tab(3).Control(3)=   "Check1(3)"
      Tab(3).Control(4)=   "Check1(4)"
      Tab(3).Control(5)=   "Check1(2)"
      Tab(3).Control(6)=   "Check1(1)"
      Tab(3).Control(7)=   "Check1(0)"
      Tab(3).Control(8)=   "Text1(22)"
      Tab(3).Control(9)=   "Label3(2)"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Producción"
      TabPicture(4)   =   "frmEmpresa.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.CheckBox Check1 
         Caption         =   "SEPA  XML"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   111
         Tag             =   "1|N|S|||empresas|SepaXML|||"
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Frame FrameHorasAcabalgadas 
         Height          =   1695
         Left            =   -68760
         TabIndex        =   108
         Top             =   1920
         Width           =   3255
         Begin VB.CheckBox Check5 
            Alignment       =   1  'Right Justify
            Caption         =   "Acabalgado es del dia siguiente "
            Height          =   255
            Left            =   90
            TabIndex        =   46
            Tag             =   "1|N|N|||empresas|AcabalDiaTrabajado|||"
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   36
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   47
            Tag             =   "1|H|S|||empresas|AcabalHora|hh:mm||"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   35
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   48
            Tag             =   "Acaba. hoas|N|S|||empresas|AcabalIncrementoxDia|||"
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Incremento horas trabajadas"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   110
            Top             =   1230
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Hora"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   109
            Top             =   810
            Width           =   975
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Horario nocturno"
         Height          =   255
         Index           =   5
         Left            =   -70440
         TabIndex        =   45
         Tag             =   "1|N|S|||empresas|horarionocturno|||"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Crea calendario cada trabajador"
         Height          =   255
         Index           =   7
         Left            =   -70440
         TabIndex        =   44
         Tag             =   "1|N|S|||empresas|CreaCalDiariaTra|||"
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   34
         Left            =   -69840
         MaxLength       =   255
         TabIndex        =   106
         Tag             =   "1|T|S|||empresas|Pathproces|||"
         Top             =   2640
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   33
         Left            =   7080
         MaxLength       =   3
         TabIndex        =   104
         Tag             =   "1|T|S|||empresas|sufijoN34|||"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   32
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "1|T|S|||empresas|iban|||"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -68400
         TabIndex        =   31
         Text            =   "Text3"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   -66480
         MaxLength       =   4
         TabIndex        =   35
         Tag             =   "1|N|N|||empresas|repeticion|||"
         Top             =   2475
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   30
         Left            =   -67920
         MaxLength       =   35
         TabIndex        =   19
         Tag             =   "1|T|S|||empresas|Nomproces|||"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   -74760
         MaxLength       =   255
         TabIndex        =   21
         Tag             =   "1|T|S|||empresas|dirmarcajes|||"
         Top             =   2640
         Width           =   4575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Produccion"
         Height          =   255
         Left            =   -72960
         TabIndex        =   15
         Tag             =   "1|N|N|||empresas|produccion|||"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Laboral"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Tag             =   "1|N|S|||empresas|laboral|||"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Frame FrameTapaCortesia 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -69720
         TabIndex        =   97
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -66240
         TabIndex        =   32
         Text            =   "Text3"
         Top             =   840
         Width           =   735
      End
      Begin VB.Frame Frame3 
         Caption         =   "Usuarios"
         Height          =   615
         Left            =   -74760
         TabIndex        =   93
         Top             =   3240
         Width           =   9495
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   28
            Left            =   7560
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   24
            Tag             =   "1|T|S|||empresas|pass|||"
            Top             =   195
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   27
            Left            =   4560
            MaxLength       =   20
            TabIndex        =   23
            Tag             =   "1|T|S|||empresas|usuario|||"
            Top             =   195
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   26
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   22
            Tag             =   "1|T|S|||empresas|servidor|||"
            Top             =   195
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Password"
            Height          =   195
            Index           =   7
            Left            =   6720
            TabIndex        =   96
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   6
            Left            =   3960
            TabIndex        =   95
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label3 
            Caption         =   "SERVIDOR"
            Height          =   195
            Index           =   5
            Left            =   960
            TabIndex        =   94
            Top             =   240
            Width           =   840
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   -73080
         MaxLength       =   255
         TabIndex        =   20
         Tag             =   "1|T|S|||empresas|configreloj|||"
         Top             =   1800
         Width           =   7815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Generación marcajes automatica"
         Height          =   255
         Index           =   6
         Left            =   -68040
         TabIndex        =   17
         Tag             =   "1|N|S|||empresas|todoslosdias|||"
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Imagenes trabajadores"
         Height          =   255
         Left            =   -70680
         TabIndex        =   16
         Tag             =   "1|N|N|||empresas|imgtrabaj|||"
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmEmpresa.frx":0098
         Left            =   -73560
         List            =   "frmEmpresa.frx":00AE
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "1|N|S|||empresas|reloj|||"
         Top             =   1290
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   -73440
         MaxLength       =   4
         TabIndex        =   25
         Tag             =   "Incidencias|N|N|||empresas|IncHoraExtra|||"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -73440
         MaxLength       =   4
         TabIndex        =   26
         Tag             =   "Incidencias|N|N|||empresas|incretraso|||"
         Top             =   1392
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -73440
         MaxLength       =   4
         TabIndex        =   27
         Tag             =   "Incidencias|N|N|||empresas|incmarcaje|||"
         Top             =   1944
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   -73440
         MaxLength       =   4
         TabIndex        =   28
         Tag             =   "Incidencias|N|N|||empresas|incvacaciones|||"
         Top             =   2496
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -73440
         MaxLength       =   4
         TabIndex        =   29
         Tag             =   "Incidencias|N|N|||empresas|IncHoraExceso|||"
         Top             =   3105
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72600
         TabIndex        =   75
         Text            =   "Text2"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -72600
         TabIndex        =   74
         Text            =   "Text2"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72600
         TabIndex        =   73
         Text            =   "Text2"
         Top             =   1944
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72600
         TabIndex        =   72
         Text            =   "Text2"
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -72600
         TabIndex        =   71
         Text            =   "Text2"
         Top             =   3105
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   -69720
         MaxLength       =   4
         TabIndex        =   98
         Tag             =   "1|N|N|||empresas|maxretraso|||"
         Top             =   840
         Width           =   150
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   -69720
         MaxLength       =   4
         TabIndex        =   99
         Tag             =   "1|N|N|||empresas|maxexceso|||"
         Top             =   840
         Width           =   150
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmEmpresa.frx":00F0
         Left            =   -68640
         List            =   "frmEmpresa.frx":0100
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Tag             =   "1|N|S|||empresas|redondeo|||"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   -68400
         MaxLength       =   4
         TabIndex        =   34
         Tag             =   "1|N|N|||empresas|minutosredondeo|||"
         Top             =   2475
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   -68400
         MaxLength       =   4
         TabIndex        =   36
         Tag             =   "1|N|N|||empresas|ajusteentrada|||"
         Text            =   " "
         Top             =   3555
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   -66480
         MaxLength       =   4
         TabIndex        =   37
         Tag             =   "1|N|N|||empresas|ajustesalida|||"
         Top             =   3555
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -72600
         TabIndex        =   70
         Text            =   "Text2"
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   -73440
         MaxLength       =   4
         TabIndex        =   30
         Tag             =   "Incidencias|N|N|||empresas|IncTarjError|||"
         Top             =   3600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Horas compensables (sem 40): Extra"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   41
         Tag             =   "1|N|S|||empresas|EmpresaHoraExtra|||"
         Top             =   2640
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calculo horas nom. automática"
         Height          =   255
         Index           =   4
         Left            =   -74400
         TabIndex        =   42
         Tag             =   "1|N|S|||empresas|NominaAutomatica|||"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Abonos separados en anticpos (HN/HC)"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   40
         Tag             =   "1|N|S|||empresas|abonosseparados|||"
         Top             =   2040
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Aplica Antiguedad HC"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   39
         Tag             =   "1|N|S|||empresas|AplicaAntiguedadHC|||"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Aplica Antiguedad HN"
         Height          =   255
         Index           =   0
         Left            =   -74400
         TabIndex        =   38
         Tag             =   "1|N|S|||empresas|AplicaAntiguedadHN|||"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   -69120
         MaxLength       =   10
         TabIndex        =   43
         Tag             =   "1|N|S|||empresas|irpf|||"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "1|F|S|||empresas|fechainicio|dd/mm/yyyy||"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   66
         Tag             =   "1|T|S|||empresas|cuenta|||"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   64
         Tag             =   "1|T|S|||empresas|codcontrol|||"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   3960
         MaxLength       =   4
         TabIndex        =   9
         Tag             =   "1|T|S|||empresas|sucursal|||"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "1|T|S|||empresas|entidad|||"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   5880
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "1|T|S|||empresas|cif|||"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "1|T|S|||empresas|telempresa|||"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "1|T|S|||empresas|codposempresa|||"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   360
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "1|T|S|||empresas|provempresa|||"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "1|T|S|||empresas|pobempresa|||"
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   360
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "D|T|S|||empresas|dirempresa|||"
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "PATH procesados"
         Height          =   255
         Index           =   3
         Left            =   -69840
         TabIndex        =   107
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Sufijo"
         Height          =   255
         Index           =   28
         Left            =   7080
         TabIndex        =   105
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN"
         Height          =   195
         Index           =   27
         Left            =   2400
         TabIndex        =   103
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Max. retraso"
         Height          =   255
         Index           =   16
         Left            =   -69360
         TabIndex        =   84
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Repetición"
         Height          =   195
         Index           =   26
         Left            =   -67440
         TabIndex        =   102
         Top             =   2520
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre archivo"
         Height          =   255
         Index           =   9
         Left            =   -69240
         TabIndex        =   101
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "PATH Marcajes"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   100
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Configuracion RELOJ"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   92
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Reloj presencia"
         Height          =   255
         Index           =   23
         Left            =   -74760
         TabIndex        =   91
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Hora extra"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   90
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Incidencias"
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
         Left            =   -74760
         TabIndex        =   89
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Retraso"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   88
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Error marcaje"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   87
         Top             =   1959
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Vacaciones"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   86
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Hora exceso"
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   85
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Max. exceso "
         Height          =   255
         Index           =   17
         Left            =   -67440
         TabIndex        =   83
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Redondeo"
         Height          =   255
         Index           =   18
         Left            =   -69720
         TabIndex        =   82
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Ajustes (min)"
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
         Left            =   -67920
         TabIndex        =   81
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Redondeo horas"
         Height          =   195
         Index           =   19
         Left            =   -69720
         TabIndex        =   80
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Redondeo entrada / salida"
         Height          =   195
         Index           =   20
         Left            =   -69720
         TabIndex        =   79
         Top             =   3120
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Entrada"
         Height          =   195
         Index           =   21
         Left            =   -69240
         TabIndex        =   78
         Top             =   3600
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Salida"
         Height          =   195
         Index           =   22
         Left            =   -67440
         TabIndex        =   77
         Top             =   3600
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -73680
         MouseIcon       =   "frmEmpresa.frx":013F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar población"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   -73680
         MouseIcon       =   "frmEmpresa.frx":0291
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar población"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -73680
         MouseIcon       =   "frmEmpresa.frx":03E3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar población"
         Top             =   1966
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   -73680
         MouseIcon       =   "frmEmpresa.frx":0535
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar población"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   -73680
         MouseIcon       =   "frmEmpresa.frx":0687
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar población"
         Top             =   3127
         Width           =   240
      End
      Begin VB.Line Line1 
         X1              =   -69840
         X2              =   -65280
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   -73680
         MouseIcon       =   "frmEmpresa.frx":07D9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar población"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Baja"
         Height          =   255
         Index           =   25
         Left            =   -74760
         TabIndex        =   76
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "IRPF empresa"
         Height          =   255
         Index           =   2
         Left            =   -70440
         TabIndex        =   69
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio temporada"
         Height          =   255
         Index           =   24
         Left            =   8160
         TabIndex        =   68
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta bancaria"
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
         Left            =   720
         TabIndex        =   67
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   255
         Index           =   9
         Left            =   5400
         TabIndex        =   65
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "CC"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   63
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal"
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   62
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Entidad"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   61
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "C.I.F."
         Height          =   255
         Index           =   5
         Left            =   5880
         TabIndex        =   60
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   59
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "C.P"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   58
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   57
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   56
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   55
         Top             =   600
         Width           =   975
      End
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
      Height          =   330
      Index           =   1
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "1|T|N|||empresas|nomempresa|||"
      Top             =   555
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   52
      Tag             =   "1|N|N|0|1|empresas|idempresa||S|"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7740
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9000
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9000
      TabIndex        =   50
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   2385
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
         Height          =   255
         Left            =   40
         TabIndex        =   49
         Top             =   240
         Width           =   2295
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
      TabIndex        =   51
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
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
      Index           =   10
      Left            =   240
      TabIndex        =   53
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public DeConsulta As Boolean

'
'' *** per a cridar ad atres formularis ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
'Private WithEvents frmPob As frmPoblacio
'Private WithEvents frmBan As frmBancsofi
'' *************************************


Private HaDevueltoDatos As Boolean

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
Dim FormatoCod As String 'formato del campo código
Dim NomTabla As String
Dim Ordenacion As String

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
Dim Indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos

Dim PrimeraVez As Boolean

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim NumReg As Byte

    On Error GoTo EPonerModo
    
    Modo = vModo
    If Modo = 2 Then
        lblIndicador.Caption = PonerContRegistros(Me.adodc1)
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    
    
    
     '---------------------------------------------
    B = Modo > 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
    
    
    BloquearText1 Me, Modo
    'Fecha alta siempre bloqueada
    
    
    BloquearImgBuscar Me, Modo
    BloquearCmb Combo1, (Modo <> 1 And Modo <> 3 And Modo <> 4)
    Me.Combo2.Enabled = Combo1.Enabled
    BloquearChecks Me, Modo
    
    
    
    ' ********************************************************
    
    
    
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    
    Text3(0).BackColor = Text1(19).BackColor
    Text3(1).BackColor = Text1(19).BackColor
    Text3(0).Locked = Text1(19).Locked
    Text3(1).Locked = Text1(19).Locked
    
    Me.FrameHorasAcabalgadas.Visible = False
    If Modo = 4 Then
        If DBLet(Me.adodc1.Recordset!AcabalIncrementoxDia, "N") = 1 Then FrameHorasAcabalgadas.Visible = True
    End If
EPonerModo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner modo.", Err.Description
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2) Or Modo = 0
    'Modificar
    Toolbar1.Buttons(7).Enabled = B
    'Me.mnModificar.Enabled = b
End Sub


Private Sub BotonAnyadir()
Dim NumF As String
    
    LimpiarCampos 'Vacía los TextBox
    CadB = ""
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
     '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("entradafichajes", "Secuencia")
    End If
    
    ' ******* Canviar el nom de la taula, el nom de la clau primaria, i el
    ' nom del camp que te la clau primaria si no es Text1(0) *************
    Text1(0).Text = NumF
    FormateaCampo Text1(0)
    
    
    'PosicionarCombo Me.Combo1(0), 724
    
    'PosarDescripcions
    PonerFoco Text1(1)
    ' ********************************************************************
End Sub


Private Sub BotonVerTodos()
    CadB = ""
    LimpiarCampos 'Limpia los Text1
    

        CadenaConsulta = "Select * from " & NomTabla & Ordenacion
        PonerCadenaBusqueda

End Sub




Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Me.adodc1.RecordSource = CadenaConsulta
    adodc1.Refresh
    If adodc1.Recordset.RecordCount <= 0 Then
        If CadB = "" Then
            PonerModo 3
            
'            Screen.MousePointer = vbDefault
'            Exit Sub
        Else
            If Modo = 1 Then MsgBox "Ningún registro encontrado para el criterio de búsqueda.", vbInformation
            PonerFoco Text1(Indice)
        End If
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        adodc1.Recordset.MoveFirst
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
   If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0)
'        PosicionarCombo Combo1(0), 754
        Text1(0).BackColor = vbYellow
    End If
End Sub


Private Sub BotonModificar()
    
    PonerModo 4
   
    'Como es modificar
    ' *** primer control que no siga clau primaria ***
    PonerFoco Text1(1)
    ' ************************************************
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonEliminar()
Dim Sql As String

    On Error GoTo EEliminar
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar el guía de viaje?"
    Sql = Sql & vbCrLf & "Código: " & Text1(0).Text
    Sql = Sql & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1) & "  " & Me.adodc1.Recordset!Ape1Guia & "  " & DBLet(Me.adodc1.Recordset!ape2guia, "T")
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        Sql = "Delete from " & NomTabla & " where codguiav=" & adodc1.Recordset!Codguiav
        conn.Execute Sql
        
        If SituarDataTrasEliminar(adodc1, NumRegElim) Then
            PonerCampos
        Else
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

Private Sub Check1_Click(Index As Integer)
    If Index = 5 Then
        If Modo > 2 Then
            Me.FrameHorasAcabalgadas.Visible = Me.Check1(Index).Value = 1
        End If
    End If
End Sub

Private Sub Check2_Click()
    'produccion
    SSTab1.TabVisible(4) = Check2.Value
End Sub

Private Sub Check3_Click()
    'laboral
    SSTab1.TabVisible(3) = Check3.Value
End Sub

Private Sub cmdAceptar_Click()

    Select Case Modo
         Case 1  'BUSQUEDA
            HacerBusqueda
    
        Case 3 'INSERTAR
            Text1(0).Text = 0
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CadenaConsulta = "select * from " & NomTabla
'                    CadenaConsulta = CadenaConsulta & " WHERE secuencia=" & Text1(0).Text
'                    CadenaConsulta = CadenaConsulta & Ordenacion
                    Me.adodc1.RecordSource = CadenaConsulta '"Select * from " & NomTabla & Ordenacion
                    Me.adodc1.Refresh
                    PonerModo 2
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    PonerModo 2
                    MsgBox "Se debe reiniciar la aplicación", vbExclamation
                    End
                End If
            End If
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


Private Sub cmdRegresar_Click()
'Dim cad As String
'Dim i As Integer
'Dim j As Integer
'Dim Aux As String
'
'    If adodc1.Recordset.EOF Then
'        MsgBox "Ningún registro devuelto.", vbExclamation
'        Exit Sub
'    End If
'    cad = ""
'    i = 0
'    Do
'        j = i + 1
'        i = InStr(j, DatosADevolverBusqueda, "|")
'        If i > 0 Then
'            Aux = Mid(DatosADevolverBusqueda, j, i - j)
'            j = Val(Aux)
'            cad = cad & adodc1.Recordset.Fields(j) & "|"
'        End If
'    Loop Until i = 0
'    RaiseEvent DatoSeleccionado(cad)
'    Unload Me
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        BotonVerTodos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    ' ICONITOS DE LA BARRA
    btnPrimero = 15 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        '.Buttons(2).Image = 1   'Buscar
        '.Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        '.Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        '.Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        '.Buttons(11).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        '14 y 15 separadors
        '.Buttons(btnPrimero).Image = 6  'Primero
        '.Buttons(btnPrimero + 1).Image = 7 'Anterior
        '.Buttons(btnPrimero + 2).Image = 8 'Siguiente
        '.Buttons(btnPrimero + 3).Image = 9 'Último
    End With

    'cargar IMAGES de busqueda


    LimpiarCampos   'Limpia los campos TextBox
    
    'Vemos como esta guardado el valor del check
    'chkVistaPrevia.Value = CheckValueLeer(Name)
    For Modo = 0 To 5
        Me.imgBuscar(Modo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next Modo


      
    ' ****************** SI N'HI HAN COMBOS ********************************
    CargaCombo (0)
    ' **********************************************************************
    
    '****************** canviar la consulta *********************************+
    NomTabla = "empresas"
    Ordenacion = ""
    CadenaConsulta = "select * from " & NomTabla
    
    Me.adodc1.ConnectionString = conn
    Me.adodc1.RecordSource = CadenaConsulta
    Me.adodc1.Refresh
    
    CadB = ""
    If vEmpresa Is Nothing Then
        Me.SSTab1.TabVisible(3) = False
        Me.SSTab1.TabVisible(4) = False
    End If
    
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1
'        Text1(0).BackColor = vbYellow 'codclien
'    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        If Indice = 5 Then
            Text1(24).Text = RecuperaValor(CadenaDevuelta, 1)
        Else
            Text1(12 + Indice).Text = RecuperaValor(CadenaDevuelta, 1)
        End If
        Text2(Indice).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
End Sub








Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub

    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
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
            Case 12 To 16
                KEYBusqueda KeyAscii, Index - 12 'incidencia
            End Select
        End If
    Else
        KeyPress KeyAscii
    End If
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim ConValor As Boolean
Dim Aux As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
    Case 12 To 16, 24 'Incidencias

        PonerFormatoEntero Text1(Index)

        ConValor = False
        Aux = ""
        If Text1(Index).Text <> "" Then
            ConValor = True
            Aux = DevuelveDesdeBD("nominci", "incidencias", "idinci", Text1(Index).Text, "N")
            If Aux = "" Then MsgBox "Incidencia no existe", vbExclamation
        End If
        
        If Index = 24 Then
            Text2(5).Text = Aux
        Else
            Text2(Index - 12).Text = Aux
        End If
        If ConValor Then
            If Aux = "" Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
        End If
    Case 23
        Text1(Index).Text = Trim(Text1(Index).Text)
        If Text1(Index).Text <> "" Then
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
        End If
        
    Case 36
        Text1(Index).Text = Replace(Text1(Index).Text, ".", ":")
    End Select
    
End Sub


Private Sub Text3_GotFocus(Index As Integer)
    ConseguirFoco Text3(Index), 3
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
Dim cad As String
    Text3(Index).Text = Trim(Text3(Index).Text)
    cad = ""
    If IsNumeric(Text3(Index).Text) Then
        Text3(Index).Text = Val(Text3(Index).Text)
        cad = Text3(Index).Text
        cad = Round(Val(cad) / 60, 2)
        
    End If
    Text1(Index + 17) = cad
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
    If Me.adodc1.Recordset.EOF Then Exit Sub
    DesplazamientoData adodc1, Index
    PonerCampos
End Sub



Private Sub PonerCampos()
Dim F As Date
Dim N As Integer

    If adodc1.Recordset.EOF Then Exit Sub
    
    
    'Horario nocturno
    If Val(DBLet(adodc1.Recordset!HorarioNocturno, "N")) = 1 Then FrameHorasAcabalgadas.Visible = True
    
    PonerCamposForma Me, Me.adodc1
       
       
    'En el campo 17 y 18 tengo exceso deecto
    'Lo transformo a minutos
    If Text1(17).Text <> "" Then Text3(0).Text = Round(adodc1.Recordset!MaxRetraso * 60, 0)
    If Text1(18).Text <> "" Then Text3(1).Text = Round(adodc1.Recordset!MaxExceso * 60, 0)
    
       
       
    ' ************* configurar els camps de les descripcions *************
'    text2(6).Text = PonerNombreDeCod(Text1(6), "poblacio", "despobla", "codpobla", "N")

    SSTab1.TabVisible(3) = Check3.Value
    SSTab1.TabVisible(4) = Check2.Value
    
    Me.FrameHorasAcabalgadas.Visible = Me.Check1(5).Value = 1
    
    Modo = 3
    For Indice = 12 To 16
        Text1_LostFocus CInt(Indice)
    Next Indice
    Text1_LostFocus 24
    Modo = 2
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = PonerContRegistros(Me.adodc1)
End Sub



Private Function DatosOk() As Boolean
Dim B As Boolean
Dim d As Date

    If Text3(0).Text = "" Or Text3(1).Text = "" Then
        MsgBox "Campo retraso/Exceso en blanco", vbExclamation
        Exit Function
    End If
    d = "00:" & Format(Text3(0).Text, "00")
    Text1(17).Text = DevuelveValorHora(d)
    
    d = "00:" & Format(Text3(1).Text, "00")
    Text1(18).Text = DevuelveValorHora(d)
    
    B = CompForm(Me)
    If Not B Then Exit Function
    
    
    
    'Si el tipo de reloj es configreloj
    If Combo2.ListIndex = 5 Then
        'Obligado el path
        If Trim(Text1(25).Text) = "" Then
            MsgBox "Debe indicar la BD del Terminal FingerKey Access", vbExclamation
            B = False
        End If
    End If
    DatosOk = B
End Function


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
    imgBuscar_Click (Indice)
End Sub


' ********** SI N'HI HAN COMBOS *****************************


Private Sub CargaCombo(Index As Integer)
    
    
    
End Sub



Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda(Me, False)
    
    If CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
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
    'Me.Combo1(0).ListIndex = -1
    
    ' ****************************************************
    
    If Err.Number <> 0 Then Err.Clear
End Sub

' ***** SI N'HI HAN BOTONS I CAMPS DE BUSCAR EN ATRES FORMULARIS ********
Private Sub imgBuscar_Click(Index As Integer)
Dim cad As String

    TerminaBloquear

    Select Case Index
        Case 0 To 5
            'INCIDENCIAS
            cad = "Codigo|idinci|N||20·"
            cad = cad & "Descripcion|nominci|T||70·"
            Indice = Index
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vDevuelve = "0|1|"
            frmB.vTabla = "incidencias"
          
            frmB.vTitulo = "INCIDENCIAS"
            frmB.Show vbModal
            Set frmB = Nothing

        Case 1 'BANCO
'            Set frmBan = New frmBancsofi
'            frmBan.DatosADevolverBusqueda = "4|1|3|"
'            frmBan.CodigoActual = Text1(18).Text
'            If Me.Combo1(0).ListIndex > 0 Then
'                cad = Me.Combo1(0).ItemData(Combo1(0).ListIndex)
'            Else
'                cad = "724"
'            End If
'            frmBan.NuevoPais = cad
'            frmBan.Show vbModal
'            Set frmBan = Nothing
'            PonerFoco Text1(18)
    End Select

    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
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
End Sub

