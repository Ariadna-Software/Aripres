VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTCP3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interacción con el TCP-3"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmTCP3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "prueba para muchos ticajes"
      Height          =   315
      Left            =   240
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Operaciones"
      TabPicture(0)   =   "frmTCP3.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Configuracion"
      TabPicture(1)   =   "frmTCP3.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1"
      Tab(1).Control(1)=   "txtConfig(3)"
      Tab(1).Control(2)=   "Command5"
      Tab(1).Control(3)=   "txtConfig(1)"
      Tab(1).Control(4)=   "txtConfig(0)"
      Tab(1).Control(5)=   "lblConfig(3)"
      Tab(1).Control(6)=   "lblConfig(2)"
      Tab(1).Control(7)=   "lblConfig(1)"
      Tab(1).Control(8)=   "lblConfig(0)"
      Tab(1).ControlCount=   9
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -71700
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Frame Frame6 
         Height          =   5175
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton cmdMasAcciones 
            Caption         =   "Acciones"
            Height          =   375
            Left            =   5040
            TabIndex        =   36
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "ERRORES"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   975
            Left            =   120
            TabIndex        =   40
            Top             =   3360
            Width           =   6495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "ERRORES"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   1335
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   6495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ORDENADOR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   38
            Top             =   2640
            Width           =   1830
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "RELOJ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   900
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   6480
            Y1              =   2520
            Y2              =   2520
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Comprobar hora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   1635
         Left            =   360
         TabIndex        =   33
         Top             =   3720
         Width           =   2355
         Begin VB.CommandButton cmdCompruebaFecha 
            Caption         =   "Comprobar HORA"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.TextBox txtConfig 
         Height          =   315
         Index           =   3
         Left            =   -71700
         TabIndex        =   25
         Top             =   2940
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Poner en hora el reloj"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   1635
         Left            =   360
         TabIndex        =   18
         Top             =   1920
         Width           =   2355
         Begin VB.TextBox Text4 
            Height          =   315
            Index           =   1
            Left            =   1020
            TabIndex        =   23
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   540
            Width           =   615
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Programar fecha/hora"
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1140
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   22
            Top             =   300
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Hora"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Guadar configuración"
         Height          =   375
         Left            =   -71700
         TabIndex        =   17
         Top             =   3540
         Width           =   1935
      End
      Begin VB.TextBox txtConfig 
         Height          =   315
         Index           =   1
         Left            =   -71700
         TabIndex        =   14
         Top             =   1740
         Width           =   1935
      End
      Begin VB.TextBox txtConfig 
         Height          =   315
         Index           =   0
         Left            =   -71700
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3840
         Width           =   4035
      End
      Begin VB.Frame Frame3 
         Caption         =   "Incidencias Reloj"
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
         Height          =   1455
         Left            =   2820
         TabIndex        =   6
         Top             =   420
         Width           =   4095
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   480
            TabIndex        =   9
            Text            =   "Elija una incidencia"
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Programar Incid."
            Height          =   375
            Left            =   1080
            TabIndex        =   8
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   3060
            TabIndex        =   7
            Top             =   360
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   180
            Picture         =   "frmTCP3.frx":688A
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tecla"
            Height          =   195
            Left            =   3060
            TabIndex        =   10
            Top             =   180
            Width           =   555
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Empleados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1635
         Left            =   2820
         TabIndex        =   4
         Top             =   1920
         Width           =   4095
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   27
            Text            =   "Text5"
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox Text5 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   28
            Text            =   "Text5"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   345
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Text            =   "Elija un empleado"
            Top             =   1140
            Width           =   3435
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Programar tarjeta"
            Height          =   375
            Left            =   2220
            TabIndex        =   29
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   32
            Top             =   900
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "TARJETA"
            Height          =   195
            Index           =   3
            Left            =   1140
            TabIndex        =   31
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Codigo"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   840
            Picture         =   "frmTCP3.frx":698C
            Top             =   900
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Marcajes reloj"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1455
         Left            =   360
         TabIndex        =   2
         Top             =   420
         Width           =   2355
         Begin VB.CommandButton Command2 
            Caption         =   "Traer"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Label lblConfig 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo espera borrado"
         Height          =   195
         Index           =   3
         Left            =   -73740
         TabIndex        =   26
         Top             =   3000
         Width           =   1635
      End
      Begin VB.Label lblConfig 
         AutoSize        =   -1  'True
         Caption         =   "Baudios"
         Height          =   195
         Index           =   2
         Left            =   -73800
         TabIndex        =   16
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label lblConfig 
         AutoSize        =   -1  'True
         Caption         =   "Nº TCP-3"
         Height          =   195
         Index           =   1
         Left            =   -73800
         TabIndex        =   15
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblConfig 
         AutoSize        =   -1  'True
         Caption         =   "Puerto COMM"
         Height          =   195
         Index           =   0
         Left            =   -73800
         TabIndex        =   13
         Top             =   1140
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5940
      TabIndex        =   0
      Top             =   6000
      Width           =   1275
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2040
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RTSEnable       =   -1  'True
      BaudRate        =   19200
   End
End
Attribute VB_Name = "frmTCP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Comprobar As Byte
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Private vIndice As Integer 'Indice para el formuario de seleccionar
Dim T1, T2
Dim Buffer$
Private NF As Integer 'Fichero
Private NombreFichero As String
Private kTCP As String  'Por si tiene + de un terminal TCP

'Datos configuracion
Private PuertoComm As Byte
Private Baudios As Long
Private NumTCP3 As Byte
Private EsperaBorrado As Integer 'En segundos

Private TieneConfiguracion As Boolean


Private Sub cmdCompruebaFecha_Click()
Dim cad As String
Dim HayErrores As Boolean

   Screen.MousePointer = vbHourglass
    'Comprobaremos la fecha y hora del reloj
    'Si esta configurado y demas
    
   
    
    Frame6.Visible = True
    Label5.Caption = "  LEYENDO "
    'Fecha PC
    cad = UCase(Format(Now, "ddd d, mmm  hh:mm"))
    Label6.Caption = cad

    Me.Refresh
    'Llegados aqui es donde empezamos a transmitir con el reloj
    'Abrimos el puerto
    MSComm1.PortOpen = True
    Text1.Text = ""
    HayErrores = True
    
    'Solictamos comando
    Text1.Text = Text1.Text & "Solicitando programación hora/fecha reloj al TCP-3" & vbCrLf
    Text1.Refresh
    PonerTexto kTCP
    cad = Leer(2, "Cmd:")
    LimpiaBufferRecepcion
    'Si respuesta afirmativa
    If cad = "" Then
        GoTo Salida3
    End If
        
    'Ponemos comando 5 Leer hora/fecha en TCP-3
    cad = "5"
    PonerTexto cad
    cad = Leer(2, "Cmd OK")
    LimpiaBufferRecepcion
    'No hay datos correctos
    If cad = "" Then
        GoTo Salida3
    End If
    

    
    HayErrores = False
Salida3:
   
    
    If HayErrores Then
    
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
         Label5.Caption = "  E R R O R E S "
        End If
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    
    
    'En cad tenemos lo k ha llegador
    If Not HayErrores Then
        Label5.Caption = "Ajustando valores dev."
        Label5.Refresh
        If Not PonerHoras(cad) Then
            Label5.Caption = "Valores devueltos erroneos"
            Label5.Refresh
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdMasAcciones_Click()
    Text1.Text = ""
    Frame6.Visible = False
End Sub

'PROGRAMAR TARJETA
Private Sub Command1_Click()
Dim cad As String
Dim HayErrores As Boolean

'If Text2(0).Tag = "" Then
'    MsgBox "Seleccione un empleado.", vbExclamation
'    Exit Sub
'End If

If Text5(0).Text = "" Or Text5(1).Text = "" Or Text2(0).Text = "" Then
    MsgBox "Seleccione un empleado", vbExclamation
    Exit Sub
End If


Screen.MousePointer = vbHourglass

'Abrimos el puerto
MSComm1.PortOpen = True
Text1.Text = ""
HayErrores = True

'Solictamos comando
Text1.Text = Text1.Text & "Solicitando programacion tarjeta al TCP-3"
Text1.Refresh
PonerTexto kTCP
cad = Leer(2, "Cmd:")
LimpiaBufferRecepcion
'Si respuesta afirmativa
If cad = "" Then
    GoTo salida
End If
    
'Ponemos comando 18 de PRogramar TARJETA
PonerTexto "18"
cad = Leer(2, "Zona:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo salida
End If
'Ponemos zona
PonerTexto "1"
cad = Leer(2, "Planta:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo salida
End If
'Ponemos PLANTA
PonerTexto "18"
cad = Leer(2, "Acceso:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo salida
End If
'
Text1.Text = Text1.Text & vbCrLf & "Trabajador: " & Text2(vIndice).Text
Text1.Refresh
'
'Ponemos ACCESO
PonerTexto "255"
cad = Leer(2, "Fichaje:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo salida
End If
'Ponemos NUM tarjeta  ----------------
'que esta en el tag
'If Len(Text2(0).Tag) > 16 Then
'    Cad = Mid(Text2(0).Tag, 1, 16)
'    Else
'    Cad = Text2(0).Tag
'End If

'ANtes aripres 4
'Cad = Text2(0).Tag
cad = Text5(1).Text
PonerTexto cad
cad = Leer(2, "Identf.:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo salida
End If
'Ponemos Nombre  ----------------
'que esta en el tag
If Len(Text2(0).Text) > 16 Then
    cad = Mid(Text2(0).Text, 1, 16)
    Else
    cad = Text2(0).Text
End If
Text1.Text = Text1.Text & vbCrLf & "Enviando Datos tarjeta ....."
Text1.Refresh
PonerTexto cad
cad = Leer(15, "Cmd OK")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo salida
End If
'Todo correcto
Text1.Text = vbCrLf & vbCrLf & "Traspaso de datos realizado correctamente."
Text1.Refresh
HayErrores = False
salida:
    If HayErrores Then _
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Screen.MousePointer = vbDefault
End Sub


'IMPORTAR MARCAJES
Private Sub Command2_Click()
Dim cad As String
Dim Correcto As Boolean

Screen.MousePointer = vbHourglass

'Si ya existe el fichero marcaje
cad = Dir(NombreFichero)
If cad <> "" Then
    cad = "Ya existe el archivo: " & NombreFichero & vbCrLf
    cad = cad & "Procéselo primero"
    MsgBox cad, vbExclamation
    Screen.MousePointer = vbDefault
    Exit Sub
End If


'Como el TCP3 no envia el año, se lo pondremos nosotros
'en funcion de la fecha actual.
'Primero generamos el archivo Fichajes.txt y luego procesaremos
'Sus lineas. Es ahi donde Le añadimos el año.
'Problema:
'   Si hacemos el traspaso a primeros de enero y hay algun ticaje
'   de diciembre estos tendrán el año como el de enero.
'   Solucion dividir el FICHAJES.txt en dos.
'   La primera mitad contendra todos los ticajes de diciembre,
'   del año anterior, y la segunda los de enero hasta la fecha
'   actual.
'   para procesar los de diciembre cambiaremos la fecha de sistema
'   al año anterior consiguiendo de ese modo cambiar el año.
'   Una vez procesado la primera parte volveremos a la fecha actual
'   y procearemos la segunda.
If Month(Now) = 1 And Day(Now) < 4 Then
    cad = "Puede que hayan marcajes del año anterior " & Year(Now) - 1 & "." & vbCrLf
    cad = cad & "Si hay alguno tendra que dividir el fichero fichajes.txt" & vbCrLf
    cad = cad & "Del siguiente modo." & vbCrLf
    cad = cad & "1.- Haga una copia del FICHAJES.TXT a FICHAJES.BAK" & vbCrLf
    cad = cad & "2.- Abra el fichero y seleccione las entradas de enero y borrelas. " & vbCrLf
    cad = cad & "3.- Cambie la fecha de su ordenador y ponga 31 de diciembre de " & Year(Now) - 1 & vbCrLf
    cad = cad & "4.- Procese el fichero" & vbCrLf
    cad = cad & "5.- Copie FICHAJES.BAK como Fichajes.txt" & vbCrLf
    cad = cad & "6.- Vuelva a poner la fecha actual" & vbCrLf
    cad = cad & "7.- Procese otra vez el fichero" & vbCrLf
    MsgBox cad, vbExclamation, "PRECAUCION"
End If
'Eliminamos los temporales
cad = Dir(App.Path & "\tmpMar.txt")
If cad <> "" Then
    Kill App.Path & "\tmpMar.txt"
End If
cad = Dir(App.Path & "\tmpMar2.txt")
If cad <> "" Then
    Kill App.Path & "\tmpMar2.txt"
End If

'Abrimos el puerto
MSComm1.PortOpen = True
Text1.Text = ""

'Probablemente haya que bloquear la unidad durante este proceso


'Solictamos comando
Text1.Text = Text1.Text & "Solicitando marcajes al TCP-3" & vbCrLf
Text1.Refresh
PonerTexto kTCP
cad = Leer(2, "Cmd:")
'Debug.Print cad
LimpiaBufferRecepcion

'Si respuesta afirmativa
If cad = "" Then
    Text1.Text = Text1.Text & "Error solicitando datos"
    GoTo salida
End If
    
'Ponemos comando 3 de solicitud de transferencias
PonerTexto "3"
cad = Leer(2, "Reg. Ini:")

LimpiaBufferRecepcion
'Si respuesta afirmativa
If cad = "" Then
    GoTo salida
End If

'Registro inicial 0
PonerTexto "0"
cad = Leer(2, "Reg. Fin:")
LimpiaBufferRecepcion

'Si respuesta afirmativa
If cad = "" Then
    GoTo salida
End If

'Registro Final 0
Text1.Text = Text1.Text & vbCrLf & "Devolviendo registros..."
Text1.Refresh
PonerTexto "0"

'Ahora leeremos todos los datos que nos devuelve la maquina
'sobre los tikajes
Correcto = False
If AbrirFichero = 0 Then
    Correcto = LeerDatos2(5, True) 'Leemos los datos
    Correcto = Correcto And (CerrarFichero = 0)
End If

If Not Correcto Then GoTo salida
Text1.Text = Text1.Text & vbCrLf & "Marcajes recibidos. Procesando fichero temporal."
Text1.Refresh
'Ahora procesamos el fichero
Correcto = ProcesarFichero2
If Not Correcto Then GoTo salida
    
If Correcto Then
    Text1.Text = Text1.Text & "    TODO CORRECTO  "
    Else
        Text1.Text = Text1.Text & "    Error procesando el fichero."
        Text1.Refresh
End If

If Not Correcto Then
    'Borramos marcajes.txt
    Text1.Text = Text1.Text & vbCrLf & "Error procesando el fichero temporal"
    Text1.Refresh
    Kill NombreFichero
    GoTo salida
End If


'BORRAR MARCAJES
'Todo ha ido bien, luego borramos los marcajes de la maquina
'Solictamos comando de borrar
Text1.Text = "Solicitando borrado marcajes."
Text1.Refresh
LimpiaBufferRecepcion
PonerTexto kTCP
cad = Leer(6, "Cmd:")
'Si respuesta afirmativa
If cad = "" Then
    Text1.Text = Text1.Text & vbCrLf & "Error conectando con el TCP-3(Cmd: 4). Tiempo agotado"
    Text1.Refresh
    Kill NombreFichero
    GoTo salida
End If
LimpiaBufferRecepcion

'Ponemos comando 3 de solicitud de transferencias
PonerTexto "4"
cad = Leer(2, "Confirmar:")
LimpiaBufferRecepcion
'Si respuesta afirmativa
If cad = "" Then
    Text1.Text = Text1.Text & vbCrLf & "Error conectando con el TCP-3(Confirmar: Si). Tiempo agotado"
    Text1.Refresh
    Kill NombreFichero
    GoTo salida
End If

PonerTexto "si"
'Aqui ponemos mas tiempo por que tiene que borrar registros: EsperaBorrado
cad = Leer(EsperaBorrado, "Cmd OK")
'Debug.Print cad
LimpiaBufferRecepcion

'Si no ha podido borrar entonces borramos el marcaje
If cad = "" Then
    Text1.Text = Text1.Text & vbCrLf & "Ha sido imposible eliminar los marcajes en el TCP-3"
    Text1.Refresh
    Kill NombreFichero
    GoTo salida
End If
Text1.Text = Text1.Text & vbCrLf & "Marcajes eliminados en el TCP-3"
Text1.Refresh
'LLegados a este punto cambiaremos el valor de la variable
'MostrarErrores para que el formulario ppal sepa que
'tiene datos para importar
'FALTA####   tratar esta variable
'MostrarErrores = True

'Todo Correcto
Text1.Text = vbCrLf & vbCrLf & "Traspaso realizado con éxito."
Text1.Refresh

salida:
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Screen.MousePointer = vbDefault
    If Correcto Then
        Me.Hide
        Unload Me
        DoEvents
        frmTraspaso.Opcion = 1
        frmTraspaso.Show vbModal
    End If
End Sub


Private Function Leer(segundos As Integer, CadenaEsperada As String) As String
Dim i As Integer
Dim Fin As Boolean
Dim T1, T2
Dim Buffer2$

Leer = ""
Fin = False
T1 = Timer
'Bucle
While Not Fin
    Buffer2$ = Buffer2$ & MSComm1.Input
    'Debug.Print "R: " & Buffer2$
    If InStr(1, Buffer2$, CadenaEsperada) Then
        Fin = True
        Leer = Buffer2$
        Else
            T2 = Timer - T1
            Fin = (T2 > segundos)
    End If
Wend
End Function



Private Sub PonerTexto(T As String)
MSComm1.Output = T & vbCrLf
End Sub


Private Sub espera(tiempo As Single)
T1 = Timer
Do
    T2 = Timer - T1
Loop Until T2 > tiempo
End Sub


'Private Function LeerDatos(Segundos As Integer) As Boolean
'Dim i As Integer
'Dim Fin As Boolean
'Dim tmp As String
'Dim Esta As Boolean
'
'Buffer$ = ""
'tmp = ""
'Fin = False
't1 = Timer
''Bucle
'While Not Fin
'
'    i = 0
'    Do
'        Buffer$ = Buffer$ & MSComm1.Input
'        t2 = Timer
'        i = InStr(1, Buffer$, vbCrLf)
'    Loop Until (t2 - t1) > Segundos Or i > 0
'    Text1.Text = Text1.Text & Buffer$ & vbCrLf
'    Text1.Refresh
'    If i > 0 Then
'        EscribeTextoFichero Buffer$
'        If InStr(1, Buffer$, "Cmd OK") Then
'            Fin = True
'            Esta = True
'            Else
'                Buffer$ = ""
'                t1 = Timer
'        End If
'        Else
'            'i=0
'            'fin tiempo
'            Esta = False
'            Fin = True
'    End If
'Wend
'LeerDatos = Esta
'End Function

Private Sub EscribeTextoFichero(Texto As String)
'ProcesaTexto texto
Print #NF, Texto
End Sub

Private Function AbrirFichero() As Byte
On Error GoTo ErrorAbrir
'Abrimos el fichero para escritura
AbrirFichero = 1
NF = FreeFile
        Open App.Path & "\tmpMar.txt" For Output As #NF
AbrirFichero = 0
Exit Function
ErrorAbrir:
    MsgBox "Error abriendo fichero: " & vbCrLf & "Número: " & Err.Number & _
        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
End Function

Private Function CerrarFichero() As Byte
On Error GoTo ErrorCerrar
'Abrimos el fichero para escritura
CerrarFichero = 1
Close NF
CerrarFichero = 0
Exit Function
ErrorCerrar:
    MsgBox "Error cerrando fichero: " & vbCrLf & "Número: " & Err.Number & _
        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
End Function

Private Sub LimpiaBufferRecepcion()
Dim Cade
Cade = MSComm1.Input
espera 0.25
End Sub


'Private Function ProcesarFichero() As Boolean
'Dim Linea As String
'Dim i As Long
'Dim TotalReg As Long
'Dim Fich As Integer
'
'
'Text1.Text = ""
'On Error GoTo ErrProcFich
'Fich = FreeFile
'NF = FreeFile + 1
'Open App.Path & "\tmpMar.txt" For Input As #NF
'Open NombreFichero For Output As #Fich
'i = 0 'Tendremos el contador
'While Not EOF(NF)
'    Line Input #NF, Linea
'    'Ahora procesamos la linea
'    Linea = Trim(Linea)
'    If Len(Linea) = 1 Then Linea = ""
'    'Segun este en blanco ponga unas cosas etcc..
'    If Linea <> "" Then
'        If Mid(Linea, 1, 10) = "Total Reg." Then
'            TotalReg = CLng(Val(Mid(Linea, 11)))
'            Else
'                If Mid(Linea, 1, 4) <> "Reg." Then
'                    If Mid(Linea, 1, 3) <> "Cmd" Then
'                        Print #Fich, Linea
'                        Text1.Text = Text1.Text & Linea
'                        i = i + 1
'                    End If
'                End If
'        End If
'    End If
'Wend
'Close #NF
'Close #Fich
''Vemos si esta correcto o no
'ProcesarFichero = (i = TotalReg)
'Exit Function
'ErrProcFich:
'    ProcesarFichero = False
'    MsgBox "Error procesando fichero: " & vbCrLf & "Número: " & Err.Number & _
'        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
'End Function



'INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS
'INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS
Private Sub Command3_Click()
Dim tecla As Integer
Dim HayErrores As Boolean
Dim cad As String


If Text2(1).Tag = "" And Text2(0).Text = "Elija una incidencia" Then
    MsgBox "Seleccione una incidencia.", vbExclamation
    Exit Sub
End If

If Text3.Text = "" Then
    MsgBox "Seleccione una incidencia y una tecla para asociarla.", vbExclamation
    Exit Sub
End If

If Not IsNumeric(Text3.Text) Then
    MsgBox "El codigo de tecla de incidencia tiene que ser numerico.", vbExclamation
    Exit Sub
End If

tecla = CInt(Text3.Text)
If tecla < 1 Then
    MsgBox "El código de tecla tiene que ser mayor que cero.", vbExclamation
    Exit Sub
End If
If tecla > 9 Then
    MsgBox "El código de tecla tiene que ser menor que 10 ( Del 0 al 9).", vbExclamation
    Exit Sub
End If

Screen.MousePointer = vbHourglass
'Llegados aqui es donde empezamos a transmitir con el reloj
'Abrimos el puerto
MSComm1.PortOpen = True
Text1.Text = ""
HayErrores = True
'Solictamos comando
Text1.Text = Text1.Text & "Solicitando grabación de mensajes al TCP-3" & vbCrLf
Text1.Refresh
PonerTexto kTCP
cad = Leer(2, "Cmd:")
LimpiaBufferRecepcion
'Si respuesta afirmativa
If cad = "" Then
    GoTo Salida2
End If
    
'Ponemos comando 13 grabar mensajes usuario
PonerTexto "13"
cad = Leer(2, "Usuario")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida2
End If
'Ponemos la tecla asignada
PonerTexto CStr(tecla)
cad = Leer(2, "Ind:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida2
End If
'Ponemos PLANTA
If Len(Text2(1).Text) < 16 Then
    cad = Text2(1).Text & "                           "
    Else
        cad = Mid(Text2(1).Text, 1, 16)
End If
cad = Mid(cad, 1, 16)
Text1.Text = Text1.Text & " Incidencia: " & tecla & vbCrLf
Text1.Text = Text1.Text & " Mensaje: " & cad & vbCrLf
Text1.Refresh
PonerTexto cad

cad = Leer(6, "Cmd OK")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida2
End If
Text1.Text = Text1.Text & vbCrLf & "Grabación correcta" & vbCrLf
Text1.Refresh

HayErrores = False
Salida2:
    If HayErrores Then _
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Screen.MousePointer = vbDefault



End Sub

Private Sub Command4_Click()
If MSComm1.PortOpen Then MSComm1.PortOpen = False
Unload Me
End Sub


'-------------------------------------------------
'-------------------------------------------------
'  Configuracion terminal
'  ----------------------
Private Sub Command5_Click()
Dim i As Integer

On Error GoTo ErrorGuardar
    For i = 0 To 3
        If i <> 2 Then
            If Not IsNumeric(txtConfig(i).Text) Then
                MsgBox "Los campos de configuracion tienen que ser numéricos.", vbExclamation
                Exit Sub
            End If
        End If
    Next i

    
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccine la velocidad de transmision", vbExclamation
        Exit Sub
    End If

    i = FreeFile
    Open App.Path & "\TCPConf.cfg" For Output As #i
    Print #i, Trim(txtConfig(0)) & "|" & Trim(txtConfig(1)) & "|" & Combo1.List(Combo1.ListIndex) & "|" & Trim(txtConfig(3)) & "|"
    Close #i
    

    
ErrorGuardar:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error." & vbCrLf & _
            " Número: " & Err.Number & vbCrLf & _
            " Descripción: " & Err.Description, vbExclamation
            
        Else
            'Ponemos los valores al COMM
'            MSComm1.CommPort = txtConfig(0).Text
'            MSComm1.Settings = Combo1.Text & ",N,8,1"
'            kTCP = "tcp" & Trim(txtConfig(1).Text)
'            EsperaBorrado = CInt(txtConfig(3))
            TieneConfiguracion = True
            OpcionesSegunEsteConfigurado
    End If
End Sub

'-------------------------------------------------
'-------------------------------------------------
'  Poner en hora el reloj
'  ----------------------
' ¤  este es el caracter que recibe por la ñ
' puede, no esta comprobado, que de errores en un futuro
Private Sub Command6_Click()
Dim HayErrores As Boolean
Dim cad As String


If Text4(0).Text = "" Then
    MsgBox "Ponga una hora para transmitir al TCP3", vbExclamation
    Exit Sub
End If

If Text4(1).Text = "" Then
    MsgBox "Ponga una fecha para transmitir al TCP3", vbExclamation
    Exit Sub
End If

If Not IsDate(Text4(0).Text) Then
    MsgBox "Ponga una hora correcta para transmitir al TCP3", vbExclamation
    Exit Sub
End If

If Not IsDate(Text4(1).Text) Then
    MsgBox "Ponga una fecha correcta para transmitir al TCP3", vbExclamation
    Exit Sub
End If


Screen.MousePointer = vbHourglass
'Llegados aqui es donde empezamos a transmitir con el reloj
'Abrimos el puerto
MSComm1.PortOpen = True
Text1.Text = ""
HayErrores = True

'Solictamos comando
Text1.Text = Text1.Text & "Solicitando programación hora/fecha reloj al TCP-3" & vbCrLf
Text1.Refresh
PonerTexto kTCP
cad = Leer(2, "Cmd:")
LimpiaBufferRecepcion
'Si respuesta afirmativa
If cad = "" Then
    GoTo Salida3
End If
    
'Ponemos comando 2 programar hora/fecha en TCP-3
'El 1 tb programa pero el 2 en nuestro terminal es mas completo
PonerTexto "2"
cad = Leer(2, "a¤o:")  'realmente ¤ es la ñ
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida3
End If

'Ponemos el año
cad = Year(CDate(Text4(1).Text))
Text1.Text = Text1.Text & "Año: " & cad & vbCrLf
Text1.Refresh
PonerTexto cad
cad = Leer(2, "mes:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida3
End If

'Ponemos el mes
cad = Month(CDate(Text4(1).Text))
Text1.Text = Text1.Text & "Mes: " & cad & vbCrLf
Text1.Refresh
PonerTexto cad
cad = Leer(2, "dia:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida3
End If
'Ponemos el dia
cad = Day(CDate(Text4(1).Text))
Text1.Text = Text1.Text & "Dia: " & cad & vbCrLf
Text1.Refresh
PonerTexto cad
cad = Leer(2, "ds:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida3
End If
'Ponemos el dia semana
cad = Weekday(CDate(Text4(1).Text), vbMonday)
Text1.Text = Text1.Text & "Dia seman: " & cad & "  (1 - Lunes ...)" & vbCrLf
Text1.Refresh
PonerTexto Format(cad, "00")
cad = Leer(2, "hora:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida3
End If
'Ponemos la hora
cad = Hour(CDate(Text4(0).Text))
Text1.Text = Text1.Text & "Hora: " & cad & vbCrLf
Text1.Refresh
PonerTexto cad
cad = Leer(2, "min:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida3
End If
'Los minutos
cad = Minute(CDate(Text4(0).Text))
Text1.Text = Text1.Text & "Hora: " & cad & vbCrLf
Text1.Refresh
PonerTexto cad
cad = Leer(2, "seg:")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida3
End If
'Los segundos
cad = 30
Text1.Text = Text1.Text & "Segundos: " & cad & vbCrLf
Text1.Refresh
PonerTexto cad
cad = Leer(2, "Cmd OK")
LimpiaBufferRecepcion
'No hay datos correctos
If cad = "" Then
    GoTo Salida3
End If

'llegados  a este punto todo ha ido bien
Text1.Text = Text1.Text & vbCrLf & vbCrLf
Text1.Text = Text1.Text & vbCrLf & "Comando completado con éxito."
Text1.Refresh

'
HayErrores = False
Salida3:
    If HayErrores Then _
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Screen.MousePointer = vbDefault







End Sub

Private Sub Command7_Click()
ProcesarFichero2
End Sub


Private Sub Form_Activate()
    If Me.Tag <> "" Then
        'La primera vez
        Me.Tag = ""
        If Comprobar Then cmdCompruebaFecha_Click
    End If
End Sub

Private Sub Form_Load()
    

    Me.Tag = "OK"
    
    CargaCombo
    
    'Si solo es comprobar la fecha solo le dejo salir
    Me.cmdMasAcciones.Visible = Not Comprobar
    Text2(0).Tag = ""
    Text2(1).Tag = ""
    Text5(0).Text = ""
    Text5(1).Text = ""
    Frame6.Visible = False
    ObtenerConfiguracion
    OpcionesSegunEsteConfigurado
    'NombreFichero = mConfig.DirMarcajes & "\" & mConfig.NomFich
    NombreFichero = App.Path & "\Fichajes.txt"
    SSTab1.Tab = 0
    Text4(0).Text = Format(Now, "hh:mm")
    Text4(1).Text = Format(Now, "dd/mm/yyyy")
    Screen.MousePointer = vbDefault
End Sub


Private Sub OpcionesSegunEsteConfigurado()
    
    If TieneConfiguracion Then
        MSComm1.CommPort = txtConfig(0).Text
        MSComm1.Settings = Combo1.Text & ",N,8,1"
        kTCP = "tcp" & Trim(txtConfig(1).Text)
        EsperaBorrado = CInt(txtConfig(3))
    End If
    Me.Command1.Enabled = TieneConfiguracion
    Me.Command2.Enabled = TieneConfiguracion
    Me.Command3.Enabled = TieneConfiguracion
    Me.Command6.Enabled = TieneConfiguracion
    Me.cmdCompruebaFecha.Enabled = TieneConfiguracion
End Sub



Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    Dim cad As String
    If vIndice = 0 Then
        'Cad = devuelveNombreNTarjeta(CInt(vCodigo))
        If cad <> "" Then
            'Necesito saber tambien el número de la tarjeta
            Text2(0).Tag = cad
            Text2(0).Text = vCadena
            Text5(0).Text = vCodigo
            Text5(1).Text = cad
        End If
        Else
            Text2(1).Tag = vCodigo
            Text3.Text = vCodigo
            Text2(1).Text = vCadena
    End If
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    If NF = 0 Then
        'Trabajadores
        Text5(0).Text = RecuperaValor(CadenaDevuelta, 1)
        Text5(1).Text = RecuperaValor(CadenaDevuelta, 3)
        Text2(0).Text = RecuperaValor(CadenaDevuelta, 2)
        
    Else
        'Incidencias
        Text2(1).Text = RecuperaValor(CadenaDevuelta, 2)
        Text3.Text = RecuperaValor(CadenaDevuelta, 1)
    End If
End Sub

Private Sub Image1_Click(Index As Integer)
Dim cad As String

    Set frmB = New frmBuscaGrid
    NF = Index

    
    'Ponemos los valores para abrir
    If Index = 0 Then
        vIndice = 0
        cad = "Codigo|idTrabajador|N||15·"
        cad = cad & "Nombre|nomtrabajador|T||60·"
        cad = cad & "Tarjeta|numtarjeta|T||20·"
        frmB.vTabla = "Trabajadores"
        frmB.vCampos = cad
        frmB.vDevuelve = "0|1|2|"
        frmB.vSelElem = 0
        frmB.vTitulo = "TRABAJADORES"
        Else
            cad = "Codigo|idinci|N||20·"
            cad = cad & "Descripcion|nominci|T||70·"
            frmB.vCampos = cad
            frmB.vDevuelve = "0|1|"
            frmB.vTabla = "incidencias"
            frmB.vTitulo = "INCIDENCIAS"
    End If
    frmB.Show vbModal
    Set frmB = Nothing
End Sub


Private Sub ObtenerConfiguracion()
Dim cad As String
Dim NF As Integer
Dim i As Integer
Dim ini As Integer
Dim v(3) As String
Dim C As String
Dim L As Long

On Error GoTo ErrorObtener
TieneConfiguracion = False

cad = Dir(App.Path & "\TCPConf.cfg")
If cad <> "" Then
    
    NF = FreeFile
    Open App.Path & "\TCPConf.cfg" For Input As #NF
    Input #NF, cad
    Close #NF
    'Leemos los valores
    TieneConfiguracion = True
    v(0) = "1"
    v(1) = "1"
    v(2) = "19200"
    v(3) = "40"
    
    
    ini = 1
    For NF = 0 To 3
        i = InStr(ini, cad, "|")
        If i > 0 Then
            C = Mid(cad, ini, i - ini)
            If IsNumeric(C) Then v(NF) = C
            ini = i + 1
        End If
    Next NF
Else
    Exit Sub
End If
'Valores por defecto
'NO existe el fichero. Ponemos los valores por defecto
ErrorObtener:
    L = CLng(v(0))
    If L > 255 Then L = 1
    PuertoComm = CByte(L)
    '-------------------
    L = CLng(v(1))
    If L > 255 Then L = 1
    NumTCP3 = CByte(L)
    '---------------------
    Baudios = CLng(v(2))
    '---------------------
    EsperaBorrado = CInt(v(3))
    For i = 0 To 3
        If i <> 2 Then txtConfig(i) = v(i)
    Next i
    
    'Pongo el combo
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = Baudios Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next i
    If i > Combo1.ListIndex Then
        MsgBox "Error fijando la velocidad de conexion", vbExclamation
        TieneConfiguracion = False
    End If
    
End Sub




'Segunda prueba para leer mucho datos desde el TCP 3
Private Function LeerDatos2(segundos As Integer, Leyendo As Boolean) As Boolean
Dim Seg As Integer
Dim i As Integer
Dim Fin As Boolean
Dim tmp As String
Dim Esta As Boolean
Dim J As Integer
Dim T3, T4
'De esta forma haremos:
'iremos leyendo del buffer
' Si hay texto cad segundos/2 (aproximadamente)
'lo guardamos y
'restauramos tiempo
'Si no hay tiempo salimos sin hacer nada
Buffer$ = ""
Fin = False
T1 = Timer
J = segundos + 1
'Bucle
Text1.Text = Text1.Text & vbCrLf
If Leyendo Then
    Text1.Text = Text1.Text & "Leyendo registros de marcajes. " & vbCrLf
    Else
    Text1.Text = Text1.Text & "Eliminando registros de marcajes. " & vbCrLf
End If
Text1.Refresh
tmp = Text1.Text
i = 0
T3 = Timer
While Not Fin
    
    
    Do
        Buffer$ = Buffer$ & MSComm1.Input
        T2 = Timer - T1
        If Buffer$ <> "" Then J = 2
    Loop Until T2 > segundos Or T2 > J
    T4 = Timer - T3
    If Buffer$ <> "" Then
        Text1.Text = tmp & vbCrLf & "Bloque: " & Format(i, "0000") & "     " & "  Seg: " & Format(T4, "0.00")
        Text1.Refresh
        i = i + 1
        EscribeTextoFichero Buffer$
        If InStr(1, Buffer$, "Cmd OK") Then
            Fin = True
            Esta = True
            Else
                Buffer$ = ""
                J = segundos + 1
                T1 = Timer
        End If
        Else
            'i=0
            'fin tiempo
            Esta = False
            Fin = True
    End If
Wend
LeerDatos2 = Esta
End Function






Private Function ProcesarFichero2() As Boolean
Dim Linea As String
Dim Aux As String
Dim i As Long
Dim TotalReg As Long
Dim Fich As Integer
Dim Fin As Boolean
Dim J As Integer
Dim cad As String

'Del fichero leido lo ponemos en lineas que contienen unicamente numero
' lo leido por el fichero
' Lineas del tipo linea i  : Reg. i
'                 linea i+1: 09001,10,12,10,53,0000,0000,02302
'
On Error GoTo ErrProcFich
Fich = FreeFile
NF = FreeFile + 1

'GoTo Aqui
Open App.Path & "\tmpMar.txt" For Input As #NF
Open App.Path & "\tmpMar2.txt" For Output As #Fich

If EOF(NF) Then
    MsgBox "El archivo esta vacío.", vbExclamation
    Close #NF
    Close #Fich
    ProcesarFichero2 = False
    Exit Function
End If
    Fin = False

    
'Leemos la primera linea k despreciamos
Line Input #NF, Linea
'Segunda linea
Line Input #NF, Linea
'El resto del fichero
While Not EOF(NF)

    Line Input #NF, cad
    J = InStr(1, cad, "Total Reg")
    If J > 0 Then
        'Mandamos a imprimir lo que esta antes de totalreg
        Aux = Linea & Mid(cad, 1, J - 1)
        'Escribimos en el fichero Aux
        Print #Fich, Aux
        'Recalculamos linea
        Linea = Mid(cad, J)
        Else
            i = InStr(1, cad, "Reg.")
            If i > 0 Then
                Aux = Linea & Mid(cad, 1, i - 1)
                'Escribimos en el fichero Aux
                Print #Fich, Aux
                'Recalculamos linea
                Linea = Mid(cad, i)
                'Sino
                Else
                    Linea = Linea & cad
            End If
    End If
Wend


'Lo ultimo del texto lo imprimimos
If Linea <> "" Then Print #Fich, Linea

'Cerramos los ficheros
Close #NF
Close #Fich

'Aqui:

'Ya hemos tenemos en tmpMarcajes el fichero PRE-procesado
'ahora lo terminamos de procesar
Fich = FreeFile
NF = FreeFile + 1
Open App.Path & "\tmpMar2.txt" For Input As #NF
Open NombreFichero For Output As #Fich

i = 0 'Tendremos el contador.
Fin = False
Linea = ""
While Not Fin
    Line Input #NF, cad
    J = InStr(1, cad, "Total Reg")
        If J > 0 Then
            'Es la ultima linea que nos intersa
            TotalReg = CLng(Val(Mid(cad, 11)))
            Fin = True
            Linea = ""
            Else
                'Ahora hay que procear cada linea
                Linea = ProcesarLineaMarcaje(cad)
                Print #Fich, Linea
                i = i + 1
        End If 'De j>0
    Fin = Fin Or EOF(NF)
Wend

'Cerramos los ficheros
Close #NF
Close #Fich

'Vemos si esta correcto o no
ProcesarFichero2 = (i = TotalReg)
Exit Function
ErrProcFich:
    ProcesarFichero2 = False
    MsgBox "Error procesando fichero: " & vbCrLf & "Número: " & Err.Number & _
        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
End Function


Private Function ProcesarLineaMarcaje(Linea As String) As String
Dim Aux As String
Dim i As Integer

On Error GoTo errorProcesarLineaMarcaje

ProcesarLineaMarcaje = ""
'Separamos las dos partes de la linea
i = InStr(2, Linea, vbLf)


Aux = Mid(Linea, i + 1)
Aux = Mid(Aux, 1, Len(Aux) - 1)
ProcesarLineaMarcaje = Aux
Exit Function
errorProcesarLineaMarcaje:
    Aux = "Una linea del fichero ha llegado con error." & vbCrLf
    Aux = Aux & "--> " & Linea
    Aux = Aux & vbCrLf & vbCrLf & "Revise luego a mano el fichero FICHAJES.TXT"
    MsgBox Aux, vbCritical
    ProcesarLineaMarcaje = Linea
End Function


Private Sub Label1_DblClick()
    Text2(1).Enabled = Not Text2(1).Enabled
End Sub

Private Sub PonerEmpleadoVacio()
            Text5(0).Text = ""
            Text5(1).Text = ""
            Text2(0).Text = ""
            Text2(0).Tag = ""
End Sub
Private Sub PonerEmpleado(Cod As String, Campo As String)
Dim RT As ADODB.Recordset
Dim Sql As String
    
    Sql = "Select * from Trabajadores where "
    Sql = Sql & Campo & " = " & Cod
    Set RT = New ADODB.Recordset
    RT.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RT.EOF Then
        'ponerempleadovacio
        PonerEmpleadoVacio
    Else
        'Ponemos los datos del empleado
        If IsNull(RT!Numtarjeta) Then
            MsgBox "No tiene codigo tarjeta asociado", vbExclamation
            PonerEmpleadoVacio
        Else
            Text5(0).Text = RT!idTrabajador
            Text5(1).Text = RT!Numtarjeta
            Text2(0).Text = RT!nomtrabajador
            Text2(0).Tag = RT!Numtarjeta
        End If
    End If
    RT.Close
    Set RT = Nothing
End Sub



Private Sub Text5_GotFocus(Index As Integer)
    With Text5(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text5_LostFocus(Index As Integer)
    Text5(Index).Text = Trim(Text5(Index).Text)
    If Text5(Index).Text <> "" Then
        If Not IsNumeric(Text5(Index).Text) Then
            MsgBox "Codigo incorrecto: " & Text5(Index).Text, vbExclamation
            Text5(Index).Text = ""
        End If
    End If
    If Text5(Index).Text = "" Then
        PonerEmpleadoVacio
    Else
        If Index = 0 Then
            PonerEmpleado Text5(Index).Text, "idTrabajador"
        Else
            PonerEmpleado "'" & Text5(Index).Text & "'", "NumTarjeta"
        End If
    End If
End Sub


'Pone la cadena devuelta por el reloj  y la fecha /hora del PC
Private Function PonerHoras(CADENA As String) As Boolean
Dim i As Integer
Dim Fecha As Date

    PonerHoras = False
    Buffer$ = ""
    '5 Reg. Fin:
    'mes,dia,ds,hora,min,seg:08,31,5,21,38,21
    ' a partir de los dos puntos
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        
        'BIEN
        CADENA = Mid(CADENA, i + 1)
        'Le aañdimos una coma al final para facilitar
        
        
    
            NF = InStr(1, CADENA, Chr(13))
            If NF > 0 Then CADENA = Mid(CADENA, 1, NF - 2)
            CADENA = CADENA & ","
        
        
        NF = 0
        Do
            i = InStr(1, CADENA, ",")
            If i > 0 Then
                NF = NF + 1
                Buffer$ = Mid(CADENA, i + 1)
                CADENA = Mid(CADENA, 1, i - 1) & "|" & Buffer$
            End If
        Loop Until i = 0
        
        Buffer$ = ""
        If NF = 6 Then

            'FECHA OK
            '----
            'Dia semana lo calcularemos de la primerasemana del
            'mes de noviembre de 2004 que empieza ekl 1 Lunes
            i = Val(RecuperaValor(CADENA, 3))
            If i > 7 Then Exit Function
            'I = I + 1
            Buffer$ = Buffer$ & Format(i & "/11/2004", "ddd")
            
            
            i = Val(RecuperaValor(CADENA, 2))
            If i = 0 Then Exit Function
            Buffer$ = Buffer$ & " " & i & ","
            
            i = Val(RecuperaValor(CADENA, 1))
            If i = 0 Then Exit Function
            Buffer$ = Buffer$ & "  " & Format("01/" & i & "/2004", "mmm")
            
            'Hora
            i = Val(RecuperaValor(CADENA, 4))
            Buffer$ = Buffer$ & "  " & i & ":"
            
            'Minutos
            i = Val(RecuperaValor(CADENA, 5))
            Buffer$ = Buffer$ & Format(i, "00")
            
            
            Label5.Caption = UCase(Buffer$)
            PonerHoras = True
        End If
        
    End If
    
    Buffer$ = ""
End Function

Private Sub CargaCombo()
    Combo1.Clear
    Combo1.AddItem "2400"
    Combo1.AddItem "9600"
    Combo1.AddItem "14400"
    Combo1.AddItem "19200"
    
    Combo1.AddItem "28800"
    Combo1.AddItem "38400"
    Combo1.AddItem "56000"
    Combo1.AddItem "128000"
End Sub
