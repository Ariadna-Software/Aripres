VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión listados"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfigImpre 
      Caption         =   "Sel. &impresora"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6435
      Begin VB.CheckBox chkEMAIL 
         Caption         =   "Enviar e-mail"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox chkSoloImprimir 
         Caption         =   "Previsualizar"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Sin definir"
      Top             =   180
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   5535
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Integer
            '0  .-  Listado marcajes por fecha   codigo
            '1  .-  Listado marcajes por fecha nombre

            '2  .-  Listado marcajes por trabajador  codigo
            '3  .-  Listado marcajes por trabajador  nombre
    
    
            '8  .- Presencia real NOMBRE
            '9  .-  "          "  codigo
            
            'Trabajadores
            '10 .- sin seccion sin foto   basico        codigo
            '11                                         nombre
            '12 .-  "             "       extendido     codigo
            '13                                         nombre
            
            '14 .- sin seccion con foto    basico       codigo
            
            
            
            '18 .- SECCION SIN FOTO, basico codi
            '19                       "     nombre
            '20                       exten  cod
            '21                       ext   nombre
            
            '30 .- Horarios
            '31 .- Secciones
            '32 .- Marcaje actual
            
            
            '33 .- Incidencia resumen:  codigo
            '34 .-  "                   nombre
            
            '35 .- Combinado por     nombre
            '36 .- Combinado         codigo
            '37 .- "           FECHA nombre
            '38 .- "              "  codigo
            
            '40 .- DIAS trabajaods
            
            
            
            '50 .- Incidencias generadas: codigo
            '51 .-     "          "       nombre
            '52 .-   "              "  agrupadas por trabajador   codigo
            '53 .-   "              "  agrupadas por trabajador   nombre
    
            '54 .- INCIRESUMEN agrupada por trabajador codi
            '55 .-   "           "             "         nom



            '60 .- Marcaje actual agrupado por trabajador

            '61 .- Costes
            '      Segun ordenacion ...
            '       costesfe


            '66    Horas procesadas (tipo alzira)

            '67    marcaje actual sin secciones ni nada


Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |
                                   ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public EnvioEMail As Boolean

Public NombreRPT100 As String
Public Titulo100 As String
Public ConSubreport100 As Boolean




Private MostrarTree As Boolean
Private Nombre As String
Private MIPATH As String
Private Lanzado As Boolean
Private PrimeraVez As Boolean
Private LlevaSubinforme As Boolean

'Private ReestableceSoloImprimir As Boolean
Private Sub chkEmail_Click()
    If chkEMAIL.Value = 1 Then Me.chkSoloImprimir.Value = 0
End Sub

Private Sub chkSoloImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 Then Me.chkEMAIL.Value = 0
End Sub

Private Sub cmdConfigImpre_Click()
Screen.MousePointer = vbHourglass
'Me.CommonDialog1.Flags = cdlPDPageNums
CommonDialog1.ShowPrinter
PonerNombreImpresora
Screen.MousePointer = vbDefault
End Sub


Private Sub cmdImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 And Me.chkEMAIL.Value = 1 Then
        MsgBox "Si desea enviar por mail no debe marcar vista preliminar", vbExclamation
        Exit Sub
    End If
    'Form2.Show vbModal
    If Dir(MIPATH & Nombre, vbArchive) = "" Then
        MsgBox "No existe el fichero: " & MIPATH & Nombre, vbExclamation
    Else
        Imprime
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
If PrimeraVez Then
    espera 0.1
    
    If SoloImprimir Then
        Imprime
        Unload Me
    Else
        If EnvioEMail Then
            Me.Hide
            chkEMAIL.Value = 1
            Imprime
            Unload Me
        End If
    End If
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim cad As String

PrimeraVez = True
Lanzado = False
CargaICO
cad = Dir(App.Path & "\impre.dat", vbArchive)


'ReestableceSoloImprimir = False
If cad = "" Then
    chkSoloImprimir.Value = 0
    Else
    chkSoloImprimir.Value = 1
    'ReestableceSoloImprimir = True
End If
cmdImprimir.Enabled = True
If SoloImprimir Then
    chkSoloImprimir.Value = 0
    Me.Frame2.Enabled = False
    chkSoloImprimir.Visible = False
Else
    Frame2.Enabled = True
    chkSoloImprimir.Visible = True
End If
PonerNombreImpresora
MostrarTree = False
LlevaSubinforme = False
'A partir del infome 26, se trabajaba sobre la b de datos de informes(USUARIOS)


    MIPATH = App.Path & "\Informes\"


Select Case Opcion
Case 0
    MostrarTree = True
    Text1.Text = "Listado marcajes por fecha/codigo"
    Nombre = "marfechcod.rpt"
              
Case 1
    MostrarTree = True
    Text1.Text = "Listado marcajes por fecha/nombre"
    Nombre = "marfechnom.rpt"
Case 2
    MostrarTree = True
    Text1.Text = "Listado marcajes por nombre/fecha"
    Nombre = "martrabnom.rpt"
Case 3
    MostrarTree = True
    Text1.Text = "Listado marcajes por codigo/fecha"
    Nombre = "martrabcod.rpt"


Case 6, 7
    MostrarTree = True
    Text1.Text = "Listado presencia por fecha"
    Nombre = "realfec"
    If Opcion = 6 Then
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    Else
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    End If

Case 8
    MostrarTree = True
    Text1.Text = "Listado presencia por Nombre"
    Nombre = "realnom.rpt"
Case 9
    MostrarTree = True
    Text1.Text = "Listado presencia por cógigo"
    Nombre = "realcod.rpt"



Case 10, 11
    Text1.Text = "Trabajadores basico "
    Nombre = "trabbassin"
    If Opcion = 10 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    
Case 12, 13
    Text1.Text = "Trabajadores extendido "
    Nombre = "trabextsin"
    If Opcion = 12 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    End If


Case 14, 15
    Text1.Text = "Trabajadores con foto "
    Nombre = "trabbascon"
    If Opcion = 14 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    End If

Case 16, 17
    Text1.Text = "Trabajadores extendido con foto "
    Nombre = "trabextcon"
    If Opcion = 16 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    
    
Case 18, 19
    Text1.Text = "Trabajadores por seccion basico"
    Nombre = "trabsecbassin"
    If Opcion = 18 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    End If

    MostrarTree = True
    
Case 20, 21
    Text1.Text = "Trabajadores por seccion extendido"
    Nombre = "trabsecextsin"
    If Opcion = 20 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    MostrarTree = True
Case 22, 23
       Text1.Text = "Trabajadores seccion basicos con foto "
    Nombre = "trabsecbascon"
    If Opcion = 24 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
 MostrarTree = True
'
Case 24, 25
    Text1.Text = "Trabajadores seccion extendido con foto "
    Nombre = "trabseccon"
    If Opcion = 24 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & "(Codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & "(Nombre)"
    End If
    MostrarTree = True
    
Case 30
    Text1.Text = "Horarios"
    Nombre = "rHorarios.rpt"
    LlevaSubinforme = True
Case 31
    Text1.Text = "Secciones"
    Nombre = "rSeccion.rpt"
    
Case 32
    Text1.Text = "Marcaje actual"
    Nombre = "marcactual.rpt"
    MostrarTree = True
Case 33
    Text1.Text = "Incidencia resumen codigo"
    Nombre = "incirescod.rpt"
    MostrarTree = True
    
Case 34
    Text1.Text = "Incidencia resumen nombre"
    Nombre = "inciresnom.rpt"
    MostrarTree = True

Case 35, 36, 37, 38
    Text1.Text = "Listado combinado horas"
    Nombre = "combinada"
    If Opcion > 36 Then
        Text1.Text = Text1.Text & ". Fecha"
        Nombre = Nombre & "fecha"
    End If
    
    If Opcion = 35 Or Opcion = 37 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & " (codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & " (nombre)"
    End If
    MostrarTree = True




Case 40
    Text1.Text = "Listado combinado horas"
    Nombre = "diastranajados01.rpt"
    
    LlevaSubinforme = True



Case 50, 51
    Text1.Text = "Informe incidencias generadas"
    Nombre = "incirgen"  'la dislexia es lo que tiene
    
    If Opcion = 50 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & " (codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & " (nombre)"
    End If
    MostrarTree = True
    
    
    
Case 52, 53
    
    Text1.Text = "Incidencias generadas por trabajador"
    Nombre = "incirestrab"
    
    If Opcion = 52 Then
        Nombre = Nombre & "cod"
        Text1.Text = Text1.Text & " (codigo)"
    Else
        Nombre = Nombre & "nom"
        Text1.Text = Text1.Text & " (nombre)"
    End If
    Nombre = Nombre & "desg.rpt"
    MostrarTree = True
    
Case 54, 55
    
    Text1.Text = "Incidencias resumen por trabajador"
    Nombre = "incirestrab"
    
    If Opcion = 54 Then
        Nombre = Nombre & "cod.rpt"
        Text1.Text = Text1.Text & " (codigo)"
    Else
        Nombre = Nombre & "nom.rpt"
        Text1.Text = Text1.Text & " (nombre)"
    End If
    MostrarTree = True
     
Case 60
    Text1.Text = "Marcaje actual por trabajador"
    Nombre = "marcactualt.rpt"
    MostrarTree = True

Case 61, 62, 63, 64
    If Opcion > 62 Then
        Text1.Text = "Horas con importe Trabajador"
        Nombre = "CosteTra"
    Else
        Text1.Text = "Horas con importe Fecha"
        Nombre = "CosteFech"
    End If
    If Opcion = 62 Or Opcion = 64 Then Nombre = Nombre & "2"
    Nombre = Nombre & ".rpt"
    MostrarTree = True
    'Nombre = "CosteFech.rpt"
    
    'Nombre = "CosteTra.rpt"
    
    'Nombre = "CosteFech2.rpt"
    
    'Nombre = "CosteTra2.rpt"

Case 65
    
    Text1.Text = "Horas procesadas"
    
    Nombre = NombreRPT100
    MostrarTree = True
    
    
Case 67
    Text1.Text = "Marcahe actual normal"
    Nombre = "marcactualPlano.rpt"
    MostrarTree = True
Case 100
    'GENERICO. Se le pasa el rpt y el titulo
    Nombre = NombreRPT100
    Text1.Text = Titulo100
    LlevaSubinforme = ConSubreport100
    
Case Else
    Text1.Text = "Opcion incorrecta"
    Me.cmdImprimir.Enabled = False
    
    
End Select



Screen.MousePointer = vbDefault
End Sub




Private Function Imprime() As Boolean
Dim Seguir As Boolean

    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = Me.FormulaSeleccion
        .SoloImprimir = (Me.chkSoloImprimir.Value = 0)
        .OtrosParametros = OtrosParametros
        .NumeroParametros = NumeroParametros
        .MostrarTree = MostrarTree
        .Informe = MIPATH & Nombre
        .ConSubinforme = LlevaSubinforme
        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
    If Me.chkEMAIL.Value = 1 Then
       If CadenaDesdeOtroForm <> "" Then
            frmEMail.Opcion = 0
            frmEMail.Show vbModal
       End If
       CadenaDesdeOtroForm = ""
    End If
    Unload Me
 
 
 
End Function


Private Sub Form_Unload(Cancel As Integer)
    If Me.chkEMAIL.Value = 1 Then Me.chkSoloImprimir.Value = 1
    'If ReestableceSoloImprimir Then SoloImprimir = False
    OperacionesArchivoDefecto
    
    
    NombreRPT100 = ""
    Titulo100 = ""
    ConSubreport100 = False
End Sub

Private Sub OperacionesArchivoDefecto()
Dim crear  As Boolean
Dim NF As Integer
On Error GoTo ErrOperacionesArchivoDefecto

crear = (Me.chkSoloImprimir.Value = 1)
'crear = crear And ReestableceSoloImprimir
If Not crear Then
    Kill App.Path & "\impre.dat"
    Else
        NF = FreeFile
        Open App.Path & "\impre.dat" For Output As #NF
        Print #NF, Format(Now, "Long Date")
        Close #NF
End If
ErrOperacionesArchivoDefecto:
    If Err.Number <> 0 Then Err.Clear
    End Sub


Private Sub Text1_DblClick()
    Frame2.Tag = Val(Frame2.Tag) + 1
    If Val(Frame2.Tag) > 2 Then
        Frame2.Enabled = True
        chkSoloImprimir.Visible = True
    End If
End Sub

Private Sub PonerNombreImpresora()
On Error Resume Next
    Label1.Caption = Printer.DeviceName
    If Err.Number <> 0 Then
        Label1.Caption = "No hay impresora instalada"
        Err.Clear
    End If
End Sub

Private Sub CargaICO()
    On Error Resume Next
    Image1.Picture = LoadPicture(App.Path & "\iconos\printer.ico")
    If Err.Number <> 0 Then Err.Clear
End Sub

