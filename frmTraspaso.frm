VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmTraspaso 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de ficheros"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmTraspaso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameKimalid 
      BackColor       =   &H00C0C0C0&
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdKimaldi 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   13
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdKimaldi 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5295
         Begin VB.TextBox txtFecha 
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   10
            Top             =   1020
            Width           =   1395
         End
         Begin VB.TextBox txtFecha 
            Height          =   315
            Index           =   1
            Left            =   3600
            TabIndex        =   11
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Height          =   495
            Left            =   120
            TabIndex        =   18
            Top             =   1560
            Width           =   4815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Dias completos y sin que ya se hubieran generado datos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   5175
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fechas entre las cuales se generan fichajes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Fin"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2940
            TabIndex        =   14
            Top             =   1080
            Width           =   315
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   720
            Picture         =   "frmTraspaso.frx":6852
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   3300
            Picture         =   "frmTraspaso.frx":6954
            Top             =   1080
            Width           =   240
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar  ult reg"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2700
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1155
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   5415
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importando fichero"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   2700
      Width           =   1515
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Iniciar"
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   2700
      Width           =   1275
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   915
      Left            =   480
      TabIndex        =   2
      Top             =   2220
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1614
      _Version        =   327681
      BackColor       =   12632256
      FullWidth       =   305
      FullHeight      =   61
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   180
      TabIndex        =   4
      Top             =   1620
      Width           =   5175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   5175
   End
End
Attribute VB_Name = "frmTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HoraM As Date = "08:00"   'PARA el horario nocturno
Private Const HoraT As Date = "20:00"   '

Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1
Public ContadorSecuencia As Long
Public Opcion As Byte
        'Con OPCION determinaremos lo k hacer:
        '   0.- Abrimos el form normal ( Han pulsado sobre el
        '          label )
        '   1.- Ejecutamos el cmdClick pq venimos de traer los datos
        '       desde el terminal
        
        '       Nueva 20 OCutbre 2004
        '   2.- Procesaremos las fechas de Kimaldi
        '       Generando primero el fichero Fichajes
        
'Private MiNF As Integer
Dim PrimeraVez As Boolean
Dim ImpFechaIni As String
Dim ImpFechaFin As String


Private Sub cmdAceptar_Click()
Dim CADENA As String
Dim Archivo As String
    'Primeras comprobaciones
    If vEmpresa.DirMarcajes = "" Then
        MsgBox "Sin configurar 'directorio'", vbExclamation
        Exit Sub
    End If
    
    If vEmpresa.reloj = vbRobotics Then
        'Para robotics vamos leyneod tooodos los archivos que se hayen en esa carpeta
        CADENA = vEmpresa.DirMarcajes & "\*.*"
    Else
        CADENA = vEmpresa.DirMarcajes & "\" & vEmpresa.NomFich
    End If
    If Dir(CADENA, vbArchive) = "" Then
        MsgBox "No existe archivo para procesar: " & CADENA, vbExclamation
        Exit Sub
    End If


    
    

    If Opcion <> 1 Then
        Archivo = "El proceso puede llevar mucho tiempo. ¿Continuar?"
        If MsgBox(Archivo, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    AccionDeImportarFichero
End Sub

Private Sub AccionDeImportarFichero()
Dim valor As Byte
Dim CADENA As String
Dim Archivo As String
Dim Valores As String
Dim MaxFecha As Date
Dim vLeeBienCrLF As Boolean

    Screen.MousePointer = vbHourglass
    cmdAceptar.Visible = False
    cmdSalir.Visible = False
    Command1.Visible = False
    Me.Refresh
    Me.Animation1.Open App.Path & "\FILEDELR.AVI"
    Me.Animation1.Play
    Label11.Caption = "Iniciando importación de fichero marcajes."
    Label11.Refresh
    valor = 0
    Valores = ""
    
    
    'Por si se producen errores
    InicializaErroresLinea (Now)
    
    'Si no lee bien los saltos de linea
    'como le pasa al PC de Manuel en Herbelca
    If Dir(App.Path & "\Lee2Fich.dat", vbArchive) = "" Then
        vLeeBienCrLF = True
    Else
        vLeeBienCrLF = False
    End If
    
    
    'Primero leemos los datos
    ' 0.- Todo correcto
    ' 1.- No existe el fichero o está vacio
    ' 2.- Algun fallo
    If vEmpresa.reloj = vbRobotics Then Archivo = Dir(vEmpresa.DirMarcajes & "\*.*", vbArchive)


        
    Do
        If vEmpresa.reloj <> vbRobotics Then
            Archivo = vEmpresa.DirMarcajes & "\" & vEmpresa.NomFich
        Else
            Archivo = vEmpresa.DirMarcajes & "\" & Archivo
            AsignaNomFichero Archivo
        End If
        Label11.Caption = "Archivo: " & Archivo
        Label11.Refresh
        espera 0.2
        
        Label10.Caption = Archivo
        Label10.Refresh
        valor = ProcesaFichero(Archivo, vLeeBienCrLF)
        If valor <> 1 Then
            'Ahora comprobamos si se han quedado entradas por procesar
            valor = VaciarTemporales(CADENA)
            If CADENA <> "" Then
                If vEmpresa.reloj = vbRobotics Then
                    Valores = Valores & "-->" & Archivo & vbCrLf & CADENA & vbCrLf & vbCrLf
                Else
                    MsgBox CADENA, vbExclamation
                    Valores = CADENA
                End If
            End If
        Else
            'Error en mover fichero
            Archivo = ""
        End If
        
        'Eliminamos el archivo que ya ha sido traspasado
        If valor = 0 Then
            If EliminarArchivo(Archivo) Then
                
                
                If vEmpresa.reloj = vbRobotics Then
                    Archivo = Dir(vEmpresa.DirMarcajes & "\*.*", vbArchive)
                Else
                    Archivo = ""
                End If
            Else
                Archivo = ""
                Valores = "SALIR"
            End If
        End If
    
    
    
    Loop Until Archivo = ""
    
    
    'Cerramos el fichero de errores
    FinErroresLinea

    
    
    'Renumeramos entradas
    Label10.Caption = "Comprobando ajustes secuencia"
    Label11.Caption = ""
    Me.Refresh
    Set miRsAux = New ADODB.Recordset
    CADENA = "Select min(secuencia) from entradafichajes"
    ContadorSecuencia = 0
    miRsAux.Open CADENA, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ContadorSecuencia = DBLet(miRsAux.Fields(0), "N")
        If ContadorSecuencia < 15000 Then ContadorSecuencia = 0
        
     
    End If
    miRsAux.Close
    
    If ContadorSecuencia > 0 Then
        Label10.Caption = "Realizando ajustes secuencia"
        DoEvents
        CADENA = "Select secuencia from entradafichajes order by secuencia asc"
        miRsAux.Open CADENA, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ContadorSecuencia = 1
        While Not miRsAux.EOF
            Label11.Caption = "Sec: " & ContadorSecuencia
            Label11.Refresh
            
            CADENA = "UPDATE entradafichajes set secuencia=" & ContadorSecuencia & " WHERE secuencia =" & miRsAux!Secuencia
            miRsAux.MoveNext
            ContadorSecuencia = ContadorSecuencia + 1
            EjecutaSQL CADENA
        Wend
        miRsAux.Close
        
    End If
    
    
     

    
    Set miRsAux = Nothing
    
    'Paramos el avi
    Me.Animation1.Stop
    Me.Animation1.Close
    'Restauramos lo del avi
    cmdAceptar.Visible = True
    cmdSalir.Visible = True
    Me.Refresh
    If Valores <> "" Then MsgBox "Se han producido errores.", vbExclamation
    Me.cmdAceptar.Enabled = False
    Label10.Caption = ""
    Label11.Caption = "Importación finalizada."
    Label11.Refresh
    Screen.MousePointer = vbDefault
End Sub

Private Function EliminarArchivo(ElArchivo As String) As Boolean
    On Error Resume Next
    Kill ElArchivo
    If Err.Number = 0 Then
        EliminarArchivo = True
    Else
        EliminarArchivo = False
        MsgBox "Error eliminando archivo: " & ElArchivo & vbCrLf & Err.Description, vbExclamation
    End If
End Function

Private Sub cmdKimaldi_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If

    If Dir(vEmpresa.DirMarcajes & "\" & vEmpresa.NomFich, vbArchive) <> "" Then
        MsgBox "Se ha quedado el fichero por procesar. Eliminielo." & vbCrLf & vEmpresa.DirMarcajes & "\" & vEmpresa.NomFich, vbExclamation
        Exit Sub
    End If
    
    If Me.txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        MsgBox "Debe poner las fechas", vbExclamation
        Exit Sub
    End If
    
    If CDate(Me.txtFecha(0).Text) > CDate(txtFecha(1).Text) Then
        MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
        Exit Sub
    End If
    
    If txtFecha(0).Tag <> "" Then
        If CDate(Me.txtFecha(0).Text) < CDate(txtFecha(0).Tag) Then
            If MsgBox("Fecha inicio menor minima fecha ofertada. ¿Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    ImpFechaIni = "#" & Format(txtFecha(0).Text, FormatoFecha) & "#"
    ImpFechaFin = "#" & Format(txtFecha(1).Text, FormatoFecha) & "#"
    If GeneraMarcajesKimaldi Then
        Label6.Caption = "Proceso 2: -------"
        Me.Refresh
        espera 1
        'OK
        'Luego hay k procesar el fichero como si vinieramos del TCP3
        Me.FrameKimalid.Visible = False
        Me.Refresh
        espera 1
        cmdAceptar_Click
        Unload Me
    End If
    Label6.Caption = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub



Private Sub Command1_Click()
Dim RS As ADODB.Recordset
Dim vF As String
Dim FE As Date

On Error GoTo ErrCom1
Screen.MousePointer = vbHourglass

vF = ""
Set RS = New ADODB.Recordset
RS.Open "Select MAX(Fecha) from EntradaFichajes", conn, , , adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then
        FE = RS.Fields(0)
        vF = "'" & Format(FE, FormatoFecha) & "'"
    End If
    RS.Close
    'Ahora si VF<>"0:00:" entonces buscamos la hora
    If vF <> "" Then
        RS.Open "Select MAX(Hora) from EntradaFichajes where Fecha=" & vF, conn, , , adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then
                MsgBox "    Ultimo fichaje" & vbCrLf & _
                    "------------------------" & vbCrLf & vbCrLf & _
                    "Fecha:  " & Format(FE, "dd/mm/yyyy") & vbCrLf & _
                    "Hora  :  " & Format(RS.Fields(0), "hh:mm:ss"), vbInformation
            End If
            Else
                vF = ""
        End If
        RS.Close
    End If
End If
If vF = "" Then MsgBox "No hay registros para mostrar datos.", vbExclamation
Set RS = Nothing
ErrCom1:
    If Err.Number <> 0 Then MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False
    If Opcion = 1 Then
    
        cmdAceptar_Click
        Unload Me
        'ELSE opcion=0
    Else
            If vEmpresa.reloj = vbTCP3 Then
                MsgBox "La forma de importar los marcajes es desde la ventana" & vbCrLf & _
                    " de operaciones\TCP 3 , con el boton crear fichero." & vbCrLf & " Si quiere procesar un fichero YA " & _
                    " procesado continue en esta pantalla. En caso contrario, salga de ella.", vbCritical _
                    , "I M P O R T A N T E"
            End If
            'If Opcion = 2 Then cmdKimaldi(0).SetFocus
    End If
End If
Screen.MousePointer = vbDefault
End Sub


Public Function ProcesaFichero(NombreFicher As String, LeeBienCrLF As Boolean) As Byte
'---------------------
'Valores que devuelve la function
' 0.- Todo correcto
' 1.- No existe el fichero o esta vacio
' 2.- Algun fallo
Dim Cad As String
Dim NF As Integer
Dim Errores As Byte
Dim NombreEnProcesados As String
Dim T1 As Long
Dim T2 As Long
Dim anyo As Integer

'----------------------------
Dim PuntoDeInicio As Integer
'       Si es ALZICOOP enonces el punto de inicio es el 1
'       si es catadau, entonces el punto de inicio es el 2


Dim ColLineas As Collection
    'Traspasare el fichero al collection
    'Despues ire uno a uno procesando cada linea
        
On Error GoTo ErrProcesaFichero


    If Opcion = 2 Then
        PuntoDeInicio = 2
    Else
        PuntoDeInicio = 1
    End If
    
    '------------------------------------------------------------------
    '------------------------------------------------------------------
    'Para saber que tablas auxiliares utilizara y como se procesaran
    'las lineas
    'Estas lineas tendran que ser parametrizables puesto que son excluyentes dos a dos
    
    
    'En lineas generales TIPOALZICOOP = True significa k es un control de produccion
    
    
    'TIPOALZICOOP = True
    
    
    T1 = FileLen(NombreFicher)
    T2 = 0
    
    
    'Abrimos el fichero y lo procesamos
    NF = FreeFile
    
    'Abrimos el fichero
    Open NombreFicher For Input As NF
    'Indicativos
    Label11.Caption = "Abriendo fichero para lectura"
    Label11.Refresh
    
    
    anyo = Year(Now)
    
    
    'Borramos el temporal
    conn.Execute "Delete  from TemporalFichajes"
    
    ObtenerPrimeraClave  'para saber el nº de secuencia
    
    Set ColLineas = New Collection
    If LeeBienCrLF Then
    
        While Not EOF(NF)
            'la procesamos
                Label11.Caption = "Importación.   Secuencia: " & ContadorSecuencia & "    "
                If T2 <= T1 Then
                    Label11.Caption = Label11.Caption & "(" & Round(T2 / T1, 2) * 100 & "%)"
                End If
                Label11.Refresh
                
                
                Line Input #NF, Cad
                T2 = T2 + Len(Cad)
                If Cad <> "" Then ColLineas.Add Cad
        '        Debug.Print Cad
        '        Dim Parar As String
        '        Parar = Mid(Cad, 1, 27)
        '
                
                
        Wend
                
                
                
        
        
    Else
        'lee mal el salto de linea, con lo cual ire leyendo carcarter a caracter buscando la linea
        Cad = ""
        
        While Not EOF(NF)
            T2 = T2 + 1
            Label11.Caption = "Importación(2).   Secuencia: " & ContadorSecuencia & Format(T2, "0000") & "  /  " & Format(T1, "0000")
            NombreEnProcesados = Input(1, #NF)
            If NombreEnProcesados = vbCr Then
                'Leo el siguiente caracter
                NombreEnProcesados = ""
                If Not EOF(NF) Then
                    NombreEnProcesados = Input(1, #NF)
                    T2 = T2 + 1
                    If NombreEnProcesados = vbLf Then NombreEnProcesados = ""
                End If
                If Cad <> "" Then ColLineas.Add Cad
                Cad = NombreEnProcesados
            Else
                Cad = Cad & NombreEnProcesados
            End If
        Wend
        
    End If
    Close #NF
                
    
                
    For NF = 1 To ColLineas.Count
        Cad = ColLineas.Item(NF)
        Label11.Caption = "Procesar linea: " & NF & " / " & ColLineas.Count
        Label11.Refresh
        If (NF Mod 20) = 0 Then DoEvents
        If Cad <> "" Then
                    
                Select Case vEmpresa.reloj
                Case vbAlzira
                
                        ProcesarLineaALZ Cad, ContadorSecuencia, PuntoDeInicio
                
                
                
                
                Case Else
                
                        If vEmpresa.reloj = vbRobotics Then Cad = TransformaLineaRobotics(Cad, anyo)
                        ProcesarLinea Cad, ContadorSecuencia, anyo
                End Select
            End If
            ContadorSecuencia = ContadorSecuencia + 1
            '--------------------------------------------
    Next NF
    
    
    
    'Si es del tipo ALZICOOP tendremos que generar, a partir del 1er
    'fichero generado el fichero de ENTRADAFICHAJES
    'desde el cual generaremos los marcajes
    If vEmpresa.reloj = vbAlzira Then GeneraEntradasALZ
    
    
    'Pasamos el fichero a carpeta procesados
    'Indicativos
    Label11.Caption = "Moviendo fichero a procesados"
    Label11.Refresh
    'Obtenemos el nombre del archivo en procesados
    NombreEnProcesados = NombreEnProcesado(NombreFicher)
    
    'Copiamos
    FileCopy NombreFicher, NombreEnProcesados
    
    
    'Indicativos
    Label11.Caption = "Fichero en procesados: " & NombreEnProcesados
    Label11.Refresh
    
    
    ' quitamos los marcajes de temporal
    Errores = TraspasaTemporal(vEmpresa.reloj = vbAlzira)
    
    'CierraFichero (CuantosErrores2 = 0)
    
    
    
    'Modificacion de tele-taxi.
    '--------------------------
    'Horario nocturno.
    'Tiene el hoario nocturno. Es decir, es teletaxi
    If vEmpresa.HorarioNocturno2 Then
        'Si la empresa fuera teletaxi....
        If False Then InsertarTicajesTeletaxi
    End If
    
    
    
    'mensaje de todo correcto
    If Errores <> 0 Then
        ProcesaFichero = 2
    Else
        ProcesaFichero = 0
    End If
    Set ColLineas = Nothing
    Exit Function
ErrProcesaFichero:
        Cad = "Se ha producido un error mientras se procesaba el fichero de marcajes." & vbCrLf
        Cad = Cad & " RUTA: " & NombreFicher & vbCrLf
        Cad = Cad & " ERROR: " & vbCrLf
        Cad = Cad & "      Número: " & Err.Number & vbCrLf
        Cad = Cad & "      Descripción: " & Err.Description
        MsgBox Cad, vbCritical
        EscribeErrorLinea Cad
        ProcesaFichero = 1
        Set ColLineas = Nothing
End Function





'COpiar pegar por si lña cago cambiando la funcion ProcesarFichero
''''''''Public Function ProcesaFichero(NombreFicher As String) As Byte
'''''''''---------------------
'''''''''Valores que devuelve la function
''''''''' 0.- Todo correcto
''''''''' 1.- No existe el fichero o esta vacio
''''''''' 2.- Algun fallo
''''''''Dim Cad As String
''''''''Dim NF As Integer
''''''''Dim Errores As Byte
''''''''Dim NombreEnProcesados As String
''''''''Dim T1 As Long
''''''''Dim T2 As Long
''''''''Dim Anyo As Integer
''''''''
'''''''''----------------------------
''''''''Dim PuntoDeInicio As Integer
'''''''''       Si es ALZICOOP enonces el punto de inicio es el 1
'''''''''       si es catadau, entonces el punto de inicio es el 2
''''''''On Error GoTo ErrProcesaFichero
''''''''
''''''''
''''''''If Opcion = 2 Then
''''''''    PuntoDeInicio = 2
''''''''Else
''''''''    PuntoDeInicio = 1
''''''''End If
''''''''
'''''''''------------------------------------------------------------------
'''''''''------------------------------------------------------------------
'''''''''Para saber que tablas auxiliares utilizara y como se procesaran
'''''''''las lineas
'''''''''Estas lineas tendran que ser parametrizables puesto que son excluyentes dos a dos
''''''''
''''''''
'''''''''En lineas generales TIPOALZICOOP = True significa k es un control de produccion
''''''''
''''''''
'''''''''TIPOALZICOOP = True
''''''''
''''''''
''''''''T1 = FileLen(NombreFicher)
''''''''T2 = 0
''''''''
''''''''
'''''''''Abrimos el fichero y lo procesamos
''''''''NF = FreeFile
''''''''
'''''''''Abrimos el fichero
''''''''Open NombreFicher For Input As NF
'''''''''Indicativos
''''''''Label11.Caption = "Abriendo fichero para lectura"
''''''''Label11.Refresh
''''''''
''''''''
''''''''Anyo = Year(Now)
''''''''
''''''''
'''''''''Borramos el temporal
''''''''Conn.Execute "Delete  from TemporalFichajes"
''''''''
''''''''ObtenerPrimeraClave  'para saber el nº de secuencia
''''''''
''''''''While Not EOF(NF)
''''''''    'la procesamos
''''''''        Label11.Caption = "Importación.   Secuencia: " & ContadorSecuencia & "    "
''''''''        If T2 <= T1 Then
''''''''            Label11.Caption = Label11.Caption & "(" & Round(T2 / T1, 2) * 100 & "%)"
''''''''        End If
''''''''        Label11.Refresh
''''''''
''''''''
''''''''        Line Input #NF, Cad
''''''''        T2 = T2 + Len(Cad)
'''''''''        Debug.Print Cad
'''''''''        Dim Parar As String
'''''''''        Parar = Mid(Cad, 1, 27)
'''''''''
''''''''
''''''''
''''''''        If Cad <> "" Then
''''''''
''''''''            Select Case vEmpresa.reloj
''''''''            Case vbAlzira
''''''''
''''''''                    ProcesarLineaALZ Cad, ContadorSecuencia, PuntoDeInicio
''''''''
''''''''
''''''''
''''''''
''''''''            Case Else
''''''''
''''''''                    If vEmpresa.reloj = vbRobotics Then Cad = TransformaLineaRobotics(Cad, Anyo)
''''''''                    ProcesarLinea Cad, ContadorSecuencia, Anyo
''''''''            End Select
''''''''        End If
''''''''        ContadorSecuencia = ContadorSecuencia + 1
''''''''        '--------------------------------------------
''''''''Wend
''''''''Close #NF
''''''''
'''''''''Si es del tipo ALZICOOP tendremos que generar, a partir del 1er
'''''''''fichero generado el fichero de ENTRADAFICHAJES
'''''''''desde el cual generaremos los marcajes
''''''''If vEmpresa.reloj = vbAlzira Then GeneraEntradasALZ
''''''''
''''''''
'''''''''Pasamos el fichero a carpeta procesados
'''''''''Indicativos
''''''''Label11.Caption = "Moviendo fichero a procesados"
''''''''Label11.Refresh
'''''''''Obtenemos el nombre del archivo en procesados
''''''''NombreEnProcesados = NombreEnProcesado(NombreFicher)
''''''''
'''''''''Copiamos
''''''''FileCopy NombreFicher, NombreEnProcesados
''''''''
''''''''
'''''''''Indicativos
''''''''Label11.Caption = "Fichero en procesados: " & NombreEnProcesados
''''''''Label11.Refresh
''''''''
''''''''
''''''''' quitamos los marcajes de temporal
''''''''Errores = TraspasaTemporal(vEmpresa.reloj = vbAlzira)
''''''''
'''''''''CierraFichero (CuantosErrores2 = 0)
''''''''
''''''''
''''''''
'''''''''Modificacion de tele-taxi.
'''''''''--------------------------
'''''''''Horario nocturno.
'''''''''Tiene el hoario nocturno. Es decir, es teletaxi
''''''''If vEmpresa.HorarioNocturno Then InsertarTicajesTeletaxi
''''''''
''''''''
''''''''
'''''''''mensaje de todo correcto
''''''''If Errores <> 0 Then
''''''''    ProcesaFicheroMAL_SaltoLinea = 2
''''''''Else
''''''''    ProcesaFicheroMAL_SaltoLinea = 0
''''''''End If
''''''''Exit Function
''''''''ErrProcesaFichero:
''''''''    Cad = "Se ha producido un error mientras se procesaba el fichero de marcajes." & vbCrLf
''''''''    Cad = Cad & " RUTA: " & NombreFicher & vbCrLf
''''''''    Cad = Cad & " ERROR: " & vbCrLf
''''''''    Cad = Cad & "      Número: " & Err.Number & vbCrLf
''''''''    Cad = Cad & "      Descripción: " & Err.Description
''''''''    MsgBox Cad, vbCritical
''''''''    EscribeErrorLinea Cad
''''''''    ProcesaFicheroMAL_SaltoLinea = 1
''''''''End Function






'Segun sea un control normal, o un control tipo ALZICOOP
Private Sub ObtenerPrimeraClave()
Dim RS As ADODB.Recordset
Dim Cad As String

Set RS = New ADODB.Recordset
ContadorSecuencia = 1
If vEmpresa.reloj = vbAlzira Then
        Cad = "SELECT MAX(secuencia) FROM TipoAlzicoop"
        RS.Open Cad, conn, , , adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then
                ContadorSecuencia = RS.Fields(0) + 1
            End If
        End If
    'ELSE
    Else
        Cad = "SELECT MAX(secuencia) FROM TemporalFichajes"
        RS.Open Cad, conn, , , adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then
                ContadorSecuencia = RS.Fields(0) + 1
            End If
        End If
End If
RS.Close
End Sub




'Los errores que devuelve son
' 0.- Ningun error

Private Function TraspasaTemporal(SoloRectifica As Boolean) As Integer
Dim RIni As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim RError As ADODB.Recordset
Dim CadInci As String
Dim Codigo As Long
Dim ContError As Long
Dim ContFich As Long
Dim Totales As Long
Dim KReg As Long
Dim PrimeraInsercion As Long
Dim Fecha As Date
Dim Hora As Date
Dim Ntarj As String
Set RIni = New ADODB.Recordset
Dim CuantosErrores2 As Integer



'Proceso de traspaso de datos desde temporal. Solo para TCP3
If Not SoloRectifica Then
    
    RIni.CursorType = adOpenKeyset
    RIni.Open "Select * from TemporalFichajes order by numtarjeta", conn, , , adCmdText
    If RIni.EOF Then
        TraspasaTemporal = 127
        RIni.Close
        Set RIni = Nothing
        Exit Function
    End If
    Totales = 0
    While Not RIni.EOF
        Totales = Totales + 1
        RIni.MoveNext
    Wend
  
    RIni.MoveFirst
    'Indicativos
    Label11.Caption = "Pasando temporal a tabla control de fichajes "
    Label11.Refresh
    'marcamos por si error
    On Error GoTo ErrTraspaso
    
    
 
      
    
    
    Set RFin = New ADODB.Recordset
    Set RError = New ADODB.Recordset
    RFin.CursorType = adOpenKeyset
    RFin.LockType = adLockOptimistic
    RError.CursorType = adOpenKeyset
    RError.LockType = adLockOptimistic
    'Abrimos los regsitros destino
    ' el correcto y el erroneo
    RError.Open "select * from ErrorTarjetas order by secuencia ", conn, , , adCmdText
    If RError.EOF Then
        ContError = 1
        Else
            RError.MoveLast
            ContError = RError!Secuencia + 1
    End If
    RFin.Open "Select * from EntradaFichajes order by secuencia", conn, , , adCmdText
    If RFin.EOF Then
        ContFich = 1
        Else
            RFin.MoveLast
            ContFich = RFin!Secuencia + 1
    End If
    PrimeraInsercion = ContFich   'Para despues saber cual es el minimo de los dia que acabo de traspasar
    CuantosErrores2 = 0
    KReg = 0
    Ntarj = ""
    While Not RIni.EOF
        KReg = KReg + 1
        If RIni!idInci = 0 Then
            CadInci = "OK"
            Else
            'FALTA###
            CadInci = DevuelveTextoIncidencia(RIni!idInci)
        End If
        'FALTA###
        If Ntarj = "" Then
            Ntarj = RIni!Numtarjeta
            Codigo = DevuelveCodigo(RIni!Numtarjeta)
        Else
            If Ntarj <> RIni!Numtarjeta Then
                Ntarj = RIni!Numtarjeta
                Codigo = DevuelveCodigo(RIni!Numtarjeta)
            End If
        End If
        If CadInci = "" Or Codigo < 0 Then
            CuantosErrores2 = CuantosErrores2 + 1
            'Ha habido un error
            RError.AddNew
            RError!Secuencia = ContError
            RError!Fecha = RIni!Fecha
            RError!Hora = Format(RIni!Hora, "hh:mm:ss")
            RError!idInci = RIni!idInci
            RError!Numtarjeta = RIni!Numtarjeta
            If CadInci = "" Then RError!Error = "La incidencia no es correcta"
            If Codigo < 0 Then RError!Error = RError!Error & "  ---  El codigo de tarjeta no corresponde con ningun trabajador"
            RError.Update
            
            If CuantosErrores2 = 1 Then EscribeErrorLinea "  Traspasando tabla temporal"
            EscribeErrorLinea "    " & RIni!Numtarjeta & "  -- " & RError!Error
            
            ContError = ContError + 1
            'Si es correcto
            Else
                RFin.AddNew
                RFin!Secuencia = ContFich
                RFin!Fecha = RIni!Fecha
                RFin!Hora = Format(RIni!Hora, "hh:mm:ss")
                RFin!HoraReal = RIni!Hora
                RFin!idInci = RIni!idInci
                RFin!idTrabajador = Codigo
                ContFich = ContFich + 1
                RFin.Update
                
    
    
            End If
        RIni.MoveNext
        'Indicativos
        Label11.Caption = "Registro temporal:  " & KReg & " de " & Totales
        Label11.Refresh
        If (ContFich Mod 100) = 0 Then DoEvents
    Wend
    RIni.Close
    RFin.Close
    RError.Close
    Set RFin = Nothing
    Set RError = Nothing
    
End If  ' Bloque para los k ademas de rectifcar procesan



'ELIMINAR MARCAJES REPETIDOS CON INTERVALOS DE TIEMPO PEQUEÑOS
    'Indicativos
    Label10.Caption = "Eliminando marcajes repeticion "
    Label11.Caption = ""
    Me.Refresh
    
    'CadInci = DevuelveDesdeBD("repeticion", "Empresas", "idEmpresa", "1", "N")
    
    If vEmpresa.Repeticion > 0 Then
        'Obtenemos la fecha mas baja
        Set RFin = New ADODB.Recordset
        If Opcion <> 2 Then
            CadInci = "Select min(fecha) from EntradaFichajes WHERE Secuencia >= " & PrimeraInsercion
            RFin.Open CadInci, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            CadInci = "'1900-01-01'"
            If Not RFin.EOF Then
                If Not IsNull(RFin.Fields(0)) Then CadInci = "'" & Format(RFin.Fields(0), FormatoFecha) & "'"
            End If
            RFin.Close
            CadInci = " Fecha >= " & CadInci
        Else
            
            CadInci = "Fecha >='" & Format(txtFecha(0).Text, FormatoFecha) & "' AND fecha <='"
            CadInci = CadInci & Format(txtFecha(1).Text, FormatoFecha) & "'"
            
            
        End If
        
        'Ya tenemos a partir de k fecha, y con k cadencia vamos a eliminar repetidos
        CadInci = "Select * from Entradafichajes WHERE " & CadInci
        CadInci = CadInci & " ORDER BY idTrabajador,Fecha,Hora"
        RFin.Open CadInci, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        PrimeraInsercion = 0 'Tendremos el codigo del trabajador
        CadInci = "DELETE from EntradaFichajes WHERE Secuencia = "
        While Not RFin.EOF
            If RFin!idTrabajador <> PrimeraInsercion Then
                Label11.Caption = "Trabajador: " & RFin!idTrabajador
                Label11.Refresh
                'Nuevo trabajador
                PrimeraInsercion = RFin!idTrabajador
                Fecha = RFin!Fecha
                Hora = Format(RFin!Hora, "hh:mm:ss")
            Else
                'Es el mismo trabajador.
                'Veamos la fecha
                If RFin!Fecha <> Fecha Then
                    Fecha = RFin!Fecha
                    Hora = Format(RFin!Hora, "hh:mm:ss")
                Else
                    'MISMO TRABAJADOR , MISMA FECHA
                    ContFich = DateDiff("n", Hora, Format(RFin!Hora, "hh:mm:ss"))
                    If ContFich >= vEmpresa.Repeticion Then
                        'Las horas se diferencian. NO elimino
                        Hora = Format(RFin!Hora, "hh:mm:ss")
                    Else
                        'SI elimino
                        conn.Execute CadInci & RFin!Secuencia
                    End If
                End If
            End If
            'Siguiente
            RFin.MoveNext
        Wend
        RFin.Close
    
    End If  'Eliminacion marcajes repetidos



'Si  SoloRectifica entonces ya salimos
If SoloRectifica Then Exit Function


Label10.Caption = ""

'--------------------





'-------------------
'Borramos el temporal
RIni.CursorType = 0
conn.Execute "Delete from TemporalFichajes"
Set RIni = Nothing



TraspasaTemporal = CuantosErrores2

Exit Function
ErrTraspaso:
    EscribeErrorLinea "Error traspasando lineas : " & Err.Description
    TraspasaTemporal = -1
End Function


Private Sub InsertaHCo(ByRef SQ As String)
On Error Resume Next
    conn.Execute SQ
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Load()
PrimeraVez = True
Me.FrameKimalid.Visible = (Opcion = 2)
If vEmpresa.reloj = vbTCP3 Then Label5.Caption = "Procesando fichero"
Label10.Caption = ""
Label11.Caption = ""
If Opcion = 2 Then
    'Buscar cual es la fecha mas pequeña a tratar
    PonerFechaPequeña
End If
End Sub






'A partir de la tabla TipoAlzicoop generaremos los valores
'de la tabla fichajes
Private Sub GeneraEntradasALZ()
Dim RTarj As ADODB.Recordset
Dim RFech As ADODB.Recordset
Dim vSec As Long
Dim Cod As Long
Dim SQL As String
Dim Rc As Byte
Dim HayMarcajes As Long
Dim miM As CMarcajes


vSec = 1
SQL = "Select distinct Tarjeta from TipoAlzicoop"
Set RTarj = New ADODB.Recordset
Set RFech = New ADODB.Recordset
RTarj.Open SQL, conn, , , adCmdText
'Obtenemos el numero de scuencia
vSec = ObtnerNumSecuenciaEntradaMarcajes

Label10.Caption = "Creando marcajes: "
Label10.Refresh
'Para cada tarjeta
While Not RTarj.EOF
    'FALTA###
    'Cod = DevuelveNumTrabajador(RTarj.Fields(0))
    If Cod > 0 Then
        'Ahora veremos de cuantas fecha tiene marcajes
        ' y procesaremos para cada fecha
        SQL = "Select distinct fecha from TipoAlzicoop WHERE Tarjeta='" & RTarj.Fields(0) & "'"
        SQL = SQL & " ORDER BY Fecha"
        RFech.Open SQL, conn, , , adCmdText
        While Not RFech.EOF
            '-----------------------------------------------------
            'Comprobamos si existen ya marcajes para esos valores
            HayMarcajes = YaExistenMarcajes(CInt(Cod), RFech.Fields(0))
            Rc = vbYes
            If HayMarcajes > 0 Then
                SQL = "Ya existen marcajes para el trabajador cod: " & Cod & "   y fecha: " & RFech.Fields(0) & vbCrLf
                SQL = SQL & "  ¿ Quiere eliminar el antiguo marcaje. ?" & vbCrLf
                SQL = SQL & "   .- Si --> Eliminamos los antiguos" & vbCrLf
                SQL = SQL & "   .- No --> Dejamos de procesar estos datos" & vbCrLf
                Rc = MsgBox(SQL, vbQuestion + vbYesNo)
                'Si k eliminamos el anterior
                If Rc = vbYes Then
                    Set miM = New CMarcajes
                    If miM.Leer(HayMarcajes) = 0 Then miM.Eliminar
                    Set miM = Nothing
                End If
            End If
            If Rc = vbYes Then
                Label11.Caption = "Tarjeta: " & RTarj.Fields(0) & " - Fecha: " & RFech!Fecha
                Label11.Refresh
                GeneraUnmarcajeAlzicoop RTarj.Fields(0), Cod, RFech!Fecha, vSec
            End If
            RFech.MoveNext
        Wend
        RFech.Close
    End If
    'Movemos la sigueinte
    RTarj.MoveNext
Wend
RTarj.Close
Set RTarj = Nothing
'Borramos temporal
Label11.Caption = ""
Label11.Refresh
Label10.Caption = ""
Label10.Refresh
Set RFech = Nothing
End Sub


Private Function ObtnerNumSecuenciaEntradaMarcajes() As Long
Dim RS As ADODB.Recordset

ObtnerNumSecuenciaEntradaMarcajes = 1
Set RS = New ADODB.Recordset
RS.Open "Select MAX(Secuencia) from EntradaFichajes", conn, , , adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then _
        ObtnerNumSecuenciaEntradaMarcajes = RS.Fields(0) + 1
End If
RS.Close
Set RS = Nothing
End Function



Private Function VaciarTemporales(ByRef Cad As String) As Byte
Dim NombreTabla As String
Dim RError As ADODB.Recordset
Dim RsIni As ADODB.Recordset
Dim ContError As Long
Dim txtFecha As String
Dim ContInicio As Long


On Error GoTo ErrorVaciarTemporales
VaciarTemporales = 1


txtFecha = Format(Now, "Long Date")
If vEmpresa.reloj = vbAlzira Then
    NombreTabla = "TipoAlzicoop"
Else
    NombreTabla = "TemporalFichajes"
End If

'Abrimos la tabla de errtarjetas
Set RError = New ADODB.Recordset
Set RsIni = New ADODB.Recordset
RsIni.CursorType = adOpenKeyset
RsIni.LockType = adLockOptimistic
RError.CursorType = adOpenKeyset
RError.LockType = adLockOptimistic
'Abrimos los regsitros destino
' el correcto y el erroneo
RError.Open "ErrorTarjetas", conn, , , adCmdTable
If RError.EOF Then
    ContError = 1
    Else
        RError.MoveLast
        ContError = RError!Secuencia + 1
End If
ContInicio = ContError
RsIni.Open NombreTabla, conn, , , adCmdTable
While Not RsIni.EOF
        'Ha habido un error
        RError.AddNew
        RError!Secuencia = ContError
        RError!Fecha = RsIni!Fecha
        RError!Hora = RsIni!Hora
        If vEmpresa.reloj = vbAlzira Then
            RError!idInci = 0
            RError!Numtarjeta = RsIni!Tarjeta
        Else
            RError!idInci = RsIni!idInci
            RError!Numtarjeta = RsIni!Numtarjeta
        End If
        RError!Error = "Se ha quedado sin traspasar. Fecha: " & txtFecha
        RError.Update
        ContError = ContError + 1
        RsIni.MoveNext
Wend

ContInicio = ContError - ContInicio
If ContInicio > 0 Then

    Label11.Caption = "Moviendo marcajes erroneos."
    Label11.Refresh
    
    EscribeErrorLinea "     -Marcajes erroneos"

    RsIni.MoveFirst
    'Utilizaremos nombre tabla como string
    NombreTabla = "       Tarjeta      Fecha       Hora   "
    EscribeErrorLinea NombreTabla
    While Not RsIni.EOF
        'Utilizaremos nombre tabla como string
        If vEmpresa.reloj = vbAlzira Then
            NombreTabla = Mid(RsIni!Tarjeta & "      ", 1, 6) & "    "
        Else
            NombreTabla = Mid(RsIni!Numtarjeta & "      ", 1, 6) & "    "
        End If
        NombreTabla = NombreTabla & Format(RsIni!Fecha, "dd/mm/yyyy") & "    "
        NombreTabla = NombreTabla & Format(RsIni!Hora, "hh:mm")
        EscribeErrorLinea "       " & NombreTabla
        RsIni.MoveNext
    Wend
    
    Else
        VaciarTemporales = 0
        Cad = ""
End If
RsIni.Close
RError.Close
Set RsIni = Nothing
Set RError = Nothing
If vEmpresa.reloj = vbAlzira Then
    
    NombreTabla = "TipoAlzicoop"
Else
    NombreTabla = "TemporalFichajes"
    
End If

conn.Execute "Delete  FROM " & NombreTabla
Exit Function
ErrorVaciarTemporales:
    Cad = "Error vaciando tablas temporales." & vbCrLf _
        & "Puede que alguna tabla temporal no este vacia." & vbCrLf & Err.Description
End Function


'
'Private Sub AbreFichero()
'MiNF = FreeFile
'Open App.Path & "\Error" & Format(Now, "yymmdd") & ".log" For Output As #MiNF
''End Sub
'
'Private Sub CierraFichero(Borrar As Boolean)
'Close #MiNF
'If Borrar Then
'    Kill App.Path & "\Error" & Format(Now, "yymmdd") & ".log"
'    Else
'        MsgBox "Se han producido errores." & vbCrLf & "Consulte el archivo: " & vbCrLf & _
'            "        " & App.Path & "\Error" & Format(Now, "yymmdd") & ".log" & vbCrLf & _
'            " para obtener más información.", vbExclamation
'End If
'End Sub
'
'Private Sub EscribeError(cadena)
'    Print #MiNF, cadena
'End Sub


Private Function NombreEnProcesado(Original As String) As String
Dim Cad As String
Dim Kpath As String
Dim I As Integer
Dim Aux As String
Dim NombreRobotics As String

On Error GoTo errNombreProcesados


Aux = App.Path & "\Procesados"
Cad = Dir(Aux, vbDirectory)
'Si no existe lo creo
If Cad = "" Then MkDir Aux
Kpath = App.Path & "\Procesados"

'------------------------
'Nuevo 3 Noviembre 2003
'------------------------
'   .- Los archivos se moveran a carpetas dentro de procesados
'   k tendran año mes . Esto es, para una fecha 12 -Abril - 2001
'   dentro de mconfig.carpetaprocesados crearemos una 2001_04
Aux = Format(Now, "yyyy") & "_" & Format(Now, "mm")

If Dir(Kpath & "\" & Aux, vbDirectory) = "" Then MkDir (Kpath & "\" & Aux)
Kpath = Kpath & "\" & Aux
NombreRobotics = ""
If vEmpresa.reloj = vbRobotics Then
    'Intentare llevarlo con el mismo nombre, puesto que es improbable que se solapen
    I = InStrRev(Original, "\")
    If I > 0 Then NombreRobotics = Mid(Original, I + 1)

End If

'Cojere la fecha
If NombreRobotics = "" Then NombreRobotics = Format(Now, "yymmdd")

'-- FIN NUEVO
Kpath = Kpath & "\"
I = 0
Do
    Aux = "PR" & NombreRobotics & "." & Format(I, "000")
    I = I + 1
    Cad = Dir(Kpath & Aux)
    If Cad = "" Then
        NombreEnProcesado = Kpath & Aux
        I = -1
    End If
Loop Until I < 0

Exit Function
errNombreProcesados:
    MuestraError Err.Number, "Error creando carpeta procesados. " & vbCrLf & Aux & vbCrLf & vbCrLf & "La aplicacion finalizará"
    End
End Function





Private Sub txtFecha_GotFocus(Index As Integer)
txtFecha(Index).SelStart = 0
txtFecha(Index).SelLength = Len(txtFecha(Index).Text)
End Sub



Private Sub txtFecha_LostFocus(Index As Integer)
If txtFecha(Index) = "" Then Exit Sub
If Not EsFechaOK(txtFecha(Index)) Then
    MsgBox "Fecha incorrecta: " & txtFecha(Index).Text, vbExclamation
    txtFecha(Index).Text = ""
End If

End Sub


Private Sub PonerFechaPequeña()
Dim F1 As Date
Dim SQL As String
Dim RS As ADODB.Recordset
    
    
    On Error GoTo EPonerFechaPequeña
    
    F1 = CDate("01/01/1900")
    Set RS = New ADODB.Recordset
    SQL = "Select Max(Fecha) FROM EntradaFichajes"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then F1 = RS.Fields(0)
    End If
    RS.Close
    
    SQL = "Select Max(Fecha) FROM EntradaMarcajes"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If F1 < RS.Fields(0) Then F1 = RS.Fields(0)
        End If
    End If
    RS.Close
    If F1 <> CDate("01/01/1900") Then
        F1 = DateAdd("d", 1, F1)
        txtFecha(0).Text = Format(F1, "dd/mm/yyyy")
        txtFecha(0).Tag = txtFecha(0).Text
    Else
        F1 = DateAdd("d", -1, Now)
        txtFecha(0).Text = Format(F1, "dd/mm/yyyy")
        txtFecha(0).Tag = ""
    End If
    
    txtFecha(1).Text = txtFecha(0).Text
    Set RS = Nothing
    Exit Sub
EPonerFechaPequeña:
    MuestraError Err.Number
End Sub





'-------------------------------------------------------------------------
' Coje de la tabla de MarcajesKimaldi y para cada trabajador, y fecha, genera
' las entradas en la tabla entradamarcajes para luego procesarlos
'
' Todos los registros de entradafichajes los generaremos a partir de la tabla de kimaldi
'
'
Private Function GeneraMarcajesKimaldi() As Boolean
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim SQL As String
Dim INSE As String
Dim con As Long
Dim Trab As Long
Dim FechaANT As Date
Dim Insertar As Boolean
Dim Hora As Date
Dim CodTarea As String
Dim EsperoSalida As Boolean
Dim NF As Integer


    GeneraMarcajesKimaldi = False
    On Error GoTo EGeneraMarcajesKimaldi
    'Los pasamos a tmpMarcajesKimaldi
    Set RS = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    Label6.Caption = "Creando  tabla intermedia"
    Label6.Refresh
    SQL = "Delete * from tmpMarcajesKimaldi"
    conn.Execute SQL
    Label6.Caption = "Pasando a tabla intermedia"
    Label6.Refresh
    SQL = "Insert into tmpMarcajesKimaldi Select * from MarcajesKimaldi"
    SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
    conn.Execute SQL
    
    
    'Marcar SALIDA MASIVA
    'Si tiene salida masiva meteremos la salida masiva, generando un ticaje
    'para cada trabjador vinculado a esa marca
    ' Recorro los ticajes viendo la tarea k es
    ' Cuando encuentro "SALIDA" , y mientras encuentre trabajadores, inserto
    ' en entradamarcajes
    SQL = "Select Tarjeta from Tareas where Tipo=1"   'Salida masiva
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodTarea = ""
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then CodTarea = RS.Fields(0)
    End If
    RS.Close
    If CodTarea <> "" Then
        'OK, hay una tarea k es ticada masiva de salida
        'Entre las fechas solicitadas. Buscaremos la tarea
        ' y los marcajes k siguen son salidas, y los modificaremos poniendo una S
        SQL = "Select * from tmpMarcajesKimaldi"
        SQL = SQL & "  WHERE (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
        SQL = SQL & " ORDER BY nodo,Fecha, Hora"
        RT.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Trab = -1  'Sera el nodo
        Insertar = False
        While Not RT.EOF
            If Trab <> RT!Nodo Then
                Trab = RT!Nodo
                Insertar = False
            End If
            Label6.Caption = RT.Fields(1) & "    " & RT.Fields(4)
            Label6.Refresh
            'Si no es insertar
            If Insertar Then
                'Si el ticaje empieza por codigo trabajador
                'TALTA
                If Mid(RT!Marcaje, 1, 1) = vbDigitoTrabajadores Then
                        SQL = "UPDATE tmpMarcajesKimaldi SET TipoMens ='S' "
                        SQL = SQL & " WHERE Nodo =" & RT!Nodo
                        SQL = SQL & " AND Fecha  = #" & Format(RT!Fecha, "yyyy/mm/dd") & "#"
                        SQL = SQL & " AND Hora = #" & Format(RT!Hora, "hh:mm:ss") & "#"
                        SQL = SQL & " AND Marcaje='" & RT!Marcaje & "'"
                        conn.Execute SQL
                Else
                    Insertar = False
                End If
            End If
            If Not Insertar Then
                'Estoy buscando las tarea de salida masiva
                If RT!Marcaje = CodTarea Then
                    'A partir de aqui son ticajes masivos de salida
                    Insertar = True
                End If
            End If
            RT.MoveNext
        Wend
        RT.Close
    End If
    
    
    'Eliminando datos de tareas
    SQL = "Select * from tmpMarcajesKimaldi"
    RS.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    con = 0
    While Not RS.EOF
        
        Label6.Caption = "Registro: " & con
        Label6.Refresh
        If Mid(RS!Marcaje, 1, 1) <> vbDigitoTrabajadores Then
            RS.Delete
        Else
            con = con + 1
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    If con = 0 Then
        MsgBox "ninguna entrada en este intervalo", vbExclamation
        Exit Function
    End If
    espera 1
    'AHORA GENEREAREMOS EL FICHERO FICHAJES.txt
   
    SQL = vEmpresa.DirMarcajes & "\" & vEmpresa.NomFich
    NF = FreeFile
    Open SQL For Output As #NF
    'Antes
    SQL = "Select * from tmpMarcajesKimaldi ORDER BY Fecha,Hora"
    'Ahora
    SQL = "SELECT tmpMarcajesKimaldi.Fecha, tmpMarcajesKimaldi.Marcaje, tmpMarcajesKimaldi.Hora, tmpMarcajesKimaldi.TipoMens"
    SQL = SQL & " From tmpMarcajesKimaldi"
    SQL = SQL & " ORDER BY tmpMarcajesKimaldi.Fecha, tmpMarcajesKimaldi.Marcaje, tmpMarcajesKimaldi.Hora;"
    
    
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Dim Entrada As Boolean

    If RS.EOF Then
        '-------------->>>>>
        ' MAL. Nada se ha generado
        MsgBox "Cero entradas. Error", vbExclamation
        RS.Close
        Exit Function
    Else
        Hora = CDate("01/01/1900")
        While Not RS.EOF
            'Ej. linea
            '000320409090532470000021ILOC010........
            'tarjeta
            'vector(2) = Mid(Linea, 1, 5)
            'FECHA
            'vector(0) = "20" & Mid(Linea, 6, 2) & "/" & Mid(Linea, 8, 2) & "/" & Mid(Linea, 10, 2)     'Le añadimos el 20 para que sea 2002
            'Hora
            'vector(1) = Mid(Linea, 12, 2) & ":" & Mid(Linea, 14, 2) & ":" & Mid(Linea, 16, 2)
            'seccion
            'vector(3) = 0
            'tecla
            'vector(4) = 0
            
            If RS!Fecha <> Hora Then
                Hora = RS!Fecha
                INSE = RS!Marcaje
                Entrada = True
            End If
            If INSE <> RS!Marcaje Then
                INSE = RS!Marcaje
                Entrada = True
            End If
            Label6.Caption = RS!Fecha & "    " & RS!Marcaje
            Label6.Refresh
            If Entrada = True Then
                If RS!tipomens = "S" Then
                    Entrada = True
                Else
                    Entrada = False
                End If
                Insertar = True
            
            
            Else
                'SALIDA
                If RS!tipomens = "S" Then
                    Insertar = True
                    Entrada = True
                End If
            End If
            
            
            If Insertar Then
                SQL = Right("0000" & Trim(RS!Marcaje), 5) & Format(RS!Fecha, "yymmdd")
                SQL = SQL & Format(RS!Hora, "hhmmss")
                SQL = SQL & "0002004ARIADNA........"
                Print #NF, SQL
                Insertar = False
            End If
            'Sig
            RS.MoveNext
        Wend
    End If
    Close (NF)
    RS.Close
    Set RS = Nothing
    Set RT = Nothing
    GeneraMarcajesKimaldi = True
    Exit Function
EGeneraMarcajesKimaldi:
    MuestraError Err.Number, Err.Description
End Function

Private Function DevuelveTrabajador(ByRef Texto, ByRef R As ADODB.Recordset) As Long
Dim SQL As String
    DevuelveTrabajador = -1
    SQL = "Select idTrabajador from Trabajadores where numtarjeta = '" & Texto & "';"
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then
        If Not IsNull(R.Fields(0)) Then DevuelveTrabajador = R.Fields(0)
    End If
    R.Close
End Function



Private Function InsertarTicajesTeletaxi()
Dim idTrabajador As Integer
Dim Fecha As Date
Dim RS As ADODB.Recordset
Dim Cad As String
Dim Posterior8Mañana As Boolean
Dim Anterior8Mañana As Boolean
Dim Posterior8Tarde As Boolean
Dim Anterior8Tarde As Boolean
Dim UltFecha As Date
Dim Tiene0000 As Boolean
Dim Tiene2359 As Boolean


  
    idTrabajador = -1
    
    Set RS = New ADODB.Recordset
    
    Label11.Caption = "Ajustando nocturnos"
    Me.Refresh
    
    
    Cad = "select max(fecha) from entradafichajes"
    RS.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cad = ""
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Cad = RS.Fields(0)
    End If
    RS.Close
    If Cad = "" Then
        MsgBox "Error leyendo en entradafichajes para el ajuste nocturno", vbExclamation
        Exit Function
    End If
    UltFecha = CDate(Cad)
    espera 0.2
    
    
    Cad = "Select max(secuencia) from entradafichajes"
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ContadorSecuencia = 0
    If Not RS.EOF Then ContadorSecuencia = DBLet(RS.Fields(0), "N")
    ContadorSecuencia = ContadorSecuencia + 1
    RS.Close
    espera 0.1
    
    Cad = "Select * from entradafichajes order by idtrabajador,fecha,hora"
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
           
    
        If RS!idTrabajador <> idTrabajador Then
            If idTrabajador > 0 Then
                'Comprobamos si insertamos
                '----------------------------------
                If Fecha <> UltFecha Then
                    'Si tiene el ticaje a las 6 y , o bien no ha ticado luego, o el ticaje
                    'es posterior a las 8 de la mañana
                
                
                        If Anterior8Mañana And (Not Posterior8Mañana Or Not Anterior8Tarde) Then
                            'INSERTAMOS EL DE LAS 00:00
                            If Not Tiene0000 Then InsertaNocturno idTrabajador, Fecha, False
                            
                        End If
                        
                        If Posterior8Tarde And (Posterior8Mañana And Not Anterior8Tarde) Then
                            'INSERTAMOS el de las 23:59
                            If Not Tiene2359 Then InsertaNocturno idTrabajador, Fecha, True
                        End If
                
                End If
            End If
            'Nuevo trabajador.
            Fecha = "0:00:00"
            
            idTrabajador = RS!idTrabajador
            'label catpiton
            Label11.Caption = "Ajuste nocturno: " & idTrabajador
            Label11.Refresh
        End If
    
        If Fecha <> RS!Fecha Then
                If Val(Fecha) <> 0 Then
                    If Fecha <> UltFecha Then
                        'Si tiene el ticaje a las 6 y , o bien no ha ticado luego, o el ticaje
                        'es posterior a las 8 de la mañana
                    
                    
                        If Anterior8Mañana And (Not Posterior8Mañana Or Not Anterior8Tarde) Then
                            'INSERTAMOS EL DE LAS 00:00
                            If Not Tiene0000 Then InsertaNocturno idTrabajador, Fecha, False
                            
                        End If
                        
                        If Posterior8Tarde And (Posterior8Mañana And Not Anterior8Tarde) Then
                            'INSERTAMOS el de las 23:59
                            If Not Tiene2359 Then InsertaNocturno idTrabajador, Fecha, True
                        End If
                    
                    End If
                End If
            
            Posterior8Mañana = False
            Anterior8Tarde = False
            Anterior8Mañana = False
            Posterior8Tarde = False
            Tiene0000 = False
            Tiene2359 = False
            Fecha = RS!Fecha
            
        End If
        
        
        If RS!Hora < HoraM Then
            
            If RS!Hora = "0:00:00" Then Tiene0000 = True
            Anterior8Mañana = True
        Else
            Posterior8Mañana = True
            If RS!Hora < HoraT Then
                Anterior8Tarde = True
            Else
                Posterior8Tarde = True
                If RS!Hora > "23:58:59" Then Tiene2359 = True
               
            End If
        End If
                
        





        'Sig
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Para el ultimo k hemos procesado
    If idTrabajador <> -1 Then
    
        If Val(Fecha) > 0 Then
             If Fecha <> UltFecha Then
                If Anterior8Mañana And (Not Posterior8Mañana Or Not Anterior8Tarde) Then
                    'INSERTAMOS EL DE LAS 00:00
                    If Not Tiene0000 Then InsertaNocturno idTrabajador, Fecha, False
                    
                End If
                
                If Posterior8Tarde And (Posterior8Mañana And Not Anterior8Tarde) Then
                    'INSERTAMOS el de las 23:59
                    If Not Tiene2359 Then InsertaNocturno idTrabajador, Fecha, True
                End If
            End If
        End If
    End If

End Function


Private Sub InsertaNocturno(idTrabajador As Integer, ByRef Fecha As Date, v2359 As Boolean)
Dim Cad As String
    
        Cad = "INSERT INTO Entradafichajes(Secuencia,idTrabajador,Fecha,Hora,HoraReal,idinci) VALUES ("
        Cad = Cad & ContadorSecuencia & "," & idTrabajador & ",'" & Format(Fecha, FormatoFecha) & "'"
        If v2359 Then
            'Es el de las 23:59
            Cad = Cad & ",'23:59:59','23:59:59'"
        Else
            Cad = Cad & ",'00:00:00','00:00:00'"
        End If
        Cad = Cad & ",0)"
        conn.Execute Cad
        ContadorSecuencia = ContadorSecuencia + 1
        'Debug.Print cad
End Sub



