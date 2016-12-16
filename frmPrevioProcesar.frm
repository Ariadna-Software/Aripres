VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrevioProcesar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Previo procesar marcajes"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   15030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Modificar"
      Height          =   495
      Left            =   9720
      TabIndex        =   6
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Continu&ar"
      Height          =   495
      Left            =   12000
      TabIndex        =   5
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   13440
      TabIndex        =   4
      Top             =   8880
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      Top             =   9000
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Secccion- Trabajador"
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   2
      Top             =   9000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   14843
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label0 
      Caption         =   "Martes 99/99/9999"
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
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   8880
      Width           =   4095
   End
End
Attribute VB_Name = "frmPrevioProcesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Fecha As Date

Dim cad As String
Dim I As Long



Private Sub cmdAceptar_Click()
    cad = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Tag = 2 Then cad = cad & "X"
    Next
    
    If cad <> "" Then
        cad = Len(cad)
        cad = "Ha modificado " & cad & " registro(s). Continuar con el proceso? "
        If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        
        
        
        cad = ""
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Tag = 2 Then
                cad = cad & ", (" & vUsu.Codigo & "," & ListView1.ListItems(I).Text & "," & TransformaComasPuntos(ListView1.ListItems(I).SubItems(5)) & ")"
            End If
        Next
        cad = Mid(cad, 2)
        Screen.MousePointer = vbHourglass
        cad = "INSERT INTO tmpcombinada(codusu,IdTrabajador,HR) VALUES " & cad
        conn.Execute cad
        
    End If
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub Command1_Click()
    ListView1_DblClick
End Sub

Private Sub Form_Activate()
    If ListView1.ColumnHeaders.Count < 5 Then
    
        
        CargarColumnas
        
        CargaDatos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Columnas
   Me.Icon = frmMain.Icon
    Me.Label0(0).Caption = Format(Fecha, "dddd") & " " & Format(Fecha, "dd/mmm/yyyy")
End Sub



Private Sub CargarColumnas()
Dim L As Collection
Dim I As Integer
Dim C As ColumnHeader

    ListView1.ColumnHeaders.Clear
    Set L = New Collection
    
    
    
    L.Add "Codigo|1100|"
    
    L.Add "Nombre|2900|"

    L.Add "Secc|900|"
    L.Add "Suma|1100|"
    L.Add "Ajustadas|1100|"
    L.Add "Pa.|600|"
    
    'Puede quitar almuerzo
    L.Add "PuedeQuitarAlm|0|"

    For I = 1 To 8
        L.Add "H" & I & "|800|"
    Next
    
    'TOTAL..... 11 campos
    For I = 1 To L.Count
        Set C = ListView1.ColumnHeaders.Add(, "C" & I)
        C.Text = RecuperaValor(L.Item(I), 1)
        C.Width = RecuperaValor(L.Item(I), 2)
        If I > 2 Then ListView1.ColumnHeaders(I).Alignment = lvwColumnRight
    Next I
    
    
End Sub



Private Function CargaDatos()
    On Error GoTo eCargadatos
    Set miRsAux = New ADODB.Recordset
Dim IT As ListItem
Dim SQL As String

Dim vHora As Integer

Dim RT As ADODB.Recordset

Dim PuedeQuitarParadas As Boolean
Dim Entrada As Boolean
Dim FueraIntervalo_ As Byte  'Sera 0 o 24, dependera
Dim vH As CHorarios
Dim Minutos As Integer
Dim HI As Date
Dim HF As Date
Dim HIAustada As Date
Dim difer As Currency
Dim Horas  As Currency
Dim Ajustadas As Currency

Dim QuitoMeriendaAlmuerzo As Currency
Dim QuitoMeriAlm As Byte '0 No he quitado nada     1. Ya he quitado almuerzo    2. Quito la merienda


Dim InicioHoras As Byte

    ListView1.ListItems.Clear
    InicioHoras = 7

    SQL = "Delete from tmpCombinada where codusu = " & vUsu.Codigo
    conn.Execute SQL
    Set vH = New CHorarios
    '''Sql = "Select entradafichajes.*,nomtrabajador from entradafichajes,trabajadores where entradafichajes.idtrabajador =trabajadores.idtrabajador "
    SQL = "select entradafichajes.idtrabajador,nomtrabajador,control,seccion from entradafichajes,trabajadores"
    SQL = SQL & " where trabajadores.idtrabajador=entradafichajes.idtrabajador and  fecha=" & DBSet(Fecha, "F")
    SQL = SQL & " group by 1 order by "
    If Me.Option1(0).Value Then
        SQL = SQL & " seccion,idtrabajador"
    Else
        SQL = SQL & " idtrabajador"
    End If
    
    
    Set RT = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        
        SQL = "select entradafichajes.idtrabajador,fecha,hour(hora) lahora,minute(hora) minutos,second(hora) segundos "
        SQL = SQL & ",Control,seccion,nomtrabajador from entradafichajes inner join trabajadores t on t.idtrabajador=entradafichajes.idtrabajador"
        SQL = SQL & " AND fecha ='" & Format(Fecha, FormatoFecha) & "' and entradafichajes.idtrabajador=" & RT!idTrabajador & " ORDER BY hora"
        
       
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        
    
        vHora = 0
        QuitoMeriendaAlmuerzo = 0   'Currency de cuanto he quitado
        QuitoMeriAlm = 0  '0 no he quitado nada   1. El almuezro    2 La merienda
        Entrada = True
        Ajustadas = 0
        Horas = 0
        PuedeQuitarParadas = False
        
        Set IT = ListView1.ListItems.Add()
        IT.Text = Format(miRsAux!idTrabajador, "0000")
        IT.SubItems(1) = miRsAux!nomtrabajador
        IT.SubItems(2) = miRsAux!Seccion
        
        'Si el trabjado no tiene el tipo de control 2 entonces NI miramos si quita paradas
        If miRsAux!Control = 2 Then PuedeQuitarParadas = True

          
        If PuedeQuitarParadas Then
             'Veamos el horario para el trabajador, dia
              cad = "calendariol.idcal=trabajadores.idcal and fecha=" & DBSet(miRsAux!Fecha, "F") & " and idtrabajador"
              SQL = "trabajadores.idcal"
              cad = DevuelveDesdeBD("idhorario", "calendariol,trabajadores", cad, CStr(miRsAux!idTrabajador), "N", SQL)
              If Val(cad) = 0 Then Err.Raise 513, , "Error obteniendo horario trabajador: " & miRsAux!idTrabajador
              
              If Val(cad) <> vH.IdHorario Then
                  If vH.Leer(CInt(cad), miRsAux!Fecha, CInt(SQL)) = 1 Then Err.Raise 513, , "Error obteniendo horario nº: " & cad
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
 
        'If It.Text = "0006" Then Stop
    
        SQL = ""
        HF = "0:00:00"
        While Not miRsAux.EOF
        
           
           If vHora < 16 Then   'solo ionserto 16
                   
                   
                   
               If miRsAux!LaHora >= 0 And miRsAux!LaHora <= 23 Then
                   I = miRsAux!LaHora
                   FueraIntervalo_ = 0
               Else
                   FueraIntervalo_ = 24
                   If miRsAux!LaHora < 0 Then Stop  'De momento NO deberia entrar aqui
                   I = miRsAux!LaHora - FueraIntervalo_
               End If
               
               SQL = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
               IT.SubItems(vHora + InicioHoras) = Mid(SQL, 1, 5)
               If FueraIntervalo_ <> 0 Then IT.ListSubItems(vHora + InicioHoras).ForeColor = vbBlue
               
               If Not Entrada Then
                   HF = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                   difer = DateDiff("n", HI, HF)
                   If FueraIntervalo_ > 0 Then difer = difer + 1440
                   
                   Horas = Horas + difer
           
                   'Ajustada
                   If Minutos > 0 Then
                       HF = HoraRectificada(HF, vEmpresa.AjusteSalida, Minutos)
                       difer = DateDiff("n", HIAustada, HF)
                       If FueraIntervalo_ > 0 Then difer = difer + 1440
                   End If
                   Ajustadas = Ajustadas + difer
                       
               
               
               
                Else
                    HI = Format(I, "00") & ":" & Format(miRsAux!Minutos, "00") & ":" & Format(miRsAux!segundos, "00")
                    If Minutos > 0 Then
                        HIAustada = HoraRectificada(HI, vEmpresa.AjusteSalida, Minutos)
                    Else
                        HIAustada = HI
                    End If
                  
                
                End If
               
                Debug.Print HIAustada
                If PuedeQuitarParadas Then
                    If vH.DtoAlm > 0 And FueraIntervalo_ = 0 Then
                        If QuitoMeriAlm = 0 Then
                            'Compruebo si el ticaje es menor que la hora del almuerzo
                            If HIAustada < vH.HoraDtoAlm Then
                                QuitoMeriAlm = 1
                                QuitoMeriendaAlmuerzo = vH.DtoAlm
                            End If
                        End If
                    End If
                    If vH.DtoMer > 0 Then
                        If QuitoMeriAlm < 2 Then
                            
                            
                            If HF > vH.HoraDtoMer Then
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
        
        
        'Horas calculadas y ajustadas, y paradas
        SQL = " "
        vHora = vHora + InicioHoras
        For I = vHora To ListView1.ColumnHeaders.Count - 1
            IT.SubItems(I) = SQL
        Next
        
        
        If Not Entrada Then
            
            IT.SubItems(3) = SQL
            IT.SubItems(4) = SQL
            IT.SubItems(5) = SQL
            IT.Tag = 0
        Else
            IT.Tag = 1
            
            Horas = Round(Horas / 60, 2)
            
            Ajustadas = Round(Ajustadas / 60, 2)
            If QuitoMeriendaAlmuerzo <> 0 Then
                Ajustadas = Ajustadas - QuitoMeriendaAlmuerzo
                IT.SubItems(InicioHoras - 2) = Format(QuitoMeriendaAlmuerzo, FormatoPrecio)
                IT.SubItems(InicioHoras - 1) = QuitoMeriendaAlmuerzo
            Else
                IT.SubItems(InicioHoras - 2) = " "
                IT.SubItems(InicioHoras - 1) = 0
            End If
            IT.SubItems(InicioHoras - 4) = Format(Horas, FormatoPrecio)
            IT.SubItems(InicioHoras - 3) = Format(Ajustadas, FormatoPrecio)
            
            
        End If
        
        
        RT.MoveNext
    Wend
    RT.Close
   
   
   
eCargadatos:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    Set RT = Nothing
    Set vH = Nothing
End Function


Private Sub ListView1_DblClick()
Dim HorasP As Currency
Dim Hor As Currency
Dim Par As Currency
    cad = ""
    CadenaDesdeOtroForm = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Selected Then
            cad = cad & "X"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & Replace(ListView1.ListItems(I).SubItems(1), "|", "") & vbCrLf
            NumRegElim = I   'me guardo cua les el itm a modificar
        End If
    Next
        
    
    If cad = "" Then
        MsgBox "Seleccione trabajadore(s)", vbExclamation
    Else
        FrmVarios.Opcion = 6
        'Parametros frmviarios
        'UNOoMAS|HT|HP|trabajadores
        If Len(cad) = 1 Then
            'Un solo trabajador
           cad = "1|" & ListView1.ListItems(NumRegElim).SubItems(4) & "|" & ListView1.ListItems(NumRegElim).SubItems(5) & "|"
        Else
           cad = Len(cad)
            cad = cad & "||" & ListView1.ListItems(NumRegElim).SubItems(5) & "|"  'Pondra los del ultimo seleccionado
        End If
        cad = cad & CadenaDesdeOtroForm & "|"
        CadenaDesdeOtroForm = ""
        FrmVarios.Parametros = cad
        FrmVarios.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            HorasP = CCur(CadenaDesdeOtroForm)
            
             For I = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(I).Selected Then
                    
                    Hor = ImporteFormateado(ListView1.ListItems(I).SubItems(4))
                    If Trim(ListView1.ListItems(I).SubItems(5)) = "" Then
                        Par = 0
                    Else
                        Par = ImporteFormateado(ListView1.ListItems(I).SubItems(5))
                    End If
                    
                    If Hor + Par < HorasP Then
                        Par = Hor + Par
                        Hor = 0
                    Else
                        Hor = Hor + Par
                        Par = HorasP
                        Hor = Hor - Par
                    End If
                    ListView1.ListItems(I).Bold = True
                    ListView1.ListItems(I).SubItems(4) = Format(Hor, FormatoImporte)
                    ListView1.ListItems(I).SubItems(5) = Format(Par, FormatoImporte)
                    ListView1.ListItems(I).Tag = 2
                End If
            Next
        End If
        
        ListView1.SetFocus
    End If
    
End Sub
