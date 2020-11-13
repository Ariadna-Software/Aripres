VERSION 5.00
Begin VB.Form frmProcesarEntradasMarcajes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso generacion marcajes"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   Icon            =   "frmProcesarEntradasMarcajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      ToolTipText     =   "Obtener siguiente dia proceso"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   6
      Top             =   1545
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   1545
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmProcesarEntradasMarcajes.frx":6852
      ToolTipText     =   "Buscar fecha"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Primera fecha pendiente"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Generar marcajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2505
   End
   Begin VB.Label Label2 
      Caption         =   "Ultima fecha procesada"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmProcesarEntradasMarcajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1

Dim SQL As String

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    HacerProceso
    Screen.MousePointer = vbDefault
End Sub

Private Sub HacerProceso()
Dim Co As Collection
Dim MasDeUnDiaDiferencia As Boolean
Dim UltimoDia As Date
Dim SeHaHechoPregunta As Boolean
Dim F As Date



    If Text1(0).Text = "" Then Exit Sub
    Set miRsAux = New ADODB.Recordset
    
    SeHaHechoPregunta = False
    
    
    'Que de todas las fechas que haya en la entrada fichajes, esta es la menor
    SQL = "Select min(fecha) from entradafichajes"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not miRsAux.EOF Then
        If CDate(Text1(0).Text) > miRsAux.Fields(0) Then SQL = "La fecha para procesar es mayor que la primera pendiente de procesar: " & miRsAux.Fields(0)
    End If
    miRsAux.Close
    If SQL <> "" Then
        If vEmpresa.TodosLosDias Then
            MsgBox "El tipo de proceso de datos(generacion automatica) no permite procesar dias posteriores.", vbCritical
            Exit Sub
        Else
            SQL = SQL & vbCrLf & vbCrLf & "Deberia empezar por la primera fecha a procesar." & vbCrLf & vbCrLf & "       ¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
            SeHaHechoPregunta = True
        End If
    End If
    
    'QUE LA FECHA ES LA CORRECTA, mayor que ultima precoesada....
    
    SQL = "Select max(fecha) from marcajes"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    MasDeUnDiaDiferencia = False
    If Not miRsAux.EOF Then
        If IsNull(miRsAux.Fields(0)) Then
            UltimoDia = DateAdd("d", -1, CDate(Text1(0).Text))
        Else
            UltimoDia = miRsAux.Fields(0)
            'Hay procesados
            If CDate(Text1(0).Text) <= miRsAux.Fields(0) Then
                SQL = "La fecha para procesar es menor o igual a la ultima ya procesada." & vbCrLf & "Desea continuar?"
            Else
                If DateDiff("d", miRsAux.Fields(0), CDate(Text1(0).Text)) > 1 Then
                    'Separacion de mas de un dia
                    MasDeUnDiaDiferencia = True
                    
                End If
            End If
        
        End If
        
    Else
        UltimoDia = DateAdd("d", -1, CDate(Text1(0).Text))
    End If
    miRsAux.Close
    If SQL <> "" Then
        SeHaHechoPregunta = True
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    
    
    
    'Veremos que todos los trabajadores que tienen ticajes tienen HORARIO
    'Para ese dia
    Set Co = New Collection
    ListadoTrabajadoresConErroresPrimarios Co, 1
    If Co.Count > 0 Then
        SQL = "Los siguientes trabajadores no tienen horario asignado:" & vbCrLf
        While Co.Count > 0
            SQL = SQL & Co.Item(Co.Count)
            Co.Remove Co.Count
        Wend
        SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar igualmente?"
        SeHaHechoPregunta = True
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    Set Co = Nothing
    
    
    
    
    
    'Comprobaremos que el numero de ticajes para cada
    Set Co = New Collection
    ListadoTrabajadoresConErroresPrimarios Co, 0
    If Co.Count > 0 Then
        
        SQL = "Los siguientes trabajadores tienen  marcajes impares:" & vbCrLf
        While Co.Count > 0
            SQL = SQL & Co.Item(Co.Count)
            Co.Remove Co.Count
        Wend
        Set Co = Nothing
        SeHaHechoPregunta = True
        SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar igualmente?"
        SQL = MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2)
        If CByte(SQL) <> vbYes Then
            If CByte(SQL) = vbNo Then
                If MsgBox("Quiere ver marcajes del dia " & Text1(0).Text & "?", vbQuestion + vbYesNo) = vbYes Then
            
                    frmTareaActual.QueFecha = CDate(Text1(0).Text)
                    frmTareaActual.Opcion = 1
                    frmTareaActual.Show vbModal
                End If
            End If
            Exit Sub
        End If
        
        
    End If
    Set Co = Nothing
    
    
    If vEmpresa.QueEmpresa = 4 Then
        'En Catadau si hay marcajes que no son 4 al dia avisar
        SQL = "fecha =" & DBSet(Text1(0).Text, "F") & " AND 1"
        SQL = DevuelveDesdeBD("idtrabajador", "entradafichajes", SQL, "1 GROUP BY idtrabajador HAVING count(*)<4")
        
        If SQL <> "" Then
            SeHaHechoPregunta = True
            SQL = "Trabajadores con menos de 4 fichajes"
            SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar igualmente?"
            SQL = MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2)
            If CByte(SQL) <> vbYes Then
                If CByte(SQL) = vbNo Then
                    If MsgBox("Quiere ver marcajes del dia " & Text1(0).Text & "?", vbQuestion + vbYesNo) = vbYes Then
                
                        frmTareaActual.QueFecha = CDate(Text1(0).Text)
                        frmTareaActual.Opcion = 1
                        frmTareaActual.Show vbModal
                    End If
                End If
                Exit Sub
            End If
        End If
    
    End If
    
    
    If vEmpresa.Reloj2 > 0 Then
        'MAS DE UN RELOJ. Coopic.
        'Veamos que las entradas salidas se producen sobre un mismo reloj
        If TieneErroresEntradaSalidaDiferentesRelojes Then
            SeHaHechoPregunta = True
            SQL = "Trabajadores con entradas /salidas en diferentes relojes"
            SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar igualmente?"
            SQL = MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2)
            If CByte(SQL) <> vbYes Then
                If CByte(SQL) = vbNo Then
                    If MsgBox("Quiere ver marcajes del dia " & Text1(0).Text & "?", vbQuestion + vbYesNo) = vbYes Then
                
                        frmTareaActual.QueFecha = CDate(Text1(0).Text)
                        frmTareaActual.Opcion = 1
                        frmTareaActual.Show vbModal
                    End If
                End If
                Exit Sub
            End If
        End If
    End If
    
    
    'Pregunta
    If Not SeHaHechoPregunta Then
        If vEmpresa.QueEmpresa <> 2 Then
            SQL = "Seguro que desea continuar con el proceso?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    
    
    'LLegado aqui veremos empezaremos el proceso de revision de marcajes
    '----------------------------------------------------------------------
    Label4.Caption = "Comienzo proceso"
    
    Me.Command1.Visible = False
    Screen.MousePointer = vbHourglass
    
    'Abril 2015
    'Aqui guardara de que trabajadores, no vamos a descontar paradas
    conn.Execute "Delete from tmpcombinada WHERE codusu = " & vUsu.Codigo
    
    
    If vEmpresa.QueEmpresa = vbAlzira Then
        CadenaDesdeOtroForm = ""
        frmPrevioProcesar.Modificar = True
        frmPrevioProcesar.Fecha = CDate(Text1(0).Text)
        frmPrevioProcesar.Show vbModal
        If CadenaDesdeOtroForm = "" Then Exit Sub
        Set miRsAux = New ADODB.Recordset
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    If vEmpresa.TodosLosDias Then
        If MasDeUnDiaDiferencia Then
            
            'Procesar Dias Que supuestamente seran festivos (finde etc)
            'Desde la ultima fecha +1 hasta el dia de procesar
            F = DateAdd("d", 1, UltimoDia)
            Do
                DoEvents
                SQL = CStr(F)
                GeneraEntradasSinMarcajes SQL, Label4, Label5
                F = DateAdd("d", 1, F)
            Loop Until F = CDate(Text1(0).Text)
        End If
        
    End If
        
    
    
    
    'ProcesarEntradasFichajes2 CDate(Text1(0).Text), 0, Label4, Label5
    'Rectifica horarios y valida c  alendario
    ProcesarEntradasFichajes CDate(Text1(0).Text), Label4, Label5
    
    Me.Refresh
    
    
    'Segun el tipo de control haremos unas cosas u otras
    GeneraEntradasMarcajes CDate(Text1(0).Text), Label4, Label5
    
    
    
        
    
    'Si alguno no ha fichado
    'Le generamos en vacio
    If vEmpresa.TodosLosDias Then
    
            'Los que no han ticado
            Label4.Caption = ""
            Label5.Caption = ""
            Me.Refresh
            DoEvents
            SQL = CDate(Text1(0).Text)
            GeneraEntradasSinMarcajes SQL, Label4, Label5
    
    
            '    GeneraLasVacaciones
            F = DateAdd("d", 1, UltimoDia)
            Label4.Caption = ""
            Label5.Caption = ""
            Me.Refresh
            Do
                DoEvents
                SQL = CStr(F)
                GeneraLosQueNoHanTicado SQL, Label4, Label5
                F = DateAdd("d", 1, F)
            Loop Until F >= CDate(Text1(0).Text)
    End If
    
    Me.cmdAceptar.Enabled = False
    Me.Command1.Visible = True
    Screen.MousePointer = vbDefault
    Label4.Caption = ""
    Label5.Caption = ""
    
End Sub

'vOpcion
'   0.- Ticajes Impares
'   1.- Trabajadores sin horario


Private Sub ListadoTrabajadoresConErroresPrimarios(ByRef Cole As Collection, vOpcion As Byte)
Dim L As Collection



    Select Case vOpcion
    Case 0
        SQL = "select idtrabajador,(count(*) % 2) as c1 from entradafichajes where fecha='" & Format(Text1(0).Text, FormatoFecha)
        SQL = SQL & "' group by idtrabajador having c1>0"
        
    Case 1
        'Listado de trabajadores sin asignar horario para esta fecha
        Dim CalendarioVariable As Boolean
        CalendarioVariable = vEmpresa.QueEmpresa <> 2 And vEmpresa.QueEmpresa <> 5    'alzira y picassent
        If vEmpresa.QueEmpresa = 4 Then CalendarioVariable = False
        If vEmpresa.QueEmpresa = 1 Then CalendarioVariable = False  'Los que no tienen nada "LIBRE, de momento2
        If CalendarioVariable Then
            
            SQL = "select entradafichajes.idtrabajador,idhorario as c1 from entradafichajes left join calendariot"
            SQL = SQL & " on entradafichajes.idTrabajador = calendariot.idTrabajador"
            SQL = SQL & " and calendariot.fecha=entradafichajes.fecha"
            SQL = SQL & " where entradafichajes.fecha='" & Format(Text1(0).Text, FormatoFecha)
            SQL = SQL & "' group by idtrabajador Having c1 Is Null"
            
        Else
            SQL = "select entradafichajes.idtrabajador,idhorario as c1 from entradafichajes left join"
            SQL = SQL & " calendariol on calendariol.fecha=entradafichajes.fecha "
            SQL = SQL & " where entradafichajes.fecha='" & Format(Text1(0).Text, FormatoFecha)
            SQL = SQL & "' group by idtrabajador Having c1 Is Null"
        End If
    End Select
        
        
    Set miRs = New ADODB.Recordset
    miRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set L = New Collection
    While Not miRs.EOF
        L.Add Val(miRs!idTrabajador)
        
        'Sig
        miRs.MoveNext
    Wend
    miRs.Close
    
    'Ya tengo
    NumRegElim = 0
    While L.Count > 0
    
        Select Case vOpcion
        Case 0
                SQL = "Select entradafichajes.*,concat(hora,'') LHora ,nomtrabajador from entradafichajes,trabajadores where"
                SQL = SQL & " entradafichajes.idtrabajador=trabajadores.idtrabajador and fecha='" & Format(Text1(0).Text, FormatoFecha)
                SQL = SQL & "' and entradafichajes.idtrabajador =" & L.Item(L.Count)
                SQL = SQL & " ORDER BY hora"
                miRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRs.EOF Then
                    SQL = miRs!nomtrabajador & vbCrLf & "             .- "
                    While Not miRs.EOF
                        SQL = SQL & miRs!LHora & "   "
                        miRs.MoveNext
                    Wend
                    SQL = SQL & vbCrLf
                    Cole.Add SQL
                    
                End If
                miRs.Close
        
        Case 1
        
                SQL = "Select * from trabajadores where"
                SQL = SQL & " idtrabajador =" & L.Item(L.Count)
                miRs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRs.EOF Then
                    SQL = miRs!nomtrabajador & "(" & L.Item(L.Count) & ")"
                Else
                    SQL = "Desconocido (" & L.Item(L.Count) & ")"
                End If
                SQL = SQL & vbCrLf
                Cole.Add SQL
                    
                
                miRs.Close
        
        
        End Select
        L.Remove L.Count
    Wend
    
    Set L = Nothing
    Set miRs = Nothing
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    ObtenerSiguienteFechaProceso
    If Text1(0).Text <> "" Then Me.cmdAceptar.Enabled = True
End Sub

Private Sub Form_Activate()
    If SQL = "" Then
        DoEvents
        CargaValoresInciales
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    SQL = ""
End Sub

Private Sub frmc_Selec(vFecha As Date)
    Text1(CInt(Me.imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim Obj As Object

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

    imgFec(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(Index).Text <> "" Then frmc.NovaData = Text1(Index).Text
    ' ********************************************

    frmc.Show vbModal
    Set frmc = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(0).Tag)) '<===
    ' ********************************************
End Sub



Private Sub CargaValoresInciales()
Dim F As Date
Dim B As Boolean
    On Error GoTo ECargaValoresInciales
    
    Set miRsAux = New ADODB.Recordset
    SQL = "Select max(fecha) from marcajes"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    F = CDate("01/01/1900")
    B = False
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            F = miRsAux.Fields(0)
            B = True
        End If
    End If
    miRsAux.Close
    
    If B Then
        Text1(1).Text = Format(F, "dd/mm/yyyy")
       ' Text1(0).Text = Format(DateAdd("d", 1, F), "dd/mm/yyyy")
    End If
    
    SQL = "Select min(fecha) from entradafichajes"
    
    SQL = SQL & " WHERE fecha >= " & DBSet(vEmpresa.FechaInicio, "F")
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    B = False
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            F = miRsAux.Fields(0)
            B = True
        End If
    End If
    miRsAux.Close
    
    If B Then
        Text1(2).Text = Format(F, "dd/mm/yyyy")
         Text1(0).Text = Format(F, "dd/mm/yyyy")
    End If
    Set miRsAux = Nothing
            
    
    
    Set miRsAux = Nothing
    Exit Sub
ECargaValoresInciales:
    MuestraError Err.Number
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
            If Text1(Index).Text = "" Then Exit Sub
            If Not EsFechaOK(Text1(Index)) Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
End Sub


Private Sub ObtenerSiguienteFechaProceso()

    Set miRsAux = New ADODB.Recordset
    SQL = "Select min(fecha) from entradafichajes where fecha>'" & Format(Text1(0).Text, FormatoFecha) & "' group by fecha"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then Text1(0).Text = Format(miRsAux.Fields(0), "dd/mm/yyyy")
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Me.Command1.Visible = False
    
    
End Sub


Private Function TieneErroresEntradaSalidaDiferentesRelojes() As Boolean
Dim B As Boolean
Dim Par As Boolean
Dim Rel As Integer

    On Error GoTo eTieneErroresEntradaSalidaDiferentesRelojes
    TieneErroresEntradaSalidaDiferentesRelojes = False
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    NumRegElim = -1
    SQL = "SELECT * from entradafichajes where fecha= " & DBSet(Text1(0).Text, "F") & " ORDER BY idtrabajador ,hora"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        B = True
        While B
            If miRsAux!idTrabajador <> NumRegElim Then
                NumRegElim = miRsAux!idTrabajador
                Par = False
            End If
            If Par Then
                If Rel <> miRsAux!Reloj Then
                    B = False
                    TieneErroresEntradaSalidaDiferentesRelojes = True
                End If
                Par = False
            Else
                Rel = miRsAux!Reloj
                Par = True
            End If
            If B Then
                miRsAux.MoveNext
                B = Not miRsAux.EOF
            End If
        Wend
    End If
    miRsAux.Close
eTieneErroresEntradaSalidaDiferentesRelojes:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        Set miRsAux = Nothing
        Set miRsAux = New ADODB.Recordset
    End If
End Function
