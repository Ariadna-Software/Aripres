VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHorasProcesadas2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Horas procesadas"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13725
   Icon            =   "frmHorasProcesadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Area|N|N|||jornadassemanalesalz|codarea||S|"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdTra 
      Caption         =   "+"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Empresa|N|N|||jornadassemanalesalz|ParaEmpresa||S|"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   8400
      TabIndex        =   6
      Tag             =   "laborable|T|S|||jornadassemanalesalz|ajuste|||"
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   6480
      TabIndex        =   5
      Tag             =   "Horas|N|N|||jornadassemanalesalz|horastrabajadas|#,##0.00||"
      Top             =   4920
      Width           =   1275
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Tag             =   "TipoHoras|N|N|0||jornadassemanalesalz|tipohoras|0|S|"
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11040
      TabIndex        =   8
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Tag             =   "Trabajador|N|N|||jornadassemanalesalz|idtrabajador|0000|S|"
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||jornadassemanalesalz|Fecha|dd/mm/yyyy|S|"
      Top             =   4920
      Width           =   795
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmHorasProcesadas.frx":000C
      Height          =   6195
      Left            =   240
      TabIndex        =   12
      Top             =   540
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   10927
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11040
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   6840
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
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   1320
      Top             =   7080
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
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
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8760
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
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
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
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
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro"
      Begin VB.Menu mnFIltro1 
         Caption         =   "Selcccionar fecha"
         Index           =   0
      End
      Begin VB.Menu mnFIltro1 
         Caption         =   "Ultimo dia procesado"
         Index           =   1
      End
      Begin VB.Menu mnFIltro1 
         Caption         =   "Utlima semana procesada"
         Index           =   2
      End
      Begin VB.Menu mnFIltro1 
         Caption         =   "Sin filtro"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmHorasProcesadas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1

'Private CadenaConsulta As String
Private cadSeleccion As String
Private CadB As String

Dim CadB_Guardada As String


Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean


Dim Filtro As Integer  '0 Fecha
Dim NumeroSubEmpresa As Integer
Dim NomSubEmpresas As String

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim I As Byte

    Modo = vModo
    
    B = (Modo = 2) Or Modo = 0
    If B Then
        Me.lblIndicador.Caption = PonerContRegistros(Me.adodc1)
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    
    
    For I = 0 To Me.txtAux.Count - 1
        txtAux(I).Visible = Not B
        If I = 0 Or I = 2 Then Text2(I).Visible = Not B
        
        If Not B Then
            If I < 3 Then
                BloquearTxt txtAux(I), Modo = 4
            Else
                BloquearTxt txtAux(I), B
            End If
        End If
    Next
    Combo1(0).Visible = Not B
    Combo1(1).Visible = Not B
    
    If Not B Then Combo1(0).Enabled = Modo = 3 Or Modo = 1
    If Not B Then Combo1(1).Enabled = Modo = 3 Or Modo = 1
    
    
    
    cmdTra.Visible = Not B
    
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.Visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    'Si estamo mod or insert
    
    
End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
Dim B As Boolean

    B = (Modo = 2) Or Modo = 0
    'Búsqueda
    Toolbar1.Buttons(2).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(3).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    'FALTA### bvolver a quitar comentariop
    'If vEmpresa.QueEmpresa = 2 Or vEmpresa.QueEmpresa = 5 Then B = False
    
    Toolbar1.Buttons(6).Enabled = B 'b And Not DeConsulta
    Me.mnNuevo.Enabled = B 'b And Not DeConsulta
    
    B = (B And Not adodc1.Recordset.EOF)
    
    'Modificar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Imprimir
    Toolbar1.Buttons(11).Enabled = True

End Sub


Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single


 '   CargaGrid2 True, "" 'primer de tot carregue tot el grid
 '   CadB = ""
'    '******************** canviar taula i camp **************************
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        NumF = NuevoCodigo
'    Else
'        NumF = SugerirCodigoSiguienteStr("tareas", "idtarea")
'    End If
'    '********************************************************************

    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1

    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    Limpiar Me
    'txtAux(0).Text = NumF
    'FormateaCampo txtAux(0)
    txtAux(1).Text = ""
    Combo1(0).ListIndex = 0
    Combo1(1).ListIndex = IIf(vEmpresa.QueEmpresa = vbAlzira, 1, 0)

    LLamaLineas anc, 3

    'AJUSTE
    '       0.- Sin ajustar
    '       1.- Se ajusto en proceso calculo de horas
    '       2.- Se creo a mano
    '       3.- Se modifico la que  estaba sin ajustar en proc horas
    '       4.- "            " del proceso de calculo de horas
    '       5.- "               la creada a mano
    '       7: TotalMenteManual Se ha cambiado sin respetar sumatorios
    txtAux(4).Text = "2"
    PonerFoco txtAux(0)
   
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid2 True, ""
    PonerModo 2
End Sub


Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid2 True, "ajuste = -1"
    '*******************************************************************************
    Limpiar Me

    LLamaLineas DataGrid1.Top + 206, 1
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer


    On Error GoTo eBotonModificar
    
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass


    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If

    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    
    txtAux(0).Text = DataGrid1.Columns(0).Text
    Text2(0).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(2).Text
    txtAux(2).Text = DataGrid1.Columns(3).Text
    Text2(2).Text = DataGrid1.Columns(4).Text
    
    Combo1(0).ListIndex = -1
    I = Val(adodc1.Recordset!paraempresa)
    PosicionarCombo Combo1(0), I
    If Combo1(0).ListIndex = -1 Then Err.Raise 513, , "Error situando combo Empresa"
    
    Combo1(1).ListIndex = -1
    I = Val(adodc1.Recordset!codArea)
    PosicionarCombo Combo1(1), I
    If Combo1(1).ListIndex = -1 Then Err.Raise 513, , "Error situando combo ALMACEN"
    
    
    I = 7
    txtAux(3).Text = DataGrid1.Columns(I).Text
    txtAux(4).Text = DataGrid1.Columns(I + 1).Text
    

    
    LLamaLineas anc, 4

    'Como es modificar
    PonerFoco txtAux(3)
    
eBotonModificar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        Me.cmdAceptar.Enabled = False
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    cmdTra.Top = alto
    Combo1(0).Top = alto
    Combo1(1).Top = alto
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    txtAux(4).Top = alto
    
    Text2(0).Top = alto
    Text2(2).Top = alto


End Sub


Private Sub BotonEliminar()
Dim SQL As String
'Dim temp As Boolean



    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub

    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub

    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar la linea ?"
    SQL = SQL & vbCrLf & "Trabajador: " & Format(adodc1.Recordset.Fields(0), FormatoCampo(txtAux(0)))
    SQL = SQL & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Dia: " & adodc1.Recordset.Fields(2)
    SQL = SQL & vbCrLf & "Horas " & adodc1.Recordset.Fields(4) & ": " & adodc1.Recordset.Fields(6)
    SQL = SQL & vbCrLf & "Area: " & adodc1.Recordset.Fields(6)
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
      
        SQL = "DELETE FROM jornadassemanalesalz WHERE fecha =" & DBSet(adodc1.Recordset!Fecha, "F")
        SQL = SQL & " AND tipohoras  = " & adodc1.Recordset!TipoHoras
        If vEmpresa.QueEmpresa = vbAlzira Then
            Err.Raise 513, , "NO puede eliminar "
        Else
            SQL = SQL & " AND paraempresa= 0 "
        End If
        SQL = SQL & " AND codarea=  " & adodc1.Recordset!codArea
        SQL = SQL & " AND idTrabajador =" & adodc1.Recordset!idTrabajador
        
        
        
        conn.Execute SQL
        CadB = adodc1.Recordset.Source
        
        CargaGrid2 False, CadB
        lblIndicador.Caption = PonerContRegistros(Me.adodc1)
    
        SituarDataTrasEliminar adodc1, NumRegElim, True
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub

Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub




Private Sub cmdAceptar_Click()
Dim I As Long

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOk Then
                
                If InsertarDesdeForm(Me) Then
                    
                    CadB = CadB_Guardada
                    CargaGrid2 True, CadB
                    CadB = txtAux(1).Text
                    BotonAnyadir
                    Me.txtAux(1).Text = CadB
                End If
            End If
           
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CadB = adodc1.Recordset.Source
                    
                        CargaGrid2 False, CadB
                        lblIndicador.Caption = "" & PonerContRegistros(Me.adodc1)
                    
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
           End If
           
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid2 True, CadB
                PonerModo 2
                lblIndicador.Caption = PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
        End Select
End Sub


Private Sub cmdCancelar_Click()
'    On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst

        Case 4 'MODIFICAR
            TerminaBloquear
        Case 1 'BUSQUEDA
            CargaGrid2 True, CadB
    End Select
    
    PonerModo 2
    
    If CadB <> "" Then lblIndicador.Caption = " " & PonerContRegistros(Me.adodc1)
   
   PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    Cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            Cad = Cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub



Private Sub cmdTra_Click()
    
    
    CadB = "Codigo|idTrabajador|N||15·"
    CadB = CadB & "Nombre|nomtrabajador|T||60·"
    CadB = CadB & "Tarjeta|numtarjeta|T||20·"
    Set frmB = New frmBuscaGrid
    frmB.vTabla = "Trabajadores"
    frmB.vCampos = CadB
    frmB.vDevuelve = "0|1|"
    frmB.vSelElem = 1
    frmB.vTitulo = "TRABAJADORES"
    CadB = ""
    frmB.Show vbModal
    Set frmB = Nothing
    If CadB <> "" Then
        txtAux(0).Text = RecuperaValor(CadB, 1)
        Text2(0).Text = RecuperaValor(CadB, 2)
        PonerFoco txtAux(1)
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
Dim F1 As Date

    If cmdRegresar.Visible Then
        cmdRegresar_Click
    Else
        If Modo = 2 Then
            If Not adodc1.Recordset.EOF Then
            
                If vEmpresa.QueEmpresa = 2 Then
                    'Cambio de horas entre alzira y fruxeresa-motilla
            
                    

            
            
                    frmAlzModificarHorasEmpresa.NombreAreasSubempresa = NomSubEmpresas
                    frmAlzModificarHorasEmpresa.NumeroTotalAreasSubempresa = NumeroSubEmpresa
                    frmAlzModificarHorasEmpresa.idTrabajador = adodc1.Recordset!idTrabajador
                    frmAlzModificarHorasEmpresa.Fecha = adodc1.Recordset!Fecha
                    frmAlzModificarHorasEmpresa.TipoHora = adodc1.Recordset!TipoHoras
                    frmAlzModificarHorasEmpresa.AlmacenArea = adodc1.Recordset!codArea
                    
                    frmAlzModificarHorasEmpresa.lblNombre = adodc1.Recordset!nomtrabajador
                    frmAlzModificarHorasEmpresa.lblTipoHora(0).Caption = adodc1.Recordset!Desctipohora
                    frmAlzModificarHorasEmpresa.lblTipoHora(1).Caption = adodc1.Recordset!Area  'del area
                    frmAlzModificarHorasEmpresa.Show vbModal
                          
                    If CadenaDesdeOtroForm <> "" Then
                        
                        cadSeleccion = adodc1.Recordset!idTrabajador & "|" & adodc1.Recordset!Fecha & "|" & adodc1.Recordset!TipoHoras & "|"
                    
                        'Ha modificado algo
                        'Ahora encesitamos ver las dos horas(fruixeresa y alzicoop)
                        'Con lo cual, si en el WHERE esta paraempresa , monto un where nuevo
                        NumRegElim = InStr(1, adodc1.RecordSource, "= tiposhora.tipohora")
                        If NumRegElim > 0 Then
                            CadB = Mid(adodc1.RecordSource, NumRegElim + 18)
                            'QUito el order by
                            NumRegElim = InStr(1, CadB, "ORDER BY ")
                            CadB = Mid(CadB, 1, NumRegElim - 1)
                            
                            NumRegElim = InStr(1, CadB, "ParaEmpresa")
                            
                            If NumRegElim > 0 Then
                                'Monto un nuevo select
                                CadB = " jornadassemanalesalz.idtrabajador = " & RecuperaValor(cadSeleccion, 1) & " AND fecha = " & DBSet(RecuperaValor(cadSeleccion, 2), "F")
                                CadB = DevuelveSQL(CadB)
                                
                            Else
                                CadB = adodc1.RecordSource
                            End If
                            
                            CargaGrid2 False, CadB
                            
                            'Situamos el grid
                            CadB = "idtrabajador = " & RecuperaValor(cadSeleccion, 1)
                            adodc1.Recordset.Find CadB
                            F1 = RecuperaValor(cadSeleccion, 2)
                            CadB = RecuperaValor(cadSeleccion, 3)
                            cadSeleccion = RecuperaValor(cadSeleccion, 1)
                            NumRegElim = 1
                            Do
                                If adodc1.Recordset.EOF Then
                                    NumRegElim = 2
                                Else
                                    If adodc1.Recordset!idTrabajador <> cadSeleccion Then
                                        NumRegElim = 2
                                    Else
                                        If adodc1.Recordset!Fecha = F1 Then
                                            If adodc1.Recordset!TipoHoras = Val(CadB) Then NumRegElim = 2
                                        End If
                                    End If
                                End If
                                If NumRegElim = 1 Then adodc1.Recordset.MoveNext
                            Loop Until NumRegElim = 2
                            
                            CadB = ""
                        End If 'nnumregelim>0
                    End If 'cadena desde otrofomz<>''
                End If  'que empresa=alzira
            End If  'EOF
        End If  'modo=2
    End If 'de regresar visible
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (Modo = 2 Or Modo = 0) Then
        If CadB = "" Then
            lblIndicador.Caption = PonerContRegistros(Me.adodc1)
        Else
            lblIndicador.Caption = PonerContRegistros(Me.adodc1)
        End If
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        PrimeraVez = False
        
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
        NumeroSubEmpresa = 0
        NomSubEmpresas = ""
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open "Select idSubEmr , NomSubEmpre    FROM areasubempresa", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            NumeroSubEmpresa = NumeroSubEmpresa + 1
            NomSubEmpresas = NomSubEmpresas & Format(miRsAux!idSubEmr, "0000") & miRsAux!NomSubEmpre & "|"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
'
'        CadB = DBSet(Now, "F")
'        CadB = "fecha < " & CadB & " AND 1"
'        CadB = DevuelveDesdeBD("max(fecha)", "jornadassemanalesalz", CadB, "1")
'        If CadB <> "" Then
'            Modo = 1
'            txtAux(1).Text = CadB
'            cmdAceptar_Click
'        Else
'            PonerModo 0
'        End If
        
        mnFIltro1_Click 2
        If Me.adodc1.Recordset.EOF Then
            PonerModo 0
        Else
            PonerModo 2
        End If
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmMain.Icon
'    btnPrimero = 14 'index del botó "primero"
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
        .Buttons(12).Image = 11  'Salir
    End With

    
    CadB = "Select idSubEmr AS id,NomSubEmpre as descripcion FROM areasubempresa "
    CargaComboTabla CadB, Combo1(0), False
    
    
    CadB = "Select codarea as id, descripcion FROM areas  "
    CargaComboTabla CadB, Combo1(1), False
    Filtro = 3
    CadB = ""
    CargaGrid2 True, " false "
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadB = CadenaDevuelta
End Sub

Private Sub frmc_Selec(vFecha As Date)
    CadB = vFecha
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
'    BotonEliminar
End Sub

Private Sub mnFIltro1_Click(Index As Integer)
Dim I As Integer
Dim C As String
Dim F As Date
Dim F2 As Date

    I = 0
    If Modo = 2 Then I = 1
    If Modo = 0 Then I = 1

    If I = 0 Then Exit Sub

    For I = 0 To mnFIltro1.Count - 1
        mnFIltro1(I).Checked = I = Index
        
    Next
    Me.mnFIltro1(0).Caption = "Fecha"
    C = ""
    If Index <> mnFIltro1.Count - 1 Then C = " *" 'Es el ultimo
    Me.mnFiltro.Caption = "Filtro " & C
    
    Select Case Index
    Case 0
           'Pide fecha
        CadB = ""
        Set frmc = New frmCal
   
        frmc.Left = Me.Left + 1200
        frmc.Top = Me.Top
        frmc.NovaData = Now
        frmc.Show vbModal
        Set frmc = Nothing
        If CadB = "" Then CadB = Now
             
        mnFIltro1(Index).Tag = DBSet(CadB, "F")
        mnFIltro1(Index).Caption = "Fecha " & mnFIltro1(Index).Tag
        CadB = ""
    Case 1, 2
        CadB = DevuelveDesdeBD("max(fecha)", "jornadassemanalesalz", "1", "1")
        If CadB = "" Then CadB = Now
        F = CDate(CadB)
        CadB = ""
        F2 = F
        
        If Index = 2 Then
            I = Weekday(F, vbMonday)
            I = I - 1
            F2 = DateAdd("d", -I, F)
            mnFIltro1(Index).Tag = " between " & DBSet(F2, "F") & " AND " & DBSet(F, "F")
        
        Else
            mnFIltro1(Index).Tag = DBSet(F, "F")
        End If
        
        CadB = ""
        
    End Select
    Filtro = Index
    CargaGrid2 True, ""
End Sub

Private Sub mnModificar_Click()
    'If BLOQUEADesdeFormulario2(Me, adodc1) Then BotonModificar
    BotonModificar
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
        Case 11
'                'MsgBox "Imprimir...copiar de l'atre manteniment"
'                printNou
                
                CadB = DBSet(Now, "F")
                CadB = "fecha < " & CadB & " AND 1"
                CadB = DevuelveDesdeBD("max(fecha)", "jornadassemanalesalz", CadB, "1")
                CadenaDesdeOtroForm = CadB
                
                                    
                frmListado.Opcion = 17
                frmListado.Show vbModal

        Case 12
                mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid2(MontaElSelect As Boolean, vSQL As String)
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    CadB_Guardada = vSQL
    If MontaElSelect Then
        SQL = DevuelveSQL(vSQL)
    Else
        SQL = vSQL
    End If
    
    
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, False
    
    'tots = "S|txtAux(0)|T|Código|700|;S|txtAux(1)|T|Nombre|3080|;S|txtAux(2)|T|Pob.|800|;S|btnBuscar(0)|B||0|;S|txtAux2(2)|T|Población|2200|;"
    
    tots = "S|txtAux(0)|T|Trab|850|;S|Text2(0)|T|Nombre|3500|;S|txtAux(1)|T|Fecha|1150|;S|txtAux(2)|T|Tipo|900|;"
    tots = tots & "S|Text2(2)|T|Horas|1500|;S|Combo1(1)|C|Area|1600|;S|Combo1(0)|C|Empresa|1600|"
    tots = tots & ";S|txtAux(3)|T|Horas|900|;S|txtAux(4)|T|Ajuste|500|;"
    tots = tots & "N|||||;N|||||;N|||||;"
    arregla tots, DataGrid1, Me
      
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.ScrollBars = dbgAutomatic
    Me.cmdTra.Left = Text2(0).Left - 30
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 And KeyAscii = teclaBuscar Then
        KeyAscii = 0
        cmdTra_Click
    Else
        KeyPress KeyAscii
    End If
End Sub



    
    'AJUSTE
    '       0.- Sin ajustar
    '       1.- Se ajusto en proceso calculo de horas
    '       2.- Se creo a mano
    '       3.- Se modifico la que  estaba sin ajustar en proc horas
    '       4.- "            " del proceso de calculo de horas
    '       5.- "               la creada a mano
    '       7: TotalMenteManual Se ha cambiado sin respetar sumatorios


Private Sub txtAux_LostFocus(Index As Integer)
Dim C As String
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    If Modo = 1 Then Exit Sub
    
    If Index = 0 Or Index = 2 Then
        C = ""
        If txtAux(Index).Text <> "" Then
            If Not PonerFormatoEntero(txtAux(Index)) Then
                txtAux(Index) = ""
                PonerFoco txtAux(Index)
            Else
                If Index = 0 Then
                    C = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idtrabajador", txtAux(Index).Text)
                Else
                    C = DevuelveDesdeBD("DescTipoHora", "tiposhora", "TipoHora", txtAux(Index).Text)
                End If
                If C = "" Then
                    MsgBox "No existe el " & IIf(Index = 0, "Trabajador", "tipo de hora") & ": " & txtAux(Index).Text, vbExclamation
                    txtAux(Index) = ""
                    PonerFoco txtAux(Index)
                End If
            End If
        End If
        Text2(Index).Text = C
            
    Else
        If Index = 1 Then
            If Not EsFechaOK(txtAux(Index)) Then txtAux(Index).Text = ""
            
        Else
            If Index = 3 Then If Not PonerFormatoDecimal(txtAux(Index), 1) Then txtAux(Index).Text = ""
            
        End If
    End If
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean


    If Modo = 3 Then txtAux(4).Text = "2"

    B = CompForm(Me)
    If Not B Then Exit Function

    If Modo = 3 Then
        Datos = "idtrabajador =" & txtAux(0).Text & " AND fecha =" & DBSet(txtAux(1).Text, "F")
        Datos = Datos & " AND tipohoras = " & txtAux(2).Text
        Datos = Datos & " AND codarea = " & Combo1(1).ItemData(Combo1(1).ListIndex) & " AND paraempresa"
        Datos = DevuelveDesdeBD("horastrabajadas", "jornadassemanalesalz", Datos, Combo1(0).ItemData(Combo1(0).ListIndex))
        If Datos <> "" Then
            MsgBox "Ya existe una entrada (Tra-Dia-Emr-Zona-TipoH)", vbExclamation
            B = False
        Else
            B = True
        End If
    End If
    DatosOk = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Function SepuedeBorrar(ByRef C As String) As Boolean
SepuedeBorrar = False
   
    
End Function


Private Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 And Modo = 2 Then Unload Me  'ESC
    End If
End Sub

'Private Sub printNou()
'    With frmImprimir2
'        .cadTabla2 = "clientes"
'        .Informe2 = "rClientes.rpt"
'        If CadB <> "" Then
'            .cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
'        Else
'            .cadRegSelec = ""
'        End If
'        .cadRegActua = Replace(POS2SF(adodc1, Me), "clientes", "clientes_1")
'        .cadTodosReg = ""
'        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={clientes.ape_raso}|"
'        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomEmpre & "'|"
'        .NumeroParametros2 = 1
'        .MostrarTree2 = False
'        .InfConta2 = False
'        .ConSubInforme2 = False
'
'        .Show vbModal
'    End With
'End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
 ' WheelHook DataGrid1
End Sub

Private Sub DataGrid1_Lostfocus()
 ' WheelUnHook
End Sub


Private Function DevuelveSQL(SQL As String) As String


'    DevuelveSQL = "select jornadassemanalesalz.idtrabajador,nomtrabajador,fecha,tipohoras,DescTipoHora,"
'    If vEmpresa.QueEmpresa = 2 Then
'        DevuelveSQL = DevuelveSQL & " if (paraempresa=1,'Alzicoop','" & "vEmpresa.SegundaEmpresa" & "')"
'    Else
'        DevuelveSQL = DevuelveSQL & " if (paraempresa=1,'Externo','')"
'    End If
'    DevuelveSQL = DevuelveSQL & " ,horastrabajadas,Ajuste as ajustadas"
'    DevuelveSQL = DevuelveSQL & ",laborable from jornadassemanalesalz,trabajadores,tiposhora where"
'    DevuelveSQL = DevuelveSQL & " jornadassemanalesalz.idtrabajador=trabajadores.idtrabajador and jornadassemanalesalz.tipohoras = tiposhora.tipohora"
'
'    'Si cambiamos el WHERE este, hay que tener cuidado que la regresar de editar horas busco el ultimo trozo , hasta aqui
    
    DevuelveSQL = "SELECT jornadassemanalesalz.idtrabajador,nomtrabajador,fecha,tipohoras,DescTipoHora,areas.descripcion area,NomSubEmpre destino"
    DevuelveSQL = DevuelveSQL & " ,horastrabajadas,Ajuste as ajustadas,  "
    DevuelveSQL = DevuelveSQL & " laborable , paraempresa , jornadassemanalesalz.codarea"
    DevuelveSQL = DevuelveSQL & " FROM jornadassemanalesalz,trabajadores,tiposhora,areasubempresa,areas where jornadassemanalesalz.idtrabajador=trabajadores.idtrabajador"
    DevuelveSQL = DevuelveSQL & " and jornadassemanalesalz.tipohoras = tiposhora.tipohora"
    DevuelveSQL = DevuelveSQL & " and jornadassemanalesalz.paraempresa = idSubEmr"
    DevuelveSQL = DevuelveSQL & " and jornadassemanalesalz.codarea = areas.codarea"
        
    
    
    
    If SQL <> "" Then DevuelveSQL = DevuelveSQL & " AND " & SQL
    
        
    DevuelveSQL = DevuelveSQL & CargaSQLFiltro
    
    
    DevuelveSQL = DevuelveSQL & " ORDER BY jornadassemanalesalz.idTrabajador,fecha,tipohoras,areas.codarea,paraempresa"
End Function




Private Function CargaSQLFiltro() As String
Dim C As String

    Select Case Filtro
    Case 0
        CargaSQLFiltro = " fecha = " & mnFIltro1(Filtro).Tag
    Case 1
        CargaSQLFiltro = " fecha = " & mnFIltro1(Filtro).Tag
    Case 2
        CargaSQLFiltro = " fecha  " & mnFIltro1(Filtro).Tag
    Case 3
            CargaSQLFiltro = " TRUE "
    End Select
        CargaSQLFiltro = " AND " & CargaSQLFiltro
End Function
