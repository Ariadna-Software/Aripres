VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHorasProcesadas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Horas procesadas"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12240
   Icon            =   "frmHorasProcesadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Tag             =   "Empresa|N|N|||jornadassemanalesalz|ParaEmpresa||S|"
      Text            =   "Combo1"
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   5
      Tag             =   "Importe|T|S|||jornadassemanalesalz|Observaciones|||"
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   6000
      TabIndex        =   3
      Tag             =   "Importe|N|N|||jornadassemanalesalz|importe|#,##0.00||"
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
      Tag             =   "Fecha|N|N|0||jornadassemanalesalz|tipohoras|0|S|"
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11040
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Tag             =   "Fecha|N|N|||jornadassemanalesalz|idtrabajador|0000|S|"
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
      Width           =   1155
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmHorasProcesadas.frx":000C
      Height          =   6195
      Left            =   240
      TabIndex        =   11
      Top             =   540
      Width           =   11805
      _ExtentX        =   20823
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
      TabIndex        =   10
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   8
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
         TabIndex        =   9
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
      TabIndex        =   12
      Top             =   0
      Width           =   12240
      _ExtentX        =   21590
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
         TabIndex        =   13
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
End
Attribute VB_Name = "frmHorasProcesadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: LAURA +-+-+-+
' +-+- Fecha: 03/03/06 +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' ********************************************************************************

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'Private CadenaConsulta As String
Private cadSeleccion As String
Private CadB As String

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


Private Sub PonerModo(vModo)
Dim B As Boolean
Dim i As Byte

    Modo = vModo
    
    B = (Modo = 2) Or Modo = 0
    If B Then
        Me.lblIndicador.Caption = PonerContRegistros(Me.adodc1)
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To Me.txtAux.Count - 1
        txtAux(i).Visible = Not B
        If i = 0 Or i = 2 Then Text2(i).Visible = Not B
    Next
    Combo1.Visible = Not B
    
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.Visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    'Si estamo mod or insert
    BloquearTxt txtAux(0), (Modo = 4)
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
    Toolbar1.Buttons(6).Enabled = True 'b And Not DeConsulta
    Me.mnNuevo.Enabled = True 'b And Not DeConsulta
    
    B = (B And Not adodc1.Recordset.EOF)
    
    'Modificar
    B = False
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
'
'    MsgBox "En desarrollo", vbExclamation
'    Exit Sub
'
'    CargaGrid 'primer de tot carregue tot el grid
'    CadB = ""
''    '******************** canviar taula i camp **************************
''    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
''        NumF = NuevoCodigo
''    Else
''        NumF = SugerirCodigoSiguienteStr("tareas", "idtarea")
''    End If
''    '********************************************************************
'
'    'Situamos el grid al final
'    AnyadirLinea DataGrid1, adodc1
'
'    anc = DataGrid1.Top
'    If DataGrid1.Row < 0 Then
'        anc = anc + 206
'    Else
'        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
'    End If
'    Limpiar Me
'    txtAux(0).Text = NumF
'    FormateaCampo txtAux(0)
'    txtAux(1).Text = ""
'    Combo1.ListIndex = 0
'
'    LLamaLineas anc, 3
'
'    'Ponemos el foco
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        PonerFoco txtAux(1)
'    Else
'        PonerFoco txtAux(0)
'    End If
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
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass

    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If

    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    
     
    Combo1.ListIndex = 0

    
    LLamaLineas anc, 4

    'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    Combo1.Top = alto
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    txtAux(4).Top = alto
    
    Text2(0).Top = alto
    Text2(2).Top = alto


End Sub


Private Sub BotonEliminar()
Dim Sql As String
'Dim temp As Boolean



    MsgBox "En desarrollo", vbExclamation
    Exit Sub

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Sql) Then Exit Sub

    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub

    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar la tarea?"
    Sql = Sql & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), FormatoCampo(txtAux(0)))
    Sql = Sql & vbCrLf & "Descripción: " & adodc1.Recordset.Fields(1)
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from tareas where idtarea=" & adodc1.Recordset!idtarea
        conn.Execute Sql
        If CadB <> "" Then
            CargaGrid2 True, CadB
            lblIndicador.Caption = PonerContRegistros(Me.adodc1)
        Else
            CargaGrid2 True, ""
            lblIndicador.Caption = ""
        End If
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
Dim i As Integer

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOk Then
'                If InsertarDesdeForm(Me) Then
'                    CargaGrid
'                    'If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'                    If False Then
'                        cmdCancelar_Click
'                        'If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
'                        If Not adodc1.Recordset.EOF Then
'                           '''''''''' adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
'                        End If
'                        cmdRegresar_Click
'                    Else
'                        BotonAnyadir
'                    End If
'                    CadB = ""
'                End If
            End If
           
        Case 4 'MODIFICAR
            If DatosOk Then
'                If ModificaDesdeFormulario(Me) Then
'                    TerminaBloquear
'                    I = adodc1.Recordset.Fields(0)
'                    PonerModo 2
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
'                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
'                End If
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
    
    If CadB <> "" Then lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
   
   PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
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
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
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
                    'Cambio de horas entre alzira y fruxeresa
            
                    frmAlzModificarHorasEmpresa.idTrabajador = adodc1.Recordset!idTrabajador
                    frmAlzModificarHorasEmpresa.Fecha = adodc1.Recordset!Fecha
                    frmAlzModificarHorasEmpresa.TipoHora = adodc1.Recordset!TipoHoras
                    frmAlzModificarHorasEmpresa.Label1(0).Caption = adodc1.Recordset!nomtrabajador
                    frmAlzModificarHorasEmpresa.Label1(3).Caption = adodc1.Recordset!Desctipohora
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
        
        
        
        CadB = DBSet(Now, "F")
        CadB = "fecha < " & CadB & " AND 1"
        CadB = DevuelveDesdeBD("max(fecha)", "jornadassemanalesalz", CadB, "1")
        If CadB <> "" Then
            Modo = 1
            txtAux(1).Text = CadB
            cmdAceptar_Click
        Else
            PonerModo 0
        End If
        

    End If
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

    Combo1.Clear
    If vEmpresa.QueEmpresa = 2 Then
        Combo1.AddItem "Alzicoop"
        Combo1.ItemData(Combo1.NewIndex) = 0
        Combo1.AddItem "Fruxeresa"
        Combo1.ItemData(Combo1.NewIndex) = 1
    Else
        Combo1.AddItem "Coopic"
        Combo1.ItemData(Combo1.NewIndex) = 1
    End If
    CadB = ""
    CargaGrid2 True, "TipoHoras = -1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
'    BotonEliminar
End Sub

Private Sub mnModificar_Click()
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
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If MontaElSelect Then
        Sql = DevuelveSQL(vSQL)
    Else
        Sql = vSQL
    End If
    
    
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, False
    
    'tots = "S|txtAux(0)|T|Código|700|;S|txtAux(1)|T|Nombre|3080|;S|txtAux(2)|T|Pob.|800|;S|btnBuscar(0)|B||0|;S|txtAux2(2)|T|Población|2200|;"
    
    tots = "S|txtAux(0)|T|Trab|850|;S|Text2(0)|T|Nombre|3500|;S|txtAux(1)|T|Fecha|1150|;S|txtAux(2)|T|Tipo|900|;"
    tots = tots & "S|Text2(2)|T|Horas|1500|;S|Combo1|C|Tipo|1600|;S|txtAux(3)|T|Horas|900|;S|txtAux(4)|T|Ajuste|500|;"
    tots = tots & "N|||||;"
    
    arregla tots, DataGrid1, Me
      
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.ScrollBars = dbgAutomatic
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 And KeyAscii = teclaBuscar Then
        KeyAscii = 0
        BotonBuscar
    Else
        KeyPress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    If Index = 0 Then
        PonerFormatoEntero txtAux(Index)
    Else
        If Index = 1 Then
            If Not EsFechaOK(txtAux(Index)) Then txtAux(Index).Text = ""
        End If
    End If
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean

    B = CompForm(Me)
    If Not B Then Exit Function


    'No puede existir el codigo de tarjeta
    Datos = DevuelveDesdeBD("tarjeta", "tareas", "tarjeta", txtAux(2).Text, "N")
    If Datos <> "" Then
        If Modo = 4 Then
            'Si el codigo de tarea NO es el que estamos modificando... MAL
            If Datos = adodc1.Recordset!idtarea Then Datos = ""
        End If
        If Datos <> "" Then
            MsgBox "El codigo de tarjeta signado ya pertenece a otra tarea", vbExclamation
            Exit Function
        End If
    End If
    If Modo = 3 Then
        'Estamos insertando
        'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
        'la N es pa dir que es numeric

         '******************** canviar els arguments de la funcio i el mensage ****************
         Datos = DevuelveDesdeBD("idtarea", "tareas", "idtarea", txtAux(0).Text, "N")
         If Datos <> "" Then
            MsgBox "Ya existe la tarea: " & txtAux(0).Text, vbExclamation
            DatosOk = False
            PonerFoco txtAux(0)
            Exit Function
        End If
        '*************************************************************************************
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


Private Function DevuelveSQL(Sql As String) As String


    DevuelveSQL = "select jornadassemanalesalz.idtrabajador,nomtrabajador,fecha,tipohoras,DescTipoHora,"
    If vEmpresa.QueEmpresa = 2 Then
        DevuelveSQL = DevuelveSQL & " if (paraempresa=1,'Alzicoop','Fruxeresa')"
    Else
        DevuelveSQL = DevuelveSQL & " if (paraempresa=1,'Externo','')"
    End If
    DevuelveSQL = DevuelveSQL & " ,horastrabajadas,Ajuste as ajustadas"
    DevuelveSQL = DevuelveSQL & ",laborable from jornadassemanalesalz,trabajadores,tiposhora where"
    DevuelveSQL = DevuelveSQL & " jornadassemanalesalz.idtrabajador=trabajadores.idtrabajador and jornadassemanalesalz.tipohoras = tiposhora.tipohora"
    
    'Si cambiamos el WHERE este, hay que tener cuidado que la regresar de editar horas busco el ultimo trozo , hasta aqui
    
    If Sql <> "" Then DevuelveSQL = DevuelveSQL & " AND " & Sql
    DevuelveSQL = DevuelveSQL & " ORDER BY jornadassemanalesalz.idTrabajador,fecha,tipohoras,paraempresa"
End Function
