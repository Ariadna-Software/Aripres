VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoAnticipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado pagos banco"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   Icon            =   "frmListadoAnticipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTra 
      Caption         =   "+"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   8040
      TabIndex        =   4
      Tag             =   "Importe|T|S|||pagos|Observaciones|||"
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   5760
      TabIndex        =   2
      Tag             =   "Importe|N|N|||pagos|importe|#,##0.00||"
      Top             =   4920
      Width           =   795
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      ItemData        =   "frmListadoAnticipos.frx":000C
      Left            =   8880
      List            =   "frmListadoAnticipos.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Importe|N|N|||pagos|pagado|||"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      ItemData        =   "frmListadoAnticipos.frx":0010
      Left            =   6600
      List            =   "frmListadoAnticipos.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Importe|N|N|||pagos|tipo||S|"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   4920
      Width           =   3195
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      TabIndex        =   1
      Tag             =   "Traba|N|N|0||pagos|trabajador|0|S|"
      Top             =   4920
      Width           =   1275
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Tag             =   "Fecha|F|N|||pagos|Fecha||S|"
      Top             =   4920
      Width           =   1155
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmListadoAnticipos.frx":0014
      Height          =   6195
      Left            =   240
      TabIndex        =   12
      Top             =   540
      Width           =   11565
      _ExtentX        =   20399
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
      Left            =   9720
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
      Width           =   11925
      _ExtentX        =   21034
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
End
Attribute VB_Name = "frmListadoAnticipos"
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

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

'codigo que tiene el campo en el momento que se llama desde otro formulario
'nos situamos en ese valor
Public CodigoActual As String

Public DeConsulta As Boolean


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1



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
        If i < 2 Then Combo1(i).Visible = Not B
    Next
    If Not B Then BloquearTxt txtAux(2), True
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    DataGrid1.Enabled = B
    
    Me.cmdTra.Visible = Modo = 3 Or Modo = 1
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.Visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    'Si estamo mod or insert
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    If Modo = 4 Then Combo1(0).Locked = True
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
    
    B = (B And Not adodc1.Recordset.EOF) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(11).Enabled = False

End Sub


Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
    

    
    CargaGrid 'primer de tot carregue tot el grid
  
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    Limpiar Me
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    txtAux(1).Text = ""
    Combo1(0).ListIndex = 1
    LLamaLineas anc, 3
       
    'Ponemos el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1)
    Else
        PonerFoco txtAux(0)
    End If
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub


Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "Tipo = -1"
    '*******************************************************************************
    Limpiar Me
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
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
    txtAux(3).Text = DataGrid1.Columns(3).Text
    txtAux(4).Text = DataGrid1.Columns(5).Text
    
    i = adodc1.Recordset!tipo
    PosicionarCombo Combo1(0), i
    
    
    i = adodc1.Recordset!pagado
    PosicionarCombo Combo1(1), i
    
    LLamaLineas anc, 4

    'Como es modificar
    PonerFoco txtAux(3)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    Combo1(0).Top = alto
    Combo1(1).Top = alto
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    txtAux(4).Top = alto
    cmdTra.Left = DataGrid1.Columns(2).Left + 240
    cmdTra.Top = alto
    


End Sub


Private Sub BotonEliminar()
Dim SQL As String
'Dim temp As Boolean
    
    If Modo <> 2 Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub


    On Error GoTo Error2
    'Ciertas comprobaciones
   

    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub

    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el pago seleccionado?"
    SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(1), FormatoCampo(txtAux(1))) & " " & adodc1.Recordset.Fields(2)
    SQL = SQL & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Tipo: " & adodc1.Recordset!descripcion & "  " & adodc1.Recordset!Importe & "€"
    
    If Val(adodc1.Recordset!pagado) = 1 Then SQL = SQL & vbCrLf & vbCrLf & String(30, "*") & vbCrLf & "        P A G A D O"
    
    
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        
        SQL = " Fecha =" & DBSet(adodc1.Recordset!Fecha, "F") & " AND  Tipo = " & adodc1.Recordset!tipo & " AND trabajador = " & adodc1.Recordset!Trabajador
        SQL = "Delete from pagos where " & SQL
        conn.Execute SQL
        
        
        NumRegElim = InStr(1, adodc1.RecordSource, " WHERE ")
        CadB = ""
        If NumRegElim > 0 Then
            CadB = Mid(adodc1.RecordSource, NumRegElim + 7)
            NumRegElim = InStr(1, CadB, " ORDER BY ")
            If NumRegElim = 0 Then
                CadB = " true "
            Else
                CadB = Mid(CadB, 1, NumRegElim - 1)
            End If
            
        Else
            CadB = " true "
        End If

        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        CargaGrid CadB
        lblIndicador.Caption = " " & PonerContRegistros(Me.adodc1)
        
        SituarDataTrasEliminar adodc1, NumRegElim, True
        PonerModoOpcionesMenu
      
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


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    
    MsgBox "buscar"
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub

Private Sub cmdAceptar_Click()
Dim i As Long

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        'If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
           
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    
                    i = InStr(1, adodc1.RecordSource, " WHERE ")
                    CadB = ""
                    If i > 0 Then
                        CadB = Mid(adodc1.RecordSource, i + 7)
                        i = InStr(1, CadB, " ORDER BY ")
                        If i = 0 Then
                            CadB = " true "
                        Else
                            CadB = Mid(CadB, 1, i - 1)
                        End If
                        
                    Else
                        CadB = " true "
                    End If
                    i = adodc1.Recordset.AbsolutePosition
                    PonerModo 2
                    
                    CargaGrid CadB
                    lblIndicador.Caption = PonerContRegistros(Me.adodc1)
                
                    
                    adodc1.Recordset.Move i
                End If
            End If
           
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
                lblIndicador.Caption = " " & PonerContRegistros(Me.adodc1)
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
            CargaGrid CadB
    End Select
    
    PonerModo 2
    
    If CadB <> "" Then lblIndicador.Caption = " " & PonerContRegistros(Me.adodc1)
   
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
        txtAux(1).Text = RecuperaValor(CadB, 1)
        txtAux(2).Text = RecuperaValor(CadB, 2)
        PonerFoco txtAux(3)
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (Modo = 2 Or Modo = 0) Then
        If CadB = "" Then
            lblIndicador.Caption = PonerContRegistros(Me.adodc1)
        Else
            lblIndicador.Caption = "" & PonerContRegistros(Me.adodc1)
        End If
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        PrimeraVez = False
        
        PonerModo 0

    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True

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

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)


    CargaCombos

    
    
    CadB = ""
    CargaGrid "Tipo = -1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    CadB = CadenaDevuelta
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
'    BotonEliminar
End Sub

Private Sub mnModificar_Click()
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
        Case 12
                mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional vSql As String)
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    SQL = DevuelveSQL(vSql)
    
    
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, False
    
    'tots = "S|txtAux(0)|T|Código|700|;S|txtAux(1)|T|Nombre|3080|;S|txtAux(2)|T|Pob.|800|;S|btnBuscar(0)|B||0|;S|txtAux2(2)|T|Población|2200|;"
    
    tots = "S|txtAux(0)|T|Fecha|1100|;S|txtAux(1)|T|Trab.|800|;S|txtAux(2)|T|Nombre|3200|;"
    tots = tots & "S|txtAux(3)|T|Importe|800|;S|Combo1(0)|C|Tipo|1200|;S|txtAux(4)|T|Descripcion|3000|;"
    tots = tots & "S|Combo1(1)|C|Pag.|800|;N|||||;N|||||;"
    
    arregla tots, DataGrid1, Me
      
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.ScrollBars = dbgAutomatic
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 And KeyAscii = teclaBuscar Then
        KeyAscii = 0
 '       btnBuscar_Click (0)
        cmdTra_Click
    Else
        Keypress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim C As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    If Modo = 1 Then Exit Sub
    
    If Index = 1 Then
        C = ""
        If Not PonerFormatoEntero(txtAux(Index)) Then
            txtAux(Index).Text = ""
        Else
            C = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idtrabajador", txtAux(Index).Text)
            If C = "" Then
                MsgBox "No existe el trabajador", vbExclamation
                txtAux(Index).Text = ""
            End If
        End If
        txtAux(2).Text = C
        
    Else
        If Index = 0 Then
            If Not EsFechaOK(txtAux(Index)) Then txtAux(Index).Text = ""
        End If
    End If
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean
    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function

    
    Datos = DevuelveDesdeBD("embargo", "trabajadores", "idTrabajador", txtAux(1).Text, "N")
    If Val(Datos) > 0 Then
        If MsgBox("Trabajador en situacion de embargo. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    If Modo = 3 Then
        Datos = " Fecha =" & DBSet(txtAux(0).Text, "F") & " AND  Tipo = " & Combo1(0).ItemData(Combo1(0).ListIndex) & " AND trabajador"
        
         Datos = DevuelveDesdeBD("fecha", "pagos", Datos, txtAux(1).Text, "N")
         If Datos <> "" Then
            MsgBox "Ya existe el pago con estos datos ", vbExclamation
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
'    C = DevuelveDesdeBD("tarea", "tareasrealizadas", "tarea", CStr(adodc1.Recordset!idtarea), "N")
'    If C <> "" Then
'        MsgBox "Existen tareas realizadas asignadas", vbExclamation
'    Else
'        SepuedeBorrar = True
'    End If
    
End Function


Private Sub Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 And Modo = 2 Then Unload Me  'ESC
    End If
End Sub


' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
 ' WheelHook DataGrid1
End Sub

Private Sub DataGrid1_Lostfocus()
 ' WheelUnHook
End Sub

Private Sub CargaCombos()

    CargaComboDesdeBD Combo1(0), "select idTipopago,Descripcion from tipopago"
    
    Combo1(1).AddItem "Si"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    Combo1(1).AddItem "No"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    
    
End Sub

Private Function DevuelveSQL(SQL As String) As String

    DevuelveSQL = "SELECT Pagos.Fecha, Pagos.Trabajador, Trabajadores.NomTrabajador, "
    DevuelveSQL = DevuelveSQL & "Pagos.Importe, TipoPago.Descripcion, Pagos.Observaciones"
    DevuelveSQL = DevuelveSQL & ",If(Pagado=1,""Si"","""") AS P ,tipo,Pagado"
    DevuelveSQL = DevuelveSQL & " FROM (Pagos INNER JOIN TipoPago ON Pagos.Tipo = TipoPago.idTipopago) INNER JOIN "
    DevuelveSQL = DevuelveSQL & "Trabajadores ON Pagos.Trabajador = Trabajadores.IdTrabajador"
    If SQL <> "" Then DevuelveSQL = DevuelveSQL & " WHERE " & SQL
    DevuelveSQL = DevuelveSQL & " ORDER BY Fecha,Pagos.Trabajador"
End Function
