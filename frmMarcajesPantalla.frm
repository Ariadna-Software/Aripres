VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMarcajesPantalla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor  marcajes"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
   Icon            =   "frmMarcajesPantalla.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIMpr 
      Height          =   375
      Left            =   10680
      Picture         =   "frmMarcajesPantalla.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtTrab 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtDT 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   15
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtFec 
      Height          =   285
      Index           =   1
      Left            =   7320
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtFec 
      Height          =   285
      Index           =   0
      Left            =   7320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtTrab 
      Height          =   285
      Index           =   5
      Left            =   2160
      TabIndex        =   1
      Top             =   555
      Width           =   855
   End
   Begin VB.TextBox txtDT 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   9
      Top             =   555
      Width           =   3375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Nombre"
      Height          =   195
      Index           =   1
      Left            =   9960
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Codigo"
      Height          =   195
      Index           =   0
      Left            =   11160
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   8760
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcajesPantalla.frx":7254
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcajesPantalla.frx":756E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcajesPantalla.frx":7B08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   11160
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6495
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   11456
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
      NumItems        =   0
   End
   Begin VB.Label lblDuplicado 
      Caption         =   "Duplicados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8760
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblFecha 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   6
      Left            =   6480
      TabIndex        =   14
      Top             =   600
      Width           =   420
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   1
      Left            =   7080
      Picture         =   "frmMarcajesPantalla.frx":E36A
      ToolTipText     =   "Buscar fecha"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label lblFecha 
      Caption         =   "Desde"
      Height          =   195
      Index           =   7
      Left            =   6480
      TabIndex        =   13
      Top             =   165
      Width           =   465
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   7080
      Picture         =   "frmMarcajesPantalla.frx":E3F5
      ToolTipText     =   "Buscar fecha"
      Top             =   142
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
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
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblFecha 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   10
      Left            =   1320
      TabIndex        =   11
      Top             =   600
      Width           =   420
   End
   Begin VB.Label lblFecha 
      Caption         =   "Desde"
      Height          =   195
      Index           =   11
      Left            =   1320
      TabIndex        =   10
      Top             =   165
      Width           =   465
   End
   Begin VB.Image imgTra 
      Height          =   255
      Index           =   4
      Left            =   1900
      Top             =   135
      Width           =   255
   End
   Begin VB.Image imgTra 
      Height          =   255
      Index           =   5
      Left            =   1900
      Top             =   570
      Width           =   255
   End
End
Attribute VB_Name = "frmMarcajesPantalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public QuieroVerDatos As String


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim Antiguo As String
Dim cad As String




Private Sub Check1_Click()
    Screen.MousePointer = vbHourglass
    ListView1.ListItems.Clear
    CargarColumnas
    CargaDatos
    Screen.MousePointer = vbDefault
End Sub


Private Function DevuelveHoraBD(Indice As Integer) As String
    If Trim(ListView1.ListItems(NumRegElim).SubItems(Indice)) = "" Then
        DevuelveHoraBD = "null"
    Else
        DevuelveHoraBD = "'" & ListView1.ListItems(NumRegElim).SubItems(Indice) & ":00'"
    End If
End Function

Private Sub cmdImpr_Click()
Dim Anterior As String
Dim DosAnterior As String
Dim N As Integer
Dim Aux As String

    If Me.ListView1.ListItems.Count = 0 Then Exit Sub
    
    conn.Execute "DELETE from  tmpinformehorasmes WHERE codusu =" & vUsu.Codigo
    
    Antiguo = ""
    Anterior = ""
    For NumRegElim = 1 To Me.ListView1.ListItems.Count
        
        'tmpinformehorasmes(codusu,fecha,idTrabajador,Nombre,HT,HN,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15,H16)
        
        
        'idTrabajador,Nombre,HT,HN,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15,H16,codusu,fecha)
        If Me.ListView1.ListItems(NumRegElim).Text <> "" Then
            Anterior = Me.ListView1.ListItems(NumRegElim).Text
            DosAnterior = Me.ListView1.ListItems(NumRegElim).SubItems(1)
        End If
        cad = ", (" & vUsu.Codigo & ","
        If Me.Check1.Value = 0 Then
            cad = cad & DBSet(ListView1.ListItems(NumRegElim).SubItems(2), "F") & "," & Anterior & ","
            cad = cad & DBSet(DosAnterior, "T") & ","
        
        Else
            cad = cad & DBSet(Anterior, "F") & "," & DBSet(ListView1.ListItems(NumRegElim).SubItems(1), "T") & ","
            cad = cad & DBSet(ListView1.ListItems(NumRegElim).SubItems(2), "T") & ","
        End If
        cad = cad & DBSet(ListView1.ListItems(NumRegElim).SubItems(8), "N") & ","
        If Trim(ListView1.ListItems(NumRegElim).SubItems(9)) = "" Then
            cad = cad & "0"
        Else
            If Trim(ListView1.ListItems(NumRegElim).SubItems(9)) = "-" Then
                cad = cad & "0"
            Else
                cad = cad & DBSet(ListView1.ListItems(NumRegElim).SubItems(9), "N")
            End If
        End If
        
        'Los 4 marcajes, si es que los hay
        For N = 1 To 4
            cad = cad & "," & DevuelveHoraBD(N + 2)
        Next N
        
        If Trim(ListView1.ListItems(NumRegElim).SubItems(7)) = "*" Then
            If Me.Check1.Value = 0 Then
                Aux = " fecha = " & DBSet(ListView1.ListItems(NumRegElim).SubItems(2), "F") & " AND idtrabajador = " & Anterior
            Else
                
                Aux = " fecha = " & DBSet(Anterior, "F") & " AND idtrabajador = " & ListView1.ListItems(NumRegElim).SubItems(1)
            End If
            Aux = "Select * from entradamarcajes where " & Aux & " AND hora > '" & Me.ListView1.ListItems(NumRegElim).SubItems(6) & ":59' order by hora"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                cad = cad & "," & DBSet(miRsAux!Hora, "H")
                N = N + 1
                If N = 16 Then
                    While Not miRsAux.EOF
                        miRsAux.MoveNext
                    Wend
                    N = 17
                Else
                    miRsAux.MoveNext
                End If
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            
        End If
        N = 16 - N + 1
        If N > 0 Then
            Do
                cad = cad & ", null"
                N = N - 1
            Loop Until N = 0
        End If
        cad = cad & ")"
        
        Antiguo = Antiguo & cad
        If Len(Antiguo) > 2000 Then
            Antiguo = Mid(Antiguo, 2)
            cad = "INSERT INTO tmpinformehorasmes(codusu,fecha,idTrabajador,Nombre,HT,HN,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15,H16) VALUES " & Antiguo
            conn.Execute cad
            Antiguo = ""
        End If
    Next
    
    If Antiguo <> "" Then
        Antiguo = Mid(Antiguo, 2)
        cad = "INSERT INTO tmpinformehorasmes(codusu,fecha,idTrabajador,Nombre,HT,HN,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15,H16) VALUES " & Antiguo
        If Not EjecutaSQL(cad) Then
            MsgBox "Error insertando datos informes", vbExclamation
        End If

    End If
    
    Antiguo = ""
    cad = ""
    If Me.txtTrab(4).Text <> "" Then cad = " desde " & Me.txtTrab(4).Text
    If Me.txtTrab(5).Text <> "" Then cad = cad & " hasta " & Me.txtTrab(5).Text
    If cad <> "" Then Antiguo = "Trabajador : " & cad
    cad = ""
    If Me.txtFec(0).Text <> "" Then cad = " desde " & Me.txtFec(0).Text
    If Me.txtFec(1).Text <> "" Then cad = cad & " hasta " & Me.txtFec(1).Text
    If cad <> "" Then Antiguo = Trim(Antiguo & "        Fecha : " & cad)
    
    
    cad = IIf(Me.Option1(0).Value, "idTrabajador", "Nombre")
    cad = "orden= {tmpinformehorasmes." & cad & "}"
    With frmImprimir
        .FormulaSeleccion = "{tmpinformehorasmes.codusu} = " & vUsu.Codigo
        If Me.Check1.Value = 0 Then
            .NombreRPT100 = "visorMarcajesTra.rpt"
        Else
            .NombreRPT100 = "visorMarcajesFecha.rpt"
        End If
        .Titulo100 = "Visor marcajes"
        
        .OtrosParametros = "Emp=""" & vEmpresa.NomEmpresa & """|Sel= """ & Antiguo & """|" & cad & "|"
        .Opcion = 100
        .NumeroParametros = 3
        .Show vbModal
    End With
    
    
    
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaDatos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    'AbriendoForm = False
    imgTra(4).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    imgTra(5).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    
    
    'Pongo la semana anterior.
    If QuieroVerDatos = "" Then
        txtFec(1).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
        txtFec(0).Text = Format(DateAdd("d", -8, Now), "dd/mm/yyyy")
        
    Else
        txtFec(0).Text = Format(RecuperaValor(QuieroVerDatos, 3), "dd/mm/yyyy")
        txtFec(1).Text = Format(RecuperaValor(QuieroVerDatos, 4), "dd/mm/yyyy")
        Me.txtTrab(4).Text = RecuperaValor(QuieroVerDatos, 1)
        Me.txtTrab(5).Text = Me.txtTrab(4).Text
        txtDT(4).Text = RecuperaValor(QuieroVerDatos, 2)
        txtDT(5).Text = txtDT(4).Text
    End If
    
    CargarColumnas
    Set ListView1.SmallIcons = Me.ImageList1
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    
    'EN imgTra(0).Tag  tengo que opcion  ha sido (trabajadores, incidencias...
    ' En imgTra(0).Tag  tendre que INDEX dentro del img
    
    Select Case imgTra(4).Tag
    Case 0
        'TRABAJADORES
        txtTrab(CInt(txtTrab(4).Tag)).Text = RecuperaValor(CadenaDevuelta, 1)
        txtDT(CInt(txtTrab(4).Tag)).Text = RecuperaValor(CadenaDevuelta, 2)
        
    Case 1
'        'INCIDENCIAS
'        txtInci(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 1)
'        txtDInci(CInt(imgInci(0).Tag)).Text = RecuperaValor(CadenaDevuelta, 2)
'
    End Select

End Sub

Private Sub frmc_Selec(vFecha As Date)
    txtFec(CInt(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub




Private Sub CargaDatos()
Dim v As Byte

    v = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    CargaDatos2
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaDatos2()
Dim SQL As String
Dim F As Date
Dim IT As ListItem
Dim Agrupar As String
Dim Salto As Boolean
Dim Entrada As Long
Dim SubI As Long
Dim T1 As Single


Dim UltimaFechaDelTrabajador As String

    lblDuplicado.Visible = False
    lblDuplicado.Refresh
    ListView1.ListItems.Clear
    
    SQL = "select marcajes.fecha,marcajes.idtrabajador,entrada,hora,nomtrabajador,nominci,incfinal "
    SQL = SQL & ",HorasTrabajadas, HorasIncid,excesodefecto , HOUR(HORA) LaHora , minute(hora) LosMinutos "
    SQL = SQL & " from marcajes left join entradamarcajes on marcajes.Entrada = entradamarcajes.idmarcaje,trabajadores,incidencias"
    SQL = SQL & " Where marcajes.idTrabajador = trabajadores.idTrabajador And IncFinal = incidencias.idinci"

    
    'EL SQL PARTICULAR
    If txtFec(0).Text = "" Then
        F = "1/01/2000"
    Else
        F = CDate(txtFec(0).Text)
    End If
        
    SQL = SQL & " and marcajes.fecha>='" & Format(F, FormatoFecha)
    If txtFec(1).Text = "" Then
        F = "1/01/2050"
    Else
        F = CDate(txtFec(1).Text)
    End If
    SQL = SQL & "' and marcajes.fecha <= '" & Format(F, FormatoFecha)
        
        
        
    If Me.txtTrab(4).Text <> "" Then
        SubI = Val(txtTrab(4).Text)
    Else
        SubI = 0
    End If
    SQL = SQL & "' and marcajes.idtrabajador>=" & SubI
    
    If Me.txtTrab(5).Text <> "" Then
        SubI = Val(txtTrab(5).Text)
    Else
        SubI = 32600
    End If
    SQL = SQL & "  and  marcajes.idtrabajador<= " & SubI
    
    
    SQL = SQL & " ORDER BY "
    If Check1.Value = 1 Then SQL = SQL & "fecha,"
    If Option1(0).Value Then
        SQL = SQL & "idtrabajador"
    Else
        SQL = SQL & "nomtrabajador"
    End If
    If Check1.Value = 1 Then SQL = SQL & ",entrada"   'Fecha. Metemos la entrada tambien
    If Not (Check1.Value = 1) Then SQL = SQL & ",fecha,entrada"
    SQL = SQL & ",hora"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    DoEvents
    SQL = ""
    Agrupar = ""
    Entrada = 0
    T1 = Timer
    While Not miRsAux.EOF
        If Timer - T1 > 2 Then
            T1 = Timer
            Me.Refresh
        End If
        If Check1.Value = 1 Then
            SQL = Format(miRsAux!Fecha, "dd/mm/yyyy")
            Salto = miRsAux.Fields(0) <> Agrupar
        Else
            Salto = miRsAux!idTrabajador <> Agrupar
        End If
        If Salto Then
            Set IT = ListView1.ListItems.Add()
            
            SQL = Format(miRsAux.Fields(0), "ddmmyyyy") & "_" & Format(miRsAux!idTrabajador, "000000")
            If SQL = UltimaFechaDelTrabajador Then
                
                    lblDuplicado.Visible = True
                    IT.Bold = True
                    IT.ForeColor = vbRed
                    IT.ListSubItems(2).Bold = True
                    IT.ListSubItems(2).ForeColor = vbRed
                    
            
            End If
            UltimaFechaDelTrabajador = SQL
            
            If Check1.Value = 1 Then
                Agrupar = Format(miRsAux.Fields(0), "dd/mm/yyyy")
                SQL = Agrupar
            Else
                'Aunque sea por nombre, el codigo no sirve
                Agrupar = miRsAux!idTrabajador
                SQL = Agrupar
            End If
            IT.Text = SQL
            If Check1.Value = 1 Then
                IT.SmallIcon = 3
                IT.SubItems(1) = miRsAux!idTrabajador
                IT.SubItems(2) = miRsAux!nomtrabajador
            Else
                IT.SmallIcon = 2
                'If Option1(0).Value Then
                    IT.SubItems(1) = miRsAux!nomtrabajador
                'Else
                '    IT.SubItems(1) = miRsAux!idTrabajador
                'End If
                IT.SubItems(2) = Format(miRsAux!Fecha, "dd/mm/yyyy")
            End If
            Entrada = 0
        End If
        
        'If miRsAux!Entrada = 303266 Then St op
        
        If Entrada <> miRsAux!Entrada Then
            If Not Salto Then
                'Estamos en la misma agrupacion, pero es un item nuevo
                Set IT = ListView1.ListItems.Add()
                If Check1.Value = 1 Then
                    IT.SubItems(1) = miRsAux!idTrabajador
                    IT.SubItems(2) = miRsAux!nomtrabajador
                Else
                    IT.SubItems(2) = Format(miRsAux!Fecha, "dd/mm/yyyy")
                End If
                SQL = Format(miRsAux.Fields(0), "ddmmyyyy") & "_" & Format(miRsAux!idTrabajador, "000000")
            
                If SQL = UltimaFechaDelTrabajador Then
                    
                    
                    lblDuplicado.Visible = True
                    IT.Bold = True
                    IT.ForeColor = vbRed
                    IT.ListSubItems(2).Bold = True
                    IT.ListSubItems(2).ForeColor = vbRed
                    
                End If
                UltimaFechaDelTrabajador = SQL
                
            End If
            Entrada = miRsAux!Entrada
            
            
            
            
            IT.Key = "C" & Entrada
            IT.SubItems(8) = Format(miRsAux!HorasTrabajadas, "0.00")
            If miRsAux!IncFinal <> 0 Then
                If miRsAux!ExcesoDefecto = 1 Then
                    IT.SubItems(9) = Format(miRsAux!HorasIncid, "0.00")
                Else
                    IT.SubItems(9) = "-"
                End If
                IT.SubItems(10) = miRsAux!NomInci
            End If
            'Pongo la primera hora
            IT.SubItems(3) = Format(miRsAux!Hora, "hh:mm")
            SubI = 4
        Else
            'Ponemos solo la hora
            If SubI > 6 Then
                IT.SubItems(7) = "*"
            Else
                IT.SubItems(7) = "."
                If IsNull(miRsAux!Hora) Then
                    IT.SubItems(SubI) = " "
                Else
                    IT.SubItems(SubI) = PonerTextoHoraConNull
                End If
                SubI = SubI + 1
            End If
        
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub

Private Function PonerTextoHoraConNull() As String
Dim SoloHora As Integer
Dim Minutos As Integer

    On Error Resume Next
    SoloHora = miRsAux!LaHora
    If SoloHora < 0 Then
        Stop
        SoloHora = SoloHora + 24
    ElseIf SoloHora >= 24 Then
        SoloHora = SoloHora - 24
    Else
        'no hago nada. hora correcta
        
    End If
    PonerTextoHoraConNull = Format(SoloHora, "00") & ":" & Format(miRsAux!LosMinutos, "00")
    If Err.Number <> 0 Then
        Err.Clear
        PonerTextoHoraConNull = " "
    End If
End Function

Private Sub CargarColumnas()
Dim L As Collection
Dim i As Integer
Dim C As ColumnHeader

    ListView1.ColumnHeaders.Clear
    Set L = New Collection
    
    
    
    If Check1.Value = 1 Then L.Add "Fecha|1300|"
    
    
    
    L.Add "Codigo|850|"
    L.Add "Nombre|2900|"

    If Not (Check1.Value = 1) Then L.Add "Fecha|1100|"

    'Las columnas para el resto de campos
    For i = 1 To 4
        L.Add "H" & i & "|800|"
    Next i
    'Columna para marcar si hay mas
    L.Add "+|300|"
    
    'Horas trabjadas
    L.Add "Total|1000|"
    L.Add "Inci|800|"
    L.Add "Descr.|1650|"
    
    
    'TOTAL..... 11 campos
    For i = 1 To 11
        Set C = ListView1.ColumnHeaders.Add(, "C" & i)
        C.Text = RecuperaValor(L.Item(i), 1)
        C.Width = RecuperaValor(L.Item(i), 2)
    Next i
    
    'A MANO
    '---------
    ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
  '  ListView1.ColumnHeaders(10).Alignment = lvwColumnRight
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim Obj As Object

    Antiguo = Me.txtFec(Index).Text

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
    If txtFec(Index).Text <> "" Then frmc.NovaData = txtFec(Index).Text
    ' ********************************************
    'AbriendoForm = True
    frmc.Show vbModal
    Set frmc = Nothing
    'AbriendoForm = False
    ' *** repasar si el camp es txtAux o Text1 ***
    'PonerFoco txtFec(CByte(imgFec(0).Tag)) '<===
    ' ********************************************
    If Antiguo <> txtFec(Index).Text Then
        DoEvents
        CargaDatos
    End If
    
    
End Sub


Private Sub imgTra_Click(Index As Integer)
    Antiguo = Me.txtTrab(Index).Text
    imgTra(4).Tag = 0 'Para que el devuelve grid sepa que es TRABAJADORES
    txtTrab(4).Tag = Index
    cad = "Codigo|idTrabajador|N||15·"
    cad = cad & "Nombre|nomtrabajador|T||60·"
    cad = cad & "Tarjeta|numtarjeta|T||20·"
    Set frmB = New frmBuscaGrid
    frmB.vTabla = "Trabajadores"
    frmB.vCampos = cad
    frmB.vDevuelve = "0|1|"
    frmB.vSelElem = 0
    frmB.vTitulo = "TRABAJADORES"
    frmB.Show vbModal
    Set frmB = Nothing
    If Antiguo <> Me.txtTrab(Index).Text Then
        DoEvents
        CargaDatos
    End If
End Sub


Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    CadenaDesdeOtroForm = ""
    frmRevision.MostrarUnosDatos = Val(Mid(ListView1.SelectedItem.Key, 2))
    frmRevision.Show vbModal
End Sub

Private Sub Option1_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    CargaDatos
    Screen.MousePointer = vbDefault
End Sub






Private Sub txtFec_GotFocus(Index As Integer)
    Antiguo = txtFec(Index).Text
    ConseguirFoco txtFec(Index), 3
    
End Sub

Private Sub txtFec_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtFec_LostFocus(Index As Integer)
    'If AbriendoForm Then Exit Sub
    txtFec(Index).Text = Trim(txtFec(Index).Text)
    If txtFec(Index).Text <> "" Then
        If Not EsFechaOK(txtFec(Index)) Then txtFec(Index).Text = ""
    End If
    If Antiguo <> txtFec(Index).Text Then CargaDatos
End Sub

Private Sub txtTrab_GotFocus(Index As Integer)
    Antiguo = txtTrab(Index).Text
    ConseguirFoco txtTrab(Index), 3
End Sub

Private Sub txtTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtTrab_LostFocus(Index As Integer)
    txtTrab(Index).Text = Trim(txtTrab(Index))
    If txtTrab(Index).Text = "" Then
        Me.txtDT(Index).Text = ""
    Else
        If IsNumeric(txtTrab(Index).Text) Then
            cad = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idtrabajador", txtTrab(Index).Text, "N")
        Else
            txtTrab(Index).Text = ""
            cad = ""
        End If
        txtDT(Index).Text = cad
    End If
    If Antiguo <> txtTrab(Index).Text Then CargaDatos
    
End Sub


Private Sub Keypress(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then
        Unload Me 'ESC
    End If
End Sub

