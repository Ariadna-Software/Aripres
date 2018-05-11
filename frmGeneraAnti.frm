VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGeneraAnti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación anticipos"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   11775
      Begin VB.ComboBox cboSeccion 
         Height          =   315
         ItemData        =   "frmGeneraAnti.frx":0000
         Left            =   3960
         List            =   "frmGeneraAnti.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   2
         Left            =   9600
         ScrollBars      =   1  'Horizontal
         TabIndex        =   13
         Top             =   405
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cod."
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
         Height          =   375
         Left            =   10920
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdImpr 
         Caption         =   "Recibos"
         Height          =   375
         Index           =   1
         Left            =   7560
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdImpr 
         Caption         =   "Imprimir"
         Height          =   375
         Index           =   0
         Left            =   7560
         TabIndex        =   10
         Top             =   150
         Width           =   855
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   1
         Left            =   840
         ScrollBars      =   1  'Horizontal
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Index           =   0
         Left            =   840
         ScrollBars      =   1  'Horizontal
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nom."
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Caption         =   "Secciones"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblFecha 
         Caption         =   "Anticipo"
         Height          =   195
         Index           =   1
         Left            =   8640
         TabIndex        =   14
         Top             =   450
         Width           =   585
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   9360
         Picture         =   "frmGeneraAnti.frx":0004
         ToolTipText     =   "Buscar fecha"
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmGeneraAnti.frx":008F
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   600
         Picture         =   "frmGeneraAnti.frx":011A
         ToolTipText     =   "Buscar fecha"
         Top             =   262
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde"
         Height          =   195
         Index           =   32
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Width           =   465
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   7875
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   13891
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6879
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   1773
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "HN"
         Object.Width           =   1206
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Importe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Antig."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "SS"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Ret."
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "TOTAL N"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7875
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   13891
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "HN"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Imp/h"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "HC"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Imp/h"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "%SS"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "% IRPF"
         Object.Width           =   1191
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Pagos"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "A ingresar"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "T1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "T2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Bruto"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmGeneraAnti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Antiguedad2 As Byte
Private WithEvents frmc As frmCal
Attribute frmc.VB_VarHelpID = -1



Private Sub cmdCargar_Click()
    
    If cboSeccion.ListIndex < 0 Then
        MsgBox "Seleccione la seccion", vbExclamation
        Exit Sub
    End If
     
    If txtFec(0).Text = "" Or txtFec(1).Text = "" Then
        MsgBox "Escriba las fechas de inicio y fin", vbExclamation
        Exit Sub
    End If
    If Abs(DateDiff("d", CDate(txtFec(0).Text), CDate(txtFec(1).Text))) > 31 Then
        MsgBox "Intervalo incorrecto", vbExclamation
        Exit Sub
    End If

    If Antiguedad2 = 0 Then
        CargaDatosModo1
    Else
        CargaDatosAntiguedad
    End If
End Sub






Private Sub CargaDatosModo1()
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Importe As Currency
Dim Importe2 As Currency
Dim itmX As ListItem

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders(12).Width = 0
    ListView1.ColumnHeaders(13).Width = 0
    ListView1.ColumnHeaders(14).Width = 0
    Sql = "SELECT Trabajadores.IdTrabajador, Trabajadores.NomTrabajador, "
    Sql = Sql & " tmpHoras.HorasT, Categorias.Importe1,[HorasT]*[Importe1] AS T1,"
    Sql = Sql & " tmpHoras.HorasC, Categorias.Importe2,[HorasC]*[Importe2] AS T2 "
    'SQL = SQL & " tmpHoras.HorasE, Categorias.Importe3,[HorasE]*[Importe2] AS T3 "
    Sql = Sql & " ,Trabajadores.PorcSS, Trabajadores.PorcIRPF, Trabajadores.ControlNomina"
    Sql = Sql & " FROM (Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria) INNER JOIN tmpHoras ON Trabajadores.IdTrabajador = tmpHoras.trabajador"
    Sql = Sql & " ORDER BY "
    If Option1(0).Value Then
        Sql = Sql & "idTrabajador "
    Else
        Sql = Sql & "nomtrabajador"
    End If
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set itmX = ListView1.ListItems.Add(, , RS.Fields(0))
        itmX.SubItems(1) = RS!nomtrabajador
        itmX.SubItems(2) = Format(RS!horast, FormatoImporte)
        Importe = Round(RS!Importe1 * RS!horast, 2)
        itmX.SubItems(3) = Format(RS!Importe1, FormatoImporte)

        'itmX.SubItems(3) = Format(Importe, FormatoImporte)
        
        itmX.SubItems(4) = Format(RS!HorasC, FormatoImporte)
        
        
        If RS!ControlNomina = 2 Then
            Importe = 0

        Else
            Importe = RS!Importe2
        End If
        itmX.SubItems(5) = Format(Importe, FormatoImporte)

        itmX.SubItems(6) = Format(RS!PorcSS, FormatoImporte)
        itmX.SubItems(7) = Format(RS!PorcIRPF, FormatoImporte)
        

        Importe = Round(RS!Importe1 * RS!horast, 2)
        itmX.SubItems(3) = Format(RS!Importe1, FormatoImporte)

        Importe = RS!T1
        itmX.SubItems(11) = Importe
        
        If RS!ControlNomina = 2 Then
            Importe2 = 0
        Else
            Importe2 = RS!T2
        End If
        itmX.SubItems(12) = Importe2
            
        Importe = Importe + Importe2
        Importe = Round(Importe, 2)
        'Bruto
        itmX.SubItems(13) = Importe
        
        'Iconito en funcion del tipo de control de nominas
        If RS!ControlNomina = 2 Then
            itmX.SmallIcon = 2
        Else
            itmX.SmallIcon = 1
        End If
        
        
        Importe2 = RS!PorcSS + RS!PorcIRPF
        Importe2 = (Importe2 * Importe) / 100
        Importe2 = Round(Importe2, 2)
        Importe = Importe - Importe2  'BRUTO - IRPF - SS
        itmX.SubItems(8) = Format(Importe, FormatoImporte)   'TOTAL
        'Obtner pagos efectuados en el periodo
        Importe2 = ObtenerPagosPeriodo(RS.Fields(0))
        itmX.SubItems(9) = Format(Importe2, FormatoImporte)
        'TOTAL A INGRESAR
        Importe2 = Importe - Importe2
        itmX.SubItems(10) = Format(Importe2, FormatoImporte)
            
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Function CrearAnticipos() As Boolean
Dim i As Integer
'Dim Aux As String
'Dim SQL As String

Dim Importe As Currency
'Dim Aux2 As String
Dim Importe2 As Currency
Dim i2 As Currency
Dim idTrabajador As Long


    On Error GoTo ECrearAnticipos
    CrearAnticipos = False
    
     
    
    'vEmpresa.AbonosSeparados
    
    'Para ver la secuencia
    Set miRsAux = New ADODB.Recordset
    
    'Para la insercion
    'SQL = "INSERT INTO Pagos (Fecha,Observaciones,Pagado,Trabajador,Importe,Tipo) VALUES ("

    'Generaremos los anticipos
    If Antiguedad2 = 0 Then
        For i = 1 To ListView1.ListItems.Count
            'Aux = SQL & ListView1.ListItems(I).Text & ","
            'Aux = Aux & TransformaComasPuntos(CStr(ImporteFormateado(ListView1.ListItems(I).SubItems(10)))) & ",1)"   '1 de anticipo
           '
           ' conn.Execute Aux
        Next i
    Else
        'Antiguedad
        'Si son abonos separados
        
        CadenaDesdeOtroForm = "insert into genanticiposproceso(fecha,fechaanticipo,fechaIni,fechaFin,codusu,Nombre) VALUES ("
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "now()," & DBSet(txtFec(2).Text, "F") & ","
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & DBSet(txtFec(0).Text, "F") & "," & DBSet(txtFec(1).Text, "F") & ","
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vUsu.Codigo & "," & DBSet(vUsu.Nombre, "T") & ")"
        conn.Execute CadenaDesdeOtroForm
        
        idTrabajador = 0
        For i = 1 To ListView2.ListItems.Count
            'El trabajador es el mismo
            If Trim(Me.ListView2.ListItems(i).Text) <> "" Then
                If Val(Me.ListView2.ListItems(i).Text) <> idTrabajador Then
                    
                    If idTrabajador > 0 Then InsertaEnPagos idTrabajador, Importe, Importe2
                    Importe = 0
                    Importe2 = 0
                    idTrabajador = Val(ListView2.ListItems(i).Text)
                End If
            End If
            If Trim(ListView2.ListItems(i).SubItems(8)) <> "" Then
                i2 = ImporteFormateado(ListView2.ListItems(i).SubItems(8))
            Else
                i2 = 0
            End If
            If ListView2.ListItems(i).Tag <> "" Then
                If ListView2.ListItems(i).Tag = 0 Then
                    Importe = Importe + i2
                Else
                    Importe2 = Importe2 + i2
                End If
            End If
        Next i
        If idTrabajador > 0 Then InsertaEnPagos idTrabajador, Importe, Importe2
    End If
    
    MsgBox "Proceso finalizado", vbInformation
    
    CrearAnticipos = True
    Exit Function
ECrearAnticipos:

    MuestraError Err.Number, Err.Description & vbCrLf
    
End Function


Private Sub InsertaEnPagos(idTrabajador As Long, ImporteN As Currency, ImporteE As Currency)
Dim Sql As String

    If ImporteN + ImporteE = 0 Then Exit Sub

    If Not vEmpresa.AbonosSeparados Then
        ImporteN = ImporteN + ImporteE
        ImporteE = 0
    End If
    Sql = ""
    If ImporteN <> 0 Then
        Sql = ", (" & DBSet(Me.txtFec(2).Text, "F") & ",'Anticipo:  " & Format(txtFec(0).Text, "dd/mm/yy") & " - " & Format(txtFec(1).Text, "dd/mm/yy") & "',0,"
        Sql = Sql & idTrabajador & "," & TransformaComasPuntos(CStr(ImporteN)) & ",1)"  '1 anticpo normal
    End If
    
    
    If ImporteE <> 0 Then
        Sql = Sql & ", (" & DBSet(Me.txtFec(2).Text, "F") & ",'Anticipo:  " & Format(txtFec(0).Text, "dd/mm/yy") & " - " & Format(txtFec(1).Text, "dd/mm/yy") & "',0,"
        Sql = Sql & idTrabajador & "," & TransformaComasPuntos(CStr(ImporteE)) & ",3)"   '3 anticpo extras
    End If
    If Sql = "" Then
        '
    Else
        Sql = Mid(Sql, 2)
    
        Sql = "INSERT INTO Pagos (Fecha,Observaciones,Pagado,Trabajador,Importe,Tipo) VALUES " & Sql
        conn.Execute Sql
    End If
End Sub

Private Sub CargaDatosAntiguedad()
Dim C  As String

    C = DevuelveDesdeBD("nominas", "secciones", "idseccion", CStr(Me.cboSeccion.ItemData(Me.cboSeccion.ListIndex)))
    If C = "1" Then
        CargaDatosAntiguedadNominas
    Else
        CargaDatosAntiguedadDesdeMarcajes
    End If

End Sub


Private Sub CargaDatosAntiguedadNominas()
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Importe As Currency
Dim Importe2 As Currency
Dim Importe3 As Currency
Dim itmX As ListItem
Dim RCat As ADODB.Recordset
Dim Trabajador As Long
Dim IndiceNombreTrabajador As Integer
Dim AplicoAnt As Boolean
Dim Lineas As Byte


    ListView2.ListItems.Clear


   Set RCat = New ADODB.Recordset
   Sql = "Select importe1,importe2,importe3,IdCategoria from categorias"   'UNO POR CADA TIPODE HORA
   RCat.Open Sql, conn, adOpenKeyset, adLockPessimistic
   

    Sql = "select jornadassemanalesalz.idtrabajador,tipohoras,NomTrabajador,idCategoria,PorcSS, Trabajadores.PorcIRPF "
    Sql = Sql & " ,Trabajadores.PorcAntiguedad,controlnomina,DescTipoHora,sum(horastrabajadas) lashoras"
    Sql = Sql & " from jornadassemanalesalz ,tiposhora,trabajadores where "
    Sql = Sql & " tiposhora.TipoHora = jornadassemanalesalz.tipohoras "
    Sql = Sql & " AND jornadassemanalesalz.idtrabajador = trabajadores.idtrabajador "
    Sql = Sql & " AND fecha >=" & DBSet(Me.txtFec(0).Text, "F") & " AND fecha <=" & DBSet(Me.txtFec(1).Text, "F")
    'Seccion
    Sql = Sql & " And seccion = " & Me.cboSeccion.ItemData(cboSeccion.ListIndex)
    
    
    
    Sql = Sql & " group by 1,2 order by "
    If Option1(0).Value Then
        Sql = Sql & "1,2 "
    Else
        Sql = Sql & "3,2"
    End If
    Set RS = New ADODB.Recordset
    
    Trabajador = -1
    
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'List 2
        
        If Trabajador <> RS!idTrabajador Then
        
            If Trabajador >= 0 And Lineas > 1 Then RealizaSumatorio itmX, IndiceNombreTrabajador
        
            Set itmX = ListView2.ListItems.Add(, , RS!idTrabajador)
            Trabajador = RS!idTrabajador
            itmX.SubItems(1) = RS!nomtrabajador
            Lineas = 0
            IndiceNombreTrabajador = ListView2.ListItems.Count
            
            'Iconito en funcion del tipo de control de nominas
            If RS!ControlNomina = 2 Then
                itmX.SmallIcon = 6
            Else
                itmX.SmallIcon = 7
            End If
            
                    
            'Buscamos para la categoria del trabajador
            RCat.Find "IdCategoria = " & RS!idCategoria, , adSearchForward, 1
            
        Else
            Set itmX = ListView2.ListItems.Add(, , " ")
            itmX.SubItems(1) = " "
        End If
        
        Lineas = Lineas + 1
        itmX.SubItems(2) = Mid(RS!Desctipohora, 1, 6)
        
        itmX.SubItems(3) = Format(RS!lashoras, FormatoImporte)
        'Importe 1

        Importe = 0
        If Not RCat.EOF Then Importe = RCat.Fields(CInt(RS!TipoHoras))
        
        Importe = Round(Importe * RS!lashoras, 2)
        itmX.SubItems(4) = Format(Importe, FormatoImporte)
        
        'Antiguedad
        AplicoAnt = False
        If RS!TipoHoras = 0 Then
            If vEmpresa.AplicaAntiguedadHN Then AplicoAnt = True
        Else
            If vEmpresa.AplicaAntiguedadHC Then AplicoAnt = True
        End If
        
        If AplicoAnt Then
            Importe2 = Round((Importe * RS!PorcAntiguedad) / 100, 2)
        Else
            Importe2 = 0
        End If
        itmX.SubItems(5) = Format(Importe2, FormatoImporte)
        Importe = Importe + Importe2
        
        'IRPF y RET sobre la misma BASE, importe2
        
        Importe2 = Round((Importe * RS!PorcIRPF) / 100, 2)
        itmX.SubItems(6) = Format(Importe2, FormatoImporte)
        
        Importe3 = Round((Importe * RS!PorcSS) / 100, 2)
        itmX.SubItems(7) = Format(Importe3, FormatoImporte)
        
        Importe3 = Importe3 + Importe2
        Importe = Importe - Importe3
        itmX.SubItems(8) = Format(Importe, FormatoImporte)
        
            
        itmX.Tag = RS!TipoHoras
            
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    If Trabajador >= 0 And Lineas > 1 Then RealizaSumatorio itmX, IndiceNombreTrabajador
    
End Sub


Private Sub CargaDatosAntiguedadDesdeMarcajes()
Dim Sql As String
Dim RS As ADODB.Recordset
Dim itmX As ListItem
Dim RCat As ADODB.Recordset
Dim Trabajador As Long
Dim IndiceNombreTrabajador As Integer
Dim AplicoAnt As Boolean
Dim Importe As Currency
Dim Importe2 As Currency
Dim Importe3 As Currency

'Dim Lineas As Byte


   ListView2.ListItems.Clear


   Set RCat = New ADODB.Recordset
   Sql = "Select importe1,importe2,importe3,IdCategoria from categorias"   'UNO POR CADA TIPODE HORA
   RCat.Open Sql, conn, adOpenKeyset, adLockPessimistic
   


    Sql = "select marcajes.idtrabajador,0,NomTrabajador,idCategoria,PorcSS, Trabajadores.PorcIRPF"
    Sql = Sql & " ,Trabajadores.PorcAntiguedad,controlnomina,DescTipoHora,sum(horastrabajadas) lashoras from"
    Sql = Sql & " marcajes,tiposhora,trabajadores where"
    Sql = Sql & " tiposhora.TipoHora = 0  AND"
    Sql = Sql & " marcajes.idtrabajador = trabajadores.idtrabajador  "
    Sql = Sql & " AND fecha >=" & DBSet(Me.txtFec(0).Text, "F") & " AND fecha <=" & DBSet(Me.txtFec(1).Text, "F")
    'Seccion
    Sql = Sql & " And seccion = " & Me.cboSeccion.ItemData(cboSeccion.ListIndex)
    
    
    
    Sql = Sql & " group by 1 order by "
    If Option1(0).Value Then
        Sql = Sql & "1 "
    Else
        Sql = Sql & "2"
    End If
    Set RS = New ADODB.Recordset
    
    Trabajador = -1
    
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'List 2
        
         
        
        Set itmX = ListView2.ListItems.Add(, , RS!idTrabajador)
        Trabajador = RS!idTrabajador
        itmX.SubItems(1) = RS!nomtrabajador
        
        IndiceNombreTrabajador = ListView2.ListItems.Count
        
        'Iconito en funcion del tipo de control de nominas
        If RS!ControlNomina = 2 Then
            itmX.SmallIcon = 6
        Else
            itmX.SmallIcon = 7
        End If
        
                    
        'Buscamos para la categoria del trabajador
        RCat.Find "IdCategoria = " & RS!idCategoria, , adSearchForward, 1
    
        itmX.SubItems(2) = Mid(RS!Desctipohora, 1, 6)
        
        itmX.SubItems(3) = Format(RS!lashoras, FormatoImporte)
        'Importe 1

        Importe = 0
        If Not RCat.EOF Then Importe = RCat.Fields(0)
        
        Importe = Round(Importe * RS!lashoras, 2)
        itmX.SubItems(4) = Format(Importe, FormatoImporte)
        
        'Antiguedad
        AplicoAnt = False
      '  If RS!TipoHoras = 0 Then
            If vEmpresa.AplicaAntiguedadHN Then AplicoAnt = True
      '  Else
      '      If vEmpresa.AplicaAntiguedadHC Then AplicoAnt = True
      '  End If
        
        If AplicoAnt Then
            Importe2 = Round((Importe * RS!PorcAntiguedad) / 100, 2)
        Else
            Importe2 = 0
        End If
        itmX.SubItems(5) = Format(Importe2, FormatoImporte)
        Importe = Importe + Importe2
        
        'IRPF y RET sobre la misma BASE, importe2
        
        Importe2 = Round((Importe * RS!PorcIRPF) / 100, 2)
        itmX.SubItems(6) = Format(Importe2, FormatoImporte)
        
        Importe3 = Round((Importe * RS!PorcSS) / 100, 2)
        itmX.SubItems(7) = Format(Importe3, FormatoImporte)
        
        Importe3 = Importe3 + Importe2
        Importe = Importe - Importe3
        itmX.SubItems(8) = Format(Importe, FormatoImporte)
        
            
        itmX.Tag = 0
            
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    
End Sub




Private Function ObtenerPagosPeriodo(Traba As Long) As Currency
Dim Impor As Currency
Dim SQLAUX As String
    Stop
    

    'SQLAUX = "Select importe from pagos where fecha >=#" & Format(Text1(0).Text, FormatoFecha) & "#"
    'SQLAUX = SQLAUX & " AND  fecha <=#" & Format(Text1(1).Text, FormatoFecha) & "#"
    SQLAUX = SQLAUX & " AND tipo = 0"  'Pagos adelantados al trabajador
    SQLAUX = SQLAUX & " AND Trabajador = " & Traba
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQLAUX, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Impor = 0
    If Not miRsAux.EOF Then
        While Not miRsAux.EOF
            Impor = Impor + miRsAux.Fields(0)
            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    ObtenerPagosPeriodo = Impor
End Function

Private Sub cmdGenerar_Click()
     
    If txtFec(2).Text = "" Then
        MsgBox "Ponga la fecha del anticipo", vbExclamation
        Exit Sub
    End If
    If Antiguedad2 = 0 Then
        If ListView1.ListItems.Count = 0 Then Exit Sub
    Else
        If ListView2.ListItems.Count = 0 Then Exit Sub
    End If
    
    
    
    CadenaDesdeOtroForm = "(tipo = 1 or tipo =3) AND fecha "  'ANTICIPO
    CadenaDesdeOtroForm = DevuelveDesdeBD("count(*)", "pagos", CadenaDesdeOtroForm, Me.txtFec(2).Text, "F")
    
    If Val(CadenaDesdeOtroForm) > 0 Then
        MsgBox "Ya existen anticipos con esa fecha"
    Else
    
        CadenaDesdeOtroForm = "Seguro que desea generar los anticipos con estos valores?"
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel) = vbYes Then CadenaDesdeOtroForm = ""
    End If
    
    Screen.MousePointer = vbHourglass
    If CadenaDesdeOtroForm = "" Then CrearAnticipos
    Screen.MousePointer = vbDefault
    
    
    CadenaDesdeOtroForm = ""
    
End Sub

Private Sub cmdImpr_Click(Index As Integer)
    
    ImprimirNormal Index = 1
End Sub

Private Sub Form_Load()
  
    Me.Icon = frmMain.Icon
    Me.ListView2.SmallIcons = frmPpal.imgListImages16
  
    Antiguedad2 = 0
    If vEmpresa.AplicaAntiguedadHN Then Antiguedad2 = 1
    If vEmpresa.AplicaAntiguedadHC Then
        If Antiguedad2 = 0 Then
            Antiguedad2 = 2 'Solo sobre Compensables  raro raro raro este caso
        Else
            Antiguedad2 = 3
        End If
    End If
    
    'insert into genanticiposproceso(fecha,fechaanticipo,fechaIni,fechaFin,codusu,Nombre)
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select max(fechafin) from genanticiposproceso", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            Me.txtFec(0).Text = DateAdd("d", 1, miRsAux.Fields(0))
            Me.txtFec(1).Text = DateAdd("d", 14, Me.txtFec(0).Text)
        End If
    End If
    miRsAux.Close
    
    
    cboSeccion.Clear
    
    miRsAux.Open "select * from secciones where idseccion in (select seccion from trabajadores) ORDER BY nombre", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cboSeccion.AddItem miRsAux!Nombre & " (" & miRsAux!IdSeccion & ")"
        cboSeccion.ItemData(cboSeccion.NewIndex) = miRsAux!IdSeccion
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    cboSeccion.ListIndex = 0
    
End Sub

Private Sub frmc_Selec(vFecha As Date)
       txtFec(CInt(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
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
    If txtFec(Index).Text <> "" Then frmc.NovaData = txtFec(Index).Text
    ' ********************************************

    frmc.Show vbModal
    Set frmc = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtFec(CByte(imgFec(0).Tag)) '<===
    ' ********************************************

End Sub

Private Sub txtFec_GotFocus(Index As Integer)
    txtFec(Index).SelStart = 0
    txtFec(Index).SelLength = Len(txtFec(Index).Text)
End Sub


Private Sub txtFec_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtFec_LostFocus(Index As Integer)
    txtFec(Index).Text = Trim(txtFec(Index).Text)
    If txtFec(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(txtFec(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFec(Index).Text, vbExclamation
        txtFec(Index).Text = ""
        PonerFoco txtFec(Index)
    End If
End Sub

Private Sub KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    
    End If
End Sub

Private Sub RealizaSumatorio(ByRef IT As ListItem, Inicio As Integer)
Dim i As Integer
Dim Aux As Currency
Dim columna As Byte
    Set IT = ListView2.ListItems.Add(, , " ")
    
    IT.SubItems(1) = "Total"
    IT.ListSubItems(1).ForeColor = vbBlue
    IT.SubItems(2) = "-"
    IT.SubItems(3) = "-"
    IT.SubItems(4) = "-"
        
    For columna = 3 To 8
        Aux = 0
        For i = Inicio To Me.ListView2.ListItems.Count - 1
            If Trim(Me.ListView2.ListItems(i).SubItems(columna)) <> "" Then
                If Me.ListView2.ListItems(i).SubItems(columna) <> "0,00" Then Aux = Aux + ImporteFormateado(ListView2.ListItems(i).SubItems(columna))
            End If
        Next i
        If Aux = 0 Then
            IT.SubItems(columna) = " "
        Else
            IT.SubItems(columna) = Format(Aux, FormatoImporte)
        End If
        IT.ListSubItems(columna).ForeColor = vbBlue
    Next columna
    
    

End Sub









Private Sub ImprimirNormal(Recibos As Boolean)
Dim cadAux  As String
Dim F1 As Date
Dim Nominas As Boolean

    If ListView1.ListItems.Count = 0 And ListView2.ListItems.Count = 0 Then Exit Sub
    
            
    Screen.MousePointer = vbHourglass
 
    If Antiguedad2 = 0 Then
    
        
        
    
    Else
        'CON ANTIGUEDAD
            
    
    End If
        
        
    cadAux = DevuelveDesdeBD("nominas", "secciones", "idseccion", CStr(Me.cboSeccion.ItemData(Me.cboSeccion.ListIndex)))
    Nominas = Val(cadAux) = 1
    
    If Nominas Then
        F1 = CDate(Me.txtFec(0).Text)
        cadAux = " {jornadassemanalesalz.fecha} >= Date(" & Year(F1) & "," & Month(F1) & "," & Day(F1) & ")"
        F1 = CDate(Me.txtFec(1).Text)
        cadAux = cadAux & " AND {jornadassemanalesalz.fecha} <= Date(" & Year(F1) & "," & Month(F1) & "," & Day(F1) & ")"
                        
    Else
        F1 = CDate(Me.txtFec(0).Text)
        cadAux = " {marcajes.fecha} >= Date(" & Year(F1) & "," & Month(F1) & "," & Day(F1) & ")"
        F1 = CDate(Me.txtFec(1).Text)
        cadAux = cadAux & " AND {marcajes.fecha} <= Date(" & Year(F1) & "," & Month(F1) & "," & Day(F1) & ")"
    
    End If
    
    cadAux = cadAux & " AND {trabajadores.seccion} = " & Me.cboSeccion.ItemData(cboSeccion.ListIndex)
    frmImprimir.FormulaSeleccion = cadAux
    
    cadAux = "|FechaIni= """ & Me.txtFec(0).Text & """|FechaFin= """ & Me.txtFec(1).Text & """|"
    frmImprimir.Opcion = 100
    If Recibos Then
        frmImprimir.Titulo100 = "Resumen nomina trabajador"
        
        If Nominas Then
            frmImprimir.NombreRPT100 = "NominaMes.rpt"
        Else
            frmImprimir.NombreRPT100 = "NominaMesSecc.rpt"
        End If
    Else
        frmImprimir.Titulo100 = "Listado generación"
        
         If Nominas Then
            frmImprimir.NombreRPT100 = "ResNominasMes.rpt"
        Else
            frmImprimir.NombreRPT100 = "ResNominasMesSecc.rpt"
        End If
    End If
    
    
    
    frmImprimir.ConSubreport100 = True
    frmImprimir.OtrosParametros = cadAux
    frmImprimir.NumeroParametros = 2
    frmImprimir.Show vbModal
    Screen.MousePointer = vbDefault


End Sub






