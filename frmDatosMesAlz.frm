VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatosMesAlz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver datos Nominas"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   9120
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   15690
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6597
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dias"
         Object.Width           =   1587
      EndProperty
   End
End
Attribute VB_Name = "frmDatosMesAlz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Mes As Date  'Sera 01/mes/año

Private Const ColumnaDondeEmpiezanHoras = 3

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
 
 If CargarDatosImpresion Then
       
        With frmImprimir
            .FormulaSeleccion = "{tmpcombinada.codusu} = " & vUsu.Codigo
            .NombreRPT100 = "DatosMes.rpt"
            .Titulo100 = "Mes trabajado"
            .OtrosParametros = "FechaFin= ""Mes: " & Format(Mes, "mmmm yyyy") & " ""|"
            .Opcion = 100
            .NumeroParametros = 1
            .Show vbModal
        End With
        
    End If

End Sub

Private Sub Form_Load()

    
    Me.Icon = frmMain.Icon
    
    
    
    
    CargaDatos 0 'Horas trabajadas para la coperativa
     
    Me.cmdCancelar.Left = Me.Width - Me.cmdCancelar.Width - 240
    Me.cmdImprimir.Left = Me.cmdCancelar.Left - Me.cmdImprimir.Width - 240

End Sub



Private Sub CargaDatos(ParaLaCooperativa As Byte)
Dim Cad As String
Dim idTrabajador As Long
Dim Fecha As Date
Dim IT As ListItem
Dim CuantosTiposHoraTrabaja  As Byte
Dim J As Integer

    Set miRsAux = New ADODB.Recordset
    
    
    
    
    
    
    Cad = "Select * from tiposhora ORDER BY TipoHora"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic
    CuantosTiposHoraTrabaja = 0
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & Mid(miRsAux!Desctipohora, 1, 3) & "|"
        CuantosTiposHoraTrabaja = CuantosTiposHoraTrabaja + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    For J = 1 To CuantosTiposHoraTrabaja
        Me.ListView1.ColumnHeaders.Add , , RecuperaValor(Cad, J), 800, 1
    Next
    
    Me.ListView1.Width = 6000 + (CuantosTiposHoraTrabaja * 800)
    Me.Width = Me.ListView1.Width + 120 + 240
   ' Me.cmdCancelar.Left = Me.Width - Me.cmdCancelar.Width - 240
   ' Me.cmdAceptar.Left = Me.cmdCancelar.Left - Me.cmdAceptar.Width - 240
    
    
    
    Set miRs = New ADODB.Recordset
    
    
    Cad = "select jornadassemanalesalz.idtrabajador,tipohoras,nomtrabajador,sum(horastrabajadas)  as totalh from jornadassemanalesalz,trabajadores  "
    Cad = Cad & " where jornadassemanalesalz.idtrabajador=trabajadores.idtrabajador"
    J = DiasMes(Month(Mes), Year(Mes))
    Cad = Cad & " AND fecha between '" & Format(Mes, FormatoFecha) & "' and '" & Format(Mes, "yyyy-mm-") & Format(J, "00") & "' group by 1,2 order by 1,2"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    
    idTrabajador = -1
    While Not miRsAux.EOF
        If miRsAux!idTrabajador <> idTrabajador Then
        

                
            
            
            Set IT = ListView1.ListItems.Add()
            IT.Text = miRsAux!idTrabajador
            IT.Tag = 0 'Trabajador
            IT.SubItems(1) = miRsAux!nomtrabajador
            IT.SubItems(2) = " "
            'El hco de horas
            For J = 0 To CuantosTiposHoraTrabaja - 1
                IT.SubItems(ColumnaDondeEmpiezanHoras + J) = " "
            Next J
    
            
            
            
            idTrabajador = miRsAux!idTrabajador
         
            


        End If

   
        'Que columna pinto
        IT.SubItems(ColumnaDondeEmpiezanHoras + miRsAux!TipoHoras) = Format(miRsAux!TotalH, "0.00")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Cad = "select idtrabajador,count(distinct(fecha))"
    Cad = Cad & " from jornadassemanalesalz  where "
    J = DiasMes(Month(Mes), Year(Mes))
    Cad = Cad & " fecha between '" & Format(Mes, FormatoFecha) & "' and '" & Format(Mes, "yyyy-mm-") & Format(J, "00") & "' group by 1"
   
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        For J = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(J).Text = CStr(miRsAux!idTrabajador) Then
                Me.ListView1.ListItems(J).SubItems(2) = miRsAux.Fields(1)
                Exit For
            End If
        Next J
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
       
    Set miRs = Nothing
    Set miRsAux = Nothing
    
    
End Sub




Private Function CargarDatosImpresion() As Boolean
Dim Aux As String
Dim Cad As String
Dim Byt As Byte
Dim Impor As Currency
Dim CuantosTiposHoraTrabaja As Byte '3
Dim Dias2 As Byte

    'tmpcombinada(IdTrabajador,Fecha,idinci,HT,HE,HR)
    NumRegElim = 1
    conn.Execute "DELETE FROM  tmpcombinada WHERE codusu = " & vUsu.Codigo
    

    Cad = ""
    CuantosTiposHoraTrabaja = 4
    For NumRegElim = 1 To Me.ListView1.ListItems.Count
            Aux = ""
            If ListView1.ListItems(NumRegElim).SubItems(2) = "" Then
                Dias2 = "1"
            Else
                Dias2 = ListView1.ListItems(NumRegElim).SubItems(2)
            End If
            For Byt = 0 To CuantosTiposHoraTrabaja - 2 'NO vamos a ver PACTADAS todavia
               
                Impor = ImporteFormateado(Trim(ListView1.ListItems(NumRegElim).SubItems(ColumnaDondeEmpiezanHoras + Byt)))


                ''tmpcombinada(IdTrabajador,Fecha,idinci,HT,HE,HR)
                Aux = Aux & "," & DBSet(Impor, "N", "N")
                
            Next Byt
            Cad = Cad & ", (" & vUsu.Codigo & "," & ListView1.ListItems(NumRegElim).Text & ",'1972-01-" & Dias2 & "',0" & Aux & ")"
            
                
            
        
    Next
    CargarDatosImpresion = True
    Cad = Mid(Cad, 2)
    Cad = "INSERT INTO tmpcombinada(codusu,IdTrabajador,Fecha,idinci,HT,HE,HR) VALUES " & Cad
    conn.Execute Cad
    

End Function



