VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColVacaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen vacaciones"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   495
      Left            =   11520
      TabIndex        =   1
      Top             =   7440
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   12303
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Utilizados"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Solicitados"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Pendientes"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmColVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

    If Me.Tag = 1 Then
        Screen.MousePointer = vbHourglass
        Me.Tag = 0
        CargaDatos -1
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Me.Tag = 1
End Sub




Private Sub CargaDatos(idTrabajador As Long)
Dim C As String
Dim Dias As String
Dim IT As ListItem
    On Error GoTo eCargaDatos
    
    
    Set miRsAux = New ADODB.Recordset
    
    C = "select  t.IdTrabajador ,NomTrabajador,DiasVacaciones, sum(if(situacion=1,1,0)) disfrutados ,sum(if(situacion=0,1,0)) solicitados"
    C = C & " from trabajadores t left join trabajadoresvacaciones v on t.idtrabajador=v.idtrabajador  and "
    C = C & " fecha between " & DBSet(vEmpresa.FechaInicio, "F") & "  and " & DBSet(vEmpresa.FechaFin, "F")
    C = C & " Where T.DiasVacaciones > 0 "
    If idTrabajador >= 0 Then C = C & "  AND t.idtrabajador=" & idTrabajador
    C = C & " group by 1"
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If idTrabajador > 0 Then
            Set IT = ListView1.SelectedItem
        Else
            Set IT = ListView1.ListItems.Add()
        End If
        IT.Text = Format(miRsAux.Fields(0), "0000")
        IT.SubItems(1) = miRsAux!nomtrabajador
        IT.SubItems(2) = miRsAux!DiasVacaciones
                
        Dias = miRsAux!DiasVacaciones - miRsAux!disfrutados
        IT.SubItems(3) = IIf(miRsAux!disfrutados > 0, miRsAux!disfrutados, " ")
        IT.SubItems(4) = IIf(miRsAux!solicitados > 0, miRsAux!solicitados, " ")
        IT.SubItems(5) = Dias
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
eCargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    If ListView1.SortKey = ColumnHeader.Index - 1 Then
        If ListView1.SortOrder = lvwAscending Then
            ListView1.SortOrder = lvwDescending
        Else
            ListView1.SortOrder = lvwAscending
        End If
    Else
        ListView1.SortKey = ColumnHeader.Index - 1
        ListView1.SortOrder = lvwAscending
    End If
End Sub

Private Sub ListView1_DblClick()
Dim Trab As Long
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Trab = CLng(Me.ListView1.SelectedItem.Text)
    frmCalendarioVacaciones.Trabajador = Trab
    frmCalendarioVacaciones.Show vbModal
    CargaDatos Trab
End Sub
