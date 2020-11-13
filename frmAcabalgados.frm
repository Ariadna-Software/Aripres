VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcabalgados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificar horas entre dias"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   375
      Index           =   0
      Left            =   11520
      TabIndex        =   3
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   12960
      TabIndex        =   2
      Top             =   8280
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Trabajador"
         Object.Width           =   7250
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "H1"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "H2"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "H3"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "H4"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "H5"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "H6"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "H7"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "H8"
         Object.Width           =   1693
      EndProperty
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13455
   End
End
Attribute VB_Name = "frmAcabalgados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Public Fecha As Date

Dim PrimVez As Boolean
Dim Cad As String

Private Sub Command1_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    If Index = 0 Then
        Cad = ""
        For NumRegElim = 1 To ListView1.ListItems.Count
            If Not ListView1.ListItems(NumRegElim).Checked Then
                Cad = Cad & ", " & ListView1.ListItems(NumRegElim).Text
            End If
            
        Next
        If Cad <> "" Then
            Cad = " WHERE  idtra IN (" & Mid(Cad, 2) & ")"
            conn.Execute "DELETE from tmpnotrabajo " & Cad
                        
            espera 1
            
        End If
        CadenaDesdeOtroForm = "OK"  'se por un tema U otro tiene que devolver true
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
Dim IT As ListItem
Dim Tr As Long
Dim i As Integer
Dim TieneUnoMayorHoraMenos8 As Boolean
Dim H1 As Date
Dim UnoDelDiaDeAntes As Boolean
Dim B As Boolean

    If PrimVez Then
        PrimVez = False
        Set miRsAux = New ADODB.Recordset
        ListView1.ListItems.Clear
        
        H1 = DateAdd("h", -8, vEmpresa.AcabalgadoHora)
        
        Cad = "select idtra,nomtrabajador,concat(hora,'')HoraText,if(hora>'23:59:59','23:59:59',hora) laHora"
        Cad = Cad & " ,if(hora<'0:00:00',1,0) HoraNegativa,hora from tmpnotrabajo,trabajadores ,entradafichajes where "
        Cad = Cad & " tmpnotrabajo.idtra=trabajadores.idtrabajador and trabajadores.idtrabajador ="
        Cad = Cad & " entradafichajes.idtrabajador AND fecha = " & DBSet(Fecha, "F") & " ORDER BY idtra,hora asc"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Tr = -1
        While Not miRsAux.EOF
            If miRsAux!idTRa <> Tr Then
            
                If Tr >= 0 Then
                    If vEmpresa.AcabalgadoDiaInicio Then TieneUnoMayorHoraMenos8 = Not TieneUnoMayorHoraMenos8
                    
                    If TieneUnoMayorHoraMenos8 Then
                        IT.Checked = True
                        IT.Bold = True
                        IT.ListSubItems(1).ForeColor = vbRed
                    End If
                End If
                Tr = miRsAux!idTRa
                Set IT = ListView1.ListItems.Add()
                IT.Text = Format(Tr, "0000")
                IT.SubItems(1) = miRsAux!nomtrabajador
                i = 2
                TieneUnoMayorHoraMenos8 = False
                UnoDelDiaDeAntes = False
            End If
            
            If miRsAux!HoraNegativa Then
                
                Cad = Horas_Quitar24(miRsAux!Hora, True)
                IT.SubItems(i) = Cad
                UnoDelDiaDeAntes = True
            Else
                If vEmpresa.AcabalgadoDiaInicio Then
                    If CDate(miRsAux!LaHora) >= H1 And CDate(miRsAux!LaHora) < vEmpresa.AcabalgadoHora Then TieneUnoMayorHoraMenos8 = True
                    
                   
                    
                Else
                    'Si el primer marcaje es mayor que las hora acabalgamiento, entonces lo pasamos
                    If CDate(miRsAux!LaHora) > vEmpresa.AcabalgadoHora Then
                        'Si es el primero entonces
                        If i = 2 Then
                            TieneUnoMayorHoraMenos8 = True
                        Else
                            
                            If UnoDelDiaDeAntes Then TieneUnoMayorHoraMenos8 = True
                        End If
                    End If
                End If
                IT.SubItems(i) = miRsAux!HoraText
            End If
            
            i = i + 1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        If Not IT Is Nothing Then
            If vEmpresa.AcabalgadoDiaInicio Then TieneUnoMayorHoraMenos8 = Not TieneUnoMayorHoraMenos8
            
            If TieneUnoMayorHoraMenos8 Then
                IT.Checked = True
                IT.Bold = True
                IT.ListSubItems(1).ForeColor = vbRed
            End If
        End If
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    PrimVez = True
    Label1.Caption = Format(Fecha, "dddd dd k mmmm k yyyy")
    Label1.Caption = "Dia proceso: " & Replace(Label1.Caption, "k", "de")
End Sub

