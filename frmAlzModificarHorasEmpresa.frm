VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmAlzModificarHorasEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar horas empresa"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBottom 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   7335
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Error laborables!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   4335
      End
   End
   Begin VB.TextBox txtEmpre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   21
      Text            =   "0,00"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtEmpre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   18
      Text            =   "0,00"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkAlziraPermiteNoSumarOk 
      Caption         =   "Permitir cambiar suma final"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtEmpre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Text            =   "0,00"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtEmpre 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   4800
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   2160
      Width           =   1335
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   735
      Left            =   6240
      TabIndex        =   2
      Top             =   2160
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1296
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblTipoHora 
      Caption         =   "lblTipoHora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   24
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblNomEmpre 
      Caption         =   "Empiezan visible false"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   23
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label lblAuste 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   22
      Top             =   3960
      Width           =   795
   End
   Begin VB.Label lblNomEmpre 
      Caption         =   "Empiezan visible false"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblAuste 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblTipoHora 
      Alignment       =   1  'Right Justify
      Caption         =   "lblTipoHora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   17
      Top             =   690
      Width           =   4200
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "No se "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   -120
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "ven"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   -120
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblAuste 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   9
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label lblAuste 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   8
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label lblTipoHora 
      Caption         =   "lblTipoHora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   690
      Width           =   3735
   End
   Begin VB.Label lblNomEmpre 
      Caption         =   "Alzicoop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label lblNomEmpre 
      Caption         =   "segunda emre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblNombre 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmAlzModificarHorasEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public NumeroTotalAreasSubempresa As Byte  '0 Motilla 1 Alzira...
Public NombreAreasSubempresa As String ' nom1 | nom2| ....


Public idTrabajador As Long
Public TipoHora As Byte   '0normal  1 estructural  2 extra
Public Fecha As Date
Public AlmacenArea As Integer   'Tabla Areas.


Dim SQL As String
Dim TotalHoras As Currency
Dim Incremento As Currency
Dim Laborable As Byte  'Cuando salga tiene que sumar este adato.

Dim H1 As Currency
Dim h2 As Currency
Dim ind As Integer   'para los for ...next


Private Sub Command1_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 1 Then
        Unload Me
    Else
    
        H1 = ImporteFormateado(txtEmpre(1).Text)
        If lblAuste(1).Tag = "" Then
            H1 = 1
        Else
            H1 = H1 - CCur(lblAuste(1).Tag)
        End If
        If H1 = 0 Then
            'NO ha cambiado nada
            '
        Else
            
            If Not HacerModificaciones Then Exit Sub
            CadenaDesdeOtroForm = "OK"
        End If
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
    
    
    If SQL <> "" Then Exit Sub
    Label2.Visible = False
    lblTipoHora(2).Caption = Format(Fecha, "dd/mm/yyyy")
    
    chkAlziraPermiteNoSumarOk.Value = 0
    chkAlziraPermiteNoSumarOk.Visible = vEmpresa.QueEmpresa = 2
    
    Set miRsAux = New ADODB.Recordset
    SQL = "Select * from jornadassemanalesalz where idtrabajador = " & idTrabajador
    SQL = SQL & " AND tipohoras= " & TipoHora & " AND fecha =" & DBSet(Fecha, "F")
    SQL = SQL & " AND codarea= " & AlmacenArea
    SQL = SQL & " ORDER by ParaEmpresa"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then
        Unload Me
    Else
        '-----------------
        'Son dos columnas como mucho. Fruixeres(0) y alzira 1
        TotalHoras = 0
        Laborable = 0
        
        
        While Not miRsAux.EOF
            ind = miRsAux!paraempresa + 1   'Emp
                
            Me.txtEmpre(ind).Text = Format(miRsAux!HorasTrabajadas, FormatoImporte)
            
            lblAuste(ind).Tag = CCur(miRsAux!HorasTrabajadas)
            'Para saber si es ajustado, creado a mano...
            lblAuste(ind).Caption = miRsAux!Ajuste
            
            Laborable = Laborable + miRsAux!Laborable
            TotalHoras = TotalHoras + miRsAux!HorasTrabajadas
                    
                    
                    
            miRsAux.MoveNext
        Wend
        If Laborable > 1 Then Label2.Visible = True 'Indicara el error
        Me.txtTotal.Text = Format(TotalHoras, FormatoImporte)
        PonerFocoBtn Command1(1)
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Incremento = 0.25
    SQL = ""
    
    
    H1 = 0

    
    
    For ind = 1 To NumeroTotalAreasSubempresa
        Me.lblNomEmpre(ind).Caption = Mid(RecuperaValor(NombreAreasSubempresa, ind), 5) 'quito los pimeros 3 que son el ID
        Me.lblNomEmpre(ind).Tag = Mid(RecuperaValor(NombreAreasSubempresa, ind), 1, 4)
        Me.lblNomEmpre(ind).Visible = True
        Me.txtEmpre(ind).Visible = True
        Me.lblAuste(ind).Visible = False
    Next
    For ind = NumeroTotalAreasSubempresa + 1 To 4
        lblAuste(ind).Caption = "-1"
        lblAuste(ind).Tag = 0
        txtEmpre(ind).Text = ""
        Me.lblAuste(ind).Visible = False
    Next
    
    FrameBottom.Top = txtEmpre(NumeroTotalAreasSubempresa).Top + 600
    
    
End Sub




Private Sub txtEmpre_GotFocus(Index As Integer)
    ConseguirFocoLin txtEmpre(Index)
End Sub

Private Sub txtEmpre_KeyPress(Index As Integer, KeyAscii As Integer)
    
        KeyPress KeyAscii

End Sub
Private Sub txtEmpre_LostFocus(Index As Integer)


    If Index = 1 Then
        h2 = ImporteFormateado(txtEmpre(2).Text)
        If Not PonerFormatoDecimal(txtEmpre(1), 2) Then
            H1 = TotalHoras - h2
        Else
            If Me.chkAlziraPermiteNoSumarOk.Value = 1 Then Exit Sub
            H1 = ImporteFormateado(txtEmpre(1).Text)
            If H1 > TotalHoras Then H1 = TotalHoras
            h2 = TotalHoras - H1
            txtEmpre(2).Text = Format(h2, FormatoImporte)
        End If
        txtEmpre(1).Text = Format(H1, FormatoImporte)
        
    ElseIf Index = 2 Then
        H1 = ImporteFormateado(txtEmpre(1).Text)
        If Not PonerFormatoDecimal(txtEmpre(2), 2) Then
            h2 = TotalHoras - H1
        Else
            If Me.chkAlziraPermiteNoSumarOk.Value = 1 Then Exit Sub
            h2 = ImporteFormateado(txtEmpre(2).Text)
            If h2 > TotalHoras Then h2 = TotalHoras
            H1 = TotalHoras - h2
            txtEmpre(1).Text = Format(H1, FormatoImporte)
        End If
        txtEmpre(2).Text = Format(h2, FormatoImporte)
    End If
    
End Sub

Private Sub UpDown1_DownClick()
    H1 = ImporteFormateado(txtEmpre(1).Text)
    If H1 = 0 Then Exit Sub
    HacerIncremento True
End Sub

Private Sub UpDown1_UpClick()
    H1 = ImporteFormateado(txtEmpre(2).Text)
    If H1 = 0 Then Exit Sub
    HacerIncremento False
End Sub


Private Sub HacerIncremento(BajarHoras As Boolean)
Dim Aux As Currency
    H1 = ImporteFormateado(txtEmpre(1).Text)
    h2 = ImporteFormateado(txtEmpre(2).Text)
    If BajarHoras Then
        If H1 >= Incremento Then
            Aux = Incremento
        Else
            Aux = Incremento - H1
        End If
        H1 = H1 - Aux
        h2 = h2 + Aux
        
    Else
        If h2 >= Incremento Then
            Aux = Incremento
        Else
            Aux = Incremento - h2
        End If
        H1 = H1 + Aux
        h2 = h2 - Aux
    End If
    txtEmpre(1).Text = Format(H1, FormatoImporte)
    txtEmpre(2).Text = Format(h2, FormatoImporte)
End Sub

Private Sub KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub




Private Function HacerModificaciones() As Boolean
Dim QueAjuste As Integer
Dim TotalMenteManual As Boolean

    On Error GoTo eHacerModificaciones
    HacerModificaciones = False
    TotalMenteManual = False
    
    h2 = 0
    For ind = 1 To NumeroTotalAreasSubempresa
        H1 = ImporteFormateado(txtEmpre(ind).Text)
        h2 = h2 + H1
    Next
    

    If h2 <> TotalHoras Then
        If Me.chkAlziraPermiteNoSumarOk.Value = 0 Then
            MsgBox "Error en sumas de horas", vbExclamation
            Exit Function
    
        Else
            SQL = String(30, "*") & vbCrLf
            SQL = SQL & "Sumatorio de horas no coincide." & vbCrLf & "Anterior: " & TotalHoras & vbCrLf & "Actual: " & H1 + h2 & vbCrLf & "¿CONTINUAR?" & vbCrLf & SQL
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
            TotalMenteManual = True
        End If
    
    End If
    
    'AJUSTE
    '       0.- Sin ajustar
    '       1.- Se ajusto en proceso calculo de horas
    '       2.- Se creo a mano
    '       3.- Se modifico la que  estaba sin ajustar en proc horas
    '       4.- "            " del proceso de calculo de horas
    '       5.- "               la creada a mano
    '       7: TotalMenteManual Se ha cambiado sin respetar sumatorios
    'jornadassemanalesalz(idtrabajador,fecha,TipoHoras,horastrabajadas,ParaEmpresa,Ajuste)
    '------------------------------------------------------------------------------------

    
    For ind = 1 To NumeroTotalAreasSubempresa
        H1 = ImporteFormateado(txtEmpre(ind).Text)
        If H1 = 0 Then
            SQL = "DELETE from jornadassemanalesalz "
            SQL = SQL & " WHERE idTrabajador = " & idTrabajador & " AND fecha ="
            SQL = SQL & DBSet(Fecha, "F") & " AND ParaEmpresa=" & lblNomEmpre(ind).Tag
            SQL = SQL & "  AND TipoHoras=" & TipoHora
            SQL = SQL & "  AND codarea=" & AlmacenArea
            SQL = SQL & "  AND ParaEmpresa=" & lblNomEmpre(ind).Tag
        Else
            SQL = " VALUES (" & idTrabajador & "," & DBSet(Fecha, "F") & "," & TipoHora & "," & DBSet(H1, "N") & "," & lblNomEmpre(ind).Tag & ","
            If Me.lblAuste(ind).Caption = "-1" Then
                SQL = SQL & IIf(TotalMenteManual, 7, 5)  'CREADA A desde aqui
            Else
                SQL = SQL & IIf(TotalMenteManual, 7, 2)
            End If
            
             SQL = SQL & "," & Laborable & "," & AlmacenArea & ")"
            SQL = "REPLACE INTO jornadassemanalesalz(idtrabajador,fecha,TipoHoras,horastrabajadas,ParaEmpresa,Ajuste,laborable,codarea)" & SQL
            
            If Laborable > 0 Then Laborable = Laborable - 1
            If Laborable < 0 Then Laborable = 0
            
        End If
        conn.Execute SQL
    Next
    
    
    HacerModificaciones = True
eHacerModificaciones:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
End Function
