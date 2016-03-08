VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmAlzModificarHorasEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar horas empresa"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      Left            =   4920
      TabIndex        =   6
      Text            =   "0,00"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   1920
      Width           =   1335
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   855
      Left            =   6360
      TabIndex        =   2
      Top             =   1920
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1935
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
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
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
      Left            =   0
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
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
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
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
      Index           =   5
      Left            =   6720
      TabIndex        =   9
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label Label1 
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
      Left            =   6720
      TabIndex        =   8
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label1 
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Fruxeresa"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmAlzModificarHorasEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public idTrabajador As Long
Public TipoHora As Byte   '0normal  1 estructural  2 extra
Public Fecha As Date



Dim SQL As String
Dim TotalHoras As Currency
Dim Incremento As Currency
Dim Laborable As Byte  'Cuando salga tiene que sumar este adato.

Dim H1 As Currency
Dim h2 As Currency



Private Sub Command1_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 1 Then
        Unload Me
    Else
        H1 = ImporteFormateado(Text1(1).Text)
        If H1 = CCur(Label1(4).Tag) Then
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
    Text1(0).Text = Format(Fecha, "dd/mm/yyyy")
    Set miRsAux = New ADODB.Recordset
    SQL = "Select * from jornadassemanalesalz where idtrabajador = " & idTrabajador
    SQL = SQL & " AND tipohoras= " & TipoHora & " AND fecha =" & DBSet(Fecha, "F")
    SQL = SQL & " ORDER by ParaEmpresa"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then
        Unload Me
    Else
        '-----------------
        'Son dos columnas como mucho. Fruixeres(0) y alzira 1
        
        TotalHoras = 0
        Label1(4).Tag = CCur(0)
        Label1(5).Tag = CCur(0)
        Label1(6).Caption = "-1"
        Label1(7).Caption = "-1"
        Laborable = 0
        While Not miRsAux.EOF
            If miRsAux!paraempresa = 0 Then
                Laborable = Laborable + miRsAux!Laborable
                Text1(1).Text = Format(miRsAux!HorasTrabajadas, FormatoImporte)
                Me.Label1(4).Caption = Text1(1).Text
                Label1(4).Tag = CCur(miRsAux!HorasTrabajadas)
                'Para saber si es ajustado, creado a mano...
                Label1(6).Caption = miRsAux!Ajuste
            End If
            If miRsAux!paraempresa = 1 Then
                Laborable = Laborable + miRsAux!Laborable
                Text1(2).Text = Format(miRsAux!HorasTrabajadas, FormatoImporte)
                Me.Label1(5).Caption = Text1(2).Text
                Label1(5).Tag = CCur(miRsAux!HorasTrabajadas)
                'Para saber si es ajustado, creado a mano...
                Label1(7).Caption = miRsAux!Ajuste
            End If
            TotalHoras = TotalHoras + miRsAux!HorasTrabajadas
                    
                    
                    
            miRsAux.MoveNext
        Wend
        If Laborable > 1 Then Label2.Visible = True 'Indicara el error
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Incremento = 0.25
    SQL = ""

End Sub

Private Sub UpDown2_Change()

End Sub
Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFocoLin Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    
        KeyPress KeyAscii

End Sub
Private Sub Text1_LostFocus(Index As Integer)


    If Index = 1 Then
        h2 = ImporteFormateado(Text1(2).Text)
        If Not PonerFormatoDecimal(Text1(1), 2) Then
            H1 = TotalHoras - h2
        Else
            H1 = ImporteFormateado(Text1(1).Text)
            If H1 > TotalHoras Then H1 = TotalHoras
            h2 = TotalHoras - H1
            Text1(2).Text = Format(h2, FormatoImporte)
        End If
        Text1(1).Text = Format(H1, FormatoImporte)
        
    ElseIf Index = 2 Then
        H1 = ImporteFormateado(Text1(1).Text)
        If Not PonerFormatoDecimal(Text1(2), 2) Then
            h2 = TotalHoras - H1
        Else
            h2 = ImporteFormateado(Text1(2).Text)
            If h2 > TotalHoras Then h2 = TotalHoras
            H1 = TotalHoras - h2
            Text1(1).Text = Format(H1, FormatoImporte)
        End If
        Text1(2).Text = Format(h2, FormatoImporte)
    End If
    
End Sub

Private Sub UpDown1_DownClick()
    H1 = ImporteFormateado(Text1(1).Text)
    If H1 = 0 Then Exit Sub
    HacerIncremento True
End Sub

Private Sub UpDown1_UpClick()
    H1 = ImporteFormateado(Text1(2).Text)
    If H1 = 0 Then Exit Sub
    HacerIncremento False
End Sub


Private Sub HacerIncremento(BajarHoras As Boolean)
Dim Aux As Currency
    H1 = ImporteFormateado(Text1(1).Text)
    h2 = ImporteFormateado(Text1(2).Text)
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
    Text1(1).Text = Format(H1, FormatoImporte)
    Text1(2).Text = Format(h2, FormatoImporte)
End Sub

Private Sub KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub




Private Function HacerModificaciones() As Boolean
Dim QueAjuste As Integer

    On Error GoTo eHacerModificaciones
    HacerModificaciones = False
    H1 = ImporteFormateado(Text1(1).Text)
    h2 = ImporteFormateado(Text1(2).Text)
    If H1 + h2 <> TotalHoras Then
        MsgBox "Error en sumas de horas", vbExclamation
        Exit Function
    End If
    
    'AJUSTE
    '       0.- Sin ajustar
    '       1.- Se ajusto en proceso calculo de horas
    '       2.- Se creo a mano
    '       3.- Se modifico la que  estaba sin ajustar en proc horas
    '       4.- "            " del proceso de calculo de horas
    '       5.- "               la creada a mano
    
    'jornadassemanalesalz(idtrabajador,fecha,TipoHoras,horastrabajadas,ParaEmpresa,Ajuste)
    '------------------------------------------------------------------------------------
    'Las de fruixeresa
    QueAjuste = -1
    If Label1(6).Caption = "-1" Then
        'NUEVO. NO estaba creado
        SQL = " VALUES (" & idTrabajador & "," & DBSet(Fecha, "F") & "," & TipoHora & ","
        SQL = SQL & DBSet(H1, "N") & ",0,2," & Laborable & ")"
        If Laborable > 0 Then Laborable = 0
        SQL = "INSERT INTO jornadassemanalesalz(idtrabajador,fecha,TipoHoras,horastrabajadas,ParaEmpresa,Ajuste,laborable)" & SQL
    Else
        If Val(Label1(6).Caption) < 3 Then QueAjuste = Val(Label1(6).Caption) + 3
    
        SQL = "UPDATE jornadassemanalesalz SET horastrabajadas =" & DBSet(H1, "N")
        If QueAjuste > 0 Then SQL = SQL & ", ajuste =" & QueAjuste
        If Laborable > 0 Then
            If H1 = 0 Then
                SQL = SQL & ", Laborable  =0"  'se lo pondre a las horas cooperativas
            Else
                SQL = SQL & ", Laborable  =" & Laborable
                Laborable = 0
            End If
            
        End If
        SQL = SQL & " WHERE idTrabajador = " & idTrabajador & " AND fecha ="
        SQL = SQL & DBSet(Fecha, "F") & " AND ParaEmpresa=0  AND TipoHoras=" & TipoHora
        
    End If
    conn.Execute SQL
    
    
    '------------------------------------------------------------------------------------
    'las de la cooperativa
    QueAjuste = -1
    If Label1(7).Caption = "-1" Then
        'NUEVO. NO estaba creado
        SQL = " VALUES (" & idTrabajador & "," & DBSet(Fecha, "F") & "," & TipoHora & ","
        SQL = SQL & DBSet(h2, "N") & ",1,2," & Laborable & ")"
        SQL = "INSERT INTO jornadassemanalesalz(idtrabajador,fecha,TipoHoras,horastrabajadas,ParaEmpresa,Ajuste,laborable)" & SQL
    Else
        If Val(Label1(7).Caption) < 3 Then QueAjuste = Val(Label1(7).Caption) + 3
    
        SQL = "UPDATE jornadassemanalesalz SET horastrabajadas =" & DBSet(h2, "N")
        If QueAjuste > 0 Then SQL = SQL & ", ajuste =" & QueAjuste
        SQL = SQL & ", Laborable  =" & Laborable
        SQL = SQL & " WHERE idTrabajador = " & idTrabajador & " AND fecha ="
        SQL = SQL & DBSet(Fecha, "F") & " AND ParaEmpresa=1  AND TipoHoras=" & TipoHora
        
    End If
    conn.Execute SQL
    
    HacerModificaciones = True
eHacerModificaciones:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
End Function
