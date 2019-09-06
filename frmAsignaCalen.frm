VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAsignaHorario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar calendario"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   Icon            =   "frmAsignaCalen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9000
      TabIndex        =   11
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   120
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   6480
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "Trimestre"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mes"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Semana"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dia"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5775
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10186
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5775
      Left            =   6480
      TabIndex        =   0
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   10186
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignaCalen.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignaCalen.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignaCalen.frx":D64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignaCalen.frx":D968
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignaCalen.frx":DC82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignaCalen.frx":E21C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmAsignaHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Integer
    '0.- Sera para signar el calendario
    
    '1.- Datos un (UNO) trabajador, cambiarle horarios

    '2.-
    
    
    '3.- Asignacion calendario vacaciones.
    
Public OtrosDatos As String   'Dependera de la opcion
Public FeIni As Date
Public FeFin As Date

Dim IT As ListItem
Dim CambiosRealizados As Boolean
Dim FechaOpcion3 As Date


Private Sub Check1_Click(Index As Integer)
Dim op As Integer
    If CambiosRealizados Then
        If MsgBox("Las modificaciones realizadas se perderán. Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    If Opcion = 3 Then
        'NO DEJO QUE ESTEN JUNTOS la opcion dia/semana
        'Si esta uno no estara el otro
        If Index = 3 Then
            If Check1(3).Value = 1 Then Check1(2).Value = 0
            If Check1(2).Value = 1 Then Check1(3).Value = 0
        Else
            If Check1(2).Value = 1 Then Check1(3).Value = 0
            If Check1(3).Value = 1 Then Check1(2).Value = 0
        End If
     End If
    
    'Monto la opcion
    op = (Check1(0).Value * 2) + (Check1(1).Value * 4) + (Check1(2).Value * 8) + (Check1(3).Value * 16)
'    CargarDeOtraForma op
    CargaFechas op
End Sub



Private Sub cmdAceptar_Click()
Dim cad As String
Dim TodoAsignado As Boolean
    If Not CambiosRealizados Then Exit Sub
    
    
    TodoAsignado = TodoPeriodoAsignado(TreeView1.Nodes(1))
    
    If Opcion = 0 Then
        'Si es para asignar el horario
        miSQL = "Select count(*) from calendariol where fecha >= '" & Format(FeIni, FormatoFecha)
        miSQL = miSQL & "' AND fecha <='" & Format(FeFin, FormatoFecha)
        miSQL = miSQL & "' AND idcal = " & Text2.Tag
        Set miRs = New ADODB.Recordset
        miRs.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        miSQL = "Es la primera asignación de horarios. Deberia asignar horarios a todo el peridodo." & vbCrLf & _
            "¿Desea continuar?"
        If Not miRs.EOF Then
            If DBLet(miRs.EOF, "N") > 0 Then miSQL = ""
        End If
        miRs.Close
        Set miRs = Nothing
        
        
        'Si devuelve 0 significa que es la primera asignacion que hacemos
        If miSQL <> "" Then
            'Significa que no hay dias para ese calendario
            If Not TodoAsignado Then
                If MsgBox(miSQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            
        Else
            
        End If
        
        
    End If
    
    
    
    cad = "Las modificaciones serán aplicadas. Desea continuar?"
    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    

        UPDATEACalendario TreeView1.Nodes(1), TreeView1.Nodes(1).Text
        If Not TreeView1.Nodes(1).LastSibling Is Nothing Then
            UPDATEACalendario TreeView1.Nodes(1).LastSibling, TreeView1.Nodes(1).LastSibling.Text
            Unload Me
        End If


    If Opcion = 3 Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CambiosRealizados = False
    Set TreeView1.ImageList = Me.ImageList1
    Set ListView2.SmallIcons = Me.ImageList1
    If Opcion = 3 Then
        'ASIGNACION DE CALENDARIO DE VACACIONES
    
        Set IT = ListView2.ListItems.Add
        IT.Text = "Normal"
        IT.Tag = 0
        
        
        Set IT = ListView2.ListItems.Add
        IT.Text = "VACACIONES"
        IT.Tag = 1
        IT.Bold = True
        IT.ForeColor = vbGreen

        
        Check1(2).Value = 0
        Check1(3).Value = 1
        Text2.Tag = RecuperaValor(Me.OtrosDatos, 2)
        Text2.Text = RecuperaValor(Me.OtrosDatos, 1)
        FechaOpcion3 = RecuperaValor(Me.OtrosDatos, 3)
    Else
        'El resto de opciones
        miSQL = "Select idhorario,nomhorario from horarios order by idhorario"
        Set miRs = New ADODB.Recordset
        miRs.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRs.EOF
            Set IT = ListView2.ListItems.Add
            IT.Text = miRs!NomHorario
            IT.Tag = miRs!IdHorario
            IT.Bold = True
            If miRs!IdHorario < 15 Then
                IT.ForeColor = QBColor(miRs!IdHorario)
            Else
                IT.ForeColor = QBColor(0)
            End If
            IT.SmallIcon = 6
            miRs.MoveNext
        Wend
        miRs.Close
        Set miRs = Nothing
        If Opcion = 0 Then
            Text2.Tag = RecuperaValor(Me.OtrosDatos, 1)
            Text2.Text = RecuperaValor(Me.OtrosDatos, 2)
            
        Else
            If Opcion = 1 Then
                Text2.Tag = RecuperaValor(Me.OtrosDatos, 2)
                Text2.Text = RecuperaValor(Me.OtrosDatos, 1) & " (" & Text2.Tag & ")"
            End If
        End If
    
    End If  'De opcion=3
    
    Text1(0).Text = Format(FeIni, "dd/mm/yyyy")
    Text1(1).Text = Format(FeFin, "dd/mm/yyyy")
    Check1_Click 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListView2_BeforeLabelEdit(Cancel As Integer)
    ListView2.Drag vbCancel
End Sub



Private Sub CargaFechas(Opcion As Integer)

    TreeView1.Nodes.Clear

    If Year(FeIni) = Year(FeFin) Then
        CargarDeOtraForma Opcion, FeIni, FeFin
    Else
        'Cargaremos primero la parte de el año uno
        CargarDeOtraForma Opcion, FeIni, CDate("31/12/" & Year(FeIni))
        'Cargaremos primero la parte de el año DOS
        CargarDeOtraForma Opcion, CDate("01/01/" & Year(FeFin)), FeFin
    End If
End Sub



'
' Comparacion bit a bit
'
Private Sub CargarDeOtraForma(Opcion As Integer, FI As Date, FF As Date)
Dim N As Node
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim F As Date
Dim Intervalo As Integer
Dim Anterior As String
Dim Padre As String


    Set N = TreeView1.Nodes.Add(, , "A" & Year(FI), Year(FI))
    N.Image = 1
    N.Expanded = True
    Anterior = "A"
    
    'Trimestre
    If Opcion And 2 Then
        
        
        'Inicio
        F = FI
        i = ((Month(F) - 1) \ 3) + 1
        'Fin
        F = FF
        J = ((Month(F) - 1) \ 3) + 1
        
        
        
        For Intervalo = i To J
            If (Intervalo Mod 2) = 0 Then
                cad = "º"
            Else
                cad = "er"
            End If
            cad = Intervalo & cad & " TRIM."
            Set N = TreeView1.Nodes.Add("A" & Year(FI), tvwChild, "T" & Intervalo & Year(FI), cad)
            N.Image = 2
            N.Expanded = True
        
        Next
        Anterior = "T"
    End If
    
        
    'MES
    If Opcion And 4 Then
        
        F = CDate(FI)
        i = Month(F)
        'Fin
        F = CDate(FF)
        J = Month(F)
        
        For Intervalo = i To J

            If Anterior = "A" Then
                Padre = "A" & Year(FI)
            Else
                Padre = (((Intervalo - 1) \ 3)) + 1
                Padre = "T" & Padre & Year(FI)
            End If
                
            cad = Format("01/" & Intervalo & "/2006", "mmmm")
            Set N = TreeView1.Nodes.Add(Padre, tvwChild, "M" & Intervalo & Year(FI), cad)
            N.Image = 3
        Next Intervalo
        Anterior = "M"
    End If
        
        
    'Semana
    If Opcion And 8 Then
        F = CDate(FI)
        
        i = 0
        
        While F <= CDate(FF)
            
            If i = 0 Then
                'ES el primer dia. Veremos la semana primera cuantos dias tiene
                Intervalo = Format(F, "w", vbMonday)
                Intervalo = 7 - Intervalo
            Else
                Intervalo = 6
            End If
            
            If Anterior = "A" Then
                Padre = "A"
            Else
                If Anterior = "T" Then
                    Padre = (((Month(F) - 1) \ 3)) + 1
                    Padre = "T" & Padre
                Else
                    Padre = "M" & Month(F)
                End If
            End If
            Padre = Padre & Year(FI)
            i = Format(F, "ww", vbMonday)
            
            cad = "Sem: " & Format(i, "00")
            cad = cad & "    " & Format(F, "dd/mm") & " - "
            F = DateAdd("d", Intervalo, F)  'El ultimo de la semana
            cad = cad & Format(F, "dd/mm")
            F = DateAdd("d", 1, F) 'el dia siguiente
            Set N = TreeView1.Nodes.Add(Padre, tvwChild, "S" & Format(i, "00") & Year(FI), cad)
            N.Image = 4
          
            
        
        
        Wend
                
        Anterior = "S"
    End If
    
    Intervalo = 0
    If Opcion And 16 Then
        

    
        i = TreeView1.Nodes.Count
        F = FI
        While F <= FF
            i = i + 1
            If Anterior = "A" Then
                Padre = "A"
            Else
                If Anterior = "T" Then
                    Padre = (Month(F) \ 4) + 1
                    Padre = "T" & Padre
                Else
                    If Anterior = "M" Then
                        Padre = "M" & Month(F)
                        
                    Else
                        Padre = Format(F, "ww", vbMonday)
                       
                        Padre = "S" & Format(Padre, "00")
                    End If
                End If
            End If
            Padre = Padre & Year(F)
            
            Set N = TreeView1.Nodes.Add(Padre, tvwChild, "D" & i, Format(F, "dd/mmm"))
            N.Image = 5
            F = DateAdd("d", 1, F)
                        
        Wend
    
    End If
    

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    PonerFoco Text1(Index)
End Sub



Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta", vbExclamation
        PonerFoco Text1(Index)
    Else
        Check1_Click 0
    End If
End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    ListView2.Drag vbEndDrag
End Sub

Private Sub TreeView1_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ListView2.Drag vbEndDrag
    If TreeView1.DropHighlight Is Nothing Then
          Set TreeView1.DropHighlight = Nothing
          Exit Sub
       Else
            If ListView2.SelectedItem Is Nothing Then Exit Sub
        
       
            ' comprobamos si lo podemos insertar y si es asi lo borramos de
            ' alli y lo metemos aqui  la carpeta no debe incluir subdirectorios
            miSQL = "Desea asignar el horario: " & ListView2.SelectedItem & vbCrLf
            miSQL = miSQL & " al periodo: "
            OtrosDatos = Mid(TreeView1.DropHighlight.Key, 1, 1)
            Select Case OtrosDatos
            Case "A"
                miSQL = miSQL & " AÑO : " & TreeView1.DropHighlight
            
            Case "T"
                miSQL = miSQL & " TRIMESTRE : " & TreeView1.DropHighlight
                
            Case "M"
                miSQL = miSQL & " MES : " & TreeView1.DropHighlight
                
            Case "S"
                miSQL = miSQL & " SEMANA : " & TreeView1.DropHighlight
                
            Case "D"
                miSQL = miSQL & " DIA : " & TreeView1.DropHighlight
                
            Case Else
                MsgBox "NODO incorrecto.", vbCritical
                Exit Sub
            End Select
            If MsgBox(miSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            CambiosRealizados = True
            HacerASignacionHorarios TreeView1.DropHighlight
             ListView2.SelectedItem.Selected = False
             Set ListView2.SelectedItem = Nothing
    End If
    Set TreeView1.DropHighlight = Nothing
End Sub




Private Sub HacerASignacionHorarios(N As Node)
Dim No As Node
    N.Bold = True
    N.ForeColor = ListView2.SelectedItem.ForeColor
    N.Tag = ListView2.SelectedItem.Tag
    If N.Children > 0 Then
        Set No = N.Child
        Do
            HacerASignacionHorarios No
            If No.Next Is Nothing Then
                Set No = Nothing
            Else
                Set No = No.Next
            End If
        Loop Until (No Is Nothing)
    End If
End Sub

Private Sub TreeView1_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
    'If Not TreeView1.DropHighlight Is Nothing Then TreeView1.DropHighlight.Expanded = True
End Sub



Private Function TodoPeriodoAsignado(N As Node) As Boolean
Dim No As Node
Dim B As Boolean
    
    TodoPeriodoAsignado = False

    If N.Children > 0 Then
        Set No = N.Child
        Do
            If N.Children > 0 Then
                B = TodoPeriodoAsignado(No)
                If Not B Then
                    TodoPeriodoAsignado = False
                    Exit Function
                End If
            Else
                If N.Bold Then TodoPeriodoAsignado = True
            End If
            If No.Next Is Nothing Then
                Set No = Nothing
            Else
                Set No = No.Next
            End If
        Loop Until (No Is Nothing)
        TodoPeriodoAsignado = True
    Else
        If N.Bold Then TodoPeriodoAsignado = True
    End If
    
    
    
End Function




Private Sub UPDATEACalendario(N As Node, vAno As String)
Dim No As Node

    
    

    If N.Children > 0 Then
        Set No = N.Child
        Do
            If N.Children > 0 Then
                UPDATEACalendario No, vAno
            Else
                If N.Bold Then HacerSQLUpdateCalendario N, vAno
            End If
            If No.Next Is Nothing Then
                Set No = Nothing
            Else
                Set No = No.Next
            End If
        Loop Until (No Is Nothing)
    Else
        If N.Bold Then HacerSQLUpdateCalendario N, vAno
    End If
    
    
    
End Sub

Private Sub HacerSQLUpdateCalendario(ByRef Nodo As Node, anyo As String)
Dim F1 As Date
Dim F2 As Date
Dim F3 As Date
Dim i As Integer
Dim J As Integer
Dim FESTIVOS As String
Dim C As String
Dim Aux As String
Dim C3 As String


    Select Case Mid(Nodo.Key, 1, 1)
    Case "A"
        'TOOOOOOODO el año
         F1 = CDate("01/01/" & Nodo.Text)
         F2 = CDate("31/12/" & Nodo.Text)
        
    Case "T"
        '
         i = Val(Mid(Nodo.Key, 2, 1))
         F1 = CDate("01/01/" & anyo)
         
         F1 = DateAdd("m", 3 * (i - 1), F1)   'LE sumo los trimestres correspondiente -1
         J = Month(F1) + 2
         i = DiasMes(J, Year(F1))
         F2 = CDate(i & "/" & J & "/" & Year(F1))
        
    Case "M"
        
        J = Mid(Nodo.Key, 2)
        F1 = CDate("01/" & J & "/" & anyo)
        i = DiasMes(J, CInt(TreeView1.Nodes(1).Text))
        F2 = CDate(i & "/" & J & "/" & anyo)
    Case "S"
        'Semana
        F1 = CDate("01/01/" & anyo)
        
        'Dias que caben en la primera semana +1
        i = Format(F1, "w", vbMonday) - 1
        i = (7 - i)
        
        J = Val(Mid(Nodo.Key, 2, 2))
        If J > 1 Then
            J = ((J - 2) * 7) + i 'los dias de desplazamiento
            i = 6
        Else
            J = 0
            i = i - 1
        End If
       
        
        
        
        F1 = DateAdd("d", J, F1)
        
        F2 = DateAdd("d", i, F1)
            
        
        
        
        
        
        
        
    
    Case "D"
        
        F1 = CDate(Nodo.Text & "/" & anyo)
        F2 = F1
    End Select
    
    
    Select Case Opcion
    Case 0
        'UPDATEA CALENDARIO. LAS LINESA
        
        'Ahora deberiamos updatear los dias para los trabajadores que tengan asignado ese horario
        
        If vEmpresa.CreaCalDiariaTra Then
            'EN alzira NO llevo dia a dia los trabajadores. Hay un horario conjunto
            Set miRsAux = New ADODB.Recordset
            miSQL = "SELECT idtrabajador from trabajadores where idcal =" & Val(Text2.Tag)
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'Borramos
                miSQL = "DELETE FROM calendariot where idtrabajador = " & miRsAux!idTrabajador
                miSQL = miSQL & " AND fecha >=' " & Format(F1, FormatoFecha) & "' AND fecha <='" & Format(F2, FormatoFecha) & "'"
                
                conn.Execute miSQL
                    
                J = 1
                F3 = F1
                While F3 <= F2
                    If J = 1 Then
                        miSQL = "INSERT INTO calendariot (idtrabajador, fecha, idhorario, TipoDia) VALUES "
                    Else
                        miSQL = miSQL & ","
                    End If
                    
                    miSQL = miSQL & "(" & miRsAux!idTrabajador & ",'" & Format(F3, FormatoFecha) & "'," & Nodo.Tag & ",0)"
                    If J = 155 Then
                        conn.Execute miSQL & ";"
                        J = 1
                    Else
                        J = J + 1
                    End If
                    F3 = DateAdd("d", 1, F3)
                Wend
                If J > 1 Then conn.Execute miSQL & ";"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
        End If
        
        miSQL = "DELETE from calendariol where idcal = " & Text2.Tag
        miSQL = miSQL & " AND fecha >=' " & Format(F1, FormatoFecha) & "' AND fecha <='" & Format(F2, FormatoFecha) & "'"
        conn.Execute miSQL
        miSQL = "INSERT INTO calendariol (idcal,  idhorario,fecha) VALUES (" & Val(Text2.Tag) & "," & Nodo.Tag & ",'"
         C = ""
        
        
    Case 1, 2
    
        Set miRsAux = New ADODB.Recordset
        FESTIVOS = ""
        miSQL = "Select fecha from calendariot where idtrabajador =" & Text2.Tag
        miSQL = miSQL & " AND fecha >=' " & Format(F1, FormatoFecha) & "' AND fecha <='" & Format(F2, FormatoFecha) & "'"
        miSQL = miSQL & " AND tipodia =1 "
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            FESTIVOS = FESTIVOS & Format(miRsAux!Fecha, "dd/mm/yyyy") & "|"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If FESTIVOS <> "" Then FESTIVOS = "|" & FESTIVOS
        'UPDATE TRABAJADOR
        miSQL = "DELETE from calendariot where idtrabajador = " & Text2.Tag
        miSQL = miSQL & " AND fecha >=' " & Format(F1, FormatoFecha) & "' AND fecha <='" & Format(F2, FormatoFecha) & "'"
        conn.Execute miSQL
        miSQL = "INSERT INTO calendariot (idtrabajador, idhorario, TipoDia,Fecha) VALUES (" & Text2.Tag & "," & Nodo.Tag & ","
        C = ""
    Case 3
       
        If F2 <= FechaOpcion3 Then
            'FECHA ULTIMO marcaje menor que la fecha fin del intervalo
            Exit Sub
        End If
        FESTIVOS = ""
        miSQL = "UPDATE calendariot SET TipoDia=" & Nodo.Tag & " WHERE (idTrabajador =" & Text2.Tag & " AND Fecha = '"
        C = ""
    End Select
    
    'Noviembre 2013
    DoEvents
    Aux = ""
    J = InStr(1, miSQL, " VALUES (") + 8
    
    While F1 <= F2
        If Opcion = 0 Then
           
        Else
            If Opcion = 1 Or Opcion = 2 Then
                C = "|" & Format(F1, "dd/mm/yyyy") & "|"
                If InStr(1, FESTIVOS, C) > 0 Then
                    C = "1,'"
                Else
                    C = "0,'"
                End If
            Else
                'OPCION=3
                
                
            End If
        End If
        
        'Noviembre 2013
        
        C3 = ", " & Mid(miSQL, J) & C & Format(F1, FormatoFecha) & "')"
        Aux = Aux & C3
        'conn.Execute miSQL & C & Format(F1, FormatoFecha) & "')"
        If Len(Aux) > 4000 Then
                Aux = Mid(Aux, 2) 'quitamos la primera coma
                C3 = Mid(miSQL, 1, J - 1)
                C3 = C3 & Aux
                conn.Execute C3
                Aux = ""
        End If
        
        F1 = DateAdd("d", 1, F1)
        
    Wend
   
    If Len(Aux) > 0 Then
            
        Aux = Mid(Aux, 2) 'quitamos la primera coma
        C3 = Mid(miSQL, 1, J - 1)
        C3 = C3 & Aux
        conn.Execute C3
    End If
End Sub
