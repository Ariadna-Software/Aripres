VERSION 5.00
Begin VB.Form frmRelojZKTeco 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZKTeco"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAjustarReloj 
      Caption         =   "UltFecLeida"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Leer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   3450
      Left            =   0
      Picture         =   "frmZKTeco.frx":0000
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "frmRelojZKTeco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Cad As String
Dim NF As Integer

Dim UtlFecLeidaCarpeta As Date

Private Sub cmdAjustarReloj_Click()
    
        
    For NumRegElim = 1 To 2
        'Carpeta 1
        Cad = DevuelveDesdeBD("ultimaFechaLeidaCarpeta" & NumRegElim, "relojZK", "1", "1")
        If Cad = "" Then Cad = "12/04/1972 15:00:00"
        UtlFecLeidaCarpeta = CDate(Cad)
        Cad = "ULTIMA FECHA/HORA.     Reloj: " & NumRegElim
        Cad = InputBox(Cad, "RELOJ " & NumRegElim, CStr(UtlFecLeidaCarpeta))
        If Cad <> "" Then
            If IsDate(Cad) Then
                Cad = "UPDATE relojZK set ultimaFechaLeidaCarpeta" & NumRegElim & "  =" & DBSet(Cad, "FH")
                EjecutaSQL Cad
            Else
                MsgBox "Fecha incorrecta", vbExclamation
            End If
        End If
    
    
    Next
End Sub

Private Sub Command1_Click()
    
    Unload Me
   
End Sub

Private Sub Command2_Click()

    

    Screen.MousePointer = vbHourglass
    If LeerRelojes Then
        
        Label1.Caption = "Lectura realizada correctamente"
        Command2.Enabled = False
    Else
        Me.Label1.Caption = ""
    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Label1.Caption = ""
    cmdAjustarReloj.Visible = vUsu.Codigo = 0
End Sub


Private Function LeerRelojes() As Boolean
Dim B As Boolean

    LeerRelojes = False


    'Borramos de la temporal en aripres
    conn.Execute "DELETE FROM tmppresencia where codusu =" & vUsu.Codigo
    
    Cad = DevuelveDesdeBD("max(secuencia)", "entradafichajes", "1", "1")
    NumRegElim = Val(Cad) + 1
    
    'Carpeta 1
    Cad = DevuelveDesdeBD("ultimaFechaLeidaCarpeta1", "relojZK", "1", "1")
    If Cad = "" Then Cad = "12/04/1972 15:00:00"
    UtlFecLeidaCarpeta = CDate(Cad)
    B = LeerDatos(1)
    
    
    'Carpeta 2
    Cad = DevuelveDesdeBD("ultimaFechaLeidaCarpeta2", "relojZK", "1", "1")
    If Cad = "" Then Cad = "12/04/1972 15:00:00"
    UtlFecLeidaCarpeta = CDate(Cad)
    B = B And LeerDatos(2)
    
 
    'If Not B Then Exit Function
    
    
    Label1.Caption = "Entradas repetidas"
    Label1.Refresh
    espera 0.25
    
    EntradasRepetidasProceso Me.Label1
    
    
    'Nov 2018
    HorasNocturnas Me.Label1
    espera 0.25
    
    LeerRelojes = True
End Function


'
'El reloj tiene un demonio que a las 23:59 , todos los dias, lee los datos del reloj.
' Y los guarda en un path: \\sultan\presencia\*.cop
'
' Y cuando el programa ZT, le decimos leer marcajes, lo hace en una carpeta local.

Private Function LeerDatos(Opcion As Byte) As Boolean
Dim Carpeta As String
Dim cArc As Collection
Dim fil As Collection
Dim i As Integer
Dim J As Integer
Dim Fic As String
Dim UltFecLeidaFichero As Date
Dim NumFich As Long

    On Error GoTo eLeerDatos
    LeerDatos = False
    
    Label1.Caption = "Comprobando carpeta"
    Label1.Refresh
    
    'OPCION:
    '   1:  Sera en la carpeta local. PUEDE que el equipo que esta leyendo NO tenga la carpeta local
    '   2:  \\sultan\presencia  TIENE que ser accisble. NO SE BORRARAN
    
    If Opcion = 1 Then
        Carpeta = vEmpresa.DirMarcajes
        If Dir(Carpeta, vbDirectory) = "" Then
            Err.Clear
            Exit Function
        End If
            
    Else
        Carpeta = DevuelveDesdeBD("configreloj", "empresas", "1", 1, "N")
        If Dir(Carpeta, vbDirectory) = "" Then Err.Raise 513, , "Falta configurar carpeta en servidor : " & Carpeta
    End If
    LeerDatos = True
    Set cArc = New Collection
    
    Cad = Dir(Carpeta & "\" & vEmpresa.NomFich, vbArchive)    ' Recupera la primera entrada.
    Do While Cad <> ""   ' Inicia el bucle.
        Fic = Cad
        Label1.Caption = "Ver fic: " & Fic
        Label1.Refresh
        
        'Si la opcion es 2, el fichero tiene formato numerico. YYYYMM
        If Opcion = 1 Then
            NumFich = 0
        Else
            J = InStr(1, Cad, ".")
            NumFich = Mid(Cad, 1, J - 1)
        End If
        
        Cad = Carpeta & "\" & Cad
        If preComprobacion(Opcion, NumFich) Then
            cArc.Add CStr(Fic)
        Else
            If Cad <> "" Then Err.Raise 513, Cad, Cad
        End If
      
       Cad = Dir   ' Obtiene siguiente entrada.
    Loop

        
    Set fil = New Collection
    
    
    For i = 1 To cArc.Count
    
    
        'En la opcion2 NO te
        If Opcion = 1 Then
    
            Label1.Caption = "Bloquear fichero: " & cArc.Item(i)
            Label1.Refresh
            'Renombramos para bloquear el fichero
            Cad = cArc.Item(i)
            J = InStrRev(Cad, ".")
            Cad = Mid(Cad, 1, J) & "ari"
            Cad = Carpeta & "\" & Cad
            Name Carpeta & "\" & cArc.Item(i) As Cad
            fil.Add CStr(Cad)
               
        Else
            
            Cad = cArc.Item(i)
            Cad = Carpeta & "\" & Cad
            fil.Add CStr(Cad)
        End If
    Next
   
    
    
    If fil.Count > 0 Then
        
        DoEvents
        UltFecLeidaFichero = CDate("12/04/1972")
        
        For i = 1 To fil.Count
           
            Cad = fil.Item(i)
            If InsertarDatos Then
                If Opcion = 1 Then
                    'En local
                    Cad = fil.Item(i)
                    MatarFichero CStr(Cad)
                End If
                
                'Vemos la ultima fecha leida
                Cad = DevuelveDesdeBD("concat(fecha,' ',h1)", "tmppresencia", "codusu ", vUsu.Codigo & "   order by 1  desc")
                If Cad <> "" Then
                    If CDate(Cad) > UltFecLeidaFichero Then UltFecLeidaFichero = Cad
                End If
                
            End If
            
        Next i
        'Ult fec leida
        If UltFecLeidaFichero <> CDate("12/04/1972") Then
            Cad = "UPDATE relojZK set ultimaFechaLeidaCarpeta" & Opcion & "  =" & DBSet(UltFecLeidaFichero, "FH")
            conn.Execute Cad
            
        End If
        
    End If
    
    
    If cArc.Count = 0 Then MsgBox "Ningun fichero a procesar (" & Carpeta & ")", vbExclamation
    
eLeerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    LeerDatos = False
    Set fil = Nothing
    Set cArc = Nothing
End Function



Private Function preComprobacion(Opcion As Byte, NomFichero As Long) As Boolean
Dim Tr As String
Dim AUX As String
    On Error GoTo epreComprobacion
    
    If Opcion = 2 Then
        'Como en el servidor NO podemos borrar fichero ni nada, lo que haremos sera NO porcesar
        
        If NomFichero < Val(Format(UtlFecLeidaCarpeta, "yyyymm")) Then
            Cad = "" 'PARA QUE NO DE ERROR
            Exit Function
        End If
    End If
    NF = FreeFile
    Tr = "|"
    Open Cad For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Cad
        'Primera linea: <PSD Copy File v.2.0>
        'Lineas de marca: 100720181228  Ejemplo len=12
        If UCase(Mid(Cad, 1, 5)) <> "<PSD " Then
            If Len(Cad) > 12 Then
                'Ejemplo. Va con tabulaciones
                ' tr  fe          hora
                '5   20180710    124851  1   0
                Cad = Mid(Cad, 1, InStr(1, Cad, Chr(9)) - 1)
                If InStr(1, Tr, "|" & Cad & "|") = 0 Then Tr = Tr & Cad & "|"
        
            End If
        End If
    Wend
    Close #NF
    
    Tr = Mid(Tr, 2) 'quitamos el primer |
    If Len(Tr) > 1 Then
        Do
            NF = InStr(1, Tr, "|")
            If NF = 0 Then
                Tr = ""
            Else
                Cad = Mid(Tr, 1, NF - 1)
                Tr = Mid(Tr, NF + 1)
            
                AUX = DevuelveDesdeBD("idtrabajador", "trabajadores", "numtarjeta", Cad, "T")
                If AUX = "" Then
                    Cad = "NO existe trabajador en BD: " & Cad & vbCrLf
                    Exit Function
                End If
            End If
        Loop Until Tr = ""
    End If
    preComprobacion = True
    Exit Function
epreComprobacion:
    Cad = Err.Description
    Err.Clear
    If NF > 0 Then Close #NF
End Function




Private Function InsertarDatos() As Boolean
Dim AUX As String
Dim Inser As String
Dim J As Integer

    On Error GoTo epreComprobacion
    
    Label1.Caption = Cad
    Label1.Refresh
    espera 0.2
    
    
    
    
    NF = FreeFile
    Open Cad For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Cad
        'Primera linea: <PSD Copy File v.2.0>
        'Lineas de marca: 100720181228  Ejemplo len=12
        If UCase(Mid(Cad, 1, 5)) <> "<PSD " Then
            If Len(Cad) > 12 Then
                'Ejemplo. Va con tabulaciones
                ' tr  fe          hora
                '5   20180710    124851  1   0
                'tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias)
                    
                J = InStr(1, Cad, Chr(9))
                AUX = Mid(Cad, 1, J - 1)
                Cad = Mid(Cad, J + 1)
                NumRegElim = NumRegElim + 1
                '
                Inser = Inser & ", (" & NumRegElim & "," & AUX & ","
                AUX = Mid(Cad, 1, 8)
                AUX = Mid(AUX, 7, 2) & "/" & Mid(AUX, 5, 2) & "/" & Mid(AUX, 1, 4)
                Inser = Inser & DBSet(AUX, "F") & ",'"
                AUX = Mid(Cad, 10, 6)
                AUX = Mid(AUX, 1, 2) & ":" & Mid(AUX, 3, 2) & ":" & Mid(AUX, 5, 2)
                Inser = Inser & AUX & "'," & vUsu.Codigo & ",0)"
                
            End If
        End If
    Wend
    Close #NF
    NF = -1
    
    Cad = "DELETE FROM tmppresencia where codusu =" & vUsu.Codigo
    conn.Execute Cad
    espera 0.1
        
    
    
    If Inser <> "" Then
        Label1.Caption = "Insertando tmp"
        Label1.Refresh
        espera 0.2
        Inser = Mid(Inser, 2)
        Inser = "insert tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias) VALUES " & Inser
        conn.Execute Inser
        espera 0.5
    
    
        'Nos cargamos los datos anteriores a la ultima vez leidos
        Cad = "DELETE from tmppresencia where fecha < " & DBSet(UtlFecLeidaCarpeta, "F")
        conn.Execute Cad
        Cad = "DELETE from tmppresencia where fecha = " & DBSet(UtlFecLeidaCarpeta, "F") & " AND h1 <=" & DBSet(UtlFecLeidaCarpeta, "H")
        conn.Execute Cad
    
    
    
        
        Set miRsAux = New ADODB.Recordset
        
        Cad = "    select distinct idtra from tmppresencia WHERE codusu = " & vUsu.Codigo
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = miRsAux!idTRa
            AUX = DevuelveDesdeBD("idtrabajador", "trabajadores", "numtarjeta", Cad, "T")
            If AUX = "" Then Err.Raise 513, , "NO existe trabajador. Tarjeta: " & Cad
            AUX = "UPDATE tmppresencia set seccion =" & AUX & " WHERE codusu =" & vUsu.Codigo & " AND idtra=" & miRsAux!idTRa
            conn.Execute AUX
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        

           
        Cad = "tmppresencia where codusu =" & vUsu.Codigo
        Cad = "SELECT id,seccion,fecha ,h1,incidencias,h1,0 relj FROM " & Cad & " ORDER BY fecha,h1"
        Cad = "INSERT INTO entradafichajes(Secuencia,idTrabajador,Fecha,Hora,idInci,HoraReal,reloj) " & Cad
        conn.Execute Cad
    
    
    End If
        
    InsertarDatos = True
    Exit Function
epreComprobacion:
    MuestraError Err.Number, Err.Description
    If NF > 0 Then Close #NF
    Set miRsAux = Nothing
End Function

Private Sub MatarFichero(Origen As String)
Dim J As Integer
Dim Secu As Long

    On Error GoTo eM
    
    
    'Mataremos el fichero siempre que el añomes sea menor que el actual
    J = InStrRev(Origen, "\")
    If J > 0 Then
        Cad = Mid(Origen, J + 1)
        J = InStr(1, Cad, ".")
        
        Cad = Mid(Cad, 1, J - 1)
        '                   'No lleva el dia.
        If Len(Cad) = 6 Then
            Secu = Cad & "99"
        Else
            Secu = Cad
        End If
        If Secu < Val(Format(Now, "yyyymmdd")) Then
            J = 1
        Else
            J = 0  'El del mes en curso no lo borro
        End If
    Else
        MsgBox "ERROR GRAVE. No tiene \"
    End If
    
    If J > 0 Then
        Cad = vEmpresa.DirProcesados & "\" & Format(Now, "yyyymmdd_hhnnss") & ".dat"
    
        FileCopy Origen, Cad
        Kill Origen
    Else
        
        'Le volvemos a poner el .cop
        J = InStrRev(Origen, "\")
        Cad = Mid(Origen, 1, J) & Cad & ".cop"
        Name Origen As Cad
    End If
    Exit Sub
eM:
    MsgBox "Error eliminando fichero. El programa continuará. Avise soporte técnico", vbCritical
    Err.Clear

End Sub
