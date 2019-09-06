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


Dim cad As String
Dim NF As Integer

Dim UtlFecLeidaCarpeta As Date

Private Sub cmdAjustarReloj_Click()
    
        
    For NumRegElim = 1 To 2
        'Carpeta 1
        cad = DevuelveDesdeBD("ultimaFechaLeidaCarpeta" & NumRegElim, "relojZK", "1", "1")
        If cad = "" Then cad = "12/04/1972 15:00:00"
        UtlFecLeidaCarpeta = CDate(cad)
        cad = "ULTIMA FECHA/HORA.     Reloj: " & NumRegElim
        cad = InputBox(cad, "RELOJ " & NumRegElim, CStr(UtlFecLeidaCarpeta))
        If cad <> "" Then
            If IsDate(cad) Then
                cad = "UPDATE relojZK set ultimaFechaLeidaCarpeta" & NumRegElim & "  =" & DBSet(cad, "FH")
                EjecutaSQL cad
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
    
    cad = DevuelveDesdeBD("max(secuencia)", "entradafichajes", "1", "1")
    NumRegElim = Val(cad) + 1
    
    'Carpeta 1
    cad = DevuelveDesdeBD("ultimaFechaLeidaCarpeta1", "relojZK", "1", "1")
    If cad = "" Then cad = "12/04/1972 15:00:00"
    UtlFecLeidaCarpeta = CDate(cad)
    B = LeerDatos(1)
    
    
    'Carpeta 2
    cad = DevuelveDesdeBD("ultimaFechaLeidaCarpeta2", "relojZK", "1", "1")
    If cad = "" Then cad = "12/04/1972 15:00:00"
    UtlFecLeidaCarpeta = CDate(cad)
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
    
    cad = Dir(Carpeta & "\" & vEmpresa.NomFich, vbArchive)    ' Recupera la primera entrada.
    Do While cad <> ""   ' Inicia el bucle.
        Fic = cad
        Label1.Caption = "Ver fic: " & Fic
        Label1.Refresh
        
        'Si la opcion es 2, el fichero tiene formato numerico. YYYYMM
        If Opcion = 1 Then
            NumFich = 0
        Else
            J = InStr(1, cad, ".")
            NumFich = Mid(cad, 1, J - 1)
        End If
        
        cad = Carpeta & "\" & cad
        If preComprobacion(Opcion, NumFich) Then
            cArc.Add CStr(Fic)
        Else
            If cad <> "" Then Err.Raise 513, cad, cad
        End If
      
       cad = Dir   ' Obtiene siguiente entrada.
    Loop

        
    Set fil = New Collection
    
    
    For i = 1 To cArc.Count
    
    
        'En la opcion2 NO te
        If Opcion = 1 Then
    
            Label1.Caption = "Bloquear fichero: " & cArc.Item(i)
            Label1.Refresh
            'Renombramos para bloquear el fichero
            cad = cArc.Item(i)
            J = InStrRev(cad, ".")
            cad = Mid(cad, 1, J) & "ari"
            cad = Carpeta & "\" & cad
            Name Carpeta & "\" & cArc.Item(i) As cad
            fil.Add CStr(cad)
               
        Else
            
            cad = cArc.Item(i)
            cad = Carpeta & "\" & cad
            fil.Add CStr(cad)
        End If
    Next
   
    
    
    If fil.Count > 0 Then
        
        DoEvents
        UltFecLeidaFichero = CDate("12/04/1972")
        
        For i = 1 To fil.Count
           
            cad = fil.Item(i)
            If InsertarDatos Then
                If Opcion = 1 Then
                    'En local
                    cad = fil.Item(i)
                    MatarFichero CStr(cad)
                End If
                
                'Vemos la ultima fecha leida
                cad = DevuelveDesdeBD("concat(fecha,' ',h1)", "tmppresencia", "codusu ", vUsu.Codigo & "   order by 1  desc")
                If cad <> "" Then
                    If CDate(cad) > UltFecLeidaFichero Then UltFecLeidaFichero = cad
                End If
                
            End If
            
        Next i
        'Ult fec leida
        If UltFecLeidaFichero <> CDate("12/04/1972") Then
            cad = "UPDATE relojZK set ultimaFechaLeidaCarpeta" & Opcion & "  =" & DBSet(UltFecLeidaFichero, "FH")
            conn.Execute cad
            
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
Dim Aux As String
    On Error GoTo epreComprobacion
    
    If Opcion = 2 Then
        'Como en el servidor NO podemos borrar fichero ni nada, lo que haremos sera NO porcesar
        
        If NomFichero < Val(Format(UtlFecLeidaCarpeta, "yyyymm")) Then
            cad = "" 'PARA QUE NO DE ERROR
            Exit Function
        End If
    End If
    NF = FreeFile
    Tr = "|"
    Open cad For Input As #NF
    While Not EOF(NF)
        Line Input #NF, cad
        'Primera linea: <PSD Copy File v.2.0>
        'Lineas de marca: 100720181228  Ejemplo len=12
        If UCase(Mid(cad, 1, 5)) <> "<PSD " Then
            If Len(cad) > 12 Then
                'Ejemplo. Va con tabulaciones
                ' tr  fe          hora
                '5   20180710    124851  1   0
                cad = Mid(cad, 1, InStr(1, cad, Chr(9)) - 1)
                If InStr(1, Tr, "|" & cad & "|") = 0 Then Tr = Tr & cad & "|"
        
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
                cad = Mid(Tr, 1, NF - 1)
                Tr = Mid(Tr, NF + 1)
            
                Aux = DevuelveDesdeBD("idtrabajador", "trabajadores", "numtarjeta", cad, "T")
                If Aux = "" Then
                    cad = "NO existe trabajador en BD: " & cad & vbCrLf
                    Exit Function
                End If
            End If
        Loop Until Tr = ""
    End If
    preComprobacion = True
    Exit Function
epreComprobacion:
    cad = Err.Description
    Err.Clear
    If NF > 0 Then Close #NF
End Function




Private Function InsertarDatos() As Boolean
Dim Aux As String
Dim Inser As String
Dim J As Integer

    On Error GoTo epreComprobacion
    
    Label1.Caption = cad
    Label1.Refresh
    espera 0.2
    
    
    
    
    NF = FreeFile
    Open cad For Input As #NF
    While Not EOF(NF)
        Line Input #NF, cad
        'Primera linea: <PSD Copy File v.2.0>
        'Lineas de marca: 100720181228  Ejemplo len=12
        If UCase(Mid(cad, 1, 5)) <> "<PSD " Then
            If Len(cad) > 12 Then
                'Ejemplo. Va con tabulaciones
                ' tr  fe          hora
                '5   20180710    124851  1   0
                'tmppresencia(Id,idtra,Fecha,H1,codusu,Incidencias)
                    
                J = InStr(1, cad, Chr(9))
                Aux = Mid(cad, 1, J - 1)
                cad = Mid(cad, J + 1)
                NumRegElim = NumRegElim + 1
                '
                Inser = Inser & ", (" & NumRegElim & "," & Aux & ","
                Aux = Mid(cad, 1, 8)
                Aux = Mid(Aux, 7, 2) & "/" & Mid(Aux, 5, 2) & "/" & Mid(Aux, 1, 4)
                Inser = Inser & DBSet(Aux, "F") & ",'"
                Aux = Mid(cad, 10, 6)
                Aux = Mid(Aux, 1, 2) & ":" & Mid(Aux, 3, 2) & ":" & Mid(Aux, 5, 2)
                Inser = Inser & Aux & "'," & vUsu.Codigo & ",0)"
                
            End If
        End If
    Wend
    Close #NF
    NF = -1
    
    cad = "DELETE FROM tmppresencia where codusu =" & vUsu.Codigo
    conn.Execute cad
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
        cad = "DELETE from tmppresencia where fecha < " & DBSet(UtlFecLeidaCarpeta, "F")
        conn.Execute cad
        cad = "DELETE from tmppresencia where fecha = " & DBSet(UtlFecLeidaCarpeta, "F") & " AND h1 <=" & DBSet(UtlFecLeidaCarpeta, "H")
        conn.Execute cad
    
    
    
        
        Set miRsAux = New ADODB.Recordset
        
        cad = "    select distinct idtra from tmppresencia WHERE codusu = " & vUsu.Codigo
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = miRsAux!idTRa
            Aux = DevuelveDesdeBD("idtrabajador", "trabajadores", "numtarjeta", cad, "T")
            If Aux = "" Then Err.Raise 513, , "NO existe trabajador. Tarjeta: " & cad
            Aux = "UPDATE tmppresencia set seccion =" & Aux & " WHERE codusu =" & vUsu.Codigo & " AND idtra=" & miRsAux!idTRa
            conn.Execute Aux
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        

           
        cad = "tmppresencia where codusu =" & vUsu.Codigo
        cad = "SELECT id,seccion,fecha ,h1,incidencias,h1 FROM " & cad & " ORDER BY fecha,h1"
        cad = "INSERT INTO entradafichajes(Secuencia,idTrabajador,Fecha,Hora,idInci,HoraReal) " & cad
        conn.Execute cad
    
    
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
        cad = Mid(Origen, J + 1)
        J = InStr(1, cad, ".")
        
        cad = Mid(cad, 1, J - 1)
        '                   'No lleva el dia.
        If Len(cad) = 6 Then
            Secu = cad & "99"
        Else
            Secu = cad
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
        cad = vEmpresa.DirProcesados & "\" & Format(Now, "yyyymmdd_hhnnss") & ".dat"
    
        FileCopy Origen, cad
        Kill Origen
    Else
        
        'Le volvemos a poner el .cop
        J = InStrRev(Origen, "\")
        cad = Mid(Origen, 1, J) & cad & ".cop"
        Name Origen As cad
    End If
    Exit Sub
eM:
    MsgBox "Error eliminando fichero. El programa continuará. Avise soporte técnico", vbCritical
    Err.Clear

End Sub
