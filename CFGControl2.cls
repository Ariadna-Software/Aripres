VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFGControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarBaseDatos As String 'copia local
Private mvarDirMarcajes As String 'copia local
Private mvarDirProcesados As String 'copia local
Private mvarNomFich As String 'copia local
Private mvarTCP3 As Boolean 'Si tiene TCP3 de Lipsoft Electronics
Private mvarAriadna As Boolean  'Si hace integracion con aplicaciones Ariadna Sw,
    'Supondremos que si no tiene el terminal los ficheros que nos pasan son
    'como los de alzicoop
    'Ademas, si no tiene TCP3 no se puede acceder al menu de operaciones con TCP3
    'Al igual qu si no tiene Integracion Ariadna no puede acceder al menu de integrar

Public Property Let Ariadna(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Navegador = 5
    mvarAriadna = vData
End Property


Public Property Get Ariadna() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Navegador
    Ariadna = mvarAriadna
End Property

Public Property Get DigitoTrabajadores() As Integer

End Property


Public Function Guardar() As Byte
    Dim NF As Integer
    Dim F As Integer
    Dim Leido As String
    Dim Linea As String
    Dim Param As String
    Dim Valor As String
    Dim I As Integer
    Dim vec(5) As Boolean
        'Para saber si algun parametro no se ha puesto
        'El v(0) sera la BD y el v(4) el de NomFich
        'Cuando acabemos de pasar el fichero comprobaremos si hay que poner todos los valores
          
    For I = 0 To 5
        vec(I) = False
    Next I
    
    Guardar = 1
    On Error GoTo ErrGuardar
    F = FreeFile
    '-- Control de existencia del fichero
    If Dir(App.Path & "\ControlInf.cfg") = "" Then
        Open App.Path & "\ControlInf.cfg" For Output As #F   ' Abre el archivo.
        Linea = " BaseDatos = " & mvarBaseDatos
        Print #F, Linea
        Linea = " DirMarcajes = " & mvarDirMarcajes
        Print #F, Linea
        Linea = " DirProcesados = " & mvarDirProcesados
        Print #F, Linea
        Linea = " DirNomFich = " & mvarNomFich
        Print #F, Linea
        Linea = " TCP3 = " & mvarTCP3
        Print #F, Linea
        Linea = " IntegracionAriadna = " & mvarAriadna
        Print #F, Linea
        Close #F
    Else
        NF = F + 1
        Open App.Path & "\ControlInf.cfg" For Input As #NF   ' Abre el archivo.
        Open App.Path & "\ControlInf2.cfg" For Output As #F   ' Abre el archivo.
        Do While Not EOF(NF)   ' Repite el bucle hasta el final del archivo.
            Line Input #NF, Leido   ' Lee el registro en la variable.
            Leido = Trim(Leido)
            If Leido = "" Or Left(Leido, 1) = "#" Then ' Se desprecian las líneas en blanco y comentarios
                'Tiene comentarios o es linea en blanco
                Linea = Leido
                'ELSE
                Else
                    I = InStr(1, Leido, "=")
                    If I > 0 Then
                        Param = Trim(Mid(Leido, 1, I - 1))
                        Valor = Trim(Mid(Leido, I + 1, Len(Leido) - I))
                        Select Case Param
                            Case "BaseDatos"
                                Linea = " BaseDatos = " & mvarBaseDatos
                                vec(0) = True
                            Case "DirMarcajes"
                                Linea = " DirMarcajes = " & mvarDirMarcajes
                                vec(1) = True
                            Case "DirProcesados"
                                Linea = " DirProcesados = " & mvarDirProcesados
                                vec(2) = True
                            Case "NomFich"
                                Linea = " NomFich = " & mvarNomFich
                                vec(3) = True
                            Case "TCP3"
                                Linea = " TCP3 = " & mvarTCP3
                                vec(4) = True
                            Case "IntegracionAriadna"
                                Linea = " IntegracionAriadna = " & mvarAriadna
                                vec(5) = True
                            Case Else
                                '-- No se reconoce el parámetro
                                Linea = "# " & Leido
                          End Select
                        'ELSE i>0
                        Else
                            '-- falta el signo de igualdad en el parámetro
                            'Comentamos la linea
                            Linea = "# " & Leido
                    End If
            End If
            'Escribimos en F
            Print #F, Linea
        Loop
        
        
        'Ahora comprobamos si todos los valores han sido colocados
        For I = 0 To 5
        
            If vec(I) = False Then
                Select Case I
                    Case 0     '"BaseDatos"
                        If mvarBaseDatos <> "" Then _
                        Linea = " BaseDatos = " & mvarBaseDatos
                    Case 1   '"DirMarcajes"
                        If mvarDirMarcajes <> "" Then _
                        Linea = " DirMarcajes = " & mvarDirMarcajes
                    Case 2       '"DirProcesados"
                        If mvarDirProcesados <> "" Then _
                        Linea = " DirProcesados = " & mvarDirProcesados
                    Case 3      '"NomFich"
                        If mvarNomFich <> "" Then _
                        Linea = " NomFich = " & mvarNomFich
                    Case 4      ' TCP3
                        Linea = " TCP3 = " & mvarTCP3
                    Case 5
                        Linea = " IntegracionAriadna = " & mvarAriadna
                End Select
                Print #F, Linea
            End If
        Next I
        Close #NF   ' Cierra el archivo.
        Close #F
        'Ahora, con los archivo cerrados procedemos a cambiar los valores
        FileCopy App.Path & "\ControlInf.cfg", App.Path & "\Control.BAK"
        FileCopy App.Path & "\ControlInf2.cfg", App.Path & "\ControlInf.cfg"
        Kill App.Path & "\ControlInf2.cfg"
    End If
    Guardar = 0
Exit Function
ErrGuardar:
    MsgBox "No se ha podido guardar la configuración. " & vbCrLf & "Número: " & Err.Number & vbCrLf & _
        "Descripción: " & Err.Description, vbExclamation
End Function

Public Property Let DirProcesados(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DirProcesados = 5
    mvarDirProcesados = vData
End Property


Public Property Get DirProcesados() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DirProcesados
    DirProcesados = mvarDirProcesados
End Property

Public Property Let NomFich(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DirProcesados = 5
    mvarNomFich = vData
End Property


Public Property Get NomFich() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DirProcesados
    NomFich = mvarNomFich
End Property


Public Property Let TCP3(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Navegador = 5
    mvarTCP3 = vData
End Property


Public Property Get TCP3() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Navegador
    TCP3 = mvarTCP3
End Property



Public Property Let BaseDatos(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.BaseDatos = 5
    mvarBaseDatos = vData
End Property


Public Property Get BaseDatos() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.BaseDatos
    BaseDatos = mvarBaseDatos
End Property

Public Property Let DirMarcajes(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DirInforme = 5
    mvarDirMarcajes = vData
End Property


Public Property Get DirMarcajes() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DirInforme
    DirMarcajes = mvarDirMarcajes
End Property
















Public Function Leer() As Byte
  '-- Al inicializarse la clase se cargan los valores de parámetros correspon
    '   dientes, se comprueba igualmente que el fichero de configuración exista
    Dim NF As Integer
    Dim Leido As String
    Dim Param As String
    Dim Valor As String
    Dim I As Integer
    '-- Control de existencia del fichero
    If Dir(App.Path & "\ControlInf.cfg") = "" Then
        ' El fichero no existe
        Leer = 1
        MsgBox "Falta el fichero de configuración " & App.Path & "\ControlInf.cfg", vbCritical, "Club Carrefour"
    Else
        NF = FreeFile
        Open App.Path & "\ControlInf.cfg" For Input As #NF   ' Abre el archivo.
        Do While Not EOF(NF)   ' Repite el bucle hasta el final del archivo.
            Line Input #NF, Leido   ' Lee el registro en la variable.
            If Leido <> "" Or Left(Leido, 1) <> "#" Then ' Se desprecian las líneas en blanco y comentarios
                I = InStr(1, Leido, "=")
                If I > 0 Then
                    Param = Trim(Mid(Leido, 1, I - 1))
                    Valor = Trim(Mid(Leido, I + 1, Len(Leido) - I))
                    Select Case Param
                        Case "BaseDatos"
                            mvarBaseDatos = Valor
                        Case "DirMarcajes"
                            mvarDirMarcajes = Valor
                        Case "DirProcesados"
                            mvarDirProcesados = Valor
                        Case "NomFich"
                            mvarNomFich = Valor
                        Case "TCP3"
                            mvarTCP3 = Valor
                        Case "IntegracionAriadna"
                            mvarAriadna = Valor
                    
                        Case Else
                            '-- No se reconoce el parámetro
                    End Select
                Else
                    '-- falta el signo de igualdad en el parámetro
                End If
                                     
            End If
        Loop
        Close #NF   ' Cierra el archivo.
        Leer = 0
    End If

End Function

