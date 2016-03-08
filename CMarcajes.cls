VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMarcajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarEntrada As Long 'copia local
Private mvaridTrabajador As Long 'copia local
Private mvarFecha As Date 'copia local
Private mvarCorrecto As Boolean 'copia local
Private mvarIncFinal As Integer 'copia local
Private mvarHorasTrabajadas As Currency 'copia local
Private mvarHorasIncid As Currency 'copia local
Private mvarIdHorario As Integer
Private mvarHorasDto As Currency 'copia local
Private mvarFestivo As Boolean
Private mvarBaja As Boolean
Private mvarNocturno As Boolean

Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As ADODB.Recordset
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo CMarcajes
'------------------------------------------------


Public Property Let Entrada(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Entrada = 5
     mvarEntrada = vData
End Property

Public Property Get Entrada() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Entrada
     Entrada = mvarEntrada
End Property


Public Property Let idTrabajador(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.idTrabajador = 5
     mvaridTrabajador = vData
End Property

Public Property Get idTrabajador() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.idTrabajador
     idTrabajador = mvaridTrabajador
End Property



Public Property Let IdHorario(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.idTrabajador = 5
     mvarIdHorario = vData
End Property

Public Property Get IdHorario() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.idTrabajador
     IdHorario = mvarIdHorario
End Property


Public Property Let Fecha(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Fecha = 5
     mvarFecha = vData
End Property

Public Property Get Fecha() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Fecha
     Fecha = mvarFecha
End Property


Public Property Let Correcto(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Correcto = 5
     mvarCorrecto = vData
End Property

Public Property Get Correcto() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Correcto
     Correcto = mvarCorrecto
End Property


Public Property Let IncFinal(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncFinal = 5
     mvarIncFinal = vData
End Property

Public Property Get IncFinal() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncFinal
     IncFinal = mvarIncFinal
End Property


Public Property Let HorasTrabajadas(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HorasTrabajadas = 5
     mvarHorasTrabajadas = vData
End Property

Public Property Get HorasTrabajadas() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HorasTrabajadas
     HorasTrabajadas = mvarHorasTrabajadas
End Property


Public Property Let HorasIncid(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HorasIncid = 5
     mvarHorasIncid = vData
End Property

Public Property Get HorasIncid() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HorasIncid
     HorasIncid = mvarHorasIncid
End Property

Public Property Let HorasDto(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HorasIncid = 5
     mvarHorasDto = vData
End Property

Public Property Get HorasDto() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HorasIncid
     HorasDto = mvarHorasDto
End Property



Public Property Let Festivo(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HorasIncid = 5
     mvarFestivo = vData
End Property

Public Property Get Festivo() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HorasIncid
     Festivo = mvarFestivo
End Property



Public Property Let Baja(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HorasIncid = 5
     mvarBaja = vData
End Property

Public Property Get Baja() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HorasIncid
     Baja = mvarBaja
End Property


Public Property Let Nocturno(ByVal vData As Boolean)
     mvarNocturno = vData
End Property

Public Property Get Nocturno() As Boolean
     Nocturno = mvarNocturno
End Property
'------------------------------------------------
'FUNCIONES del modulo CMarcajes
'------------------------------------------------



'****   AGREGAR  ***
Public Function Agregar() As Byte
    On Error GoTo Err1
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    SQL = "SELECT * FROM Marcajes"
    SQL = SQL & " WHERE "
    SQL = SQL & " Entrada = " & mvarEntrada

    RS.Open SQL, conn, , , adCmdText

    If Not RS.EOF Then
        Mens = "Ya existe un Registro en la Tabla Marcajes con esos valores."
        MsgBox Mens, vbExclamation
        Agregar = 1
        'ELSE
        Else
            RS.AddNew
            RS!Entrada = mvarEntrada
            RS!idTrabajador = mvaridTrabajador
            RS!Fecha = mvarFecha
            RS!Correcto = Abs(mvarCorrecto)
            RS!IncFinal = mvarIncFinal
            RS!IdHorario = mvarIdHorario
            RS!HorasTrabajadas = mvarHorasTrabajadas
            RS!HorasDto = mvarHorasDto
            RS!HorasIncid = mvarHorasIncid
            RS!Festivo = Abs(mvarFestivo)
            RS!Baja = Abs(mvarBaja)
            RS!Nocturno = Abs(mvarNocturno)
            RS.Update
            Agregar = 0
    End If
    RS.Close
    Exit Function

Err1:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Agregar = 1
End Function


'****   ELIMINAR  ***
Public Function Eliminar() As Byte
    On Error GoTo Err2
    SQL = "SELECT * FROM Marcajes"
    SQL = SQL & " WHERE "
    SQL = SQL & " Entrada = " & mvarEntrada
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open SQL, conn, , , adCmdText
    If RS.EOF Then
        Mens = "No existe un Registro en la Tabla Marcajes con esos valores."
        MsgBox Mens, vbExclamation
        Eliminar = 1
        'ELSE
        Else
            'Primero eliminamos las entradas en los fichajes
            SQL = "Delete * from EntradaMarcajes where IdMarcaje=" & mvarEntrada
            conn.Execute SQL
            'Las incidencias generadas
            SQL = "Delete * from IncidenciasGeneradas where EntradaMarcaje=" & mvarEntrada
            conn.Execute SQL

            '-----------------------------------------------------
            RS.Delete
            RS.Update
            RS.MoveNext
            Eliminar = 0
    End If

    RS.Close
    Exit Function

Err2:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Eliminar = 1
End Function


'****   LEER  ***
Public Function Leer(vIdEntrada As Long) As Byte
    On Error GoTo Err3
    SQL = "SELECT * FROM Marcajes"
    SQL = SQL & " WHERE "
    SQL = SQL & " Entrada = " & vIdEntrada
    RS.Open SQL, conn, , , adCmdText
    If RS.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarEntrada = RS!Entrada
            mvaridTrabajador = RS!idTrabajador
            mvarFecha = RS!Fecha
            mvarCorrecto = RS!Correcto
            mvarIncFinal = RS!IncFinal
            mvarHorasTrabajadas = RS!HorasTrabajadas
            mvarHorasIncid = RS!HorasIncid
            mvarHorasDto = RS!HorasDto
            mvarIdHorario = RS!IdHorario
            mvarFestivo = RS!Festivo = 1
            mvarBaja = RS!Baja = 1
            mvarNocturno = RS!Nocturno = 1
            Leer = 0
    End If

    RS.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Leer = 1
End Function


'****   MODIFICAR  ***
Public Function Modificar() As Byte
    On Error GoTo Err4
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    SQL = "SELECT * FROM Marcajes"
    SQL = SQL & " WHERE "
    SQL = SQL & " Entrada = " & mvarEntrada
    RS.Open SQL, conn, , , adCmdText
    If RS.EOF Then
        Mens = "No existe el Registro en la Tabla Marcajes con esos valores."
        MsgBox Mens, vbExclamation
        Modificar = 1
        'ELSE
        Else
            'Rs.Edit
            RS!idTrabajador = mvaridTrabajador
            RS!Fecha = mvarFecha
            RS!Correcto = Abs(mvarCorrecto)
            RS!IncFinal = mvarIncFinal
            RS!HorasTrabajadas = mvarHorasTrabajadas
            RS!HorasIncid = mvarHorasIncid
            RS!HorasDto = mvarHorasDto
            RS!Festivo = Abs(mvarFestivo)
            RS!Baja = Abs(mvarBaja)
            RS!Nocturno = Abs(mvarNocturno)
            RS!IdHorario = mvarIdHorario
            RS.Update
            Modificar = 0
    End If

    RS.Close
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Modificar = 1
End Function


'****   SIGUIENTE  ***
Public Function Siguiente() As Long

    Siguiente = 1
    SQL = "SELECT MAX(Entrada) FROM Marcajes"
    RS.Open SQL, conn, , , adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Siguiente = RS.Fields(0) + 1
    End If
    mvarEntrada = Siguiente
    RS.Close
End Function


Public Function PonerCorrecta() As Byte
    On Error GoTo Err4
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    SQL = "SELECT * FROM Marcajes"
    SQL = SQL & " WHERE "
    SQL = SQL & " Entrada = " & mvarEntrada
    RS.Open SQL, conn, , , adCmdText
    If RS.EOF Then
        Mens = "No existe el Registro en la Tabla Marcajes con esos valores."
        MsgBox Mens, vbExclamation
        PonerCorrecta = 1
        'ELSE
        Else
            'Rs.Edit
            RS!Correcto = True
            RS!HorasTrabajadas = mvarHorasTrabajadas
            RS!HorasIncid = mvarHorasIncid
            RS.Update
            PonerCorrecta = 0
    End If
    RS.Close
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    PonerCorrecta = 1

End Function


Private Sub Class_Initialize()
Set RS = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
Set RS = Nothing
End Sub
