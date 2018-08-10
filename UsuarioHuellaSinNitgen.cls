VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UsuarioHuella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarCodUsuario As String 'copia local
Private mvarFIR As String 'copia local
Private mvarMensaje As String 'copia local
Private mvarGesLabID As Long 'copia local
Private SQL As String
Private Rs As ADODB.Recordset

Public Property Let GesLabID(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.GesLabID = 5
    mvarGesLabID = vData
End Property


Public Property Get GesLabID() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.GesLabID
    GesLabID = mvarGesLabID
End Property



Public Property Let Mensaje(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Mensaje = 5
    mvarMensaje = vData
End Property


Public Property Get Mensaje() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Mensaje
    Mensaje = mvarMensaje
End Property



Public Property Let FIR(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FIR = 5
    mvarFIR = vData
End Property


Public Property Get FIR() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FIR
    FIR = mvarFIR
End Property



Public Property Let CodUsuario(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CodUsuario = 5
    mvarCodUsuario = vData
    '-- Forzamos a que la longitud sea 10 con ceros por delante
    mvarCodUsuario = Right(String(10, "0") & Trim(mvarCodUsuario), 10)
End Property


Public Property Get CodUsuario() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CodUsuario
    CodUsuario = mvarCodUsuario
End Property

Public Function Leer(vCodUsuario As String, Optional Comprobando As Boolean) As Boolean
End Function

Public Function Guardar() As Boolean
End Function

Public Sub CapturaHuella(byt As Byte)


End Sub
