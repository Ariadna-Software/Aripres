VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrabajadorNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvarCodigo As Long 'copia local
Private mvarNombre As String 'copia local
Private mvarHOficiales As Currency 'copia local
Private mvarDiasOficiales As Integer 'copia local
Private mvarDiasReales As Integer 'copia local

'variables locales para almacenar los valores de las propiedades
Private mvarFecAlta As Date 'copia local
Private mvarFecBaja As Date 'copia local
Private mvarHNReales As Currency 'copia local
Private mvarHEReales As Currency 'copia local


Private mvarHBolsa As Currency

Private mvarHorasCompensadasNomina As Currency
Private mvarDiasCompensables As Currency

Public Property Let HEReales(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HEReales = 5
    mvarHEReales = vData
End Property


Public Property Get HEReales() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HEReales
    HEReales = mvarHEReales
End Property



Public Property Let HNReales(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HNReales = 5
    mvarHNReales = vData
End Property


Public Property Get HNReales() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HNReales
    HNReales = mvarHNReales
End Property


Public Property Let DiasCompensables(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiasOficiales = 5
    mvarDiasCompensables = vData
End Property


Public Property Get DiasCompensables() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiasOficiales
    DiasCompensables = mvarDiasCompensables
End Property



Public Property Let HorasCompensadasNomina(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HNReales = 5
    mvarHorasCompensadasNomina = vData
End Property


Public Property Get HorasCompensadasNomina() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HNReales
    HorasCompensadasNomina = mvarHorasCompensadasNomina
End Property



Public Property Let HBolsa(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HNReales = 5
    mvarHBolsa = vData
End Property


Public Property Get HBolsa() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HNReales
    HBolsa = mvarHBolsa
End Property



Public Property Let FecBaja(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FecBaja = 5
    mvarFecBaja = vData
End Property




Public Property Get FecBaja() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FecBaja
    FecBaja = mvarFecBaja
End Property



Public Property Let FecAlta(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FecAlta = 5
    mvarFecAlta = vData
End Property


Public Property Get FecAlta() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FecAlta
    FecAlta = mvarFecAlta
End Property



Public Property Let DiasReales_(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiasOficiales = 5
    mvarDiasReales = vData
End Property


Public Property Get DiasReales_() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiasOficiales
    DiasReales_ = mvarDiasReales
End Property



Public Property Let DiasOficiales(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiasOficiales = 5
    mvarDiasOficiales = vData
End Property


Public Property Get DiasOficiales() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiasOficiales
    DiasOficiales = mvarDiasOficiales
End Property



Public Property Let HOficiales(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HOficiales = 5
    mvarHOficiales = vData
End Property


Public Property Get HOficiales() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HOficiales
    HOficiales = mvarHOficiales
End Property



Public Property Let Nombre(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Nombre = 5
    mvarNombre = vData
End Property


Public Property Get Nombre() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property



Public Property Let Codigo(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Codigo = 5
    mvarCodigo = vData
End Property


Public Property Get Codigo() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Codigo
    Codigo = mvarCodigo
End Property





Public Function DatosLineaDep() As String
Dim C As String
    C = "   Ofi (D/H)                  Traba(D/H /)          Compensadoas(D/H)  " & vbCrLf
    DatosLineaDep = C & String(70, "=") & vbCrLf
    C = Right(Space(40) & Me.DiasOficiales & "/" & Me.HOficiales, 15)
    DatosLineaDep = DatosLineaDep & C
    C = Right(Space(50) & Me.DiasReales_ & " / " & Me.HNReales & " / " & Me.HEReales, 30)
    DatosLineaDep = DatosLineaDep & C
    C = Right(Space(40) & Me.DiasCompensables & " / " & Me.HorasCompensadasNomina, 20)
    DatosLineaDep = DatosLineaDep & C

    
End Function
