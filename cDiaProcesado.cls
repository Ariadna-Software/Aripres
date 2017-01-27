VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDiaProcesado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Private mvarDiaSemana As Byte 'copia local
Private mvarFestivo As Boolean 'copia local
Private mvarHT_Reales As Currency 'copia local
Private mvarHE_Reales As Currency 'copia local
Private mvarDiaNomina As Currency 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarDiaProcesable As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarBaja As Boolean 'copia local

'Los sabados puede haberlos trabajado, pero DIASnimmona se pone a cero para que  no le sume en el computo total
Private mvarSabadoSiHabiaTrabajado As Boolean


Private mvarNumeroSemana As Integer

Public Property Let Baja(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Baja = 5
    mvarBaja = vData
End Property


Public Property Get Baja() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Baja
    Baja = mvarBaja
End Property



Public Property Let DiaProcesable(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiaProcesable = 5
    mvarDiaProcesable = vData
End Property


Public Property Get DiaProcesable() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiaProcesable
    DiaProcesable = mvarDiaProcesable
End Property


Public Property Let SabadoSiHabiaTrabajado(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiaProcesable = 5
    mvarSabadoSiHabiaTrabajado = vData
End Property


Public Property Get SabadoSiHabiaTrabajado() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiaProcesable
    SabadoSiHabiaTrabajado = mvarSabadoSiHabiaTrabajado
End Property




Public Property Let DiaNomina(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiaNomina = 5
    mvarDiaNomina = vData
End Property


Public Property Get DiaNomina() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiaNomina
    DiaNomina = mvarDiaNomina
End Property



Public Property Let HE_Reales(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HE_Reales = 5
    mvarHE_Reales = vData
End Property


Public Property Get HE_Reales() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HE_Reales
    HE_Reales = mvarHE_Reales
End Property



Public Property Let HT_Reales(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HT_Reales = 5
    mvarHT_Reales = vData
End Property


Public Property Get HT_Reales() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HT_Reales
    HT_Reales = mvarHT_Reales
End Property



Public Property Let Festivo(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Festivo = 5
    mvarFestivo = vData
End Property


Public Property Get Festivo() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Festivo
    Festivo = mvarFestivo
End Property



Public Property Let DiaSemana(ByVal vData As Byte)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiaSemana = 5
    mvarDiaSemana = vData
End Property


Public Property Get DiaSemana() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiaSemana
    DiaSemana = mvarDiaSemana
End Property

Public Property Let NumeroSemana(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiaSemana = 5
    mvarNumeroSemana = vData
End Property


Public Property Get NumeroSemana() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiaSemana
    NumeroSemana = mvarNumeroSemana
End Property



