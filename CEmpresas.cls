VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cempresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarIdEmpresa As Long 'copia local
Private mvarNomEmpresa As String 'copia local
Private mvarDirEmpresa As String 'copia local
Private mvarPobEmpresa As String 'copia local
Private mvarProvEmpresa As String 'copia local
Private mvarTelEmpresa As String 'copia local
Private mvarCodPosEmpresa As String 'copia local
Private mvarMaxRetraso As Currency 'copia local
Private mvarMaxExceso As Currency 'copia local
Private mvarIncHoraExtra As Integer 'copia local
Private mvarIncRetraso As Integer 'copia local
Private mvarIncMarcaje As Integer 'copia local
Private mvarIncVacaciones As Integer 'copia local
Private mvarIncTarjError As Integer 'copia local
Private mvarIncHoraExceso As Integer 'copia local
Private mvarCIF As String 'copia local
Private mvarMinutosRedondeo As Integer 'copia local
Private mvarAjusteEntrada As Integer 'copia local
Private mvarAjusteSalida As Integer 'copia local
Private mvarHorasJornada As Currency 'copia local
Private mvarRecuperacionDias As Integer 'copia local
Private mvarEntidad As String 'copia local
Private mvarSucursal As String 'copia local
Private mvarCodControl As String 'copia local
Private mvarCuenta As String 'copia local
Private mvarRepeticion As Integer 'copia local
Private mvarAplicaAntiguedadHN As Boolean 'copia local
Private mvarAplicaAntiguedadHC As Boolean 'copia local
Private mvarAbonosSeparados As Integer 'copia local
Private mvarIRPF As Currency 'copia local
Private mvarEmpresaHoraExtra As Boolean 'copia local
Private mvarNominaAutomatica As Boolean 'copia local

Private mvarredondeo As Byte 'copia local
Private mvarlaboral As Boolean 'copia local
Private mvarproduccion As Boolean 'copia local
Private mvarimgtrabaj As Boolean 'copia local
Private mvarReloj As Byte 'copia local
Private mvarTodosLosDias As Boolean  'copia local
Private mvarFechaInicio As Date
Private mvarFechaFin As Date
Private mvarServer As String
Private mvarDirMarcajes As String
Private mvarDirProcesados As String
Private mvarNomFich As String


'
'Solo tienen GET  '0. TEinsa     1.- Libre    2.- Alzira   4.- "Catadau "    5.- Coopic
                  '6.- BELGIDA
                
Private mvarQueEmpresa As Byte
Private mvarDirHuellas As String



'Febrero 2014
'Si crea una entrada por cada dia y trabajador o
'el programa trabaja leyendo en calendario
Private mvarCreaCalDiariaTra As Boolean  'copia local


'Enero 2015
Private mvarHorarioNocturno As Boolean  '0 NO
                                            
'nuevos
' Dias acabalgados.
' Ejemplo.  Empieza a las 22:30  del dia 17  y acaba a las 06:00 del segundo
' AcabalDiaTrabajado        0.- El dia es el que tiene el primer marcaje  Ejem: dia 17
'                           1.- El dia es el que tiene el primer marcaje  Ejem: dia 18
'
' AcabalHora                Hora que compararemos para ver si tiene un fichaje posterior
'                           para considerarlo acabalgado
'
' AcabalIncrementoxDia      Si el marcaje es acabalgado, le sumara a las HT una el valor que ponga esta variable
Private mvarAcabalgadoDiaInicio As Boolean
Private mvarAcabalgadoHora As Date
Private mvarAcabalgadoIncremento As Currency

'Se pueden ir a casa a las 12:00 (o mas ) del dia siguiente
Private mvarMaximaHoraDiaSiguiente As Date   'Es decir , si pongo las 4:00 significa que hasta esa hora, los fichajes son del dia de antes
Private mvarAcabaJornadaDiaSiguiente As Boolean


'Catadu
Private mvarpathCostesServer As String

'A3
Private mvarTieneCentrosA3 As Boolean



Private mvarCompensaHorasNominaMES As Boolean
Private mvarCompensaMES_HorasDia As Integer
Private mvarHoraSabadoExtras As String


' Para poder añadir otro reloj mas a  la lectura. De momento un BIostar2 (y solo COOPIC)
'       0: NO tiene     1: Tiene
Private mvarReloj2 As Byte 'copia local
Private mvarRelojConfig2 As String 'copia local


Private mvarNominaConEstructurales As Boolean  '











Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As ADODB.Recordset 'Para la base de datos
Private RC As Byte
'------------------------------------------------
'Propiedades del modulo Cempresas
'------------------------------------------------


Public Property Let IdEmpresa(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IdEmpresa = 5
     mvarIdEmpresa = vData
End Property

Public Property Get IdEmpresa() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IdEmpresa
     IdEmpresa = mvarIdEmpresa
End Property


Public Property Let NomEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NomEmpresa = 5
     mvarNomEmpresa = vData
End Property

Public Property Get NomEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NomEmpresa
     NomEmpresa = mvarNomEmpresa
End Property


Public Property Let DirEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DirEmpresa = 5
     mvarDirEmpresa = vData
End Property

Public Property Get DirEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DirEmpresa
     DirEmpresa = mvarDirEmpresa
End Property


Public Property Let PobEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.PobEmpresa = 5
     mvarPobEmpresa = vData
End Property

Public Property Get PobEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PobEmpresa
     PobEmpresa = mvarPobEmpresa
End Property


Public Property Let ProvEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ProvEmpresa = 5
     mvarProvEmpresa = vData
End Property

Public Property Get ProvEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ProvEmpresa
     ProvEmpresa = mvarProvEmpresa
End Property


Public Property Let TelEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.TelEmpresa = 5
     mvarTelEmpresa = vData
End Property

Public Property Get TelEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.TelEmpresa
     TelEmpresa = mvarTelEmpresa
End Property


Public Property Let CodPosEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.CodPosEmpresa = 5
     mvarCodPosEmpresa = vData
End Property

Public Property Get CodPosEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.CodPosEmpresa
     CodPosEmpresa = mvarCodPosEmpresa
End Property


Public Property Let MaxRetraso(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.MaxRetraso = 5
     mvarMaxRetraso = vData
End Property

Public Property Get MaxRetraso() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.MaxRetraso
     MaxRetraso = mvarMaxRetraso
End Property


Public Property Let MaxExceso(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.MaxExceso = 5
     mvarMaxExceso = vData
End Property

Public Property Get MaxExceso() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.MaxExceso
     MaxExceso = mvarMaxExceso
End Property


Public Property Let IncHoraExtra(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncHoraExtra = 5
     mvarIncHoraExtra = vData
End Property

Public Property Get IncHoraExtra() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncHoraExtra
     IncHoraExtra = mvarIncHoraExtra
End Property


Public Property Let IncRetraso(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncRetraso = 5
     mvarIncRetraso = vData
End Property

Public Property Get IncRetraso() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncRetraso
     IncRetraso = mvarIncRetraso
End Property


Public Property Let IncMarcaje(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncMarcaje = 5
     mvarIncMarcaje = vData
End Property

Public Property Get IncMarcaje() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncMarcaje
     IncMarcaje = mvarIncMarcaje
End Property


Public Property Let IncVacaciones(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncVacaciones = 5
     mvarIncVacaciones = vData
End Property

Public Property Get IncVacaciones() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncVacaciones
     IncVacaciones = mvarIncVacaciones
End Property


Public Property Let IncTarjError(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncTarjError = 5
     mvarIncTarjError = vData
End Property

Public Property Get IncTarjError() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncTarjError
     IncTarjError = mvarIncTarjError
End Property


Public Property Let IncHoraExceso(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncHoraExceso = 5
     mvarIncHoraExceso = vData
End Property

Public Property Get IncHoraExceso() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncHoraExceso
     IncHoraExceso = mvarIncHoraExceso
End Property


Public Property Let CIF(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.CIF = 5
     mvarCIF = vData
End Property

Public Property Get CIF() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.CIF
     CIF = mvarCIF
End Property


Public Property Let MinutosRedondeo(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.MinutosRedondeo = 5
     mvarMinutosRedondeo = vData
End Property

Public Property Get MinutosRedondeo() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.MinutosRedondeo
     MinutosRedondeo = mvarMinutosRedondeo
End Property


Public Property Let AjusteEntrada(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.AjusteEntrada = 5
     mvarAjusteEntrada = vData
End Property

Public Property Get AjusteEntrada() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.AjusteEntrada
     AjusteEntrada = mvarAjusteEntrada
End Property


Public Property Let AjusteSalida(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.AjusteSalida = 5
     mvarAjusteSalida = vData
End Property

Public Property Get AjusteSalida() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.AjusteSalida
     AjusteSalida = mvarAjusteSalida
End Property


Public Property Let HorasJornada(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HorasJornada = 5
     mvarHorasJornada = vData
End Property

Public Property Get HorasJornada() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HorasJornada
     HorasJornada = mvarHorasJornada
End Property


Public Property Let RecuperacionDias(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.RecuperacionDias = 5
     mvarRecuperacionDias = vData
End Property

Public Property Get RecuperacionDias() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.RecuperacionDias
     RecuperacionDias = mvarRecuperacionDias
End Property


Public Property Let Entidad(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Entidad = 5
     mvarEntidad = vData
End Property

Public Property Get Entidad() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Entidad
     Entidad = mvarEntidad
End Property


Public Property Let Sucursal(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Sucursal = 5
     mvarSucursal = vData
End Property

Public Property Get Sucursal() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Sucursal
     Sucursal = mvarSucursal
End Property


Public Property Let CodControl(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.CodControl = 5
     mvarCodControl = vData
End Property

Public Property Get CodControl() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.CodControl
     CodControl = mvarCodControl
End Property


Public Property Let Cuenta(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Cuenta = 5
     mvarCuenta = vData
End Property

Public Property Get Cuenta() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Cuenta
     Cuenta = mvarCuenta
End Property


Public Property Let Repeticion_(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Repeticion = 5
     mvarRepeticion = vData
End Property

Public Property Get Repeticion_() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Repeticion
     Repeticion_ = mvarRepeticion
End Property


Public Property Let AplicaAntiguedadHN(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.AplicaAntiguedadHN = 5
     mvarAplicaAntiguedadHN = vData
End Property

Public Property Get AplicaAntiguedadHN() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.AplicaAntiguedadHN
     AplicaAntiguedadHN = mvarAplicaAntiguedadHN
End Property


Public Property Let AplicaAntiguedadHC(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.AplicaAntiguedadHC = 5
     mvarAplicaAntiguedadHC = vData
End Property

Public Property Get AplicaAntiguedadHC() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.AplicaAntiguedadHC
     AplicaAntiguedadHC = mvarAplicaAntiguedadHC
End Property


Public Property Let AbonosSeparados(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.AbonosSeparados = 5
     mvarAbonosSeparados = vData
End Property

Public Property Get AbonosSeparados() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.AbonosSeparados
     AbonosSeparados = mvarAbonosSeparados
End Property


Public Property Let IRPF_(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IRPF = 5
     mvarIRPF = vData
End Property

Public Property Get IRPF_() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IRPF
     IRPF_ = mvarIRPF
End Property


Public Property Let EmpresaHoraExtra(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.EmpresaHoraExtra = 5
     mvarEmpresaHoraExtra = vData
End Property

Public Property Get EmpresaHoraExtra() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.EmpresaHoraExtra
     EmpresaHoraExtra = mvarEmpresaHoraExtra
End Property


Public Property Let NominaAutomatica(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NominaAutomatica = 5
     mvarNominaAutomatica = vData
End Property

Public Property Get NominaAutomatica() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NominaAutomatica
     NominaAutomatica = mvarNominaAutomatica
End Property


Public Property Let HorarioNocturno2(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HorarioNocturno = 5
     mvarHorarioNocturno = vData
End Property

Public Property Get HorarioNocturno2() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HorarioNocturno
     HorarioNocturno2 = mvarHorarioNocturno
End Property


Public Property Let redondeo(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.redondeo = 5
     mvarredondeo = vData
End Property

Public Property Get redondeo() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.redondeo
     redondeo = mvarredondeo
End Property


Public Property Let laboral(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.laboral = 5
     mvarlaboral = vData
End Property

Public Property Get laboral() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.laboral
     laboral = mvarlaboral
End Property


Public Property Let produccion(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.produccion = 5
     mvarproduccion = vData
End Property

Public Property Get produccion() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.produccion
     produccion = mvarproduccion
End Property


Public Property Let imgtrabaj(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.imgtrabaj = 5
     mvarimgtrabaj = vData
End Property

Public Property Get imgtrabaj() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.imgtrabaj
     imgtrabaj = mvarimgtrabaj
End Property


Public Property Let Reloj(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.reloj = 5
     mvarReloj = vData
End Property

'vbTCP3 = 1  vbKimaldi = 2      vbAlzira = 3     vbRobotics = 4
Public Property Get Reloj() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.reloj
     Reloj = mvarReloj
End Property




'Si genera datos para
Public Property Get TodosLosDias() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.reloj
     TodosLosDias = mvarTodosLosDias
End Property






Public Property Get FechaInicio() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PobEmpresa
     FechaInicio = mvarFechaInicio
End Property


Public Property Get FechaFin() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PobEmpresa
     FechaFin = mvarFechaFin
End Property


Public Property Get Server() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PobEmpresa
     Server = mvarServer
End Property






'-----------------------------------------
' Esto estaba antes en la clase config
' Ahora esta aqui dentro


Public Property Let NomFich(ByVal vData As String)

    mvarNomFich = vData
End Property


Public Property Get NomFich() As String
    NomFich = mvarNomFich
End Property



Public Property Let DirMarcajes(ByVal vData As String)
    mvarDirMarcajes = vData
End Property


Public Property Get DirMarcajes() As String

    DirMarcajes = mvarDirMarcajes
End Property


Public Property Let DirProcesados(ByVal vData As String)

    mvarDirProcesados = vData
End Property


Public Property Get DirProcesados() As String

    DirProcesados = mvarDirProcesados
End Property



'Solo tienen GET  '0. TEinsa     1.- Libre    2.- Alzira   4.- "Catadau "    5.- Coopic
                  '6.- BELGIDA
Public Property Get QueEmpresa() As Byte

     QueEmpresa = mvarQueEmpresa
End Property


'Solo tienen GET
Public Property Get DirHuellas() As String
     DirHuellas = mvarDirHuellas
End Property



'el programa trabaja leyendo en calendario  o en la tabla calendariotrabajador
Public Property Get CreaCalDiariaTra() As Boolean
     CreaCalDiariaTra = mvarCreaCalDiariaTra
End Property




'**********************************************************************************************
' Dias acabalgados.
' Ejemplo.  Empieza a las 22:30  del dia 17  y acaba a las 06:00 del segundo
' AcabalDiaTrabajado        0.- El dia es el que tiene el primer marcaje  Ejem: dia 17
'                           1.- El dia es el que tiene el primer marcaje  Ejem: dia 18
'
' AcabalHora                Hora que compararemos para ver si tiene un fichaje posterior
'                           para considerarlo acabalgado
'
' AcabalIncrementoxDia      Si el marcaje es acabalgado, le sumara a las HT una el valor que ponga esta variable
Public Property Get AcabalgadoDiaInicio() As Boolean
     AcabalgadoDiaInicio = mvarAcabalgadoDiaInicio
End Property
Public Property Get AcabalgadoHora() As Date
     AcabalgadoHora = mvarAcabalgadoHora
End Property
Public Property Get AcabalgadoIncremento() As Currency
     AcabalgadoIncremento = mvarAcabalgadoIncremento
End Property


Public Property Get pathCostesServer() As String
     pathCostesServer = mvarpathCostesServer
End Property

Public Property Get TieneCentrosA3() As Boolean
     TieneCentrosA3 = mvarTieneCentrosA3
End Property








'0 NO     1 Tiene un biostar2
Public Property Get Reloj2() As Integer
     Reloj2 = mvarReloj2
End Property

Public Property Get RelojConfig2() As String
     RelojConfig2 = mvarRelojConfig2
End Property


Public Property Get MaximaHoraDiaSiguiente() As String
     MaximaHoraDiaSiguiente = mvarMaximaHoraDiaSiguiente
End Property
Public Property Get AcabaJornadaDiaSiguiente() As Boolean
     AcabaJornadaDiaSiguiente = mvarAcabaJornadaDiaSiguiente
End Property




'La copmpensacion es al final de mes. Se suman HN, HEstr,  y frente a las que tiene que trabajar , si necesita coje, si no
Public Property Get CompensaHorasNominaMES() As Boolean
     CompensaHorasNominaMES = mvarCompensaHorasNominaMES
End Property


Public Property Get CompensaMES_HorasDia() As Integer
     CompensaMES_HorasDia = mvarCompensaMES_HorasDia
End Property

Public Property Get HoraSabadoExtras() As String
     HoraSabadoExtras = mvarHoraSabadoExtras
End Property

Public Property Get NominaConEstructurales() As Boolean
     NominaConEstructurales = mvarNominaConEstructurales
End Property






'****   LEER  ***
Public Function Leer(vIdIdEmpresa As Long) As Byte
Dim I As Integer

    On Error GoTo Err3

    SQL = "SELECT * FROM empresas"
    SQL = SQL & " WHERE "
    SQL = SQL & " IdEmpresa = " & vIdIdEmpresa
    RS.Open SQL, conn, , , adCmdText
    If RS.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarIdEmpresa = RS!IdEmpresa
            mvarNomEmpresa = DBLet(RS!NomEmpresa)
            mvarDirEmpresa = DBLet(RS!DirEmpresa)
            mvarPobEmpresa = DBLet(RS!PobEmpresa)
            mvarProvEmpresa = DBLet(RS!ProvEmpresa)
            mvarTelEmpresa = DBLet(RS!TelEmpresa)
            mvarCodPosEmpresa = DBLet(RS!CodPosEmpresa)
            mvarMaxRetraso = RS!MaxRetraso
            mvarMaxExceso = RS!MaxExceso
            mvarIncHoraExtra = RS!IncHoraExtra
            mvarIncRetraso = RS!IncRetraso
            mvarIncMarcaje = RS!IncMarcaje
            mvarIncVacaciones = RS!IncVacaciones
            mvarIncTarjError = RS!IncTarjError
            mvarIncHoraExceso = RS!IncHoraExceso
            mvarCIF = DBLet(RS!CIF)
            mvarMinutosRedondeo = RS!MinutosRedondeo
            mvarAjusteEntrada = RS!AjusteEntrada
            mvarAjusteSalida = RS!AjusteSalida
            mvarHorasJornada = DBLet(RS!HorasJornada, "N")
            mvarRecuperacionDias = RS!RecuperacionDias
            mvarEntidad = DBLet(RS!Entidad)
            mvarSucursal = DBLet(RS!Sucursal)
            mvarCodControl = DBLet(RS!CodControl)
            mvarCuenta = DBLet(RS!Cuenta)
            mvarRepeticion = RS!Repeticion
            mvarAplicaAntiguedadHN = RS!AplicaAntiguedadHN = 1
            mvarAplicaAntiguedadHC = RS!AplicaAntiguedadHC = 1
            mvarAbonosSeparados = RS!AbonosSeparados = 1
            mvarIRPF = DBLet(RS!IRPF, "N")
            mvarEmpresaHoraExtra = RS!EmpresaHoraExtra = 1
            mvarNominaAutomatica = RS!NominaAutomatica = 1
            mvarHorarioNocturno = RS!HorarioNocturno = 1
            mvarredondeo = RS!redondeo
            mvarlaboral = RS!laboral = 1
            mvarproduccion = RS!produccion = 1
            mvarimgtrabaj = RS!imgtrabaj = 1
            mvarReloj = RS!Reloj
            mvarTodosLosDias = (RS!TodosLosDias = 1)
            'mvarFechaInicio="01/04/2006"  'Para probar años partidos
            mvarFechaInicio = Format(RS!FechaInicio, "dd/mm/yyyy")
            mvarFechaFin = DateAdd("yyyy", 1, mvarFechaInicio)
            mvarFechaFin = DateAdd("d", -1, mvarFechaFin)
            
            '---------------------
            'Antiguamente estaba en el config.cfg
            mvarDirMarcajes = DBLet(RS!DirMarcajes)
            mvarDirProcesados = DBLet(RS!Pathproces)
            mvarNomFich = DBLet(RS!Nomproces)
            
            
            'Cotubre 2013
            mvarQueEmpresa = DBLet(RS!QueEmpresa, "N")
    
            
            
            mvarCreaCalDiariaTra = DBLet(RS!CreaCalDiariaTra, "N") = 1
            
            
            mvarDirHuellas = DBLet(RS!DirHuellas, "T")
            
            
            
            'Enero 2015
            If mvarHorarioNocturno Then
                mvarAcabalgadoDiaInicio = DBLet(RS!AcabalDiaTrabajado, "N") = 0
                mvarAcabalgadoHora = "22:30:00"
                If Not IsNull(RS!AcabalHora) Then mvarAcabalgadoHora = Format(RS!AcabalHora, "hh:mm")
                mvarAcabalgadoIncremento = DBLet(RS!AcabalIncrementoxDia, "N")
            
                
                
                
            End If
            'USuarios
            mvarServer = DBLet(RS!Servidor)
            
            
            mvarpathCostesServer = DBLet(RS!pathCostesServer, "T")
            
            '-------------------------------------
            'Vere si son el mismo
            I = InStr(1, conn.ConnectionString, ";SERVER=")
            If I = 0 Then
            
                'Voy a suponer que es el mismo servidor que el de usuarios
                mvarServer = ""
            
            Else
                
                
                SQL = Mid(conn.ConnectionString, I + 8)
                I = InStr(1, SQL, ";")
                If I = 0 Then
                    MsgBox "Error grave. Cadena no encontrada. 2º;" & vbCrLf & SQL, vbCritical
                    End
                End If
                SQL = UCase(Mid(SQL, 1, I - 1))
                If SQL = UCase(mvarServer) Then
                    'PERFECTO. SOn el mismo servidor
                    mvarServer = ""
                Else
                    mvarServer = DBLet(RS!Servidor) & "|" & DBLet(RS!Usuario) & "|" & DBLet(RS!Pass)
                End If
            End If
            
            mvarReloj2 = DBLet(RS!Reloj2, "N")
            mvarRelojConfig2 = DBLet(RS!configreloj2, "T")
            
            'Noviembre2018
            mvarAcabaJornadaDiaSiguiente = False
            If Not IsNull(RS!HoraDiaSiguiente) Then
                mvarMaximaHoraDiaSiguiente = Format(RS!HoraDiaSiguiente, "hh:nn:ss")
                mvarAcabaJornadaDiaSiguiente = True
            End If
            
            
            
            'Sept 19
            mvarCompensaHorasNominaMES = DBLet(RS!CompensaHorasNominaMES, "N") = 1
            mvarCompensaMES_HorasDia = DBLet(RS!CompensaMES_HorasDia, "N")
              
            mvarHoraSabadoExtras = DBLet(RS!HoraSabadoExtras, "T")
            'OCt 1
            
            mvarNominaConEstructurales = False
            mvarNominaConEstructurales = DBLet(RS!NominaConEstructurales, "N") = 1
            
            
            Leer = 0
    End If
 
    RS.Close
    
    If Leer = 0 Then
        FijaCentrosA3
    End If
    
    
    Exit Function

Err3:
    Mens = "Se ha producido un error parametros." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Leer = 1
End Function



Private Sub FijaCentrosA3()
    On Error GoTo eF
    mvarTieneCentrosA3 = False
    
    SQL = "SELECT * FROM CentrosA3"
    RS.Open SQL, conn, , , adCmdText
    If Not RS.EOF Then
        If RS.Fields(0).Value >= 0 Then mvarTieneCentrosA3 = True
    End If
    RS.Close
eF:
    Err.Clear
    conn.Errors.Clear
End Sub


Public Sub ActualizarDatosReloj(Datos As String)
    On Error GoTo eAC
    If vEmpresa.Reloj = vbBiostarXpass Then
    
        
    End If
    
eAC:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description

End Sub






'****   MODIFICAR  ***
'Public Function Modificar() As Byte
'    On Error GoTo Err4
'
'    Sql = "SELECT * FROM empresas"
'    Sql = Sql & " WHERE "
'    Sql = Sql & " IdEmpresa = " & mvarIdEmpresa
'    Rs.Open Sql, Conn, , , adCmdText
'    If Rs.EOF Then
'        Mens = "No existe el Registro en la Tabla empresas con esos valores."
'        MsgBox Mens, vbExclamation
'        Modificar = 1
'        'ELSE
'        Else
'            'Rs.Edit
'            Rs!NomEmpresa = mvarNomEmpresa
'            Rs!DirEmpresa = mvarDirEmpresa
'            Rs!PobEmpresa = mvarPobEmpresa
'            Rs!ProvEmpresa = mvarProvEmpresa
'            Rs!TelEmpresa = mvarTelEmpresa
'            Rs!CodPosEmpresa = mvarCodPosEmpresa
'            Rs!MaxRetraso = mvarMaxRetraso
'            Rs!MaxExceso = mvarMaxExceso
'            Rs!IncHoraExtra = mvarIncHoraExtra
'            Rs!IncRetraso = mvarIncRetraso
'            Rs!IncMarcaje = mvarIncMarcaje
'            Rs!IncVacaciones = mvarIncVacaciones
'            Rs!IncTarjError = mvarIncTarjError
'            Rs!IncHoraExceso = mvarIncHoraExceso
'            Rs!CIF = mvarCIF
'            Rs!MinutosRedondeo = mvarMinutosRedondeo
'            Rs!AjusteEntrada = mvarAjusteEntrada
'            Rs!AjusteSalida = mvarAjusteSalida
'            Rs!HorasJornada = mvarHorasJornada
'            Rs!RecuperacionDias = mvarRecuperacionDias
'            Rs!Entidad = mvarEntidad
'            Rs!Sucursal = mvarSucursal
'            Rs!CodControl = mvarCodControl
'            Rs!Cuenta = mvarCuenta
'            Rs!Repeticion = mvarRepeticion
'            Rs!AplicaAntiguedadHN = mvarAplicaAntiguedadHN
'            Rs!AplicaAntiguedadHC = mvarAplicaAntiguedadHC
'            Rs!AbonosSeparados = mvarAbonosSeparados
'            Rs!IRPF = mvarIRPF
'            Rs!EmpresaHoraExtra = mvarEmpresaHoraExtra
'            Rs!NominaAutomatica = mvarNominaAutomatica
'            Rs!HorarioNocturno = mvarHorarioNocturno
'            Rs!redondeo = mvarredondeo
'            Rs!laboral = mvarlaboral
'            Rs!produccion = mvarproduccion
'            Rs!imgtrabaj = mvarimgtrabaj
'            Rs!reloj = mvarreloj
'            Rs.Update
'            Modificar = 0
'    End If
'
'    Rs.Close
'    Exit Function
'
'Err4:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Modificar = 1
'End Function



Private Sub Class_Initialize()
    Set RS = New ADODB.Recordset
    'Ponemos el recordset en modo:
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
End Sub


Private Sub Class_Terminate()
    Set RS = Nothing
End Sub



