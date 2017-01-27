VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarIdHorario As Integer 'copia local
Private mvarNomHorario As String 'copia local
Private mvarDiaSemana As Byte 'copia local
Private mvarHoraE1 As Date 'copia local
Private mvarHoraS1 As Date 'copia local
Private mvarHoraE2 As Date 'copia local
Private mvarHoraS2 As Date 'copia local
Private mvarNumTikadas As Byte 'copia local
Private mvarTotalHoras As Single 'copia local
'Private mvarTotalHoras As byte 'copia local
Private mvarEsDiaFestivo As Boolean 'copia local
Private mvarDtoAlm As Single 'copia local
Private mvarHoraDtoAlm As Date 'copia local
Private mvarDtoMer As Single 'copia local
Private mvarHoraDtoMer As Date 'copia local
Private mvarDiaNomina As Currency 'copia local
Private mvarFecha As Date
Private mvarRectificar As Byte
Private mvarIdCal As Integer


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As Recordset
Private RC As Byte



'------------------------------------------------
'Propiedades del modulo CHorarios
'------------------------------------------------


Public Property Let IdHorario(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IdHorario = 5
     mvarIdHorario = vData
End Property

Public Property Get IdHorario() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IdHorario
     IdHorario = mvarIdHorario
End Property


Public Property Let IdCal(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IdHorario = 5
     mvarIdCal = vData
End Property

Public Property Get IdCal() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IdHorario
     IdCal = mvarIdCal
End Property




Public Property Let NomHorario(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NomHorario = 5
     mvarNomHorario = vData
End Property

Public Property Get NomHorario() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NomHorario
     NomHorario = mvarNomHorario
End Property


Public Property Let DiaSemana(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DiaSemana = 5
     mvarDiaSemana = vData
End Property

Public Property Get DiaSemana() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DiaSemana
     DiaSemana = mvarDiaSemana
End Property


Public Property Let HoraE1(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraE1 = 5
     mvarHoraE1 = vData
End Property

Public Property Get HoraE1() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraE1
     HoraE1 = mvarHoraE1
End Property


Public Property Let HoraS1(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraS1 = 5
     mvarHoraS1 = vData
End Property

Public Property Get HoraS1() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraS1
     HoraS1 = mvarHoraS1
End Property


Public Property Let HoraE2(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraE2 = 5
     mvarHoraE2 = vData
End Property

Public Property Get HoraE2() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraE2
     HoraE2 = mvarHoraE2
End Property


Public Property Let HoraS2(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraS2 = 5
     mvarHoraS2 = vData
End Property

Public Property Get HoraS2() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraS2
     HoraS2 = mvarHoraS2
End Property


Public Property Let NumTikadas(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NumTikadas = 5
     mvarNumTikadas = vData
End Property

Public Property Get NumTikadas() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NumTikadas
     NumTikadas = mvarNumTikadas
End Property

Public Property Let TotalHoras(ByVal vData As Single)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NumTikadas = 5
     mvarTotalHoras = vData
End Property

Public Property Get TotalHoras() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NumTikadas
     TotalHoras = mvarTotalHoras
End Property



Public Property Let EsDiaFestivo(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NumTikadas = 5
     mvarEsDiaFestivo = vData
End Property

Public Property Get EsDiaFestivo() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NumTikadas
     EsDiaFestivo = mvarEsDiaFestivo
End Property



Public Property Let DtoAlm(ByVal vData As Single)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DtoAlm = 5
     mvarDtoAlm = vData
End Property

Public Property Get DtoAlm() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DtoAlm
     DtoAlm = mvarDtoAlm
End Property


Public Property Let HoraDtoAlm(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraDtoAlm = 5
     mvarHoraDtoAlm = vData
End Property

Public Property Get HoraDtoAlm() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraDtoAlm
     HoraDtoAlm = mvarHoraDtoAlm
End Property


Public Property Get Fecha() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraDtoAlm
     Fecha = mvarFecha
End Property




Public Property Let DtoMer(ByVal vData As Single)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DtoMer = 5
     mvarDtoMer = vData
End Property

Public Property Get DtoMer() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DtoMer
     DtoMer = mvarDtoMer
End Property



Public Property Let Rectificar(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DtoMer = 5
     mvarRectificar = vData
End Property

Public Property Get Rectificar() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DtoMer
     Rectificar = mvarRectificar
End Property




Public Property Let HoraDtoMer(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraDtoMer = 5
     mvarHoraDtoMer = vData
End Property

Public Property Get HoraDtoMer() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraDtoMer
     HoraDtoMer = mvarHoraDtoMer
End Property



Public Property Let DiaNomina(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraDtoMer = 5
     mvarDiaNomina = vData
End Property

Public Property Get DiaNomina() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraDtoMer
     DiaNomina = mvarDiaNomina
End Property





'------------------------------------------------
'FUNCIONES del modulo CHorarios
'------------------------------------------------


'****   LEER  ***
Public Function Leer(vIdIdHorario As Integer, vFecha, Calendario As Integer) As Byte
Dim RT As ADODB.Recordset
Dim DiaSem As Byte
Dim HayHorarios As Boolean
Dim EsFinD As Boolean 'Festivo de Fin de semana

    On Error GoTo Err3
    SQL = "SELECT * FROM Horarios"
    SQL = SQL & " WHERE "
    SQL = SQL & " idHorario = " & vIdIdHorario
    RS.Open SQL, conn, , , adCmdText
    mvarIdCal = Calendario
    If RS.EOF Then
        Leer = 1
       
        'ELSE
        Else
            'Variables de HORARIOS
            
                mvarIdHorario = RS!IdHorario
                
                mvarFecha = CDate(vFecha)
                
                mvarNomHorario = RS!NomHorario
                mvarDtoAlm = DBLet(RS!DtoAlm, "N")
                mvarHoraDtoAlm = Format(DBLet(RS!HoraDtoAlm, "F"), "hh:mm:ss")
                mvarDtoMer = DBLet(RS!DtoMer, "N")
                mvarHoraDtoMer = Format(DBLet(RS!HoraDtoMer, "F"), "hh:mm:ss")
                mvarRectificar = RS!Rectificar
                
                DiaSem = Weekday(vFecha, vbMonday)
                mvarDiaSemana = DiaSem
                SQL = "select * from Subhorarios where idHorario=" & vIdIdHorario
                SQL = SQL & " AND DiaSemana=" & DiaSem
                Set RT = New ADODB.Recordset
                RT.Open SQL, conn, , , adCmdText
                If RT.EOF Then
                    HayHorarios = False
                    Else
                        HayHorarios = True
                        If DBLet(RT!Festivo, "N") = 0 Then
                            mvarHoraE1 = Format(DBLet(RT!HEntrada1, "H"), "hh:mm")
                            mvarHoraS1 = Format(DBLet(RT!HSalida1, "H"), "hh:mm")
                            If IsNull(RT!hentrada2) Then
                                'mvarHoraE2 = ""
                            Else
                                mvarHoraE2 = Format(DBLet(RT!hentrada2, "H"), "hh:mm")
                            End If
                            If IsNull(RT!HSalida2) Then
                                'mvarHoraS2 = ""
                            Else
                                mvarHoraS2 = Format(DBLet(RT!HSalida2, "H"), "hh:mm")
                            End If
                        End If
                        mvarNumTikadas = RT!N_Tikadas
                        mvarTotalHoras = RT!HorasDia
                        If Not IsNull(RT!DiaNomina) Then
                            mvarDiaNomina = RT!DiaNomina
                        Else
                            mvarDiaNomina = 0
                        End If
                        EsFinD = RT!Festivo = 1
                End If
                RT.Close
                Set RT = Nothing
                
            
            'Comprobamos si es un dia festivo
            SQL = EsFestivo(vFecha)
            If SQL <> "" Then
                If Not HayHorarios Then
                    mvarHoraE1 = "0:00:00"
                    mvarHoraS1 = "0:00:00"
                    mvarHoraE2 = "0:00:00"
                    mvarHoraS2 = "0:00:00"
                End If
                mvarEsDiaFestivo = True
                mvarNumTikadas = 0
                mvarTotalHoras = 0
                mvarNomHorario = mvarNomHorario & " (" & SQL & ")"
                Leer = 0
                Else
                    'No es festivo de calendario
                    'Comprobaremos que no es festivo de
    
                    If HayHorarios Then
                        
                        If EsFinD Then
                            mvarEsDiaFestivo = True
                            mvarNomHorario = mvarNomHorario & " (Festivo)"
                        Else
                             mvarEsDiaFestivo = False
                        End If
                        Leer = 0
                        Else
                            Leer = 1
                    End If
            End If 'DE FESTIVO
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

Private Sub Class_Initialize()
Set RS = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
Set RS = Nothing
End Sub

Public Sub CambioCalendario(NuevoCal As Integer)
    mvarIdCal = NuevoCal
    SQL = EsFestivo(mvarFecha)
    mvarEsDiaFestivo = SQL <> ""
    
End Sub


Private Function EsFestivo(Fecha) As String
Dim Cad As String
Dim RF As ADODB.Recordset

'Devuelve una cadena que dice que fiesta es( y se añade el nombre de horario.")
EsFestivo = ""
Set RF = New ADODB.Recordset
Cad = "select * from calendariof  where IdCal=" & mvarIdCal
Cad = Cad & " and Fecha='" & Format(Fecha, FormatoFecha) & "'"
RF.Open Cad, conn, , , adCmdText
If Not RF.EOF Then
    EsFestivo = RF.Fields(0)
End If
RF.Close
Set RF = Nothing

End Function



'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
Public Function LeerDiasFestivos(vIdIdHorario As Integer, Fini As Date, FFin As Date) As String
Dim RT As ADODB.Recordset
Dim HayHorarios As Boolean
Dim EsFinD As Boolean 'Festivo de Fin de semana
Dim FE As Date
Dim D As Integer
Dim FESTIVOS As String
Dim RecuperaEnSabado As Boolean

    On Error GoTo Err3
    SQL = "SELECT * FROM Horarios"
    SQL = SQL & " WHERE "
    SQL = SQL & " idHorario = " & vIdIdHorario
    RS.Open SQL, conn, , , adCmdText
    LeerDiasFestivos = ""
    FESTIVOS = ""
    If Not RS.EOF Then
        
            'Variables de HORARIOS
            mvarIdHorario = RS!IdHorario
            mvarNomHorario = RS!NomHorario
            
            'Vemos si los sabados son recuperables . COOPIC seguro que no lo hace
            RecuperaEnSabado = False
            If vEmpresa.QueEmpresa <> 5 Then RecuperaEnSabado = RecuperaDiasEnSabado(RS)
            
            Set RT = New ADODB.Recordset
                        
            'FALTA###
            SQL = "Select * from calendariof WHERE idcal = 1"
            SQL = SQL & " AND Fecha >= '" & Format(Fini, FormatoFecha) & "'"
            SQL = SQL & " AND Fecha <= '" & Format(FFin, FormatoFecha) & "'"
            SQL = SQL & " ORDER By FECHA"
            RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                FESTIVOS = FESTIVOS & Format(RT!Fecha, "dd/mm/yyyy") & "|"
                RT.MoveNext
            Wend
            RT.Close
            
            'COOPIC no hace nada. SOLO deberia ser, cuando venga, PICASSENT
            If vEmpresa.QueEmpresa <> 5 Then
                SQL = "select * from Subhorarios where idHorario=" & vIdIdHorario
                RT.Open SQL, conn, , , adCmdText
                If RT.EOF Then
                    
                    Else
                        'En el SQL Meteremos los dias de las semana k son festivos
                        'Es decir , empimaremos los dias k sean festivos
                        ' Doming|Sabado|  por ejemplo, pero en numeros 5|6|
                        SQL = ""
                        While Not RT.EOF
                            
                            If RT!Festivo Then
                               
                                If RecuperaEnSabado Then
                                    If RT!DiaSemana <> 6 Then
                                        SQL = SQL & RT!DiaSemana & "|"
                                    End If
                                Else
                                    SQL = SQL & RT!DiaSemana & "|"
                                End If
                            End If
                            RT.MoveNext
                        Wend
                        
                        
                        FE = Fini
                        While FE <= FFin
                            D = Weekday(FE, vbMonday)
                            Mens = D & "|"
                            If InStr(1, SQL, Mens) > 0 Then
                                    FESTIVOS = FESTIVOS & Format(FE, "dd/mm/yyyy") & "|"
                            End If
                                
                            'Dia mas
                            FE = DateAdd("d", 1, FE)
                        Wend
                End If
                RT.Close
            End If
            Set RT = Nothing
            'Comprobamos si es un dia festivo
            LeerDiasFestivos = FESTIVOS
    
    End If
    RS.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    LeerDiasFestivos = ""
End Function





'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
Public Function LeerMediosDias(vIdIdHorario As Integer, Fini As Date, FFin As Date) As String
Dim RT As ADODB.Recordset
Dim HayHorarios As Boolean
Dim EsFinD As Boolean 'Festivo de Fin de semana
Dim FE As Date
Dim D As Integer
Dim MediosDias As String

    On Error GoTo Err3
    SQL = "SELECT * FROM Horarios"
    SQL = SQL & " WHERE "
    SQL = SQL & " idHorario = " & vIdIdHorario
    RS.Open SQL, conn, , , adCmdText
    LeerMediosDias = ""
    MediosDias = ""
    If Not RS.EOF Then
        
            'Variables de HORARIOS
            mvarIdHorario = RS!IdHorario
            mvarNomHorario = RS!NomHorario
            Set RT = New ADODB.Recordset
            
            SQL = "select * from Subhorarios where idHorario=" & vIdIdHorario
            RT.Open SQL, conn, , , adCmdText
            If RT.EOF Then
                
                Else
                    'En el SQL Meteremos los dias de las semana k s trabajan 0.5 horas
                    ' Miercoles|Sabado|  por ejemplo, pero en numeros 3|6|
                    SQL = ""
                    While Not RT.EOF
                        If Not RT!Festivo Then
                            If RT!DiaNomina < 1 Then SQL = SQL & RT!DiaSemana & "|"
                        End If
                        RT.MoveNext
                    Wend
                    
                    
                    FE = Fini
                    While FE <= FFin
                        D = Weekday(FE, vbMonday)
                        Mens = D & "|"
                        If InStr(1, SQL, Mens) > 0 Then
                                MediosDias = MediosDias & Format(FE, "dd/mm/yyyy") & "|"
                        End If
                            
                        'Dia mas
                        FE = DateAdd("d", 1, FE)
                    Wend
            End If
            RT.Close
            Set RT = Nothing
            'Comprobamos si es un dia festivo
            LeerMediosDias = MediosDias
    
    End If
    RS.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    LeerMediosDias = ""
End Function



Private Function RecuperaDiasEnSabado(ByRef R As Recordset) As Boolean
    On Error Resume Next
    RecuperaDiasEnSabado = False
    RecuperaDiasEnSabado = RS!RecuperaSabados
    If Err.Number <> 0 Then Err.Clear
End Function
