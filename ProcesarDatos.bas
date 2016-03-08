Attribute VB_Name = "ProcesarDatos"
Option Explicit
 

Public TIPOALZICOOP As Boolean
'Public CadenaProv As String


Public vH As CHorarios
Public vE As Cempresas

Private RedondeoMarcajes As Byte
Private Signo As Integer
Private QuitarAlmuerzo As Boolean
Private quitarmerienda As Boolean

Public Sub GeneraIncidencia(Inci As Integer, marca As Long, Horas As Single)
Dim RS As ADODB.Recordset
Dim Id As Long

    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open "Select * from IncidenciasGeneradas", Conn, , , adCmdText
    If RS.EOF Then
        Id = 1
        Else
            RS.MoveLast
            Id = RS!Id + 1
    End If
    RS.AddNew
    RS!Id = Id
    RS!Incidencia = Inci
    RS!EntradaMarcaje = marca
    RS!Horas = Horas
    RS.Update
    RS.Close
    Set RS = Nothing
End Sub

Public Sub ProcesarMarcaje_Tipo1(ByRef vMar As CMarcajes)
Dim Rss As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim NumTikadas As Integer
Dim T1 As Single
Dim T2 As Single
Dim kIncidencia As Single
Dim TieneIncidencia As Boolean
Dim MarcajeCorrecto As Boolean
Dim Exceso As Date
Dim Retraso As Date
Dim I As Long
Dim v(3) As Single
Dim Cad As String
Dim HoraH As Date
Dim InciManual As Integer
Dim N As Integer
Dim TotalH As Single


'Ahora ya tenemos las horas tikadas reflejadas
'Comprobamos las horas en funcion de los horarios
'  y calculamos las horas comprobadas

Conn.BeginTrans
Set Rss = New ADODB.Recordset
'Vector para incidencias
For I = 0 To 3
    v(I) = 0
Next I
'Seleccionamos todas las horas de este
Cad = "Select * from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
Cad = Cad & " AND Fecha=#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"
Cad = Cad & " ORDER BY Hora"
Rss.CursorType = adOpenStatic
Rss.Open Cad, Conn, , , adCmdText

If Rss.EOF Then
    'Si no hay ninguna entrada
    Rss.Close
    GoTo ErrorProcesaMarcaje
End If
InciManual = 0

'If vMar.idTrabajador = 307 Then Stop

NumTikadas = Rss.RecordCount
If vH.EsDiaFestivo Then
    'Si es festivo asignamos las tikadas segun vengan
    ' y todo pasa a ser horas extras
    If (NumTikadas Mod 2) > 0 Then
        'Numero de marcajes impares. No podemos calcular horas
        'trabajadas. Generamos error en marcaje
        vMar.IncFinal = vE.IncMarcaje
        vMar.HorasIncid = 0
        vMar.HorasTrabajadas = 0
        GeneraIncidencia vE.IncMarcaje, vMar.Entrada, 0
        Else
            N = NumTikadas \ 2
            TotalH = 0
            'NUMERO DE MARCAJES PAR
            Rss.MoveFirst
            For I = 1 To N
                T1 = DevuelveValorHora(Rss!Hora)
                Rss.MoveNext
                T2 = DevuelveValorHora(Rss!Hora)
                Rss.MoveNext
                TotalH = TotalH + (T2 - T1)
            Next I
            
            'Contabilizaremos los descuentos relativos al almuerzo y merienda
            'si procede
                
            QuitarAlmuerzo = False
            quitarmerienda = False
            
'            If vH.DtoAlm > 0 Then
'                Rss.MoveFirst
'                For I = 1 To N
'                    PrimerTicaje = Rss!Hora
'                    Rss.MoveNext
'                    If PrimerTicaje < vH.HoraDtoAlm Then
'                        If Rss!Hora > vH.HoraDtoAlm Then QuitarAlmuerzo = True
'                    End If
'                Next I
'            End If
                    
                    
                    
            'Nuevo. Revision pedida por Catadau. Si el trabajador NO esta , no puede quitarsele el almuerzo
            If vH.DtoAlm > 0 Then
                    QuitarAlmuerzo = LeQuitamosElAmluerzo(Rss, vH)
            End If
                    
            If vH.DtoMer > 0 Then
                For I = 1 To N
                    PrimerTicaje = Rss!Hora
                    Rss.MoveNext
                    If PrimerTicaje <= vH.HoraDtoMer Then
                        If Rss!Hora > vH.HoraDtoMer Then quitarmerienda = True
                    End If
                Next I
            End If
        
            'Ahora ya sabemos las horas trabajadas
            TotalH = Round(TotalH, 2)
            
            'Asignamos a la incidencia
            T2 = TotalH
            If QuitarAlmuerzo Then T2 = T2 - vH.DtoAlm
            If quitarmerienda Then T2 = T2 - vH.DtoMer
            
            TotalH = Round(T2, 2)
             vMar.HorasTrabajadas = TotalH
             vMar.HorasIncid = TotalH
             vMar.IncFinal = vE.IncHoraExtra
    End If  'de NUMTIKADAS es numero par

'------------------
'------------------
'
'ELSE de DIA FESTIVO
'
Else
    If NumTikadas = vH.NumTikadas Then
        'Ha ticado las mismas veces que le correspondian
        'Comprobamos si ha habido algun retraso, o exceso
        Exceso = DevuelveHora(vE.MaxExceso)
        Retraso = DevuelveHora(vE.MaxRetraso)
        I = 0
        PrimerTicaje = Rss!Hora
        While Not Rss.EOF
            If Rss!IdInci > 0 Then InciManual = Rss!IdInci
            Select Case I
            Case 0
                HoraH = vH.HoraE1
            Case 1
                HoraH = vH.HoraS1
            Case 2
                HoraH = vH.HoraE2
            Case 3
                HoraH = vH.HoraS2
            End Select
            kIncidencia = EntraDentro(Rss!Hora, HoraH, Exceso, Retraso, (I Mod 2) = 0)
            v(I) = kIncidencia
            I = I + 1
            UltimoTicaje = Rss!Hora
            Rss.MoveNext
        Wend
        
        'Ahora ya tenmos si ha llegado tarde, ha salido antes etc, por lo tanto
        ' realizamos los calculos de las horas y generaremos, si cabe
        'las incidencias
        'En v() tenemos que si es 0 nada, pero si es menor tenemos la horas extras
        ' y si es mayor las horas de retraso
        'En t1 tendremos las horas en las incidencias
        T1 = 0
        TieneIncidencia = False
        For I = 0 To 3
            T1 = T1 + v(I)
            If v(I) > 0 Then
                GeneraIncidencia vE.IncRetraso, vMar.Entrada, v(I)
                TieneIncidencia = True
                Else
                    If v(I) < 0 Then
                        GeneraIncidencia vE.IncHoraExceso, vMar.Entrada, Abs(v(I))
                        TieneIncidencia = True
                    End If
            End If
        Next I
        'Debug.Print vMar.IdTrabajador & ": " & T1
        'Stop
        'si tiene dto. Le sumaremos al valor obtenido en T1 el valor de los dtos
        'Comprobamos los dtos almuerzo merienda
        '******************************************************
        QuitarAlmuerzo = False
        quitarmerienda = False
        N = (Rss.RecordCount \ 2)
        If vH.DtoAlm > 0 Then
'            Rss.MoveFirst
'            For I = 1 To N
'                PrimerTicaje = Rss!Hora
'                Rss.MoveNext
'                If PrimerTicaje < vH.HoraDtoAlm Then
              If LeQuitamosElAmluerzo(Rss, vH) Then
                    QuitarAlmuerzo = True
'                    If Rss!Hora > vH.HoraDtoAlm Then QuitarAlmuerzo = True
                End If
'            Next I
        End If
                
        If vH.DtoMer > 0 Then
            For I = 1 To N
                PrimerTicaje = Rss!Hora
                Rss.MoveNext
                If PrimerTicaje <= vH.HoraDtoMer Then
                    If Rss!Hora > vH.HoraDtoMer Then quitarmerienda = True
                End If
            Next I
        End If
            
        
        'Asignamos a la incidencia
        T2 = vH.TotalHoras
        'If QuitarAlmuerzo Then T2 = T2 - vH.DtoAlm
        'If quitarmerienda Then T2 = T2 - vH.DtoMer
        
        'CREO QUE ESTABA MAL PQ le quitaba el almuerzo tb a las horas extra
        'NO se quita el amuerzo. Antes los signos estaban al reves
        '------------------
        If vH.DtoAlm > 0 Then
            If Not QuitarAlmuerzo Then
                
                If T1 >= 0 Then
                    'Stop
                    'Me debe mas horas
                    T1 = T1 - vH.DtoAlm
                Else
                    'Horas extra. Le quito el almuerzo
                    'Stop
                    If T1 < 0 Then
                        T1 = T1 - vH.DtoAlm
                    End If
                End If
            End If

        End If

            
     
        
        TotalH = Round(T2, 2)
        
         
         
         
         
         
        
        '----------------------------------------------
                
        'Una vez asignadas calculamos las horas que le corresponden
        'En el tipo uno, las horas son las horas menos el almuerzo y la merienda
        T2 = TotalH
        T2 = Round(T2 - T1, 2)
        vMar.HorasTrabajadas = T2
        'Asignaremos la incidencia
        'Si tiene manual se queda la manual, si no se queda, si tuviera, la automatica
        If InciManual > 0 Then
            vMar.IncFinal = InciManual
            vMar.HorasIncid = Round(vH.TotalHoras - vMar.HorasTrabajadas, 2)
            Else
                'Vemos si tiene automatica
                If T1 = 0 Then
                    If TieneIncidencia Then
                        'La suma de horas da 0, pero tiene incidencias
                        vMar.IncFinal = vE.IncMarcaje
                        Else
                            vMar.IncFinal = 0
                    End If
                    Else
                        'Falta o sobran horas
                        If T1 > 0 Then
                            'Retraso
                            vMar.IncFinal = vE.IncRetraso
                            Else
                                vMar.IncFinal = vE.IncHoraExtra
                        End If
                        vMar.HorasIncid = Abs(T1)
                End If 't2=0
        End If
        
        
    '   El numero de tikadas no coincide
    Else
        While Not Rss.EOF
             If Rss!IdInci > 0 Then InciManual = Rss!IdInci
             Rss.MoveNext
        Wend
        If InciManual > 0 Then
            vMar.IncFinal = InciManual
            GeneraIncidencia InciManual, vMar.Entrada, 0
            Else
                vMar.IncFinal = vE.IncMarcaje
                GeneraIncidencia vE.IncMarcaje, vMar.Entrada, 0
        End If
        
        
        'Ahora pondremos las horas trabajadas por diferencias
        Rss.MoveFirst
        TotalH = 0
        If (Rss.RecordCount Mod 2) = 0 Then
            While Not Rss.EOF
                'Son pares
                T1 = DevuelveValorHora(Rss!Hora)
                'Siguiente
                Rss.MoveNext
                T2 = DevuelveValorHora(Rss!Hora)
                T2 = T2 - T1
                TotalH = TotalH + T2
                'siguiente par
                Rss.MoveNext
            Wend
            TotalH = Round(TotalH, 2)
        End If
        T1 = 0
        
        
        'Contabilizaremos los descuentos relativos al almuerzo y merienda
            'si procede
                
        QuitarAlmuerzo = False
        quitarmerienda = False
        N = (Rss.RecordCount \ 2)
        If vH.DtoAlm > 0 Then
'            Rss.MoveFirst
'
'            For I = 1 To N
'                PrimerTicaje = Rss!Hora
'                Rss.MoveNext
'                If PrimerTicaje < vH.HoraDtoAlm Then
'                    If Rss!Hora > vH.HoraDtoAlm Then QuitarAlmuerzo = True
'                End If
'            Next I
            QuitarAlmuerzo = LeQuitamosElAmluerzo(Rss, vH)
        End If
                
        If vH.DtoMer > 0 Then
            For I = 1 To N
                PrimerTicaje = Rss!Hora
                Rss.MoveNext
                If PrimerTicaje <= vH.HoraDtoMer Then
                    If Rss!Hora > vH.HoraDtoMer Then quitarmerienda = True
                End If
            Next I
        End If
    
        'Ahora ya sabemos las horas trabajadas
        TotalH = Round(TotalH, 2)
        
        
        
        'Asignamos a la incidencia
        T2 = TotalH
        If QuitarAlmuerzo Then T2 = T2 - vH.DtoAlm
        If quitarmerienda Then T2 = T2 - vH.DtoMer
        
        TotalH = Round(T2, 2)
        
        
        
        'Deberia haber trabajado
        If TotalH > 0 Then
            'Cuanto tiene k trabajar al dia
            T2 = vH.TotalHoras
            'If QuitarAlmuerzo Then T2 = T2 - vH.DtoAlm
            'If quitarmerienda Then T2 = T2 - vH.DtoMer
            T1 = T2 - TotalH
            T1 = Abs(T1)
        Else
            TotalH = 0
        End If
        vMar.HorasTrabajadas = TotalH
        vMar.HorasIncid = T1
    End If   'de numero de tikadas=vh.numtikadas
End If 'De DIAFESTIVO
'Por ultimo marcamos o no el campo correcto
vMar.Correcto = vMar.IncFinal = 0


'Comprobamos si esta de baja
If EsBajaTrabajo(vMar.idTrabajador) Then
    vMar.Correcto = False
   If vMar.IncFinal <> vE.IncMarcaje Then vMar.IncFinal = vE.IncVacaciones    'Es la incidencia de baja
End If



'Grabamos el marcaje
vMar.Modificar
'-------------------------------------------------------------------------
'Cerramos y borramos todos los fichajes pasandolos a una tabla de marcajes
    Rss.MoveFirst
    Set RFin = New ADODB.Recordset
    RFin.CursorType = adOpenKeyset
    RFin.LockType = adLockOptimistic
    RFin.Open "Select * from EntradaMarcajes", Conn, , , adCmdText
    If RFin.EOF Then
        I = 1
        Else
            RFin.MoveLast
            I = RFin!Secuencia + 1
    End If
    While Not Rss.EOF
        RFin.AddNew
        RFin!Secuencia = I
        RFin!idTrabajador = vMar.idTrabajador
        RFin!idMarcaje = vMar.Entrada
        RFin!IdInci = Rss!IdInci
        RFin!Fecha = Rss!Fecha
        RFin!Hora = Rss!Hora
        RFin!horareal = Rss!horareal
        RFin.Update
        I = I + 1
        Rss.MoveNext
    Wend
    RFin.Close
    
    
    'Borramos los ticajes
    Cad = "Delete * from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
    Cad = Cad & " AND Fecha=#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"
    RFin.Open Cad, Conn, , , adCmdText
'Cerramos los recordsets
Rss.Close

Set Rss = Nothing
Set RFin = Nothing

'Adelante con las operaciones
Conn.CommitTrans
Exit Sub
ErrorProcesaMarcaje:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
    Conn.RollbackTrans
End Sub








' El tipo 3 solo controla si es festivo por lo cual todas
' las horas son horas extraso
' y si no es festivo donde todas, repito todas las horas
' son horas trabajadas
Public Sub ProcesarMarcaje_Tipo3(ByRef vMar As CMarcajes)
Dim Rss As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim NumTikadas As Integer
Dim T1 As Single
Dim T2 As Single
Dim kIncidencia As Single
'Dim TieneIncidencia As Boolean
'Dim MarcajeCorrecto As Boolean
'Dim Exceso As Date
'Dim Retraso As Date
Dim I As Long
'Dim v(3) As Single
Dim Cad As String
'Dim HoraH As Date
Dim InciManual As Integer
Dim N As Integer
Dim TotalH As Single
Dim PrimerTicaje As Date
Dim UltimoTicaje As Date

'Ahora ya tenemos las horas tikadas reflejadas
'Comprobamos las horas en funcion de los horarios
'  y calculamos las horas comprobadas

Conn.BeginTrans
Set Rss = New ADODB.Recordset

'Seleccionamos todas las horas de este
Cad = "Select * from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
Cad = Cad & " AND Fecha=#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"
Cad = Cad & " ORDER BY Hora"
Rss.CursorType = adOpenStatic
Rss.Open Cad, Conn, , , adCmdText

If Rss.EOF Then
    'Si no hay ninguna entrada
    Rss.Close
    GoTo ErrorProcesaMarcaje3
End If

InciManual = 0
NumTikadas = Rss.RecordCount

If (NumTikadas Mod 2) > 0 Then
        'Numero de marcajes impares. No podemos calcular horas
        'trabajadas. Generamos error en marcaje
        vMar.IncFinal = vE.IncMarcaje
        vMar.HorasIncid = 0
        vMar.HorasTrabajadas = 0
        GeneraIncidencia vE.IncMarcaje, vMar.Entrada, 0
        Else
            N = NumTikadas \ 2
            TotalH = 0
            'NUMERO DE MARCAJES PAR
            Rss.MoveFirst
            
            'Lo utilizaremos despues para saber si quitamos minutos de almuerzo
            PrimerTicaje = Rss!Hora
            
            '----------------------------------------------
            
                For I = 1 To N
                    T1 = DevuelveValorHora(Rss!Hora)
                    'por si acaso; traen; incidencias; manuales
                    If InciManual = 0 Then InciManual = Rss!IdInci
                    Rss.MoveNext
                    T2 = DevuelveValorHora(Rss!Hora)
                    'Por si trae incidencias manuales
                    If InciManual = 0 Then InciManual = Rss!IdInci
                    UltimoTicaje = Rss!Hora
                    Rss.MoveNext
                    TotalH = TotalH + (T2 - T1)
                Next I
            '******************************************************
            'Comprobamos si hay que quitar los minutos del almuerzo
            If vH.DtoAlm > 0 Then
            
                'aqui aqui aqui
                If LeQuitamosElAmluerzo(Rss, vH) Then
            
                'If PrimerTicaje < vH.HoraDtoAlm Then
                    TotalH = TotalH - vH.DtoAlm
                    If TotalH < 0 Then TotalH = 0
                End If
            End If
            '----------------------------------------------
                
            'Comprobamos si hay que quitar los minutos de la MER
            'Como esta ya en el ultimo
            If vH.DtoMer > 0 Then
                If UltimoTicaje > vH.HoraDtoMer Then
                    TotalH = TotalH - vH.DtoMer
                    If TotalH < 0 Then TotalH = 0
                End If
            End If
            '----------------------------------------------
            '******************************************************
            
            
            'Ahora ya sabemos las horas trabajadas, y las redondeamos
            TotalH = RealizaRedondeo(TotalH)
            
            'Asignamos a la incidencia
             vMar.HorasTrabajadas = TotalH
             'Aqui comprobamos si es festivo o no para asignarle los valores correspondientes
             If vH.EsDiaFestivo Then
                vMar.HorasIncid = TotalH
                vMar.IncFinal = vE.IncHoraExtra
                Else
                    If InciManual > 0 Then GeneraIncidencia InciManual, vMar.Entrada, 0
                    vMar.HorasIncid = 0
                    vMar.IncFinal = InciManual
            End If
End If 'De DIAFESTIVO
'Por ultimo marcamos el campo correcto a FALSE para que los revise a mano
vMar.Correcto = False


'Comprobamos si esta de baja
If EsBajaTrabajo(vMar.idTrabajador) Then
    vMar.Correcto = False
    vMar.IncFinal = DevuelveINC_MARCAJE(vMar.idTrabajador)
End If
'Grabamos el marcaje
vMar.Modificar
'-------------------------------------------------------------------------
'Cerramos y borramos todos los fichajes pasandolos a una tabla de marcajes
    Rss.MoveFirst
    Set RFin = New ADODB.Recordset
    RFin.CursorType = adOpenKeyset
    RFin.LockType = adLockOptimistic
    RFin.Open "Select * from EntradaMarcajes", Conn, , , adCmdText
    If RFin.EOF Then
        I = 1
        Else
            RFin.MoveLast
            I = RFin!Secuencia + 1
    End If
    While Not Rss.EOF
        RFin.AddNew
        RFin!Secuencia = I
        RFin!idTrabajador = vMar.idTrabajador
        RFin!idMarcaje = vMar.Entrada
        RFin!IdInci = Rss!IdInci
        RFin!Fecha = Rss!Fecha
        RFin!Hora = Rss!Hora
        RFin!horareal = Rss!horareal
        RFin.Update
        I = I + 1
        Rss.MoveNext
    Wend
    RFin.Close
    
    
    
    Cad = "Delete * from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
    Cad = Cad & " AND Fecha=#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"
    RFin.Open Cad, Conn, , , adCmdText
'Cerramos los recordsets
Rss.Close

Set Rss = Nothing
Set RFin = Nothing

'Adelante con las operaciones
Conn.CommitTrans
Exit Sub
ErrorProcesaMarcaje3:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
    Conn.RollbackTrans
End Sub





Public Function RealizaRedondeo(ByRef T1 As Single, Optional TipoRedondeo As Byte) As Single

'De momento solo se aplica redondeo al tipo de marcaje 3 (ej. Belgida)
'puesto que en los demas el redondeo se realiza
'revisando marcajes ya que si trabaja las
'horas que le corresponden no hace falta redondear

Dim Entera As Single
Dim resto As Single
Dim Divisor As Integer '
Dim cociente As Integer
Dim v As Single
Dim margen As Single

'Si no hay que redondear
T1 = Round(T1, 2)
RealizaRedondeo = T1
If TipoRedondeo = 0 Then Exit Function

'Seguimos
Select Case TipoRedondeo
Case 2
    Divisor = 25
    margen = 18
Case 3
    Divisor = 50
    margen = 37
Case Else  'Por si acso los recogemos en ELSE que es decima de punto
    Divisor = 10
    margen = 7
End Select
'Cambiamos el valor de t1
Entera = Int(T1)
resto = Round((T1 - Entera) * 100, 0)


v = resto Mod Divisor
cociente = resto \ Divisor
If v > margen Then
    cociente = cociente + 1
End If

v = cociente * Divisor  'Resto redondeado
v = v / 100
RealizaRedondeo = Entera + v
End Function
'Calcularemos si para las horas es correcto o generamos incidencia
Public Function EntraDentro(HoraTicada As Date, HoraHorario As Date, Exc As Date, Ret As Date, EsEntrada As Boolean) As Single
Dim Resul

EntraDentro = 0
If EsEntrada Then
    If HoraTicada >= HoraHorario Then
            'ha llegado tarde
            Resul = HoraTicada - (HoraHorario + Ret)
            If Resul > 0 Then
                'GEneramos la incidencia
                EntraDentro = DevuelveValorHora(HoraTicada - HoraHorario)
            End If
            Else
                'ha llegado antes
                Resul = HoraHorario - (HoraTicada + Exc)
                If Resul > 0 Then
                    'Generamos incidencia H_extra
                    EntraDentro = -1 * DevuelveValorHora(HoraHorario - HoraTicada)
                End If
     End If
     'ELSE
     Else    'es una salida
        'se queda un poco
         If HoraTicada >= HoraHorario Then
               Resul = HoraTicada - (HoraHorario + Exc)
               
               If Resul > 0 Then
                   'GEneramos la incidencia de hora extra
                   EntraDentro = -1 * DevuelveValorHora(HoraTicada - HoraHorario)
               End If
               Else
                   'ha salido antes
                   Resul = HoraHorario - (HoraTicada + Exc)
                   If Resul > 0 Then
                       EntraDentro = DevuelveValorHora(HoraHorario - HoraTicada)
                   End If
        End If
End If
End Function


'La funcion devolvera un 0 si las horas han cuadrado y un 1
'si no han cuadrado

Private Function CalcularHorasComprobadas()
'Dim vMM As CEntradaFichajes
'dimvHH As CHorarios
'Dim i As Integer
'Dim H_Exceso As Date
'Dim H_Defecto As Date
'Dim hora As Single
'Dim TodoOk As Boolean
'
''Por defecto ponemos que no
'CalcularHorasComprobadas = 1
'TodoOk = False
''Vemos la empresa que margenes horarios que tiene
'DevuelveMargenesHorarios H_Exceso, H_Defecto, vMM.idTrabajador
'
''para cada hora miramos si entra dentro de los márgenes
'For i = 1 To vHH.NumTikadas
'    Select Case i
'    Case 1
'        If EntraDentro(vMM.HoraE1, vHH.HoraE1, H_Defecto) Then
'            vMM.HoraE1C = vHH.HoraE1
'            Else
'                vMM.HoraE1C = vMM.HoraE1
'                TodoOk = False
'        End If
'    Case 2
'        If EntraDentro(vMM.HoraS1, vHH.HoraS1, H_Exceso) Then
'            vMM.HoraS1C = vHH.HoraS1
'            Else
'                vMM.HoraS1C = vMM.HoraS1
'                TodoOk = False
'        End If
'    Case 3
'        If EntraDentro(vMM.HoraE2, vHH.HoraE2, H_Defecto) Then
'            vMM.HoraE2C = vHH.HoraE2
'            Else
'                vMM.HoraE2C = vMM.HoraE2
'                TodoOk = False
'        End If
'    Case 4
'        If EntraDentro(vMM.HoraS2, vHH.HoraS2, H_Exceso) Then
'            vMM.HoraS2C = vHH.HoraS2
'            Else
'                vMM.HoraS2C = vMM.HoraS2
'                TodoOk = False
'        End If
'
'    End Select
'Next i
'CalcularHorasComprobadas = Abs(TodoOk)
End Function


Public Function DevuelveNumTrabajador(Cad As String) As Long
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Dim SQL As String

    SQL = "SELECT Trabajadores.IdTrabajador, Trabajadores.NumTarjeta"
    SQL = SQL & " From Trabajadores"
    SQL = SQL & " WHERE (((Trabajadores.NumTarjeta)='" & Cad & "'))"
    RS.Open SQL, Conn, , , adCmdText
    If Not RS.EOF Then
            DevuelveNumTrabajador = DBLet(RS.Fields(0), "N")
        Else
            DevuelveNumTrabajador = -1
    End If
    RS.Close
    Set RS = Nothing
End Function







'Ahora, para este trabajador generaremos los marcajes definitivos
'Es decir, entrada salida etc
'En vSec tenemos el numero de secuencia para insertar en fichajes
Public Function GeneraUnmarcajeAlzicoop(NTarjeta As String, Codigo As Long, vFecha As Date, ByRef vSec As Long) As Byte
'  ANTES Public Function GeneraUnmarcajeAlzicoop(NTarjeta As String, Codigo As Long, vFecha As Date, ByRef vSec As Long) As Byte
Dim RS As ADODB.Recordset
Dim RsAUX As Recordset
Dim Cad As String
Dim I As Integer
Dim H1 As Date
Dim h2 As Date
Dim Entrada As Boolean
Dim Aux As Byte


On Error GoTo ErrGeneraUnmarcajeAlzicoop
GeneraUnmarcajeAlzicoop = 1
Cad = "Select * from TipoAlzicoop WHERE Tarjeta='" & NTarjeta & "'"
Cad = Cad & " AND Fecha=#" & Format(vFecha, "yyyy/mm/dd") & "#  ORDER BY Hora"
Set RS = New ADODB.Recordset
RS.Open Cad, Conn, , , adCmdText

If Not RS.EOF Then
    Set RsAUX = New ADODB.Recordset
    RsAUX.CursorType = adOpenKeyset
    RsAUX.LockType = adLockOptimistic
    RsAUX.Open "EntradaFichajes", Conn, , , adCmdTable

    
'--------->  ANTES GENERAMOS NOSOTROS LAS ENTRADAS Y SALIDAS EN FUNCION DE BLA BLA
' --- AORA CAD MARCAJE SE RECOGE EN LA TABLA
''Entrada = False
''
''
''While Not Rs.EOF
''    'Vemos si el marcaje es una salida
''    'Si lo es mandaremos a generar la entrada y la salida
''    If Rs.Fields(4) = "233" And Rs.Fields(5) = "045" Then
''            'Cad = Cad & "(Salida)"
''            h2 = Rs.Fields(2)
''            'Aqui mandaremos a generar
''            If Entrada Then
''                aux = 0
''                Else
''                    aux = 2
''            End If
''            GeneraEntradaFichajesALZ h1, h2, aux, vSec, Codigo, vFecha
''
''            'Una vez generado ponemos entrada a FALSE
''            Entrada = False
''            Else
''                If Not Entrada Then
''                    h1 = Rs.Fields(2)
''                    Entrada = True
''                End If
''    End If
''    Rs.MoveNext
''Wend
''Rs.Close
''If Entrada Then
''    GeneraEntradaFichajesALZ h1, h2, 1, vSec, Codigo, vFecha
''    'El 1 signifca solo la entrada
''End If
    '-------------  AHORA  -------------------
    While Not RS.EOF
    
    
        RsAUX.AddNew
        RsAUX!Secuencia = vSec
        RsAUX!idTrabajador = Codigo
        RsAUX!Fecha = vFecha
        RsAUX!Hora = RS!Hora
        
        'Nuevo
        RsAUX!horareal = RS!Hora
        
        RsAUX!IdInci = 0
        RsAUX.Update
        vSec = vSec + 1
        'Siguiente
        RS.MoveNext
    Wend
    RsAUX.Close
End If 'De rs.eof
RS.Close
'Borramos los marcajes en TABLAALZICOOP
Cad = "DELETE from TipoAlzicoop WHERE Tarjeta='" & NTarjeta & "'"
Cad = Cad & " AND Fecha=#" & Format(vFecha, "yyyy/mm/dd") & "#"
Conn.Execute Cad

'Salida
Set RS = Nothing
GeneraUnmarcajeAlzicoop = 0
Exit Function
ErrGeneraUnmarcajeAlzicoop:
    
    
End Function

'Los marcajes tipo ALZICOOP proceden de un control de produccion
'Por ello se recogen tantos marcajes como cambios de actividad realiza el empleado.
'Para calcular las horas tendremos que ir buscando los marcajes
'de salida. Una vez encontrado, el primer marcaje anterior a el sera de entrada

Private Function GeneraEntradaFichajesALZ(tick1 As Date, tick2 As Date, Kmarcaje As Byte, ByRef vContador As Long, vCod As Long, vFe As Date) As Byte
Dim I As Integer
Dim RS As ADODB.Recordset
Dim C As Integer
Dim Final As Integer

'En kMarcaje tendremos
    '0  .- Entrada y salida
    '1  .- Solo la entrada
    '2  .- Solo la salida
If Kmarcaje = 0 Then
    C = 1
    Final = 2
    Else
    If Kmarcaje = 1 Then
        C = 1
        Final = 1
        Else
            C = 2
            Final = 2
    End If
End If
Set RS = New ADODB.Recordset
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open "EntradaFichajes", Conn, , , adCmdTable
    For I = C To Final
        RS.AddNew
        RS!Secuencia = vContador
        RS!idTrabajador = vCod
        RS!Fecha = vFe
        If I = 1 Then
            RS!Hora = tick1
            Else
            RS!Hora = tick2
        End If
        RS!IdInci = 0
        RS.Update
        vContador = vContador + 1
    Next I
    RS.Close
    Set RS = Nothing
End Function





'--------------------------------------------------------------------------------
'Cuando solo vemos si el num de tikadas es par y calculamos las horas 'trabajadas
'en funcion del horario. Si es
'
Public Sub ProcesarMarcaje_Tipo2(ByRef vMar As CMarcajes)
Dim Rss As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim NumTikadas As Integer
Dim T1 As Single
Dim T2 As Single
Dim I As Long
Dim Cad As String
Dim N As Integer
Dim TotalH As Single
Dim HoE As Single

'Ahora ya tenemos las horas tikadas reflejadas
'Comprobamos las horas en funcion de los horarios
'  y calculamos las horas comprobadas
 Conn.BeginTrans
Set Rss = New ADODB.Recordset

'Seleccionamos todas las horas de este
Cad = "Select * from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
Cad = Cad & " AND Fecha=#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"
Cad = Cad & " ORDER BY Hora"
Rss.CursorType = adOpenStatic
Rss.Open Cad, Conn, , , adCmdText

If Rss.EOF Then
    'Si no hay ninguna entrada
    Rss.Close
    GoTo ErrorProcesaMarcaje_Tipo2
End If


'Si el numero de tikadas es par entonces calculamos las horas
NumTikadas = Rss.RecordCount
If (NumTikadas Mod 2) > 0 Then
    'Numero de marcajes impares. No podemos calcular horas
    'trabajadas. Generamos error en marcaje
    vMar.IncFinal = vE.IncMarcaje
    GeneraIncidencia vE.IncMarcaje, vMar.Entrada, 0
    vMar.HorasIncid = 0
    vMar.HorasTrabajadas = 0
    Else
        N = NumTikadas \ 2
        TotalH = 0
        'NUMERO DE MARCAJES PAR
        Rss.MoveFirst
        PrimerTicaje = Rss!Hora  'Almacenamos el primer ticaje
        For I = 1 To N
            T1 = DevuelveValorHora(Rss!Hora)
            Rss.MoveNext
            UltimoTicaje = Rss!Hora 'Obtendremos el ultimo marcaje
            T2 = DevuelveValorHora(Rss!Hora)
            Rss.MoveNext
            TotalH = TotalH + (T2 - T1)
        Next I
        
        'Comprobamos los detos almuerzo merienda
        '******************************************************
        'Comprobamos si hay que quitar los minutos del almuerzo
        If vH.DtoAlm > 0 Then
        '    If PrimerTicaje < vH.HoraDtoAlm Then
         '       TotalH = TotalH - vH.DtoAlm
         '       If TotalH < 0 Then TotalH = 0
         '   End If
         
            If LeQuitamosElAmluerzo(Rss, vH) Then
                TotalH = TotalH - vH.DtoAlm
                If TotalH < 0 Then TotalH = 0
            End If
        End If
        '----------------------------------------------
            
        'Comprobamos si hay que quitar los minutos de la MER
        'Como esta ya en el ultimo
        If vH.DtoMer > 0 Then
            If UltimoTicaje > vH.HoraDtoMer Then
                TotalH = TotalH - vH.DtoMer
                If TotalH < 0 Then TotalH = 0
            End If
        End If
        '----------------------------------------------
        '******************************************************
        'Ahora ya sabemos las horas trabajadas
        TotalH = Round(TotalH, 2)
        'Vemos si es diafestivo o no
        'si lo es todas son horas extras, si no
        'calculamos
        If vH.EsDiaFestivo Then
            vMar.HorasTrabajadas = TotalH
            vMar.HorasIncid = TotalH
            vMar.IncFinal = vE.IncHoraExtra
            'ELSE
            Else     'No es festivo
            'ELSE
            HoE = EntraDentro2(TotalH, vH.TotalHoras, vE.MaxExceso, vE.MaxRetraso)
            If HoE = 0 Then
                vMar.HorasTrabajadas = vH.TotalHoras
                vMar.HorasIncid = 0
                vMar.Correcto = True
                Else
                    vMar.Correcto = False
                    If HoE < 0 Then
                        'Horas extras
                        vMar.HorasTrabajadas = vH.TotalHoras - HoE
                        vMar.HorasIncid = Abs(HoE)
                        vMar.IncFinal = vE.IncHoraExtra
                        Else
                            'retraso, no ha llegado al minimo exigible
                            vMar.HorasTrabajadas = vH.TotalHoras - HoE
                            vMar.HorasIncid = HoE
                            vMar.IncFinal = vE.IncRetraso
                            'Ya que despues no quedara constancia ya que sera anulada
                            'para pasar anominas
                            'Ademas genreamos la incidencia de retraso correspondiente
                            GeneraIncidencia vE.IncRetraso, vMar.Entrada, HoE
                    End If
            End If
        End If
End If


''Comprobamos si esta de baja
If EsBajaTrabajo(vMar.idTrabajador) Then
    vMar.Correcto = False
   If vMar.IncFinal <> vE.IncMarcaje Then vMar.IncFinal = vE.IncVacaciones    'Es la incidencia de baja
End If



'Grabamos el marcaje
vMar.Modificar
'-------------------------------------------------------------------------
'Cerramos y borramos todos los fichajes pasandolos a una tabla de marcajes
    Rss.MoveFirst
    Set RFin = New ADODB.Recordset
    RFin.CursorType = adOpenKeyset
    RFin.LockType = adLockOptimistic
    RFin.Open "Select * from EntradaMarcajes", Conn, , , adCmdText
    If RFin.EOF Then
        I = 1
        Else
            RFin.MoveLast
            I = RFin!Secuencia + 1
    End If
    While Not Rss.EOF
        RFin.AddNew
        RFin!Secuencia = I
        RFin!idTrabajador = vMar.idTrabajador
        RFin!idMarcaje = vMar.Entrada
        RFin!IdInci = Rss!IdInci
        RFin!horareal = Rss!horareal
        RFin!Fecha = Rss!Fecha
        RFin!Hora = Rss!Hora
        RFin.Update
        I = I + 1
        Rss.MoveNext
    Wend
    RFin.Close
    
    Cad = "Delete * from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
    Cad = Cad & " AND Fecha=#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"
    RFin.Open Cad, Conn, , , adCmdText
'Cerramos los recordsets
Rss.Close

Set Rss = Nothing
Set RFin = Nothing
'Adelante con las operaciones
Conn.CommitTrans
Exit Sub
ErrorProcesaMarcaje_Tipo2:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
    Conn.RollbackTrans
End Sub



'Dadas unas horas a trbajar y unas trabajadas me dira si
'esta dentro de las que le tocaban o
'devolvera horas extras o retrasos
Public Function EntraDentro2(HoraTotales As Single, HorasHorario As Single, Exc As Single, Ret As Single) As Single
Dim Resul
Dim Valor As Single

        Valor = 0
        'se queda un poco
         If HoraTotales >= HorasHorario Then
               Resul = HoraTotales - (HorasHorario + Exc)
               If Resul > 0 Then
                   'GEneramos la incidencia de hora extra
                   Valor = -1 * (HoraTotales - HorasHorario)
               End If
               Else
                   'ha salido antes
                   Resul = HorasHorario - (HoraTotales + Ret)
                   If Resul > 0 Then
                       Valor = HorasHorario - HoraTotales
                   End If
        End If
        EntraDentro2 = Round(Valor, 2)
End Function



Public Function YaExistenMarcajes(Cod As Integer, Fecha As Date) As Long
Dim RS As ADODB.Recordset
Dim SQL As String
    YaExistenMarcajes = -1
    Set RS = New ADODB.Recordset
    SQL = "SELECT Entrada" & _
        " FROM Marcajes WHERE " & _
        " IdTrabajador=" & Cod & _
        " AND Fecha=#" & Format(Fecha, "yyyy/mm/dd") & "#"
    RS.Open SQL, Conn, , , adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then _
            YaExistenMarcajes = RS.Fields(0)
    End If
    RS.Close
    Set RS = Nothing
End Function


'Public Function QuitarHoras(Horas As Single, TipoMarcaje As Byte) As Single
'Dim AUx As Single
'
'AUx = Horas - VH.DtoAlm
''De momento no utilizamos tipomarcaje
'QuitarHoras = 1
'End Function




Public Function LeQuitamosElAmluerzo(ByRef dRS As ADODB.Recordset, ByRef dH As CHorarios) As Boolean
Dim Fin As Boolean

    LeQuitamosElAmluerzo = False
    
    dRS.MoveFirst
    Fin = False
    Do
        'Si el primer ticaje, ya es posterior a la hora del almuerzo
        If dRS!Hora > dH.HoraDtoAlm Then Exit Function
    
        dRS.MoveNext
        
        If dRS.EOF Then Exit Function
        'Segundo ticaje
        'Ticaje menor. k la hora de almuerzo. Vemos si no ha salido
        If dRS!Hora < dH.HoraDtoAlm Then
            'Ha salido antes de comienzo almuerzo
            'No hago nada
        Else
            LeQuitamosElAmluerzo = True
            Exit Function
        End If
        
        dRS.MoveNext
            
        If dRS.EOF Then Fin = True
    Loop Until Fin
    
    
End Function
