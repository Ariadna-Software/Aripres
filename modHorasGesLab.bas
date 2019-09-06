Attribute VB_Name = "modHorasGesLab"
Option Explicit

Private Const HoraIntermediaMiercolesSabado = "14:00:00"


'Para las PUTAS compensaciones de los miercoles / sabado
Private SemanaMesPrimera As Integer
Private SemanaMesUltima As Integer

Public Depuracion As Boolean
Private TrabajadoresDepuracion As String
Private MiNF As Integer  'si depuramos meteremos los datos aqui
Private cDep2 As Collection 'iremos metiendo las lineas aqui y luego las pasremos al fichero




'Modificacion COOPIC . SI es finde semana, NO cuentan
Public Function CalculaHorasHorario(idCal As Integer, IdHor As Integer, ByRef Dias As Integer, Fini As Date, FFin As Date) As Currency

Dim Sum As Currency
Dim vH As CHorarios
Dim F As Date
Dim d As Currency
Dim Semana As Integer
Dim UltimoMiercolesTrabajado As Integer
Dim DiaDeLaSemana As Integer
Dim ProcesamosDia As Boolean

    Set vH = New CHorarios
    
    CalculaHorasHorario = -1
    Sum = 0
    d = 0
    F = Fini
    
    'Para cada dia del mes
    
    Do
        If vH.Leer(IdHor, F, idCal) = 1 Then Exit Function
        
        
        If Not vH.EsDiaFestivo Then
        
            DiaDeLaSemana = Weekday(F)
            ProcesamosDia = True
            If vEmpresa.QueEmpresa = 5 Then
                If DiaDeLaSemana = 7 Or DiaDeLaSemana = 1 Then ProcesamosDia = False
            End If
            
            
            If ProcesamosDia Then
                If vH.DiaNomina = 0.5 Then
                    Semana = Format(F, "ww")
                    If Weekday(F) = 4 Then
                        UltimoMiercolesTrabajado = Semana
                        d = d + 1
                    Else
                        If UltimoMiercolesTrabajado <> Semana Then d = d + 1
                    End If
            
                Else
                    d = d + vH.DiaNomina
                End If
                Sum = Sum + vH.TotalHoras
            End If 'fin de semana
        End If
        F = DateAdd("d", 1, F)
    Loop Until F > FFin
    'Redondeamos siempre hacia arriba
    Dias = Int(d)
    If d > Int(d) Then
        'Tiene fraccion de dia
        Dias = Int(d) + 1
    End If
    CalculaHorasHorario = Sum
End Function





'Para las bajas.
'Puede ser que un trabajador tenga varias bajas el mismo mes.
'Tendremos que enviar cual fue el utlimo miercoles trabajado
Public Function CalculaHorasHorarioBaja(idCal As Integer, IdHor As Integer, ByRef Dias As Integer, Fini As Date, FFin As Date) As Currency
Dim Sum As Currency
Dim vH As CHorarios
Dim F As Date
Dim d As Currency
Dim Semana As Integer

    Set vH = New CHorarios
    CalculaHorasHorarioBaja = 0
    Sum = 0
    d = 0
    F = Fini
    
    'De lunes a domingo  5 dias. Festivos saba y domingo
    
    Do
        If vH.Leer(IdHor, F, idCal) = 1 Then Exit Function
        
        If Not vH.EsDiaFestivo Then
            
            Semana = Format(F, "w")
            If Semana <= 6 Then
               d = d + 1
    
               Sum = Sum + vH.TotalHoras

            End If
            
       End If
       F = DateAdd("d", 1, F)
    Loop Until F > FFin
    'Redondeamos siempre hacia arriba
    Dias = Int(d)
    If d > Int(d) Then
        'Tiene fraccion de dia
        Dias = Int(d) + 1
    End If
    CalculaHorasHorarioBaja = Sum
End Function




'Calculo de horas. Simplemente es dias * 8
Public Function CalculaHorasHorarioALZ(IdHor As Integer, ByRef Dias As Integer, Fini As Date, FFin As Date) As Currency
Dim vH As CHorarios
Dim F As Date
Dim d As Currency
Dim EsFestivo As Boolean
Dim k As Integer

    Set vH = New CHorarios
    
    CalculaHorasHorarioALZ = -1
'    Sum = 0
    d = 0
    F = Fini
    'Para cada dia del mes
    Do
        If vH.Leer(IdHor, F, 1) = 1 Then Exit Function
        
        EsFestivo = False
        If vH.EsDiaFestivo Then
            EsFestivo = True
        Else
            'De momento para todas las cooperativas
            k = Format(F, "w")
            If k = 7 Then EsFestivo = True
        End If
        If Not EsFestivo Then
            d = d + 1   'vH.DiaNomina
        End If
        F = DateAdd("d", 1, F)
    Loop Until F > FFin
    'Redondeamos siempre hacia arriba
    Dias = Int(d)
    If d > Int(d) Then
        'Tiene fraccion de dia
        Dias = Int(d) + 1
    End If
    CalculaHorasHorarioALZ = Dias * 8
End Function



Public Function CalculaHorasHorarioALZConVector(IdHor As Integer, ByRef Dias As Integer, Fini As Date, FFin As Date, VectorDiasTrabajados As String) As Currency
Dim vH As CHorarios
Dim F As Date
Dim d As Currency
Dim EsFestivo As Boolean
Dim k As Integer
Dim ComoEsElDia As String
    Set vH = New CHorarios
    
    CalculaHorasHorarioALZConVector = -1

    d = 0
    F = Fini
    'Para cada dia del mes
    
    Do
        
        If vH.Leer(IdHor, F, 1) = 1 Then Exit Function
        
        EsFestivo = False
        ComoEsElDia = "N"
        If vH.EsDiaFestivo Then
            EsFestivo = True
            
        Else
            'De momento para todas las cooperativas
            k = Format(F, "w")
            If k = 7 Then EsFestivo = True
        End If
        If Not EsFestivo Then
            d = d + 1   'vH.DiaNomina
        Else
            ComoEsElDia = "F"
        End If
        VectorDiasTrabajados = VectorDiasTrabajados & ComoEsElDia
        
        
        
        
        F = DateAdd("d", 1, F)
    Loop Until F > FFin
    'Redondeamos siempre hacia arriba
    Dias = Int(d)
    If d > Int(d) Then
        'Tiene fraccion de dia
        Dias = Int(d) + 1
    End If
    CalculaHorasHorarioALZConVector = Dias * 8
End Function






'Calcula las horas trabajadas para los trabajadores k tiene la marca puesta
Public Sub CalculaHorasTrabajadas(Fini As Date, FFin As Date, ControlNomina As Byte, UnaSeccionSolo As Integer)
Dim FAux As Date
Dim FAux2 As Date
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Dias As Currency
Dim Trabajador As Long
Dim Aux As String
Dim SQL As String
Dim vH As CHorarios
Dim FESTIVOS2 As String
Dim MEDIODIA As String
Dim strControlNomina As String
'----------------------------------------------
'FALTA### parametrizar esto
Dim UltimoMiercolesTrabajado As Integer
Dim Semana As Integer
Dim BAJAS As String
Dim RF As ADODB.Recordset


Dim idCal As Integer
    
    idCal = 1
    
    
    Select Case ControlNomina
    Case 0
        strControlNomina = " AND Trabajadores.ControlNomina >0  AND Trabajadores.ControlNomina <=2 "
    Case 1
        strControlNomina = " AND Trabajadores.ControlNomina = 3"
    Case 2
        strControlNomina = " AND (Trabajadores.ControlNomina =1  OR Trabajadores.ControlNomina =3) "
    Case 3
        'Sera para el listado que se entraga a los trabbajdores en PICASSENT
        ' Es para los tipos 1,2,3
        strControlNomina = " AND Trabajadores.ControlNomina >0"
    Case Else
        strControlNomina = ""
    End Select
    
    If UnaSeccionSolo >= 0 Then strControlNomina = " AND Trabajadores.seccion =" & UnaSeccionSolo
        
    conn.Execute "Delete from tmpHoras "
    
   
     Set RS = New ADODB.Recordset
    
    If vEmpresa.QueEmpresa = 4 Then
        'CATADAU. Van separadas
        SQL = "INSERT INTO tmpHoras(trabajador,HorasT,HorasC,HorasE) "
        SQL = SQL & " SELECT  jornadassemanalesalz.idtrabajador,sum(if(tipohoras=0,horastrabajadas,0)) normales ,sum(if(tipohoras=1,horastrabajadas,0)) estruc"
        SQL = SQL & " ,sum(if(tipohoras=2,horastrabajadas,0)) extra"
        SQL = SQL & " from jornadassemanalesalz ,trabajadores where jornadassemanalesalz.idtrabajador=trabajadores.idtrabajador"
        SQL = SQL & " AND jornadassemanalesalz.Fecha >= " & DBSet(Fini, "F")
        SQL = SQL & " and jornadassemanalesalz.Fecha <= " & DBSet(FFin, "F")
        SQL = SQL & strControlNomina
        SQL = SQL & " GROUP BY jornadassemanalesalz.idTrabajador;"
        conn.Execute SQL
        
        
        
    Else
        'COOPIC. Las normales y esctructurales van JUNTAS
        SQL = "INSERT INTO tmpHoras(trabajador,HorasT) "
        SQL = SQL & " SELECT  jornadassemanalesalz.idtrabajador,sum(if(tipohoras<2,horastrabajadas,0)) "
        SQL = SQL & " from jornadassemanalesalz ,trabajadores where jornadassemanalesalz.idtrabajador=trabajadores.idtrabajador"
        SQL = SQL & " AND jornadassemanalesalz.Fecha >= " & DBSet(Fini, "F")
        SQL = SQL & " and jornadassemanalesalz.Fecha <= " & DBSet(FFin, "F")
        SQL = SQL & strControlNomina
        SQL = SQL & " GROUP BY jornadassemanalesalz.idTrabajador;"
        conn.Execute SQL
        
        
        
        
        'Las horas EXTRA
        SQL = "SELECT  jornadassemanalesalz.idtrabajador,sum(if(tipohoras=2,horastrabajadas,0)) sumadehoras "
        SQL = SQL & " from jornadassemanalesalz ,trabajadores where jornadassemanalesalz.idtrabajador=trabajadores.idtrabajador"
        SQL = SQL & " AND jornadassemanalesalz.Fecha >= " & DBSet(Fini, "F")
        SQL = SQL & " and jornadassemanalesalz.Fecha <= " & DBSet(FFin, "F")
        SQL = SQL & strControlNomina
        SQL = SQL & " GROUP BY jornadassemanalesalz.idTrabajador;"
        
       
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            If DBLet(RS!sumadehoras, "N") <> 0 Then
                SQL = "UPDATE tmpHoras Set HorasC = " & TransformaComasPuntos(RS!sumadehoras)
                SQL = SQL & " WHERE Trabajador = " & RS!idTrabajador
            
                conn.Execute SQL
            End If
            RS.MoveNext
        Wend
        RS.Close
        
    End If
    
    'Updatemos con los dias trabajados.
    '
    'Acciones:
    '       -En una variable cargaremos los dias festivos de
    '       -En Otra Cargaremos los medios dias.
    '       -Para cada dia trabajado, para cada trabajador, veremos
    '       - Si los dias trabajados es un festivo o unidad fraccionarai
    
    SQL = "SELECT idHorario,idcal "
    SQL = SQL & " FROM calendariol"
    SQL = SQL & " Where calendariol.Fecha >= '" & Format(Fini, FormatoFecha) & "'"
    SQL = SQL & " and calendariol.Fecha <= '" & Format(FFin, FormatoFecha) & "'"
    SQL = SQL & " GROUP BY idHorario,idcal;"
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    
    While Not RS.EOF
        Set vH = New CHorarios
        idCal = RS!idCal
        If vH.Leer(RS!IdHorario, Now, idCal) = 0 Then
            FESTIVOS2 = vH.LeerDiasFestivos(vH.IdHorario, Fini, FFin)
            
            'NO HAY medios dias
            'Lo dejo por si pasa PICASSENT
            'MEDIODIA = vH.LeerMediosDias(vH.IdHorario, Fini, FFin)
             
             Set Rs2 = New ADODB.Recordset
    

    
           
    
    
    

                        
            SQL = "SELECT jornadassemanalesalz.idTrabajador,jornadassemanalesalz.fecha, "
            'Febrero 2019
            'En catadau, toooodos los dias que vienen son LABORABLES
            If vEmpresa.QueEmpresa = 4 Then SQL = SQL & " 1 "
            
            
            SQL = SQL & " laborable FROM Trabajadores INNER JOIN jornadassemanalesalz ON Trabajadores.IdTrabajador = jornadassemanalesalz.idTrabajador "
            SQL = SQL & " WHERE jornadassemanalesalz.Fecha >= '" & Format(Fini, FormatoFecha) & "'"
            SQL = SQL & " and jornadassemanalesalz.Fecha <= '" & Format(FFin, FormatoFecha) & "'"
            SQL = SQL & strControlNomina
            SQL = SQL & " AND idcal = " & idCal
            SQL = SQL & " GROUP BY jornadassemanalesalz.idTrabajador, Fecha"
            SQL = SQL & " ORDER BY jornadassemanalesalz.idTrabajador, Fecha"
            Set Rs2 = New ADODB.Recordset
            Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            If Not Rs2.EOF Then
                Trabajador = -1
                Do
                   
                          
                
                   If Trabajador <> Rs2!idTrabajador Then
                         
                        
                        
                        If Trabajador > 0 Then
                            SQL = "UPDATE tmpHoras Set Dias = "
                            If Dias > Int(Dias) Then
                                Dias = Int(Dias) + 1
                            Else
                                Dias = Int(Dias)
                            End If
                            SQL = SQL & Int(Dias)
                            SQL = SQL & " WHERE Trabajador = " & Trabajador
                            conn.Execute SQL
                        End If
                   
                   
                   
                        'Por si tiene bajas
                        'Por si esta de baja.  Solo podria trabajar el primer dia de la baja
                        'Picassent
                        If True Then
                             Set RF = New ADODB.Recordset
                             BAJAS = "Select * from bajas where idtrab=" & Rs2!idTrabajador
                             BAJAS = BAJAS & " AND Fechabaja >= '" & Format(Fini, FormatoFecha) & "'"
                             BAJAS = BAJAS & " and Fechabaja <= '" & Format(FFin, FormatoFecha) & "'"
                             RF.Open BAJAS, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                             BAJAS = ""
                             If Not RF.EOF Then
                                 While Not RF.EOF
                                     BAJAS = BAJAS & Format(RF!FechaBaja, "dd/mm/yyyy") & "|"
                                     RF.MoveNext
                                 Wend
                             End If
                             RF.Close
                             Set RF = Nothing
                        End If
                   
                   
                   
                   
                   
                   
                   
                   
                        Trabajador = Rs2!idTrabajador
                        Dias = 0
                        UltimoMiercolesTrabajado = 0
                    End If
    
                    'Si el dia esta en FESTIVOS no lo sumo
                    Aux = Format(Rs2!Fecha, "dd/mm/yyyy") & "|"
    
    
                    If vEmpresa.QueEmpresa = 4 Then
                        'EN CATADAU, si es el dia lo ha trabajado (max(5)) semana
    
                        'Si esta de baja
                        If InStr(1, BAJAS, Aux) > 0 Then MsgBox "Trabaja dia BAJA"
    
                        If Rs2!Laborable = 1 Then
                            Dias = Dias + 1
                        Else
                            'SI ES FESTIVO  TAMBIEN LO MARCAMOS
                            If InStr(1, FESTIVOS2, Aux) = 1 Then Dias = Dias + 1
                        End If
                    Else
                        'NO esta en festivos
                        If InStr(1, FESTIVOS2, Aux) = 0 Then
                        
                            'Si esta de baja
                            If InStr(1, BAJAS, Aux) = 0 Then
                                'Si es medio dia sumo medio
                                If InStr(1, MEDIODIA, Aux) > 0 Then
                                    Semana = Format(Rs2!Fecha, "ww")
                                    If Weekday(Rs2!Fecha) = 4 Then
                                        Dias = Dias + 1
                                        UltimoMiercolesTrabajado = Semana
                                    Else
                                        If UltimoMiercolesTrabajado <> Semana Then Dias = Dias + 1
        
                                    End If
                                Else
                                    Dias = Dias + 1
                                End If
                            Else
                                'QUITAR###. Trabajan el primer dia de baja
                               '
                               
                               'En CATADAU. Los festivos tambien SON cotizables
                               'S top
                               MsgBox "Trabaja el dia de la baja: " & Rs2!Fecha & "  " & Trabajador, vbExclamation
                            End If   'bajas
                        End If       'festivos
                    End If 'DE CATADAU
                    'Sig
                    Rs2.MoveNext
                Loop Until Rs2.EOF
                
                
                
                'Ahora faltara por hacer el ultimo trabajador
                SQL = "UPDATE tmpHoras Set Dias = "
                If Dias > Int(Dias) Then
                    Dias = Int(Dias) + 1
                Else
                    Dias = Int(Dias)
                End If
                SQL = SQL & Int(Dias)
                SQL = SQL & " WHERE Trabajador = " & Trabajador
                conn.Execute SQL
            End If
        End If
        RS.MoveNext 'Siguiente horario
    Wend
        
        
        
        
        
        
        
        
        
        
        
    'Por si acaso algun trabajador tiene numeros negativos
    SQL = "UPDATE tmpHoras Set Dias = 0"
    SQL = SQL & " WHERE Dias < 0 "
    conn.Execute SQL
    
    
    Set RS = Nothing
End Sub




Public Sub CalculaDatosMes(Fini As Date, FFin As Date, ControlNomina As Byte, UnaSeccionSolo As Integer)
Dim FAux As Date
Dim FAux2 As Date
Dim RS As ADODB.Recordset
Dim Horas As Currency
Dim d As Integer
Dim Aux As String
Dim SQL As String
Dim strControlNomina As String
Dim D22 As Integer
Dim h22 As Currency
Dim IDT As Integer
Dim IDH As Integer
Dim vM As CMarcajes
Dim HorasBaja As Currency
Dim UltMierTrabajado As Integer

Dim idCal As Integer
Dim varHorario As Integer

Dim maxDias As Integer
Dim maxHoras As Currency


    'IMPORTANTE
    'Ahora hay un control nomina mas, k es el 2
    'El tipo de control 2: Tiene un suledo fijo al mes
    'Pero en anticpos solo anticipa hNormales
    'luego el calculo de horas es el mismo que el 1
    ' por lo tanto donde ponia
        'SQL = SQL & " AND Trabajadores.ControlNomina = 1"
    ' pondra ahora
        'SQL = SQL & " AND Trabajadores.ControlNomina > 0"


    'Otro MAS. El tipo 3
    '   40 Horas semanales. 5 dias semana
    '
    'Con lo cual si en
    ' controlnomina
        ' 1.-   NORMAL ControlNomina >0 and ControlNomina <3
        ' 2.- Solo para el tipo  3
    If ControlNomina = 0 Then
        strControlNomina = " AND Trabajadores.ControlNomina >0  AND Trabajadores.ControlNomina <3"
    Else
        strControlNomina = " AND Trabajadores.ControlNomina = 3"
    End If
    
    If UnaSeccionSolo >= 0 Then strControlNomina = " AND Trabajadores.seccion =" & UnaSeccionSolo


    
    
    '-------   Datos teroicos del mes
    conn.Execute "Delete from tmpDatosMes"
    
    'Creamos todos los trabajadores con las horas y dias k
    'Deberian haber trabajado en el mes completo( y no esten de baja)
    SQL = "INSERT INTO tmpDatosMes(Mes,Trabajador,MesHoras,MesDias,idHorarioTra)"
    SQL = SQL & " SELECT " & Month(Fini) & ", Trabajadores.IdTrabajador, tmpHorasMesHorario.Horas, tmpHorasMesHorario.Dias"   ', Trabajadores.FecBaja"
    SQL = SQL & " ,idcal FROM Trabajadores INNER JOIN tmpHorasMesHorario ON Trabajadores.idcal = tmpHorasMesHorario.idHorario"
    SQL = SQL & " WHERE (Trabajadores.FecBaja Is Null) "
    SQL = SQL & strControlNomina
    SQL = SQL & " AND (Trabajadores.FecAlta < '" & Format(Fini, FormatoFecha) & "')"
    conn.Execute SQL


    Set RS = New ADODB.Recordset
    idCal = 1
    
    

    
    'Ahora vemo los k entraron a trabajar este periodo.
    '¡Descontaremos de las horas laborables los dias k no han trabajado
    'o dicho de otra forma. Le contamos solo las horas k debia haber trabajado en fechas de alta
    SQL = "Select idTrabajador,idcal,fecalta,fecbaja,idcal from Trabajadores WHERE"
    SQL = SQL & " fecalta >='" & Format(Fini, FormatoFecha) & "'"
    SQL = SQL & " and fecalta <='" & Format(FFin, FormatoFecha) & "'"
    SQL = SQL & strControlNomina
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "INSERT INTO tmpDatosMes(Mes,Trabajador,MesHoras,MesDias,idHorarioTra) VALUES (" & Month(Fini) & ","
    varHorario = -1
    While Not RS.EOF
    
        
        
        If varHorario <> RS!idCal Then
            varHorario = RS!idCal
            'Veamos el maximo de dias a trabajar
            Aux = DevuelveDesdeBD("concat(dias,'|',horas,'|')", "tmphorasmeshorario", "idhorario", CStr(varHorario))
            If Aux = "" Then Aux = "60|500|"
            maxDias = RecuperaValor(Aux, 1)
            maxHoras = TransformaPuntosComas(RecuperaValor(Aux, 2))
        End If
    
        If IsNull(RS!FecBaja) Then
            FAux = FFin
        Else
            FAux = RS!FecBaja
            If FAux > FFin Then FAux = FFin
        End If
        Horas = CalculaHorasHorario(idCal, RS!idCal, d, RS!FecAlta, FAux)
        
        
        If d > maxDias Then d = maxDias
        If Horas > maxHoras Then Horas = maxHoras
    
        
        
        Aux = RS.Fields!idTrabajador & "," & TransformaComasPuntos(CStr(Horas)) & "," & d & "," & RS!idCal & ")"
        conn.Execute SQL & Aux
        RS.MoveNext
    Wend
    RS.Close
    varHorario = -1
    
    'AHora vemos los k se han dado de baja en este periodo
    SQL = "Select idTrabajador,idcal,fecalta,fecbaja,idcal from Trabajadores WHERE"
    SQL = SQL & " fecalta <'" & Format(Fini, FormatoFecha) & "'"
    SQL = SQL & " AND fecbaja >='" & Format(Fini, FormatoFecha) & "'"
    SQL = SQL & " AND fecbaja <='" & Format(FFin, FormatoFecha) & "'"
    SQL = SQL & strControlNomina
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "INSERT INTO tmpDatosMes(Mes,Trabajador,MesHoras,MesDias,idHorarioTra) VALUES (" & Month(Fini) & ","
    While Not RS.EOF
        Horas = CalculaHorasHorario(idCal, RS!idCal, d, Fini, RS!FecBaja)
        Aux = RS.Fields!idTrabajador & "," & TransformaComasPuntos(CStr(Horas)) & "," & d & "," & RS!idCal & ")"
        conn.Execute SQL & Aux
        RS.MoveNext
    Wend
    RS.Close
    
 
    'Aquellos que entran de baja enfermedad durante este mes
    'Cambio 3 Diciembre
        '-----------------
        'Calcularemos los dias que tenia que haber trabajado,
        'no los que le faltban para completar el mes y leugo restar
    Aux = ""
    If True Then  'por empresas?
        SQL = "Select bajas.*,trabajadores.idcal from bajas,trabajadores where idtrab=idTrabajador"
        SQL = SQL & strControlNomina
        SQL = SQL & " AND fechabaja >='" & Format(Fini, FormatoFecha) & "'"
        SQL = SQL & " AND fechabaja <='" & Format(FFin, FormatoFecha) & "'"
        SQL = SQL & " ORDER BY idtrabajador"
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        d = 0
        While Not RS.EOF
            If d <> RS!idTrab Then
                d = RS!idTrab
                Aux = Aux & d & "|"
            End If
            
            RS.MoveNext
        Wend
        RS.Close
    End If
    'Ya tengo los trabajadores. Ahora ire uno a uno por si han tenido mas dias de baja y eso
    While Aux <> ""
        
        d = InStr(1, Aux, "|")
        If d = 0 Then
            Aux = ""
        Else
            SQL = "Select bajas.*,trabajadores.idcal,trabajadores.fecalta  from bajas,trabajadores where idtrab=idTrabajador"
            SQL = SQL & strControlNomina
            SQL = SQL & " AND fechabaja >='" & Format(Fini, FormatoFecha) & "'"
            SQL = SQL & " AND fechabaja <='" & Format(FFin, FormatoFecha) & "' AND idtrab = "
            SQL = SQL & Mid(Aux, 1, d - 1)
            SQL = SQL & " ORDER BY fechabaja"
            
            
            
            
            
            
            Aux = Mid(Aux, d + 1)
            RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            FAux = Fini
            If Not IsNull(RS!FecAlta) Then
                If RS!FecAlta > Fini Then FAux = RS!FecAlta
            End If
            D22 = 0
            h22 = 0
            HorasBaja = 0
            UltMierTrabajado = -1
            varHorario = -1
            
            While Not RS.EOF
                If varHorario < 0 Then
                    SQL = DevuelveDesdeBD("idhorario", "calendariol", "idcal", CStr(RS!idCal))
                    If SQL = "" Then SQL = "1"
                    varHorario = SQL
                End If
                IDH = varHorario
                IDT = RS!idTrab
                'Tramo anterior a la baja
                If RS!FechaBaja > FAux Then
                    FAux2 = DateAdd("d", -1, RS!FechaBaja)
                    
                    
                    
                    Horas = CalculaHorasHorarioBaja(idCal, IDH, d, FAux, FAux2)
                   
                    h22 = h22 + Horas
                    D22 = D22 + d
                End If
                
                
                
                
                FAux2 = FFin
                If Not IsNull(RS!fechaalta) Then
                    If RS!fechaalta < FFin Then FAux2 = RS!fechaalta
                End If
                
                
                'Vemos si trabajao el dia de la baja
                Set vM = New CMarcajes
                                                            'Trabajo el dia de la baja
                If vM.Leer2(CLng(IDT), RS!FechaBaja) = 0 Then HorasBaja = HorasBaja + vM.HorasIncid
                Set vM = Nothing
                    
                
                'Insertamos en temporal de bajas para comprobar luego quien ha estado
                SQL = "INSERT INTO tmpCombinada(idTrabajador,Fecha,H1) VALUES (" & IDT
                SQL = SQL & ",'" & Format(RS!FechaBaja, FormatoFecha) & "','" & Format(FAux2, FormatoFecha) & "')"
                conn.Execute SQL
                        
                'Pongo la fecha aux a la fecha baja
                FAux = DateAdd("d", 1, FAux2)
                RS.MoveNext
                
                'Tiene mas de uno. Para pruebas
                'If Not RS.EOF Then
            Wend
            RS.Close
            
            If FAux2 < FFin Then
                'Significa que aun trabaja algo a final del mes
                FAux = DateAdd("d", 1, FAux2)
                FAux2 = FFin
                Horas = CalculaHorasHorarioBaja(idCal, IDH, d, FAux, FAux2)
                'Horas = CalculaHorasHorario(IDH, D, FAux, FAux2, False)   CalculaHorasHorarioBaja
                h22 = h22 + Horas
                D22 = D22 + d
            End If

            SQL = "UPDATE tmpDatosMes SET meshoras= " & TransformaComasPuntos(CStr(h22))
            SQL = SQL & " , mesdias = " & D22
            'If horas de baja >0 siginifica que trabajo. Luego tiene que tener +hc y -hn
            If HorasBaja > 0 Then
                SQL = SQL & " , HorasN = horasN - " & TransformaComasPuntos(CStr(HorasBaja))
                SQL = SQL & " , HorasC = horasC - " & TransformaComasPuntos(CStr(HorasBaja))
                
            End If
            SQL = SQL & " WHERE mes= " & Month(Fini) & " AND Trabajador =" & IDT
            conn.Execute SQL
            'Debug.Print "Tra      " & IDT & "     " & D22 & "    Horas " & h22
        End If
    Wend
    
    'LA FECHA DE ALTA NOOOOOOOO se trabaja
    'Cmprobar el procedimiento de baja
    'Aquellos que entraron de baja en dias anteriores al mes
    'Y se dieron de alta en el mes de calculo
    Aux = ""
    If True Then  'MiEmpresa.QueEmpresa = 0
            SQL = "Select bajas.*,trabajadores.idcal,fechaalta as altaTrabajador from bajas,trabajadores where idtrab=idTrabajador"
            SQL = SQL & strControlNomina
            SQL = SQL & " AND fechabaja <'" & Format(Fini, FormatoFecha) & "'"
            SQL = SQL & " AND fechaalta >='" & Format(Fini, FormatoFecha) & "'"
            SQL = SQL & " AND fechaalta <='" & Format(FFin, FormatoFecha) & "'"
            RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            d = 0
            While Not RS.EOF
                If d <> RS!idTrab Then
                    d = RS!idTrab
                    Aux = Aux & d & "|"
                End If
                
                RS.MoveNext
            Wend
            RS.Close
    End If
    
    'Ya tengo los trabajadores. Ahora ire uno a uno por si han tenido mas dias de baja y eso
    While Aux <> ""
        
        d = InStr(1, Aux, "|")
        If d = 0 Then
            Aux = ""
        Else
    
            'Empieza a trabajar este mes despues de una baja. No hacemos nada
            SQL = "Select bajas.*,idcal,FecAlta as altaTrabajador from bajas,trabajadores where idtrab=idTrabajador"
            SQL = SQL & strControlNomina
            SQL = SQL & " AND fechabaja <'" & Format(Fini, FormatoFecha) & "'"
            SQL = SQL & " AND fechaalta >='" & Format(Fini, FormatoFecha) & "'"
            SQL = SQL & " AND fechaalta <='" & Format(FFin, FormatoFecha) & "'"
            SQL = SQL & " AND idtrab = " & Mid(Aux, 1, d - 1)
            RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            
            
            'FALTA###
            IDH = 1 'RS!IdHorario
            IDT = RS!idTrab
            
            
             'If IDT = 31 Then MsgBox " -----"   '
            
            
            Aux = Mid(Aux, d + 1)
            'Si se ha dado de
            FAux = Fini
            If RS!altaTrabajador > FAux Then FAux = RS!altaTrabajador
            FAux2 = RS!fechaalta
            
            RS.Close
            
            SQL = DBSet(FAux, "F")
            SQL = DevuelveDesdeBD("idhorario", "calendariol", "idcal =" & idCal & " AND fecha", SQL)
            If SQL = "" Then SQL = "1"
            IDH = CInt(SQL)

            'Nuevo
            If FAux < FAux2 Then
            
                
                Horas = CalculaHorasHorario(idCal, IDH, d, FAux, FAux2)
            
    
                SQL = "INSERT INTO tmpCombinada(idTrabajador,Fecha,H1) VALUES (" & IDT
                SQL = SQL & ",'" & Format(FAux, FormatoFecha) & "','" & Format(FAux2, FormatoFecha) & "')"
                Ejecuta SQL
                
    
                SQL = "UPDATE tmpDatosMes SET meshoras= meshoras - " & TransformaComasPuntos(CStr(Horas))
                SQL = SQL & " , mesdias =mesdias - " & d
                SQL = SQL & " WHERE mes= " & Month(Fini) & " AND Trabajador =" & IDT
                conn.Execute SQL
            End If
        End If
      
            
    Wend
    
    
    
    
    
    'A titulo informtivo pondremos aquellos trabajadores
    'que estan de baja todavia. Es decir la fecha de alta es menor
    'ALTA temporada<inicio
    'baja temporada o null o >ffin
    'En bajas esta con la fecha de alta a null y fecha baja < finicio
    If ControlNomina = 0 Then
    
        'PARA QUE APAREZCAN LAS BAJAS EN EL MOMENTO
    
'        SQL = "SELECT Bajas.idTrab"
'        SQL = SQL & " FROM Trabajadores INNER JOIN Bajas ON Trabajadores.IdTrabajador = Bajas.idTrab"
'        SQL = SQL & " WHERE Bajas.FechaAlta Is Null AND Trabajadores.FecAlta<#" & Format(Fini, FormatoFecha) & "# AND"
'        SQL = SQL & " (Trabajadores.FecBaja Is Null  OR Trabajadores.Fecbaja>#" & Format(FFin, FormatoFecha) & "#) AND"
'        SQL = SQL & " (Bajas.Fechabaja Is Null  OR Bajas.Fechabaja<#" & Format(Fini, FormatoFecha) & "#)"
'
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        SQL = "INSERT INTO tmpDatosMes(Mes,Trabajador,MesHoras,MesDias) VALUES (" & Month(Fini) & ","
'        While Not RS.EOF
'
'            aux = RS.Fields!idTrab & ",0,0)"
'            'Conn.Execute SQL & Aux
'            RS.MoveNext
'        Wend
'        RS.Close
    End If
    
    Set RS = Nothing
End Sub

Private Sub Ejecuta(SQL As String)
    On Error Resume Next
    conn.Execute SQL
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description & vbCrLf & "El proceso continua"
End Sub

'Cojeremos y uniremos en tmpDatosMes todos los datos relativos a los trabajadores , para
'el periodo procesado anteriormente
Public Sub CombinaDatos(Fini As Date, FFin As Date)
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim i As Integer
Dim Tot As Currency
'Dim Importe As Currency
Dim Aux As String
Dim SQL As String
Dim Rs2 As ADODB.Recordset
Dim meshoras As Currency

    Set RS = New ADODB.Recordset
    SQL = "Select Trabajador,MEsDias,meshoras from tmpDatosMes "
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    SQL = "SELECT tmpHoras.trabajador, tmpHoras.HorasT, tmpHoras.HorasC, tmpHoras.HorasE, tmpHoras.Dias, trabajadoresbolsahoras.HorasBolsa"
    SQL = SQL & " FROM  tmpHoras left join trabajadoresbolsahoras ON trabajadoresbolsahoras.IdTrabajador = tmpHoras.trabajador and"
    SQL = SQL & " tipohora=1 " 'No deberia tener bolsa horas extra
    SQL = SQL & " WHERE tmpHoras.trabajador = "

    Set RT = New ADODB.Recordset
    While Not RS.EOF
       

        
        RT.Open SQL & RS.Fields(0), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            Aux = "UPDATE tmpDatosMes Set HorasN=" & TransformaComasPuntos(CStr(RT!horast))
            Aux = Aux & " ,HorasC=" & TransformaComasPuntos(CStr(RT!HorasC))  'Las compensables SON las extras en COOPIC
            Aux = Aux & " ,HorasE=" & TransformaComasPuntos(CStr(RT!horase))
            'Tot = RT!horasc + RT!horast  extras mas normales
            Tot = RT!horast
            Aux = Aux & " ,HorasT=" & TransformaComasPuntos(CStr(Tot))
            i = RS!MesDias - RT!Dias
            meshoras = RS!meshoras
            If i < 0 Then
                'EN CATADAU dejo pasar.
                'NO compensan dias desde bolsa, con lo cual, cualquier dia que haya venido entra
                
                If vEmpresa.QueEmpresa = 4 Then
                    i = RT!Dias
                Else
                    i = RS!MesDias
                End If
            Else
               
               
               If vEmpresa.QueEmpresa = 4 Then If i > 0 Then meshoras = RT!Dias * 8
                                    
                 i = RT!Dias
            End If
            
            Aux = Aux & " ,DiasTrabajados=" & i
            Aux = Aux & " ,meshoras=" & TransformaComasPuntos(CStr(meshoras))
            Aux = Aux & " ,BolsaAntes =" & TransformaComasPuntos(CStr(DBLet(RT!horasbolsa, "N")))
            Aux = Aux & " ,Anticipos = " & "0"    ' & TransformaComasPuntos(CStr(Tot))
            Aux = Aux & " WHERE trabajador = " & RS.Fields(0)
            conn.Execute Aux
        Else
            'MIRARE SI TIENE BOLSA DE HORAS. Con lo cual puede que no haya trabajado NINGUN dia
            'pero si tenia bolsa le seguiremos generando dias
            RT.Close
            Aux = "SELECT trabajadoresbolsahoras.HorasBolsa"
            Aux = Aux & " FROM  trabajadoresbolsahoras WHERE "
            Aux = Aux & " tipohora=1 " 'No deberia tener bolsa horas extra
            Aux = Aux & " AND  idtrabajador = " & RS.Fields(0)
        
            
            RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RT.EOF Then
                If DBLet(RT!horasbolsa, "N") > 0 Then
                    Aux = "UPDATE tmpDatosMes SET "
                    Aux = Aux & " BolsaAntes =" & TransformaComasPuntos(CStr(DBLet(RT!horasbolsa, "N")))
                    Aux = Aux & " WHERE trabajador = " & RS.Fields(0)
                    conn.Execute Aux
                End If
            End If
        End If
        RT.Close
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    Set RS = Nothing
End Sub




'Total horas y total dias
Public Sub CalculoDatosACompensar()
Dim RS As ADODB.Recordset
Dim i As Integer
Dim SQL As String
Dim Diferencia As Currency
    SQL = "Select * from tmpDatosMes"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not RS.EOF
    
        

          
        Diferencia = RS!horast - RS!meshoras
        If Diferencia < 0 Then
            'Veremos si coge horas de la bolsa o no
            '
            
        End If
        
        RS!saldoh = Diferencia
        
        
        
        
        RS!saldodias = RS!MesDias - RS!diasTrabajados
        RS.Update
        
        'sgi
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
End Sub


Public Sub HacerCompensaciones(FInicio As Date, FFin As Date, lbl As Label)
Dim HCompMes As Currency
Dim HPaBolsa As Currency
Dim DiasOF As Integer
Dim HorasOf As Currency
Dim H As Currency
Dim h2 As Currency
Dim SQL As String
Dim ModoCompensacion As String
Dim HorasJornadaRecuperacion As Currency
Dim Horario As Integer
Dim vH As CHorarios
Dim FESTIVOS As String
Dim MEDIODIA As String
Dim DiasReajusteXSTrabajados As Integer
Dim BajaTodoElMes As Boolean
Dim Canti As Currency
Dim ImporAux As Currency
Dim LlevaPlus2 As Currency
                      
                      'El brutoN sera la suma de las horasN * importe
                      'LlevaPlus2=Hest * PlusHoraTrabajador + HNormales*PlusHoraTrabajador + HExtr*PlusHoraTrabajador
                      'PlusNormal=extra * IMporteDehoraExtra
                      
                      

Dim BrutoN As Currency
Dim ImportePlus As Currency
Dim HazUpdate As Boolean
    Dim RS As ADODB.Recordset
Dim HorasC As Currency  'Catadau, las estructuradas para los de importe fijo, pueden cambiar
Dim ImporEstrcut As Currency

Dim H_D_mas As Currency
    'ModoCompensacion
    'Vemos cual es el modo de compensacion
    '   0 .- NO compensa
    '   1 .- A partir de los dias trabajados del trabajador
    '         vemos cuantos dias le puedo compensar
    '   2 .- X horas hacen una jornada laboral a compensar
    '   3 .- Picassen cotubre 2008.
    '           -Compensaran por semana /dia con cuidado a los miercoles sabados
    '           -si trabaja una hora un dia, el resto de horas NO las tiene que compensar para la nomina
    SQL = "HorasJornada"
    ModoCompensacion = DevuelveDesdeBD("RecuperacionDias", "Empresas", "idEmpresa", "1", "N", SQL)
    

    If Val(ModoCompensacion) = 0 Then
        ModoCompensacion = "0"
        HorasJornadaRecuperacion = 0
    Else
        HorasJornadaRecuperacion = CCur(SQL)
    End If
    
    

    'De momento NO lo necesito
    If ModoCompensacion = "3" Then
    
        'Fijo cual es ñla primera semana del mes, y la utima
        SemanaMesPrimera = Format(FInicio, "ww", vbMonday)
        SemanaMesUltima = Format(FFin, "ww", vbMonday)
    
        'Ajustes ponemos HN las que tiene menos las que sean extra
        lbl.Caption = "Ajuste horas normales"
        AjustarHorasNormales
        
        'Utlizaremos una tabla mas para guardar lios dias que XyS no deberean ser contabilizados como tal en nomina
        conn.Execute "DELETE FROM tmpNOTrabajo"
        conn.Execute "DELETE FROM tmpPagosMes"
        
        
        'VEmos el miercoles
        RecalculoHorasMiercolesSabados FInicio, FFin, lbl, True
        'Sabado
        RecalculoHorasMiercolesSabados FInicio, FFin, lbl, False

        'Vemos cuantos miercoles /sabado han trabajado pero no deben entrar en nomina
        lbl.Caption = "Procesar datos"
        lbl.Refresh
    End If

  
    SQL = "Select tmpDatosMes.*,idHorarioTra idHorario,FecAlta,FecBaja,controlnomina,NomTrabajador,ImporteFijoNomina from tmpDatosMes,Trabajadores"
    SQL = SQL & " WHERE tmpDatosMes.trabajador = Trabajadores.idTrabajador"
    SQL = SQL & " ORDER BY idHorario"
    Horario = -1
    FESTIVOS = ""
    Set RS = New ADODB.Recordset
    
   
    
    RS.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        
        If ModoCompensacion = "1" Or ModoCompensacion = "3" Then
            If Horario <> RS!IdHorario Then
                Set vH = Nothing
                Set vH = New CHorarios
                Horario = RS!IdHorario
                FESTIVOS = vH.LeerDiasFestivos(Horario, FInicio, FFin)
                MEDIODIA = vH.LeerMediosDias(Horario, FInicio, FFin)
            End If
        End If
            
        HazUpdate = True
        DiasReajusteXSTrabajados = 0
        HorasC = -1
        
        'If RS!Trabajador = 20208 Then S top

        
        
        If vEmpresa.RecuperacionDias = 2 Then
            'Solo se crecuperan dias. Las horas trabjadas van como van
            HCompMes = 0
            HorasOf = RS!horasn
            HPaBolsa = RS!bolsaantes
            DiasOF = RS!diasTrabajados
            H = RS!MesDias - RS!diasTrabajados
            If H < 0 Then
                MsgBox RS!nomtrabajador & "    Dias " & RS!MesDias & "   Traba: " & RS!diasTrabajados, vbExclamation
            Else
            
                If H > 0 Then
                           
                    'Modificacion NOVIEMBRE 2018
                    'Desdoblamos CATADAU y el resto
                    If vEmpresa.QueEmpresa = 4 Then
                            'Puede compensar dias
                            If HPaBolsa >= HorasJornadaRecuperacion Then
                                
                                H = CCur(HPaBolsa)
                                DiasOF = CuantosDiasCompensas(RS!saldodias, H, HorasJornadaRecuperacion)
                                H = DiasOF * HorasJornadaRecuperacion
                                
                                
                                
                                
                                DiasOF = RS!diasTrabajados + DiasOF
                                HPaBolsa = HPaBolsa - H
                            End If
                    
                    
                             'Abril 2018
                             'Si los nuevos dias, generaran NUEVAS horasmes. Habra que ver las que se pasen
                             h2 = 8 * DiasOF
                             h2 = RS!horasn - h2
                             If h2 > 0 Then
                                 SQL = "Trabajador: " & RS!Trabajador & vbCrLf
                                 SQL = SQL & "Total dias/horas mes: " & DiasOF & "/" & 8 * DiasOF & vbCrLf
                                 SQL = SQL & "Horas trabajadas: " & RS!horasn & "   a bolsa: " & h2 & vbCrLf
                                 'SQL = SQL & "Utlizadas dias copensados: " & H
                                 'MsgBox SQL, vbExclamation
                                 Debug.Print "Comepnsa " & Replace(SQL, vbCrLf, "    ")
                                 
                                'En catadau Cambiamos las horas ficiales
                                 HorasOf = 8 * DiasOF
                                
                                 HPaBolsa = HPaBolsa + h2
                             End If
                        
                        
                    Else
                        'RESTO. COOPIC
                        H_D_mas = 8 * DiasOF
                        H_D_mas = RS!horasn - H_D_mas
                        If H_D_mas > 0 Then
                            Debug.Print "Comepnsa " & RS!Trabajador & "   " & H_D_mas & "  +  " & HPaBolsa
                            HPaBolsa = HPaBolsa + H_D_mas
                            
                        End If
                        If HPaBolsa >= HorasJornadaRecuperacion Then
                                
                            H = CCur(HPaBolsa)
                            DiasOF = CuantosDiasCompensas(RS!saldodias, H, HorasJornadaRecuperacion)
                            H = DiasOF * HorasJornadaRecuperacion
                            
                            
                            
                            
                            DiasOF = RS!diasTrabajados + DiasOF
                            HPaBolsa = HPaBolsa - H
                        End If
                        
                        
                    End If
                Else
                    'Veremos si ha trabjado mas horas de las que deberia
                    
                    H = RS!horasn - RS!meshoras
                    If H > 0 Then
                        'Ha trabajado mas horas de las que debe
                        HorasOf = RS!meshoras
                        HPaBolsa = HPaBolsa + H
                    Else
                        ' '
                        
                    End If
                End If
                
                
                                
                
                
            End If
        
            
        
        Else
        
            If ModoCompensacion <> "0" Then
                'Me debe dias trabajados
                'Tengo k ver si en las horas que tiene tiene suficiente
                'Para esos dias trabajados. Si no no le compenso los dias
         
            
                
                
                H = RS!HorasC + RS!bolsaantes
                If (RS!horasn + H) >= RS!meshoras Then
                
                    
                
                    If ModoCompensacion <> "3" Then
                            'Tiene bastantes horas para compensar el mes entero
                            DiasOF = RS!MesDias
                            HorasOf = RS!meshoras
                            
                            HCompMes = RS!meshoras - RS!horasn
                            If RS!HorasC >= HCompMes Then
                                'Las coje todas de las compensadas de este mes
                                H = RS!HorasC - HCompMes
                                HPaBolsa = RS!bolsaantes + H
                        
                            Else
                                H = HCompMes - RS!HorasC
                                'Necesito h horas de la bolsa
                                HPaBolsa = RS!bolsaantes - H
                            End If
                    Else
                        'Picassent 2008
                        
                        
                        HCompMes = H
                        DiasOF = CompensacionesDiaTrabajadoYSemana(RS!saldodias, RS, FESTIVOS, FInicio, FFin, vH, HorasJornadaRecuperacion, HCompMes, DiasReajusteXSTrabajados, BajaTodoElMes)
                        
                        
                        If DiasOF < RS!saldodias Then
                                'Le faltan dias por compensar
                                If H - HCompMes > HorasJornadaRecuperacion Then
                                    DiasOF = DiasOF + 1
                                    If H - HCompMes > 8 Then
                                        HCompMes = HCompMes + 8
                                    Else
                                        HCompMes = H
                                    End If
                                End If
                         End If
                        
                        
                        
                        
                        HPaBolsa = H - HCompMes
                        HorasOf = RS!horasn
                        DiasOF = RS!diasTrabajados + DiasOF
                        
                    End If
                    
                Else
                    
                    'Si no tengo bastante le dejo la bolsa como esta
                    'Y le pongo los dias k ha hecho, sin modificar
                    
                    If H = 0 Then   'Si NO tiene nada a compensar
                        DiasOF = RS!diasTrabajados
                        HorasOf = RS!horast
                        HCompMes = 0   'Este mes no le quedara nada para compensar
                        HPaBolsa = 0
                    Else
                    
                   
                       HPaBolsa = 0
                       'En funcion del tipo de compensacion
                       Select Case Val(ModoCompensacion)
                       Case 1, 2
                                         
                            HorasOf = RS!horasn + H
                        
                            'Vemos esas h -horas cuantos dias me puede compensar
                            If ModoCompensacion = 2 Then
                                DiasOF = CuantosDiasCompensas(RS!saldodias, H, HorasJornadaRecuperacion)
                      
                            Else
                                'EN HorasJornadaRecuperacion: tengo el minimo de horas para que le compensen a una persona un dia sin llegar a las 8 horas
                                DiasOF = CompensacionesDiaTrabajado(RS!saldodias, H, RS, FESTIVOS, FInicio, FFin, vH, HorasJornadaRecuperacion)
                               
                            End If
                      
                            DiasOF = RS!diasTrabajados + DiasOF
                            HCompMes = H
                       Case 3
                            'Nueva forma de compensar en PICASSENT. Oct 2008
                            
                            'EN HorasJornadaRecuperacion: tengo el minimo de horas para que le compensen a una persona un dia sin llegar a las 8 horas
                            HCompMes = H
                            DiasOF = CompensacionesDiaTrabajadoYSemana(RS!saldodias, RS, FESTIVOS, FInicio, FFin, vH, HorasJornadaRecuperacion, HCompMes, DiasReajusteXSTrabajados, BajaTodoElMes)
                            
                            If Not BajaTodoElMes Then
                                If DiasOF < RS!saldodias Then
                                    'Le faltan dias por compensar
                                    If H - HCompMes >= HorasJornadaRecuperacion Then
                                        DiasOF = DiasOF + 1
                                        If H - HCompMes > 8 Then
                                            HCompMes = HCompMes + 8
                                        Else
                                            HCompMes = H
                                        End If
                                    End If
                                End If
                                
                                DiasOF = RS!diasTrabajados + DiasOF
                                                        
                            
                                'Aui esta la difencia
                                'Las horas que no se utilizan, no se compensan
                                HPaBolsa = H - HCompMes
                                'HorasOf = RS!horast - RS!horasc
                                'Ahora
                                HorasOf = RS!horasn
                            Else
                                'Todo el mes de baja
                                DiasOF = 0
                                HorasOf = 0
                                HPaBolsa = H
                                H = 0
                                
                            End If
                       Case Else
                            'NOOOOOO compensamos nada
                            DiasOF = RS!diasTrabajados
                            HPaBolsa = 0
                            HorasOf = RS!horast
                            HCompMes = 0
                       End Select
                   
                    End If
                    
                End If
            Else
                    HazUpdate = False
                        'Aqui creo que no entra
                        If vEmpresa.QueEmpresa <> 4 Then
                            MsgBox "Error gen true 2!"
              
                            HCompMes = 0
                            
                            'Aui esta la difencia
                            'Las horas que no se utilizan, no se compensan
                            HPaBolsa = H - HCompMes
                            HorasOf = RS!horast - RS!HorasC
                        Else
                            'Si tiene importe fiho,
                            If DBLet(RS!ImporteFijoNomina, "N") > 0 Then
                                HPaBolsa = 0
                                DiasOF = RS!MesDias
                                
                                HCompMes = RS!horase
                                If RS!horasn + RS!HorasC >= RS!meshoras Then
                                    
                                    HorasC = RS!meshoras - RS!horasn
                                    HorasC = RS!HorasC - HorasC
                                    HorasOf = RS!meshoras
                                Else
                                    HorasC = 0
                                    HorasOf = RS!horasn + RS!HorasC
                                End If
                                HazUpdate = True
                                
                            Else
                                'Resto de trabajadores dejamos los valores como estaban
                                
                                HPaBolsa = DBLet(RS!bolsadespues, "N")
                                HorasOf = DBLet(RS!HorasPeriodo, "N")
                                
                                HCompMes = DBLet(RS!horase, "N")
                                DiasOF = RS!diasTrabajados
                                HorasOf = RS!horast
                                
                                HazUpdate = True
                            End If
                        End If
                        
                
            End If   'Diastrabajados no es igual k los k debia hber trabajado
        End If 'de compensa solo solo dias

        'Updateamos con los valores calculados
        
        'May18 ModoCompensacion_2 <> "0"
        If HazUpdate Then
             SQL = "UPDATE tmpDatosMes SET"
            
             
             SQL = SQL & "  BolsaDespues =" & TransformaComasPuntos(CStr(HPaBolsa))
             SQL = SQL & ", HorasPeriodo = " & TransformaComasPuntos(CStr(HorasOf))
             
             If DiasReajusteXSTrabajados > 0 Then
                 'Debug.Print "Dias reajus trabajador: " & RS!Trabajador
                 SQL = SQL & ", DiasTrabajados  = DiasTrabajados - " & DiasReajusteXSTrabajados
                 'El saldo de dias se incrementa
                 SQL = SQL & ", SaldoDias  = SaldoDias + " & DiasReajusteXSTrabajados
                 DiasOF = DiasOF - DiasReajusteXSTrabajados
                 SQL = SQL & ", DiasPeriodo  = " & TransformaComasPuntos(CStr(DiasOF))
             Else
                 SQL = SQL & ", DiasPeriodo  = " & TransformaComasPuntos(CStr(DiasOF))
             End If
             'Para PICASSENT, machaco los datos
             SQL = SQL & ", HorasN = " & TransformaComasPuntos(CStr(HorasOf))
             
             
             
             'Las horas extras
             If HCompMes < 0 Then HCompMes = 0
             SQL = SQL & ", Extras = " & TransformaComasPuntos(CStr(HCompMes))
             
             
             'Si es horasc>-1 sigmifica que ha entrado en el sistema de nominas catadau
             If HorasC > -1 Then
                SQL = SQL & ", horasc = " & TransformaComasPuntos(CStr(HorasC))
             End If
             
             'Trabajador
             SQL = SQL & " WHERE Trabajador = " & RS!Trabajador
             conn.Execute SQL
             
        
        End If
        
        
        If BajaTodoElMes Then
            AjustaDatosBajaMesEnteroTrabajador CLng(RS!Trabajador)
        End If
        
        
        
        
        'sgi
        RS.MoveNext
    Wend
    RS.Close
    espera 0.5
    
    
    'FALTA###
    If vUsu.Login = "root" Then
        
        'MsgBox "Proceso en funcinamiento. Falta ver si son todas las horas o solo las de nomina", vbExclamation
    End If

    
    
    'AHora obtenemos los anticpos en NOMINA
    '-----------------------------------------
    SQL = "SELECT Trabajador,tmpDatosMEs.HorasN, tmpDatosMEs.extras,horasc, Categorias.Importe1, Categorias.Importe2,Importe3, Trabajadores.PorcSS, Trabajadores.PorcIRPF,Trabajadores.PorcAntiguedad,saldoH,tmpDatosMEs.HorasPlus"
    SQL = SQL & ",PlusHN , PlusHE , ImporteFijoNomina"
    SQL = SQL & " FROM tmpDatosMEs INNER JOIN (Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria) ON tmpDatosMEs.Trabajador = Trabajadores.IdTrabajador"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
    
  
    
        'Se pagan TODAS las horas
        ' Las que vayan a bolsa, se pagan tambien
        ' En saldoH tenemos las que iran a bolsa. Siempre que sean positivas
        LlevaPlus2 = 0
        ImportePlus = 0
        ImporEstrcut = 0
        HorasOf = 0
        If vEmpresa.QueEmpresa = 4 Then
            'CATADAU NO compensa horas, todas se pagan en su franja.
            'ADemas CATADAU lleva plus
            
            
            
            'NORMALES
               
            'If RS!horasn > 0 Then St op
            'If RS!Trabajador = 199 Then S top
            
            Canti = RS!Importe1
            BrutoN = Round(RS!horasn * Canti, 2)
            
            'El plus en normales ya va a la columna de plus
            If DBLet(RS!PlusHN, "N") > 0 Then
                LlevaPlus2 = Round(RS!horasn * RS!PlusHN, 2)
            End If
            
            If Not IsNull(RS!ImporteFijoNomina) Then BrutoN = RS!ImporteFijoNomina
                       
            
            'Estrcuturales
            Canti = RS!Importe2
            If DBLet(RS!PlusHN, "N") > 0 Then
                LlevaPlus2 = LlevaPlus2 + Round(RS!HorasC * RS!PlusHN, 2)
            End If
            ImporEstrcut = (RS!HorasC * Canti)
            
            H = 0
            If vEmpresa.AplicaAntiguedadHC Then
                If DBLet(RS!PorcAntiguedad, "N") > 0 Then H = RS!PorcAntiguedad
            End If
            H = Round((ImporEstrcut * H) / 100, 2)
           
            HorasOf = ImporEstrcut + BrutoN - H
            
            
            'EXTRAS
            Canti = RS!Importe3
            If DBLet(RS!PlusHe, "N") > 0 Then
                LlevaPlus2 = LlevaPlus2 + Round(RS!extras * RS!PlusHe, 2)
            End If
            ImporAux = (RS!extras * Canti)
            H = 0
            If vEmpresa.AplicaAntiguedadHC Then
                If DBLet(RS!PorcAntiguedad, "N") > 0 Then H = RS!PorcAntiguedad
            End If
            H = Round((ImporAux * H) / 100, 2)
            ImportePlus = ImportePlus + ImporAux
            HorasOf = ImporAux + HorasOf - H + LlevaPlus2


        Else
            'Resto
            H = RS!saldoh
            If H < 0 Then H = 0
            HorasOf = (RS!horasn * RS!Importe1) + (H * RS!Importe2) + (RS!HorasC * RS!Importe3)
            
            H = 0
            If DBLet(RS!PorcAntiguedad, "N") > 0 Then H = RS!PorcAntiguedad
            H = Round((HorasOf * H) / 100, 2)
            HorasOf = HorasOf + H

            BrutoN = HorasOf
        End If
        
        
        

        
        'Quitamos IRPF y SS
        H = (HorasOf * RS!PorcIRPF) + (HorasOf * RS!PorcSS)
        H = Round((H / 100), 2)
        HorasOf = HorasOf - H
        H = (ImportePlus * RS!PorcIRPF) + (ImportePlus * RS!PorcSS)
        H = Round((H / 100), 2)
        HorasOf = HorasOf - H + ImporEstrcut + ImportePlus
        
        SQL = "UPDATE tmpDatosMes SET"
        SQL = SQL & " Anticipos = " & TransformaComasPuntos(CStr(HorasOf))
        
        SQL = SQL & ", Bruto = " & TransformaComasPuntos(CStr(BrutoN))
        SQL = SQL & ", Extras = " & TransformaComasPuntos(CStr(ImportePlus))
        SQL = SQL & ", LlevaPlus = " & TransformaComasPuntos(CStr(LlevaPlus2))
        SQL = SQL & ", ImporEstruc = " & TransformaComasPuntos(CStr(ImporEstrcut))
        If Not IsNull(RS!ImporteFijoNomina) Then SQL = SQL & ",  ImporteFijo =1 "
        
        'Trabajador
        SQL = SQL & " WHERE Trabajador = " & RS!Trabajador
    
        conn.Execute SQL
    
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    
    
    
    Set RS = Nothing
    
    

End Sub


Private Function MiercolesSabadoNoCuentaTrabajado(T As String, F As Date) As Boolean
Dim RT As ADODB.Recordset
Dim C As String

    C = "Select * from tmpNoTrabajo where idtra=" & T & " AND idFech='" & Format(F, FormatoFecha) & "'"
    Set RT = New ADODB.Recordset
    RT.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    MiercolesSabadoNoCuentaTrabajado = RT.EOF
    RT.Close
    Set RT = Nothing
        
End Function








Private Function CuantosDiasCompensas(Dias As Integer, HorasCompensar As Currency, HorasJornadaCompensable As Currency) As Integer
Dim J As Currency
Dim i As Integer
    'Compensamos dias a partir de HorasJornada  horas trabajadas
    J = (HorasCompensar / HorasJornadaCompensable)
    i = Int(J)
    If i > Dias Then i = Dias
    CuantosDiasCompensas = i
End Function



Private Function CompensacionesDiaTrabajado(Dias As Integer, HorasCompensar As Currency, ByRef Rec As Recordset, ByRef FEST As String, ByVal FI As Date, ByVal FF As Date, ByRef vHO As CHorarios, HorasMinimoDia As Currency) As Integer
Dim RF As ADODB.Recordset
Dim cad As String
Dim Fin  As Boolean
Dim Horas As Currency
Dim Sig As Boolean
Dim DiaC As Currency
Dim FechaReferencia As Date

On Error GoTo ECompensacionesDiaTrabajado

    CompensacionesDiaTrabajado = 0
    'Si fecha alta > fecha inicio mes enonces finicio mes=fecha alta
    If Rec!FecAlta > FI Then FI = Rec!FecAlta

    'Si fecha baja < fecha baja mes entonces finicio mes=fecha alta
    If Not IsNull(Rec!FecBaja) Then
        If Rec!FecBaja < FF Then FF = Rec!FecBaja
    End If

    cad = "Select distinct(fecha) from marcajes"
    cad = cad & " WHERE Fecha >='" & Format(FI, FormatoFecha) & "'"
    cad = cad & " AND Fecha <='" & Format(FF, FormatoFecha) & "'"
    cad = cad & " AND idTrabajador = " & Rec!Trabajador

    Set RF = New ADODB.Recordset
    RF.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Horas = 0
    DiaC = 0
    Fin = False
    If Not RF.EOF Then
        '
        FechaReferencia = RF!Fecha
        While Not Fin
            
            If FI > FF Then
                Fin = True
                Sig = False 'Para k no mueva el recordset
            Else
                If FI = FechaReferencia Then
                    FI = DateAdd("d", 1, FI)
                    Sig = True
                Else
                    If InStr(1, FEST, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                        'Es un dia festivo
                        FI = DateAdd("d", 1, FI)
                        Sig = False
                    Else
                        'Es un dia k no ha  trabajado. Vemos cuantas horas son
                        vHO.Leer vHO.IdHorario, FI, 1
                        'Ya tenog las horas k debia haber trabajado
                        If Horas + vHO.TotalHoras <= HorasCompensar Then
                            'Le puedo compensar este dia
                            DiaC = DiaC + vHO.DiaNomina
                            Horas = Horas + vHO.TotalHoras
                            
                        Else
                            'Este dia no se lo puedo compensar
                            'No hago nada
                        End If
                                     
                        FI = DateAdd("d", 1, FI)
                        Sig = False
                                     
                        'Por si acaso ya ha compensado todos los dias
                        If DiaC >= Dias Then
                            If DiaC > Dias Then DiaC = Dias
                            Fin = True
                        End If
                        
                        'Si no le kedan horas para compensar tampoco seguimos
                        If HorasCompensar - Horas < 3 Then Fin = True
                    End If
                End If
            
        
            End If
            If Sig Then
                If RF.EOF Then
                    'Deberiamos salir
                    '
                Else
                    RF.MoveNext
                    'ANTES
                    'If RF.EOF Then Fin = True
                    If Not RF.EOF Then FechaReferencia = RF!Fecha
                End If
            End If
        Wend
    Else
        'NO HA TRABAJADO, pero tiene Horas de otros meses
        DiaC = HorasCompensar \ 8    'Cuantos dias de 8 horas le entran
        Horas = HorasCompensar - (DiaC * 8) 'Horas sobrantes
        If Horas >= HorasMinimoDia Then DiaC = DiaC + 1   'Veo si el resto me comepnsa un dia o no
        If DiaC >= Dias Then DiaC = Dias                  'NO puede compensar mas dias de los que pueden ir en nomina

    End If
    RF.Close
    Set RF = Nothing
    If DiaC > Int(DiaC) Then
        DiaC = Int(DiaC) + 1
        If DiaC > Dias Then DiaC = Dias
    End If
        
    CompensacionesDiaTrabajado = DiaC
        
    Exit Function
ECompensacionesDiaTrabajado:
    MuestraError Err.Number, "CompensacionesDiaTrabajado"

End Function









Public Sub AjustaDatosBajaMesEntero()
Dim SQL As String
    SQL = "UPDATE tmpDatosMes SET "
    SQL = SQL & " MesHoras=0, Mesdias = 0, SaldoH=0, SaldoDias=0,HorasPeriodo =0, BolsaDespues=0, DiasPeriodo=0"
    SQL = SQL & " WHERE (((tmpDatosMEs.DiasTrabajados)=0) AND ((tmpDatosMEs.HorasN)=0) AND ((tmpDatosMEs.HorasC)=0) AND ((tmpDatosMEs.bolsaAntes)=0)) ;"
    conn.Execute SQL
End Sub

Public Sub AjustaDatosBajaMesEnteroTrabajador(T As Long)
Dim SQL As String
    SQL = "UPDATE tmpDatosMes SET "
    SQL = SQL & " MesHoras=0, Mesdias = 0, SaldoH=0, SaldoDias=0,HorasPeriodo =0"
    SQL = SQL & " WHERE tmpDatosMEs.trabajador= " & T
    conn.Execute SQL
End Sub



'Un trabajador, entre unas fechas, si ha trabajado
Public Function HaTrabajadoConBaja(ByRef R As ADODB.Recordset) As Boolean
Dim Rec As ADODB.Recordset
Dim SQL As String

    HaTrabajadoConBaja = False
    SQL = "Select * from Marcajes WHERE"
    SQL = SQL & " idTrabajador =" & R!idTrabajador
    SQL = SQL & " AND fecha >='" & Format(R!Fecha, FormatoFecha) & "'"
    'Ambos inclusive de baja
    SQL = SQL & " AND fecha <='" & Format(R!H1, FormatoFecha) & "'"
    Set Rec = New ADODB.Recordset
    Rec.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rec.EOF Then HaTrabajadoConBaja = True
    Rec.Close
    Set Rec = Nothing
        
End Function




'Calcula las horas trabajadas para los trabajadores k tiene la marca puesta
Public Sub CalculaHorasTrabajadasConEXTRAS(Fini As Date, FFin As Date, ControlNomina As Byte)
Dim FAux As Date
Dim FAux2 As Date
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Dias As Currency
Dim Trab As Long
Dim Aux As String
Dim SQL As String
Dim vH As CHorarios
Dim FESTIVOS As String
Dim MEDIODIA As String
Dim strControlNomina As String
Dim Horario As Integer
Dim HC As Currency
Dim HE As Currency

    'IMPORTANTE
    'Ahora hay un control nomina mas, k es el 2
    'El tipo de control 2: Tiene un suledo fijo al mes
    'Pero en anticpos solo anticipa hNormales
    'luego el calculo de horas es el mismo que el 1
    ' por lo tanto donde ponia
        'SQL = SQL & " AND Trabajadores.ControlNomina = 1"
    ' pondra ahora
        'SQL = SQL & " AND Trabajadores.ControlNomina > 0"


    'Otro MAS. El tipo 3
    '   40 Horas semanales. 5 dias semana
    '
    'Con lo cual si en
    ' controlnomina
        ' 1.-   NORMAL ControlNomina >0 and ControlNomina <3
        ' 2.- Solo para el tipo  3
    Select Case ControlNomina
    Case 0
        strControlNomina = " AND Trabajadores.ControlNomina >0  AND Trabajadores.ControlNomina <2 "
    Case 1
        strControlNomina = " AND Trabajadores.ControlNomina = 3"
    Case 2
        strControlNomina = " AND (Trabajadores.ControlNomina =1  OR Trabajadores.ControlNomina =3) "
    Case 3
        'Sera para el listado que se entraga a los trabbajdores en PICASSENT
        ' Es para los tipos 1,2,3
        strControlNomina = " AND Trabajadores.ControlNomina >0"
    Case Else
        strControlNomina = ""
    End Select
    


    conn.Execute "Delete from tmpHoras"
    
    'Calculamos las horas para el mes
    'Primero las normales con un simple insert into
    SQL = "INSERT INTO tmpHoras(trabajador,HorasT) "
    SQL = SQL & "SELECT Marcajes.idTrabajador, Sum(Marcajes.HorasTrabajadas) AS SumaDeHorasTrabajadas"
    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
    SQL = SQL & " Where Marcajes.Fecha >= '" & Format(Fini, FormatoFecha) & "'"
    SQL = SQL & " and Marcajes.Fecha <= '" & Format(FFin, FormatoFecha) & "'"
    
    SQL = SQL & strControlNomina
    SQL = SQL & " GROUP BY Marcajes.idTrabajador;"
    conn.Execute SQL
    
    
    
    '----HORAS COMPENSAR
    'Las horas para la bolsa de trabajor
    SQL = "SELECT Marcajes.idTrabajador,Marcajes.Horasincid,Fecha,Trabajadores.idHorario,IncFinal"
    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
    SQL = SQL & " Where Marcajes.Fecha >= '" & Format(Fini, FormatoFecha) & "'"
    SQL = SQL & " and Marcajes.Fecha <= '" & Format(FFin, FormatoFecha) & "'"
    SQL = SQL & strControlNomina
    

    SQL = SQL & " ORDER BY idHorario,Marcajes.idTrabajador,Fecha"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Horario = -1
    Trab = -1
    Set vH = New CHorarios
    While Not RS.EOF
        If RS!IdHorario <> Horario Then
            
            If vH.Leer(1, RS!IdHorario, Now) = 0 Then
                FESTIVOS = vH.LeerDiasFestivos(vH.IdHorario, Fini, FFin)
                MEDIODIA = vH.LeerMediosDias(vH.IdHorario, Fini, FFin)
            Else
                MsgBox "Error leyendo datos del horario:" & RS!IdHorario & ". El programa finalizara", vbExclamation
                Exit Sub
            End If
            Horario = RS!IdHorario
        End If
        
        If Trab <> RS!idTrabajador Then
            If Trab <> -1 Then
                If Dias > Int(Dias) Then
                    Dias = Int(Dias) + 1
                Else
                    Dias = Int(Dias)
                End If
                UpdateaHoras HC, HE, Dias, Trab
            End If
        
        
            HE = 0
            HC = 0
            Trab = RS!idTrabajador
            Dias = 0
                    
        End If
        
        'Si el dia esta en FESTIVOS no lo sumo
        Aux = Format(RS!Fecha, "dd/mm/yyyy") & "|"

        
        If InStr(1, FESTIVOS, Aux) = 0 Then
            'Si es medio dia sumo medio
            'NO esta en festivos  'NO esta en festivos   'NO esta en festivos  'NO esta en festivos
            If RS!IncFinal = vEmpresa.IncHoraExtra Then
                HC = HC + RS!HorasIncid
            End If
                
            If InStr(1, MEDIODIA, Aux) > 0 Then
                Dias = Dias + 0.5
            Else
                Dias = Dias + 1
            End If
        Else
            'FIESTA
            HE = HE + RS!HorasIncid
            

        End If

        
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    'Updatemaos el ultimo
    If Trab > 0 Then
        If Dias > Int(Dias) Then
            Dias = Int(Dias) + 1
        Else
            Dias = Int(Dias)
        End If
        UpdateaHoras HC, HE, Dias, Trab
    End If
    
'    'Updatemos con los dias trabajados.
'    '
'    'Acciones:
'    '       -En una variable cargaremos los dias festivos de
'    '       -En Otra Cargaremos los medios dias.
'    '       -Para cada dia trabajado, para cada trabajador, veremos
'    '       - Si los dias trabajados es un festivo o unidad fraccionarai
'
'    SQL = "SELECT idHorario"
'    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
'    SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
'    SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
'    SQL = SQL & strControlNomina
'    SQL = SQL & " GROUP BY Trabajadores.idHorario;"
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'
'    While Not RS.EOF
'        Set vH = New CHorarios
'        If vH.Leer(RS!IdHorario, Now) = 0 Then
'            FESTIVOS = vH.LeerDiasFestivos(vH.IdHorario, Fini, FFin)
'            MEDIODIA = vH.LeerMediosDias(vH.IdHorario, Fini, FFin)
'
'
'            'AHora para cada trabajador k haya trabajado entre las fechas le sumare dias trabajados
'            '.... o no(ej festivo)
'            SQL = "SELECT Marcajes.*"
'            SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
'            SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
'            SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
'            SQL = SQL & strControlNomina
'            SQL = SQL & " And Trabajadores.IdHorario = " & RS!IdHorario
'            SQL = SQL & " ORDER BY Marcajes.idTrabajador, Fecha"
'            Set RS2 = New ADODB.Recordset
'            RS2.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'
'            If Not RS2.EOF Then
'                Trabajador = -1
'                Do
'                   If Trabajador <> RS2!idTrabajador Then
'
'                        If Trabajador > 0 Then
'
'                            SQL = "UPDATE tmpHoras Set Dias = "
'                            If Dias > Int(Dias) Then
'                                Dias = Int(Dias) + 1
'                            Else
'                                Dias = Int(Dias)
'                            End If
'                            SQL = SQL & Int(Dias)
'                            SQL = SQL & " WHERE Trabajador = " & Trabajador
'                            Conn.Execute SQL
'                        End If
'
'                        Trabajador = RS2!idTrabajador
'                        Dias = 0
'
'                    End If
'
'
'                    'Sig
'                    RS2.MoveNext
'                Loop Until RS2.EOF
'
'                'Ahora faltara por hacer el ultimo trabajador
'                SQL = "UPDATE tmpHoras Set Dias = "
'                If Dias > Int(Dias) Then
'                    Dias = Int(Dias) + 1
'                Else
'                    Dias = Int(Dias)
'                End If
'                SQL = SQL & Int(Dias)
'                SQL = SQL & " WHERE Trabajador = " & Trabajador
'                Conn.Execute SQL
'            End If
'        End If
'        RS.MoveNext 'Siguiente horario
'    Wend
'
        
    'Por si acaso algun trabajador tiene numeros negativos
    SQL = "UPDATE tmpHoras Set Dias = 0"
    SQL = SQL & " WHERE Dias < 0 "
    conn.Execute SQL
    
    Set RS = Nothing
End Sub

Private Sub UpdateaHoras(ByRef vHC As Currency, ByRef vHE As Currency, ByRef vDias As Currency, ByRef T As Long)
Dim SQL As String

        SQL = "UPDATE tmpHoras Set HorasC = " & TransformaComasPuntos(CStr(vHC))
        SQL = SQL & " ,HorasE =  " & TransformaComasPuntos(CStr(vHE))
        SQL = SQL & " ,Dias =  " & vDias
        '
        vHC = vHC + vHE
        SQL = SQL & " ,HorasT = HorasT - " & TransformaComasPuntos(CStr(vHC))
        SQL = SQL & " WHERE Trabajador = " & T
        conn.Execute SQL
End Sub



Public Sub PonHorasExtraDeBolsa()
        conn.Execute "UPDATE tmpDatosMEs set ExtrasPeriodo = HorasE + Bolsadespues"
        espera 0.2
        conn.Execute "UPDATE tmpDatosMEs set Bolsadespues=0"
End Sub






'------------------------------------------------
'
' El objetivo final de los trabajadores semana
' es k trabajan 5 dias a la semans 8 horas
' Por lo tanto, no va a poder trabajar a la semana mas horas de
' las k las oficiales. Por lo tanto , con este sub
' Revisamos k las horas trabajadas no son mas de:
' dias * 8
' Si fuera mayor le sumariamos la diferencia a la bolsa k tuviera
' y en horas como mucho pondriamos
Public Sub HacerCompensacionSememana(FI As Date, FF As Date)
Dim SQLAUX As String
Dim RT As ADODB.Recordset

Dim H As Currency
Dim Def As Currency
Dim d As Integer


    SQLAUX = "SELECT  *"
    SQLAUX = SQLAUX & " , [diasperiodo]*8 AS Expr1, [HorasPeriodo]-[expr1] AS Diferencia"
    SQLAUX = SQLAUX & " FROM tmpDatosMes"
    d = Month(FI)
    SQLAUX = SQLAUX & " WHERE mes = " & d

    
    Set RT = New ADODB.Recordset

    RT.Open SQLAUX, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
            If RT!Diferencia > 0 Then
                Def = Abs(RT!Diferencia)
                SQLAUX = "UPDATE tmpdatosmes SET HorasPeriodo = " & TransformaComasPuntos(CStr((RT!expr1)))
                Def = Def + RT!bolsadespues
                SQLAUX = SQLAUX & " , BolsaDespues = " & TransformaComasPuntos(CStr(Def))
                SQLAUX = SQLAUX & " WHERE Mes = " & d & " AND Trabajador = " & RT!Trabajador
            Else
                SQLAUX = ""
            End If
            RT.MoveNext
            If SQLAUX <> "" Then conn.Execute SQLAUX
    Wend
    RT.Close
    Set RT = Nothing
    
End Sub




'Obtener anticpos pagados
'Pondremos los tipos
'               0, Pagos
'               1.- Anticpos

Public Sub ObtenerAnticposPagadosPorPrograma(FI As Date, FF As Date)
Dim SQLAUX As String
Dim RT As ADODB.Recordset

'    SQLAUX = "Select sum(importe) as impor,trabajador from pagos where tipo <2 "
'    SQLAUX = SQLAUX & " AND fecha>=#" & Format(FI, FormatoFecha)
'    SQLAUX = SQLAUX & "# AND fecha<=#" & Format(FF, FormatoFecha) & "#"
'    SQLAUX = SQLAUX & " GROUP BY trabajador"
'    Set RT = New ADODB.Recordset
'    RT.Open SQLAUX, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not RT.EOF
'        SQLAUX = "UPDATE tmpDatosMEs SET Anticipos=" & TransformaComasPuntos(RT!Impor)
'        SQLAUX = SQLAUX & " WHERE trabajador = " & RT!Trabajador
'        conn.Execute SQLAUX
'        RT.MoveNext
'    Wend
'    RT.Close
    Set RT = Nothing
End Sub


Public Sub CalculaDiferenciasDiasHoras()
Dim SQLAUX As String


    SQLAUX = "UPDATE tmpDatosMes SET SaldoH = Meshoras - HorasN, "
    SQLAUX = SQLAUX & " SaldoDias= MesDias - DiasTrabajados"
    SQLAUX = SQLAUX & " ,DiasPeriodo = DiasTrabajados"
    conn.Execute SQLAUX
End Sub




'--------------------------------
' EN tipo Alz la bolsa de horas
' pas directamente las HORASC as bolsa de horas
Public Sub ValoresBolsaDespues()
Dim SQLAUX As String
Dim RBolsa As ADODB.Recordset
Dim Bolsa As Currency

Dim RT As ADODB.Recordset
Dim I1 As Currency
Dim i2 As Currency
Dim Bruto As Currency

    Set RBolsa = New ADODB.Recordset
    SQLAUX = "select idtrabajador,HorasBolsa from trabajadoresbolsahoras where tipohora=1"
    RBolsa.Open SQLAUX, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    'SQLAUX = "SELECT Trabajadores.bolsahoras, Trabajadores.bolsaBRUTO, Trabajadores.IdTrabajador"
    SQLAUX = "SELECT Trabajadores.IdTrabajador"
    SQLAUX = SQLAUX & " , tmpDatosMEs.HorasC, Categorias.Importe2"
    SQLAUX = SQLAUX & " ,Trabajadores.porcss,Trabajadores.porcIRPF"
    SQLAUX = SQLAUX & " FROM tmpDatosMEs INNER JOIN (Categorias INNER JOIN Trabajadores ON"
    SQLAUX = SQLAUX & " Categorias.IdCategoria = Trabajadores.idCategoria) ON tmpDatosMEs.Trabajador = Trabajadores.IdTrabajador;"

    Set RT = New ADODB.Recordset
    RT.Open SQLAUX, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RT.EOF
        
        'Bolsa despues
        Bolsa = 0
        RBolsa.Find "idtrabajador =" & RT!idTrabajador, , adSearchForward, 1
        If Not RBolsa.EOF Then Bolsa = DBLet(RBolsa!horasbolsa, "N")
        
        i2 = Bolsa + RT!HorasC
        'SQL
        SQLAUX = "UPDATE tmpDatosMes SET bolsadespues = " & TransformaComasPuntos(CStr(i2))
        
        
        
        If False Then
                        'La bolsa k importe supone bruto
                        I1 = RT!HorasC * DBLet(RT!Importe2, "N")
                        I1 = Round(I1, 2)
                        Bruto = I1
                        i2 = 0   'DBLet(RT!bolsabruto, "N")
                        i2 = i2 + I1
                        SQLAUX = SQLAUX & ",brutodespues = " & TransformaComasPuntos(CStr(i2))
                        
                        'El neto
                        i2 = DBLet(RT!PorcSS, "N") + DBLet(RT!PorcIRPF, "N")
                        i2 = i2 / 100
                        i2 = Round(Bruto * i2, 2)
                        i2 = Bruto - i2
                        
                        i2 = i2 + 0 ' DBLet(RT!bolsaneto, "N")
                        SQLAUX = SQLAUX & ",netodespues = " & TransformaComasPuntos(CStr(i2))
                    
        End If
                    
                    
        'idTrabajador
        SQLAUX = SQLAUX & " WHERE Trabajador = " & RT!idTrabajador
        
        RT.MoveNext
        'Ejecutamos
        conn.Execute SQLAUX
    Wend
    RT.Close
    Set RT = Nothing
End Sub



'-------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------
'
'   Compensacion PICASSENT a partir de octubre 2008
'
'   -La compensación de horas extras se realizar siempre por días que hayan faltado al trabajo y
'    NUNCA para completar horas de los días que no hayan realizado las 8 horas.
'
'
'   - Si un dia, las horas minimo no llegan al minimo por dia NO entra en nomina


'  Hay X y S que pueden haber trabajado pero no contar. Por ello,
'   devuelve cuantos dias reajusta(XS_NoCuentaTrabajados), pero en este procedimiento tambien
'    al incremetar XS_NoCuentaTrabajados debe permitir compensarle este dia

'las horas a compensar se pasan por referencia
Private Function CompensacionesDiaTrabajadoYSemana(Dias As Integer, ByRef Rec As Recordset, ByRef FEST As String, ByVal FI As Date, ByVal FF As Date, ByRef vHO As CHorarios, HorasMinimoDia2 As Currency, ByRef HorasCompensarQueUtiliza As Currency, ByRef DiasQueReajusteXSTrabajados As Integer, ByRef EstaDeBajaTodoElMes As Boolean) As Integer
Dim RF As ADODB.Recordset
Dim cad As String
Dim Fin  As Boolean
Dim Horas As Currency
Dim Sig As Boolean
Dim DiaC As Currency
Dim FechaReferencia As Date
Dim DiaSem As Byte
Dim Semana As Integer
Dim Js As Integer
Dim TrabajaMier As Boolean
Dim TrabajaSab As Boolean
Dim CompensaMitad As Boolean
Dim vFechaINicio As Date
Dim vfechaFin As Date
Dim HorasCompensar2 As Currency
Dim Rhoras As ADODB.Recordset
'Dim XS_NoCuentaTrabajados As String
'Dim XS_NoCtanYPuedoCompensarlos As String
Dim HorasTrabajadas As Currency
Dim TieneSab As Boolean
Dim TieneMier As Boolean
Dim NoCuentaM As Boolean   'No cuenta como trabajado
Dim NoCuentaS As Boolean   'No cuenta como trabajado
Dim BAJAS As String
Dim HorasAuxiliar As Currency
Dim MinimoPorDia As Currency

Dim SemanaNormal As Boolean   'Si tiene la semana miercoles y sabado. Para la primera del mes
On Error GoTo ECompensacionesDiaTrabajado

    CompensacionesDiaTrabajadoYSemana = 0
    'Si fecha alta > fecha inicio mes enonces finicio mes=fecha alta
    If Rec!FecAlta > FI Then FI = Rec!FecAlta
    vFechaINicio = FI
    
    'Si fecha baja < fecha baja mes entonces finicio mes=fecha alta
    If Not IsNull(Rec!FecBaja) Then
        If Rec!FecBaja < FF Then FF = Rec!FecBaja
    End If
    vfechaFin = FF
    
    DiasQueReajusteXSTrabajados = 0
    HorasCompensar2 = HorasCompensarQueUtiliza
    EstaDeBajaTodoElMes = False
    'XS_NoCuentaTrabajados = ""

'    If Rec!Trabajador = 8 Or Rec!Trabajador = 8 Then
'
'    Else
'        Exit Function
'    End If

    Set RF = New ADODB.Recordset
    
    
    'Las bajas
    
    cad = "Select * from tmpcombinada where idtrabajador=" & Rec!Trabajador
    RF.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    BAJAS = ""
    If Not RF.EOF Then
        
        BAJAS = "|"
        While Not RF.EOF
            FechaReferencia = RF!Fecha
            Do
                BAJAS = BAJAS & Format(FechaReferencia, "dd/mm/yyyy") & "|"
                FechaReferencia = DateAdd("d", 1, FechaReferencia)
            Loop Until FechaReferencia > RF!H1
            RF.MoveNext
        Wend
    End If
    RF.Close


    cad = "Select fecha,horastrabajadas,horasincid from marcajes"
    cad = cad & " WHERE Fecha >='" & Format(FI, FormatoFecha) & "'"
    cad = cad & " AND Fecha <='" & Format(FF, FormatoFecha) & "'"
    cad = cad & " AND idTrabajador = " & Rec!Trabajador
    cad = cad & " ORDER BY fecha"
    RF.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    
    
    Horas = 0
    DiaC = 0
    Fin = False
    Set Rhoras = New ADODB.Recordset
    HorasTrabajadas = 0
    
    If Not RF.EOF Then
    
        'Pasada CERO. Para el trabajador, si hay dias que ha trabajado MENOS horas de las 3.5 pactadas
        While Not RF.EOF
            
            FI = RF!Fecha
            DiaSem = Weekday(FI, vbMonday)
            If DiaSem <> 6 And DiaSem <> 3 Then
                If InStr(1, FEST, Format(FI, "dd/mm/yyyy") & "|") = 0 Then
                    
                    MinimoPorDia = RF!HorasTrabajadas - 3.5
                
                    If MinimoPorDia < 0 Then
                        Horas = Horas + Abs(MinimoPorDia)
                    End If
                End If
            End If
            RF.MoveNext
        Wend
        

        
        RF.MoveFirst
    
        'PRIMERA PASADA--------------------------------------------------------------------------
        
        FI = vFechaINicio
        Semana = -1

        While Not Fin
            
            Js = CInt(Format(FI, "ww", vbMonday))
            If Js <> Semana Then    '//PUNTO 1
                'Ha cambiado de semana
                If Semana > 0 Then
                    'Haremos los calculos
                    
                    
                    ActualizarDatosDeUnMiercolesSabado TieneMier, TrabajaMier, TieneSab, TrabajaSab, Horas, DiaC, _
                        Rec!Trabajador, HorasTrabajadas, HorasMinimoDia2, NoCuentaM, NoCuentaS, DiasQueReajusteXSTrabajados, HorasCompensar2
                        
                    
                    
                    If DiaC > Dias Then DiaC = Dias 'NO puede compensar mas dias de los que pueden ir en nomina
                    

                    
      
        
                End If
                
                'Reestablecemos variables
                Semana = Js
                'Puede que la semana NO tenga miercoles o sabado, pq sean festivos
                ' o pq sea la primera semana del mes
                If Semana = SemanaMesPrimera Or Semana = SemanaMesUltima Then
                    'Primera o ultima semana del mes. Si es el primer dia es jueves no tiene miercoles y si es domingo no tiene ninguno
                    If Semana = SemanaMesUltima Then
                        'Compruebo frente al ULTIMO DIA DEL MES
                        DiaSem = Weekday(vfechaFin, vbMonday)
                        TieneSab = DiaSem >= 6
                        TieneMier = DiaSem >= 3
                        
                        
                    Else
                        'Compruebo frente al primer dia del mes
                        DiaSem = Weekday(vFechaINicio, vbMonday)
                        TieneSab = DiaSem <= 6
                        TieneMier = DiaSem <= 3
                        
                    End If
                Else
                    'En cualquier otro caso, deberia tener miercoles sabado
                    TieneSab = True
                    TieneMier = True
                End If
                TrabajaSab = False
                TrabajaMier = False
                NoCuentaM = False
                NoCuentaS = False
                HorasTrabajadas = 0
                
        
            End If  '//end PUNTO 1
            
            If RF.EOF Then
                FechaReferencia = DateAdd("d", -1, FI)  'Para que pueda entrar en los ifs
            Else
                FechaReferencia = RF!Fecha
            End If
            
  
                'Igual o mayor a la fecha en BD
                If FechaReferencia = FI Then                                           '// PUNTO 3
                '---------------------------------
                    DiaSem = Weekday(FI, vbMonday)
                    If (DiaSem = 3 Or DiaSem = 6) Then                          '// PUNTO 4. Tranabaja el miercoles o sab
                    'Ha trabajado. Veremos si se lo cuento como trabajado o no
                    'Comprobaremos con el sQL si le cuento como trabajado o no
                        If DiaSem = 3 Then
                            cad = "Select max(hora) from entradamarcajes where fecha = '"
                            cad = cad & Format(FI, FormatoFecha) & "' AND idTrabajador = " & Rec!Trabajador
                        Else
                            cad = "Select min(hora) from entradamarcajes where fecha = '"
                            cad = cad & Format(FI, FormatoFecha) & "' AND idTrabajador = " & Rec!Trabajador
                        End If
                        Rhoras.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If Rhoras.EOF Then
                            MsgBox "Error grave: " & cad, vbExclamation
                            End
                        End If
                        
                        
                        'Si que es miercoles o sabado
                        HorasAuxiliar = 0
                        If DiaSem = 3 Then
                            If Rhoras.Fields(0) > CDate(HoraIntermediaMiercolesSabado) Then
                            
                                TrabajaMier = True
                                'HorasTrabajadas = HorasTrabajadas + RF!HorasTrabajadas
                                'HorasAuxiliar = RF!HorasTrabajadas - RF!HorasIncid
                                HorasAuxiliar = HorasTrabajadasDeUnMiercolesSabado(Rec!Trabajador, RF!HorasTrabajadas, RF!HorasIncid, FI)
                                HorasTrabajadas = HorasTrabajadas + HorasAuxiliar
                            Else
                                'XS_NoCuentaTrabajados = XS_NoCuentaTrabajados & FI & "|"
                                NoCuentaM = True
                            End If
                            
                            
                        Else
                            If Rhoras.Fields(0) < CDate(HoraIntermediaMiercolesSabado) Then
                                TrabajaSab = True
                                HorasAuxiliar = HorasTrabajadasDeUnMiercolesSabado(Rec!Trabajador, RF!HorasTrabajadas, RF!HorasIncid, FI)
                                'HorasTrabajadas = HorasTrabajadas + RF!HorasTrabajadas
                                HorasTrabajadas = HorasTrabajadas + HorasAuxiliar
                            Else
                                NoCuentaS = True
                            End If
                        End If
                        If HorasAuxiliar < 0 Then
                            MsgBox "Horas trabajadas menor que CERO : " & FI & "  - " & Rec!Trabajador, vbExclamation
                            
                        End If
                            
                        Rhoras.Close
                    End If
                    RF.MoveNext
                    FI = DateAdd("d", 1, FI)
                Else
                    'NO trabaja este dia
                    DiaSem = Weekday(FI, vbMonday)
                    If Not (DiaSem = 3 Or DiaSem = 6) Then
                        'Otros dias semanas.
                        
                    Else
                        'MIERCOLES Y SABADO
                        'Si es FESTIVO o de baja pongo que NO tenia que trabajar el dia en cuestion
                        If InStr(1, BAJAS, Format(FI, "dd/mm/yyyy") & "|") > 0 Or InStr(1, FEST, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                            If DiaSem = 3 Then
                               TieneMier = False
                            Else
                               TieneSab = False
                            End If
                        
                        Else
                            'NO ha trabajado este dia
                            'NO  HACEMOS NADA
                        
                        End If
                        
                    End If                                          '//end PUNTO 4
                    FI = DateAdd("d", 1, FI)
                End If                                              '//end PUNTO 3

            If Not Fin Then
                If FI > vfechaFin Then
                    Fin = True
                    If CInt(Format(FI, "ww", vbMonday)) = Semana Then
                        Js = Js + 1 'Para que entre en el if de abajo
                    End If
                End If
            End If
        Wend
        
        
        'Para la ultima semana
        If Js <> Semana And Semana >= 0 Then
            
            ActualizarDatosDeUnMiercolesSabado TieneMier, TrabajaMier, TieneSab, TrabajaSab, Horas, DiaC _
                , Rec!Trabajador, HorasTrabajadas, HorasMinimoDia2, NoCuentaM, NoCuentaS, DiasQueReajusteXSTrabajados, HorasCompensar2
        Else
            '
        End If
        
        
        'QUTIAR#
        If DiasQueReajusteXSTrabajados > 0 Then
            'Debug.Print Rec!Trabajador
            '
        End If
        If DiaC > Dias Then DiaC = Dias 'NO puede compensar mas dias de los que pueden ir en nomina


        
    
    
    
    
    
            'SEGUNDA PASADA--------------------------------------------------------------------------
            'SEGUNDA PASADA--------------------------------------------------------------------------
            'SEGUNDA PASADA--------------------------------------------------------------------------
            'SEGUNDA PASADA--------------------------------------------------------------------------
            'busco compensar dias de 8 en 8 horas
            RF.MoveFirst
            FI = vFechaINicio
            FF = vfechaFin
            Fin = False
            
    
            'Si no tiene OCHO horas no puedo compensarle NI UN SOLO DIA mas
            If HorasCompensar2 - Horas < 8 Then Fin = True
            FechaReferencia = RF!Fecha
            While Not Fin
                
                If FI > FF Then
                    Fin = True
                    Sig = False 'Para k no mueva el recordset
                Else
                    If FI = FechaReferencia Then
                        FI = DateAdd("d", 1, FI)
                        Sig = True
                        If RF!Fecha = FechaReferencia Then
                            If InStr(1, BAJAS, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                                'Ha trabajaod estando de baja
                                'El primer dia puede darse el caso
                                '
                                'Debug.Print FI & " " & Rec!Trabajador
                            Else
                                
                                DiaSem = Weekday(FechaReferencia, vbMonday)
                                If DiaSem <> 3 And DiaSem <> 6 Then
                                    If RF!HorasTrabajadas < HorasMinimoDia2 Then
                                        'Auqi compensamos
                                        'Le debemos compensar las horas hasta
                                        
                                        'ESTE PASO lo hace ahora en el PASO 0
                                        'HorasTrabajadas = HorasMinimoDia2 - RF!HorasTrabajadas
                                        'Horas = Horas + HorasTrabajadas
                                        'Si no le kedan horas para compensar tampoco seguimos
                                        If HorasCompensar2 - Horas < HorasMinimoDia2 Then Fin = True
                                    End If
                                Else
       
                                End If
                            End If
                        End If
                    Else
                        If InStr(1, FEST, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                            'Es un dia festivo
                            FI = DateAdd("d", 1, FI)
                            Sig = False
                        Else
                            'Es de bajas
                            If InStr(1, BAJAS, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                                FI = DateAdd("d", 1, FI)
                                Sig = False
                            Else
                                DiaSem = Weekday(FI, vbMonday)
                                If DiaSem = 3 Or DiaSem = 6 Then
                                    'YA LO HEMOS PROCESADO
                                    'Pero podria ser que tiene a compensar. ha trabajado por la mañana cuando debia trabajar por la tarde
                                    FI = DateAdd("d", 1, FI)
                                    Sig = False
                                Else
                                        'Es un dia k no ha  trabajado. Vemos cuantas horas son
                                        vHO.Leer vHO.IdHorario, FI, 1
                                        'Ya tenog las horas k debia haber trabajado
                                        If Horas + vHO.TotalHoras <= HorasCompensar2 Then
                                            If DiaC < Dias Then
                                                'Si el dia es de miercoles o sabado SI que quito las horas
                                                'Le puedo compensar este dia
                                                DiaC = DiaC + vHO.DiaNomina
                                                
                                                Horas = Horas + vHO.TotalHoras
                                             Else
                                               '
                                            End If
                                        Else
                                            'Este dia no se lo puedo compensar
                                            'No hago nada
                                        End If
                                                     
                                        FI = DateAdd("d", 1, FI)
                                        Sig = False
                                                     
                                        'Por si acaso ya ha compensado todos los dias
                                        If DiaC >= Dias Then
                                            If DiaC > Dias Then DiaC = Dias
                                            'Fin = True
                                            'No pongo el FIN, pq puede compensarles horas todavia
                                            
                                        End If
                                        
                                        'Si no le kedan horas para compensar tampoco seguimos
                                        If HorasCompensar2 - Horas < 8 Then Fin = True
                                        
                                End If 'de diasem
                            End If 'de bajas
                        End If
                    End If
                
            
                End If
                If Sig Then
                    If RF.EOF Then
                        'Deberiamos salir
                        '
                    Else
                        RF.MoveNext
                        'ANTES
                        'If RF.EOF Then Fin = True
                        If Not RF.EOF Then FechaReferencia = RF!Fecha
                    End If
                End If
            Wend
        
        
        
        
        
        
    Else   'rt.eof
    
        RF.Close
        cad = "Select * from bajas where idtrab = " & Rec!Trabajador
        cad = cad & " AND fechaalta is null and fechabaja < '" & Format(FI, FormatoFecha) & "'"
        RF.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        If Not RF.EOF Then
            If Not IsNull(RF!idTrab) Then
                cad = "BAJA"
                EstaDeBajaTodoElMes = True
            End If
        End If
        
        'NO HA TRABAJADO, pero tiene Horas de otros meses
        If cad = "" Then
                DiaC = Int(HorasCompensar2) \ 8    'Cuantos dias de 8 horas le entran
                
                If DiaC > Dias Then DiaC = Dias
                
                Horas = (DiaC * 8)
                HorasTrabajadas = HorasCompensar2 - Horas
                If HorasTrabajadas >= HorasMinimoDia2 Then
                    If DiaC < Dias Then
                        DiaC = DiaC + 1
                        If HorasTrabajadas > 8 Then HorasTrabajadas = 8
                        Horas = Horas + HorasTrabajadas
                    End If
                End If
                If DiaC >= Dias Then DiaC = Dias                  'NO puede compensar mas dias de los que pueden ir en nomina
        End If
    End If
    RF.Close
    Set RF = Nothing
    Set Rhoras = Nothing
    If DiaC < 0 Then
        ' DiaC = 0
        'Debug.Print Rec!Trabajador
       '
    End If
    If DiaC > Int(DiaC) Then
        DiaC = Int(DiaC) + 1
        If DiaC > Dias Then DiaC = Dias 'NO puede compensar mas dias de los que pueden ir en nomina
    End If
        
    HorasCompensarQueUtiliza = Horas
    CompensacionesDiaTrabajadoYSemana = DiaC
        
    Exit Function
ECompensacionesDiaTrabajado:
    MuestraError Err.Number, "CompensacionesDiaTrabajado" & vbCrLf & Err.Description

End Function

'Devuelve horas de TRABAJADAS (no las trab + incid) de un sabado y miercoles
'Puede ser que el dias trabajado tengamos que compensarle un hora
Private Function HorasTrabajadasDeUnMiercolesSabado(idTRa As Long, HMarcaje As Currency, Hincid As Currency, Fecha As Date) As Currency
Dim RN As ADODB.Recordset
Dim Horas As Currency
Dim C As Currency
Dim h2 As Currency
    Horas = HMarcaje - Hincid
    
    HorasTrabajadasDeUnMiercolesSabado = Horas
    Set RN = New ADODB.Recordset
    RN.Open "Select HT from tmpPagosMes where idTrabajador = " & idTRa, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RN.EOF Then
        If Not IsNull(RN!HT) Then
            h2 = NuevoCalculoHorasDiaXS(idTRa, Fecha, Weekday(Fecha, vbMonday) = 3)
            If h2 = RN!HT Then conn.Execute "DELETE FROM tmppagosmes WHERE idtrabajador = " & idTRa
            
            Horas = Horas - h2
            
        End If
    End If
    RN.Close
    Set RN = Nothing
    HorasTrabajadasDeUnMiercolesSabado = Horas
End Function


'Habran unas variabes que se pasan por referencia para actualizarlas
Private Sub ActualizarDatosDeUnMiercolesSabado(TieneTrabajarX As Boolean, TrabajaMiercoles As Boolean, TieneTrabajarS As Boolean _
    , TrabajaSabado As Boolean, ByRef HorasQueCompensa As Currency, ByRef DiasQueCompensa As Currency, Trabajador As Long _
        , HorasTrabajadas As Currency, HorasMinimoDia As Currency, NoCuentaMier As Boolean, NoCuentaSab As Boolean, ByRef DiasXySNoCuentan As Integer, HorasMaximasACompensar As Currency)
Dim UnDiaMenosTrabajadoUnDiaMasCompensable As Boolean
                
        Dim Diferencia As Currency
        Dim HaCompensadoEnEstaFuncion As Boolean
        
               HaCompensadoEnEstaFuncion = False
               If Not TrabajaSabado And Not TrabajaMiercoles Then
                    If Not TieneTrabajarX And Not TieneTrabajarS Then
                        'NO hacemos nada. No ha trabajado, pero tampoc tenia pq trabajar TrabajaMiercoles
                
                
                    Else
                        Diferencia = HorasMaximasACompensar - HorasQueCompensa
                        
                        
                        If TieneTrabajarX And TieneTrabajarS Then
                            'Debia haber trabajado los dos dias y no ha trabajado ninguno
                            If Diferencia >= 8 Then
                                HaCompensadoEnEstaFuncion = True
                                HorasQueCompensa = HorasQueCompensa + 8
                                DiasQueCompensa = DiasQueCompensa + 1   'Compensa UN DIA
                            Else
                                If Diferencia >= 3.5 Then
                                    HaCompensadoEnEstaFuncion = True
                                    DiasQueCompensa = DiasQueCompensa + 1   'Compensa UN DIA
                                    HorasQueCompensa = HorasQueCompensa + Diferencia
                                End If
                            End If
                        Else
                            'Tenia que haber trabajado el sabado
                            If TieneTrabajarX Then
                                If Diferencia > 3.5 Then
                                    HorasQueCompensa = HorasQueCompensa + 3.5
                                    DiasQueCompensa = DiasQueCompensa + 1   'Compensa UN DIA
                                Else
                                
                                    HorasQueCompensa = HorasMaximasACompensar
                                End If
                            Else
                                If Diferencia > 4.5 Then
                                    HorasQueCompensa = HorasQueCompensa + 4.5
                                    DiasQueCompensa = DiasQueCompensa + 1   'Compensa UN DIA
                                Else
                                    'NO podemos comepnsarle el dia
                                    'HorasQueCompensa = HorasMaximasACompensar
                                    DiasQueCompensa = DiasQueCompensa - 1
                                End If
                            End If
                            
                        End If
                            
                    End If
                Else
                    'Ha trabajado alguno de los dos dias
                    'Si entre los dos dias no suma 3.5 tb le compensaremos
                        
                        If HorasTrabajadas < 3.5 Then   'HorasMinimoDia2=3.5
                            'MsgBox "Compensa horas X y S trabajador: " & Trabajador & "       Horas: " & HorasTrabajadas, vbExclamation
                            If HorasTrabajadas < 0 Then
                                
                        
                                HorasTrabajadas = Abs(HorasTrabajadas)
                                HorasQueCompensa = HorasQueCompensa + HorasTrabajadas
                            Else
                                'Le debemos compensar las horas hasta
                                HorasTrabajadas = HorasMinimoDia - HorasTrabajadas
                                HorasQueCompensa = HorasQueCompensa + HorasTrabajadas
                            End If
                        End If
    
                End If

            If NoCuentaMier Or NoCuentaSab Then
                UnDiaMenosTrabajadoUnDiaMasCompensable = False
                If NoCuentaMier And NoCuentaSab Then
                    'Un dia mas de compensar
                    UnDiaMenosTrabajadoUnDiaMasCompensable = True
                    
                Else
                    If NoCuentaMier And Not TrabajaSabado Then
                        'Solo habia trabajado el miercoles, y como ademas no le debe contar como dia trabajado
                        'incrementamos la variable de dias ...
                        
                        'y le dejamos compensar un dia mas
                        UnDiaMenosTrabajadoUnDiaMasCompensable = True
                    End If
                    If NoCuentaSab And Not TrabajaMiercoles Then
                        'Solo habia trabajado el sabado, y como ademas no le debe contar como dia trabajado
                        'incrementamos la variable de dias ...
                        
                        'y le dejamos compensar un dia mas
                        UnDiaMenosTrabajadoUnDiaMasCompensable = True
                    End If
                End If
                If UnDiaMenosTrabajadoUnDiaMasCompensable Then
'                    If HaCompensadoEnEstaFuncion Then DiasQueCompensa = DiasQueCompensa - 1
                    'If Not HaCompensadoEnEstaFuncion Then
                    DiasXySNoCuentan = DiasXySNoCuentan + 1
                End If
            End If
End Sub



Private Sub RecalculoHorasMiercolesSabados(F1 As Date, F2 As Date, vLbl As Label, Miercoles As Boolean)
Dim cad As String
Dim RF As ADODB.Recordset
Dim Trab As Long
Dim HT As Currency
Dim Horas As Currency

    vLbl.Caption = "Recalculo horas miercoles"
    vLbl.Refresh
    If Miercoles Then
        Trab = 4
    Else
        Trab = 7
    End If
    cad = "SELECT EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha, Weekday([Fecha]) AS Expr1"
    cad = cad & " From EntradaMarcajes"
    cad = cad & " Where EntradaMarcajes.Fecha >= #" & Format(F1, FormatoFecha) & "# And"
    cad = cad & " EntradaMarcajes.Fecha <= #" & Format(F2, FormatoFecha) & "# And "
    cad = cad & " Weekday([Fecha]) = " & Trab & " And Hora "
    If Miercoles Then
        cad = cad & " <"
    Else
        cad = cad & " >"
    End If
    cad = cad & " #14:00:00# group by  EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha,  Weekday([Fecha])"
    cad = cad & " ORDER BY EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha"
    Set RF = New ADODB.Recordset
    RF.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Trab = -1
    While Not RF.EOF
    
        'If RF!idTrabajador = 65 Then St op
    
        If Trab <> RF!idTrabajador Then
            vLbl.Caption = "R.H. (" & Val(RF!expr1) - 1 & ")  trab:" & RF!idTrabajador
            vLbl.Refresh
            'Nuevo trabajador
            If Trab > 0 Then
                'Updateamos los nuevos valores en tmphoras
                If HT > 0 Then UpdateaNuevosValoresMiercolesSabado Trab, HT, Month(RF!Fecha)
            End If
            'Reseteamos variables
            Trab = RF!idTrabajador
            HT = 0
            
        End If
        
        Horas = NuevoCalculoHorasDiaXS(Trab, RF!Fecha, Miercoles)
        HT = HT + Horas
        RF.MoveNext
        
    Wend
    RF.Close
    Set RF = Nothing
    'EL ultimo
    If Trab > 0 And HT > 0 Then UpdateaNuevosValoresMiercolesSabado Trab, HT, Month(F1)
End Sub

Private Sub UpdateaNuevosValoresMiercolesSabado(idTRa As Long, Hor As Currency, KMes As Integer)
Dim cad As String


    cad = "UPDATE tmpDatosMes SET horasN= horasN - " & TransformaComasPuntos(CStr(Hor))
    cad = cad & " , horasc=horasc +  " & TransformaComasPuntos(CStr(Hor))
    cad = cad & " WHERE mes= " & KMes & " AND Trabajador =" & idTRa
    conn.Execute cad
    
    'Updateamos en tmpPagosMes, para que luego se las compense en nomina
    cad = "INSERT INTO tmpPagosMes (idTrabajador,nombre,HT) VALUES (" & idTRa & ",'Compensaciones'," & TransformaComasPuntos(CStr(Hor)) & ")"
    conn.Execute cad
    Debug.Print "Nuevos vlaores XyS:" & idTRa
End Sub

Public Function NuevoCalculoHorasDiaXS(Trabajador As Long, Fecha As Date, DeMiercoles As Boolean) As Currency
Dim RH As ADODB.Recordset
Dim C As String
Dim T1 As Currency
Dim T2 As Currency
Dim E As Boolean
Dim Seguir As Boolean
Dim NuevaHora As Currency

    NuevoCalculoHorasDiaXS = 0
    Set RH = New ADODB.Recordset
    
    C = "Select * from entradamarcajes where idtrabajador=" & Trabajador
    C = C & " AND fecha = #" & Format(Fecha, FormatoFecha) & "# "
    

    C = C & " ORDER BY hora"

    RH.Open C, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    Seguir = True
    'Si tiene algun --> Miercoles y es por la tarde ---->NORMAL, no hago nada
    '                     sabado   y es por la mañana ---> "
    If Not RH.EOF Then
        If DeMiercoles Then
       
            If RH!Hora > CDate(HoraIntermediaMiercolesSabado) Then
                Seguir = False
            Else
                'RH.MoveFirst
            End If
        Else
            RH.MoveLast
           'sabado
            If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                Seguir = False
            Else
                RH.MoveFirst
            End If
            
        End If
    End If
    
    
    If Not Seguir Then
        RH.Close
        Exit Function
    End If
    E = True
    NuevaHora = 0
    While Not RH.EOF

        If E Then
            If DeMiercoles Then
                If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                    T1 = CCur(DevuelveValorHora(RH!Hora))
                Else
                    'Ya es cuando le toca
                    RH.MoveLast
                End If
                
             Else
                If RH!Hora >= CDate(HoraIntermediaMiercolesSabado) Then
                    T1 = CCur(DevuelveValorHora(RH!Hora))
                Else
                    'Ya es cuando le toca
                  
                End If
             End If
                
        Else
            'Si tiene valor t1 calculamos dif
            If T1 > 0 Then
                T2 = CCur(DevuelveValorHora(RH!Hora))
                T1 = T2 - T1
            
                NuevaHora = NuevaHora + T1
                T1 = 0

            End If
        End If
        E = Not E
        RH.MoveNext
    Wend
        
        
    'veremos si este dia realmente es como si no lo trabajadra
    'ya que si tenia que haber venido por la mañana, pero solo viene por
    'la tarde a efectos de nomina no lo cuento
    
    If Trabajador < 700 Then
        If DeMiercoles Then
            RH.MoveLast
            'nos vamos al ultimo. Si la hora es mayor que las 2 NO loañado a la lista
            If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                'Este dia no lo contare para la nomina
                C = "INSERT INTO tmpNoTrabajo (idtra,idfech) VALUES (" & Trabajador & ",#" & Format(Fecha, FormatoFecha) & "#)"
                conn.Execute C
            End If
        Else
            RH.MoveFirst
            If RH!Hora > CDate(HoraIntermediaMiercolesSabado) Then
                
                C = "INSERT INTO tmpNoTrabajo (idtra,idfech) VALUES (" & Trabajador & ",#" & Format(Fecha, FormatoFecha) & "#)"
                conn.Execute C
            End If
         End If
    End If
    RH.Close    'para que no coja los 700 y 900
    If NuevaHora > 0 And Trabajador < 700 Then
        C = "Select * from marcajes where idtrabajador=" & Trabajador
        C = C & " AND fecha = #" & Format(Fecha, FormatoFecha) & "# "
        RH.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RH.EOF Then
            MsgBox "Mal: " & C
        Else
            
            
            
            
            'Si deberia tenre mas horas comepnsables
            
            If RH!IncFinal = vEmpresa.IncRetraso Then
                'No hacemos nada, ya que el valor calculado sera el bueno
             
            Else
                If NuevaHora > RH!HorasIncid Then
                    T1 = NuevaHora - RH!HorasIncid
                    'Tenemos lo que incrementa y decrementa en comepnsables y en normales
                    'nunca puede aumentar mas que lo que ha trabajador
                    T2 = RH!HorasTrabajadas - RH!HorasIncid
                    If T1 > T2 Then T1 = T2
                    NuevaHora = T1
                Else
                   'No hacemos nada
                   NuevaHora = 0
                End If
                
            End If
            RH.Close
            NuevoCalculoHorasDiaXS = NuevaHora
        End If
    End If
End Function

Private Sub AjustarHorasNormales()
Dim RT As ADODB.Recordset
    
    conn.Execute "UPDATE tmpdatosmes set horasn=horasn-horasc"
    espera 0.5
    Set RT = New ADODB.Recordset
    RT.Open "Select * from tmpdatosmes where horasn<0 ", conn, adOpenForwardOnly, adLockOptimistic
    If Not RT.EOF Then
        MsgBox "Avise soporte tecnico." & vbCrLf & "HorasN<0", vbExclamation
    End If
    RT.Close
End Sub

'----------------------------
'dentro del nuevo proc
'                        If TrabajaSab Then
'                            'COMPENSA miercoles.
'                            'Habra que comprobar que el miercoles de esa semana
'                            'Esta en el periodo de calculo
'                            '
'                            CompensaMitad = True
'                            If CInt(Format(vFechaINicio, "ww")) = Js Then
'                                'Primera semana del calculo.
'                                'Si el dia es mayor que miercoles NO tiene miercoels a comepnsar
'                                DiaSem = Weekday(vFechaINicio, vbMonday)
'                                If DiaSem > 3 Then CompensaMitad = False                             'NO TIENE miercoles
'                            End If
'                            If CompensaMitad Then Horas = Horas + 3.5
'
'                        End If
'
'                        If TrabajaMier Then
'                            'COMPENSA SABADO.
'                            'Habra que comprobar que el sabado de esa semana
'                            'Esta en el periodo de calculo
'                            '
'                            CompensaMitad = True
'                            If CInt(Format(vfechaFin, "ww")) = Js Then
'                                'Primera semana del calculo.
'                                'Si el dia es mayor que miercoles NO tiene miercoels a comepnsar
'                                DiaSem = Weekday(vfechaFin, vbMonday)
'                                If DiaSem < 6 Then CompensaMitad = False                 'NO TIENE sabado
'                            End If
'                            If CompensaMitad Then Horas = Horas + 4.5
'
'                        End If
'                        If HorasCompensar2 - Horas < 4.5 Then Fin = True
'                    End If
'                End If











'------------------------------------------------------------
Public Sub HacerCompensacionesPicassent(FInicio As Date, FFin As Date, lbl As Label)
Dim NumeroDiasMes As Integer
Dim HCompMes As Currency
Dim HPaBolsa As Currency
Dim DiasOF As Integer
Dim HorasOf As Currency
Dim H As Currency
Dim SQL As String
Dim Horario As Integer
Dim vH As CHorarios
Dim FESTIVOS As String
Dim MEDIODIA As String
Dim PrimerMiercoles As Integer
Dim BajaTodoElMes As Boolean
Dim RS As ADODB.Recordset
Dim Dias As Integer
Dim cT As CTrabajadorNomina
Dim SumatorioHorasAntes As Currency
Dim SumatorioAux As Currency
Dim RtAnticipos As ADODB.Recordset


    'Vemos cual es el modo de compensacion
    '   0 .- NO compensa
    '   1 .- A partir de los dias trabajados del trabajador
    '         vemos cuantos dias le puedo compensar
    '   2 .- X horas hacen una jornada laboral a compensar
    '   3 .- Picassen cotubre 2008.
    '           -Compensaran por semana /dia con cuidado a los miercoles sabados
    '           -si trabaja una hora un dia, el resto de horas NO las tiene que compensar para la nomina
    ' SIEMPRE 3
    
    If Depuracion Then
        'VAmos a depurar algunos trabajadores
        lbl.Caption = "Depuracion"
        lbl.Refresh
        TrabajadoresDepuracion = "|"
        
        Do
            SQL = InputBox("Cod. trabajador a depurar?" & vbCrLf & "(Vacio=Fin)", "Depuracion")
            If SQL <> "" Then
                If IsNumeric(SQL) Then
                    If InStr(1, TrabajadoresDepuracion, "|" & SQL & "|") = 0 Then TrabajadoresDepuracion = TrabajadoresDepuracion & SQL & "|"
                End If
            End If
        Loop Until SQL = ""
        If TrabajadoresDepuracion = "|" Then Depuracion = False
    End If
        
    'Ajustes ponemos HN las que tiene menos las que sean extra
    lbl.Caption = "Ajuste horas normales"
    lbl.Refresh
    
    'Vemos cual es el primer miercoles del mes
    Horario = Weekday(FInicio, vbMonday)
    '               en funcion del primer dia....
    Select Case Horario
    Case 0
        'Primer dia domingo
        PrimerMiercoles = 4 'primer X dia 4
    Case 1
        PrimerMiercoles = 3
    Case 2
        PrimerMiercoles = 2
    Case 3
        PrimerMiercoles = 1  '1er miercoles dia 1
    Case 4
        PrimerMiercoles = 7
    Case 5
        PrimerMiercoles = 6
    Case Else
        PrimerMiercoles = 5
    End Select
    'Con lo cual el primer miercoles estara en el arrray en la poscion -1
    PrimerMiercoles = PrimerMiercoles - 1
    'Fijo cual es la primera semana del mes, y la utima
    SemanaMesPrimera = Format(FInicio, "ww", vbMonday)
    SemanaMesUltima = Format(FFin, "ww", vbMonday)
     
    'Meto en tmpdatosmes
    
    
    'Incializamos el array
    SQL = "Select tmpDatosMes.*,idHorario,FecAlta,FecBaja,controlnomina,nomtrabajador from tmpDatosMes,Trabajadores"
    SQL = SQL & " WHERE tmpDatosMes.trabajador = Trabajadores.idTrabajador"
    SQL = SQL & " ORDER BY idHorario,idtrabajador"
    Horario = -1
    FESTIVOS = ""
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Depuracion Then FichDepuracion True
    While Not RS.EOF
       
        
        If Horario <> RS!IdHorario Then
            Set vH = Nothing
            Set vH = New CHorarios
            Horario = RS!IdHorario
            FESTIVOS = vH.LeerDiasFestivos(Horario, FInicio, FFin)
            MEDIODIA = vH.LeerMediosDias(Horario, FInicio, FFin)
        End If
    
                
        lbl.Caption = RS!nomtrabajador
        lbl.Refresh
    
    
        
   
    
        Set cT = New CTrabajadorNomina
        'Asignamos los campos
        cT.Codigo = RS!Trabajador
        cT.DiasOficiales = RS!MesDias
        cT.HOficiales = RS!meshoras
        cT.HBolsa = DBLet(RS!bolsaantes, "N")
        cT.HEReales = DBLet(RS!HorasC, "N")   'Para saber si dispongo o NO
        cT.HNReales = DBLet(RS!horasn, "N") - cT.HEReales
        If cT.HNReales < 0 Then
            MsgBox "Horas menor que cero para : " & cT.Codigo & "?"
        End If
        cT.DiasReales_ = DBLet(RS!diasTrabajados, "N")
        If IsNull(RS!FecBaja) Then
             cT.FecBaja = "01/10/2999"
        Else
            cT.FecBaja = RS!FecBaja
        End If
        
        If IsNull(RS!FecAlta) Then
            cT.FecAlta = FInicio
        Else
            cT.FecAlta = RS!FecAlta
        End If

        
        Set cDep2 = Nothing
        Set cDep2 = New Collection
       
        BajaTodoElMes = False
        
'        If RS!Trabajador = 43 Then St op
        SumatorioHorasAntes = cT.HBolsa + cT.HEReales + cT.HNReales
        SumatorioAux = cT.HBolsa
        
        
        ProcesarTrabajadorPicasent cT, FInicio, FFin, vH, PrimerMiercoles, FESTIVOS, BajaTodoElMes
    
        '                           bolsa antes
        SumatorioAux = cT.HBolsa - SumatorioAux + cT.HEReales + cT.HNReales '- cT.HorasCompensadasNomina
        
        
        If SumatorioHorasAntes - SumatorioAux <> 0 Then
           ' MsgBox "Horas antes / despues distintas: " & SumatorioHorasAntes & " / " & SumatorioAux & "(" & cT.Codigo & ")", vbExclamation
        End If
        
        
        'Updateamos los valores
        Dias = cT.DiasCompensables + cT.DiasReales_
        If cT.DiasOficiales < Dias Then
            'FALTA###  ver pk llega hasta aqui
            MsgBox "Error en compensaciones. Exceso dias. Trabajador: " & cT.Codigo & " - " & lbl.Caption, vbExclamation
            Dias = cT.DiasOficiales
        End If

        'Updateamos con los valores calculados
        SQL = "UPDATE tmpDatosMes SET"
        If RS!ControlNomina = 2 Then
            'No puede tener horas en bolsa
            HPaBolsa = 0
        End If
        
        SQL = SQL & "  BolsaDespues =" & TransformaComasPuntos(CStr(cT.HBolsa))
        'SQL = SQL & ", HorasPeriodo = " & TransformaComasPuntos(CStr(cT.HEReales))
        SQL = SQL & ", DiasPeriodo = " & TransformaComasPuntos(CStr(Dias))  ' Dias = cT.DiasCompensables + cT.DiasReales
    
        'Para PICASSENT, machaco los datos
        SQL = SQL & ", HorasN = " & TransformaComasPuntos(CStr(cT.HNReales))
        SQL = SQL & ", HorasC = " & TransformaComasPuntos(CStr(cT.HEReales))
        
        
        'Si los dias va benne
        If Dias >= cT.DiasReales_ And cT.DiasReales_ <= cT.DiasOficiales Then
            SQL = SQL & ", DiasTrabajados = " & cT.DiasReales_
        Else
            
            MsgBox "Dias reales: " & cT.DiasReales_, vbExclamation
        End If
        
        'Reseteo estos campos
        SQL = SQL & ", SaldoDias = " & cT.DiasCompensables  'ahora pongo los dias que le compenso
        SQL = SQL & ", SaldoH = 0"
        
        SQL = SQL & ", Extras = " & TransformaComasPuntos(CStr(cT.HorasCompensadasNomina))
        'Trabajador
        SQL = SQL & " WHERE Trabajador = " & RS!Trabajador
        conn.Execute SQL
        
        
        
        
        
        If BajaTodoElMes Then
            AjustaDatosBajaMesEnteroTrabajador CLng(RS!Trabajador)
        End If
        
        
        'Imprimimos en el fichero
        If cDep2.Count > 0 Then
            cDep2.Add "POST-compensacion:"
            cDep2.Add ""
            cDep2.Add cT.DatosLineaDep
            cDep2.Add vbCrLf
        
            ImprimeFichero cT.Codigo & "   -   " & RS!nomtrabajador
        End If
        'sgi
        RS.MoveNext
    Wend
    RS.Close
    espera 0.5
    
    
    
    
    
    'AHora obtenemos los anticpos en NOMINA
    '-----------------------------------------
    
    Set RtAnticipos = New ADODB.Recordset
    SQL = "SELECT Trabajador,tmpDatosMEs.HorasN, tmpDatosMEs.extras, Categorias.Importe1, Categorias.Importe2, Trabajadores.PorcSS, Trabajadores.PorcIRPF,embargado"
    SQL = SQL & " FROM tmpDatosMEs INNER JOIN (Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria) ON tmpDatosMEs.Trabajador = Trabajadores.IdTrabajador"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF

        'Ajustes ponemos HN las que tiene menos las que sean extra
        lbl.Caption = "Anticipo: " & RS!Trabajador
        lbl.Refresh

        
        'Los anticipos llevan toooodos los importes
        HorasOf = (RS!horasn * RS!Importe1) + (RS!extras * RS!Importe2)
        'Quitamos IRPF y SS
        H = (HorasOf * RS!PorcIRPF) + (HorasOf * RS!PorcSS)
        H = Round((H / 100), 2)
        HorasOf = HorasOf - H
        
        'sql=devuelvedesdebd("sum(
        SQL = "Select sum(importe) from pagos where trabajador = " & RS!Trabajador
        'SQL = SQL & " AND tipo <=1 " 'pago o anticipo nomina
        SQL = SQL & " AND Fecha >= #" & Format(FInicio, FormatoFecha) & "#"
        SQL = SQL & " AND Fecha <= #" & Format(FFin, FormatoFecha) & "#"
        RtAnticipos.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SumatorioAux = 0
        If Not RtAnticipos.EOF Then
            If Not IsNull(RtAnticipos.Fields(0)) Then SumatorioAux = RtAnticipos.Fields(0)
        End If
        RtAnticipos.Close
            
        SumatorioHorasAntes = SumatorioAux - HorasOf
        If Abs(SumatorioHorasAntes) > 5 Then
            
        End If
        
        SQL = "UPDATE tmpDatosMes SET"
        SQL = SQL & " Anticipos = " & TransformaComasPuntos(CStr(HorasOf))
        'Trabajador
        SQL = SQL & " WHERE Trabajador = " & RS!Trabajador
    
        conn.Execute SQL
    
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    Set RtAnticipos = Nothing
    
    
    
    
    
    Set RS = Nothing
    If MiNF > 0 Then
        FichDepuracion False
        'Lanzar notepad
        LanzaNotepad
    End If
End Sub

Private Sub LanzaNotepad()
    On Error Resume Next
    Shell "NOTEPAD.EXE " & App.Path & "\Depu.txt"
    Err.Clear
End Sub


Private Sub ProcesarTrabajadorPicasent(ByRef cT As CTrabajadorNomina, FI As Date, FF As Date, ByRef vHor As CHorarios, PrimerMiercoles As Integer, FESTIVOS As String, ByRef BajaTodoElMes2 As Boolean)
Dim vMEs(30) As cDiaProcesado
Dim vM As cDiaProcesado
Dim i As Integer
Dim J As Integer
Dim N As Integer
Dim HN As Currency
Dim HC As Currency
Dim RF As ADODB.Recordset
Dim BAJAS As String
Dim NumeroDiasMes As Integer
Dim cad As String
Dim FechaReferencia As Date
Dim FinalCompensar As Boolean
Dim UltDiaSemana As Integer
Dim ElOtroProcesado As Boolean
Dim Dia As Integer
Dim vM2 As cDiaProcesado
Dim VariablesDepuracion_Dias As String
Dim TrabajadorDepuracion As Integer
Dim NUmeroSemanasMes As Currency
Dim DiasPartidos As Integer
Dim HorPart As Currency
Dim MiercolesDeSemanaProcesando2 As Integer
Dim SabadoSemanaProcesando2 As Integer

Dim HorasCambiadasMierSab As Currency

Dim CadenaDiasMes As String
    TrabajadorDepuracion = 0
  
    
  
    If Depuracion Then
        BAJAS = "|" & cT.Codigo & "|"
        If InStr(1, TrabajadoresDepuracion, BAJAS) > 0 Then TrabajadorDepuracion = cT.Codigo
        
        If TrabajadorDepuracion > 0 Then
            Debug.Print ""
            'St op
        End If
    End If
    
    
    
    
    J = Weekday(FI, vbMonday)
    NumeroDiasMes = DiasMes(Month(FI), Year(FF)) - 1
    BAJAS = "/" & Month(FI) & "/" & Year(FF)
    For i = 0 To NumeroDiasMes
        Set vMEs(i) = New cDiaProcesado
        vMEs(i).DiaProcesable = True
        vMEs(i).DiaSemana = (J + i) Mod 7  'L,M,X,J,
        'Debug.Print I & ": " & I + 1 & "  " & vMes(I).DiaSemana
        vMEs(i).Festivo = vMEs(i).DiaSemana = 0   'Los domingos son festivos SEGURO
        cad = Format(i + 1 & BAJAS, "dd/mm/yyyy")
        If InStr(1, FESTIVOS, cad) > 0 Then
            
            vMEs(i).Festivo = True
        End If
        
        vMEs(i).NumeroSemana = CInt(Format(i + 1 & "/" & Month(FI) & "/" & Year(FF), "ww"))
    Next
    Set RF = New ADODB.Recordset
    
    If cT.FecAlta > FI Then
        FechaReferencia = FI
        While FechaReferencia < cT.FecAlta
            i = Day(FechaReferencia) - 1
            vMEs(i).DiaProcesable = False
            FechaReferencia = DateAdd("d", 1, FechaReferencia)
        Wend
    End If
    
    
   
    If cT.FecBaja <= FF Then
        FechaReferencia = FF
        While FechaReferencia > cT.FecBaja
            i = Day(FechaReferencia) - 1
            vMEs(i).DiaProcesable = False
            FechaReferencia = DateAdd("d", -1, FechaReferencia)
        Wend
    End If
    
    'Las bajas
    
    cad = "Select * from tmpcombinada where idtrabajador=" & cT.Codigo
    RF.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    BAJAS = ""
    If Not RF.EOF Then
        
      
        BAJAS = "|"
        While Not RF.EOF
            FechaReferencia = RF!Fecha
            Do
                i = Day(FechaReferencia) - 1
                vMEs(i).DiaProcesable = True
                vMEs(i).Baja = True
                BAJAS = BAJAS & Format(FechaReferencia, "dd/mm/yyyy") & "|"
                FechaReferencia = DateAdd("d", 1, FechaReferencia)
            Loop Until FechaReferencia > RF!H1
            RF.MoveNext
        Wend
    End If
    RF.Close



    cad = "Select * from bajas where idtrab=" & cT.Codigo & " AND fechabaja <= #" & Format(FI, FormatoFecha) & "# "
    cad = cad & " AND ( fechaalta is null )"
    RF.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not RF.EOF Then
        'Trabajador de BAJA
        BajaTodoElMes2 = True
        If TrabajadorDepuracion > 0 Then cDep2.Add "De baja"
        RF.Close
        Exit Sub
    End If
    RF.Close




    cad = "Select fecha,horastrabajadas,horasincid from marcajes"
    cad = cad & " WHERE Fecha >=#" & Format(FI, FormatoFecha) & "#"
    cad = cad & " AND Fecha <=#" & Format(FF, FormatoFecha) & "#"
    cad = cad & " AND idTrabajador = " & cT.Codigo
    cad = cad & " ORDER BY fecha"
    RF.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText

    

    If RF.EOF Then
        If TrabajadorDepuracion > 0 Then cDep2.Add "No tiene datos"
        'St op  FALTA###
        'Exit Sub
    End If
    
    'Pasada CERO.
    '-------------------------------------------------------------------------------
    'Trabaja dias de baja /festivos
    '
    'Y miercoles o sabados trabja a horas cambiadas (mañana tarde)

    While Not RF.EOF
        N = Day(RF!Fecha) - 1
        'Debug.Print RF!Fecha & ": " & N & "  " & vMes(N).DiaSemana
        If Not vMEs(N).DiaProcesable Then
            'HA trabajado un dia que no deberia haber trabajado
            MsgBox "Trabaja dia no procesable: " & RF!Fecha & "   Trab: " & cT.Codigo, vbCritical
           ' End
        End If
            
        
        HN = RF!HorasTrabajadas
        i = 0
        If vMEs(N).DiaSemana = 3 Then
            If HN > 3.5 Then i = 1
        ElseIf vMEs(N).DiaSemana = 6 Then
            If HN > 4.5 Then i = 1
        ElseIf HN > 8 Then i = 1
        End If
        
        
        If vMEs(N).Festivo Then i = 1 'Ha trabajado el festivo. TOdas son EXTRA pero la suma de horas del mes YA estabien, el update horasn ya lo recoge
        
        'Si es dia festivo..
        'If N = 5 Then St op
        vMEs(N).DiaNomina = 1
        vMEs(N).SabadoSiHabiaTrabajado = True  'Me da lo mismo que no sea sabado
        
        If i = 1 Then 'Horas extra
            
            HC = RF!HorasIncid
            HN = HN - HC
        Else
           
            HC = 0
        End If
        
        vMEs(N).HE_Reales = HC
        vMEs(N).HT_Reales = HN
        
        If vMEs(N).Baja Or vMEs(N).Festivo Then
            'Ha trabajado el dia de la baja  o festivo
            'If vMes(N).Baja Then St op
            
            
            'Todas las horas deberian ser Ex
            If vMEs(N).Baja Then
                If vMEs(N).HT_Reales > 0 Then
                    HN = vMEs(N).HT_Reales  'Guardo las T
                    cad = "Trabaja baja : " & N + 1 & "   HT: " & vMEs(N).HT_Reales & " /" & vMEs(N).HE_Reales
                    cad = cad & vbCrLf & "       HComp(antes/despues)=" & cT.HEReales & "/"
                    cT.HEReales = cT.HEReales + HN   'le sumo a las comepnsables del mes las de este dia
                    cad = cad & cT.HEReales
                    cad = cad & vbCrLf & "       Trabajadas(antes/despues)=" & cT.HNReales & "/"
                    cT.HNReales = cT.HNReales - HN
                    cad = cad & cT.HNReales
    
                    
                    vMEs(N).HE_Reales = vMEs(N).HT_Reales + vMEs(N).HE_Reales
                    vMEs(N).HT_Reales = 0  'tot reales=0
                Else
                    
                    cad = "Trabaja baja. NO tiene horas. Dia : " & N + 1
                    
                End If
            Else
                'Es festivo. Lo indicamos en el depur
                cad = "Trabaja festivo.  " & N + 1
            End If
            If TrabajadorDepuracion > 0 Then cDep2.Add cad
            vMEs(N).DiaNomina = 0
            vMEs(N).SabadoSiHabiaTrabajado = False
        End If
        
            
        RF.MoveNext
        'RF.MoveLast
    Wend
        
        
    'Compensacion de horas por dia que no llegan al minimo
    i = CInt(Format(FI, "ww"))
    J = CInt(Format(FF, "ww"))
    NUmeroSemanasMes = J - i + 1
    
    'Variables
    HC = cT.HEReales + cT.HBolsa  'Las extra del mes mas la bolsa
    
    If TrabajadorDepuracion > 0 Then
        cDep2.Add "Pre-compensacion:"
        cDep2.Add ""
        cDep2.Add cT.DatosLineaDep
        cDep2.Add vbCrLf
        cDep2.Add vbCrLf
    End If
    
    
    Dia = 0
    VariablesDepuracion_Dias = ""
    
    
    'Para todas las senamas del mes
    For N = 1 To NUmeroSemanasMes
        cad = ""  'para depurar
        DiasPartidos = 0
        HorPart = 0 'ht para mierco y sabado
    
    
        'Procesamos por semanas
        '-------------------------------------
        If N = 1 Then
            'Primera semana. Veremos el ultimo dia semana cual es
            UltDiaSemana = vMEs(0).DiaSemana
            If UltDiaSemana <> 0 Then
                UltDiaSemana = 8 - UltDiaSemana
                UltDiaSemana = UltDiaSemana - 1 'pk es un aarray desde el CERO
            End If
                        
        Else
            UltDiaSemana = UltDiaSemana + 7
            If UltDiaSemana > NumeroDiasMes Then UltDiaSemana = NumeroDiasMes
        End If
        
        
        'Si tiene miercoles y sabado miro sus horas. Igual pasan todas a hextras
        MiercolesDeSemanaProcesando2 = -1
        SabadoSemanaProcesando2 = -1
        For i = Dia To UltDiaSemana
            If vMEs(i).DiaProcesable And Not vMEs(i).Baja And Not vMEs(i).Festivo Then
                
                If vMEs(i).DiaSemana = 3 Or vMEs(i).DiaSemana = 6 Then
                    
                    'MIERCOLES SABADO
                    If vMEs(i).DiaSemana = 3 Then
                        MiercolesDeSemanaProcesando2 = i
                    Else
                        SabadoSemanaProcesando2 = i
                    End If
                    HorasCambiadasMierSab = 0
                    cad = ""
                    FijarHorasMiercolesSabadoPicassent cT, vMEs(i), CDate(CStr(i + 1) & Format(FI, "/mm/yyyy")), cad, HorasCambiadasMierSab
                    
                    If vMEs(i).HT_Reales > 0 Then
                        HorPart = HorPart + vMEs(i).HT_Reales
                        DiasPartidos = 1
                    Else
                        'Horas trabajadas=0
                        vMEs(i).DiaNomina = 0
                        vMEs(i).SabadoSiHabiaTrabajado = False  'Todas son compensables
                    End If
                    If HorasCambiadasMierSab <> 0 Then
                        
                        
                        'Luego tenemos esas horas de mas o menos en HC, que es la que lleva Bolsa+HE
                        HC = HC + HorasCambiadasMierSab
                        cT.HNReales = cT.HNReales - HorasCambiadasMierSab
                    End If
                    
                    If TrabajadorDepuracion > 0 Then VariablesDepuracion_Dias = VariablesDepuracion_Dias & cad
                End If
            End If
        Next i
    
         
        'Ahora ya vamos a ver si necesitan horas extra  para completar
        cad = ""
        ElOtroProcesado = True
        
        If Dia < UltDiaSemana Then
           If UltDiaSemana - Dia < 6 Then
                'Semanas incompletas
                If vMEs(Dia).DiaSemana = 0 Then ElOtroProcesado = False
            End If
        End If
        If HC <= 0 Then ElOtroProcesado = False 'NO tengo suficiente
        
        'Significa que tengo que ver la dupla XyS
        
     
        If ElOtroProcesado Then
            If DiasPartidos = 0 Then
                'Nada. NO ha trabajado la dupla
            Else
               
                If HorPart < 3.5 Then
                    HN = 3.5 - HorPart
                    HC = HC - HN
                    cad = "XyS semana : " & Dia + 1 & " - " & UltDiaSemana + 1 & "      Falta " & HN & vbCrLf
                    cT.HorasCompensadasNomina = cT.HorasCompensadasNomina + HN 'Cuantas compenso en nomina
                Else
                    'Ha trbajado sificiente
                    If MiercolesDeSemanaProcesando2 >= 0 And SabadoSemanaProcesando2 >= 0 Then
                        'Si ha trabajado los dos dias, suficiente, uno de los dos lo pongo a CERO, para el conteo posterior
                        If vMEs(MiercolesDeSemanaProcesando2).DiaNomina = 1 And vMEs(SabadoSemanaProcesando2).DiaNomina = 1 Then
                            vMEs(SabadoSemanaProcesando2).DiaNomina = 0
                            'AQUI NO TOCO EL SABADO
                            '|||||
                        End If
                    End If
                End If
            End If
        End If
        
        
        For i = Dia To UltDiaSemana
            
     
                If vMEs(i).DiaProcesable And Not vMEs(i).Baja And Not vMEs(i).Festivo Then
                    If vMEs(i).DiaSemana <> 3 And vMEs(i).DiaSemana <> 6 Then
                        If vMEs(i).DiaNomina = 1 And vMEs(i).HT_Reales < 3.5 Then
                            Dim AuxH As Currency
                            'No ha llegado a 3.5 trabajadas
                            AuxH = HC + vMEs(i).HT_Reales
                            If AuxH >= 3.5 Then
                                'Utilizamos horas para compensar
                                HN = 3.5 - vMEs(i).HT_Reales
                                HC = HC - HN
                                cad = cad & "Dia : " & i + 1 & " Falta " & HN & vbCrLf
                                
                                
                                cT.HorasCompensadasNomina = cT.HorasCompensadasNomina + HN
                            Else
                                'NO tiene bastante para compensar
                                HN = vMEs(i).HE_Reales + vMEs(i).HT_Reales
                                
                                cad = cad & "Dia : " & i + 1 & "   a compensar.  " & HN & ". Antes " & vMEs(i).HE_Reales
                                
                                vMEs(i).HT_Reales = 0
                                vMEs(i).HE_Reales = HN
                                vMEs(i).DiaNomina = 0
                                'Quito las del dia
                                cT.HNReales = cT.HNReales - HN
                                cT.HEReales = cT.HEReales + HN
                                'cT.HorasCompensadasNomina = cT.HorasCompensadasNomina + HN
                                'LAs compensables
                                cad = cad & "       Variable(HC):" & HC & " --> " & HC + HN & vbCrLf
                                HC = HC + HN
                            End If
                        End If
                    End If
                End If
            
        Next i
        If TrabajadorDepuracion > 0 And cad <> "" Then cDep2.Add cad
        Dia = UltDiaSemana + 1
        If Dia > 30 Then N = 9 'QUe se salga
    Next
    
    
    
    
    'Vemos dias trabajados que le han quedado
   'If cT.Codigo = 4 Then St op
   J = 0
   HN = 0
   cad = ""
   N = vMEs(0).DiaSemana
   If N = 1 Then
        N = 7
   Else
    If N = 0 Then
        N = 1
         cad = Space(6 * 8)
    Else
        N = 7 - vMEs(0).DiaSemana  'ult dia semana
        N = N + 1
        cad = Space(CInt((vMEs(0).DiaSemana) - 1) * 8)
    End If
   
   End If
  ''' If TrabajadorDepuracion > 0 Then St op
   
   For i = 0 To NumeroDiasMes
        If i = N Then
            N = i + 7
            cad = cad & vbCrLf
        End If
        If vMEs(i).DiaProcesable And Not vMEs(i).Baja And Not vMEs(i).Festivo Then
            FinalCompensar = True
            
            J = J + vMEs(i).DiaNomina
            'HN = HN + vMes(i).HT_Reales
        Else
            FinalCompensar = False
            
        End If
        cad = cad & PintaDia(i, vMEs(i), FinalCompensar)
        
    Next i
    CadenaDiasMes = ""
    If TrabajadorDepuracion > 0 Then CadenaDiasMes = cad

    
    
    
    If cT.DiasReales_ > J Then
        If TrabajadorDepuracion > 0 Then
            cDep2.Add ""
            cDep2.Add "Reajuste dias trabajados(reales/ajustados):" & cT.DiasReales_ & " / " & J
            cDep2.Add ""
        End If
        cT.DiasReales_ = J
    End If
    
    
    
    J = cT.DiasOficiales - J
    If J < 0 Then
        If TrabajadorDepuracion > 0 Then
            cad = "Dias oficiales= " & cT.DiasOficiales & "  Trabajados: " & cT.DiasOficiales + J
            cDep2.Add cad
        End If
        J = 0
    End If
    
    cad = ""
    Dia = 0
    N = 0
    
    FinalCompensar = False
    While Not FinalCompensar
            If Not FinalCompensar Then FinalCompensar = J = 0 Or HC < 3.5
            If Not FinalCompensar Then
                    'Proceso una semana
                    '--------------------------------
                    If Dia <= NumeroDiasMes Then
                        If Dia = 0 Then   '1ª a procesar
                            UltDiaSemana = vMEs(0).DiaSemana
                            If UltDiaSemana <> 0 Then
                                UltDiaSemana = 8 - UltDiaSemana
                                UltDiaSemana = UltDiaSemana - 1 'pk es un aarray desde el CERO
                            End If
                        Else
                            UltDiaSemana = UltDiaSemana + 7
                            If UltDiaSemana > NumeroDiasMes Then UltDiaSemana = NumeroDiasMes
                        End If
                        
                        'Proceso una semana ENTERA. Primero X y Saba
                        'Despues el resto de dias
                        '------------------------------------------------------------------
                        Set vM = Nothing
                        If vMEs(Dia).DiaSemana > 3 Then
                            'NO tiene Miercoles la semana
                            '
                        Else
                            If UltDiaSemana = NumeroDiasMes Then
                                'Ultima semana
                                i = Dia + 2
                            Else
                                'Semanas intemedias
                                i = UltDiaSemana - 4
                            End If
                            If i > Dia Then
                                If i <= NumeroDiasMes Then Set vM = vMEs(i)
                            Else
                                If i = 0 Then Set vM = vMEs(i)
                            End If
                        End If
                        If UltDiaSemana = NumeroDiasMes Then
                            'Ultima semana del mes
                            If vMEs(UltDiaSemana).DiaSemana < 6 Then
                                Set vM2 = Nothing
                            End If
                        Else
                            
                            i = UltDiaSemana - 1
                            If i < 0 Then
                                'domingo es primer y utlimo dia semana
                                Set vM2 = Nothing
                            Else
                                Set vM2 = vMEs(i)
                            End If
                        End If
                        
                        'Debug.Print "Mierc: (" & vM.DiaSemana & ")      Saba: " & vM2.DiaSemana
                        If Not (vM Is Nothing) Or Not (vM2 Is Nothing) Then
                            HN = HayQueCompensarMiercolesSabado(vM, vM2, HC)
                        Else
                            HN = 0
                        End If
                        
                        If HN > 0 Then
                            
                            
                            J = J - 1
                            cT.DiasCompensables = cT.DiasCompensables + 1
                            'Mayo 2012
                            'Existe la posibilidad que sea <0
                            If HC - HN < 0 Then
                                cad = "         <0.  Quedan:" & HC & "  Necesito:" & HN
                                
                                cT.HorasCompensadasNomina = cT.HorasCompensadasNomina + HC
                                HC = 0
                            Else
                                cad = ""
                                HC = HC - HN
                                cT.HorasCompensadasNomina = cT.HorasCompensadasNomina + HN
                            End If
                            VariablesDepuracion_Dias = VariablesDepuracion_Dias & " Compensa dia xs " & Dia + 1 & " al " & Dia + 7 & " (" & HN & ")" & "     #" & HC & cad & vbCrLf
                        End If
                    End If
                    
                    For i = Dia To UltDiaSemana
                        If J > 0 And HC >= 8 Then
                            If vMEs(i).DiaProcesable Then
                                If Not (vMEs(i).Festivo Or vMEs(i).Baja) Then
                                    If vMEs(i).DiaNomina = 0 Then
                                        If vMEs(i).DiaSemana <> 3 And vMEs(i).DiaSemana <> 6 Then
                                            'DIA NORMAL NO TRABAJADO
                                            
                                            J = J - 1
                                            HC = HC - 8
                                            cT.HorasCompensadasNomina = cT.HorasCompensadasNomina + 8
                                            cT.DiasCompensables = cT.DiasCompensables + 1
                                            
                                            VariablesDepuracion_Dias = VariablesDepuracion_Dias & " Compensamos " & i + 1 & "                    #" & HC & vbCrLf   'Para insertar en la variable
                                            
                                        End If
                                    End If 'dianomina=0
                                End If 'no de baja o festivo
                            End If 'diaprocesable
                        End If 'suficiente para compensar
                    Next i ' Siguiente dia
                    Dia = i  'para el siguiente bucle
                    'Ya hemos procesado todo el mes
                    If Dia > NumeroDiasMes Then FinalCompensar = True
            End If
            
        
    Wend
    
    'Si llegamos aqui y tiene dias a compensar y mas de cuatro horas y menos de
    If J > 0 And HC >= 4 Then
        cT.DiasCompensables = cT.DiasCompensables + 1
        If HC > 8 Then
            cT.HorasCompensadasNomina = cT.HorasCompensadasNomina + 8
            'EXTRAÑO
            MsgBox "Dias a compensar teniendo mas de 8 horas a compensar", vbExclamation
            HN = 8
            HC = HC - 8
        Else
            cT.HorasCompensadasNomina = cT.HorasCompensadasNomina + HC
            HN = HC
            HC = 0 'Todas las he llavado
        End If
        VariablesDepuracion_Dias = VariablesDepuracion_Dias & " Horas de sobra y puede comensar (" & HN & ")" & vbCrLf 'Para insertar en la variable
        
    End If
    
    
    
    
   If cT.DiasCompensables > 0 And TrabajadorDepuracion > 0 Then
       cad = ""
       N = vMEs(0).DiaSemana
       If N = 1 Then
            N = 7
       Else
        If N = 0 Then
            N = 1
             cad = Space(6 * 8)
        Else
            N = 7 - vMEs(0).DiaSemana  'ult dia semana
            N = N + 1
            cad = Space(CInt((vMEs(0).DiaSemana) - 1) * 8)
        End If
       
       End If
       For i = 0 To NumeroDiasMes
            If i = N Then
                N = i + 7
                cad = cad & vbCrLf
            End If
            If vMEs(i).DiaProcesable And Not vMEs(i).Baja And Not vMEs(i).Festivo Then
                FinalCompensar = True
    
                J = J + vMEs(i).DiaNomina
                HN = HN + vMEs(i).HT_Reales
            Else
                FinalCompensar = False
            End If
            cad = cad & PintaDia(i, vMEs(i), FinalCompensar)
            
        Next i
    
        cDep2.Add ""
        cDep2.Add CadenaDiasMes
        cDep2.Add vbCrLf & vbCrLf
        cDep2.Add cad
    End If
    
    
    
    'Cuanta bolsa queda
    cT.HBolsa = HC
    
    If TrabajadorDepuracion > 0 Then
        For i = 1 To 3
            cDep2.Add ""
        Next
        If VariablesDepuracion_Dias <> "" Then
            cDep2.Add "Proceso compensacion"
            cDep2.Add "Dia.       Horas"
            cDep2.Add "------------------------------------------"
            cDep2.Add VariablesDepuracion_Dias
        
            For i = 1 To 3
                cDep2.Add ""
            Next
        End If

        

    End If

    
    For i = 0 To NumeroDiasMes
        Set vMEs(i) = Nothing
    Next
End Sub




Private Function PintaDia(Indice As Integer, ByRef vn As cDiaProcesado, Deberiatrabajar As Boolean) As String
    PintaDia = Right(" " & Indice + 1, 2)
    If Deberiatrabajar Then
        PintaDia = PintaDia & " S"
        If vn.DiaNomina = 0 Then
            PintaDia = PintaDia & "-"
        Else
            PintaDia = PintaDia & " "
        End If
        
    Else
        PintaDia = PintaDia & " n"
        If vn.DiaNomina = 1 Then
            PintaDia = PintaDia & "%"
        Else
            PintaDia = PintaDia & " "
        End If
        
    End If
    PintaDia = PintaDia & "   "

End Function


Private Function HayQueCompensarMiercolesSabado(ByRef Mier As cDiaProcesado, ByRef Sab As cDiaProcesado, DeCuantasHorasCompensablesDispongo As Currency) As Currency
Dim TrabajaSabado As Boolean
Dim TrabajaMiercoles As Boolean
Dim TieneTrabajarX As Boolean
Dim TieneTrabajarS As Boolean
Dim bol As Boolean

    HayQueCompensarMiercolesSabado = 0
    If Mier Is Nothing Then
        TieneTrabajarX = False
    Else
        If Mier.Baja Then
            TieneTrabajarX = False
        Else
            If Mier.DiaProcesable And Not Mier.Festivo Then
                TieneTrabajarX = True
                'If Mier.DiaNomina = 0 And Mier.HE_Reales = 0 And Mier.HT_Reales = 0 Then
                If Mier.DiaNomina = 0 Then
                    TrabajaMiercoles = False
                Else
                    TrabajaMiercoles = True
                End If
            Else
                TieneTrabajarX = False
            End If
        End If
    End If
    If Sab Is Nothing Then
        TieneTrabajarS = False
    Else
        If Sab.Baja Then
            TieneTrabajarS = False
        Else
            If Sab.DiaProcesable And Not Sab.Festivo Then
                TieneTrabajarS = True
                'If Sab.DiaNomina = 0 And Sab.HE_Reales = 0 And Sab.HT_Reales = 0 Then
                bol = True
                If Sab.DiaNomina = 0 Then
                    If Not Sab.SabadoSiHabiaTrabajado Then bol = False
                End If
                If Not bol Then
                    TrabajaSabado = False
                Else
                    TrabajaSabado = True
                End If
            Else
                TieneTrabajarS = False
            End If
        End If
    End If
    'No tenia que trabajar NINGUNO de los dos
    If Not TieneTrabajarS And Not TieneTrabajarX Then Exit Function
    'Ha trabajado alguno de los dos
    If TrabajaSabado Or TrabajaMiercoles Then Exit Function
    
    
    'No ha trabajado NINGUNO de los dos dias
    'Y tenia que haber trabajado alguno de ellos
    If TieneTrabajarS And TieneTrabajarX Then
        HayQueCompensarMiercolesSabado = 8 'OCHO horas a compensar
    Else
        If TieneTrabajarS Then
            If DeCuantasHorasCompensablesDispongo >= 4.5 Then HayQueCompensarMiercolesSabado = 4.5
        Else
            If DeCuantasHorasCompensablesDispongo >= 3.5 Then HayQueCompensarMiercolesSabado = 3.5
        End If
    End If
End Function


Private Sub FijarHorasMiercolesSabadoPicassent(ByRef cTr As CTrabajadorNomina, ByRef DN As cDiaProcesado, Dia As Date, ByRef Depurac As String, ByRef CambiaHoras As Currency)
Dim RH As ADODB.Recordset
Dim C As String
Dim T1 As Currency
Dim T2 As Currency
Dim E As Boolean
Dim Seguir As Boolean
Dim NuevaHora As Currency
Dim DeMiercoles As Boolean
Dim C2 As String
    If cTr.Codigo >= 700 Then Exit Sub

    Set RH = New ADODB.Recordset
    
    C = "Select * from entradamarcajes where idtrabajador=" & cTr.Codigo
    C = C & " AND fecha = #" & Format(Dia, FormatoFecha) & "# "
    

    C = C & " ORDER BY hora"
    DeMiercoles = DN.DiaSemana = 3
    RH.Open C, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    Seguir = True
    'Si tiene algun --> Miercoles y es por la tarde ---->NORMAL, no hago nada
    '                     sabado   y es por la mañana ---> "
    If Not RH.EOF Then
        If DeMiercoles Then
       
            If RH!Hora > CDate(HoraIntermediaMiercolesSabado) Then
                Seguir = False
            Else
                'RH.MoveFirst
                
            End If
        Else
            RH.MoveLast
           'sabado
            If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                Seguir = False
            Else
                RH.MoveFirst
            End If
            
        End If
    End If
    
    
    If Not Seguir Then
        
        If DeMiercoles Then
            'Todas las horas que ha trabajado son por la tarde
            
        Else
            'Todas las horas que ha trabajado son por la mañana
            
            
            
        End If
'        'Son todas de mañana o de tarde
'        If DN.HT_Reales < 3.5 Then
'            'Ho ha llegado a trabajar las 3.5 horas
'
'            DN.HE_Reales = DN.HT_Reales
'            DN.HT_Reales = 0
'            DN.DiaNomina = 0
'        Else
        DN.DiaNomina = 1
   

        RH.Close
        Exit Sub
    End If
    E = True
    NuevaHora = 0

    T1 = 0
    C2 = Day(Dia) & "    "
    While Not RH.EOF
        C2 = C2 & RH!Hora & "     "
        If E Then
            If DeMiercoles Then
                If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                    T1 = CCur(DevuelveValorHora(RH!Hora))
                Else
                    'Ya es cuando le toca
                    RH.MoveLast
                End If
                
             Else
                If RH!Hora >= CDate(HoraIntermediaMiercolesSabado) Then
                    T1 = CCur(DevuelveValorHora(RH!Hora))
                Else
                    'Ya es cuando le toca
                  
                End If
             End If
                
        Else
            'Si tiene valor t1 calculamos dif
            If T1 > 0 Then
                T2 = CCur(DevuelveValorHora(RH!Hora))
                T1 = T2 - T1
            
                NuevaHora = NuevaHora + T1
                T1 = 0

            End If
        End If
        E = Not E
        RH.MoveNext
    Wend
    
    


   'veremos si este dia realmente es como si no lo trabajadra
    'ya que si tenia que haber venido por la mañana, pero solo viene por
    'la tarde a efectos de nomina no lo cuento
    



    RH.Close
    If NuevaHora > 0 Then
                
                'TOTAL HORAS TRABAJADAS
                T1 = DN.HE_Reales + DN.HT_Reales
                DN.DiaNomina = 1
                    If NuevaHora > DN.HE_Reales Then
                        C2 = C2 & vbCrLf
                        C2 = C2 & "    Real: " & DN.HT_Reales & " / " & DN.HE_Reales & vbCrLf
                        
    
                        If T1 >= NuevaHora Then
                        
                            'Las extra
                            T2 = NuevaHora - DN.HE_Reales
                            cTr.HEReales = cTr.HEReales + T2
                            DN.HE_Reales = NuevaHora   'Nuevas horas extra
                            
      
                            DN.HT_Reales = DN.HT_Reales - T2
                            If DN.HT_Reales < 0 Then MsgBox "Error revisndo dia(" & Dia & ")   HTra < 0"

                            CambiaHoras = T2
                        Else
                           ' Stop
                        End If
                        C2 = C2 & "    Ahora: " & DN.HT_Reales & " / " & DN.HE_Reales & "          Dif: " & T2
                        Depurac = CStr(C2) & vbCrLf
                        
                    Else
                        If NuevaHora < DN.HE_Reales Then
                            If DN.HT_Reales < 3.5 Then MsgBox "Avisa a david ;) " & Dia & "    Trab: " & cTr.Codigo
                        End If
                    End If
         
         
                    
    Else
       If DN.HE_Reales + DN.HT_Reales > 0 Then
            If DN.DiaNomina = 0 Then MsgBox "Dia nomina=0"
       
            DN.DiaNomina = 1
            
        End If
    End If

    

End Sub


Private Sub FichDepuracion(abrir As Boolean)
    On Error Resume Next
    If abrir Then
        MiNF = FreeFile
        Open App.Path & "\depu.txt" For Output As #MiNF
        
    Else
        Close #MiNF
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Sub ImprimeFichero(Trabajador As String)
Dim i As Integer
    On Error GoTo EI
    Print #MiNF, "Trabajador: " & Trabajador
    Print #MiNF, "": Print #MiNF, ""
    For i = 1 To cDep2.Count
        Print #MiNF, cDep2.Item(i)
    Next i
    
    'Damos espacios
    Print #MiNF, String(70, "*")
    Print #MiNF, String(70, "*")
    Print #MiNF, String(70, "*")
    For i = 1 To 5
        Print #MiNF, ""
    Next i
EI:
    Err.Clear
End Sub





Public Function DiasLaborablesSemana(Horario As Integer) As Integer
Dim SQL As String
Dim RS As ADODB.Recordset
    SQL = "SELECT Count(*) From SubHorarios Where SubHorarios.IdHorario = " & Horario & " And SubHorarios.Festivo = False"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        DiasLaborablesSemana = -1
    Else
        DiasLaborablesSemana = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    Set RS = Nothing
End Function





