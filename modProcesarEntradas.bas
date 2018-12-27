Attribute VB_Name = "modProcesarEntradas"
Option Explicit



Private idIncidenciaGenerada As Long

Public PrimerTicaje As Date
Public UltimoTicaje As Date

Dim QuitarAlmuerzo
Dim quitarmerienda
Dim TotalParadas2 As Currency

Dim vSQL2 As String


Public Sub ProcesarEntradasFichajes(Fecha As Date, ByRef lblPpal As Label, ByRef lblDetall As Label)
'Public Sub ProcesarEntradasFichajes(Fecha As Date, IdTra As Long, ByRef lblPpal As Label, ByRef lblDetall As Label)
Dim SQL As String
Dim RsH As ADODB.Recordset
Dim vH As CHorarios

'Noviembre 2018
'El redondeo de ajustes se hará:
'   ó porque esta selecionado en el horario
' o pq aun teniendo ajustes de otro tipo, tiene ajustes puestos
Dim HacerRedondeoDeAjustes As Boolean

    'En el lbl podremos cambiar los textos

    'MOTANMOS UNA SELECT PARA SABER LOS DITINTOS HORARIOS
    'QUE HAY EN EL TOTAL A PROCESAR, PARA NO TENER QUE IR LEYENDO
    'LOS HORARIOS SALTEADOS
    If vEmpresa.CreaCalDiariaTra Then
    
        SQL = "select distinct(idhorario) from entradafichajes,calendariot"
        SQL = SQL & " where entradafichajes.idtrabajador=calendariot.idtrabajador"
        SQL = SQL & " and entradafichajes.fecha=calendariot.fecha "
        SQL = SQL & " and entradafichajes.fecha='" & Format(Fecha, FormatoFecha) & "'"
        
        
    Else
        'ALZIRA - CATADAU
        SQL = "select distinct(idhorario) from entradafichajes,calendariol ,trabajadores where"
        SQL = SQL & " entradafichajes.idtrabajador=trabajadores.idtrabajador and trabajadores.idcal=calendariol.idcal"
        SQL = SQL & " and entradafichajes.fecha=calendariol.fecha"
        SQL = SQL & " and entradafichajes.fecha='" & Format(Fecha, FormatoFecha) & "'"
    End If
        
    'Para el trabajador
    'If IdTra > 0 Then SQL = SQL & " and entradafichajes.idtrabajador = " & IdTra
        
    
    Set RsH = New ADODB.Recordset
    RsH.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set vH = New CHorarios
    While Not RsH.EOF
        If vH.IdHorario <> RsH!IdHorario Then
            If vH.Leer(RsH!IdHorario, Fecha, 0) = 1 Then
                'ERROR LEYENDO HORARIO
                MsgBox "leyendo horario"
                Stop
                
            End If
        End If
        
        'Para cada horario hacemos rectificados
        '---------------------------------------------------
        
        '  vbRecAjustes = 1    'Por ajustes Lleva ajustes
        '  vbRecNormCuarto = 2   'normal, cuarto hora
        '  vbRecNormMedia = 3    'normal  media hora
        '  vbRecESCuarto = 4     ' Entrada/Salida cuarto
        '  vbRecESMedia = 5      'E/S 30'
        HacerRedondeoDeAjustes = False
        If vH.Rectificar = 1 Then
            'SOLO tiene redondeo por ajustes
            HacerRedondeoDeAjustes = True
        Else
            SQL = DevuelveDesdeBD("idhorario", "ModificarFichajes", "idhorario", vH.IdHorario)
            'Tiene ajustes puestos
            If SQL <> "" Then HacerRedondeoDeAjustes = True
        End If
        
        If HacerRedondeoDeAjustes Then Rectificar vH, lblPpal, lblDetall, vbRecAjustes
        
        'If vH.Rectificar > 0 : Este caso ya lo hemos contemplado aqui arriba
        If vH.Rectificar > 1 Then Rectificar vH, lblPpal, lblDetall, vH.Rectificar
        
        
        

        
        
        'Siguiente horario a procesar en esa fecha
        RsH.MoveNext
    Wend
    RsH.Close
    
    
    
    
    
    
    Set RsH = Nothing



    

End Sub


'TIPO ajustes=vH.Rectificar-->> vbRecAjustes ....
Private Function Rectificar(ByRef vH As CHorarios, ByRef l1 As Label, ByRef L2 As Label, TipoAjustes As Byte) As Byte
Dim Recortes As ADODB.Recordset
Dim vRs As ADODB.Recordset
Dim H1, h2, H3
Dim cad As String
Dim Aux As String
Dim Trabajador As Long
Dim i As Integer
Dim Hora As Date
Dim HoraAnt As Date
Dim HoraFin As Date
Dim H8 As Integer
Dim IncremeHora As Integer
Dim k As Integer


    l1.Caption = "Rectificar"
    l1.Refresh
    Select Case TipoAjustes
    Case vbRecAjustes
        'Tenemos k recortar en funcion de lo k haya puewto
        'En ajuste manuales
        cad = "SELECT * FROM ModificarFichajes "
        cad = cad & " WHERE Idhorario= " & vH.IdHorario
        Set Recortes = New ADODB.Recordset
        Recortes.Open cad, conn, , , adCmdText
        While Not Recortes.EOF
            
            H1 = "'" & Format(Recortes.Fields(1), "hh:mm") & "'"
            h2 = "'" & Format(Recortes.Fields(2), "hh:mm") & "'"
            'H3 = "#" & Format(Recortes.Fields(4), "hh:mm") & "#"
            H3 = Format(Recortes.Fields(3), "hh:mm")
            'Label
            
            l1.Caption = H1 & " - " & h2 & "   --> " & H3
            l1.Refresh
            DoEvents
            'Creamos la consulta de acutalizacion
            'Para cada recortes modificamos la tabla
            
            
            'SELECT EntradaFichajes.idTrabajador, Trabajadores.IdTrabajador, Secciones.IdSeccion
            'FROM Secciones INNER JOIN (EntradaFichajes INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador) ON Secciones.IdSeccion = Trabajadores.Seccion
            'WHERE (((Secciones.IdSeccion)=1));
            
            'VER que estamos lincando con calendarioL
            
            Set vRs = New ADODB.Recordset
            cad = "select secuencia,hora from entradafichajes,calendariol where entradafichajes.fecha=calendariol.fecha and"
            cad = cad & " entradafichajes.fecha =calendariol.fecha"
            cad = cad & " and calendariol.idhorario = " & vH.IdHorario
            cad = cad & " and calendariol.fecha= '" & Format(vH.Fecha, FormatoFecha) & "'"
            cad = cad & " and hora >= " & H1 & " and hora <= " & h2
            Set vRs = New ADODB.Recordset
            vRs.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            While Not vRs.EOF
                cad = "UPDATE entradafichajes SET HORA = '" & H3 & "' where secuencia = " & vRs!Secuencia
                L2.Caption = H3
                'Siguiente
                vRs.MoveNext
                
                conn.Execute cad
            Wend
            vRs.Close

            
            'Siguiente
            Recortes.MoveNext
        Wend
        'Cerramos el recordset de modificar marcjaes
        Recortes.Close
    
    Case vbRecNormCuarto, vbRecNormMedia
        'Ajuste normal en funcion de un valor determinadao en parametros
            
        
            l1.Caption = "Ajuste por fraccion"
            l1.Refresh
            '------------------------
            'Ajustes por redondeo
            '-----------------------
          'Tenemos k redondear a cuartos, o a media hora en funcion del valor en datos de empresa
          'Entonces, a partir de las doce de la mañana vamos haciendo hasta las 11:30 de la noche
          If vH.Rectificar = vbRecNormCuarto Then
              Aux = "15"
          Else
              Aux = "30"
          End If
          
          'Primero vemos los minutos para cortar los intervalos
          H1 = vEmpresa.MinutosRedondeo
          
          
          Dim redondeo
          Dim mihora
          
          redondeo = Val(H1)
          
          'Cogemos el minimo y el maximo
          Set vRs = New ADODB.Recordset
          
          'Febrero  2014
          If vEmpresa.CreaCalDiariaTra Then
            cad = "select min(hora),max(hora) from entradafichajes,calendariot where"
            cad = cad & " entradafichajes.fecha=calendariot.fecha and entradafichajes.idtrabajador =calendariot.idtrabajador and"
            cad = cad & " entradafichajes.fecha =calendariot.fecha and calendariot.idhorario = " & vH.IdHorario
            cad = cad & " and calendariot.fecha= '" & Format(vH.Fecha, FormatoFecha) & "' "
            cad = cad & " AND hora< '23:59:59'"
              
          Else
             cad = "select min(hora),max(hora) from entradafichajes,calendariol,trabajadores  where"
             cad = cad & " entradafichajes.fecha=calendariol.fecha and entradafichajes.idtrabajador =trabajadores.idtrabajador"
             cad = cad & " and calendariol.idhorario = " & vH.IdHorario
             cad = cad & " and calendariol.fecha= '" & Format(vH.Fecha, FormatoFecha) & "' "
             cad = cad & " AND hora< '23:59:59'"
          End If
          vRs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

          
          
          'Ajuste hora fin
          mihora = vRs.Fields(1) 'fin
          If mihora < CDate("23:30:00") Then
                mihora = DateAdd("n", Val(Aux), mihora)
          
          
          
                i = Minute(mihora)
                If vH.Rectificar = vbRecNormCuarto Then
                   i = (i \ 15)
                   
                   i = 15 * i
                   
                Else
                     If i < 31 Then
                          i = 0
                     Else
                          i = 30
                     End If
                     
                End If
                HoraFin = CDate(Hour(mihora) & ":" & Format(i, "00"))
                
                
        Else
                If vH.Rectificar = vbRecNormCuarto Then
                    i = 15
                Else
                    i = 30
                End If
                HoraFin = Format(DateAdd("n", -i, CDate("0:00:00")), "hh:mm:ss")
          'Minimo   ######### INICIO
        End If
          mihora = vRs.Fields(0) 'minimo
          
          i = Minute(mihora)
          If vH.Rectificar = vbRecNormCuarto Then
             i = (i \ 15)
             
             i = 15 * i
             
          Else
               If i < 31 Then
                    i = 0
               Else
                    i = 30
               End If
               
          End If
          Hora = CDate(Hour(mihora) & ":" & Format(i, "00"))

          vRs.Close
              
          
          'El bucle

              HoraAnt = Hora
              While Hora <= HoraFin
              
                    L2.Caption = Hora
                    L2.Refresh
              
                    'FALTA CONTEMPLAR EL ULTIMO INTERVALO
                    
                    

                    
                      mihora = DateAdd("n", redondeo, Hora)
                  
                      H1 = "'" & Format(HoraAnt, "hh:mm:ss") & "'"
                      h2 = "'" & Format(mihora, "hh:mm") & ":59'"
                      'H3 = "#" & Format(Hora, "hh:mm") & "#"
                      'FALTA####
                      'If Hora > CDate("23:00") Then St op
                      H3 = Format(Hora, "hh:mm") & ":00"
                      
                      l1.Caption = HoraAnt & " - " & mihora & "   --> " & H3
                      l1.Refresh
                      
                      If vEmpresa.CreaCalDiariaTra Then
                          cad = "select secuencia,hora from entradafichajes,calendariot where"
                          cad = cad & " entradafichajes.fecha=calendariot.fecha and entradafichajes.idtrabajador =calendariot.idtrabajador and"
                          cad = cad & " entradafichajes.fecha =calendariot.fecha and calendariot.idhorario = " & vH.IdHorario
                          cad = cad & " and calendariot.fecha= '" & Format(vH.Fecha, FormatoFecha) & "' "
                          
                      Else
                          cad = "select secuencia,hora from entradafichajes,calendariol,trabajadores  where"
                          cad = cad & " entradafichajes.fecha=calendariol.fecha and entradafichajes.idtrabajador =trabajadores.idtrabajador"
                          cad = cad & " and calendariol.idhorario = " & vH.IdHorario
                          cad = cad & " and calendariol.fecha= '" & Format(vH.Fecha, FormatoFecha) & "' "
                      End If
                      cad = cad & " AND hora >= " & H1 & " AND hora <= " & h2
                      
                      vRs.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                        While Not vRs.EOF
                            
                            cad = "UPDATE entradafichajes SET HORA = '" & H3 & "' where secuencia = " & vRs!Secuencia
                            conn.Execute cad
                            'Siguiente
                            vRs.MoveNext
                        Wend
                        vRs.Close
                          
                      'Subimos hora y hora post
                      Hora = DateAdd("n", Val(Aux), Hora)
                      HoraAnt = DateAdd("n", 1, mihora)
              Wend
          
          'Hacemos el ultimo, desde las 12-algo de la noche hasta las 23:59 son las 23:59
                      H1 = "'" & Format(HoraAnt, "hh:mm") & "'"
                      h2 = "'23:59:58'"
                      H3 = "'23:59:59'"
                      
                      cad = "select secuencia,hora from entradafichajes,calendariot where"
                      cad = cad & " entradafichajes.fecha=calendariot.fecha and entradafichajes.idtrabajador =calendariot.idtrabajador and"
                      cad = cad & " entradafichajes.fecha =calendariot.fecha and calendariot.idhorario = " & vH.IdHorario
                      cad = cad & " and calendariot.fecha= '" & Format(vH.Fecha, FormatoFecha) & "' "
                      cad = cad & " AND hora >= " & H1 & " AND hora <= " & h2
                      
                      Set vRs = New ADODB.Recordset
                      vRs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        While Not vRs.EOF
                            cad = "UPDATE entradafichajes SET HORA = '" & H3 & "' where secuencia = " & vRs!Secuencia
                            conn.Execute cad
                            'Siguiente
                            vRs.MoveNext
                        Wend
                        vRs.Close
                      'Ejecutamos el SQL
                      
                      
                      
                      
                      
                      
       Case vbRecESCuarto, vbRecESMedia
                      
         '----------------------------------------------------
        
               
        '----------------------------
         '  AJUSTES por entrada salida
         '----------------------------
            
         'Cojeremos para cada trabajador, cada fecha e iremos viendo entrada salida
         'Los marcajes, y por conteo iremos viendo
         ' Entrada--> ajuste entrada.... salida---> ajuste salida
      
               
          If vH.Rectificar = vbRecESCuarto Then
              Aux = "15"
          Else
              Aux = "30"   'Entradas salidas cada media hora
          End If
                   
         
                   

      
          'Primero vemos los ajustes. Medias horas, cuartos
          cad = "AjusteSalida"
          
          
          If vEmpresa.CreaCalDiariaTra Then
            cad = "select entradafichajes.* ,hour(hora) LaHora,minute(hora) minutos,second(hora) segundos, if(hora<'0:00:00',1,0) Negativa"
            cad = cad & ",if(hora<'0:00:00',ADDTIME(hora , '24:00:00' ),if(hour(hora)>24,ADDTIME(hora , '-24:00:00' ),hora)) HoraPintarneg"

            cad = cad & " from entradafichajes,calendariot where entradafichajes.fecha=calendariot.fecha"
            cad = cad & " and entradafichajes.idtrabajador =calendariot.idtrabajador and entradafichajes.fecha =calendariot.fecha and"
            cad = cad & " calendariot.idhorario = " & vH.IdHorario & " and calendariot.fecha= '" & Format(vH.Fecha, FormatoFecha) & "'"
          Else
            cad = "select entradafichajes.* ,hour(hora) LaHora,minute(hora) minutos,second(hora) segundos ,if(hora<'0:00:00',1,0) Negativa"
            
            'Si fuera negativa la hora
            cad = cad & ",if(hora<'0:00:00',ADDTIME(hora , '24:00:00' ),if(hour(hora)>24,ADDTIME(hora , '-24:00:00' ),hora)) HoraPintarneg"
            
            cad = cad & " from entradafichajes,calendariol,trabajadores where "
            cad = cad & " entradafichajes.idtrabajador =trabajadores.idtrabajador and  trabajadores.idcal=calendariol.idcal and"
            cad = cad & " entradafichajes.Fecha = calendariol.Fecha And calendariol.IdHorario = " & vH.IdHorario
            cad = cad & " and calendariol.fecha= '" & Format(vH.Fecha, FormatoFecha) & "'"

          End If
          cad = cad & " ORDER By entradafichajes.idTrabajador,Fecha,Hora"
          Trabajador = -1
          Set vRs = New ADODB.Recordset
          vRs.Open cad, conn, , , adCmdText
          While Not vRs.EOF
          
          
              If Trabajador <> vRs!idTrabajador Then
                  'label
                   Trabajador = vRs!idTrabajador
                    L2.Caption = "Trab: " & Trabajador
                    DoEvents
                    i = 0
                    
                    
                    
                    'If InStr(1, ",901,178,169,196,193,150,182,154,", "," & vRs!idTrabajador & ",") > 0 Then St op
        
                    'If Trabajador = 30 Then St op
                   
              End If
 
              If vRs!LaHora >= 24 Then
                  IncremeHora = -24
                  H8 = vRs!LaHora
              ElseIf vRs!Negativa = 1 Then
                  IncremeHora = 24
                  H8 = -vRs!LaHora
              Else
                    IncremeHora = 0
                    HoraAnt = Format(vRs!Hora, "hh:mm:ss")
              End If
              If IncremeHora <> 0 Then
                'Acabalgada
                'LaHora,minute(hora) minutos,second(hora) segundos
                If IncremeHora = -24 Then
                    HoraAnt = Format(H8 + IncremeHora, "00") & ":" & vRs!Minutos & ":" & vRs!segundos
                Else
                    
                    HoraAnt = Format(vRs!HoraPintarneg, "hh:mm:ss")
                End If
                
              End If
              
              If (i Mod 2) = 0 Then
                  'Entrada
                  Hora = HoraRectificada(HoraAnt, vEmpresa.AjusteEntrada, CInt(Aux))
              Else
                  'Salida
                  Hora = HoraRectificada(HoraAnt, vEmpresa.AjusteSalida, CInt(Aux))
              End If
                'If Hora <> HoraAnt Then St op
              'reajusto la hora
              If IncremeHora <> 0 Then
                    
                    If IncremeHora = -24 Then
                        k = Hour(Hora)
                        k = k - IncremeHora
                        cad = Format(k, "00") & Format(Hora, ":nn:ss")
                    Else
                        cad = Horas_Quitar24(CDate(Hora), False)
                    End If
              Else
                cad = Format(Hora, "hh:mm:ss")
              End If
              
             
              
              cad = "UPDATE entradafichajes SET HORA = '" & cad & "' where secuencia = " & vRs!Secuencia
              
              vRs.MoveNext
              espera 0.03
              conn.Execute cad
              'Siguiente
              i = i + 1
           Wend
           vRs.Close
            
    End Select
    
    Set vRs = Nothing
'Todo correcto
Rectificar = 0
Exit Function
ErrorRectificacionDeMarcajes:
    MuestraError Err.Number
    Rectificar = 1
End Function

'
'Private Function RectifcaTipoCOntrolEstricto(ByRef vH As CHorarios, ByRef l1 As Label, ByRef L2 As Label) As Byte
'Dim SQL As String
'Dim Rs As ADODB.Recordset
'
'Dim ListaTraba As String
'
'    On Error GoTo ERectifcaTipoCOntrolEstricto
'    RectifcaTipoCOntrolEstricto = 1
'
'    If vH.EsDiaFestivo Then Exit Function
'
'    SQL = "Select calendariot.idtrabajador,nomtrabajador from calendariot,trabajadores where "
'    SQL = SQL & " calendariot.fecha='" & Format(vH.Fecha, FormatoFecha) & "'"
'    SQL = SQL & " AND calendariot.idHOrario = " & vH.IdHorario
'    SQL = SQL & " AND calendariot.idtrabajador =  trabajadores.idtrabajador "
'    Set Rs = New ADODB.Recordset
'    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    ListaTraba = ""
'    While Not Rs.EOF
'        L2.Caption = Rs!nomtrabajador
'        L2.Refresh
'
'        If Len(ListaTraba) < 220 Then
'            ListaTraba = ListaTraba & " OR idtrabajador = " & Rs!idTrabajador
'        Else
'            L2.Caption = "Proceso revision"
'            L2.Refresh
'
'
'        End If
'
'
'
'
'        Rs.MoveNext
'
'    Wend
'    RectifcaTipoCOntrolEstricto_SQL vH, ListaTraba
'
'    Rs.Close
'    Set Rs = Nothing
'
'    Exit Function
'ERectifcaTipoCOntrolEstricto:
'    MuestraError Err.Number, Err.Description
'    Set Rs = Nothing
'
'End Function


'Private Function RectifcaTipoCOntrolEstricto_SQL(ByRef vH1 As CHorarios, LaListaTrabajadores As String) As Byte
'Dim Cade As String
'Dim RT As ADODB.Recordset
'Dim Contador As Byte
'Dim idTrab As Long
'Dim Diferencia As Long
'Dim HoraComparacion As Date
'
'
'    On Error GoTo ERectifcaTipoCOntrolEstricto2
'    'QUito el primer or
'    LaListaTrabajadores = Mid(LaListaTrabajadores, 4)
'    Cade = "Select * from entradafichajes where "
'    Cade = Cade & " Fecha = '" & Format(vH1.Fecha, FormatoFecha) & "' AND ("
'    Cade = Cade & LaListaTrabajadores & ") ORDER BY idtrabajador,hora"
'    Set RT = New ADODB.Recordset
'    RT.Open Cade, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    idTrab = -1
'    While Not RT.EOF
'        If idTrab <> RT!idTrabajador Then
'            Contador = 1
'            idTrab = RT!idTrabajador
'        End If
'
'
'
'
'
'
'
'        If Diferencia <> 0 Then
'
'            If Diferencia < 0 Then
'                If Diferencia < vEmpresa.MaxRetraso Then
'                    'SI K AJUSTO. Si no, na de na
'
'                End If
'            Else
'                If Diferencia <= vEmpresa.MaxExceso Then
'                    'Corrigo
'
'                Else
'
'
'                End If
'            End If
'
'        End If
'
'
'        RT.MoveNext
'    Wend
'    RT.Close
'    Set RT = Nothing
'    Exit Function
'ERectifcaTipoCOntrolEstricto2:
'    MuestraError Err.Number
'    Set RT = Nothing
'End Function



Private Function ComparaHoraSobreHorario2(Entrada As Boolean, ByRef ElHorario As CHorarios) As Date
    
End Function


Public Function HoraRectificada(Hora As Date, Ajuste As Integer, FraccionHora As Integer) As Date
Dim Nueva As Date
Dim Minu As Integer
Dim Salir As Boolean

    
        HoraRectificada = Hora
        Nueva = CDate(Hour(Hora) & ":00")
        Salir = False
        Do
            Minu = DateDiff("n", Nueva, Hora)
            If Minu > Ajuste Then
                If DateDiff("n", Nueva, CDate("23:59")) < 15 Then
                    HoraRectificada = CDate("23:59")
                    Exit Function
                End If
                Nueva = DateAdd("n", FraccionHora, Nueva)
            Else
                Salir = True
                HoraRectificada = Nueva
            End If
        Loop Until Salir
        
End Function




'---------------------------------------------------------------------------
'
'       GENERACION DE MARCAJES
'
'---------------------------------------------------------------------------

Public Sub GeneraEntradasMarcajes(Fecha As Date, ByRef l1 As Label, ByRef L2 As Label)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim vH As CHorarios
Dim Control As Integer
Dim FechaBaja As Date
Dim vM As CMarcajes
Dim C1 As Collection
Dim C2 As Collection
Dim Num As Long
Dim Tot As Long
Dim RTra As ADODB.Recordset
Dim MiCal As Integer
Dim CalAux As Integer

Dim ModificaLasParadas As Boolean
Dim ValorModificadoParadas As Currency

    l1.Caption = "Obtener conjunto de registros"
    L2.Caption = ""
    l1.Refresh
    L2.Refresh
    DoEvents
    
    
    idIncidenciaGenerada = 0 'Contador de incidencas para trabajador/dia
    'Vemos cual es la que le toca
    SQL = "select max(id) from incidenciasgeneradas"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        idIncidenciaGenerada = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close

    If vEmpresa.CreaCalDiariaTra Then
        SQL = "select entradafichajes.idtrabajador,idhorario from entradafichajes,calendariot where"
        SQL = SQL & " entradafichajes.idTrabajador = calendariot.idTrabajador"
        SQL = SQL & " and entradafichajes.fecha =calendariot.fecha and"
        SQL = SQL & " entradafichajes.fecha='" & Format(Fecha, FormatoFecha) & "'"
        SQL = SQL & " group by 1,2 order by idhorario,idtrabajador"
    Else
        'TIPO ALZIRA. No llevan una entrada en calendariot para cada dia
        SQL = "SElect idhorario,entradafichajes.idtrabajador from entradafichajes,calendariol ,trabajadores where"
        SQL = SQL & " entradafichajes.idTrabajador = trabajadores.idTrabajador And trabajadores.idCal = calendariol.idCal"
        SQL = SQL & " And entradafichajes.Fecha = calendariol.Fecha"
        SQL = SQL & " AND entradafichajes.fecha='" & Format(Fecha, FormatoFecha) & "'"
        SQL = SQL & " group by 1,2 order by idhorario,idtrabajador"
    
    End If
    

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set vH = New CHorarios
    
    
    l1.Caption = "Procesar marcaje.   "
    l1.Refresh
     
     
    If RS.EOF Then
        MsgBox "Ninguna entrada a procesar de : " & Fecha, vbExclamation
        RS.Close
        Exit Sub
    Else
        While Not RS.EOF
            Tot = Tot + 1
            RS.MoveNext
        Wend
        RS.MoveFirst
    End If
    'EL SQL para los inserts
    vSQL2 = "INSERT INTO entradamarcajes (Secuencia, idTrabajador, idMarcaje, Fecha, Hora, "
    vSQL2 = vSQL2 & "idInci, HoraReal, reloj) VALUES ( "
    Set RTra = New ADODB.Recordset
    Num = 0
    MiCal = 0
    While Not RS.EOF
        Num = Num + 1
        
        If (Num Mod 30) = 0 Then
            l1.Caption = "Leyendo BD .... "
            l1.Refresh
            L2.Caption = "entrada fichajes ...."
            L2.Refresh
            DoEvents
            espera 0.5
        End If
        
        'sql = "," & RS!idTrabajador & ","
        'If InStr(1, ",901,178,169,196,193,150,182,154,", sql) > 0 Then St op
        
        
        l1.Caption = "Procesar marcaje.   (" & Num & " de " & Tot & ")"
        l1.Refresh
        L2.Caption = "Trab: " & RS!idTrabajador
        L2.Refresh
    
       
        SQL = "Select control,fecbaja,idcal,nomtrabajador from trabajadores WHERE idtrabajador=" & RS!idTrabajador
        RTra.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Control = DBLet(RTra!Control, "N")
        SQL = DBLet(RTra!FecBaja, "T")
        CalAux = RTra!idCal
        L2.Caption = "Trab: " & Mid(DBLet(RTra!nomtrabajador, "T"), 1, 30)
        L2.Refresh
    
        
        If SQL <> "" Then
            FechaBaja = CDate(SQL)
        Else
            FechaBaja = CDate("01/01/2300")
        End If
        RTra.Close
        
        ModificaLasParadas = False
        If vEmpresa.QueEmpresa = 2 Then
        
            SQL = "Select * from tmpcombinada WHERE codusu=" & vUsu.Codigo & " AND  idtrabajador=" & RS!idTrabajador
            RTra.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RTra.EOF Then
                If Not IsNull(RTra!codusu) Then
                    ModificaLasParadas = True
                    ValorModificadoParadas = RTra!hr
                End If
            End If
            RTra.Close
        End If
            
        
        If vH.IdHorario <> RS!IdHorario Then
            MiCal = CalAux
            If vH.Leer(RS!IdHorario, Fecha, CalAux) = 1 Then
                'ERROR
                
            End If
        Else
            'Son el mismo horario pero de distinto calendario. Creare un function nueva
            If MiCal <> CalAux Then
                vH.Leer RS!IdHorario, Fecha, CalAux
                MiCal = CalAux
            End If
        End If
        
        
        
        
        
        Set vM = New CMarcajes
        vM.Siguiente
        vM.Fecha = Fecha
        vM.Nocturno = False
        vM.idTrabajador = RS!idTrabajador
        vM.IdHorario = vH.IdHorario
        vM.Correcto = False
        TotalParadas2 = 0
        
        
        
        'Aqui aqui
        'Abril 2015.
        'En el previo indicaremos que trabajadores les vamos a quitar almuerzo (pudiendo quitarlo), y en cuales NO
        'Entonces, para un trabajador que le hemos dicho que NO le quitamos , pararemos y
        
        Select Case Control
        Case 3
            'TIPO 3.
            'Se contabilizaran las horas totales y punto
            ProcesarMarcaje_Tipo3 vM, vH, False
            
            
        Case 2
            ' Se contabilizaran las horas totales y se compararan con las horas
            'que debia haberse trabajado generando incidencias o no
            
            ProcesarMarcaje_Tipo2 vM, vH, False, ModificaLasParadas, ValorModificadoParadas
            
            
        Case Else
            'El seguimiento es exahustivo
            'Se comparan las entradas con los margenes de cortesia
            'generando entradas por cada una de ellas
            
            ProcesarMarcaje_Tipo1 vM, vH, False
            
        End Select
            
            
    
        
        'Comprobamos si esta de baja
        If FechaBaja <= Fecha Then
            vM.Correcto = False
            vM.IncFinal = vEmpresa.IncTarjError
            vM.Modificar
        End If
        
        
        
    


    
    
    
        RS.MoveNext
        Set vM = Nothing
    Wend
    RS.Close
    Set RTra = Nothing
    
    
    'Aqui contemplaremos todos los que para esa fecha estan de vacaciones
    If vEmpresa.TodosLosDias Then
        l1.Caption = "Lectura datos trabajadores.... vacaciones "
        l1.Refresh
        
        SQL = "Select idtrabajador,idhorario from  calendariot where "
        SQL = SQL & "  fecha='" & Format(Fecha, FormatoFecha) & "'"
        SQL = SQL & " and tipodia = " & vbDiaVacas
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set C1 = New Collection
        Set C2 = New Collection
        While Not RS.EOF
            C1.Add Val(RS!idTrabajador)
            C2.Add Val(RS!IdHorario)
            RS.MoveNext
        Wend
        RS.Close
        
        'Abrimos el recodset con los que han trabajado este dia
        SQL = "Select entrada,idtrabajador from marcajes where fecha='" & Format(Fecha, FormatoFecha) & "'"
        SQL = SQL & " ORDER BY idtrabajador"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
        
            For Control = C1.Count To 1 Step -1
                l1.Caption = "Vacaciones: " & C1.Item(Control)
                l1.Refresh
                
                RS.Find "idtrabajador = " & C1.Item(Control), , adSearchForward, 1
                If Not RS.EOF Then
                    ''  HA TENIDO MARCAJES !!!!!!!!
                    '  ..... ESTANDO DE VACACIONES. OOOOOOOohhhhhhhhh
                    'Le UPDATEO el marcaje a vacaciones y se lo pongo incorrecto pq no debia de estar
                    SQL = "UPDATE marcajes set incfinal =" & vEmpresa.IncVacaciones
                    SQL = SQL & " ,correcto=0 where entrada =" & RS!Entrada
                    conn.Execute SQL
                    
                    'Lo quito de la lista para que asi luego
                    'no le genere los marcajes
                    C1.Remove Control
                    C2.Remove Control
                End If
            Next Control
            
        End If
        
        
        If C1.Count > 0 Then
            DoEvents
            Set vM = New CMarcajes
            vM.HorasDto = 0
            vM.HorasTrabajadas = 0
            vM.HorasIncid = 0
            vM.Correcto = True
            vM.Fecha = Fecha
            vM.IncFinal = vEmpresa.IncVacaciones
            
            For Control = 1 To C1.Count
                l1.Caption = "Vacaciones: " & C1.Item(Control)
                l1.Refresh
            
                'Generamos el marcaje con la incidencia de vacaciones
                'Pero lo pongo correcto
                vM.Siguiente
                
                
                
                vM.IdHorario = C2.Item(Control)
                vM.idTrabajador = C1.Item(Control)
                
                vM.Agregar
            Next Control
            Set vM = Nothing
       End If
       Set C1 = Nothing
       Set C2 = Nothing
    End If 'De todos los dias
End Sub


Public Sub GeneraIncidencia(Inci As Integer, marca As Long, Horas As Currency)
Dim cad As String
    On Error Resume Next
    cad = "INSERT INTO incidenciasgeneradas (Id, EntradaMarcaje, Incidencia, horas) VALUES ("
    idIncidenciaGenerada = idIncidenciaGenerada + 1
    cad = cad & idIncidenciaGenerada & "," & marca & "," & Inci & "," & TransformaComasPuntos(CStr(Horas)) & ")"
    conn.Execute cad
    If Err.Number <> 0 Then MuestraError Err.Number, "Generando incidencia. " & vbCrLf & cad

End Sub
'-------------------------------------------------------------------------------
Public Sub ProcesarMarcaje_Tipo1(ByRef vMar As CMarcajes, ByRef vH As CHorarios, RevisionEnMarcajes As Boolean)
Dim Rss As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim NumTikadas As Integer
Dim T1 As Currency
Dim T2 As Currency
Dim kIncidencia As Currency
Dim TieneIncidencia As Boolean
Dim MarcajeCorrecto As Boolean
Dim Exceso As Date
Dim Retraso As Date
Dim i As Long
Dim v(3) As Currency
Dim vI(3) As Integer
Dim cad As String
Dim HoraH As Date
Dim InciManual As Integer
Dim N As Integer
Dim TotalH As Currency
Dim SQLUpdateHora As String

    'Ahora ya tenemos las horas tikadas reflejadas
    'Comprobamos las horas en funcion de los horarios
    '  y calculamos las horas comprobadas
    
    
    Set Rss = New ADODB.Recordset
    'Vector para incidencias
    For i = 0 To 3
        v(i) = 0
        vI(i) = 0
    Next i
    'Seleccionamos todas las horas de este
    If RevisionEnMarcajes Then
    
        cad = "Select * from EntradaMarcajes WHERE idmarcaje=" & vMar.Entrada
        'cad = cad & " AND Fecha='" & Format(vMar.Fecha, FormatoFecha) & "'"
        cad = cad & " ORDER BY Hora"
    Else
    
        cad = "Select * from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
        cad = cad & " AND Fecha='" & Format(vMar.Fecha, FormatoFecha) & "'"
        cad = cad & " ORDER BY Hora"
    End If
    Rss.CursorType = adOpenStatic
    Rss.Open cad, conn, , , adCmdText
    
    If Rss.EOF Then
        'Si no hay ninguna entrada
        Rss.Close
        GoTo ErrorProcesaMarcaje
    End If
    InciManual = 0

    
    NumTikadas = 0
    While Not Rss.EOF
        NumTikadas = NumTikadas + 1
        Rss.MoveNext
    Wend


If vH.EsDiaFestivo Then
    'Si es festivo asignamos las tikadas segun vengan
    ' y todo pasa a ser horas extras
    vMar.Festivo = True
    If (NumTikadas Mod 2) > 0 Then
        'Numero de marcajes impares. No podemos calcular horas
        'trabajadas. Generamos error en marcaje
        
        vMar.IncFinal = vEmpresa.IncMarcaje
        vMar.HorasIncid = 0
        vMar.HorasTrabajadas = 0
        GeneraIncidencia vEmpresa.IncMarcaje, vMar.Entrada, 0
    Else
        N = NumTikadas \ 2
        TotalH = 0
        'NUMERO DE MARCAJES PAR
        Rss.MoveFirst
        For i = 1 To N
            T1 = DevuelveValorHora(Rss!Hora)
            Rss.MoveNext
            T2 = DevuelveValorHora(Rss!Hora)
            Rss.MoveNext
            TotalH = TotalH + (T2 - T1)
        Next i
        
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
            For i = 1 To N
                PrimerTicaje = Rss!Hora
                Rss.MoveNext
                If PrimerTicaje <= vH.HoraDtoMer Then
                    If Rss!Hora > vH.HoraDtoMer Then quitarmerienda = True
                End If
            Next i
        End If
        
        'Ahora ya sabemos las horas trabajadas
        TotalH = RealizaRedondeo(TotalH)
        
        'Asignamos a la incidencia
        T2 = TotalH
        If QuitarAlmuerzo Then T2 = T2 - vH.DtoAlm
        If quitarmerienda Then T2 = T2 - vH.DtoMer
        
        TotalH = Round(T2, 2)
        
        
        
        vMar.HorasTrabajadas = TotalH
        vMar.HorasIncid = TotalH
        vMar.IncFinal = vEmpresa.IncHoraExtra
        
        
        
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
        Exceso = DevuelveHora(vEmpresa.MaxExceso)
        Retraso = DevuelveHora(vEmpresa.MaxRetraso)
        vMar.HorasDto = 0
        i = 0
        Rss.MoveFirst
        PrimerTicaje = Format(Rss!Hora, "hh:mm:ss")
        SQLUpdateHora = ""
        While Not Rss.EOF
            If Rss!IdInci > 0 Then
                InciManual = Rss!IdInci
                vI(i) = InciManual
            End If
            Select Case i
            Case 0
                HoraH = vH.HoraE1
            Case 1
                HoraH = vH.HoraS1
            Case 2
                HoraH = vH.HoraE2
            Case 3
                HoraH = vH.HoraS2
            End Select
            kIncidencia = EntraDentro(Format(Rss!Hora, "hh:mm:ss"), HoraH, Exceso, Retraso, (i Mod 2) = 0)
            v(i) = kIncidencia
            If kIncidencia = 0 Then
                'Como ha entrado dentro entonces UPDATE la hora a hora
                If RevisionEnMarcajes Then
                    SQLUpdateHora = "entradamarcajes"
                Else
                    SQLUpdateHora = "entradafichajes"
                End If
                SQLUpdateHora = "UPDATE " & SQLUpdateHora & " SET hora ='" & Format(HoraH, "hh:mm:ss") & "' WHERE Secuencia =" & Rss!Secuencia
                EjecutaSQL SQLUpdateHora
            End If
            i = i + 1
            UltimoTicaje = Format(Rss!Hora, "hh:mm:ss")
            Rss.MoveNext
        Wend
        
        If SQLUpdateHora <> "" Then
            'Como he hecho unos updates
            'Refresco la tabla
            SQLUpdateHora = Rss.Source
            Rss.Close
            Rss.Open SQLUpdateHora, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        End If
        
        'Ahora ya tenmos si ha llegado tarde, ha salido antes etc, por lo tanto
        ' realizamos los calculos de las horas y generaremos, si cabe
        'las incidencias
        'En v() tenemos que si es 0 nada, pero si es menor tenemos la horas extras
        ' y si es mayor las horas de retraso
        'En t1 tendremos las horas en las incidencias
        T1 = 0
        TieneIncidencia = False
        For i = 0 To 3
            T1 = T1 + v(i)
            If v(i) > 0 Then
                'Si tenia incidencia manul la pongo
                If vI(i) <> 0 Then
                    N = vI(i)
                Else
                    N = vEmpresa.IncRetraso
                End If
                GeneraIncidencia N, vMar.Entrada, v(i)
                TieneIncidencia = True
                Else
                    
                    If v(i) < 0 Then
                    
                        If vI(i) <> 0 Then
                            N = vI(i)
                        Else
                            N = vEmpresa.IncHoraExceso
                        End If
                        GeneraIncidencia N, vMar.Entrada, Abs(v(i))
                        TieneIncidencia = True
                    End If
            End If
        Next i
        'Debug.Print vMar.IdTrabajador & ": " & T1
        
        'si tiene dto. Le sumaremos al valor obtenido en T1 el valor de los dtos
        'Comprobamos los dtos almuerzo merienda
        '******************************************************
        QuitarAlmuerzo = False
        quitarmerienda = False
        N = (NumTikadas \ 2)
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
            Rss.MoveFirst
            For i = 1 To N
                PrimerTicaje = Format(Rss!Hora, "hh:mm:ss")
                Rss.MoveNext
                If PrimerTicaje <= vH.HoraDtoMer Then
                    If Format(Rss!Hora, "hh:mm:ss") > vH.HoraDtoMer Then quitarmerienda = True
                End If
                Rss.MoveNext
            Next i
        End If
            
        
        'Asignamos a la incidencia
        T2 = vH.TotalHoras
        If QuitarAlmuerzo Then
            vMar.HorasDto = vMar.HorasDto + vH.DtoAlm
        End If
        If quitarmerienda Then
            't2 = T2 - vH.DtoMer  es tipo estricto. No deberia llevar dtos
            vMar.HorasDto = vMar.HorasDto + vH.DtoMer
        End If
        
        
        'CREO QUE ESTABA MAL PQ le quitaba el almuerzo tb a las horas extra
        'NO se quita el amuerzo. Antes los signos estaban al reves
        '------------------
'        If vH.DtoAlm > 0 Then
'            If Not QuitarAlmuerzo Then
'
'                If T1 >= 0 Then
'                    '
'                    'Me debe mas horas
'                    T1 = T1 - vH.DtoAlm
'                Else
'                    'Horas extra. Le quito el almuerzo
'                    '
'                    If T1 < 0 Then
'                        T1 = T1 - vH.DtoAlm
'                    End If
'                End If
'            End If
'
'        End If

            
     
        
        TotalH = RealizaRedondeo(T2)
        T1 = RealizaRedondeo(T1)
         
         
        
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
            vMar.HorasIncid = Abs(Round(vH.TotalHoras - vMar.HorasTrabajadas, 2))
            Else
                'Vemos si tiene automatica
                If T1 = 0 Then
                    vMar.HorasIncid = 0
                    'No hace falta ponerle incidencia de error
                    
                    vMar.IncFinal = 0
                    
                Else
                        'Falta o sobran horas
                        If T1 > 0 Then
                            'Retraso
                            vMar.IncFinal = vEmpresa.IncRetraso
                            Else
                                vMar.IncFinal = vEmpresa.IncHoraExtra
                        End If
                        vMar.HorasIncid = Abs(T1)
                End If 't2=0
        End If
        
        
    '   El numero de tikadas no coincide
    Else
        Rss.MoveFirst
        While Not Rss.EOF
             If Rss!IdInci > 0 Then InciManual = Rss!IdInci
             Rss.MoveNext
        Wend
        If InciManual > 0 Then
            vMar.IncFinal = InciManual
            GeneraIncidencia InciManual, vMar.Entrada, 0
            Else
                vMar.IncFinal = vEmpresa.IncMarcaje
                GeneraIncidencia vEmpresa.IncMarcaje, vMar.Entrada, 0
        End If
        
        
        'Ahora pondremos las horas trabajadas por diferencias
        Rss.MoveFirst
        TotalH = 0
        If (NumTikadas Mod 2) = 0 Then
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
        N = (NumTikadas \ 2)
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
            For i = 1 To N
                PrimerTicaje = Rss!Hora
                Rss.MoveNext
                If PrimerTicaje <= vH.HoraDtoMer Then
                    If Rss!Hora > vH.HoraDtoMer Then quitarmerienda = True
                End If
            Next i
        End If
    
        'Ahora ya sabemos las horas trabajadas
        'TotalH = Round(TotalH, 2)
        TotalH = RealizaRedondeo(TotalH)
        
        
        'Asignamos a la incidencia
        T2 = TotalH
        If QuitarAlmuerzo Then
            T2 = T2 - vH.DtoAlm
            TotalParadas2 = vH.DtoAlm
        End If
        If quitarmerienda Then
            T2 = T2 - vH.DtoMer
            TotalParadas2 = TotalParadas2 + vH.DtoMer
        End If
        
        TotalH = Round(T2, 2)
        vMar.HorasDto = TotalParadas2
        
        
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
        
        
        
        'Nuevo 09/11/2006
        'Updato las horas de la incidencia generada a la que resultan de la incidencia total
        cad = "UPDATE incidenciasgeneradas SET horas=" & TransformaComasPuntos(CStr(T1))
        cad = cad & " WHERE EntradaMarcaje =" & vMar.Entrada & " AND incidencia = " & vMar.IncFinal
        EjecutaSQL cad
        
        
    End If   'de numero de tikadas=vh.numtikadas
End If 'De DIAFESTIVO



'Por ultimo marcamos o no el campo correcto
vMar.Correcto = vMar.IncFinal = 0





'Grabamos el marcaje
If RevisionEnMarcajes Then
    vMar.Modificar
Else
    vMar.Agregar
End If

    '-------------------------------------------------------------------------
    'Cerramos y borramos todos los fichajes pasandolos a una tabla de marcajes
    If Not RevisionEnMarcajes Then
        Rss.MoveFirst
        espera 0.2
        Set RFin = New ADODB.Recordset
        RFin.Open "Select max(secuencia) from EntradaMarcajes ", conn, , , adCmdText
        If RFin.EOF Then
            i = 1
            Else
                i = DBLet(RFin.Fields(0), "N") + 1
        End If
        RFin.Close
        While Not Rss.EOF
            cad = i & "," & vMar.idTrabajador & "," & vMar.Entrada
            cad = cad & ",'" & Format(Rss!Fecha, FormatoFecha) & "','" & Format(Rss!Hora, "hh:mm:ss")
            cad = cad & "'," & Rss!IdInci & ",'" & Format(Rss!HoraReal, "hh:mm:ss") & "'," & Rss!Reloj & ")"
            conn.Execute vSQL2 & cad
            i = i + 1
            Rss.MoveNext
        Wend
        
    
    
    'Borramos los ticajes
    cad = "Delete from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
    cad = cad & " AND Fecha= '" & Format(vMar.Fecha, FormatoFecha) & "'"
    conn.Execute cad

End If

'Cerramos los recordsets
Rss.Close

Set Rss = Nothing
Set RFin = Nothing


Exit Sub
ErrorProcesaMarcaje:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
    
End Sub




'-------------------------------------------------------------------------------


Public Sub ProcesarMarcaje_Tipo2(ByRef vMar As CMarcajes, ByRef vH As CHorarios, RevisionEnMarcajes As Boolean, ModificaLasParadas As Boolean, CuantoPara As Currency)
Dim Rss As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim NumTikadas As Integer
Dim T1 As Currency
Dim T2 As Currency
Dim i As Long
Dim cad As String
Dim N As Integer
Dim TotalH As Currency
Dim HoE As Currency
Dim IncManual As Integer

'ALZIRA  no quiere que se cambien las paradas cuando revisa desde "Marcajes correctos"
Dim PuedeQuitarAlmuerzoMerienda As Boolean

'ENERO 2015
' Fichadas acabalgadas
Dim IncreHora As Integer
Dim HoraPintar As String
Dim HoraNocturna As Boolean


Dim EsAcabalgado As Boolean
Dim C2 As String

'Ahora ya tenemos las horas tikadas reflejadas
'Comprobamos las horas en funcion de los horarios
'  y calculamos las horas comprobadas

Set Rss = New ADODB.Recordset
IncManual = 0

'Seleccionamos todas las horas de este
If RevisionEnMarcajes Then
    cad = "Select EntradaMarcajes.*,hour(hora) lahora,minute(hora) minutos,second(hora) segundos, concat(horareal,' ') LaReal ,if(hora<'0:00:00',1,0) Negativa"
    cad = cad & " from EntradaMarcajes WHERE idmarcaje=" & vMar.Entrada
    'Cad = Cad & " AND Fecha='" & Format(vMar.Fecha, FormatoFecha) & "'"
    cad = cad & " ORDER BY Hora"
Else
    cad = "Select EntradaFichajes.*,hour(hora) lahora,minute(hora) minutos,second(hora) segundos , concat(horareal,' ') LaReal,if(hora<'0:00:00',1,0) Negativa"
    cad = cad & " from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
    cad = cad & " AND Fecha='" & Format(vMar.Fecha, FormatoFecha) & "'"
    cad = cad & " ORDER BY Hora"
End If
Rss.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText

If Rss.EOF Then
    'Si no hay ninguna entrada
    Rss.Close
    GoTo ErrorProcesaMarcaje_Tipo2
End If


'Si el numero de tikadas es par entonces calculamos las horas
    NumTikadas = 0
    While Not Rss.EOF
        NumTikadas = NumTikadas + 1
        Rss.MoveNext
    Wend




'If vMar.idTrabajador = 144 Then St op

If (NumTikadas Mod 2) > 0 Then
    'Numero de marcajes impares. No podemos calcular horas
    'trabajadas. Generamos error en marcaje
    vMar.IncFinal = vEmpresa.IncMarcaje
    GeneraIncidencia vEmpresa.IncMarcaje, vMar.Entrada, 0
    vMar.HorasIncid = 0
    vMar.HorasTrabajadas = 0
    Else
        N = NumTikadas \ 2
        TotalH = 0
        TotalParadas2 = 0
        'NUMERO DE MARCAJES PAR
        Rss.MoveFirst
        
        

        
        PrimerTicaje = Format(Rss!Hora, "hh:nn:ss") 'Almacenamos el primer ticaje  siempre es entre las 00y las 24
        PuedeQuitarAlmuerzoMerienda = True
        If vEmpresa.QueEmpresa = 2 Then
            'ALZIRA. Desde Revision de marcajes
            If RevisionEnMarcajes Then
                PuedeQuitarAlmuerzoMerienda = False
                TotalParadas2 = vMar.HorasDto   'Las que tuviere
            End If
        End If
        HoraNocturna = False
        For i = 1 To N
        
            
            If Rss!Negativa = 1 Or Rss!LaHora > 23 Then
                IncreHora = 2
                If Rss!Negativa = 1 Then IncreHora = 1
                
                HoraPintar = Format(Rss!LaHora, "00") & ":" & Format(Rss!Minutos, "00") & ":" & Format(Rss!segundos, "00")
                PuedeQuitarAlmuerzoMerienda = False
                HoraNocturna = True
            Else
                IncreHora = 0
                HoraPintar = Format(Rss!Hora, "hh:nn:ss")
            End If
                
            If Rss!IdInci <> 0 Then
                IncManual = Rss!IdInci
                If vEmpresa.QueEmpresa = 4 And IncManual = 2 Then IncManual = 0
            End If
            T1 = DevuelveValorHora3(CByte(IncreHora), HoraPintar)
            
            
            Rss.MoveNext
            If Rss!Negativa = 1 Or Rss!LaHora > 23 Then
                If Rss!Negativa Then
                     IncreHora = 1
                    HoraPintar = Format(Rss!LaHora, "00") & ":" & Format(Rss!Minutos, "00") & ":" & Format(Rss!segundos, "00")
                    
                Else
                    IncreHora = 2
                    HoraPintar = Format(Rss!LaHora, "00") & ":" & Format(Rss!Minutos, "00") & ":" & Format(Rss!segundos, "00")
                    C2 = Rss!LaHora - 24
                End If
                
                
                EsAcabalgado = True
                If vEmpresa.AcabaJornadaDiaSiguiente Then
                    C2 = Format(C2, "00") & ":" & Format(Rss!Minutos, "00") & ":" & Format(Rss!segundos, "00")
                    If CDate(C2) <= vEmpresa.MaximaHoraDiaSiguiente Then EsAcabalgado = False
                End If
                
                If EsAcabalgado Then
                
                    UltimoTicaje = "23:59:59"
                    
                    PuedeQuitarAlmuerzoMerienda = False
                    Debug.Print vEmpresa.MaximaHoraDiaSiguiente
                    
                    HoraNocturna = True
                End If
            Else
                
                
                IncreHora = 0
                HoraPintar = Format(Rss!Hora, "hh:nn:ss")
                UltimoTicaje = Format(Rss!Hora, "hh:nn:ss") 'Obtendremos el ultimo marcaje
                
            End If
            
            
            
            T2 = DevuelveValorHora3(CByte(IncreHora), HoraPintar)
            If Rss!IdInci <> 0 Then
                IncManual = Rss!IdInci
                If vEmpresa.QueEmpresa = 4 And IncManual = 2 Then IncManual = 0 'SALIDA
            End If
            Rss.MoveNext
            TotalH = TotalH + (T2 - T1)
        Next i
        
        
        'ALZIRA. Los ticajes NOCTURNOS llevan una hora mas trabajada
        If HoraNocturna Then
            vMar.Nocturno = True
            TotalH = TotalH + vEmpresa.AcabalgadoIncremento
        
        End If
        
        
        
        
        
        'Comprobamos los detos almuerzo merienda
        '******************************************************
        'Comprobamos si hay que quitar los minutos del almuerzo
        
        If ModificaLasParadas Then PuedeQuitarAlmuerzoMerienda = True   'Fuerza la parada
        
        If PuedeQuitarAlmuerzoMerienda Then
        
            'Si viene forazado que el valor q
            If ModificaLasParadas Then
                TotalParadas2 = CuantoPara
            Else
                'Lo que le corresponda
                If vH.DtoAlm > 0 Then
                    If LeQuitamosElAmluerzo(Rss, vH) Then TotalParadas2 = vH.DtoAlm
                End If
            
                '----------------------------------------------
                'Comprobamos si hay que quitar los minutos de la MER
                'Como esta ya en el ultimo
                If vH.DtoMer > 0 Then
                    If UltimoTicaje > vH.HoraDtoMer Then TotalParadas2 = TotalParadas2 + vH.DtoMer
                End If
            End If
            
            If TotalParadas2 > 0 Then
                If TotalParadas2 > TotalH Then TotalParadas2 = TotalH
                TotalH = TotalH - TotalParadas2
            End If
            
            
        Else
            
            TotalH = TotalH - TotalParadas2
        End If
            
            
        '----------------------------------------------
        '******************************************************
        'Ahora ya sabemos las horas trabajadas
        TotalH = RealizaRedondeo(TotalH)
        vMar.HorasDto = TotalParadas2
        
        
        'Vemos si es diafestivo o no
        'si lo es todas son horas extras, si no
        'calculamos
        If vH.EsDiaFestivo Then
            vMar.Festivo = True
            vMar.HorasTrabajadas = TotalH
            vMar.HorasIncid = TotalH
            vMar.IncFinal = vEmpresa.IncHoraExtra
            'ELSE
            Else     'No es festivo
            'ELSE
            HoE = EntraDentro2(TotalH, vH.TotalHoras, vEmpresa.MaxExceso, vEmpresa.MaxRetraso)
            If HoE = 0 Then
                vMar.HorasTrabajadas = vH.TotalHoras
                vMar.HorasIncid = 0
                vMar.Correcto = True
                If vEmpresa.QueEmpresa = 4 Then
                    If IncManual = 2 Then IncManual = 0
                    vMar.IncFinal = 0
                End If
                If IncManual <> 0 Then GeneraIncidencia IncManual, vMar.Entrada, 0
            Else
                    vMar.Correcto = False
                    If HoE < 0 Then
                        'Horas extras
                        vMar.HorasTrabajadas = vH.TotalHoras - HoE
                        vMar.HorasIncid = Abs(HoE)
                        vMar.IncFinal = vEmpresa.IncHoraExtra
                        GeneraIncidencia vEmpresa.IncHoraExtra, vMar.Entrada, Abs(HoE)  'Genera tb la incidenciagenerada horaextra
                        If IncManual <> 0 And IncManual <> vEmpresa.IncHoraExtra Then GeneraIncidencia IncManual, vMar.Entrada, 0
                            If vEmpresa.QueEmpresa <> 2 Then vMar.Correcto = True
                        Else
                            'retraso, no ha llegado al minimo exigible
                            vMar.HorasTrabajadas = vH.TotalHoras - HoE
                            vMar.HorasIncid = HoE
                            If IncManual = vEmpresa.IncMarcaje Then
                                vMar.IncFinal = vEmpresa.IncRetraso
                            Else
                                If IncManual <> 0 Then
                                    vMar.IncFinal = IncManual
                                Else
                                    vMar.IncFinal = vEmpresa.IncRetraso
                                End If
                            End If
                            vMar.Correcto = True
                            'Ya que despues no quedara constancia ya que sera anulada
                            'para pasar a nominas
                            'Ademas genreamos la incidencia de retraso correspondiente
                            GeneraIncidencia vMar.IncFinal, vMar.Entrada, HoE
                    End If
            End If
        End If
End If


    


    
    'Grabamos el marcaje
    If RevisionEnMarcajes Then
        vMar.Modificar
    Else
        vMar.Agregar
    End If
    
    '-------------------------------------------------------------------------
    'Cerramos y borramos todos los fichajes pasandolos a una tabla de marcajes
    If Not RevisionEnMarcajes Then
        Rss.MoveFirst
        espera 0.2
        Set RFin = New ADODB.Recordset
        RFin.Open "Select max(secuencia) from EntradaMarcajes ", conn, , , adCmdText
        If RFin.EOF Then
            i = 1
            Else
                i = DBLet(RFin.Fields(0), "N") + 1
        End If
        RFin.Close
        
        While Not Rss.EOF
                    
            If Rss!Negativa = 1 Then
                HoraPintar = "-" & Format(Rss!LaHora, "00") & ":" & Format(Rss!Minutos, "00") & ":" & Format(Rss!segundos, "00")
            Else
                If Rss!LaHora > 23 Then
                    HoraPintar = Format(Rss!LaHora, "00") & ":" & Format(Rss!Minutos, "00") & ":" & Format(Rss!segundos, "00")
                Else
                    HoraPintar = Format(Rss!Hora, "hh:nn:ss")
                End If
            End If
            cad = Trim(Rss!lareal)
            cad = Replace(cad, Chr(0), "")
            cad = "'," & Rss!IdInci & ",'" & cad & "'," & Rss!Reloj & ")"
            cad = ",'" & Format(Rss!Fecha, FormatoFecha) & "','" & HoraPintar & cad
            
            cad = i & "," & vMar.idTrabajador & "," & vMar.Entrada & cad
            Debug.Print cad
            conn.Execute vSQL2 & cad
            i = i + 1
            Rss.MoveNext
        Wend
                
        cad = "Delete  from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
        cad = cad & " AND Fecha='" & Format(vMar.Fecha, FormatoFecha) & "'"
        conn.Execute cad
    End If
    'Cerramos los recordsets
    Rss.Close
    
    Set Rss = Nothing
    Set RFin = Nothing
    'Adelante con las operaciones
    
Exit Sub
ErrorProcesaMarcaje_Tipo2:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation

End Sub










'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'   Tipo 3
'--------------------------------------------------------------------------

' El tipo 3 solo controla si es festivo por lo cual todas
' las horas son horas extraso
' y si no es festivo donde todas, repito todas las horas
' son horas trabajadas

'Llevara un variable para ver si estamos revisnado el marcaje.
'Con lo cual en lugar de la tabla de entrada fichajes, sera entradamarcajes

Public Sub ProcesarMarcaje_Tipo3(ByRef vMar As CMarcajes, ByRef vH As CHorarios, RevisionEnMarcajes As Boolean)
Dim Rss As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim NumTikadas As Integer
Dim T1 As Currency
Dim T2 As Currency
Dim kIncidencia As Currency
'Dim TieneIncidencia As Boolean
'Dim MarcajeCorrecto As Boolean
'Dim Exceso As Date
'Dim Retraso As Date
Dim i As Long
'Dim v(3) As Single
Dim cad As String
'Dim HoraH As Date
Dim InciManual As Integer
Dim N As Integer
Dim TotalH As Currency
Dim PrimerTicaje As Date
Dim UltimoTicaje As Date

'Ahora ya tenemos las horas tikadas reflejadas
'Comprobamos las horas en funcion de los horarios
'  y calculamos las horas comprobadas


Set Rss = New ADODB.Recordset

'Seleccionamos todas las horas de este
If RevisionEnMarcajes Then

    cad = "Select * from EntradaMarcajes WHERE idmarcaje=" & vMar.Entrada
    'cad = cad & " AND Fecha='" & Format(vMar.Fecha, FormatoFecha) & "'"
    cad = cad & " ORDER BY Hora"
Else
    cad = "Select * from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
    cad = cad & " AND Fecha='" & Format(vMar.Fecha, FormatoFecha) & "'"
    cad = cad & " ORDER BY Hora"

End If
Rss.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText

If Rss.EOF Then
    'Si no hay ninguna entrada
    Rss.Close
    GoTo ErrorProcesaMarcaje3
End If

    InciManual = 0
    NumTikadas = 0
    While Not Rss.EOF
        NumTikadas = NumTikadas + 1
        Rss.MoveNext
    Wend
   

If (NumTikadas Mod 2) > 0 Then
        'Numero de marcajes impares. No podemos calcular horas
        'trabajadas. Generamos error en marcaje
        vMar.IncFinal = vEmpresa.IncMarcaje
        vMar.HorasIncid = 0
        vMar.HorasTrabajadas = 0
        GeneraIncidencia vEmpresa.IncMarcaje, vMar.Entrada, 0
        vMar.Correcto = False
        Else
            N = NumTikadas \ 2
            TotalH = 0
            TotalParadas2 = 0
            'NUMERO DE MARCAJES PAR
            Rss.MoveFirst
            
            'Lo utilizaremos despues para saber si quitamos minutos de almuerzo
            PrimerTicaje = Rss!Hora
            
            '----------------------------------------------
            
                For i = 1 To N
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
            Next i
                
            'Ahora ya sabemos las horas trabajadas, y las redondeamos
            TotalH = RealizaRedondeo(TotalH)
            
            
            '******************************************************
            'Comprobamos si hay que quitar los minutos del almuerzo
            If vH.DtoAlm > 0 Then
                If LeQuitamosElAmluerzo(Rss, vH) Then TotalParadas2 = vH.DtoAlm
            End If
                
            'Comprobamos si hay que quitar los minutos de la MER
            'Como esta ya en el ultimo
            If vH.DtoMer > 0 Then
                If UltimoTicaje > vH.HoraDtoMer Then TotalParadas2 = TotalParadas2 + vH.DtoMer
            End If
            
            If TotalParadas2 > 0 Then
                If TotalParadas2 > TotalH Then TotalParadas2 = TotalH
                TotalH = TotalH - TotalParadas2
            End If

            
            
            '----------------------------------------------
            '******************************************************
            
            
            
            
            'Asignamos a la incidencia
             vMar.HorasTrabajadas = TotalH
             vMar.HorasDto = TotalParadas2
             
             'Aqui comprobamos si es festivo o no para asignarle los valores correspondientes
             If vH.EsDiaFestivo Then
                vMar.HorasIncid = TotalH
                vMar.IncFinal = vEmpresa.IncHoraExtra
                
            Else
                If InciManual > 0 Then GeneraIncidencia InciManual, vMar.Entrada, 0
                vMar.HorasIncid = 0
                vMar.IncFinal = InciManual
                
            End If
            vMar.Correcto = True
End If 'De DIAFESTIVO

    



    'Grabamos el marcaje
    If RevisionEnMarcajes Then
        vMar.Modificar
    Else
        vMar.Agregar
    End If

    '
    
    '-------------------------------------------------------------------------
    'Cerramos y borramos todos los fichajes pasandolos a una tabla de marcajes
    If Not RevisionEnMarcajes Then
        Rss.MoveFirst
        espera 0.2
        Set RFin = New ADODB.Recordset
        RFin.Open "Select max(secuencia) from EntradaMarcajes ", conn, , , adCmdText
        If RFin.EOF Then
            i = 1
            Else
                i = DBLet(RFin.Fields(0), "N") + 1
        End If
        RFin.Close
        While Not Rss.EOF
            If ProcesandomarcajesHoraOk(Rss) Then
                
                cad = i & "," & vMar.idTrabajador & "," & vMar.Entrada
                cad = cad & ",'" & Format(Rss!Fecha, FormatoFecha) & "','" & Format(Rss!Hora, "hh:mm:ss")
                cad = cad & "'," & Rss!IdInci & ",'" & Format(Rss!HoraReal, "hh:mm:ss") & "'," & Rss!Reloj & ")"
                conn.Execute vSQL2 & cad
            
            Else
                cad = "ERROR " & vbCrLf & vbCrLf & "Trab: " & vMar.idTrabajador & " Secuencia: " & Rss!Secuencia
                cad = cad & "Fecha: " & Format(Rss!Fecha, FormatoFecha)
                MsgBox cad, vbExclamation
            End If
            i = i + 1
            Rss.MoveNext
        Wend
        
        
        
        
        cad = "Delete  from EntradaFichajes WHERE IdTrabajador=" & vMar.idTrabajador
        cad = cad & " AND Fecha='" & Format(vMar.Fecha, FormatoFecha) & "'"
        conn.Execute cad
    End If

'Cerramos los recordsets
Rss.Close

Set Rss = Nothing
Set RFin = Nothing



Exit Sub
ErrorProcesaMarcaje3:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
    
End Sub



Private Function ProcesandomarcajesHoraOk(ByRef R As ADODB.Recordset) As Boolean
Dim H1 As Date
    On Error Resume Next
    H1 = R!Hora
    If Err.Number <> 0 Then
        Err.Clear
        ProcesandomarcajesHoraOk = False
    Else
        ProcesandomarcajesHoraOk = True
    End If
End Function



Public Function LeQuitamosElAmluerzo(ByRef dRS As ADODB.Recordset, ByRef dH As CHorarios) As Boolean
Dim Fin As Boolean
Dim H As Date

    LeQuitamosElAmluerzo = False
    
    dRS.MoveFirst
    Fin = False
    Do
        'Si el primer ticaje, ya es posterior a la hora del almuerzo
        If Format(dRS!Hora, "hh:mm:ss") > dH.HoraDtoAlm Then Exit Function
    
        dRS.MoveNext
        
        If dRS.EOF Then Exit Function
        'Segundo ticaje
        'Ticaje menor. k la hora de almuerzo. Vemos si no ha salido
        If Format(dRS!Hora, "hh:mm:ss") < dH.HoraDtoAlm Then
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



Public Function FijarCodigoIncidenciaGenerada(Codigo As Long)
    idIncidenciaGenerada = Codigo
End Function


Public Function RealizaRedondeo(ByRef T1 As Currency) As Currency
    
   
    'puesto que en los demas el redondeo se realiza
    'revisando marcajes ya que si trabaja las
    'horas que le corresponden no hace falta redondear
    
    Dim Entera As Currency
    Dim resto As Currency
    Dim Divisor As Integer '
    Dim cociente As Integer
    Dim v As Currency
    Dim margen As Currency
    
    'Si no hay que redondear
    T1 = Round(T1, 2)
    RealizaRedondeo = T1
    If vEmpresa.redondeo = 0 Then Exit Function
    
    margen = vEmpresa.MinutosRedondeo
    'Seguimos
    Select Case vEmpresa.redondeo
    Case 2
        Divisor = 25
        
    Case 3
        Divisor = 50
        
    Case Else  'Por si acso los recogemos en ELSE que es decima de punto
        Divisor = 10
        
    End Select
    'Cambiamos el valor de t1
    If T1 < 0 Then
        Entera = Fix(T1)
    Else
        Entera = Int(T1)
    End If
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

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'
' Entradentro:   Para los tipo de marcaje 1.  Comprobara que para cada hora,
'                cuadra con al entrada - salida correspondiente + -  la
'                cortesia
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

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'
' Entradentro:   Para los tipo de marcaje 2.  Comprobara que para el total horas
'                trabajado, cuadra con el total horas a trabajar + - los excesos etc
Public Function EntraDentro2(HoraTotales As Currency, HorasHorario As Currency, Exc As Currency, Ret As Currency) As Currency
Dim Resul
Dim valor As Single

        valor = 0
        'se queda un poco
         If HoraTotales >= HorasHorario Then
               Resul = HoraTotales - (HorasHorario + Exc)
               If Resul > 0 Then
                   'GEneramos la incidencia de hora extra
                   valor = -1 * (HoraTotales - HorasHorario)
               End If
               Else
                   'ha salido antes
                   Resul = HorasHorario - (HoraTotales + Ret)
                   If Resul > 0 Then
                       valor = HorasHorario - HoraTotales
                   End If
        End If
        EntraDentro2 = Round(valor, 2)
End Function




'--------------------------------------------------------
'
'
Public Function YaExistenMarcajes(Cod As Integer, Fecha As Date) As Long
Dim RS As ADODB.Recordset
Dim SQL As String
    YaExistenMarcajes = -1
    Set RS = New ADODB.Recordset
    SQL = "SELECT Entrada" & _
        " FROM Marcajes WHERE " & _
        " IdTrabajador=" & Cod & _
        " AND Fecha=#" & Format(Fecha, "yyyy/mm/dd") & "#"
    RS.Open SQL, conn, , , adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then _
            YaExistenMarcajes = RS.Fields(0)
    End If
    RS.Close
    Set RS = Nothing
End Function





'-------------------------------------------------------------------
'
'           ALZICOOP
'Ahora, para este trabajador generaremos los marcajes definitivos
'Es decir, entrada salida etc
'En vSec tenemos el numero de secuencia para insertar en fichajes
Public Function GeneraUnmarcajeAlzicoop(NTarjeta As String, Codigo As Long, vFecha As Date, ByRef vSec As Long) As Byte
'  ANTES Public Function GeneraUnmarcajeAlzicoop(NTarjeta As String, Codigo As Long, vFecha As Date, ByRef vSec As Long) As Byte
Dim RS As ADODB.Recordset
Dim RsAUX As Recordset
Dim cad As String
Dim i As Integer
Dim H1 As Date
Dim h2 As Date
Dim Entrada As Boolean
Dim Aux As Byte


On Error GoTo ErrGeneraUnmarcajeAlzicoop
GeneraUnmarcajeAlzicoop = 1
cad = "Select * from TipoAlzicoop WHERE Tarjeta='" & NTarjeta & "'"
cad = cad & " AND Fecha=#" & Format(vFecha, "yyyy/mm/dd") & "#  ORDER BY Hora"
Set RS = New ADODB.Recordset
RS.Open cad, conn, , , adCmdText

If Not RS.EOF Then
    Set RsAUX = New ADODB.Recordset
    RsAUX.CursorType = adOpenKeyset
    RsAUX.LockType = adLockOptimistic
    RsAUX.Open "EntradaFichajes", conn, , , adCmdTable

    
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
        RsAUX!HoraReal = RS!Hora
        
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
cad = "DELETE from TipoAlzicoop WHERE Tarjeta='" & NTarjeta & "'"
cad = cad & " AND Fecha=#" & Format(vFecha, "yyyy/mm/dd") & "#"
conn.Execute cad

'Salida
Set RS = Nothing
GeneraUnmarcajeAlzicoop = 0
Exit Function
ErrGeneraUnmarcajeAlzicoop:
    
    
End Function









'-----------------------------------------------------------------
' Funcion:          Generara los marcajes de los dias en los cuales
'               no se ha ticado.  Esta opciones para empresas que
'               generan marcajes todos los dias.
'               Comprobaremos en los dias anterior

Public Sub GeneraEntradasSinMarcajes(Fecha As String, ByRef l1 As Label, ByRef L2 As Label)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim vM As CMarcajes
Dim Lista
Dim vH As CHorarios
Dim k As Integer
Dim J As Integer
Dim RT As ADODB.Recordset


        l1.Caption = "Incidencias continuadas: " & Fecha
        L2.Caption = ""
        l1.Refresh
        L2.Refresh
        DoEvents
        espera 0.3
    
        'Cojeremos los marcajes del dia cuya incidencia este marcada como
        'continuada, y le generaremos los marcajes
        SQL = Format(DateAdd("d", -1, CDate(Fecha)), FormatoFecha)
        SQL = "select idtrabajador,idinci from marcajes,incidencias where idinci=incfinal and fecha='" & SQL
        SQL = SQL & "' and continuada=1 "
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Lista = "|"
        
        'VEo los marcajes de hoy
        SQL = "Select idtrabajador from marcajes where fecha ='" & Format(Fecha, FormatoFecha) & "'"
        Set RT = New ADODB.Recordset
        RT.Open SQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        While Not RS.EOF
            RT.Find "idtrabajador = " & RS!idTrabajador, , adSearchForward, 1
            If RT.EOF Then Lista = Lista & RS!idTrabajador & ":" & RS!IdInci & "|"
            RS.MoveNext
        Wend
        RS.Close
        RT.Close
        Set RT = Nothing
        
        If Len(Lista) < 2 Then
            Lista = ""
            DoEvents
            espera 0.5
            Exit Sub
        End If
            
        'EN lista tenemos los que tienen incidencias continuadas
        'Ahora cojeremos el horario  que tienen
        

        SQL = "Select calendariot.*,idcal from calendariot,trabajadores where fecha = '" & Format(Fecha, FormatoFecha) & "'"
        SQL = SQL & "  and calendariot.idtrabajador=trabajadores.idtrabajador ORDER BY idhorario,idtrabajador"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        'Para generar los marcajes
        Set vM = New CMarcajes
        vM.HorasDto = 0
        vM.HorasIncid = 0
        vM.HorasTrabajadas = 0
        vM.Fecha = Fecha
        vM.Siguiente   'Luego, el contador sera +1 todo el rato
        
        
        
        'Los horarios
        Set vH = New CHorarios
        
        While Not RS.EOF
            'AQUI###
            'Faltara la funcion que dado un calendario mirara si elk horario es festivo al cambio de calendario
            If vH.IdHorario <> RS!IdHorario Then
                l1.Caption = "Dia: " & Fecha & "    Horario: " & RS!IdHorario
                DoEvents
                vH.Leer RS!IdHorario, Fecha, RS!idCal
                espera 0.3
                
            Else
                'Veremos si es festivo
                If vH.idCal <> RS!idCal Then vH.CambioCalendario RS!idCal
            End If
        
                    
            L2.Caption = RS!idTrabajador
            L2.Refresh
                
                
            
                    
    
            
  
            SQL = "|" & RS!idTrabajador & ":"
            k = InStr(1, Lista, SQL)
            If k > 0 Then
                'El trabajador tenia una incidencia continuada
                k = k + Len(SQL)
                J = InStr(k, Lista, "|")
                SQL = Mid(Lista, k, J - k)
            Else
                SQL = ""
            End If
           
                
                If SQL <> "" Then
                    vM.idTrabajador = RS!idTrabajador
                    vM.IdHorario = vH.IdHorario
                    vM.Festivo = vH.EsDiaFestivo
                    vM.Correcto = True
                    
                        'FALTA#####
                        
                    
                        'No tenia incidencia continuada
                        'Con lo cual, si el horario NO pone que es festivo... es un error
'                        If vH.EsDiaFestivo Then
'                            vM.IncFinal = Val(SQL)
'                        Else
                            vM.IncFinal = Val(SQL)
                            
'                        End If
                    
                    vM.Agregar
                    
                    vM.Entrada = vM.Entrada + 1
                End If
            
        
            'Siguiente
            RS.MoveNext
            
        
        Wend
        RS.Close
        l1.Caption = ""
        L2.Caption = ""
        Set vM = Nothing
        Set vH = Nothing
End Sub










Public Sub GeneraLosQueNoHanTicado(Fecha As String, ByRef l1 As Label, ByRef L2 As Label)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim vM As CMarcajes
Dim vH As CHorarios
Dim k As Integer
Dim J As Integer
Dim FESTIVOS As String


        l1.Caption = "Trabajadores que no han fichado: " & Fecha
        L2.Caption = "Paso 1"
        l1.Refresh
        L2.Refresh
        DoEvents
        Set RS = New ADODB.Recordset
        
        SQL = "Select idtrabajador from calendariot where fecha= '" & Format(Fecha, FormatoFecha) & "' AND tipodia=2"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        FESTIVOS = "|"
        While Not RS.EOF
            FESTIVOS = FESTIVOS & RS!idTrabajador & "|"
            RS.MoveNext
        Wend
        RS.Close
    
        SQL = "delete from tmpConMarcajes where codusu = " & vUsu.Codigo
        conn.Execute SQL
        
        SQL = "INSERT INTO tmpConMarcajes select " & vUsu.Codigo & ",idtrabajador from marcajes  where fecha='" & Format(Fecha, FormatoFecha) & "' group by idtrabajador"
        conn.Execute SQL
        
        SQL = "select trabajadores.idtrabajador as c1,tmpConMarcajes.idTrabajador,trabajadores.idcal from trabajadores left join tmpConMarcajes on tmpConMarcajes.idTrabajador=trabajadores.idtrabajador and codusu =" & vUsu.Codigo & " where "
        If vEmpresa.laboral Then
            SQL = SQL & "fecalta <='" & Format(Fecha, FormatoFecha) & "' AND "
        End If
        SQL = SQL & " (fecbaja is null or fecbaja>'" & Format(Fecha, FormatoFecha) & "') AND "
        SQL = SQL & "  tmpConMarcajes.idTrabajador is null"
        
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        
        
        'Version 4.01 de MYSQL
'        SQL = "select idtrabajador from trabajadores where "
'        If vEmpresa.laboral Then
'            SQL = SQL & "fecalta <='" & Format(Fecha, FormatoFecha) & "' AND "
'        End If
'        SQL = SQL & " (fecbaja is null)"
'        SQL = SQL & " and idtrabajador not in (select idtrabajador from marcajes where fecha='" & Format(Fecha, FormatoFecha) & "')"
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS.EOF Then
        
        
        
        
            'Para generar los marcajes
            Set vM = New CMarcajes
            vM.HorasDto = 0
            vM.HorasIncid = 0
            vM.HorasTrabajadas = 0
            vM.Fecha = Fecha
            vM.Siguiente   'Luego, el contador sera +1 todo el rato
            Set miRsAux = New ADODB.Recordset
            Set vH = New CHorarios
            
            While Not RS.EOF
        
                L2.Caption = "Trab: " & RS.Fields(0)
                L2.Refresh
                
                SQL = "Select idhorario from calendariot where fecha='" & Format(Fecha, FormatoFecha) & "'"
                SQL = SQL & " and idtrabajador = " & RS.Fields(0)
                    
                miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    'ERROR. El trabajador no tenia horario asignado. Cogeremos uno cualquiera
                    SQL = ""
                Else
                    SQL = miRsAux!IdHorario
                End If
                miRsAux.Close
                If SQL = "" Then
                    'NO TIENE HORARIO ASIGNADO. COJO UNO
                    miRsAux.Open "Select * from horarios", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    SQL = miRsAux!IdHorario
                    miRsAux.Close
                End If
                
                If vH.IdHorario <> Val(SQL) Then
                    vH.Leer CInt(SQL), Fecha, RS!idCal
                Else
                    If vH.idCal <> RS!idCal Then
                        'Vuelvo a leer faltara esto
                        vH.CambioCalendario RS!idCal
                    End If
                End If
                

                
                
                'No tenia incidencia continuada
                'Con lo cual, si el horario NO pone que es festivo... es un error
                If vH.EsDiaFestivo Then
                    'Si es festivo NO lo meto
                Else
                    'AHORA insertamos el marcaje
                    vM.idTrabajador = RS.Fields(0)
                    vM.IdHorario = vH.IdHorario
                    If InStr(1, FESTIVOS, "|" & RS!C1 & "|") > 0 Then
                        'ESTA DE VACACIONES
                        vM.Festivo = True
                        vM.IncFinal = vEmpresa.IncVacaciones
                        vM.Correcto = True
                        
                    Else
                        GeneraIncidencia vEmpresa.IncMarcaje, vM.Entrada, 0
                        vM.Festivo = False
                        vM.IncFinal = vEmpresa.IncMarcaje
                        vM.Correcto = False
                    End If
            
            
                    vM.Agregar
                    
                
                    
                    vM.Entrada = vM.Entrada + 1
                End If
                'Siguiente
                RS.MoveNext
                
                
            Wend
            RS.Close

        End If

        Set RS = Nothing
        Set miRsAux = Nothing
        Set vM = Nothing
        Set vH = Nothing
End Sub



'



'''''Private Sub ConversionRedondeo(ByRef T1 As Currency, ByRef T2 As Currency)
'''''Dim T3 As Currency
'''''Dim Entera As Currency
'''''Dim resto As Currency
'''''Dim Divisor As Integer '
'''''Dim margen As Currency
'''''Dim cociente As Integer
'''''Dim v As Currency
'''''
'''''
'''''
'''''
'''''    'Seguimos
'''''    Select Case redondeo
'''''    Case 2
'''''        Divisor = 25
'''''        margen = 18
'''''    Case 3
'''''        Divisor = 50
'''''        margen = 38
'''''    Case Else  'Por si acso los recogemos en ELSE que es decima de punto
'''''        Divisor = 10
'''''        margen = 3
'''''    End Select
'''''    T3 = T1 + T2
'''''    'Cambiamos el valor de t1
'''''    Entera = Int(T1)
'''''    resto = Round((T1 - Entera) * 100, 0)
'''''
'''''
'''''    v = resto Mod Divisor
'''''    cociente = resto \ Divisor
'''''    'No se redondea nunca hacia arriba, luego la instrucciones van comentadas
'''''    If v >= margen Then
'''''            cociente = cociente + 1
'''''    End If
'''''    v = cociente * Divisor  'Resto redondeado
'''''    v = v / 100
'''''    T1 = Entera + v
'''''    T2 = Round(T3 - T1, 2)
'''''    T2 = Abs(T2)
'''''
'''''End Sub
'''''
'''''




Public Sub EntradasRepetidasProceso(ByRef lbl As Label)
Dim RFin As ADODB.Recordset
Dim idTrabajador As Long
Dim CadInci As String
Dim Fecha As Date
Dim Hora As Date
Dim Diferencia As Long
Dim paso As Byte
Dim RelojDistintos As Byte


    

    If vEmpresa.Repeticion_ <= 0 Then Exit Sub
    
    RelojDistintos = 1
    If vEmpresa.Reloj2 > 0 Then RelojDistintos = 2
        
    
    
    lbl.Caption = "Entradas duplicadas"
    lbl.Refresh
    Set RFin = New ADODB.Recordset
        
    For paso = 1 To RelojDistintos
        
        lbl.Caption = "Repetidas: " & paso
        lbl.Refresh
        espera 0.25
        
        'Ya tenemos a partir de k fecha, y con k cadencia vamos a eliminar repetidos
        CadInci = "Select * from Entradafichajes WHERE hora <='23:59:59'"
        If RelojDistintos > 1 Then
            'IREMOS POR RELOJ
            CadInci = CadInci & " AND reloj =" & paso - 1
            
        End If
        CadInci = CadInci & " ORDER BY idTrabajador,Fecha,Hora"
        RFin.Open CadInci, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        idTrabajador = 0 'Tendremos el codigo del trabajador
        CadInci = "DELETE from EntradaFichajes WHERE Secuencia = "
        While Not RFin.EOF
           
        
            If RFin!idTrabajador <> idTrabajador Then
                
                lbl.Caption = "Trabajador: " & RFin!idTrabajador
                lbl.Refresh
                
                'Nuevo trabajador
                idTrabajador = RFin!idTrabajador
                Fecha = RFin!Fecha
                Hora = Format(RFin!Hora, "hh:mm:ss")
            Else
                'Es el mismo trabajador.
                'Veamos la fecha
                If RFin!Fecha <> Fecha Then
                    Fecha = RFin!Fecha
                    Hora = Format(RFin!Hora, "hh:mm:ss")
                Else
                    'MISMO TRABAJADOR , MISMA FECHA
                    Diferencia = DateDiff("n", Hora, Format(RFin!Hora, "hh:mm:ss"))
                    If Diferencia >= vEmpresa.Repeticion_ Then
                        'Las horas se diferencian. NO elimino
                        Hora = Format(RFin!Hora, "hh:mm:ss")
                    Else
                        'SI elimino
                        conn.Execute CadInci & RFin!Secuencia
                    End If
                End If
            End If
            'Siguiente
            RFin.MoveNext
        Wend
        RFin.Close
    
    Next
    Set RFin = Nothing




End Sub

'Seran las horas trabajadas desde las 0:00 hasta las 4:00 De momento NO esta en parametros. Va "a piñon"
'Estas horas son del dia de antes!!!!
'Se han quedado a trarabjar hasta mas alla de la medianoche (antes de las 4;00)
Public Sub HorasNocturnas(ByRef lbl As Label)
Dim RT As ADODB.Recordset
Dim cad As String


    On Error GoTo eHorasNocturnas


    If Not vEmpresa.AcabaJornadaDiaSiguiente Then Exit Sub  'lo sque no lleven horas estas seguimos
        

    If Not lbl Is Nothing Then
        lbl.Caption = "Horas nocturnas"
        lbl.Refresh
    End If

    
    cad = "Select * from entradafichajes where hora>='0:00:00' and hora<'" & vEmpresa.MaximaHoraDiaSiguiente & "'"
    
    cad = cad & " ORDER BY fecha,hora"
    Set RT = New ADODB.Recordset
    RT.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        If Not lbl Is Nothing Then
            lbl.Caption = RT!Fecha & " " & Format(RT!Hora, "hh:mm")
            lbl.Refresh
        End If
        cad = RT!Fecha
        cad = "UPDATE entradafichajes SET "
        cad = cad & " hora = ADDTIME(hora , '24:00:00' ) "
        cad = cad & ",horareal = ADDTIME(horareal , '24:00:00' ) "
        cad = cad & ",fecha = DATE_ADD(fecha, INTERVAL -1 DAY)"
        cad = cad & " WHERE Secuencia =" & RT!Secuencia
        conn.Execute cad
        RT.MoveNext
    Wend
    RT.Close
    
eHorasNocturnas:
    If Err.Number <> 0 Then MuestraError Err.Number, "HorasNocturnas", Err.Description
    Set RT = Nothing
End Sub


