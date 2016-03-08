# MySQL-Front Dump 2.5
#
# Host: localhost   Database: aripres1
# --------------------------------------------------------
# Server version 4.1.18-nt


#
# Table structure for table 'bajas'
#

CREATE TABLE bajas (
  idTrab int(11) NOT NULL default '0',
  idTipobaja smallint(6) NOT NULL default '0',
  Fechabaja date NOT NULL default '0000-00-00',
  FechaAlta date default NULL,
  PRIMARY KEY  (idTrab,Fechabaja),
  KEY Baja_idTipobaja (idTipobaja)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'calendario'
#

CREATE TABLE calendario (
  idcal smallint(1) unsigned NOT NULL default '0',
  descripcion varchar(50) default NULL,
  PRIMARY KEY  (idcal)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'calendariof'
#

CREATE TABLE calendariof (
  idcal smallint(1) unsigned NOT NULL default '0',
  fecha date NOT NULL default '0000-00-00',
  descripcion varchar(50) default NULL,
  PRIMARY KEY  (idcal,fecha)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'calendariol'
#

CREATE TABLE calendariol (
  idcal smallint(1) unsigned NOT NULL default '0',
  fecha date NOT NULL default '0000-00-00',
  idhorario smallint(6) NOT NULL default '0',
  PRIMARY KEY  (idcal,fecha)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'calendariot'
#

CREATE TABLE calendariot (
  idtrabajador int(1) NOT NULL default '0',
  fecha date NOT NULL default '0000-00-00',
  idhorario smallint(6) NOT NULL default '0',
  TipoDia tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (idtrabajador,fecha)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'categorias'
#

CREATE TABLE categorias (
  IdCategoria int(11) NOT NULL default '0',
  nomCategoria char(50) default NULL,
  Importe1 decimal(10,2) default '0.00',
  Importe2 decimal(10,2) default '0.00',
  Importe3 decimal(10,2) default '0.00',
  PRIMARY KEY  (IdCategoria)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;



#
# Table structure for table 'diasemana'
#

CREATE TABLE diasemana (
  IdDia int(11) NOT NULL default '0',
  TextoDia char(15) default NULL,
  PRIMARY KEY  (IdDia)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'empresas'
#

CREATE TABLE empresas (
  IdEmpresa int(11) NOT NULL default '0',
  NomEmpresa varchar(50) default NULL,
  DirEmpresa varchar(50) default NULL,
  PobEmpresa varchar(50) default NULL,
  ProvEmpresa varchar(50) default NULL,
  TelEmpresa varchar(50) default NULL,
  CodPosEmpresa varchar(5) default NULL,
  MaxRetraso decimal(10,2) default '0.00',
  MaxExceso decimal(10,2) default '0.00',
  IncHoraExtra int(11) default '0',
  IncRetraso smallint(6) default '0',
  IncMarcaje smallint(6) default '0',
  IncVacaciones smallint(6) default '0',
  IncTarjError smallint(6) default '0',
  IncHoraExceso smallint(11) default '0',
  CIF varchar(15) default NULL,
  MinutosRedondeo int(11) default '0',
  AjusteEntrada smallint(11) default '0',
  AjusteSalida smallint(50) default NULL,
  HorasJornada smallint(6) default NULL,
  RecuperacionDias tinyint(1) default '0',
  Entidad varchar(4) default NULL,
  Sucursal varchar(4) default NULL,
  CodControl char(2) default NULL,
  Cuenta varchar(10) default NULL,
  Repeticion int(11) default '0',
  AplicaAntiguedadHN tinyint(1) default '0',
  AplicaAntiguedadHC tinyint(1) default '0',
  AbonosSeparados tinyint(1) default '0',
  IRPF decimal(10,2) default '0.00',
  EmpresaHoraExtra tinyint(1) default NULL,
  NominaAutomatica tinyint(1) default NULL,
  HorarioNocturno tinyint(1) default NULL,
  redondeo tinyint(3) unsigned default NULL,
  laboral tinyint(3) unsigned default NULL,
  produccion tinyint(3) unsigned default NULL,
  imgtrabaj tinyint(3) unsigned NOT NULL default '0',
  reloj tinyint(3) unsigned default '0',
  todoslosdias tinyint(3) unsigned default '0',
  fechainicio date default NULL,
  servidor varchar(30) default NULL,
  usuario varchar(30) default NULL,
  pass varchar(30) default NULL,
  configreloj varchar(255) default NULL,
  Pathproces varchar(255) default NULL,
  Nomproces varchar(35) default NULL,
  PRIMARY KEY  (IdEmpresa)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;



#
# Table structure for table 'entradafichaj2'
#

CREATE TABLE entradafichaj2 (
  Secuencia int(11) NOT NULL default '0',
  idTrabajador int(11) default '0',
  Fecha date default NULL,
  Hora time default NULL,
  idInci smallint(6) default '0',
  HoraReal time default NULL,
  PRIMARY KEY  (Secuencia),
  KEY Entr_idTrabajador (idTrabajador),
  KEY Entr_idInci (idInci)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'entradafichajes'
#

CREATE TABLE entradafichajes (
  Secuencia int(11) NOT NULL default '0',
  idTrabajador int(11) default '0',
  Fecha date default NULL,
  Hora time default NULL,
  idInci smallint(6) default '0',
  HoraReal time default NULL,
  PRIMARY KEY  (Secuencia),
  KEY Entr_idTrabajador (idTrabajador),
  KEY Entr_idInci (idInci)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'entradamarcajes'
#

CREATE TABLE entradamarcajes (
  Secuencia int(11) NOT NULL default '0',
  idTrabajador int(11) default '0',
  idMarcaje int(11) default '0',
  Fecha date default NULL,
  Hora time default NULL,
  idInci smallint(6) default '0',
  HoraReal time default NULL,
  PRIMARY KEY  (Secuencia),
  KEY Entr_idTrabajador (idTrabajador),
  KEY Entr_idMarcaje (idMarcaje),
  KEY Entr_idInci (idInci)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'errores'
#

CREATE TABLE errores (
  Id int(11) NOT NULL default '0',
  Campo1 decimal(10,0) default NULL,
  Campo2 date default NULL,
  Campo3 char(255) default NULL,
  PRIMARY KEY  (Id)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'errortarjetas'
#

CREATE TABLE errortarjetas (
  Secuencia int(11) NOT NULL default '0',
  numTarjeta char(50) default NULL,
  Fecha date default NULL,
  Hora time default NULL,
  idInci smallint(6) default '0',
  Error char(200) default NULL,
  PRIMARY KEY  (Secuencia),
  KEY Erro_idInci (idInci)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'festivos'
#

CREATE TABLE festivos (
  Id int(11) NOT NULL default '0',
  IdHorario smallint(6) default '0',
  Anyo smallint(6) default '0',
  Fecha date NOT NULL default '0000-00-00',
  Descripcion char(50) default NULL,
  PRIMARY KEY  (Id),
  KEY Fest_IdHorario (IdHorario),
  KEY F_IdHorario (IdHorario),
  CONSTRAINT festivos_ibfk_1 FOREIGN KEY (IdHorario) REFERENCES horarios (IdHorario)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;



#
# Table structure for table 'horarios'
#

CREATE TABLE horarios (
  IdHorario smallint(6) NOT NULL default '0',
  NomHorario char(50) NOT NULL default '',
  TotalHoras decimal(10,2) default NULL,
  DtoAlm decimal(10,2) default NULL,
  HoraDtoAlm time default NULL,
  DtoMer decimal(10,2) default NULL,
  HoraDtoMer time default NULL,
  RecuperaSabados tinyint(1) default NULL,
  Rectificar tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (IdHorario)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;



#
# Table structure for table 'horastrabajadas'
#

CREATE TABLE horastrabajadas (
  id int(11) NOT NULL default '0',
  Empresa char(50) default NULL,
  Seccion char(50) default NULL,
  Fecha date default NULL,
  Nombre char(50) default NULL,
  HorasN decimal(10,0) default '0',
  HorasE decimal(10,0) default '0',
  HorasF decimal(10,0) default '0',
  PRIMARY KEY  (id)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'incidencias'
#

CREATE TABLE incidencias (
  IdInci smallint(6) NOT NULL default '0',
  NomInci char(50) default NULL,
  Continuada tinyint(1) default NULL,
  ExcesoDefecto tinyint(1) default NULL,
  PRIMARY KEY  (IdInci)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'incidenciasgeneradas'
#

CREATE TABLE incidenciasgeneradas (
  Id int(11) NOT NULL default '0',
  EntradaMarcaje int(11) default '0',
  Incidencia smallint(6) default '0',
  horas decimal(10,2) default '0.00',
  PRIMARY KEY  (Id)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'incidenciasgeneradashco'
#

CREATE TABLE incidenciasgeneradashco (
  Id int(11) NOT NULL default '0',
  Traspaso tinyint(1) NOT NULL default '0',
  EntradaMarcajehco int(11) default '0',
  Incidencia smallint(6) default '0',
  horas decimal(10,2) default '0.00',
  PRIMARY KEY  (Id,Traspaso)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'incidenciashco'
#

CREATE TABLE incidenciashco (
  IdInci smallint(6) NOT NULL default '0',
  Traspaso tinyint(1) NOT NULL default '0',
  NomInci char(50) default NULL,
  Continuada tinyint(1) default NULL,
  ExcesoDefecto tinyint(1) default NULL,
  PRIMARY KEY  (IdInci,Traspaso)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'jornadassemanales'
#

CREATE TABLE jornadassemanales (
  idTrabajador int(11) NOT NULL default '0',
  Fecha date NOT NULL default '0000-00-00',
  HorasOfi decimal(10,2) default '0.00',
  DiasOfi int(11) default '0',
  HN decimal(10,2) default '0.00',
  HC decimal(10,2) default '0.00',
  Dias int(11) default '0',
  BolsaAntes decimal(10,2) default '0.00',
  BolsaDespues decimal(10,2) default '0.00',
  HE decimal(10,2) default '0.00',
  PRIMARY KEY  (idTrabajador,Fecha)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'marcajes'
#

CREATE TABLE marcajes (
  Entrada int(11) NOT NULL default '0',
  idTrabajador int(11) default '0',
  Fecha date default NULL,
  Correcto tinyint(1) default NULL,
  IncFinal smallint(6) default '0',
  HorasTrabajadas decimal(10,2) default '0.00',
  HorasIncid decimal(10,2) default '0.00',
  idHorario smallint(5) unsigned NOT NULL default '0',
  HorasDto decimal(10,2) NOT NULL default '0.00',
  Festivo tinyint(3) unsigned NOT NULL default '0',
  Baja tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (Entrada),
  KEY F_idTrabajador (idTrabajador),
  KEY id_Horario (idHorario),
  KEY inci (IncFinal),
  CONSTRAINT marcajes_ibfk_1 FOREIGN KEY (idTrabajador) REFERENCES trabajadores (IdTrabajador)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;



#
# Table structure for table 'marcajeshco'
#

CREATE TABLE marcajeshco (
  Entrada int(11) NOT NULL default '0',
  Traspaso tinyint(1) NOT NULL default '0',
  idTrabajador int(11) default '0',
  Fecha date default NULL,
  Correcto tinyint(1) default NULL,
  IncFinal smallint(6) default '0',
  HorasTrabajadas decimal(10,2) default '0.00',
  HorasIncid decimal(10,2) default '0.00',
  idHorario smallint(5) unsigned NOT NULL default '0',
  HorasDto decimal(10,2) NOT NULL default '0.00',
  Festivo tinyint(3) unsigned NOT NULL default '0',
  Baja tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (Entrada,Traspaso),
  KEY F_idTrabajador (idTrabajador),
  KEY id_Horario (idHorario),
  KEY inci (IncFinal)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'marcajeskimaldi'
#

CREATE TABLE marcajeskimaldi (
  Nodo int(11) default '0',
  Fecha date default NULL,
  Hora time default NULL,
  TipoMens char(50) default NULL,
  Marcaje char(50) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'marcajeskimaldierror'
#

CREATE TABLE marcajeskimaldierror (
  Nodo int(11) default '0',
  Fecha date default NULL,
  Hora time default NULL,
  TipoMens char(50) default NULL,
  Marcaje char(50) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'modificarfichajes'
#

CREATE TABLE modificarfichajes (
  idhorario smallint(6) NOT NULL default '0',
  Inicio time NOT NULL default '00:00:00',
  Fin time NOT NULL default '00:00:00',
  Modificada time NOT NULL default '00:00:00',
  PRIMARY KEY  (idhorario,Inicio)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'nominas'
#

CREATE TABLE nominas (
  idTrabajador int(11) NOT NULL default '0',
  Fecha date NOT NULL default '0000-00-00',
  Dias tinyint(1) default '0',
  HN decimal(10,2) default '0.00',
  HC decimal(10,2) default '0.00',
  PLUS decimal(10,2) default '0.00',
  Anticipos decimal(10,2) default '0.00',
  BolsaAntes decimal(10,2) default '0.00',
  BolsaDespues decimal(10,2) default '0.00',
  HP decimal(10,2) default '0.00',
  PRIMARY KEY  (idTrabajador,Fecha)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'pagos'
#

CREATE TABLE pagos (
  Trabajador int(11) NOT NULL default '0',
  Fecha date NOT NULL default '0000-00-00',
  Tipo tinyint(1) NOT NULL default '0',
  Pagado tinyint(1) default NULL,
  Importe decimal(10,2) default '0.00',
  Observaciones char(50) default NULL,
  PRIMARY KEY  (Trabajador,Fecha,Tipo)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'secciones'
#

CREATE TABLE secciones (
  IdSeccion smallint(6) NOT NULL default '0',
  Nombre char(50) default NULL,
  idCal smallint(6) default '0',
  ControlEmpleados tinyint(1) default '0',
  Nominas tinyint(1) default NULL,
  PRIMARY KEY  (IdSeccion)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;



#
# Table structure for table 'stipocontrol'
#

CREATE TABLE stipocontrol (
  tipocontrol smallint(1) unsigned NOT NULL default '0',
  desccontrol varchar(30) default NULL,
  PRIMARY KEY  (tipocontrol)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'subhorarios'
#

CREATE TABLE subhorarios (
  IdHorario smallint(6) NOT NULL default '0',
  DiaSemana tinyint(1) NOT NULL default '0',
  Festivo tinyint(1) default NULL,
  HEntrada1 time default NULL,
  HSalida1 time default NULL,
  HEntrada2 time default NULL,
  HSalida2 time default NULL,
  N_Tikadas tinyint(1) default '0',
  HorasDia decimal(10,0) default '0',
  DiaNomina decimal(10,2) default '0.00',
  PRIMARY KEY  (DiaSemana,IdHorario),
  KEY SubH_IdHorario (IdHorario),
  KEY F_IdHorario (IdHorario)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;



#
# Table structure for table 'tareas'
#

CREATE TABLE tareas (
  idTarea int(11) NOT NULL default '0',
  Descripcion char(50) default NULL,
  Tarjeta char(50) default NULL,
  Tipo tinyint(1) default '0',
  PRIMARY KEY  (idTarea)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tareasrealizadas'
#

CREATE TABLE tareasrealizadas (
  Trabajador int(11) default '0',
  Fecha date default NULL,
  HoraInicio time default NULL,
  HoraFin time default NULL,
  Tarea int(11) default '0',
  HorasTrabajadas decimal(10,2) default '0.00',
  Horas1 decimal(10,2) default '0.00',
  Horas2 decimal(10,2) default '0.00',
  Horas3 decimal(10,2) default '0.00',
  Importe1 decimal(10,2) default '0.00',
  Importe2 decimal(10,2) default '0.00',
  Importe3 decimal(10,2) default '0.00',
  Total decimal(10,2) default '0.00'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'temporalfichajes'
#

CREATE TABLE temporalfichajes (
  Secuencia int(11) NOT NULL default '0',
  numTarjeta char(50) default NULL,
  Fecha date default NULL,
  Hora time default NULL,
  idInci smallint(6) default '0',
  PRIMARY KEY  (Secuencia),
  KEY Temp_idInci (idInci)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'timagen'
#

CREATE TABLE timagen (
  idTrabajador int(3) unsigned NOT NULL default '0',
  imagen longblob,
  PRIMARY KEY  (idTrabajador)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Imagenes';



#
# Table structure for table 'tipoalzicoop'
#

CREATE TABLE tipoalzicoop (
  Secuencia int(11) NOT NULL default '0',
  Fecha date default NULL,
  Hora time default NULL,
  Tarjeta char(50) default NULL,
  seccion char(3) default NULL,
  tecla char(3) default NULL,
  HoraReal time default NULL,
  PRIMARY KEY  (Secuencia)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tipobaja'
#

CREATE TABLE tipobaja (
  idbaja smallint(6) NOT NULL default '0',
  descbaja char(50) default NULL,
  PRIMARY KEY  (idbaja)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tipocontrato'
#

CREATE TABLE tipocontrato (
  idContrato smallint(6) NOT NULL default '0',
  DescContrato char(50) default NULL,
  PRIMARY KEY  (idContrato)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tipopago'
#

CREATE TABLE tipopago (
  idTipopago tinyint(1) NOT NULL default '0',
  Descripcion char(15) default NULL,
  PRIMARY KEY  (idTipopago)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmp_proc1'
#

CREATE TABLE tmp_proc1 (
  codusu smallint(6) NOT NULL default '0',
  codigo int(11) NOT NULL default '0',
  valor1 int(11) NOT NULL default '0',
  fecha date NOT NULL default '0000-00-00',
  texto1 varchar(30) default NULL,
  texto2 varchar(30) default NULL,
  texto3 varchar(30) default NULL,
  PRIMARY KEY  (codusu,codigo)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpcambiohor'
#

CREATE TABLE tmpcambiohor (
  Trabajador int(11) NOT NULL default '0',
  codusu tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (Trabajador)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpcombinada'
#

CREATE TABLE tmpcombinada (
  IdTrabajador int(11) NOT NULL default '0',
  Fecha date NOT NULL default '0000-00-00',
  HT decimal(10,2) default '0.00',
  HE decimal(10,2) default '0.00',
  H1 time default NULL,
  H2 time default NULL,
  H3 time default NULL,
  H4 time default NULL,
  H5 time default NULL,
  H6 time default NULL,
  H7 time default NULL,
  H8 time default NULL,
  codusu int(11) NOT NULL default '0',
  idinci smallint(5) unsigned default '0',
  HR decimal(5,2) default NULL,
  PRIMARY KEY  (IdTrabajador,Fecha,codusu)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpconincres'
#

CREATE TABLE tmpconincres (
  id int(11) NOT NULL default '0',
  idempresa int(11) default '0',
  NomEmpresa char(50) default NULL,
  idIncidencia int(11) default '0',
  NomIncidencia char(50) default NULL,
  IdTrabajador int(11) default '0',
  NomTrabajador char(50) default NULL,
  HE decimal(10,0) default '0',
  HD decimal(10,0) default '0',
  ExcesoDefecto tinyint(1) default NULL,
  fecha date default NULL,
  Seccion char(50) default NULL,
  PRIMARY KEY  (id),
  KEY tmpC_idempresa (idempresa),
  KEY tmpC_idIncidencia (idIncidencia),
  KEY tmpC_IdTrabajador (IdTrabajador)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpdatosmes'
#

CREATE TABLE tmpdatosmes (
  Mes tinyint(1) NOT NULL default '0',
  Trabajador int(11) NOT NULL default '0',
  MesHoras decimal(10,2) default '0.00',
  MesDias tinyint(1) default '0',
  DiasTrabajados int(11) default '0',
  HorasN decimal(10,2) default '0.00',
  HorasC decimal(10,2) default '0.00',
  HorasPlus decimal(10,2) default '0.00',
  HorasT decimal(10,2) default '0.00',
  HorasE decimal(10,2) default '0.00',
  SaldoH decimal(10,2) default '0.00',
  SaldoDias smallint(6) default '0',
  bolsaAntes decimal(10,2) default '0.00',
  bolsaDespues decimal(10,2) default '0.00',
  HorasPeriodo decimal(10,2) default '0.00',
  DiasPeriodo smallint(6) default '0',
  Extras decimal(10,2) default '0.00',
  Anticipos decimal(10,2) default '0.00',
  PLUS int(11) default '0',
  ExtrasPeriodo decimal(10,2) default '0.00',
  codusu mediumint(9) NOT NULL default '0',
  PRIMARY KEY  (Mes,Trabajador,codusu)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpdiastrabajinci'
#

CREATE TABLE tmpdiastrabajinci (
  idtrabajador int(11) NOT NULL default '0',
  incidencia smallint(6) NOT NULL default '0',
  horas decimal(10,2) default NULL,
  codusu mediumint(9) NOT NULL default '0',
  dias smallint(5) unsigned NOT NULL default '0',
  PRIMARY KEY  (codusu,idtrabajador,incidencia)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpfechas'
#

CREATE TABLE tmpfechas (
  fechas date default NULL,
  Horario char(255) default NULL,
  Descr char(255) default NULL,
  idHor int(11) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmphoras'
#

CREATE TABLE tmphoras (
  trabajador int(11) NOT NULL default '0',
  HorasT decimal(10,2) default '0.00',
  HorasC decimal(10,2) default '0.00',
  HorasE decimal(10,2) default '0.00',
  Dias tinyint(1) default '0',
  PRIMARY KEY  (trabajador)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmphorasmeshorario'
#

CREATE TABLE tmphorasmeshorario (
  idHorario smallint(6) NOT NULL default '0',
  Horas decimal(10,2) default '0.00',
  Dias smallint(6) default '0',
  PRIMARY KEY  (idHorario)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpinformehorasmes'
#

CREATE TABLE tmpinformehorasmes (
  idTrabajador int(11) NOT NULL default '0',
  Asesoria char(10) default NULL,
  Nombre char(35) default NULL,
  HT decimal(10,2) default '0.00',
  HN decimal(10,2) default '0.00',
  DT int(11) default '0',
  H1 char(5) default NULL,
  H2 char(5) default NULL,
  H3 char(5) default NULL,
  H4 char(5) default NULL,
  H5 char(5) default NULL,
  H6 char(5) default NULL,
  H7 char(5) default NULL,
  H8 char(5) default NULL,
  H9 char(5) default NULL,
  H10 char(5) default NULL,
  H11 char(5) default NULL,
  H12 char(5) default NULL,
  H13 char(5) default NULL,
  H14 char(5) default NULL,
  H15 char(5) default NULL,
  H16 char(5) default NULL,
  H17 char(5) default NULL,
  H18 char(5) default NULL,
  H19 char(5) default NULL,
  H20 char(5) default NULL,
  H21 char(5) default NULL,
  H22 char(5) default NULL,
  H23 char(5) default NULL,
  H24 char(5) default NULL,
  H25 char(5) default NULL,
  H26 char(5) default NULL,
  H27 char(5) default NULL,
  H28 char(5) default NULL,
  H29 char(5) default NULL,
  H30 char(5) default NULL,
  H31 char(5) default NULL,
  C1 char(5) default NULL,
  C2 char(5) default NULL,
  C3 char(5) default NULL,
  C4 char(5) default NULL,
  C5 char(5) default NULL,
  C6 char(5) default NULL,
  C7 char(5) default NULL,
  C8 char(5) default NULL,
  C9 char(5) default NULL,
  C10 char(5) default NULL,
  C11 char(5) default NULL,
  C12 char(5) default NULL,
  C13 char(5) default NULL,
  C14 char(5) default NULL,
  C15 char(5) default NULL,
  C16 char(5) default NULL,
  C17 char(5) default NULL,
  C18 char(5) default NULL,
  C19 char(5) default NULL,
  C20 char(5) default NULL,
  C21 char(5) default NULL,
  C22 char(5) default NULL,
  C23 char(5) default NULL,
  C24 char(5) default NULL,
  C25 char(5) default NULL,
  C26 char(5) default NULL,
  C27 char(5) default NULL,
  C28 char(5) default NULL,
  C29 char(5) default NULL,
  C30 char(5) default NULL,
  C31 char(5) default NULL,
  PRIMARY KEY  (idTrabajador)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpmarcajes'
#

CREATE TABLE tmpmarcajes (
  Entrada int(11) NOT NULL default '0',
  idTrabajador int(11) default '0',
  Fecha date default NULL,
  Correcto tinyint(1) default NULL,
  IncFinal smallint(6) default '0',
  HorasTrabajadas decimal(10,0) default '0',
  HorasIncid decimal(10,0) default '0',
  PRIMARY KEY  (Entrada),
  KEY tmpM_idTrabajador (idTrabajador),
  KEY tmpM_HorasIncid (HorasIncid)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpmarcajeskimaldi'
#

CREATE TABLE tmpmarcajeskimaldi (
  Nodo int(11) default '0',
  Fecha date default NULL,
  Hora time default NULL,
  TipoMens char(50) default NULL,
  Marcaje char(50) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpnorma34'
#

CREATE TABLE tmpnorma34 (
  CodSoc int(11) NOT NULL default '0',
  Nombre char(50) default NULL,
  Banco1 char(4) default NULL,
  Banco2 char(4) default NULL,
  Banco3 char(2) default NULL,
  Banco4 char(10) default NULL,
  Domicilio char(35) default NULL,
  Codpos char(5) default NULL,
  Poblacion char(50) default NULL,
  Concepto char(50) default NULL,
  Importe decimal(10,2) NOT NULL default '0.00',
  tipo tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (CodSoc,tipo)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmpnotrabajo'
#

CREATE TABLE tmpnotrabajo (
  idTra int(11) default NULL,
  idFech date default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmppagosmes'
#

CREATE TABLE tmppagosmes (
  idTrabajador int(11) NOT NULL default '0',
  Nombre char(50) default NULL,
  HT decimal(10,2) default '0.00',
  Importe1 decimal(10,2) default '0.00',
  HC decimal(10,2) default '0.00',
  Importe2 decimal(10,2) default '0.00',
  IRPF char(50) default NULL,
  SS char(50) default NULL,
  Neto decimal(10,2) default '0.00',
  Pagos decimal(10,2) default '0.00',
  Ingresar decimal(10,2) default '0.00',
  Bruto decimal(10,2) default '0.00',
  PrecioHora1 decimal(10,2) default '0.00',
  PRIMARY KEY  (idTrabajador)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmppresencia'
#

CREATE TABLE tmppresencia (
  Id int(11) NOT NULL default '0',
  NomTrabajador char(50) default NULL,
  NomEmpresa char(50) default NULL,
  Fecha date default NULL,
  H1 time default NULL,
  H2 time default NULL,
  H3 time default NULL,
  H4 time default NULL,
  H5 time default NULL,
  H6 time default NULL,
  H7 time default NULL,
  H8 time default NULL,
  Incidencias char(50) default NULL,
  Seccion char(50) default NULL,
  idtra int(5) unsigned NOT NULL default '0',
  codusu smallint(5) unsigned NOT NULL default '0',
  semana tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (Id,codusu)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmptareaactual'
#

CREATE TABLE tmptareaactual (
  Trabajador int(11) NOT NULL default '0',
  Tarea int(11) default '0',
  Hora time default NULL,
  PRIMARY KEY  (Trabajador)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmptareasrealizadas'
#

CREATE TABLE tmptareasrealizadas (
  Fecha date default NULL,
  Hora time default NULL,
  trabajador int(11) default '0',
  Tarea int(11) default '0'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'tmptareasrealizadas2'
#

CREATE TABLE tmptareasrealizadas2 (
  Fecha date default NULL,
  Hora time default NULL,
  trabajador int(11) default NULL,
  Tarea int(11) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



#
# Table structure for table 'trabajadores'
#

CREATE TABLE trabajadores (
  IdTrabajador int(11) NOT NULL default '0',
  NumTarjeta varchar(50) default NULL,
  idCategoria int(11) default '0',
  NomTrabajador varchar(50) NOT NULL default '',
  DomTrabajador varchar(50) default NULL,
  PobTrabajador varchar(50) default NULL,
  ProvTrabajador varchar(50) default NULL,
  CodPosTrabajador varchar(5) default NULL,
  TelTrabajador varchar(50) default NULL,
  MovTrabajador varchar(50) default NULL,
  FecAlta date default NULL,
  FecBaja date default NULL,
  FecEntVac date default NULL,
  FecSalVac date default NULL,
  Control tinyint(1) NOT NULL default '0',
  InciCont tinyint(1) default '0',
  numSS varchar(50) default NULL,
  numMat varchar(50) default NULL,
  numDNI varchar(50) default NULL,
  Seccion smallint(6) NOT NULL default '0',
  PorcAntiguedad decimal(10,2) default '0.00',
  PorcSS decimal(10,2) default '0.00',
  PorcIRPF decimal(10,2) default '0.00',
  TipoContrato smallint(6) default '0',
  pagobancario tinyint(1) default NULL,
  entidad varchar(4) default NULL,
  oficina varchar(4) default NULL,
  controlcta char(2) default NULL,
  cuenta varchar(10) default NULL,
  bolsahoras decimal(10,2) default '0.00',
  idAsesoria varchar(50) default NULL,
  Antiguedad date default NULL,
  ControlNomina tinyint(1) default '0',
  sexo int(11) default '0',
  bolsaNETO decimal(10,2) default NULL,
  bolsaBRUTO decimal(10,2) default NULL,
  email varchar(100) default NULL,
  idcal smallint(5) unsigned NOT NULL default '0',
  PRIMARY KEY  (IdTrabajador),
  KEY Trab_NumTarjeta (NumTarjeta),
  KEY Trab_idCategoria (idCategoria),
  KEY Trab_idAsesoria (idAsesoria),
  KEY F_Seccion (Seccion),
  CONSTRAINT trabajadores_ibfk_1 FOREIGN KEY (idCategoria) REFERENCES categorias (IdCategoria),
  CONSTRAINT trabajadores_ibfk_4 FOREIGN KEY (Seccion) REFERENCES secciones (IdSeccion)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

