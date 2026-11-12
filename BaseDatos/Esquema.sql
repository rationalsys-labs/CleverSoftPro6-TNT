-- Esquema final Clever SoftPro 6.0. SQL Server 2008 (compatibilidad 100).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
CREATE TABLE dbo.[CargasInformacion](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [IdUsuario] int NOT NULL,
 [Archivo] nvarchar(255) NOT NULL,
 [Filas] int NOT NULL);
GO
CREATE TABLE dbo.[CartillasRP](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Numero] numeric(18,0) NOT NULL,
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [IdUsuario] int NOT NULL);
GO
CREATE TABLE dbo.[ClavesCobranzas](
 [Clave] uniqueidentifier NOT NULL,
 [Modulo] tinyint NOT NULL,
 [Accion] nvarchar(20) NOT NULL,
 [IdUsuario] int NOT NULL,
 [Solicitud] xml NOT NULL,
 [IdCaso] int NULL,
 [IdGestion] int NULL,
 [IdCheque] int NULL,
 [IdRecojo] int NULL,
 [IdCarta] int NULL,
 [IdDevolucion] int NULL);
GO
CREATE TABLE dbo.[ConfiguracionCobranzas](
 [Item] tinyint NOT NULL,
 [Empresa] nvarchar(150) NOT NULL,
 [CuentaSoles] nvarchar(150) NOT NULL,
 [CuentaDolares] nvarchar(150) NOT NULL,
 [Contacto] nvarchar(150) NOT NULL,
 [CorreoConfirmacion] nvarchar(300) NOT NULL,
 [Anexo] nvarchar(50) NOT NULL,
 [DestinoCheques] nvarchar(1000) NOT NULL,
 [DestinoFacturas] nvarchar(1000) NOT NULL,
 [DestinoCartas] nvarchar(1000) NOT NULL,
 [HostSMTP] nvarchar(255) NOT NULL,
 [PuertoSMTP] int NOT NULL,
 [UsarTLS] bit NOT NULL,
 [UsuarioSMTP] nvarchar(255) NOT NULL,
 [RemitenteSMTP] nvarchar(255) NOT NULL,
 [ClaveSMTP] varbinary(max) NULL,
 [EquipoSMTP] nvarchar(128) NOT NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[CorreosTerritorioCobranza](
 [IdTerritorio] int NOT NULL,
 [Destinatarios] nvarchar(1000) NOT NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[DbCheques](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Asient] nvarchar(10) NOT NULL,
 [Client] nvarchar(300) NOT NULL,
 [CodCli] nvarchar(20) NOT NULL,
 [Direcc] nvarchar(255) NOT NULL,
 [Contac] nvarchar(50) NOT NULL,
 [Cargo] nvarchar(50) NOT NULL,
 [Banco] int NULL,
 [FecGir] datetime NOT NULL,
 [NroChe] nvarchar(50) NULL,
 [Monto] money NOT NULL,
 [Divisa] nvarchar(50) NOT NULL,
 [Motivo] int NULL,
 [Gastos] money NULL,
 [MonCan] money NULL,
 [Estado] nvarchar(20) NULL,
 [SitDoc] nvarchar(5) NOT NULL,
 [FecCar] datetime NULL,
 [NroRUC] nvarchar(20) NULL,
 [IdCliente] int NULL,
 [VersionFila] timestamp NOT NULL,
 [IdCreador] int NULL,
 [EstadoGestion] nvarchar(10) NOT NULL DEFAULT ('Abierto'),
 [FechaCierre] datetime NULL,
 [IdCerrador] int NULL,
 [MotivoCierre] nvarchar(1000) NULL,
 [Solicitud] xml NULL);
GO
CREATE TABLE dbo.[DbClientes](
 [Asient] nvarchar(20) NOT NULL,
 [CLIENT] nvarchar(255) NOT NULL,
 [RUCDNI] nvarchar(20) NOT NULL,
 [DIRECC] nvarchar(255) NOT NULL,
 [TELEF1] nvarchar(50) NULL,
 [TELEF2] nvarchar(50) NULL,
 [TELEF3] nvarchar(50) NULL,
 [IdTerritorio] int NOT NULL,
 [CLITNT] nvarchar(10) NULL,
 [CLITOP] nvarchar(10) NULL,
 [REMARK] nvarchar(255) NULL,
 [CLIIMP] nvarchar(10) NULL,
 [CBZATNT] nvarchar(5) NULL,
 [DATAPAS] nvarchar(50) NULL,
 [FECDATA] datetime NULL,
 [Contac] nvarchar(100) NULL,
 [ContacEnt] nvarchar(100) NULL,
 [DireccEnt] nvarchar(255) NULL,
 [TelefoEnt] nvarchar(50) NULL,
 [Recarg] money NULL,
 [Item] int IDENTITY(1,1) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [VersionFila] timestamp NOT NULL,
 [EsInternacional] bit NOT NULL DEFAULT ((0)),
 [IdPais] int NULL,
 [IdDistrito] int NULL,
 [IdDistritoEntrega] int NOT NULL,
 [IdNivel] tinyint NOT NULL,
 [IdCobranza] tinyint NOT NULL,
 [IdComprobante] tinyint NOT NULL,
 [IdMoneda] tinyint NOT NULL,
 [IdTipoPersona] tinyint NOT NULL,
 [IdTipoIdentificacion] tinyint NOT NULL,
 [EsLocal] bit NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[DbClientes2](
 [Asient] nvarchar(20) NOT NULL,
 [CLIENT] nvarchar(255) NOT NULL,
 [RUCDNI] nvarchar(20) NOT NULL,
 [DIRECC] nvarchar(255) NOT NULL,
 [DISTRI] nvarchar(50) NULL,
 [TELEF1] nvarchar(50) NULL,
 [TELEF2] nvarchar(50) NULL,
 [TELEF3] nvarchar(50) NULL,
 [TERRIT] nvarchar(50) NULL,
 [CLITNT] nvarchar(10) NULL,
 [CLITOP] nvarchar(10) NULL,
 [REMARK] nvarchar(255) NULL,
 [CLIIMP] nvarchar(10) NULL,
 [CBZATNT] nvarchar(5) NULL,
 [DATAPAS] tinyint NULL,
 [FECDATA] datetime NULL,
 [Contac] nvarchar(100) NULL,
 [ContacEnt] nvarchar(100) NULL,
 [DireccEnt] nvarchar(255) NULL,
 [DistriEnt] nvarchar(50) NULL,
 [TelefoEnt] nvarchar(50) NULL,
 [Recarg] money NULL,
 [TipCom] nvarchar(10) NULL,
 [Divisa] nvarchar(5) NULL,
 [TipCli] nvarchar(10) NULL,
 [Cobran] nvarchar(10) NULL,
 [Item] int IDENTITY(1,1) NOT NULL);
GO
CREATE TABLE dbo.[DbDevol](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NroCar] nvarchar(10) NOT NULL,
 [FecCar] datetime NOT NULL,
 [Usuari] nvarchar(50) NOT NULL,
 [ImpExp] nvarchar(15) NOT NULL,
 [Client] nvarchar(255) NOT NULL,
 [CodCli] nvarchar(20) NOT NULL,
 [NroGui] nvarchar(100) NOT NULL,
 [Cancel] nvarchar(100) NOT NULL,
 [Abonos] nvarchar(100) NOT NULL,
 [Motivo] nvarchar(100) NOT NULL,
 [Transf] nvarchar(50) NOT NULL,
 [Otros] nvarchar(50) NOT NULL,
 [ParDep] nvarchar(50) NOT NULL,
 [CtaCte] nvarchar(50) NOT NULL,
 [Monto] money NULL,
 [Moneda] nvarchar(50) NULL,
 [Contac] nvarchar(50) NULL,
 [Estado] tinyint NOT NULL,
 [Final] tinyint NULL,
 [IdCliente] int NULL,
 [IdUsuario] int NULL,
 [VersionFila] timestamp NOT NULL,
 [IdCreador] int NULL,
 [EstadoGestion] nvarchar(10) NOT NULL DEFAULT ('Abierto'),
 [FechaCierre] datetime NULL,
 [IdCerrador] int NULL,
 [MotivoCierre] nvarchar(1000) NULL,
 [Solicitud] xml NULL);
GO
CREATE TABLE dbo.[DBEmiFac](
 [TipDoc] tinyint NULL,
 [NumDoc] nvarchar(14) NOT NULL,
 [FecEmi] datetime NOT NULL,
 [FecVen] datetime NULL,
 [CodCli] nvarchar(10) NULL,
 [NomCli] nvarchar(255) NOT NULL,
 [DivFac] nvarchar(50) NOT NULL,
 [SubFac] money NULL,
 [RecFac] money NULL,
 [IgvFac] money NULL,
 [ImpFac] money NOT NULL,
 [ImpSal] money NULL,
 [SitDoc] nvarchar(15) NOT NULL,
 [Estado] tinyint NULL,
 [FecEst] datetime NULL,
 [NumGui] nvarchar(10) NULL,
 [Courier] nvarchar(100) NULL,
 [Querie] nvarchar(20) NULL,
 [Raiz] nvarchar(14) NULL,
 [Recojo] nvarchar(14) NULL,
 [CarDev] nvarchar(10) NULL,
 [PorsRC] money NULL,
 [GasOtr] money NULL,
 [Servic] nvarchar(30) NULL,
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdCliente] int NULL,
 [IdEstado] int NULL,
 [IdDocumentoReferencia] int NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[DbGuias](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NroGui] nvarchar(10) NOT NULL,
 [Origen] nvarchar(5) NOT NULL,
 [PickUP] datetime NOT NULL,
 [Produc] nvarchar(5) NOT NULL,
 [Client] nvarchar(255) NOT NULL,
 [Tipo] nvarchar(10) NOT NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[DbHisCourier](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [TipDoc] nvarchar(5) NOT NULL,
 [NumDoc] nvarchar(20) NOT NULL,
 [Courier] nvarchar(100) NOT NULL,
 [Fecha] datetime NULL,
 [DATAPAS] tinyint NULL,
 [FECDATA] datetime NULL,
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL);
GO
CREATE TABLE dbo.[DbHisFac](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [TipDoc] tinyint NULL,
 [NumDoc] nvarchar(20) NOT NULL,
 [Estado] nvarchar(50) NOT NULL,
 [FecEst] datetime NOT NULL,
 [NomUsu] nvarchar(50) NULL,
 [DATAPAS] tinyint NULL,
 [FECDATA] datetime NULL,
 [IdUsuario] int NULL,
 [IdEstado] int NULL,
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL,
 [IdRegistroCargo] int NULL,
 [Activo] bit NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[DbNCredito](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [NumDoc] nvarchar(50) NOT NULL,
 [FecEmi] datetime NOT NULL,
 [CodCli] nvarchar(50) NOT NULL,
 [DivMon] nvarchar(50) NOT NULL,
 [Monto] money NULL,
 [NumRef] nvarchar(50) NULL,
 [Aplicad] nvarchar(10) NULL,
 [FacApli] nvarchar(15) NULL,
 [Estado] tinyint NULL,
 [FecEst] datetime NULL,
 [DATAPAS] tinyint NULL,
 [FECDATA] datetime NULL,
 [NomCli] nvarchar(300) NULL,
 [Tarifa] money NULL,
 [SitDoc] nvarchar(3) NULL,
 [IdCliente] int NULL,
 [Recargo] money NULL,
 [Igv] money NULL,
 [FecVen] datetime NULL,
 [IdDocumentoReferencia] int NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[DbQuejas_Cas](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Asient] nvarchar(10) NOT NULL,
 [CodCli] nvarchar(20) NOT NULL,
 [TipCas] tinyint NULL,
 [FecCas] datetime NOT NULL,
 [FecFin] datetime NULL,
 [Estado] nvarchar(20) NULL,
 [Usuari] nvarchar(20) NULL,
 [UsuFin] nvarchar(20) NULL,
 [DataPas] tinyint NOT NULL,
 [FecData] datetime NOT NULL,
 [IdCliente] int NULL,
 [IdUsuario] int NULL,
 [VersionFila] timestamp NOT NULL,
 [DescripcionCaso] nvarchar(150) NULL);
GO
CREATE TABLE dbo.[DbQuejas_Cli](
 [CodCli] nvarchar(20) NOT NULL,
 [Estado] nvarchar(20) NOT NULL,
 [CliTop] tinyint NOT NULL,
 [Usuario] nvarchar(20) NOT NULL,
 [FecApe] datetime NOT NULL,
 [ContCas] numeric(18,0) NOT NULL,
 [DataPas] tinyint NOT NULL,
 [FecData] datetime NOT NULL,
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdCliente] int NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[DbQuejas_Ges](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Asient] nvarchar(10) NOT NULL,
 [Raiz] nvarchar(10) NOT NULL,
 [Status] tinyint NULL,
 [FecSta] datetime NOT NULL,
 [Contac] nvarchar(50) NULL,
 [Telef1] nvarchar(20) NULL,
 [Gestion] ntext NOT NULL,
 [Usuari] nvarchar(50) NOT NULL,
 [DataPas] tinyint NOT NULL,
 [FecData] datetime NOT NULL,
 [IdUsuario] int NULL,
 [IdCaso] int NULL,
 [VersionFila] timestamp NOT NULL,
 [CodigoEstado] nvarchar(20) NULL,
 [DescripcionEstado] nvarchar(150) NULL);
GO
CREATE TABLE dbo.[DbQuerie](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Asient] nvarchar(10) NOT NULL,
 [FecSol] datetime NOT NULL,
 [CliDoc] nvarchar(255) NOT NULL,
 [TipDoc] tinyint NULL,
 [NumDoc] nvarchar(14) NOT NULL,
 [CodCli] nvarchar(50) NULL,
 [FecEmi] datetime NOT NULL,
 [MonSol] money NOT NULL,
 [DivSol] nvarchar(50) NOT NULL,
 [RcSol] money NULL,
 [SolTer] nvarchar(150) NULL,
 [SolDa1] tinyint NULL,
 [SolDa2] tinyint NULL,
 [SolDa3] tinyint NULL,
 [SoAnl1] tinyint NULL,
 [SoAnl2] tinyint NULL,
 [Solib1] tinyint NULL,
 [Solib2] tinyint NULL,
 [SolOtr] nvarchar(20) NULL,
 [SolCli] nvarchar(255) NULL,
 [SolRuc] nvarchar(11) NULL,
 [SolDir] nvarchar(255) NULL,
 [SOlTel] nvarchar(20) NULL,
 [SolTa1] tinyint NULL,
 [SolTa2] tinyint NULL,
 [SolTa3] tinyint NULL,
 [SolTa4] tinyint NULL,
 [SolTa5] tinyint NULL,
 [SolMa1] tinyint NULL,
 [SolMa2] tinyint NULL,
 [SolMa3] tinyint NULL,
 [SolMa4] tinyint NULL,
 [SolMa5] tinyint NULL,
 [SoCon1] tinyint NULL,
 [SoCon2] tinyint NULL,
 [SoCon3] tinyint NULL,
 [SoCli2] nvarchar(255) NULL,
 [SoRuc2] nvarchar(11) NULL,
 [SoDir2] nvarchar(255) NULL,
 [SoTel2] nvarchar(20) NULL,
 [SoOtr1] tinyint NULL,
 [SolEsp] nvarchar(50) NULL,
 [SoGui1] nvarchar(50) NULL,
 [SoGui2] nvarchar(50) NULL,
 [SoGui3] nvarchar(50) NULL,
 [SoGui4] nvarchar(50) NULL,
 [SoGui5] nvarchar(50) NULL,
 [SoGui6] nvarchar(50) NULL,
 [SolAdi] nvarchar(255) NULL,
 [SolDes] money NULL,
 [SolRes] nvarchar(150) NULL,
 [SolEnc] nvarchar(50) NULL,
 [SitDoc] nvarchar(15) NULL,
 [TipQue] nvarchar(20) NULL,
 [Estado] nvarchar(50) NULL,
 [DATAPAS] tinyint NULL,
 [FECDATA] datetime NULL,
 [ORIGEN] nvarchar(50) NULL,
 [TipCas] nvarchar(100) NULL,
 [Proble] nvarchar(1000) NULL,
 [FecIni] datetime NULL,
 [FecVis] datetime NULL,
 [IdCliente] int NULL,
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL,
 [VersionFila] timestamp NOT NULL,
 [IdCreador] int NULL,
 [IdResponsable] int NULL,
 [IdServicio] int NULL,
 [ServicioSolicitud] nvarchar(150) NULL,
 [OpcionesSolicitud] xml NULL,
 [SolicitudOriginal] xml NULL,
 [ClaveSolicitud] uniqueidentifier NULL,
 [IdNotaResolucion] int NULL,
 [FechaCierre] datetime NULL,
 [IdCerrador] int NULL,
 [Resolucion] nvarchar(1000) NULL,
 [MontoAplicado] money NULL,
 [FechaAnulacion] datetime NULL,
 [IdAnulador] int NULL);
GO
CREATE TABLE dbo.[DbRecojos](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Asient] nvarchar(10) NULL,
 [FecEnv] datetime NULL,
 [ModGes] nvarchar(10) NULL,
 [CodCli] nvarchar(20) NULL,
 [Mensaj] ntext NULL,
 [Estado] tinyint NULL,
 [Destino] nvarchar(1000) NULL,
 [IdCliente] int NULL,
 [VersionFila] timestamp NOT NULL,
 [IdCreador] int NULL,
 [EstadoGestion] nvarchar(10) NOT NULL DEFAULT ('Abierto'),
 [FechaCierre] datetime NULL,
 [IdCerrador] int NULL,
 [MotivoCierre] nvarchar(1000) NULL,
 [Solicitud] xml NULL);
GO
CREATE TABLE dbo.[DbsCartas](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Correl] nvarchar(50) NOT NULL,
 [FecCar] datetime NOT NULL,
 [CliCar] nvarchar(255) NOT NULL,
 [CodCli] nvarchar(20) NOT NULL,
 [DirCar] nvarchar(255) NOT NULL,
 [ContCar] nvarchar(100) NULL,
 [CargCar] nvarchar(50) NULL,
 [RefeCar] nvarchar(50) NULL,
 [DivCar] nvarchar(50) NULL,
 [MontCar] money NULL,
 [SgdaCar] tinyint NULL,
 [IdCliente] int NULL,
 [VersionFila] timestamp NOT NULL,
 [IdCreador] int NULL,
 [EstadoGestion] nvarchar(10) NOT NULL DEFAULT ('Abierto'),
 [FechaCierre] datetime NULL,
 [IdCerrador] int NULL,
 [MotivoCierre] nvarchar(1000) NULL,
 [Solicitud] xml NULL);
GO
CREATE TABLE dbo.[DbsCompRp](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [ASIENT] nvarchar(20) NOT NULL,
 [CORREL] nvarchar(20) NULL,
 [NROPLA] nvarchar(20) NULL,
 [FECPLA] datetime NULL,
 [NROGUI] nvarchar(15) NULL,
 [CLIENT] nvarchar(255) NULL,
 [CODCLI] nvarchar(20) NULL,
 [TERCLI] nvarchar(5) NULL,
 [PESFIS] decimal(18,4) NULL,
 [PESVOL] decimal(18,4) NULL,
 [MONNET] money NULL,
 [MONRC] money NULL,
 [MONIGV] money NULL,
 [MONTOT] money NULL,
 [MONEDA] tinyint NULL,
 [PUGUI] datetime NULL,
 [ORIGEN] nvarchar(50) NULL,
 [TIPGUI] nvarchar(10) NULL,
 [POLIZA] nvarchar(20) NULL,
 [NRODOC] nvarchar(15) NULL,
 [NRODUA] nvarchar(50) NULL,
 [VALCIF] money NULL,
 [FECNUM] datetime NULL,
 [CANAL] nvarchar(20) NULL,
 [FINAL] nvarchar(50) NULL,
 [INFO] ntext NULL,
 [ServSP] tinyint NULL,
 [ServSS] tinyint NULL,
 [ServSO] tinyint NULL,
 [IdCliente] int NULL,
 [IdGuia] int NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[DbsDiarRp](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Asient] nvarchar(20) NOT NULL,
 [NroGui] nvarchar(10) NOT NULL,
 [Client] nvarchar(255) NOT NULL,
 [CodCli] nvarchar(15) NULL,
 [PesFis] decimal(18,4) NULL,
 [PesVol] decimal(18,4) NULL,
 [MonNet] money NULL,
 [MonRc] money NULL,
 [RcPor] money NULL,
 [MonIgv] money NULL,
 [MonTot] money NULL,
 [Moneda] tinyint NULL,
 [PuGui] datetime NULL,
 [Origen] nvarchar(50) NULL,
 [TipGui] nvarchar(10) NULL,
 [Estado] tinyint NULL,
 [Salida] tinyint NULL,
 [Info] ntext NULL,
 [Reten] tinyint NULL,
 [ServSP] tinyint NULL,
 [ServSS] tinyint NULL,
 [ServSO] tinyint NULL,
 [TipRet] tinyint NULL,
 [IdCliente] int NULL,
 [IdGuia] int NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[DbServGui](
 [TipDoc] tinyint NULL,
 [NumDoc] nvarchar(15) NOT NULL,
 [NroGui] nvarchar(10) NOT NULL,
 [Servic] nvarchar(50) NOT NULL,
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL,
 [IdGuia] int NULL,
 [Linea] int NULL);
GO
CREATE TABLE dbo.[DbsHisRP](
 [Item] int IDENTITY(1,1) NOT NULL,
 [HERENCIA] nvarchar(20) NULL,
 [STATUS] nvarchar(5) NULL,
 [INFO] ntext NULL,
 [FECHA] datetime NULL,
 [IdDiario] int NULL,
 [IdCompilado] int NULL);
GO
CREATE TABLE dbo.[Departamentos](
 [Item] int NOT NULL,
 [Nombre] nvarchar(150) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [Codigo] char(2) NOT NULL,
 [IdPais] int NOT NULL);
GO
CREATE TABLE dbo.[DetallesCargos](
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdRegistro] int NOT NULL,
 [Tipo] nvarchar(2) NOT NULL,
 [IdDocumento] int NOT NULL,
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL,
 [Documento] nvarchar(50) NOT NULL,
 [Emision] datetime NOT NULL,
 [Cliente] nvarchar(300) NOT NULL,
 [Divisa] nvarchar(50) NOT NULL,
 [Importe] money NOT NULL,
 [Situacion] nvarchar(15) NULL,
 [Guias] nvarchar(max) NOT NULL,
 [Servicios] nvarchar(max) NOT NULL,
 [ContactoEntrega] nvarchar(255) NOT NULL,
 [DireccionEntrega] nvarchar(500) NOT NULL,
 [DistritoEntrega] nvarchar(100) NOT NULL,
 [TelefonoEntrega] nvarchar(255) NOT NULL);
GO
CREATE TABLE dbo.[DetallesCartillaRP](
 [Planilla] numeric(18,0) NOT NULL,
 [IdGuia] int NOT NULL,
 [NroGui] nvarchar(10) NOT NULL,
 [Cliente] nvarchar(255) NOT NULL,
 [Territorio] nvarchar(150) NULL,
 [PesFis] decimal(18,4) NULL,
 [PesVol] decimal(18,4) NULL,
 [Neto] money NULL,
 [Recargo] money NULL,
 [IGV] money NULL,
 [Total] money NULL,
 [Moneda] nvarchar(50) NULL,
 [FechaEnvio] datetime NULL,
 [Producto] nvarchar(10) NULL);
GO
CREATE TABLE dbo.[Distritos](
 [Item] int NOT NULL,
 [Nombre] nvarchar(150) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [Codigo] char(6) NOT NULL,
 [IdProvincia] int NOT NULL);
GO
CREATE TABLE dbo.[DocumentosCobranzas](
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdRecojo] int NULL,
 [IdCarta] int NULL,
 [IdDevolucion] int NULL,
 [IdFactura] int NOT NULL,
 [Tipo] nvarchar(2) NOT NULL,
 [Documento] nvarchar(14) NOT NULL,
 [Emision] datetime NOT NULL,
 [Cliente] nvarchar(255) NOT NULL,
 [Divisa] nvarchar(50) NOT NULL,
 [Importe] money NOT NULL);
GO
CREATE TABLE dbo.[DocumentosResolucionQuery](
 [IdQuery] int NOT NULL,
 [IdDocumento] int NOT NULL,
 [NumeroDocumento] nvarchar(14) NOT NULL,
 [Cliente] nvarchar(255) NOT NULL,
 [Emision] datetime NOT NULL,
 [Divisa] nvarchar(50) NOT NULL,
 [Importe] money NOT NULL);
GO
CREATE TABLE dbo.[EmisionesCarta](
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdCarta] int NOT NULL,
 [Numero] tinyint NOT NULL,
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [IdUsuario] int NOT NULL,
 [Modelo] tinyint NOT NULL,
 [Datos] xml NOT NULL);
GO
CREATE TABLE dbo.[EnviosCobranzas](
 [Item] uniqueidentifier NOT NULL,
 [IdRecojo] int NULL,
 [IdEmision] int NULL,
 [IdUsuario] int NOT NULL,
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [Destinatario] nvarchar(1000) NOT NULL,
 [Asunto] nvarchar(255) NOT NULL,
 [Cuerpo] nvarchar(max) NOT NULL,
 [Estado] nvarchar(12) NOT NULL,
 [Inicio] datetime NULL,
 [Fin] datetime NULL,
 [Detalle] nvarchar(1000) NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[EstadosInicialesCargos](
 [Tipo] nvarchar(2) NOT NULL,
 [IdDocumento] int NOT NULL,
 [IdEstado] int NULL,
 [Estado] tinyint NULL,
 [FechaEstado] datetime NULL);
GO
CREATE TABLE dbo.[EventosRegistrosCargos](
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdRegistro] int NOT NULL,
 [IdUsuario] int NOT NULL,
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [Activo] bit NOT NULL);
GO
CREATE TABLE dbo.[GuiasDevolucion](
 [IdDevolucion] int NOT NULL,
 [IdGuia] int NOT NULL,
 [Numero] nvarchar(10) NOT NULL);
GO
CREATE TABLE dbo.[GuiasQuery](
 [IdQuery] int NOT NULL,
 [IdGuia] int NOT NULL,
 [NumeroGuia] nvarchar(10) NOT NULL);
GO
CREATE TABLE dbo.[HistorialCobranzas](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [IdUsuario] int NOT NULL,
 [Modulo] tinyint NOT NULL,
 [Accion] nvarchar(20) NOT NULL,
 [Detalle] nvarchar(1000) NOT NULL,
 [IdCaso] int NULL,
 [IdCheque] int NULL,
 [IdRecojo] int NULL,
 [IdCarta] int NULL,
 [IdDevolucion] int NULL);
GO
CREATE TABLE dbo.[HistorialQueries](
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdQuery] int NOT NULL,
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [IdUsuario] int NOT NULL,
 [Accion] nvarchar(20) NOT NULL,
 [Detalle] nvarchar(1000) NOT NULL);
GO
CREATE TABLE dbo.[INDICES](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [SOLICITUD] nvarchar(20) NOT NULL,
 [CONTEO] nvarchar(20) NOT NULL);
GO
CREATE TABLE dbo.[Int_Access](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Usuari] nvarchar(10) NOT NULL,
 [CodMod] nvarchar(10) NOT NULL,
 [Access] tinyint NOT NULL,
 [IdUsuario] int NULL);
GO
CREATE TABLE dbo.[Int_Bancos](
 [CodBank] int IDENTITY(1,1) NOT NULL,
 [NomBank] nvarchar(50) NOT NULL,
 [ActEst] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[Int_Casos](
 [CodCas] tinyint NOT NULL,
 [NomCas] nvarchar(50) NOT NULL,
 [ActEst] bit NOT NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[Int_ComentaryRP](
 [CodCli] nvarchar(20) NOT NULL,
 [Comentary] ntext NULL,
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdCliente] int NULL);
GO
CREATE TABLE dbo.[Int_Correl](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Usuari] nvarchar(5) NOT NULL,
 [DivCor] nvarchar(3) NOT NULL,
 [Correl] numeric(18,0) NOT NULL,
 [Concep] nvarchar(50) NULL,
 [ActDiv] tinyint NULL,
 [instance] nvarchar(10) NULL);
GO
CREATE TABLE dbo.[Int_DivRp](
 [NomCli] nvarchar(255) NOT NULL,
 [CodCli] nvarchar(20) NOT NULL,
 [DesCli] nvarchar(255) NULL,
 [Item] int IDENTITY(1,1) NOT NULL,
 [IdCliente] int NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[Int_EstadCheq](
 [CodEst] int IDENTITY(1,1) NOT NULL,
 [Estado] nvarchar(50) NOT NULL,
 [ActEst] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[int_IATA](
 [Origen] nvarchar(50) NULL,
 [codigo] nvarchar(10) NOT NULL);
GO
CREATE TABLE dbo.[Int_InfoMail](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Usuari] nvarchar(10) NULL,
 [LogUsu] nvarchar(50) NULL,
 [PasUsu] nvarchar(50) NULL,
 [MailChe] nvarchar(350) NULL,
 [MailFac] nvarchar(350) NULL,
 [IdUsuario] int NULL);
GO
CREATE TABLE dbo.[Int_InfoSys](
 [Item] int IDENTITY(1,1) NOT NULL,
 [ModAct] nvarchar(50) NULL,
 [Infor1] nvarchar(100) NULL,
 [Infor2] nvarchar(100) NULL,
 [Infor3] nvarchar(100) NULL);
GO
CREATE TABLE dbo.[Int_MailDatQue](
 [Item] int IDENTITY(1,1) NOT NULL,
 [VarMail] nvarchar(20) NOT NULL,
 [TxMail] nvarchar(250) NULL,
 [Concep] nvarchar(50) NOT NULL);
GO
CREATE TABLE dbo.[Int_ModulosUsuario](
 [Item] int NOT NULL,
 [Nombre] nvarchar(100) NOT NULL,
 [IdPadre] int NULL,
 [CodMod] char(6) NOT NULL,
 [Vigente] bit NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[Int_Moneda](
 [CodEst] tinyint NOT NULL,
 [Moneda] nvarchar(50) NOT NULL,
 [ActEst] bit NOT NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[Int_PreRP](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NroGui] nvarchar(50) NULL,
 [SolRP] nvarchar(255) NULL,
 [Cliente] nvarchar(300) NULL,
 [CodCli] nvarchar(50) NULL,
 [FecSol] datetime NULL,
 [Estado] numeric(18,0) NULL,
 [Usuario] nvarchar(10) NULL,
 [IdCliente] int NULL,
 [IdUsuario] int NULL);
GO
CREATE TABLE dbo.[Int_PublicMail](
 [Item] int IDENTITY(1,1) NOT NULL,
 [ModMai] nvarchar(10) NULL,
 [MailCo] nvarchar(350) NULL);
GO
CREATE TABLE dbo.[Int_StatusCas](
 [CodSta] tinyint NOT NULL,
 [Status] nvarchar(3) NOT NULL,
 [DesSta] nvarchar(50) NOT NULL,
 [ActEst] bit NOT NULL,
 [IteAcc] nvarchar(10) NULL,
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[Int_StatusFac](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [TitEst] nvarchar(50) NOT NULL,
 [Status] nvarchar(20) NOT NULL,
 [RGB1] numeric(18,0) NOT NULL,
 [RGB2] numeric(18,0) NOT NULL,
 [RGB3] numeric(18,0) NOT NULL,
 [ActEst] bit NOT NULL,
 [IteAcc] nvarchar(10) NULL,
 [Correl] numeric(18,0) NULL,
 [ActCor] tinyint NULL);
GO
CREATE TABLE dbo.[Int_StatusRP](
 [CodEst] tinyint NOT NULL,
 [Status] nvarchar(5) NOT NULL,
 [TitSta] nvarchar(50) NOT NULL,
 [ActEst] bit NOT NULL);
GO
CREATE TABLE dbo.[Int_TipDoc](
 [CodIte] tinyint NOT NULL,
 [TipDoc] nvarchar(2) NOT NULL,
 [NamVar] nvarchar(50) NOT NULL,
 [VersionFila] timestamp NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[Int_Usuario](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [NomUsu] nvarchar(50) NOT NULL,
 [ApeUsu] nvarchar(50) NOT NULL,
 [Usuari] nvarchar(50) NOT NULL,
 [Paswor] ntext NOT NULL,
 [ActPas] bit NOT NULL,
 [Access] nvarchar(50) NULL,
 [SitUsu] nvarchar(8) NOT NULL,
 [Email] nvarchar(100) NULL,
 [Origen] nvarchar(50) NULL,
 [VersionFila] timestamp NOT NULL,
 [IdPerfilAcceso] int NULL);
GO
CREATE TABLE dbo.[IntAccesPlant](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Planti] int NULL,
 [CodMod] numeric(18,0) NOT NULL,
 [Access] tinyint NOT NULL,
 [IdMenu] int NULL,
 [IdEstado] int NULL);
GO
CREATE TABLE dbo.[IntAccess](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Usuari] int NULL,
 [CodMod] numeric(18,0) NOT NULL,
 [Access] tinyint NOT NULL,
 [IdMenu] int NULL,
 [IdEstado] int NULL);
GO
CREATE TABLE dbo.[IntBackup](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Archivo] nvarchar(50) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntCargosEnterprice](
 [CodCar] int IDENTITY(1,1) NOT NULL,
 [DesCar] nvarchar(150) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((0)));
GO
CREATE TABLE dbo.[IntComentaryRP](
 [Item] int IDENTITY(1,1) NOT NULL,
 [CodCli] numeric(32,0) NULL,
 [Coment] nvarchar(500) NOT NULL);
GO
CREATE TABLE dbo.[IntDominioRP](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Dominio] nvarchar(50) NOT NULL,
 [Codigo] numeric(18,0) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntInfoBody](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Body] nvarchar(50) NOT NULL,
 [Asunto] nvarchar(250) NOT NULL,
 [Mensaj] nvarchar(2000) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntInfoMail](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Grupos] nvarchar(50) NOT NULL,
 [Destin] nvarchar(500) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntInforQuer](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Domin] tinyint NOT NULL,
 [Descrip] nvarchar(150) NOT NULL);
GO
CREATE TABLE dbo.[IntInfoSys](
 [Item] int IDENTITY(1,1) NOT NULL,
 [ModAct] nvarchar(50) NOT NULL,
 [Infor1] nvarchar(100) NOT NULL);
GO
CREATE TABLE dbo.[IntInfoTerMail](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Territ] int NULL,
 [DirTer] nvarchar(max) NOT NULL DEFAULT (''),
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntMailOrigen](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Origen] nvarchar(50) NOT NULL,
 [Modal] tinyint NOT NULL,
 [Grupos] int NULL DEFAULT (NULL),
 [Mensaj] int NULL DEFAULT (NULL),
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntMotivoDevol](
 [CodMot] int IDENTITY(1,1) NOT NULL,
 [DesMot] nvarchar(150) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((0)));
GO
CREATE TABLE dbo.[IntNiveles](
 [item] int IDENTITY(1,1) NOT NULL,
 [Descrip] nvarchar(150) NOT NULL DEFAULT (''),
 [Planti] int NULL);
GO
CREATE TABLE dbo.[IntPlantillas](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NomPla] nvarchar(150) NOT NULL);
GO
CREATE TABLE dbo.[IntPrealertRP](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NroGui] numeric(18,0) NOT NULL,
 [SolRP] nvarchar(255) NOT NULL,
 [Cliente] nvarchar(300) NOT NULL,
 [CodCli] numeric(32,0) NULL,
 [FecSol] datetime NOT NULL DEFAULT (getdate()),
 [Estado] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntProdGuias](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Product] nvarchar(10) NOT NULL,
 [Descrip] nvarchar(50) NOT NULL);
GO
CREATE TABLE dbo.[IntRefeModalDevol](
 [CodMod] int IDENTITY(1,1) NOT NULL,
 [DesMod] nvarchar(150) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((0)));
GO
CREATE TABLE dbo.[IntReferEnterprice](
 [CodRef] int IDENTITY(1,1) NOT NULL,
 [DesRef] nvarchar(150) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((0)));
GO
CREATE TABLE dbo.[IntReplaceRecojos](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Origen] tinyint NOT NULL DEFAULT ((0)),
 [VarData] nvarchar(50) NOT NULL,
 [Replace] int NULL DEFAULT (NULL));
GO
CREATE TABLE dbo.[IntRespQuerie](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Descrip] nvarchar(150) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntSaveActualizar](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Consulta] nvarchar(max) NOT NULL,
 [ActSav] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntServFac](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Descrip] nvarchar(150) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntStaRecojos](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Status] nvarchar(50) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntStatusCart](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Status] nvarchar(50) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[IntTerritorios](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Descrip] nvarchar(150) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)),
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[IntVarReplace](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Variable] nvarchar(20) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[MdbAsientCancel](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NumDoc] nvarchar(15) NOT NULL,
 [Monto] money NOT NULL,
 [Divisa] tinyint NULL,
 [TipCam] money NOT NULL DEFAULT ((0.00)),
 [Tipo] tinyint NOT NULL DEFAULT ((1)),
 [Active] tinyint NOT NULL DEFAULT ((1)),
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL);
GO
CREATE TABLE dbo.[MdbCartas](
 [Item] int IDENTITY(1,1) NOT NULL,
 [CodCar] numeric(18,0) NOT NULL,
 [FecEmi] datetime NOT NULL DEFAULT (getdate()),
 [CliCar] nvarchar(250) NOT NULL,
 [CodCli] numeric(32,0) NULL,
 [DirCar] nvarchar(250) NOT NULL,
 [Contac] nvarchar(100) NOT NULL,
 [Cargo] int NULL,
 [Refere] int NULL,
 [Divisa] tinyint NULL,
 [MontCar] money NOT NULL,
 [SgdaCar] tinyint NOT NULL DEFAULT ((0)),
 [Status] int NULL DEFAULT (NULL),
 [FecSta] datetime NOT NULL DEFAULT (getdate()),
 [Descrip] nvarchar(250) NOT NULL DEFAULT (''),
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[MdbCheques](
 [Item] int IDENTITY(1,1) NOT NULL,
 [CodChe] numeric(18,0) NOT NULL,
 [FecCar] datetime NOT NULL DEFAULT (getdate()),
 [Client] nvarchar(250) NOT NULL,
 [RucCli] nvarchar(50) NOT NULL,
 [DirCli] nvarchar(250) NOT NULL,
 [Contac] nvarchar(50) NOT NULL,
 [Cargo] int NULL,
 [CodCli] numeric(32,0) NULL,
 [Banco] int NULL,
 [FecGir] datetime NOT NULL,
 [NroChe] nvarchar(50) NOT NULL,
 [Motivo] int NULL,
 [Gastos] money NOT NULL,
 [MonCan] money NOT NULL,
 [Divisa] tinyint NULL DEFAULT ((1)),
 [Status] int NULL,
 [FecSta] datetime NOT NULL,
 [Descrip] nvarchar(250) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[MDbClientes](
 [Asient] numeric(32,0) NOT NULL,
 [Client] nvarchar(255) NOT NULL,
 [RucDni] nvarchar(20) NOT NULL,
 [Direcc] nvarchar(255) NOT NULL,
 [Distri] nvarchar(50) NOT NULL,
 [Telef1] nvarchar(50) NOT NULL DEFAULT (''),
 [Telef2] nvarchar(50) NOT NULL DEFAULT (''),
 [Telef3] nvarchar(50) NOT NULL DEFAULT (''),
 [Territ] int NULL,
 [CliTnt] tinyint NOT NULL DEFAULT ((0)),
 [CliTop] tinyint NOT NULL DEFAULT ((0)),
 [Remark] nvarchar(255) NOT NULL DEFAULT (''),
 [CliImp] tinyint NOT NULL DEFAULT ((0)),
 [Contac] nvarchar(100) NOT NULL DEFAULT (''),
 [ContacEnt] nvarchar(100) NOT NULL DEFAULT (''),
 [DireccEnt] nvarchar(255) NOT NULL DEFAULT (''),
 [DistriEnt] nvarchar(50) NOT NULL DEFAULT (''),
 [TelefoEnt] nvarchar(50) NOT NULL DEFAULT (''),
 [Recarg] money NOT NULL DEFAULT ((0)),
 [TipCom] tinyint NULL DEFAULT (NULL),
 [Divisa] tinyint NULL DEFAULT (NULL));
GO
CREATE TABLE dbo.[MdbCompRp](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NroPla] numeric(18,0) NOT NULL DEFAULT ((0)),
 [FecPla] datetime NOT NULL DEFAULT (getdate()),
 [NroGui] numeric(18,0) NOT NULL DEFAULT ((0)),
 [Client] nvarchar(300) NOT NULL DEFAULT (''),
 [CodCli] numeric(32,0) NULL DEFAULT (NULL),
 [PesFis] money NOT NULL DEFAULT ((0.00)),
 [PesVol] money NOT NULL DEFAULT ((0.00)),
 [Neto] money NOT NULL DEFAULT ((0.00)),
 [MonRC] money NOT NULL DEFAULT ((0.00)),
 [RcPor] money NOT NULL DEFAULT ((0.00)),
 [MonIGV] money NOT NULL DEFAULT ((0.00)),
 [Total] money NOT NULL DEFAULT ((0.00)),
 [Divisa] tinyint NULL,
 [PuGui] datetime NOT NULL,
 [Origen] nvarchar(50) NOT NULL,
 [Produc] int NULL,
 [Status] tinyint NULL DEFAULT (NULL),
 [Info] nvarchar(max) NOT NULL DEFAULT (''),
 [Servic] nvarchar(10) NOT NULL DEFAULT ('0,0,0'),
 [InterMen] nvarchar(255) NOT NULL DEFAULT (''),
 [IdGuia] int NULL);
GO
CREATE TABLE dbo.[MdbDevol](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NroCar] numeric(18,0) NOT NULL,
 [FecCar] datetime NOT NULL DEFAULT (getdate()),
 [Usuari] int NULL,
 [ImpExp] int NULL,
 [Client] nvarchar(250) NOT NULL,
 [CodCli] numeric(32,0) NULL,
 [NroGui] numeric(18,0) NOT NULL,
 [Cancel] nvarchar(150) NOT NULL,
 [Abonos] nvarchar(150) NOT NULL,
 [Motivo] int NULL,
 [Factura] nvarchar(150) NOT NULL DEFAULT (''),
 [Modali] int NOT NULL,
 [Banco] int NULL,
 [CtaCte] nvarchar(50) NOT NULL,
 [Monto] money NOT NULL,
 [Divisa] tinyint NULL,
 [Contac] nvarchar(50) NOT NULL,
 [Status] int NULL,
 [FecSta] datetime NOT NULL,
 [Descrip] nvarchar(250) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[MdbDiarRp](
 [Item] int NOT NULL,
 [NroGui] numeric(18,0) NOT NULL DEFAULT ((0)),
 [Client] nvarchar(300) NOT NULL DEFAULT (''),
 [CodCli] numeric(32,0) NULL DEFAULT (NULL),
 [PesFis] money NOT NULL DEFAULT ((0.00)),
 [PesVol] money NOT NULL DEFAULT ((0.00)),
 [Neto] money NOT NULL DEFAULT ((0.00)),
 [MonRC] money NOT NULL DEFAULT ((0.00)),
 [RcPor] money NOT NULL DEFAULT ((0.00)),
 [MonIGV] money NOT NULL DEFAULT ((0.00)),
 [Total] money NOT NULL DEFAULT ((0.00)),
 [Divisa] tinyint NULL,
 [PuGui] datetime NOT NULL,
 [Origen] nvarchar(50) NOT NULL,
 [Produc] int NULL,
 [Status] tinyint NULL DEFAULT (NULL),
 [Salida] tinyint NOT NULL DEFAULT ((0)),
 [Info] nvarchar(max) NOT NULL DEFAULT (''),
 [Reten] tinyint NOT NULL DEFAULT ((0)),
 [Servic] nvarchar(10) NOT NULL DEFAULT ('0,0,0'),
 [InterMen] nvarchar(255) NOT NULL DEFAULT (''),
 [IdGuia] int NULL);
GO
CREATE TABLE dbo.[MdbDocument](
 [TipDoc] tinyint NULL,
 [NumDoc] nvarchar(14) NOT NULL,
 [FecEmi] datetime NOT NULL DEFAULT (getdate()),
 [FecVen] datetime NOT NULL DEFAULT (getdate()),
 [CodCli] numeric(32,0) NULL,
 [NomCli] nvarchar(255) NOT NULL,
 [DivFac] tinyint NULL,
 [SubFac] money NOT NULL DEFAULT ((0.00)),
 [RecFac] money NOT NULL DEFAULT ((0.00)),
 [PorsRC] money NOT NULL DEFAULT ((0.00)),
 [GasOtr] money NOT NULL DEFAULT ((0.00)),
 [IgvFac] money NOT NULL DEFAULT ((0.00)),
 [ImpFac] money NOT NULL DEFAULT ((0.00)),
 [SitDoc] tinyint NOT NULL DEFAULT ((1)),
 [Item] int IDENTITY(1,1) NOT NULL);
GO
CREATE TABLE dbo.[MdbGuias](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Guia] numeric(18,0) NOT NULL,
 [Origen] nvarchar(10) NOT NULL,
 [Fecha] datetime NOT NULL,
 [Produc] int NULL,
 [Client] nvarchar(300) NOT NULL,
 [Tipo] nvarchar(10) NOT NULL);
GO
CREATE TABLE dbo.[MdbHisRP](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Herencia] int NOT NULL,
 [Status] tinyint NULL,
 [Info] nvarchar(1000) NOT NULL DEFAULT (''),
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [Active] tinyint NOT NULL DEFAULT ((1)),
 [IdDiario] int NULL,
 [IdCompilado] int NULL);
GO
CREATE TABLE dbo.[MDBHistory](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [TipDoc] tinyint NULL,
 [NumDoc] nvarchar(14) NOT NULL,
 [Status] int NULL,
 [FecEmi] datetime NOT NULL DEFAULT (getdate()),
 [Usuari] int NULL,
 [Correl] numeric(18,0) NOT NULL DEFAULT ((0)),
 [Active] tinyint NOT NULL DEFAULT ((1)),
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL);
GO
CREATE TABLE dbo.[MdbNCredito](
 [Numdoc] nvarchar(50) NOT NULL,
 [Emision] datetime NOT NULL,
 [CodCli] numeric(32,0) NULL,
 [Client] nvarchar(300) NOT NULL,
 [Divisa] tinyint NULL,
 [Total] money NOT NULL,
 [SitDoc] tinyint NOT NULL,
 [Refere] nvarchar(50) NOT NULL,
 [Item] int IDENTITY(1,1) NOT NULL);
GO
CREATE TABLE dbo.[MdbNotas](
 [NumDoc] nvarchar(50) NOT NULL,
 [FecEmi] datetime NOT NULL,
 [CodCli] numeric(32,0) NULL,
 [Client] nvarchar(300) NOT NULL,
 [Divisa] tinyint NULL,
 [Total] money NOT NULL,
 [SitDoc] tinyint NOT NULL,
 [Refere] nvarchar(50) NOT NULL,
 [Item] int IDENTITY(1,1) NOT NULL);
GO
CREATE TABLE dbo.[MdbPlantillaPri](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Destino] nvarchar(150) NOT NULL,
 [Codigo] numeric(18,0) NOT NULL);
GO
CREATE TABLE dbo.[MDbQuerie](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [Asient] nvarchar(15) NOT NULL,
 [FecSol] datetime NOT NULL,
 [CliDoc] nvarchar(255) NOT NULL,
 [TipDoc] tinyint NULL,
 [NumDoc] nvarchar(14) NOT NULL,
 [CodCli] numeric(32,0) NULL,
 [FecEmi] datetime NOT NULL,
 [MonSol] money NOT NULL,
 [DivSol] tinyint NULL,
 [RcSol] money NOT NULL,
 [SolTer] int NULL,
 [SolDa1] tinyint NOT NULL,
 [SolDa2] tinyint NOT NULL,
 [SolDa3] tinyint NOT NULL,
 [SoAnl1] tinyint NOT NULL,
 [SoAnl2] tinyint NOT NULL,
 [Solib1] tinyint NOT NULL,
 [Solib2] tinyint NOT NULL,
 [SolCli] nvarchar(255) NOT NULL,
 [SolRuc] nvarchar(15) NOT NULL,
 [SolDir] nvarchar(255) NOT NULL,
 [SOlTel] nvarchar(20) NOT NULL,
 [SolTa1] tinyint NOT NULL,
 [SolTa2] tinyint NOT NULL,
 [SolTa3] tinyint NOT NULL,
 [SolTa4] tinyint NOT NULL,
 [SolTa5] tinyint NOT NULL,
 [SolMa1] tinyint NOT NULL,
 [SolMa2] tinyint NOT NULL,
 [SolMa3] tinyint NOT NULL,
 [SolMa4] tinyint NOT NULL,
 [SolMa5] tinyint NOT NULL,
 [SoCon1] tinyint NOT NULL,
 [SoCon2] tinyint NOT NULL,
 [SoCon3] tinyint NOT NULL,
 [CheDes] tinyint NOT NULL DEFAULT ((0)),
 [SoCli2] nvarchar(255) NOT NULL,
 [SoRuc2] nvarchar(11) NOT NULL,
 [SoDir2] nvarchar(255) NOT NULL,
 [SoTel2] nvarchar(20) NOT NULL,
 [Servic] int NULL,
 [SoGuia] nvarchar(200) NOT NULL,
 [SolAdi] nvarchar(255) NOT NULL,
 [SolDes] money NOT NULL,
 [SolRes] int NULL,
 [SolEnc] int NOT NULL,
 [SitDoc] tinyint NOT NULL DEFAULT ((1)),
 [TipQue] tinyint NOT NULL,
 [Estado] tinyint NOT NULL DEFAULT ((1)),
 [ORIGEN] tinyint NOT NULL,
 [TipCas] int NULL,
 [Proble] int NULL,
 [FecIni] datetime NOT NULL,
 [FecVis] datetime NOT NULL,
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL);
GO
CREATE TABLE dbo.[MDbRecojos](
 [Asient] int NOT NULL,
 [Origen] nvarchar(50) NOT NULL,
 [FecEnv] datetime NOT NULL,
 [Mensaj] nvarchar(2000) NOT NULL,
 [Destino] nvarchar(300) NOT NULL,
 [Status] int NULL,
 [FecSta] datetime NOT NULL,
 [Descrip] nvarchar(250) NOT NULL,
 [Active] tinyint NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[MdbServGuias](
 [Item] int IDENTITY(1,1) NOT NULL,
 [NroDoc] nvarchar(20) NOT NULL,
 [TipoDoc] tinyint NULL,
 [Guia] numeric(18,0) NOT NULL,
 [Servic] int NULL,
 [IdFactura] int NULL,
 [IdNotaCredito] int NULL,
 [IdGuia] int NULL);
GO
CREATE TABLE dbo.[MdbUsers](
 [ITEM] int IDENTITY(1,1) NOT NULL,
 [NomUsu] nvarchar(50) NOT NULL,
 [ApeUsu] nvarchar(50) NOT NULL,
 [Usuari] nvarchar(5) NOT NULL,
 [Access] int NOT NULL DEFAULT ((0)),
 [Paswor] nvarchar(150) NOT NULL,
 [ActPas] tinyint NOT NULL DEFAULT ((1)),
 [SitUsu] tinyint NOT NULL DEFAULT ((1)),
 [Origen] tinyint NULL,
 [USUARIO] nvarchar(50) NULL,
 [PASWORD] nvarchar(50) NULL);
GO
CREATE TABLE dbo.[MigracionIncidencias](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Tabla] nvarchar(128) NOT NULL,
 [Clave] nvarchar(150) NULL,
 [Columna] nvarchar(128) NULL,
 [ValorOriginal] nvarchar(max) NULL,
 [Motivo] nvarchar(250) NOT NULL,
 [DatosOriginales] xml NULL);
GO
CREATE TABLE dbo.[NivelesCliente](
 [Item] tinyint NOT NULL,
 [Nombre] nvarchar(50) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[OtrosCasosQuery](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Descripcion] nvarchar(150) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[Paises](
 [Item] int NOT NULL,
 [Nombre] nvarchar(150) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [Codigo] char(2) NOT NULL);
GO
CREATE TABLE dbo.[PerfilesAcceso](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Nombre] nvarchar(50) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[PermisosPerfil](
 [IdPerfil] int NOT NULL,
 [IdModulo] int NOT NULL,
 [Acceso] bit NOT NULL);
GO
CREATE TABLE dbo.[Provincias](
 [Item] int NOT NULL,
 [Nombre] nvarchar(150) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [Codigo] char(4) NOT NULL,
 [IdDepartamento] int NOT NULL);
GO
CREATE TABLE dbo.[RegistrosCargos](
 [Item] int IDENTITY(1,1) NOT NULL,
 [Numero] int NOT NULL,
 [Fecha] datetime NOT NULL DEFAULT (getdate()),
 [FechaEstado] datetime NOT NULL,
 [IdEstado] int NOT NULL,
 [IdUsuario] int NOT NULL,
 [Titulo] nvarchar(100) NOT NULL,
 [Correlativo] int NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[SeriesCargos](
 [Serie] char(3) NOT NULL,
 [Numero] int NOT NULL);
GO
CREATE TABLE dbo.[SeriesClientes](
 [Serie] char(2) NOT NULL,
 [Numero] int NOT NULL);
GO
CREATE TABLE dbo.[SeriesCobranzas](
 [Codigo] varchar(3) NOT NULL,
 [Numero] int NOT NULL);
GO
CREATE TABLE dbo.[SeriesQueries](
 [Serie] char(1) NOT NULL,
 [Numero] int NOT NULL);
GO
CREATE TABLE dbo.[sysdiagrams](
 [name] sysname NOT NULL,
 [principal_id] int NOT NULL,
 [diagram_id] int IDENTITY(1,1) NOT NULL,
 [version] int NULL,
 [definition] varbinary(max) NULL);
GO
CREATE TABLE dbo.[TB_DISTRITO](
 [DISTRITO_ID] char(5) NOT NULL,
 [DISTRITO_DES] varchar(50) NOT NULL,
 [DISTRITO_ESTADO] bit NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[TB_MONEDA](
 [MONEDA_ID] tinyint NOT NULL,
 [MONEDA_SIMBOLO] varchar(10) NOT NULL,
 [MONEDA_ESTADO] bit NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[TiposCobranza](
 [Item] tinyint NOT NULL,
 [Nombre] nvarchar(50) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)),
 [VersionFila] timestamp NOT NULL);
GO
CREATE TABLE dbo.[TiposIdentificacion](
 [Item] tinyint NOT NULL,
 [Nombre] nvarchar(50) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)));
GO
CREATE TABLE dbo.[TiposPersona](
 [Item] tinyint NOT NULL,
 [Nombre] nvarchar(50) NOT NULL,
 [Activo] bit NOT NULL DEFAULT ((1)));
GO
ALTER TABLE dbo.[CargasInformacion] ADD CONSTRAINT [PK__CargasIn__9CBB9A4FC8EA5DEA] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[CartillasRP] ADD CONSTRAINT [PK__Cartilla__9CBB9A4F2A19D4DE] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[CartillasRP] ADD CONSTRAINT [UQ__Cartilla__7E532BC6BB21EBE3] UNIQUE NONCLUSTERED ([Numero] ASC);
GO
ALTER TABLE dbo.[ClavesCobranzas] ADD CONSTRAINT [PK__ClavesCo__E8181E10354F1DA5] PRIMARY KEY CLUSTERED ([Clave] ASC);
GO
ALTER TABLE dbo.[ConfiguracionCobranzas] ADD CONSTRAINT [PK__Configur__9CBB9A4FC553013F] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[CorreosTerritorioCobranza] ADD CONSTRAINT [PK__CorreosT__B16E1E55554EEE91] PRIMARY KEY CLUSTERED ([IdTerritorio] ASC);
GO
ALTER TABLE dbo.[DbCheques] ADD CONSTRAINT [PK_DbCheques] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[DbCheques] ADD CONSTRAINT [UQ_DbCheques_Codigo] UNIQUE NONCLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbCheques_IdCliente] ON dbo.[DbCheques] ([IdCliente] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbCheques_Banco] ON dbo.[DbCheques] ([Banco] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbCheques_Motivo] ON dbo.[DbCheques] ([Motivo] ASC);
GO
ALTER TABLE dbo.[DbClientes] ADD CONSTRAINT [PK_DbClientes] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DbClientes] ADD CONSTRAINT [UQ_DbClientes_Codigo] UNIQUE NONCLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbClientes_Territorio] ON dbo.[DbClientes] ([IdTerritorio] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Cliente_Distrito] ON dbo.[DbClientes] ([IdDistrito] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Cliente_Pais] ON dbo.[DbClientes] ([IdPais] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Cliente_DistritoEntrega] ON dbo.[DbClientes] ([IdDistritoEntrega] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Cliente_Nivel] ON dbo.[DbClientes] ([IdNivel] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Cliente_Cobranza] ON dbo.[DbClientes] ([IdCobranza] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Cliente_Comprobante] ON dbo.[DbClientes] ([IdComprobante] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Cliente_Moneda] ON dbo.[DbClientes] ([IdMoneda] ASC);
GO
ALTER TABLE dbo.[DbClientes2] ADD CONSTRAINT [PK_DbClientes2] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DbDevol] ADD CONSTRAINT [PK_DbDevol] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DbDevol] ADD CONSTRAINT [UQ_DbDevol_Codigo] UNIQUE NONCLUSTERED ([NroCar] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbDevol_IdCliente] ON dbo.[DbDevol] ([IdCliente] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbDevol_IdUsuario] ON dbo.[DbDevol] ([IdUsuario] ASC);
GO
ALTER TABLE dbo.[DBEmiFac] ADD CONSTRAINT [PK_DBEmiFac] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DBEmiFac] ADD CONSTRAINT [UQ_DBEmiFac_Codigo] UNIQUE NONCLUSTERED ([NumDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DBEmiFac_IdCliente] ON dbo.[DBEmiFac] ([IdCliente] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DBEmiFac_IdEstado] ON dbo.[DBEmiFac] ([IdEstado] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DBEmiFac_TipDoc] ON dbo.[DBEmiFac] ([TipDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DBEmiFac_Emision] ON dbo.[DBEmiFac] ([FecEmi] ASC);
GO
ALTER TABLE dbo.[DbGuias] ADD CONSTRAINT [PK_DbGuias] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DbGuias] ADD CONSTRAINT [UQ_DbGuias_Negocio] UNIQUE NONCLUSTERED ([NroGui] ASC,[Origen] ASC,[PickUP] ASC,[Produc] ASC,[Client] ASC,[Tipo] ASC);
GO
ALTER TABLE dbo.[DbGuias] ADD CONSTRAINT [UQ_DbGuias_Numero] UNIQUE NONCLUSTERED ([NroGui] ASC);
GO
ALTER TABLE dbo.[DbHisCourier] ADD CONSTRAINT [PK_DbHisCourier] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisCourier_IdFactura] ON dbo.[DbHisCourier] ([IdFactura] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisCourier_IdNotaCredito] ON dbo.[DbHisCourier] ([IdNotaCredito] ASC);
GO
ALTER TABLE dbo.[DbHisFac] ADD CONSTRAINT [PK_DbHisFac] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisFac_IdUsuario] ON dbo.[DbHisFac] ([IdUsuario] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisFac_IdEstado] ON dbo.[DbHisFac] ([IdEstado] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisFac_TipDoc] ON dbo.[DbHisFac] ([TipDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisFac_IdFactura] ON dbo.[DbHisFac] ([IdFactura] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisFac_IdNotaCredito] ON dbo.[DbHisFac] ([IdNotaCredito] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisFac_DocumentoFecha] ON dbo.[DbHisFac] ([NumDoc] ASC,[FecEst] DESC,[ITEM] DESC);
GO
CREATE NONCLUSTERED INDEX [IX_DbHisFac_Cargo] ON dbo.[DbHisFac] ([IdRegistroCargo] ASC);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_DbHisFac_CargoFac] ON dbo.[DbHisFac] ([IdRegistroCargo] ASC,[IdFactura] ASC) WHERE ([IdRegistroCargo] IS NOT NULL AND [IdFactura] IS NOT NULL);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_DbHisFac_CargoNC] ON dbo.[DbHisFac] ([IdRegistroCargo] ASC,[IdNotaCredito] ASC) WHERE ([IdRegistroCargo] IS NOT NULL AND [IdNotaCredito] IS NOT NULL);
GO
ALTER TABLE dbo.[DbNCredito] ADD CONSTRAINT [PK_DbNCredito] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[DbNCredito] ADD CONSTRAINT [UQ_DbNCredito_Codigo] UNIQUE NONCLUSTERED ([NumDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbNCredito_IdCliente] ON dbo.[DbNCredito] ([IdCliente] ASC);
GO
ALTER TABLE dbo.[DbQuejas_Cas] ADD CONSTRAINT [PK_DbQuejas_Cas] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[DbQuejas_Cas] ADD CONSTRAINT [UQ_DbQuejas_Cas_Codigo] UNIQUE NONCLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuejas_Cas_IdCliente] ON dbo.[DbQuejas_Cas] ([IdCliente] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuejas_Cas_IdUsuario] ON dbo.[DbQuejas_Cas] ([IdUsuario] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuejas_Cas_TipCas] ON dbo.[DbQuejas_Cas] ([TipCas] ASC);
GO
ALTER TABLE dbo.[DbQuejas_Cli] ADD CONSTRAINT [PK_DbQuejas_Cli] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DbQuejas_Cli] ADD CONSTRAINT [UQ_DbQuejas_Cli_Codigo] UNIQUE NONCLUSTERED ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuejas_Cli_IdCliente] ON dbo.[DbQuejas_Cli] ([IdCliente] ASC);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_Quejas_ClienteGestion] ON dbo.[DbQuejas_Cli] ([IdCliente] ASC) WHERE ([IdCliente] IS NOT NULL);
GO
ALTER TABLE dbo.[DbQuejas_Ges] ADD CONSTRAINT [PK_DbQuejas_Ges] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[DbQuejas_Ges] ADD CONSTRAINT [UQ_DbQuejas_Ges_Codigo] UNIQUE NONCLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuejas_Ges_IdUsuario] ON dbo.[DbQuejas_Ges] ([IdUsuario] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuejas_Ges_IdCaso] ON dbo.[DbQuejas_Ges] ([IdCaso] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuejas_Ges_Status] ON dbo.[DbQuejas_Ges] ([Status] ASC);
GO
ALTER TABLE dbo.[DbQuerie] ADD CONSTRAINT [PK_DbQuerie] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[DbQuerie] ADD CONSTRAINT [UQ_DbQuerie_Codigo] UNIQUE NONCLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuerie_IdCliente] ON dbo.[DbQuerie] ([IdCliente] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuerie_TipDoc] ON dbo.[DbQuerie] ([TipDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuerie_IdFactura] ON dbo.[DbQuerie] ([IdFactura] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbQuerie_IdNotaCredito] ON dbo.[DbQuerie] ([IdNotaCredito] ASC);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_DbQuerie_ClaveSolicitud] ON dbo.[DbQuerie] ([ClaveSolicitud] ASC) WHERE ([ClaveSolicitud] IS NOT NULL);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_DbQuerie_NotaResolucion] ON dbo.[DbQuerie] ([IdNotaResolucion] ASC) WHERE ([IdNotaResolucion] IS NOT NULL);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_DbQuerie_Abierta] ON dbo.[DbQuerie] ([IdFactura] ASC,[TipQue] ASC,[TipCas] ASC) WHERE ([IdFactura] IS NOT NULL AND [Estado]='Abierto' AND [SitDoc]='ACT');
GO
ALTER TABLE dbo.[DbRecojos] ADD CONSTRAINT [PK_DbRecojos] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DbRecojos] ADD CONSTRAINT [UQ_DbRecojos_Negocio] UNIQUE NONCLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbRecojos_IdCliente] ON dbo.[DbRecojos] ([IdCliente] ASC);
GO
ALTER TABLE dbo.[DbsCartas] ADD CONSTRAINT [PK_DbsCartas] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DbsCartas] ADD CONSTRAINT [UQ_DbsCartas_Negocio] UNIQUE NONCLUSTERED ([Correl] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsCartas_IdCliente] ON dbo.[DbsCartas] ([IdCliente] ASC);
GO
ALTER TABLE dbo.[DbsCompRp] ADD CONSTRAINT [PK_DbsCompRp] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[DbsCompRp] ADD CONSTRAINT [UQ_DbsCompRp_Codigo] UNIQUE NONCLUSTERED ([ASIENT] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsCompRp_IdCliente] ON dbo.[DbsCompRp] ([IdCliente] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsCompRp_MONEDA] ON dbo.[DbsCompRp] ([MONEDA] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsCompRp_IdGuia] ON dbo.[DbsCompRp] ([IdGuia] ASC);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_DbsCompRp_Guia] ON dbo.[DbsCompRp] ([IdGuia] ASC) WHERE ([IdGuia] IS NOT NULL);
GO
ALTER TABLE dbo.[DbsDiarRp] ADD CONSTRAINT [PK_DbsDiarRp] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[DbsDiarRp] ADD CONSTRAINT [UQ_DbsDiarRp_Codigo] UNIQUE NONCLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsDiarRp_IdCliente] ON dbo.[DbsDiarRp] ([IdCliente] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsDiarRp_Estado] ON dbo.[DbsDiarRp] ([Estado] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsDiarRp_Moneda] ON dbo.[DbsDiarRp] ([Moneda] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsDiarRp_IdGuia] ON dbo.[DbsDiarRp] ([IdGuia] ASC);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_DbsDiarRp_Guia] ON dbo.[DbsDiarRp] ([IdGuia] ASC) WHERE ([IdGuia] IS NOT NULL);
GO
ALTER TABLE dbo.[DbServGui] ADD CONSTRAINT [PK_DbServGui] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_DbServGui_Linea] ON dbo.[DbServGui] ([NumDoc] ASC,[Linea] ASC) WHERE ([Linea] IS NOT NULL);
GO
CREATE NONCLUSTERED INDEX [IX_DbServGui_TipDoc] ON dbo.[DbServGui] ([TipDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbServGui_IdFactura] ON dbo.[DbServGui] ([IdFactura] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbServGui_IdNotaCredito] ON dbo.[DbServGui] ([IdNotaCredito] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbServGui_IdGuia] ON dbo.[DbServGui] ([IdGuia] ASC);
GO
ALTER TABLE dbo.[DbsHisRP] ADD CONSTRAINT [PK_DbsHisRP] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsHisRP_IdDiario] ON dbo.[DbsHisRP] ([IdDiario] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_DbsHisRP_IdCompilado] ON dbo.[DbsHisRP] ([IdCompilado] ASC);
GO
ALTER TABLE dbo.[Departamentos] ADD CONSTRAINT [PK_Departamentos] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Departamentos] ADD CONSTRAINT [UQ__Departam__06370DAC1E62A3E7] UNIQUE NONCLUSTERED ([Codigo] ASC);
GO
ALTER TABLE dbo.[DetallesCargos] ADD CONSTRAINT [PK__Detalles__9CBB9A4FA7C4E44D] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[DetallesCargos] ADD CONSTRAINT [UQ_DetallesCargos_Documento] UNIQUE NONCLUSTERED ([IdRegistro] ASC,[Tipo] ASC,[IdDocumento] ASC);
GO
ALTER TABLE dbo.[DetallesCartillaRP] ADD CONSTRAINT [PK__Detalles__AA74F5D587BE8DAC] PRIMARY KEY CLUSTERED ([Planilla] ASC,[IdGuia] ASC);
GO
ALTER TABLE dbo.[Distritos] ADD CONSTRAINT [PK_Distritos] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Distritos] ADD CONSTRAINT [UQ__Distrito__06370DAC9EE1BA74] UNIQUE NONCLUSTERED ([Codigo] ASC);
GO
ALTER TABLE dbo.[DocumentosCobranzas] ADD CONSTRAINT [PK__Document__9CBB9A4FA76D4246] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_CobDocRecojo] ON dbo.[DocumentosCobranzas] ([IdRecojo] ASC,[IdFactura] ASC) WHERE ([IdRecojo] IS NOT NULL);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_CobDocCarta] ON dbo.[DocumentosCobranzas] ([IdCarta] ASC,[IdFactura] ASC) WHERE ([IdCarta] IS NOT NULL);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_CobDocDevol] ON dbo.[DocumentosCobranzas] ([IdDevolucion] ASC,[IdFactura] ASC) WHERE ([IdDevolucion] IS NOT NULL);
GO
ALTER TABLE dbo.[DocumentosResolucionQuery] ADD CONSTRAINT [PK__Document__A3EA1AA9D3897A4F] PRIMARY KEY CLUSTERED ([IdQuery] ASC,[IdDocumento] ASC);
GO
ALTER TABLE dbo.[EmisionesCarta] ADD CONSTRAINT [PK__Emisione__9CBB9A4F0D48E867] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[EmisionesCarta] ADD CONSTRAINT [UQ__Emisione__0B7AE2F223AAF467] UNIQUE NONCLUSTERED ([IdCarta] ASC,[Numero] ASC);
GO
ALTER TABLE dbo.[EnviosCobranzas] ADD CONSTRAINT [PK__EnviosCo__9CBB9A4FA49FB2F1] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[EstadosInicialesCargos] ADD CONSTRAINT [PK__EstadosI__E0242B8134023407] PRIMARY KEY CLUSTERED ([Tipo] ASC,[IdDocumento] ASC);
GO
ALTER TABLE dbo.[EventosRegistrosCargos] ADD CONSTRAINT [PK__EventosR__9CBB9A4F5AC7FB1E] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[GuiasDevolucion] ADD CONSTRAINT [PK__GuiasDev__630D4AB6EC768D9A] PRIMARY KEY CLUSTERED ([IdDevolucion] ASC,[IdGuia] ASC);
GO
ALTER TABLE dbo.[GuiasQuery] ADD CONSTRAINT [PK__GuiasQue__D580D289E437B0AB] PRIMARY KEY CLUSTERED ([IdQuery] ASC,[IdGuia] ASC);
GO
ALTER TABLE dbo.[HistorialCobranzas] ADD CONSTRAINT [PK__Historia__9CBB9A4F7F563D35] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[HistorialQueries] ADD CONSTRAINT [PK__Historia__9CBB9A4FEAE06EB7] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[INDICES] ADD CONSTRAINT [PK_INDICES] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[Int_Access] ADD CONSTRAINT [PK_Int_Access] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[Int_Access] ADD CONSTRAINT [UQ_Int_Access_Negocio] UNIQUE NONCLUSTERED ([Usuari] ASC,[CodMod] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Int_Access_IdUsuario] ON dbo.[Int_Access] ([IdUsuario] ASC);
GO
ALTER TABLE dbo.[Int_Bancos] ADD CONSTRAINT [PK_Int_Bancos] PRIMARY KEY CLUSTERED ([CodBank] ASC);
GO
ALTER TABLE dbo.[Int_Bancos] ADD CONSTRAINT [UQ_Int_Bancos_Codigo] UNIQUE NONCLUSTERED ([NomBank] ASC);
GO
ALTER TABLE dbo.[Int_Casos] ADD CONSTRAINT [PK_Int_Casos] PRIMARY KEY CLUSTERED ([CodCas] ASC);
GO
ALTER TABLE dbo.[Int_ComentaryRP] ADD CONSTRAINT [PK_Int_ComentaryRP] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Int_ComentaryRP] ADD CONSTRAINT [UQ_Int_ComentaryRP_Codigo] UNIQUE NONCLUSTERED ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Int_ComentaryRP_IdCliente] ON dbo.[Int_ComentaryRP] ([IdCliente] ASC);
GO
ALTER TABLE dbo.[Int_Correl] ADD CONSTRAINT [PK_Int_Correl] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Int_Correl] ADD CONSTRAINT [UQ_Int_Correl_Codigo] UNIQUE NONCLUSTERED ([Usuari] ASC,[DivCor] ASC);
GO
ALTER TABLE dbo.[Int_DivRp] ADD CONSTRAINT [PK_Int_DivRp] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Int_DivRp] ADD CONSTRAINT [UQ_Int_DivRp_Codigo] UNIQUE NONCLUSTERED ([CodCli] ASC);
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_Int_DivRp_Cliente] ON dbo.[Int_DivRp] ([IdCliente] ASC) WHERE ([IdCliente] IS NOT NULL);
GO
ALTER TABLE dbo.[Int_EstadCheq] ADD CONSTRAINT [PK_Int_EstadCheq] PRIMARY KEY CLUSTERED ([CodEst] ASC);
GO
ALTER TABLE dbo.[Int_EstadCheq] ADD CONSTRAINT [UQ_Int_EstadCheq_Codigo] UNIQUE NONCLUSTERED ([Estado] ASC);
GO
ALTER TABLE dbo.[int_IATA] ADD CONSTRAINT [PK_int_IATA] PRIMARY KEY CLUSTERED ([codigo] ASC);
GO
ALTER TABLE dbo.[Int_InfoMail] ADD CONSTRAINT [PK_Int_InfoMail] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Int_InfoMail] ADD CONSTRAINT [UQ_Int_InfoMail_Negocio] UNIQUE NONCLUSTERED ([Usuari] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Int_InfoMail_IdUsuario] ON dbo.[Int_InfoMail] ([IdUsuario] ASC);
GO
ALTER TABLE dbo.[Int_InfoSys] ADD CONSTRAINT [PK_Int_InfoSys] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Int_MailDatQue] ADD CONSTRAINT [PK_Int_MailDatQue] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Int_MailDatQue] ADD CONSTRAINT [UQ_Int_MailDatQue_Codigo] UNIQUE NONCLUSTERED ([VarMail] ASC);
GO
ALTER TABLE dbo.[Int_ModulosUsuario] ADD CONSTRAINT [PK__Int_Modu__9CBB9A4F11614458] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Int_ModulosUsuario] ADD CONSTRAINT [UQ__Int_Modu__837D533B45C5251C] UNIQUE NONCLUSTERED ([CodMod] ASC);
GO
ALTER TABLE dbo.[Int_Moneda] ADD CONSTRAINT [PK_Int_Moneda] PRIMARY KEY CLUSTERED ([CodEst] ASC);
GO
ALTER TABLE dbo.[Int_PreRP] ADD CONSTRAINT [PK_Int_PreRP] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Int_PreRP_IdCliente] ON dbo.[Int_PreRP] ([IdCliente] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Int_PreRP_IdUsuario] ON dbo.[Int_PreRP] ([IdUsuario] ASC);
GO
ALTER TABLE dbo.[Int_PublicMail] ADD CONSTRAINT [PK_Int_PublicMail] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Int_StatusCas] ADD CONSTRAINT [PK_Int_StatusCas] PRIMARY KEY CLUSTERED ([CodSta] ASC);
GO
ALTER TABLE dbo.[Int_StatusFac] ADD CONSTRAINT [PK_Int_StatusFac] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[Int_StatusFac] ADD CONSTRAINT [UQ_Int_StatusFac_Negocio] UNIQUE NONCLUSTERED ([Status] ASC);
GO
ALTER TABLE dbo.[Int_StatusRP] ADD CONSTRAINT [PK_Int_StatusRP] PRIMARY KEY CLUSTERED ([CodEst] ASC);
GO
ALTER TABLE dbo.[Int_TipDoc] ADD CONSTRAINT [PK_Int_TipDoc] PRIMARY KEY CLUSTERED ([CodIte] ASC);
GO
ALTER TABLE dbo.[Int_Usuario] ADD CONSTRAINT [PK_Int_Usuario] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[Int_Usuario] ADD CONSTRAINT [UQ_Int_Usuario_Codigo] UNIQUE NONCLUSTERED ([Usuari] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_Int_Usuario_PerfilAcceso] ON dbo.[Int_Usuario] ([IdPerfilAcceso] ASC);
GO
ALTER TABLE dbo.[IntAccesPlant] ADD CONSTRAINT [PK_IntAccesPlant] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[IntAccesPlant] ADD CONSTRAINT [UQ_IntAccesPlant_Negocio] UNIQUE NONCLUSTERED ([Planti] ASC,[CodMod] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntAccesPlant_Planti] ON dbo.[IntAccesPlant] ([Planti] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntAccesPlant_IdMenu] ON dbo.[IntAccesPlant] ([IdMenu] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntAccesPlant_IdEstado] ON dbo.[IntAccesPlant] ([IdEstado] ASC);
GO
ALTER TABLE dbo.[IntAccess] ADD CONSTRAINT [PK_IntAccess] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[IntAccess] ADD CONSTRAINT [UQ_IntAccess_Negocio] UNIQUE NONCLUSTERED ([Usuari] ASC,[CodMod] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntAccess_Usuari] ON dbo.[IntAccess] ([Usuari] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntAccess_IdMenu] ON dbo.[IntAccess] ([IdMenu] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntAccess_IdEstado] ON dbo.[IntAccess] ([IdEstado] ASC);
GO
ALTER TABLE dbo.[IntBackup] ADD CONSTRAINT [PK_IntBackup] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntBackup] ADD CONSTRAINT [UQ_IntBackup_Codigo] UNIQUE NONCLUSTERED ([Archivo] ASC);
GO
ALTER TABLE dbo.[IntCargosEnterprice] ADD CONSTRAINT [PK_IntCargosEnterprice] PRIMARY KEY CLUSTERED ([CodCar] ASC);
GO
ALTER TABLE dbo.[IntCargosEnterprice] ADD CONSTRAINT [UQ_IntCargosEnterprice_Codigo] UNIQUE NONCLUSTERED ([DesCar] ASC);
GO
ALTER TABLE dbo.[IntComentaryRP] ADD CONSTRAINT [PK_IntComentaryRP] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntComentaryRP] ADD CONSTRAINT [UQ_IntComentaryRP_Codigo] UNIQUE NONCLUSTERED ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntComentaryRP_CodCli] ON dbo.[IntComentaryRP] ([CodCli] ASC);
GO
ALTER TABLE dbo.[IntDominioRP] ADD CONSTRAINT [PK_IntDominioRP] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntInfoBody] ADD CONSTRAINT [PK_IntInfoBody] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntInfoBody] ADD CONSTRAINT [UQ_IntInfoBody_Codigo] UNIQUE NONCLUSTERED ([Body] ASC);
GO
ALTER TABLE dbo.[IntInfoMail] ADD CONSTRAINT [PK_IntInfoMail] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntInfoMail] ADD CONSTRAINT [UQ_IntInfoMail_Codigo] UNIQUE NONCLUSTERED ([Grupos] ASC);
GO
ALTER TABLE dbo.[IntInforQuer] ADD CONSTRAINT [PK_IntInforQuer] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntInforQuer] ADD CONSTRAINT [UQ_IntInforQuer_Codigo] UNIQUE NONCLUSTERED ([Descrip] ASC);
GO
ALTER TABLE dbo.[IntInfoSys] ADD CONSTRAINT [PK_IntInfoSys] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntInfoSys] ADD CONSTRAINT [UQ_IntInfoSys_Codigo] UNIQUE NONCLUSTERED ([ModAct] ASC);
GO
ALTER TABLE dbo.[IntInfoTerMail] ADD CONSTRAINT [PK_IntInfoTerMail] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntInfoTerMail_Territ] ON dbo.[IntInfoTerMail] ([Territ] ASC);
GO
ALTER TABLE dbo.[IntMailOrigen] ADD CONSTRAINT [PK_IntMailOrigen] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntMailOrigen] ADD CONSTRAINT [UQ_IntMailOrigen_Codigo] UNIQUE NONCLUSTERED ([Origen] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntMailOrigen_Grupos] ON dbo.[IntMailOrigen] ([Grupos] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntMailOrigen_Mensaj] ON dbo.[IntMailOrigen] ([Mensaj] ASC);
GO
ALTER TABLE dbo.[IntMotivoDevol] ADD CONSTRAINT [PK_IntMotivoDevol] PRIMARY KEY CLUSTERED ([CodMot] ASC);
GO
ALTER TABLE dbo.[IntMotivoDevol] ADD CONSTRAINT [UQ_IntMotivoDevol_Codigo] UNIQUE NONCLUSTERED ([DesMot] ASC);
GO
ALTER TABLE dbo.[IntNiveles] ADD CONSTRAINT [PK_IntNiveles] PRIMARY KEY CLUSTERED ([item] ASC);
GO
ALTER TABLE dbo.[IntNiveles] ADD CONSTRAINT [UQ_IntNiveles_Codigo] UNIQUE NONCLUSTERED ([Descrip] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntNiveles_Planti] ON dbo.[IntNiveles] ([Planti] ASC);
GO
ALTER TABLE dbo.[IntPlantillas] ADD CONSTRAINT [PK_IntPlantillas] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntPlantillas] ADD CONSTRAINT [UQ_IntPlantillas_Codigo] UNIQUE NONCLUSTERED ([NomPla] ASC);
GO
ALTER TABLE dbo.[IntPrealertRP] ADD CONSTRAINT [PK_IntPrealertRP] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntPrealertRP_CodCli] ON dbo.[IntPrealertRP] ([CodCli] ASC);
GO
ALTER TABLE dbo.[IntProdGuias] ADD CONSTRAINT [PK_IntProdGuias] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntProdGuias] ADD CONSTRAINT [UQ_IntProdGuias_Codigo] UNIQUE NONCLUSTERED ([Product] ASC);
GO
ALTER TABLE dbo.[IntRefeModalDevol] ADD CONSTRAINT [PK_IntRefeModalDevol] PRIMARY KEY CLUSTERED ([CodMod] ASC);
GO
ALTER TABLE dbo.[IntRefeModalDevol] ADD CONSTRAINT [UQ_IntRefeModalDevol_Codigo] UNIQUE NONCLUSTERED ([DesMod] ASC);
GO
ALTER TABLE dbo.[IntReferEnterprice] ADD CONSTRAINT [PK_IntReferEnterprice] PRIMARY KEY CLUSTERED ([CodRef] ASC);
GO
ALTER TABLE dbo.[IntReferEnterprice] ADD CONSTRAINT [UQ_IntReferEnterprice_Codigo] UNIQUE NONCLUSTERED ([DesRef] ASC);
GO
ALTER TABLE dbo.[IntReplaceRecojos] ADD CONSTRAINT [PK_IntReplaceRecojos] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_IntReplaceRecojos_Replace] ON dbo.[IntReplaceRecojos] ([Replace] ASC);
GO
ALTER TABLE dbo.[IntRespQuerie] ADD CONSTRAINT [PK_IntRespQuerie] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntRespQuerie] ADD CONSTRAINT [UQ_IntRespQuerie_Codigo] UNIQUE NONCLUSTERED ([Descrip] ASC);
GO
ALTER TABLE dbo.[IntSaveActualizar] ADD CONSTRAINT [PK_IntSaveActualizar] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntServFac] ADD CONSTRAINT [PK_IntServFac] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntServFac] ADD CONSTRAINT [UQ_IntServFac_Codigo] UNIQUE NONCLUSTERED ([Descrip] ASC);
GO
ALTER TABLE dbo.[IntStaRecojos] ADD CONSTRAINT [PK_IntStaRecojos] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntStaRecojos] ADD CONSTRAINT [UQ_IntStaRecojos_Codigo] UNIQUE NONCLUSTERED ([Status] ASC);
GO
ALTER TABLE dbo.[IntStatusCart] ADD CONSTRAINT [PK_IntStatusCart] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntStatusCart] ADD CONSTRAINT [UQ_IntStatusCart_Codigo] UNIQUE NONCLUSTERED ([Status] ASC);
GO
ALTER TABLE dbo.[IntTerritorios] ADD CONSTRAINT [PK_IntTerritorios] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntTerritorios] ADD CONSTRAINT [UQ_IntTerritorios_Codigo] UNIQUE NONCLUSTERED ([Descrip] ASC);
GO
ALTER TABLE dbo.[IntVarReplace] ADD CONSTRAINT [PK_IntVarReplace] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[IntVarReplace] ADD CONSTRAINT [UQ_IntVarReplace_Codigo] UNIQUE NONCLUSTERED ([Variable] ASC);
GO
ALTER TABLE dbo.[MdbAsientCancel] ADD CONSTRAINT [PK_MdbAsientCancel] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbAsientCancel_Divisa] ON dbo.[MdbAsientCancel] ([Divisa] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbAsientCancel_IdFactura] ON dbo.[MdbAsientCancel] ([IdFactura] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbAsientCancel_IdNotaCredito] ON dbo.[MdbAsientCancel] ([IdNotaCredito] ASC);
GO
ALTER TABLE dbo.[MdbCartas] ADD CONSTRAINT [PK_MdbCartas] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbCartas] ADD CONSTRAINT [UQ_MdbCartas_Codigo] UNIQUE NONCLUSTERED ([CodCar] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCartas_CodCli] ON dbo.[MdbCartas] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCartas_Cargo] ON dbo.[MdbCartas] ([Cargo] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCartas_Refere] ON dbo.[MdbCartas] ([Refere] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCartas_Divisa] ON dbo.[MdbCartas] ([Divisa] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCartas_Status] ON dbo.[MdbCartas] ([Status] ASC);
GO
ALTER TABLE dbo.[MdbCheques] ADD CONSTRAINT [PK_MdbCheques] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbCheques] ADD CONSTRAINT [UQ_MdbCheques_Codigo] UNIQUE NONCLUSTERED ([CodChe] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCheques_CodCli] ON dbo.[MdbCheques] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCheques_Cargo] ON dbo.[MdbCheques] ([Cargo] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCheques_Banco] ON dbo.[MdbCheques] ([Banco] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCheques_Motivo] ON dbo.[MdbCheques] ([Motivo] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCheques_Divisa] ON dbo.[MdbCheques] ([Divisa] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCheques_Status] ON dbo.[MdbCheques] ([Status] ASC);
GO
ALTER TABLE dbo.[MDbClientes] ADD CONSTRAINT [PK_MDbClientes] PRIMARY KEY CLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbClientes_Territ] ON dbo.[MDbClientes] ([Territ] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbClientes_TipCom] ON dbo.[MDbClientes] ([TipCom] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbClientes_Divisa] ON dbo.[MDbClientes] ([Divisa] ASC);
GO
ALTER TABLE dbo.[MdbCompRp] ADD CONSTRAINT [PK_MdbCompRp] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCompRp_CodCli] ON dbo.[MdbCompRp] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCompRp_Divisa] ON dbo.[MdbCompRp] ([Divisa] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCompRp_Produc] ON dbo.[MdbCompRp] ([Produc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCompRp_Status] ON dbo.[MdbCompRp] ([Status] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbCompRp_IdGuia] ON dbo.[MdbCompRp] ([IdGuia] ASC);
GO
ALTER TABLE dbo.[MdbDevol] ADD CONSTRAINT [PK_MdbDevol] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbDevol] ADD CONSTRAINT [UQ_MdbDevol_Codigo] UNIQUE NONCLUSTERED ([NroCar] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDevol_CodCli] ON dbo.[MdbDevol] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDevol_Usuari] ON dbo.[MdbDevol] ([Usuari] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDevol_ImpExp] ON dbo.[MdbDevol] ([ImpExp] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDevol_Motivo] ON dbo.[MdbDevol] ([Motivo] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDevol_Banco] ON dbo.[MdbDevol] ([Banco] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDevol_Divisa] ON dbo.[MdbDevol] ([Divisa] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDevol_Status] ON dbo.[MdbDevol] ([Status] ASC);
GO
ALTER TABLE dbo.[MdbDiarRp] ADD CONSTRAINT [PK_MdbDiarRp] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDiarRp_CodCli] ON dbo.[MdbDiarRp] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDiarRp_Divisa] ON dbo.[MdbDiarRp] ([Divisa] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDiarRp_Produc] ON dbo.[MdbDiarRp] ([Produc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDiarRp_Status] ON dbo.[MdbDiarRp] ([Status] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDiarRp_IdGuia] ON dbo.[MdbDiarRp] ([IdGuia] ASC);
GO
ALTER TABLE dbo.[MdbDocument] ADD CONSTRAINT [PK_MdbDocument] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbDocument] ADD CONSTRAINT [UQ_MdbDocument_Codigo] UNIQUE NONCLUSTERED ([NumDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDocument_CodCli] ON dbo.[MdbDocument] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDocument_TipDoc] ON dbo.[MdbDocument] ([TipDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDocument_DivFac] ON dbo.[MdbDocument] ([DivFac] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbDocument_Emision] ON dbo.[MdbDocument] ([FecEmi] ASC);
GO
ALTER TABLE dbo.[MdbGuias] ADD CONSTRAINT [PK_MdbGuias] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbGuias] ADD CONSTRAINT [UQ_MdbGuias_Negocio] UNIQUE NONCLUSTERED ([Guia] ASC,[Origen] ASC,[Fecha] ASC,[Produc] ASC,[Client] ASC,[Tipo] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbGuias_Produc] ON dbo.[MdbGuias] ([Produc] ASC);
GO
ALTER TABLE dbo.[MdbHisRP] ADD CONSTRAINT [PK_MdbHisRP] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbHisRP_Status] ON dbo.[MdbHisRP] ([Status] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbHisRP_IdDiario] ON dbo.[MdbHisRP] ([IdDiario] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbHisRP_IdCompilado] ON dbo.[MdbHisRP] ([IdCompilado] ASC);
GO
ALTER TABLE dbo.[MDBHistory] ADD CONSTRAINT [PK_MDBHistory] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDBHistory_Status] ON dbo.[MDBHistory] ([Status] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDBHistory_Usuari] ON dbo.[MDBHistory] ([Usuari] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDBHistory_TipDoc] ON dbo.[MDBHistory] ([TipDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDBHistory_IdFactura] ON dbo.[MDBHistory] ([IdFactura] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDBHistory_IdNotaCredito] ON dbo.[MDBHistory] ([IdNotaCredito] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDBHistory_Ultimo] ON dbo.[MDBHistory] ([NumDoc] ASC,[Active] ASC,[FecEmi] DESC,[ITEM] DESC) INCLUDE ([Status],[Usuari]);
GO
ALTER TABLE dbo.[MdbNCredito] ADD CONSTRAINT [PK_MdbNCredito] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbNCredito] ADD CONSTRAINT [UQ_MdbNCredito_Codigo] UNIQUE NONCLUSTERED ([Numdoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbNCredito_CodCli] ON dbo.[MdbNCredito] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbNCredito_Divisa] ON dbo.[MdbNCredito] ([Divisa] ASC);
GO
ALTER TABLE dbo.[MdbNotas] ADD CONSTRAINT [PK_MdbNotas] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbNotas] ADD CONSTRAINT [UQ_MdbNotas_Codigo] UNIQUE NONCLUSTERED ([NumDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbNotas_CodCli] ON dbo.[MdbNotas] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbNotas_Divisa] ON dbo.[MdbNotas] ([Divisa] ASC);
GO
ALTER TABLE dbo.[MdbPlantillaPri] ADD CONSTRAINT [PK_MdbPlantillaPri] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbPlantillaPri] ADD CONSTRAINT [UQ_MdbPlantillaPri_Codigo] UNIQUE NONCLUSTERED ([Codigo] ASC);
GO
ALTER TABLE dbo.[MDbQuerie] ADD CONSTRAINT [PK_MDbQuerie] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[MDbQuerie] ADD CONSTRAINT [UQ_MDbQuerie_Codigo] UNIQUE NONCLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_CodCli] ON dbo.[MDbQuerie] ([CodCli] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_TipDoc] ON dbo.[MDbQuerie] ([TipDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_DivSol] ON dbo.[MDbQuerie] ([DivSol] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_SolTer] ON dbo.[MDbQuerie] ([SolTer] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_SolRes] ON dbo.[MDbQuerie] ([SolRes] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_Servic] ON dbo.[MDbQuerie] ([Servic] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_TipCas] ON dbo.[MDbQuerie] ([TipCas] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_Proble] ON dbo.[MDbQuerie] ([Proble] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_IdFactura] ON dbo.[MDbQuerie] ([IdFactura] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbQuerie_IdNotaCredito] ON dbo.[MDbQuerie] ([IdNotaCredito] ASC);
GO
ALTER TABLE dbo.[MDbRecojos] ADD CONSTRAINT [PK_MDbRecojos] PRIMARY KEY CLUSTERED ([Asient] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MDbRecojos_Status] ON dbo.[MDbRecojos] ([Status] ASC);
GO
ALTER TABLE dbo.[MdbServGuias] ADD CONSTRAINT [PK_MdbServGuias] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[MdbServGuias] ADD CONSTRAINT [UQ_MdbServGuias_Negocio] UNIQUE NONCLUSTERED ([NroDoc] ASC,[TipoDoc] ASC,[Guia] ASC,[Servic] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbServGuias_TipoDoc] ON dbo.[MdbServGuias] ([TipoDoc] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbServGuias_Servic] ON dbo.[MdbServGuias] ([Servic] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbServGuias_IdFactura] ON dbo.[MdbServGuias] ([IdFactura] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbServGuias_IdNotaCredito] ON dbo.[MdbServGuias] ([IdNotaCredito] ASC);
GO
CREATE NONCLUSTERED INDEX [IX_MdbServGuias_IdGuia] ON dbo.[MdbServGuias] ([IdGuia] ASC);
GO
ALTER TABLE dbo.[MdbUsers] ADD CONSTRAINT [PK_MdbUsers] PRIMARY KEY CLUSTERED ([ITEM] ASC);
GO
ALTER TABLE dbo.[MdbUsers] ADD CONSTRAINT [UQ_MdbUsers_Codigo] UNIQUE NONCLUSTERED ([Usuari] ASC);
GO
ALTER TABLE dbo.[MigracionIncidencias] ADD CONSTRAINT [PK_MigracionIncidencias] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[NivelesCliente] ADD CONSTRAINT [PK_NivelesCliente] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[NivelesCliente] ADD CONSTRAINT [UQ__NivelesC__75E3EFCF07E3C2CE] UNIQUE NONCLUSTERED ([Nombre] ASC);
GO
ALTER TABLE dbo.[OtrosCasosQuery] ADD CONSTRAINT [PK__OtrosCas__9CBB9A4FB51387A1] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[OtrosCasosQuery] ADD CONSTRAINT [UQ__OtrosCas__92C53B6C723C4C5A] UNIQUE NONCLUSTERED ([Descripcion] ASC);
GO
ALTER TABLE dbo.[Paises] ADD CONSTRAINT [PK_Paises] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Paises] ADD CONSTRAINT [UQ__Paises__06370DACFA49E63C] UNIQUE NONCLUSTERED ([Codigo] ASC);
GO
ALTER TABLE dbo.[PerfilesAcceso] ADD CONSTRAINT [PK_PerfilesAcceso] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[PerfilesAcceso] ADD CONSTRAINT [UQ_PerfilesAcceso_Nombre] UNIQUE NONCLUSTERED ([Nombre] ASC);
GO
ALTER TABLE dbo.[PermisosPerfil] ADD CONSTRAINT [PK_PermisosPerfil] PRIMARY KEY CLUSTERED ([IdPerfil] ASC,[IdModulo] ASC);
GO
ALTER TABLE dbo.[Provincias] ADD CONSTRAINT [PK_Provincias] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[Provincias] ADD CONSTRAINT [UQ__Provinci__06370DAC40983594] UNIQUE NONCLUSTERED ([Codigo] ASC);
GO
ALTER TABLE dbo.[RegistrosCargos] ADD CONSTRAINT [PK__Registro__9CBB9A4F6B8B5121] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[RegistrosCargos] ADD CONSTRAINT [UQ__Registro__7E532BC610D54B8C] UNIQUE NONCLUSTERED ([Numero] ASC);
GO
ALTER TABLE dbo.[SeriesCargos] ADD CONSTRAINT [PK__SeriesCa__798299D67C171D77] PRIMARY KEY CLUSTERED ([Serie] ASC);
GO
ALTER TABLE dbo.[SeriesClientes] ADD CONSTRAINT [PK__SeriesCl__798299D621043507] PRIMARY KEY CLUSTERED ([Serie] ASC);
GO
ALTER TABLE dbo.[SeriesCobranzas] ADD CONSTRAINT [PK__SeriesCo__06370DAD0FB26738] PRIMARY KEY CLUSTERED ([Codigo] ASC);
GO
ALTER TABLE dbo.[SeriesQueries] ADD CONSTRAINT [PK__SeriesQu__798299D66435D598] PRIMARY KEY CLUSTERED ([Serie] ASC);
GO
ALTER TABLE dbo.[sysdiagrams] ADD CONSTRAINT [PK_sysdiagrams] PRIMARY KEY CLUSTERED ([diagram_id] ASC);
GO
ALTER TABLE dbo.[sysdiagrams] ADD CONSTRAINT [UQ_sysdiagrams_Nombre] UNIQUE NONCLUSTERED ([principal_id] ASC,[name] ASC);
GO
ALTER TABLE dbo.[TB_DISTRITO] ADD CONSTRAINT [PK_TB_DISTRITO] PRIMARY KEY CLUSTERED ([DISTRITO_ID] ASC);
GO
ALTER TABLE dbo.[TB_MONEDA] ADD CONSTRAINT [PK_TB_MONEDA] PRIMARY KEY CLUSTERED ([MONEDA_ID] ASC);
GO
ALTER TABLE dbo.[TiposCobranza] ADD CONSTRAINT [PK_TiposCobranza] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[TiposCobranza] ADD CONSTRAINT [UQ__TiposCob__75E3EFCF814AE705] UNIQUE NONCLUSTERED ([Nombre] ASC);
GO
ALTER TABLE dbo.[TiposIdentificacion] ADD CONSTRAINT [PK_TiposIdentificacion] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[TiposIdentificacion] ADD CONSTRAINT [UQ__TiposIde__75E3EFCF49805377] UNIQUE NONCLUSTERED ([Nombre] ASC);
GO
ALTER TABLE dbo.[TiposPersona] ADD CONSTRAINT [PK_TiposPersona] PRIMARY KEY CLUSTERED ([Item] ASC);
GO
ALTER TABLE dbo.[TiposPersona] ADD CONSTRAINT [UQ__TiposPer__75E3EFCF974C6851] UNIQUE NONCLUSTERED ([Nombre] ASC);
GO
ALTER TABLE dbo.[SeriesQueries] WITH CHECK ADD CONSTRAINT [CK__SeriesQue__Serie__03DB89B3] CHECK ([Serie]='Q');
GO
ALTER TABLE dbo.[SeriesQueries] WITH CHECK ADD CONSTRAINT [CK__SeriesQue__Numer__04CFADEC] CHECK ([Numero]>=(0) AND [Numero]<=(999999999));
GO
ALTER TABLE dbo.[DbHisFac] WITH CHECK ADD CONSTRAINT [CK_DbHisFac_Documento] CHECK ([IdFactura] IS NULL OR [IdNotaCredito] IS NULL);
GO
ALTER TABLE dbo.[DbHisCourier] WITH CHECK ADD CONSTRAINT [CK_DbHisCourier_Documento] CHECK ([IdFactura] IS NULL OR [IdNotaCredito] IS NULL);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [CK_DbQuerie_Documento] CHECK ([IdFactura] IS NULL OR [IdNotaCredito] IS NULL);
GO
ALTER TABLE dbo.[DbServGui] WITH CHECK ADD CONSTRAINT [CK_DbServGui_Documento] CHECK ([IdFactura] IS NULL OR [IdNotaCredito] IS NULL);
GO
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [CK_MDBHistory_Documento] CHECK ([IdFactura] IS NULL OR [IdNotaCredito] IS NULL);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [CK_MDbQuerie_Documento] CHECK ([IdFactura] IS NULL OR [IdNotaCredito] IS NULL);
GO
ALTER TABLE dbo.[MdbServGuias] WITH CHECK ADD CONSTRAINT [CK_MdbServGuias_Documento] CHECK ([IdFactura] IS NULL OR [IdNotaCredito] IS NULL);
GO
ALTER TABLE dbo.[MdbAsientCancel] WITH CHECK ADD CONSTRAINT [CK_MdbAsientCancel_Documento] CHECK ([IdFactura] IS NULL OR [IdNotaCredito] IS NULL);
GO
ALTER TABLE dbo.[MdbHisRP] WITH CHECK ADD CONSTRAINT [CK_MdbHisRP_Padre] CHECK ([IdDiario] IS NULL OR [IdCompilado] IS NULL);
GO
ALTER TABLE dbo.[Departamentos] WITH CHECK ADD CONSTRAINT [CK_Departamentos_Peru] CHECK ([IdPais]=(1));
GO
ALTER TABLE dbo.[DbsHisRP] WITH CHECK ADD CONSTRAINT [CK_DbsHisRP_Padre] CHECK ([IdDiario] IS NULL OR [IdCompilado] IS NULL);
GO
ALTER TABLE dbo.[DbQuejas_Cli] WITH CHECK ADD CONSTRAINT [CK_DbQuejas_Cli_CliTop] CHECK ([CliTop]=(1) OR [CliTop]=(0));
GO
ALTER TABLE dbo.[DbsCartas] WITH CHECK ADD CONSTRAINT [CK_DbsCartas_SgdaCar] CHECK ([SgdaCar]=(1) OR [SgdaCar]=(0));
GO
ALTER TABLE dbo.[DbsDiarRp] WITH CHECK ADD CONSTRAINT [CK_DbsDiarRp_Salida] CHECK ([Salida]=(1) OR [Salida]=(0));
GO
ALTER TABLE dbo.[DbsDiarRp] WITH CHECK ADD CONSTRAINT [CK_DbsDiarRp_Reten] CHECK ([Reten]=(1) OR [Reten]=(0));
GO
ALTER TABLE dbo.[Int_Bancos] WITH CHECK ADD CONSTRAINT [CK_Int_Bancos_ActEst] CHECK ([ActEst]=(1) OR [ActEst]=(0));
GO
ALTER TABLE dbo.[Int_Casos] WITH CHECK ADD CONSTRAINT [CK_Int_Casos_ActEst] CHECK ([ActEst]=(1) OR [ActEst]=(0));
GO
ALTER TABLE dbo.[Int_EstadCheq] WITH CHECK ADD CONSTRAINT [CK_Int_EstadCheq_ActEst] CHECK ([ActEst]=(1) OR [ActEst]=(0));
GO
ALTER TABLE dbo.[Int_Moneda] WITH CHECK ADD CONSTRAINT [CK_Int_Moneda_ActEst] CHECK ([ActEst]=(1) OR [ActEst]=(0));
GO
ALTER TABLE dbo.[Int_StatusCas] WITH CHECK ADD CONSTRAINT [CK_Int_StatusCas_ActEst] CHECK ([ActEst]=(1) OR [ActEst]=(0));
GO
ALTER TABLE dbo.[Int_StatusFac] WITH CHECK ADD CONSTRAINT [CK_Int_StatusFac_ActEst] CHECK ([ActEst]=(1) OR [ActEst]=(0));
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [CK_Cliente_Ubicacion] CHECK ([EsInternacional]=(0) AND [IdDistrito] IS NOT NULL AND [IdPais] IS NULL OR [EsInternacional]=(1) AND [IdPais] IS NOT NULL AND [IdPais]<>(1) AND [IdDistrito] IS NULL);
GO
ALTER TABLE dbo.[Int_StatusRP] WITH CHECK ADD CONSTRAINT [CK_Int_StatusRP_ActEst] CHECK ([ActEst]=(1) OR [ActEst]=(0));
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [CK_Cliente_Comprobante] CHECK ([IdComprobante]=(2) OR [IdComprobante]=(1));
GO
ALTER TABLE dbo.[Int_Usuario] WITH CHECK ADD CONSTRAINT [CK_Int_Usuario_ActPas] CHECK ([ActPas]=(1) OR [ActPas]=(0));
GO
ALTER TABLE dbo.[IntBackup] WITH CHECK ADD CONSTRAINT [CK_IntBackup_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntCargosEnterprice] WITH CHECK ADD CONSTRAINT [CK_IntCargosEnterprice_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntDominioRP] WITH CHECK ADD CONSTRAINT [CK_IntDominioRP_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntInfoBody] WITH CHECK ADD CONSTRAINT [CK_IntInfoBody_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntInfoMail] WITH CHECK ADD CONSTRAINT [CK_IntInfoMail_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntInfoTerMail] WITH CHECK ADD CONSTRAINT [CK_IntInfoTerMail_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntMailOrigen] WITH CHECK ADD CONSTRAINT [CK_IntMailOrigen_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntMotivoDevol] WITH CHECK ADD CONSTRAINT [CK_IntMotivoDevol_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntRefeModalDevol] WITH CHECK ADD CONSTRAINT [CK_IntRefeModalDevol_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntReferEnterprice] WITH CHECK ADD CONSTRAINT [CK_IntReferEnterprice_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntRespQuerie] WITH CHECK ADD CONSTRAINT [CK_IntRespQuerie_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntServFac] WITH CHECK ADD CONSTRAINT [CK_IntServFac_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[SeriesCobranzas] WITH CHECK ADD CONSTRAINT [CK__SeriesCob__Numer__2DD1C37F] CHECK ([Numero]>=(0) AND [Numero]<=(9999999));
GO
ALTER TABLE dbo.[IntStaRecojos] WITH CHECK ADD CONSTRAINT [CK_IntStaRecojos_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntStatusCart] WITH CHECK ADD CONSTRAINT [CK_IntStatusCart_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntTerritorios] WITH CHECK ADD CONSTRAINT [CK_IntTerritorios_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[IntVarReplace] WITH CHECK ADD CONSTRAINT [CK_IntVarReplace_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[MdbAsientCancel] WITH CHECK ADD CONSTRAINT [CK_MdbAsientCancel_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[MdbCartas] WITH CHECK ADD CONSTRAINT [CK_MdbCartas_SgdaCar] CHECK ([SgdaCar]=(1) OR [SgdaCar]=(0));
GO
ALTER TABLE dbo.[MdbCartas] WITH CHECK ADD CONSTRAINT [CK_MdbCartas_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[MdbCheques] WITH CHECK ADD CONSTRAINT [CK_MdbCheques_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[MDbClientes] WITH CHECK ADD CONSTRAINT [CK_MDbClientes_CliTnt] CHECK ([CliTnt]=(1) OR [CliTnt]=(0));
GO
ALTER TABLE dbo.[MDbClientes] WITH CHECK ADD CONSTRAINT [CK_MDbClientes_CliTop] CHECK ([CliTop]=(1) OR [CliTop]=(0));
GO
ALTER TABLE dbo.[ClavesCobranzas] WITH CHECK ADD CONSTRAINT [CK__ClavesCobranzas__375B2DB9] CHECK ((((((case when [IdCaso] IS NULL then (0) else (1) end+case when [IdGestion] IS NULL then (0) else (1) end)+case when [IdCheque] IS NULL then (0) else (1) end)+case when [IdRecojo] IS NULL then (0) else (1) end)+case when [IdCarta] IS NULL then (0) else (1) end)+case when [IdDevolucion] IS NULL then (0) else (1) end)=(1));
GO
ALTER TABLE dbo.[MDbClientes] WITH CHECK ADD CONSTRAINT [CK_MDbClientes_CliImp] CHECK ([CliImp]=(1) OR [CliImp]=(0));
GO
ALTER TABLE dbo.[MdbDevol] WITH CHECK ADD CONSTRAINT [CK_MdbDevol_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[MdbDiarRp] WITH CHECK ADD CONSTRAINT [CK_MdbDiarRp_Salida] CHECK ([Salida]=(1) OR [Salida]=(0));
GO
ALTER TABLE dbo.[MdbDiarRp] WITH CHECK ADD CONSTRAINT [CK_MdbDiarRp_Reten] CHECK ([Reten]=(1) OR [Reten]=(0));
GO
ALTER TABLE dbo.[MdbHisRP] WITH CHECK ADD CONSTRAINT [CK_MdbHisRP_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [CK_MDBHistory_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[MDbRecojos] WITH CHECK ADD CONSTRAINT [CK_MDbRecojos_Active] CHECK ([Active]=(1) OR [Active]=(0));
GO
ALTER TABLE dbo.[MdbUsers] WITH CHECK ADD CONSTRAINT [CK_MdbUsers_ActPas] CHECK ([ActPas]=(1) OR [ActPas]=(0));
GO
ALTER TABLE dbo.[MdbUsers] WITH CHECK ADD CONSTRAINT [CK_MdbUsers_SitUsu] CHECK ([SitUsu]=(1) OR [SitUsu]=(0));
GO
ALTER TABLE dbo.[SeriesClientes] WITH CHECK ADD CONSTRAINT [CK__SeriesCli__Serie__40257DE4] CHECK ([Serie]='CS');
GO
ALTER TABLE dbo.[HistorialCobranzas] WITH CHECK ADD CONSTRAINT [CK__HistorialCobranz__40E497F3] CHECK (((((case when [IdCaso] IS NULL then (0) else (1) end+case when [IdCheque] IS NULL then (0) else (1) end)+case when [IdRecojo] IS NULL then (0) else (1) end)+case when [IdCarta] IS NULL then (0) else (1) end)+case when [IdDevolucion] IS NULL then (0) else (1) end)=(1));
GO
ALTER TABLE dbo.[SeriesClientes] WITH CHECK ADD CONSTRAINT [CK__SeriesCli__Numer__4119A21D] CHECK ([Numero]>=(0));
GO
ALTER TABLE dbo.[DocumentosCobranzas] WITH CHECK ADD CONSTRAINT [CK__DocumentosCobran__47919582] CHECK (((case when [IdRecojo] IS NULL then (0) else (1) end+case when [IdCarta] IS NULL then (0) else (1) end)+case when [IdDevolucion] IS NULL then (0) else (1) end)=(1));
GO
ALTER TABLE dbo.[EmisionesCarta] WITH CHECK ADD CONSTRAINT [CK__Emisiones__Numer__5026DB83] CHECK ([Numero]=(2) OR [Numero]=(1));
GO
ALTER TABLE dbo.[ConfiguracionCobranzas] WITH CHECK ADD CONSTRAINT [CK__Configurac__Item__54EB90A0] CHECK ([Item]=(1));
GO
ALTER TABLE dbo.[Int_DivRp] WITH CHECK ADD CONSTRAINT [CK_Int_DivRp_Cliente] CHECK ([CodCli]='0000000000' AND [IdCliente] IS NULL OR [CodCli]<>'0000000000' AND [IdCliente] IS NOT NULL);
GO
ALTER TABLE dbo.[EnviosCobranzas] WITH CHECK ADD CONSTRAINT [CK__EnviosCob__Estad__5E74FADA] CHECK ([Estado]='Incierto' OR [Estado]='Error' OR [Estado]='Enviado' OR [Estado]='Enviando' OR [Estado]='Pendiente');
GO
ALTER TABLE dbo.[EnviosCobranzas] WITH CHECK ADD CONSTRAINT [CK__EnviosCobranzas__5F691F13] CHECK ((case when [IdRecojo] IS NULL then (0) else (1) end+case when [IdEmision] IS NULL then (0) else (1) end)=(1));
GO
ALTER TABLE dbo.[SeriesCargos] WITH CHECK ADD CONSTRAINT [CK__SeriesCar__Serie__627A95E8] CHECK ([Serie]='CRS');
GO
ALTER TABLE dbo.[SeriesCargos] WITH CHECK ADD CONSTRAINT [CK__SeriesCar__Numer__636EBA21] CHECK ([Numero]>=(0));
GO
ALTER TABLE dbo.[DetallesCargos] WITH CHECK ADD CONSTRAINT [CK_DetallesCargos_Documento] CHECK ([IdFactura] IS NOT NULL AND [IdNotaCredito] IS NULL AND [Tipo]<>'NC' AND [IdFactura]=[IdDocumento] OR [IdFactura] IS NULL AND [IdNotaCredito] IS NOT NULL AND [Tipo]='NC' AND [IdNotaCredito]=[IdDocumento]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK__DbQuerie__IdCerr__000AF8CF] FOREIGN KEY ([IdCerrador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[MdbAsientCancel] WITH CHECK ADD CONSTRAINT [FK_MdbAsientCancel_IdFactura] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[MdbDocument] ([Item]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK__DbQuerie__IdAnul__00FF1D08] FOREIGN KEY ([IdAnulador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[Provincias] WITH CHECK ADD CONSTRAINT [FK__Provincia__IdDep__01342732] FOREIGN KEY ([IdDepartamento]) REFERENCES dbo.[Departamentos] ([Item]);
GO
ALTER TABLE dbo.[MdbAsientCancel] WITH CHECK ADD CONSTRAINT [FK_MdbAsientCancel_IdNotaCredito] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[MdbNCredito] ([Item]);
GO
ALTER TABLE dbo.[DbServGui] WITH CHECK ADD CONSTRAINT [FK_DbServGui_IdGuia] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[DbGuias] ([Item]);
GO
ALTER TABLE dbo.[DbsDiarRp] WITH CHECK ADD CONSTRAINT [FK_DbsDiarRp_IdGuia] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[DbGuias] ([Item]);
GO
ALTER TABLE dbo.[DbsCompRp] WITH CHECK ADD CONSTRAINT [FK_DbsCompRp_IdGuia] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[DbGuias] ([Item]);
GO
ALTER TABLE dbo.[MdbServGuias] WITH CHECK ADD CONSTRAINT [FK_MdbServGuias_IdGuia] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[MdbGuias] ([Item]);
GO
ALTER TABLE dbo.[Distritos] WITH CHECK ADD CONSTRAINT [FK__Distritos__IdPro__05F8DC4F] FOREIGN KEY ([IdProvincia]) REFERENCES dbo.[Provincias] ([Item]);
GO
ALTER TABLE dbo.[MdbDiarRp] WITH CHECK ADD CONSTRAINT [FK_MdbDiarRp_IdGuia] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[MdbGuias] ([Item]);
GO
ALTER TABLE dbo.[MdbCompRp] WITH CHECK ADD CONSTRAINT [FK_MdbCompRp_IdGuia] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[MdbGuias] ([Item]);
GO
ALTER TABLE dbo.[HistorialQueries] WITH CHECK ADD CONSTRAINT [FK__Historial__IdQue__07AC1A97] FOREIGN KEY ([IdQuery]) REFERENCES dbo.[DbQuerie] ([ITEM]);
GO
ALTER TABLE dbo.[MdbHisRP] WITH CHECK ADD CONSTRAINT [FK_MdbHisRP_IdDiario] FOREIGN KEY ([IdDiario]) REFERENCES dbo.[MdbDiarRp] ([Item]);
GO
ALTER TABLE dbo.[MdbHisRP] WITH CHECK ADD CONSTRAINT [FK_MdbHisRP_IdCompilado] FOREIGN KEY ([IdCompilado]) REFERENCES dbo.[MdbCompRp] ([Item]);
GO
ALTER TABLE dbo.[HistorialQueries] WITH CHECK ADD CONSTRAINT [FK__Historial__IdUsu__09946309] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[IntAccess] WITH CHECK ADD CONSTRAINT [FK_IntAccess_IdMenu] FOREIGN KEY ([IdMenu]) REFERENCES dbo.[MdbPlantillaPri] ([Item]);
GO
ALTER TABLE dbo.[IntAccess] WITH CHECK ADD CONSTRAINT [FK_IntAccess_IdEstado] FOREIGN KEY ([IdEstado]) REFERENCES dbo.[Int_StatusFac] ([ITEM]);
GO
ALTER TABLE dbo.[IntAccesPlant] WITH CHECK ADD CONSTRAINT [FK_IntAccesPlant_IdMenu] FOREIGN KEY ([IdMenu]) REFERENCES dbo.[MdbPlantillaPri] ([Item]);
GO
ALTER TABLE dbo.[GuiasQuery] WITH CHECK ADD CONSTRAINT [FK__GuiasQuer__IdQue__0C70CFB4] FOREIGN KEY ([IdQuery]) REFERENCES dbo.[DbQuerie] ([ITEM]);
GO
ALTER TABLE dbo.[IntAccesPlant] WITH CHECK ADD CONSTRAINT [FK_IntAccesPlant_IdEstado] FOREIGN KEY ([IdEstado]) REFERENCES dbo.[Int_StatusFac] ([ITEM]);
GO
ALTER TABLE dbo.[GuiasQuery] WITH CHECK ADD CONSTRAINT [FK__GuiasQuer__IdGui__0D64F3ED] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[DbGuias] ([Item]);
GO
ALTER TABLE dbo.[DocumentosResolucionQuery] WITH CHECK ADD CONSTRAINT [FK__Documento__IdQue__10416098] FOREIGN KEY ([IdQuery]) REFERENCES dbo.[DbQuerie] ([ITEM]);
GO
ALTER TABLE dbo.[DocumentosResolucionQuery] WITH CHECK ADD CONSTRAINT [FK__Documento__IdDoc__113584D1] FOREIGN KEY ([IdDocumento]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[DbCheques] WITH CHECK ADD CONSTRAINT [FK_DbCheques_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbDevol] WITH CHECK ADD CONSTRAINT [FK_DbDevol_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DBEmiFac] WITH CHECK ADD CONSTRAINT [FK_DBEmiFac_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdPai__18178C8A] FOREIGN KEY ([IdPais]) REFERENCES dbo.[Paises] ([Item]);
GO
ALTER TABLE dbo.[DbNCredito] WITH CHECK ADD CONSTRAINT [FK_DbNCredito_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdDis__190BB0C3] FOREIGN KEY ([IdDistrito]) REFERENCES dbo.[Distritos] ([Item]);
GO
ALTER TABLE dbo.[DbQuejas_Cas] WITH CHECK ADD CONSTRAINT [FK_DbQuejas_Cas_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdDis__19FFD4FC] FOREIGN KEY ([IdDistritoEntrega]) REFERENCES dbo.[Distritos] ([Item]);
GO
ALTER TABLE dbo.[DbQuejas_Cli] WITH CHECK ADD CONSTRAINT [FK_DbQuejas_Cli_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdNiv__1AF3F935] FOREIGN KEY ([IdNivel]) REFERENCES dbo.[NivelesCliente] ([Item]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK_DbQuerie_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdCob__1BE81D6E] FOREIGN KEY ([IdCobranza]) REFERENCES dbo.[TiposCobranza] ([Item]);
GO
ALTER TABLE dbo.[DbRecojos] WITH CHECK ADD CONSTRAINT [FK_DbRecojos_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdCom__1CDC41A7] FOREIGN KEY ([IdComprobante]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[DbsCartas] WITH CHECK ADD CONSTRAINT [FK_DbsCartas_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdMon__1DD065E0] FOREIGN KEY ([IdMoneda]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[DbsCompRp] WITH CHECK ADD CONSTRAINT [FK_DbsCompRp_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdTip__1EC48A19] FOREIGN KEY ([IdTipoPersona]) REFERENCES dbo.[TiposPersona] ([Item]);
GO
ALTER TABLE dbo.[DbsDiarRp] WITH CHECK ADD CONSTRAINT [FK_DbsDiarRp_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK__DbCliente__IdTip__1FB8AE52] FOREIGN KEY ([IdTipoIdentificacion]) REFERENCES dbo.[TiposIdentificacion] ([Item]);
GO
ALTER TABLE dbo.[Int_ComentaryRP] WITH CHECK ADD CONSTRAINT [FK_Int_ComentaryRP_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[Int_PreRP] WITH CHECK ADD CONSTRAINT [FK_Int_PreRP_IdCliente] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[DbRecojos] WITH CHECK ADD CONSTRAINT [FK__DbRecojos__IdCre__216BEC9A] FOREIGN KEY ([IdCreador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbHisFac] WITH CHECK ADD CONSTRAINT [FK_DbHisFac_IdUsuario] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbRecojos] WITH CHECK ADD CONSTRAINT [FK__DbRecojos__IdCer__226010D3] FOREIGN KEY ([IdCerrador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbDevol] WITH CHECK ADD CONSTRAINT [FK_DbDevol_IdUsuario] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbQuejas_Cas] WITH CHECK ADD CONSTRAINT [FK_DbQuejas_Cas_IdUsuario] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbCheques] WITH CHECK ADD CONSTRAINT [FK__DbCheques__IdCre__24485945] FOREIGN KEY ([IdCreador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbQuejas_Ges] WITH CHECK ADD CONSTRAINT [FK_DbQuejas_Ges_IdUsuario] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbCheques] WITH CHECK ADD CONSTRAINT [FK__DbCheques__IdCer__253C7D7E] FOREIGN KEY ([IdCerrador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[Int_Access] WITH CHECK ADD CONSTRAINT [FK_Int_Access_IdUsuario] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[Int_InfoMail] WITH CHECK ADD CONSTRAINT [FK_Int_InfoMail_IdUsuario] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbsCartas] WITH CHECK ADD CONSTRAINT [FK__DbsCartas__IdCre__2724C5F0] FOREIGN KEY ([IdCreador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[Int_PreRP] WITH CHECK ADD CONSTRAINT [FK_Int_PreRP_IdUsuario] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbsCartas] WITH CHECK ADD CONSTRAINT [FK__DbsCartas__IdCer__2818EA29] FOREIGN KEY ([IdCerrador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbQuejas_Ges] WITH CHECK ADD CONSTRAINT [FK_DbQuejas_Ges_IdCaso] FOREIGN KEY ([IdCaso]) REFERENCES dbo.[DbQuejas_Cas] ([ITEM]);
GO
ALTER TABLE dbo.[DbsHisRP] WITH CHECK ADD CONSTRAINT [FK_DbsHisRP_IdDiario] FOREIGN KEY ([IdDiario]) REFERENCES dbo.[DbsDiarRp] ([ITEM]);
GO
ALTER TABLE dbo.[DbDevol] WITH CHECK ADD CONSTRAINT [FK__DbDevol__IdCread__2A01329B] FOREIGN KEY ([IdCreador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbsHisRP] WITH CHECK ADD CONSTRAINT [FK_DbsHisRP_IdCompilado] FOREIGN KEY ([IdCompilado]) REFERENCES dbo.[DbsCompRp] ([ITEM]);
GO
ALTER TABLE dbo.[DbDevol] WITH CHECK ADD CONSTRAINT [FK__DbDevol__IdCerra__2AF556D4] FOREIGN KEY ([IdCerrador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DBEmiFac] WITH CHECK ADD CONSTRAINT [FK_DBEmiFac_IdEstado] FOREIGN KEY ([IdEstado]) REFERENCES dbo.[Int_StatusFac] ([ITEM]);
GO
ALTER TABLE dbo.[Int_ModulosUsuario] WITH CHECK ADD CONSTRAINT [FK__Int_Modul__IdPad__2C1E8537] FOREIGN KEY ([IdPadre]) REFERENCES dbo.[Int_ModulosUsuario] ([Item]);
GO
ALTER TABLE dbo.[DbHisFac] WITH CHECK ADD CONSTRAINT [FK_DbHisFac_IdEstado] FOREIGN KEY ([IdEstado]) REFERENCES dbo.[Int_StatusFac] ([ITEM]);
GO
ALTER TABLE dbo.[DbCheques] WITH CHECK ADD CONSTRAINT [FK_DbCheques_Banco] FOREIGN KEY ([Banco]) REFERENCES dbo.[Int_Bancos] ([CodBank]);
GO
ALTER TABLE dbo.[DbCheques] WITH CHECK ADD CONSTRAINT [FK_DbCheques_Motivo] FOREIGN KEY ([Motivo]) REFERENCES dbo.[Int_EstadCheq] ([CodEst]);
GO
ALTER TABLE dbo.[DbQuejas_Cas] WITH CHECK ADD CONSTRAINT [FK_DbQuejas_Cas_TipCas] FOREIGN KEY ([TipCas]) REFERENCES dbo.[Int_Casos] ([CodCas]);
GO
ALTER TABLE dbo.[DbQuejas_Ges] WITH CHECK ADD CONSTRAINT [FK_DbQuejas_Ges_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[Int_StatusCas] ([CodSta]);
GO
ALTER TABLE dbo.[ClavesCobranzas] WITH CHECK ADD CONSTRAINT [FK__ClavesCob__IdUsu__30AE302A] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DbsDiarRp] WITH CHECK ADD CONSTRAINT [FK_DbsDiarRp_Estado] FOREIGN KEY ([Estado]) REFERENCES dbo.[Int_StatusRP] ([CodEst]);
GO
ALTER TABLE dbo.[ClavesCobranzas] WITH CHECK ADD CONSTRAINT [FK__ClavesCob__IdCas__31A25463] FOREIGN KEY ([IdCaso]) REFERENCES dbo.[DbQuejas_Cas] ([ITEM]);
GO
ALTER TABLE dbo.[DbsDiarRp] WITH CHECK ADD CONSTRAINT [FK_DbsDiarRp_Moneda] FOREIGN KEY ([Moneda]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[ClavesCobranzas] WITH CHECK ADD CONSTRAINT [FK__ClavesCob__IdGes__3296789C] FOREIGN KEY ([IdGestion]) REFERENCES dbo.[DbQuejas_Ges] ([ITEM]);
GO
ALTER TABLE dbo.[DbsCompRp] WITH CHECK ADD CONSTRAINT [FK_DbsCompRp_MONEDA] FOREIGN KEY ([MONEDA]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[ClavesCobranzas] WITH CHECK ADD CONSTRAINT [FK__ClavesCob__IdChe__338A9CD5] FOREIGN KEY ([IdCheque]) REFERENCES dbo.[DbCheques] ([ITEM]);
GO
ALTER TABLE dbo.[DBEmiFac] WITH CHECK ADD CONSTRAINT [FK_DBEmiFac_TipDoc] FOREIGN KEY ([TipDoc]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[ClavesCobranzas] WITH CHECK ADD CONSTRAINT [FK__ClavesCob__IdRec__347EC10E] FOREIGN KEY ([IdRecojo]) REFERENCES dbo.[DbRecojos] ([Item]);
GO
ALTER TABLE dbo.[DbHisFac] WITH CHECK ADD CONSTRAINT [FK_DbHisFac_TipDoc] FOREIGN KEY ([TipDoc]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[ClavesCobranzas] WITH CHECK ADD CONSTRAINT [FK__ClavesCob__IdCar__3572E547] FOREIGN KEY ([IdCarta]) REFERENCES dbo.[DbsCartas] ([Item]);
GO
ALTER TABLE dbo.[PermisosPerfil] WITH CHECK ADD CONSTRAINT [FK__PermisosP__IdPer__35A7EF71] FOREIGN KEY ([IdPerfil]) REFERENCES dbo.[PerfilesAcceso] ([Item]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK_DbQuerie_TipDoc] FOREIGN KEY ([TipDoc]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[ClavesCobranzas] WITH CHECK ADD CONSTRAINT [FK__ClavesCob__IdDev__36670980] FOREIGN KEY ([IdDevolucion]) REFERENCES dbo.[DbDevol] ([Item]);
GO
ALTER TABLE dbo.[PermisosPerfil] WITH CHECK ADD CONSTRAINT [FK__PermisosP__IdMod__369C13AA] FOREIGN KEY ([IdModulo]) REFERENCES dbo.[Int_ModulosUsuario] ([Item]);
GO
ALTER TABLE dbo.[DbServGui] WITH CHECK ADD CONSTRAINT [FK_DbServGui_TipDoc] FOREIGN KEY ([TipDoc]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[Int_Usuario] WITH CHECK ADD CONSTRAINT [FK_Int_Usuario_PerfilAcceso] FOREIGN KEY ([IdPerfilAcceso]) REFERENCES dbo.[PerfilesAcceso] ([Item]);
GO
ALTER TABLE dbo.[MdbDocument] WITH CHECK ADD CONSTRAINT [FK_MdbDocument_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[MdbNCredito] WITH CHECK ADD CONSTRAINT [FK_MdbNCredito_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[MdbNotas] WITH CHECK ADD CONSTRAINT [FK_MdbNotas_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[HistorialCobranzas] WITH CHECK ADD CONSTRAINT [FK__Historial__IdUsu__3B2BBE9D] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[MdbDiarRp] WITH CHECK ADD CONSTRAINT [FK_MdbDiarRp_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[HistorialCobranzas] WITH CHECK ADD CONSTRAINT [FK__Historial__IdCas__3C1FE2D6] FOREIGN KEY ([IdCaso]) REFERENCES dbo.[DbQuejas_Cas] ([ITEM]);
GO
ALTER TABLE dbo.[MdbCompRp] WITH CHECK ADD CONSTRAINT [FK_MdbCompRp_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[HistorialCobranzas] WITH CHECK ADD CONSTRAINT [FK__Historial__IdChe__3D14070F] FOREIGN KEY ([IdCheque]) REFERENCES dbo.[DbCheques] ([ITEM]);
GO
ALTER TABLE dbo.[MdbCartas] WITH CHECK ADD CONSTRAINT [FK_MdbCartas_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[HistorialCobranzas] WITH CHECK ADD CONSTRAINT [FK__Historial__IdRec__3E082B48] FOREIGN KEY ([IdRecojo]) REFERENCES dbo.[DbRecojos] ([Item]);
GO
ALTER TABLE dbo.[MdbCheques] WITH CHECK ADD CONSTRAINT [FK_MdbCheques_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[HistorialCobranzas] WITH CHECK ADD CONSTRAINT [FK__Historial__IdCar__3EFC4F81] FOREIGN KEY ([IdCarta]) REFERENCES dbo.[DbsCartas] ([Item]);
GO
ALTER TABLE dbo.[MdbDevol] WITH CHECK ADD CONSTRAINT [FK_MdbDevol_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[HistorialCobranzas] WITH CHECK ADD CONSTRAINT [FK__Historial__IdDev__3FF073BA] FOREIGN KEY ([IdDevolucion]) REFERENCES dbo.[DbDevol] ([Item]);
GO
ALTER TABLE dbo.[DbClientes] WITH CHECK ADD CONSTRAINT [FK_DbClientes_Territorio] FOREIGN KEY ([IdTerritorio]) REFERENCES dbo.[IntTerritorios] ([Item]);
GO
ALTER TABLE dbo.[IntComentaryRP] WITH CHECK ADD CONSTRAINT [FK_IntComentaryRP_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[IntPrealertRP] WITH CHECK ADD CONSTRAINT [FK_IntPrealertRP_CodCli] FOREIGN KEY ([CodCli]) REFERENCES dbo.[MDbClientes] ([Asient]);
GO
ALTER TABLE dbo.[MDbClientes] WITH CHECK ADD CONSTRAINT [FK_MDbClientes_Territ] FOREIGN KEY ([Territ]) REFERENCES dbo.[IntTerritorios] ([Item]);
GO
ALTER TABLE dbo.[DBEmiFac] WITH CHECK ADD CONSTRAINT [FK__DBEmiFac__IdDocu__4301EA8F] FOREIGN KEY ([IdDocumentoReferencia]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[MDbClientes] WITH CHECK ADD CONSTRAINT [FK_MDbClientes_TipCom] FOREIGN KEY ([TipCom]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[DocumentosCobranzas] WITH CHECK ADD CONSTRAINT [FK__Documento__IdRec__43C1049E] FOREIGN KEY ([IdRecojo]) REFERENCES dbo.[DbRecojos] ([Item]);
GO
ALTER TABLE dbo.[DbNCredito] WITH CHECK ADD CONSTRAINT [FK__DbNCredit__IdDoc__43F60EC8] FOREIGN KEY ([IdDocumentoReferencia]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[MDbClientes] WITH CHECK ADD CONSTRAINT [FK_MDbClientes_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[DocumentosCobranzas] WITH CHECK ADD CONSTRAINT [FK__Documento__IdCar__44B528D7] FOREIGN KEY ([IdCarta]) REFERENCES dbo.[DbsCartas] ([Item]);
GO
ALTER TABLE dbo.[MdbDocument] WITH CHECK ADD CONSTRAINT [FK_MdbDocument_TipDoc] FOREIGN KEY ([TipDoc]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[DocumentosCobranzas] WITH CHECK ADD CONSTRAINT [FK__Documento__IdDev__45A94D10] FOREIGN KEY ([IdDevolucion]) REFERENCES dbo.[DbDevol] ([Item]);
GO
ALTER TABLE dbo.[MdbDocument] WITH CHECK ADD CONSTRAINT [FK_MdbDocument_DivFac] FOREIGN KEY ([DivFac]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[DocumentosCobranzas] WITH CHECK ADD CONSTRAINT [FK__Documento__IdFac__469D7149] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[MdbNCredito] WITH CHECK ADD CONSTRAINT [FK_MdbNCredito_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[CargasInformacion] WITH CHECK ADD CONSTRAINT [FK__CargasInf__IdUsu__47C69FAC] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[MdbNotas] WITH CHECK ADD CONSTRAINT [FK_MdbNotas_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[Int_StatusFac] ([ITEM]);
GO
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_Usuari] FOREIGN KEY ([Usuari]) REFERENCES dbo.[MdbUsers] ([ITEM]);
GO
ALTER TABLE dbo.[GuiasDevolucion] WITH CHECK ADD CONSTRAINT [FK__GuiasDevo__IdDev__4A6E022D] FOREIGN KEY ([IdDevolucion]) REFERENCES dbo.[DbDevol] ([Item]);
GO
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_TipDoc] FOREIGN KEY ([TipDoc]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[GuiasDevolucion] WITH CHECK ADD CONSTRAINT [FK__GuiasDevo__IdGui__4B622666] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[DbGuias] ([Item]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_TipDoc] FOREIGN KEY ([TipDoc]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[CartillasRP] WITH CHECK ADD CONSTRAINT [FK__Cartillas__IdUsu__4C8B54C9] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_DivSol] FOREIGN KEY ([DivSol]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_SolTer] FOREIGN KEY ([SolTer]) REFERENCES dbo.[IntTerritorios] ([Item]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_SolRes] FOREIGN KEY ([SolRes]) REFERENCES dbo.[IntRespQuerie] ([Item]);
GO
ALTER TABLE dbo.[EmisionesCarta] WITH CHECK ADD CONSTRAINT [FK__Emisiones__IdCar__4F32B74A] FOREIGN KEY ([IdCarta]) REFERENCES dbo.[DbsCartas] ([Item]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_Servic] FOREIGN KEY ([Servic]) REFERENCES dbo.[IntServFac] ([Item]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_TipCas] FOREIGN KEY ([TipCas]) REFERENCES dbo.[IntInforQuer] ([Item]);
GO
ALTER TABLE dbo.[DetallesCartillaRP] WITH CHECK ADD CONSTRAINT [FK__DetallesC__Plani__515009E6] FOREIGN KEY ([Planilla]) REFERENCES dbo.[CartillasRP] ([Numero]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_Proble] FOREIGN KEY ([Proble]) REFERENCES dbo.[IntInforQuer] ([Item]);
GO
ALTER TABLE dbo.[EmisionesCarta] WITH CHECK ADD CONSTRAINT [FK__Emisiones__IdUsu__520F23F5] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[DetallesCartillaRP] WITH CHECK ADD CONSTRAINT [FK__DetallesC__IdGui__52442E1F] FOREIGN KEY ([IdGuia]) REFERENCES dbo.[DbGuias] ([Item]);
GO
ALTER TABLE dbo.[MdbDiarRp] WITH CHECK ADD CONSTRAINT [FK_MdbDiarRp_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[MdbDiarRp] WITH CHECK ADD CONSTRAINT [FK_MdbDiarRp_Produc] FOREIGN KEY ([Produc]) REFERENCES dbo.[IntProdGuias] ([Item]);
GO
ALTER TABLE dbo.[MdbDiarRp] WITH CHECK ADD CONSTRAINT [FK_MdbDiarRp_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[Int_StatusRP] ([CodEst]);
GO
ALTER TABLE dbo.[Int_DivRp] WITH CHECK ADD CONSTRAINT [FK__Int_DivRp__IdCli__55209ACA] FOREIGN KEY ([IdCliente]) REFERENCES dbo.[DbClientes] ([Item]);
GO
ALTER TABLE dbo.[MdbCompRp] WITH CHECK ADD CONSTRAINT [FK_MdbCompRp_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[MdbCompRp] WITH CHECK ADD CONSTRAINT [FK_MdbCompRp_Produc] FOREIGN KEY ([Produc]) REFERENCES dbo.[IntProdGuias] ([Item]);
GO
ALTER TABLE dbo.[MdbCompRp] WITH CHECK ADD CONSTRAINT [FK_MdbCompRp_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[Int_StatusRP] ([CodEst]);
GO
ALTER TABLE dbo.[CorreosTerritorioCobranza] WITH CHECK ADD CONSTRAINT [FK__CorreosTe__IdTer__57C7FD4B] FOREIGN KEY ([IdTerritorio]) REFERENCES dbo.[IntTerritorios] ([Item]);
GO
ALTER TABLE dbo.[MdbGuias] WITH CHECK ADD CONSTRAINT [FK_MdbGuias_Produc] FOREIGN KEY ([Produc]) REFERENCES dbo.[IntProdGuias] ([Item]);
GO
ALTER TABLE dbo.[MdbServGuias] WITH CHECK ADD CONSTRAINT [FK_MdbServGuias_TipoDoc] FOREIGN KEY ([TipoDoc]) REFERENCES dbo.[Int_TipDoc] ([CodIte]);
GO
ALTER TABLE dbo.[MdbServGuias] WITH CHECK ADD CONSTRAINT [FK_MdbServGuias_Servic] FOREIGN KEY ([Servic]) REFERENCES dbo.[IntServFac] ([Item]);
GO
ALTER TABLE dbo.[EnviosCobranzas] WITH CHECK ADD CONSTRAINT [FK__EnviosCob__IdRec__5AA469F6] FOREIGN KEY ([IdRecojo]) REFERENCES dbo.[DbRecojos] ([Item]);
GO
ALTER TABLE dbo.[MdbHisRP] WITH CHECK ADD CONSTRAINT [FK_MdbHisRP_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[Int_StatusRP] ([CodEst]);
GO
ALTER TABLE dbo.[EnviosCobranzas] WITH CHECK ADD CONSTRAINT [FK__EnviosCob__IdEmi__5B988E2F] FOREIGN KEY ([IdEmision]) REFERENCES dbo.[EmisionesCarta] ([Item]);
GO
ALTER TABLE dbo.[MDbRecojos] WITH CHECK ADD CONSTRAINT [FK_MDbRecojos_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[IntStaRecojos] ([Item]);
GO
ALTER TABLE dbo.[EnviosCobranzas] WITH CHECK ADD CONSTRAINT [FK__EnviosCob__IdUsu__5C8CB268] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[MdbAsientCancel] WITH CHECK ADD CONSTRAINT [FK_MdbAsientCancel_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[RegistrosCargos] WITH CHECK ADD CONSTRAINT [FK__Registros__IdEst__5DB5E0CB] FOREIGN KEY ([IdEstado]) REFERENCES dbo.[Int_StatusFac] ([ITEM]);
GO
ALTER TABLE dbo.[IntAccess] WITH CHECK ADD CONSTRAINT [FK_IntAccess_Usuari] FOREIGN KEY ([Usuari]) REFERENCES dbo.[MdbUsers] ([ITEM]);
GO
ALTER TABLE dbo.[RegistrosCargos] WITH CHECK ADD CONSTRAINT [FK__Registros__IdUsu__5EAA0504] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[IntAccesPlant] WITH CHECK ADD CONSTRAINT [FK_IntAccesPlant_Planti] FOREIGN KEY ([Planti]) REFERENCES dbo.[IntPlantillas] ([Item]);
GO
ALTER TABLE dbo.[IntNiveles] WITH CHECK ADD CONSTRAINT [FK_IntNiveles_Planti] FOREIGN KEY ([Planti]) REFERENCES dbo.[IntPlantillas] ([Item]);
GO
ALTER TABLE dbo.[IntInfoTerMail] WITH CHECK ADD CONSTRAINT [FK_IntInfoTerMail_Territ] FOREIGN KEY ([Territ]) REFERENCES dbo.[IntTerritorios] ([Item]);
GO
ALTER TABLE dbo.[IntMailOrigen] WITH CHECK ADD CONSTRAINT [FK_IntMailOrigen_Grupos] FOREIGN KEY ([Grupos]) REFERENCES dbo.[IntInfoMail] ([Item]);
GO
ALTER TABLE dbo.[IntMailOrigen] WITH CHECK ADD CONSTRAINT [FK_IntMailOrigen_Mensaj] FOREIGN KEY ([Mensaj]) REFERENCES dbo.[IntInfoBody] ([Item]);
GO
ALTER TABLE dbo.[IntReplaceRecojos] WITH CHECK ADD CONSTRAINT [FK_IntReplaceRecojos_Replace] FOREIGN KEY ([Replace]) REFERENCES dbo.[IntVarReplace] ([Item]);
GO
ALTER TABLE dbo.[MdbCartas] WITH CHECK ADD CONSTRAINT [FK_MdbCartas_Cargo] FOREIGN KEY ([Cargo]) REFERENCES dbo.[IntCargosEnterprice] ([CodCar]);
GO
ALTER TABLE dbo.[MdbCartas] WITH CHECK ADD CONSTRAINT [FK_MdbCartas_Refere] FOREIGN KEY ([Refere]) REFERENCES dbo.[IntReferEnterprice] ([CodRef]);
GO
ALTER TABLE dbo.[MdbCartas] WITH CHECK ADD CONSTRAINT [FK_MdbCartas_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[DetallesCargos] WITH CHECK ADD CONSTRAINT [FK__DetallesC__IdReg__673F4B05] FOREIGN KEY ([IdRegistro]) REFERENCES dbo.[RegistrosCargos] ([Item]);
GO
ALTER TABLE dbo.[MdbCartas] WITH CHECK ADD CONSTRAINT [FK_MdbCartas_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[IntStatusCart] ([Item]);
GO
ALTER TABLE dbo.[DetallesCargos] WITH CHECK ADD CONSTRAINT [FK__DetallesC__IdFac__68336F3E] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[MdbCheques] WITH CHECK ADD CONSTRAINT [FK_MdbCheques_Cargo] FOREIGN KEY ([Cargo]) REFERENCES dbo.[IntCargosEnterprice] ([CodCar]);
GO
ALTER TABLE dbo.[DetallesCargos] WITH CHECK ADD CONSTRAINT [FK__DetallesC__IdNot__69279377] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[DbNCredito] ([ITEM]);
GO
ALTER TABLE dbo.[MdbCheques] WITH CHECK ADD CONSTRAINT [FK_MdbCheques_Banco] FOREIGN KEY ([Banco]) REFERENCES dbo.[Int_Bancos] ([CodBank]);
GO
ALTER TABLE dbo.[MdbCheques] WITH CHECK ADD CONSTRAINT [FK_MdbCheques_Motivo] FOREIGN KEY ([Motivo]) REFERENCES dbo.[Int_EstadCheq] ([CodEst]);
GO
ALTER TABLE dbo.[MdbCheques] WITH CHECK ADD CONSTRAINT [FK_MdbCheques_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[DbHisFac] WITH CHECK ADD CONSTRAINT [FK__DbHisFac__IdRegi__6C040022] FOREIGN KEY ([IdRegistroCargo]) REFERENCES dbo.[RegistrosCargos] ([Item]);
GO
ALTER TABLE dbo.[MdbCheques] WITH CHECK ADD CONSTRAINT [FK_MdbCheques_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[IntStatusCart] ([Item]);
GO
ALTER TABLE dbo.[MdbDevol] WITH CHECK ADD CONSTRAINT [FK_MdbDevol_Usuari] FOREIGN KEY ([Usuari]) REFERENCES dbo.[MdbUsers] ([ITEM]);
GO
ALTER TABLE dbo.[MdbDevol] WITH CHECK ADD CONSTRAINT [FK_MdbDevol_ImpExp] FOREIGN KEY ([ImpExp]) REFERENCES dbo.[IntRefeModalDevol] ([CodMod]);
GO
ALTER TABLE dbo.[MdbDevol] WITH CHECK ADD CONSTRAINT [FK_MdbDevol_Motivo] FOREIGN KEY ([Motivo]) REFERENCES dbo.[IntMotivoDevol] ([CodMot]);
GO
ALTER TABLE dbo.[MdbDevol] WITH CHECK ADD CONSTRAINT [FK_MdbDevol_Banco] FOREIGN KEY ([Banco]) REFERENCES dbo.[Int_Bancos] ([CodBank]);
GO
ALTER TABLE dbo.[EventosRegistrosCargos] WITH CHECK ADD CONSTRAINT [FK__EventosRe__IdReg__70C8B53F] FOREIGN KEY ([IdRegistro]) REFERENCES dbo.[RegistrosCargos] ([Item]);
GO
ALTER TABLE dbo.[MdbDevol] WITH CHECK ADD CONSTRAINT [FK_MdbDevol_Divisa] FOREIGN KEY ([Divisa]) REFERENCES dbo.[Int_Moneda] ([CodEst]);
GO
ALTER TABLE dbo.[EventosRegistrosCargos] WITH CHECK ADD CONSTRAINT [FK__EventosRe__IdUsu__71BCD978] FOREIGN KEY ([IdUsuario]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[MdbDevol] WITH CHECK ADD CONSTRAINT [FK_MdbDevol_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[IntStatusCart] ([Item]);
GO
ALTER TABLE dbo.[DbHisFac] WITH CHECK ADD CONSTRAINT [FK_DbHisFac_IdFactura] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[DbHisFac] WITH CHECK ADD CONSTRAINT [FK_DbHisFac_IdNotaCredito] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[DbNCredito] ([ITEM]);
GO
ALTER TABLE dbo.[DbHisCourier] WITH CHECK ADD CONSTRAINT [FK_DbHisCourier_IdFactura] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[DbHisCourier] WITH CHECK ADD CONSTRAINT [FK_DbHisCourier_IdNotaCredito] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[DbNCredito] ([ITEM]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK_DbQuerie_IdFactura] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK_DbQuerie_IdNotaCredito] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[DbNCredito] ([ITEM]);
GO
ALTER TABLE dbo.[DbServGui] WITH CHECK ADD CONSTRAINT [FK_DbServGui_IdFactura] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[DBEmiFac] ([Item]);
GO
ALTER TABLE dbo.[DbServGui] WITH CHECK ADD CONSTRAINT [FK_DbServGui_IdNotaCredito] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[DbNCredito] ([ITEM]);
GO
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_IdFactura] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[MdbDocument] ([Item]);
GO
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_IdNotaCredito] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[MdbNCredito] ([Item]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK__DbQuerie__IdCrea__7C3A67EB] FOREIGN KEY ([IdCreador]) REFERENCES dbo.[Int_Usuario] ([ITEM]);
GO
ALTER TABLE dbo.[Departamentos] WITH CHECK ADD CONSTRAINT [FK__Departame__IdPai__7C6F7215] FOREIGN KEY ([IdPais]) REFERENCES dbo.[Paises] ([Item]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_IdFactura] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[MdbDocument] ([Item]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK__DbQuerie__IdResp__7D2E8C24] FOREIGN KEY ([IdResponsable]) REFERENCES dbo.[IntRespQuerie] ([Item]);
GO
ALTER TABLE dbo.[MDbQuerie] WITH CHECK ADD CONSTRAINT [FK_MDbQuerie_IdNotaCredito] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[MdbNCredito] ([Item]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK__DbQuerie__IdServ__7E22B05D] FOREIGN KEY ([IdServicio]) REFERENCES dbo.[IntServFac] ([Item]);
GO
ALTER TABLE dbo.[MdbServGuias] WITH CHECK ADD CONSTRAINT [FK_MdbServGuias_IdFactura] FOREIGN KEY ([IdFactura]) REFERENCES dbo.[MdbDocument] ([Item]);
GO
ALTER TABLE dbo.[DbQuerie] WITH CHECK ADD CONSTRAINT [FK__DbQuerie__IdNota__7F16D496] FOREIGN KEY ([IdNotaResolucion]) REFERENCES dbo.[DbNCredito] ([ITEM]);
GO
ALTER TABLE dbo.[MdbServGuias] WITH CHECK ADD CONSTRAINT [FK_MdbServGuias_IdNotaCredito] FOREIGN KEY ([IdNotaCredito]) REFERENCES dbo.[MdbNCredito] ([Item]);
GO
CREATE PROCEDURE [dbo].[AlterAsignarRP]
	@Guia numeric(18,0),
	@Cliente nvarchar(300),
	@Codigo numeric(18,0),
	@Producto nvarchar(10),
	@Origen nvarchar(50),
 @Registro int=NULL
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY

 IF @Registro IS NULL BEGIN
 IF(SELECT COUNT(*)FROM dbo.MdbDiarRp WHERE NroGui=@Guia)<>1 RAISERROR('Seleccione el registro RP por su identificador; guía inexistente o ambigua.',16,1);
 SELECT @Registro=Item FROM dbo.MdbDiarRp WHERE NroGui=@Guia;
 END;
 IF NOT EXISTS(SELECT 1 FROM dbo.MdbDiarRp WHERE Item=@Registro AND NroGui=@Guia) RAISERROR('Registro RP y guía no coinciden.',16,1);
	Declare @Prod int, @Ori int
	Set @Prod=ISNULL((SELECT Item FROM IntProdGuias WHERE Product=@Producto),(SELECT Item FROM IntProdGuias WHERE Product='NNE'))
	Update MdbDiarRp Set Client=@Cliente,CodCli=@Codigo,Produc=@Prod,Origen=@Origen WHERE Item=@Registro

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.ViewGuias
AS
SELECT     dbo.MdbGuias.Guia, ISNULL(dbo.int_IATA.Origen, dbo.MdbGuias.Origen) AS Origen, dbo.MdbGuias.Fecha, dbo.IntProdGuias.Product, 
                      dbo.MdbGuias.Client, dbo.MdbGuias.Tipo
FROM         dbo.MdbGuias LEFT OUTER JOIN
                      dbo.int_IATA ON dbo.MdbGuias.Origen = dbo.int_IATA.codigo INNER JOIN
                      dbo.IntProdGuias ON dbo.MdbGuias.Produc = dbo.IntProdGuias.Item
GO
CREATE VIEW dbo.ViewQueries
AS
SELECT     dbo.MDbQuerie.Asient AS Asiento, dbo.MDbQuerie.FecSol AS Emision, dbo.MDbQuerie.CliDoc AS Cliente, dbo.Int_TipDoc.TipDoc AS Tipo, 
                      dbo.MDbQuerie.NumDoc AS Documento, dbo.MDbQuerie.CodCli AS Codigo, dbo.MDbQuerie.FecEmi AS FechaEmi, dbo.MDbQuerie.MonSol AS Monto, 
                      dbo.Int_Moneda.Moneda AS Divisa, dbo.MDbQuerie.RcSol AS Recargo, dbo.IntTerritorios.Descrip AS Territorio, dbo.MDbQuerie.SolDa1, 
                      dbo.MDbQuerie.SolDa2, dbo.MDbQuerie.SolDa3, dbo.MDbQuerie.SoAnl1, dbo.MDbQuerie.SoAnl2, dbo.MDbQuerie.Solib1, dbo.MDbQuerie.Solib2, 
                      dbo.MDbQuerie.SolCli, dbo.MDbQuerie.SolRuc, dbo.MDbQuerie.SolDir, dbo.MDbQuerie.SOlTel, dbo.MDbQuerie.SolTa1, dbo.MDbQuerie.SolTa2, 
                      dbo.MDbQuerie.SolTa3, dbo.MDbQuerie.SolTa4, dbo.MDbQuerie.SolTa5, dbo.MDbQuerie.SolMa1, dbo.MDbQuerie.SolMa2, dbo.MDbQuerie.SolMa3, 
                      dbo.MDbQuerie.SolMa4, dbo.MDbQuerie.SolMa5, dbo.MDbQuerie.SoCon1, dbo.MDbQuerie.SoCon2, dbo.MDbQuerie.SoCon3, dbo.MDbQuerie.SoCli2, 
                      dbo.MDbQuerie.SoRuc2, dbo.MDbQuerie.SoDir2, dbo.MDbQuerie.SoTel2, dbo.IntServFac.Descrip AS Servicio, dbo.MDbQuerie.SoGuia AS Guias, 
                      dbo.MDbQuerie.SolAdi AS Adicional, dbo.MDbQuerie.SolDes AS Descuento, dbo.IntRespQuerie.Descrip AS Responsable, 
                      CASE dbo.MDbQuerie.SolEnc WHEN 1 THEN 'GAF' WHEN 2 THEN 'CONTABILIDAD' ELSE 'GAF' END AS Encargado, 
                      CASE dbo.MDbQuerie.SitDoc WHEN 1 THEN 'Activo' ELSE 'Anulado' END AS SitDoc, 
                      CASE dbo.MDbQuerie.TipQue WHEN 1 THEN 'ANL ANL' WHEN 2 THEN 'ANL NC' WHEN 3 THEN 'DES NC' WHEN 4 THEN 'DES 100' ELSE 'ANL ANL' END
                       AS TipoQuerie, CASE dbo.MDbQuerie.Estado WHEN 1 THEN 'Abierto' ELSE 'Cerrado' END AS Estado, 
                      CASE dbo.MDbQuerie.ORIGEN WHEN 1 THEN 'JPA' WHEN 2 THEN 'CAC' WHEN 3 THEN 'GGU' END AS Origen, ISNULL(dbo.IntInforQuer.Descrip, '') 
                      AS Caso, ISNULL(IntInforQuer_1.Descrip, '') AS Problema, dbo.MDbQuerie.FecIni AS Inicio, dbo.MDbQuerie.FecVis AS Visado
FROM dbo.MDbQuerie LEFT OUTER JOIN dbo.Int_TipDoc ON dbo.Int_TipDoc.CodIte = dbo.MDbQuerie.TipDoc LEFT OUTER JOIN
                      dbo.Int_Moneda ON dbo.MDbQuerie.DivSol = dbo.Int_Moneda.CodEst LEFT OUTER JOIN
                      dbo.IntTerritorios ON dbo.MDbQuerie.SolTer = dbo.IntTerritorios.Item LEFT OUTER JOIN
                      dbo.IntRespQuerie ON dbo.MDbQuerie.SolRes = dbo.IntRespQuerie.Item LEFT OUTER JOIN
                      dbo.IntInforQuer ON dbo.MDbQuerie.TipCas = dbo.IntInforQuer.Item LEFT OUTER JOIN
                      dbo.IntInforQuer AS IntInforQuer_1 ON dbo.MDbQuerie.Proble = IntInforQuer_1.Item LEFT OUTER JOIN
                      dbo.IntServFac ON dbo.MDbQuerie.Servic = dbo.IntServFac.Item
GO
CREATE VIEW dbo.InformeGuias AS
SELECT g.NroGui Codigo,g.Origen,g.PickUP FechaEnvio,g.Client Consignatario,g.Tipo Gestion,g.Produc Producto,(SELECT COUNT(*)FROM dbo.DbServGui s WHERE s.IdGuia=g.Item) ServiciosFacturados FROM dbo.DbGuias g;
GO
CREATE VIEW dbo.ViewCancelaciones
AS
SELECT     dbo.MdbAsientCancel.NumDoc AS Documento, dbo.Int_Moneda.Moneda AS Divisa, 
                      SUM(CASE divisa WHEN 1 THEN Monto / (CASE tipcam WHEN 0 THEN 1 ELSE tipcam END) 
                      WHEN 2 THEN Monto * (CASE tipcam WHEN 0 THEN 1 ELSE tipcam END) END) AS Monto
FROM         dbo.MdbAsientCancel INNER JOIN
                      dbo.MdbDocument ON dbo.MdbAsientCancel.NumDoc = dbo.MdbDocument.NumDoc INNER JOIN
                      dbo.Int_Moneda ON dbo.MdbDocument.DivFac = dbo.Int_Moneda.CodEst
GROUP BY dbo.MdbAsientCancel.NumDoc, dbo.Int_Moneda.Moneda
GO
CREATE VIEW dbo.InformeDiarioRP AS
SELECT d.Asient Codigo,d.NroGui Guia,d.Client Cliente,d.PuGui FechaEnvio,d.Origen,d.TipGui Producto,e.Status Estado,CASE WHEN d.Salida=1 THEN N'Sí'ELSE N'No'END ConSalida,CASE WHEN d.Reten=1 THEN N'Sí'ELSE N'No'END Retenido,d.PesFis PesoFisico,d.PesVol PesoVolumen,d.MonNet Neto,d.MonRc Recargo,d.MonIgv IGV,d.MonTot Total,d.Moneda,CONVERT(nvarchar(1000),d.Info) Observaciones FROM dbo.DbsDiarRp d LEFT JOIN dbo.Int_StatusRP e ON e.CodEst=d.Estado;
GO
CREATE PROCEDURE [dbo].[InsertServGuias]
	@Facturas nvarchar(20),
	@Tipo nvarchar(10),
	@Guia numeric(18,0),
	@Servicio nvarchar(100)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY


 IF NULLIF(LTRIM(RTRIM(@Servicio)),'') IS NULL RAISERROR('Indique el servicio.',16,1);
 DECLARE @ServicioId int,@TipoId tinyint,@Factura int,@Nota int,@GuiaId int;
 SELECT @TipoId=CodIte FROM dbo.Int_TipDoc WHERE TipDoc=@Tipo;
 IF @TipoId IS NULL RAISERROR('Tipo de documento inexistente.',16,1);
 SELECT @Factura=Item FROM dbo.MdbDocument WHERE NumDoc=@Facturas;
 SELECT @Nota=Item FROM dbo.MdbNCredito WHERE Numdoc=@Facturas;
 IF (@Factura IS NULL AND @Nota IS NULL) OR (@Factura IS NOT NULL AND @Nota IS NOT NULL) RAISERROR('Documento inexistente o ambiguo.',16,1);
 SELECT @ServicioId=Item FROM dbo.IntServFac WITH(UPDLOCK,HOLDLOCK) WHERE Descrip=@Servicio;
 IF @ServicioId IS NULL BEGIN INSERT dbo.IntServFac(Descrip)VALUES(@Servicio);SET @ServicioId=SCOPE_IDENTITY();END;
 SELECT @GuiaId=MIN(Item) FROM dbo.MdbGuias WHERE Guia=@Guia HAVING COUNT(*)=1;
 IF @GuiaId IS NULL RAISERROR('Seleccione una guía identificable sin ambigüedad.',16,1);
 IF NOT EXISTS(SELECT 1 FROM dbo.MdbServGuias WITH(UPDLOCK,HOLDLOCK) WHERE NroDoc=@Facturas AND TipoDoc=@TipoId AND Guia=@Guia AND Servic=@ServicioId)
 INSERT dbo.MdbServGuias(NroDoc,TipoDoc,Guia,Servic,IdFactura,IdNotaCredito,IdGuia) VALUES(@Facturas,@TipoId,@Guia,@ServicioId,@Factura,@Nota,@GuiaId);

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.ViewDevol
AS
SELECT     dbo.MdbDevol.Item, CONVERT(NVARCHAR(50), dbo.MdbDevol.NroCar) AS Codigo, dbo.MdbDevol.FecCar AS Emision, 
                      dbo.MdbUsers.Usuari AS Usuario, dbo.IntRefeModalDevol.DesMod AS Modalidad, dbo.MdbDevol.Client AS Cliente, dbo.MdbDevol.CodCli, 
                      CONVERT(nvarchar(20), dbo.MdbDevol.NroGui) AS Guia, dbo.MdbDevol.Cancel AS Cancelado, dbo.MdbDevol.Abonos, 
                      dbo.IntMotivoDevol.DesMot AS Motivo, dbo.MdbDevol.Factura, 
                      CASE dbo.MdbDevol.Modali WHEN 1 THEN 'TRANSFERENCIA' WHEN 2 THEN 'DEPOSITO' WHEN 3 THEN 'OTROS' END AS Modal, 
                      ISNULL(dbo.Int_Bancos.NomBank, '') AS Banco, dbo.MdbDevol.CtaCte AS Cuenta, dbo.MdbDevol.Monto, dbo.Int_Moneda.Moneda, 
                      dbo.MdbDevol.Contac AS Contacto, ISNULL(dbo.IntStatusCart.Status, '') AS Status, dbo.MdbDevol.FecSta AS Fecha, 
                      dbo.MdbDevol.Descrip AS Descripcion
FROM         dbo.MdbDevol INNER JOIN
                      dbo.MdbUsers ON dbo.MdbDevol.Usuari = dbo.MdbUsers.ITEM INNER JOIN
                      dbo.IntRefeModalDevol ON dbo.MdbDevol.ImpExp = dbo.IntRefeModalDevol.CodMod INNER JOIN
                      dbo.IntMotivoDevol ON dbo.MdbDevol.Motivo = dbo.IntMotivoDevol.CodMot LEFT OUTER JOIN
                      dbo.Int_Bancos ON dbo.MdbDevol.Banco = dbo.Int_Bancos.CodBank INNER JOIN
                      dbo.Int_Moneda ON dbo.MdbDevol.Divisa = dbo.Int_Moneda.CodEst LEFT OUTER JOIN
                      dbo.IntStatusCart ON dbo.MdbDevol.Status = dbo.IntStatusCart.Item
GO
-- Guardado de clientes normalizados; SQL Server 2008.
CREATE PROCEDURE dbo.GuardarCliente @Item int,@Version binary(8),
@Asient nvarchar(20),
@CLIENT nvarchar(255),
@RUCDNI nvarchar(20),
@DIRECC nvarchar(255),
@TELEF1 nvarchar(50),
@IdTerritorio int,
@Contac nvarchar(100),
@ContacEnt nvarchar(100),
@DireccEnt nvarchar(255),
@TelefoEnt nvarchar(50),
@Recarg money,
@IdComprobante tinyint,
@IdMoneda tinyint,
@IdNivel tinyint,
@IdCobranza tinyint,
@EsInternacional bit,
@IdPais int,
@IdDistrito int,
@IdDistritoEntrega int,
@IdTipoPersona tinyint,
@IdTipoIdentificacion tinyint, @Importando bit=0
AS
BEGIN
SET NOCOUNT ON;SET XACT_ABORT ON;
DECLARE @Inicio int,@Punto varchar(32);SET @Inicio=@@TRANCOUNT;SET @Punto=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
IF @Inicio=0 BEGIN TRAN;ELSE SAVE TRAN @Punto;
BEGIN TRY
IF @Item<>0 AND EXISTS(SELECT 1 FROM dbo.DbClientes WITH(UPDLOCK,HOLDLOCK) WHERE Item=@Item AND EsLocal<>CASE WHEN @Importando=1 THEN 0 ELSE 1 END) RAISERROR('La procedencia del cliente no permite esta operación.',16,1);
IF @Importando=0 AND @Item=0 BEGIN
 DECLARE @NumeroLocal int;
 UPDATE dbo.SeriesClientes WITH(UPDLOCK,HOLDLOCK) SET Numero=Numero+1;
 SELECT @NumeroLocal=Numero FROM dbo.SeriesClientes;
 IF @NumeroLocal>99999999 RAISERROR('Serie de clientes agotada.',16,1);
 SET @Asient='CS'+RIGHT('00000000'+CONVERT(varchar(8),@NumeroLocal),8);
END;
IF @Importando=0 AND @Item<>0 AND EXISTS(SELECT 1 FROM dbo.DbClientes WHERE Item=@Item AND Asient<>@Asient)RAISERROR('El código del cliente es inmutable.',16,1);
IF @Importando=1 AND @Asient LIKE 'CS%' RAISERROR('Código reservado para clientes locales.',16,1);
IF @Item IS NULL OR @Item<0 RAISERROR('Identificador inválido.',16,1);
IF NULLIF(LTRIM(RTRIM(@Asient)),'')IS NULL OR NULLIF(LTRIM(RTRIM(@CLIENT)),'')IS NULL OR NULLIF(LTRIM(RTRIM(@DIRECC)),'')IS NULL RAISERROR('Código, nombre y dirección son obligatorios.',16,1);
IF @Recarg IS NULL OR @Recarg<0 RAISERROR('El recargo porcentual no puede ser negativo.',16,1);
IF @IdTipoPersona NOT IN(1,2)OR @IdTipoPersona IS NULL OR @IdTipoIdentificacion IS NULL RAISERROR('Indique persona e identificación.',16,1);
IF NULLIF(LTRIM(RTRIM(@RUCDNI)),'')IS NULL RAISERROR('Ingrese identificación.',16,1);
IF @IdTipoIdentificacion=1 AND(LEN(@RUCDNI)<>8 OR @RUCDNI LIKE '%[^0-9]%'OR @IdTipoPersona<>1)RAISERROR('DNI: 8 dígitos, solo persona natural.',16,1);
IF @IdTipoIdentificacion=2 AND(LEN(@RUCDNI)<>11 OR @RUCDNI LIKE '%[^0-9]%' OR(@IdTipoPersona=1 AND LEFT(@RUCDNI,2)<>'10')OR(@IdTipoPersona=2 AND LEFT(@RUCDNI,2)<>'20'))RAISERROR('RUC: 11 dígitos; prefijo 10 para natural o 20 para jurídica.',16,1);
IF @IdTipoIdentificacion=3 AND(LEN(@RUCDNI)<>9 OR @RUCDNI LIKE '%[^0-9]%' OR @IdTipoPersona<>1)RAISERROR('Carné: 9 dígitos, persona natural.',16,1);
IF @IdTipoIdentificacion=4 AND @IdTipoPersona<>1 RAISERROR('Pasaporte solo para persona natural.',16,1);
IF @IdTipoIdentificacion=5 AND @IdTipoPersona<>2 RAISERROR('Identificación fiscal extranjera: persona jurídica.',16,1);
IF EXISTS(SELECT 1 FROM DbClientes WITH(UPDLOCK,HOLDLOCK)WHERE Asient=@Asient AND Item<>@Item)RAISERROR('El código ya existe, incluso si está eliminado.',16,1);
IF @Item<>0 AND NOT EXISTS(SELECT 1 FROM DbClientes WITH(UPDLOCK,HOLDLOCK)WHERE Item=@Item AND Activo=1 AND VersionFila=@Version)RAISERROR('El cliente cambió o fue eliminado. Recargue.',16,1);
IF @EsInternacional IS NULL OR(@EsInternacional=0 AND(@IdDistrito IS NULL OR @IdPais IS NOT NULL))OR(@EsInternacional=1 AND(@IdPais IS NULL OR @IdPais=1 OR @IdDistrito IS NOT NULL))RAISERROR('Ubicación principal: indique solo distrito nacional o país extranjero.',16,1);
IF @IdDistritoEntrega IS NULL RAISERROR('La entrega es nacional: seleccione distrito.',16,1);
IF NOT EXISTS(SELECT 1 FROM dbo.IntTerritorios WITH(HOLDLOCK)WHERE Item=@IdTerritorio AND(Active=1 OR EXISTS(SELECT 1 FROM DbClientes WHERE Item=@Item AND IdTerritorio=@IdTerritorio)))RAISERROR('Seleccione Territorio válido y activo.',16,1);
IF NOT EXISTS(SELECT 1 FROM dbo.NivelesCliente WITH(HOLDLOCK)WHERE Item=@IdNivel AND(Activo=1 OR EXISTS(SELECT 1 FROM DbClientes WHERE Item=@Item AND IdNivel=@IdNivel)))RAISERROR('Seleccione Nivel válido y activo.',16,1);
IF NOT EXISTS(SELECT 1 FROM dbo.TiposCobranza WITH(HOLDLOCK)WHERE Item=@IdCobranza AND(Activo=1 OR EXISTS(SELECT 1 FROM DbClientes WHERE Item=@Item AND IdCobranza=@IdCobranza)))RAISERROR('Seleccione Cobranza válido y activo.',16,1);
IF NOT EXISTS(SELECT 1 FROM dbo.TiposPersona WITH(HOLDLOCK)WHERE Item=@IdTipoPersona AND(Activo=1 OR EXISTS(SELECT 1 FROM DbClientes WHERE Item=@Item AND IdTipoPersona=@IdTipoPersona)))RAISERROR('Seleccione TipoPersona válido y activo.',16,1);
IF NOT EXISTS(SELECT 1 FROM dbo.TiposIdentificacion WITH(HOLDLOCK)WHERE Item=@IdTipoIdentificacion AND(Activo=1 OR EXISTS(SELECT 1 FROM DbClientes WHERE Item=@Item AND IdTipoIdentificacion=@IdTipoIdentificacion)))RAISERROR('Seleccione TipoIdentificacion válido y activo.',16,1);
IF @IdDistrito IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.Distritos WITH(HOLDLOCK)WHERE Item=@IdDistrito AND Activo=1)RAISERROR('Ubicación inexistente o inactiva.',16,1);
IF @IdPais IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.Paises WITH(HOLDLOCK)WHERE Item=@IdPais AND Activo=1)RAISERROR('Ubicación inexistente o inactiva.',16,1);
IF @IdDistritoEntrega IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.Distritos WITH(HOLDLOCK)WHERE Item=@IdDistritoEntrega AND Activo=1)RAISERROR('Ubicación inexistente o inactiva.',16,1);
IF @IdComprobante NOT IN(1,2) OR @IdComprobante IS NULL RAISERROR('Seleccione boleta o factura.',16,1);
IF NOT EXISTS(SELECT 1 FROM dbo.Int_Moneda WITH(HOLDLOCK)WHERE CodEst=@IdMoneda AND(ActEst=1 OR EXISTS(SELECT 1 FROM dbo.DbClientes WHERE Item=@Item AND IdMoneda=@IdMoneda)))RAISERROR('Seleccione una moneda activa.',16,1);
IF NOT EXISTS(SELECT 1 FROM dbo.Int_TipDoc WITH(HOLDLOCK)WHERE CodIte=@IdComprobante AND(Activo=1 OR EXISTS(SELECT 1 FROM dbo.DbClientes WHERE Item=@Item AND IdComprobante=@IdComprobante)))RAISERROR('Seleccione un comprobante activo.',16,1);
IF @Item=0 BEGIN
INSERT dbo.DbClientes(Asient,CLIENT,RUCDNI,DIRECC,TELEF1,IdTerritorio,Contac,ContacEnt,DireccEnt,TelefoEnt,Recarg,IdComprobante,IdMoneda,IdNivel,IdCobranza,EsInternacional,IdPais,IdDistrito,IdDistritoEntrega,IdTipoPersona,IdTipoIdentificacion)VALUES(@Asient,@CLIENT,@RUCDNI,@DIRECC,@TELEF1,@IdTerritorio,@Contac,@ContacEnt,@DireccEnt,@TelefoEnt,@Recarg,@IdComprobante,@IdMoneda,@IdNivel,@IdCobranza,@EsInternacional,@IdPais,@IdDistrito,@IdDistritoEntrega,@IdTipoPersona,@IdTipoIdentificacion);
SET @Item=SCOPE_IDENTITY(); UPDATE dbo.DbClientes SET EsLocal=CASE WHEN @Importando=1 THEN 0 ELSE 1 END WHERE Item=@Item;
END ELSE UPDATE dbo.DbClientes SET Asient=@Asient,CLIENT=@CLIENT,RUCDNI=@RUCDNI,DIRECC=@DIRECC,TELEF1=@TELEF1,IdTerritorio=@IdTerritorio,Contac=@Contac,ContacEnt=@ContacEnt,DireccEnt=@DireccEnt,TelefoEnt=@TelefoEnt,Recarg=@Recarg,IdComprobante=@IdComprobante,IdMoneda=@IdMoneda,IdNivel=@IdNivel,IdCobranza=@IdCobranza,EsInternacional=@EsInternacional,IdPais=@IdPais,IdDistrito=@IdDistrito,IdDistritoEntrega=@IdDistritoEntrega,IdTipoPersona=@IdTipoPersona,IdTipoIdentificacion=@IdTipoIdentificacion WHERE Item=@Item;
IF @Inicio=0 COMMIT;SELECT Item=@Item;
END TRY
BEGIN CATCH
IF @Inicio=0 AND XACT_STATE()<>0 ROLLBACK;ELSE IF @Inicio>0 AND XACT_STATE()=1 ROLLBACK TRAN @Punto;DECLARE @Mensaje nvarchar(2048);SET @Mensaje=ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);
END CATCH
END
GO
CREATE PROCEDURE [dbo].[ProSaveStatus]
	@Registro int,
	@Status nvarchar(150),
	@Descripcion nvarchar(250),
	@Fecha datetime,
	@Modal tinyint
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
IF @Modal NOT IN(1,2,3) OR NOT EXISTS(SELECT 1 FROM dbo.IntStatusCart WHERE Status=@Status) RAISERROR('Modalidad o estado inexistente.',16,1);
IF (@Modal=1 AND NOT EXISTS(SELECT 1 FROM dbo.MdbCartas WHERE Item=@Registro)) OR (@Modal=2 AND NOT EXISTS(SELECT 1 FROM dbo.MdbCheques WHERE Item=@Registro)) OR (@Modal=3 AND NOT EXISTS(SELECT 1 FROM dbo.MdbDevol WHERE Item=@Registro)) RAISERROR('Seleccione un registro existente.',16,1);

	Declare @Sta int
	Set @Sta=isnull((Select Item From IntStatusCart Where(Status=@Status)),0)
	If @Modal=1
		BEGIN
			Update MdbCartas Set Status=@Sta, FecSta=@Fecha, Descrip=@Descripcion Where(Item=@Registro)
		END
	Else
		If @Modal=2
			BEGIN
				Update MdbCheques Set Status=@Sta, FecSta=@Fecha, Descrip=@Descripcion Where(Item=@Registro)
			END
		Else
			BEGIN
				Update MdbDevol Set Status=@Sta, FecSta=@Fecha, Descrip=@Descripcion Where(Item=@Registro)
			END


 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE [dbo].[ProSaveCheque]
	@CodCli numeric(18,0),
	@Contac nvarchar(150),
	@Cargo nvarchar(150),
	@Banco nvarchar(150),
	@FecGir datetime,	
	@Cheque nvarchar(50),
	@Motivo nvarchar(50),
	@Gastos money,
	@Monto money,
	@Divisa nvarchar(10)

AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
DECLARE @Bloqueo numeric(18,0); SELECT @Bloqueo=Correl FROM dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) WHERE DivCor IN('CDC');
IF(SELECT COUNT(*)FROM dbo.Int_Correl WHERE DivCor='CDC')<>1 RAISERROR('Correlativo inexistente o ambiguo.',16,1);
	Insert into MdbCheques (CodChe,FecCar,Client,RucCli,DirCli,Contac,Cargo,CodCli,Banco,FecGir,NroChe,Motivo,Divisa,Gastos,MonCan,Status,FecSta,Descrip)
		Select top 1
			Codigo=(Select Correlativo=Correl+1 from Int_Correl Where(DivCor='CDC')),
			Emision=getdate(),
			Cliente=(Select top 1 Client FROM MDbClientes Where(Asient=@CodCli)),
			RucDni=(Select top 1 RucDni FROM MDbClientes Where(Asient=@CodCli)),
			Direccion=(Select top 1 Direcc FROM MDbClientes Where(Asient=@CodCli)),
			Contacto=@Contac,
			Cargo=(Select CodCar From IntCargosEnterprice Where(DesCar=@Cargo)),
			CodCli=@codcli,
			Banco=(Select CodBank From Int_Bancos Where(NomBank=@Banco)),
			FecGir=@FecGir,
			Cheque=@Cheque,
			Motivo=(Select CodEst From Int_EstadCheq Where(Estado=@Motivo)),
			Divisa=(Select CodEst From Int_Moneda Where(Moneda=@Divisa)),
			Gastos=@Gastos,
			MonCan=@Monto,
			Status=1,
			FecSta=getdate(),
			Descrip=''
		
		Update Int_Correl Set Correl=correl+1 Where(DivCor='CDC')

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE [dbo].[InsertDocument]
	@Tipo nvarchar(10),
	@Documento nvarchar(30),
	@Emision datetime,
	@Vencimiento datetime,
	@Codigo Numeric(18,0),
	@Cliente nvarchar(300),
	@Divisa nvarchar(10),
	@SubTotal Money,
	@Recargo Money,
	@Porsentaje Money,
	@Gastos Money,
	@IGV Money,
	@Importe Money,
	@Situacion nvarchar(50)

AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
IF NULLIF(LTRIM(RTRIM(@Documento)),'') IS NULL OR LEN(@Tipo+'-'+@Documento)>14 RAISERROR('Número de documento vacío o demasiado largo.',16,1);
IF @Emision IS NULL OR @Vencimiento IS NULL OR @Vencimiento<@Emision RAISERROR('Revise las fechas del documento.',16,1);

	
	Insert into MdbDocument (TipDoc,NumDoc,FecEmi,FecVen,CodCli,NomCli,DivFac,SubFac,RecFac,PorsRC,GasOtr,IgvFac,ImpFac,SitDoc)
	Select 
		Tipo=(SELECT CodIte From Int_TipDoc Where(TipDoc=@Tipo)),
		Documento=@Tipo + '-' + @Documento,
		Emision=@Emision,
		Vencimiento=@Vencimiento,
		Codigo=@Codigo,
		Cliente=@Cliente,
		Divisa=(Select CodEst From Int_Moneda Where(Moneda=@Divisa)),
		SubTotal=@SubTotal,
		Recargo=@Recargo,
		Porsentaje=@Porsentaje,
		Gastos=@Gastos,
		IGV=@IGV,
		Importe=@Importe,
		Situacion=(Case @Situacion WHEN 'ANULADO' THEN 0 ELSE 1 END)

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.ViewPlantillaPri
AS
SELECT     'Menus' AS Origen, Destino, Codigo
FROM         dbo.MdbPlantillaPri
UNION
SELECT     'Status' AS Origen, Status, CONVERT(int, IteAcc) AS Expr1
FROM         dbo.Int_StatusFac
GO
CREATE PROCEDURE dbo.ConsultarPermisosUsuario @Actor int,@Item int AS
BEGIN
 SET NOCOUNT ON;
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE Item=@Actor AND Access=N'ADMINISTRADOR'AND ActPas=1 AND SitUsu=N'Activo')BEGIN RAISERROR('Solo un administrador activo puede gestionar usuarios.',16,1);RETURN;END;
 SELECT m.Item,m.Nombre,m.IdPadre,m.Vigente,CONVERT(bit,CASE WHEN m.Vigente=0 THEN 0 WHEN m.Item IN(1,2)OR u.Access=N'ADMINISTRADOR'THEN 1 WHEN u.IdPerfilAcceso IS NULL THEN ISNULL(a.Access,0)ELSE ISNULL(p.Acceso,0)END)AS Acceso,
 (SELECT VersionFila FROM dbo.PerfilesAcceso WHERE Item=u.IdPerfilAcceso)AS VersionPerfil
 FROM dbo.Int_ModulosUsuario m LEFT JOIN dbo.Int_Usuario u ON u.Item=@Item
 LEFT JOIN dbo.Int_Access a ON a.IdUsuario=@Item AND a.CodMod=m.CodMod
 LEFT JOIN dbo.PermisosPerfil p ON p.IdPerfil=u.IdPerfilAcceso AND p.IdModulo=m.Item ORDER BY m.Item;
END
GO
CREATE PROCEDURE dbo.ReservarCorrelativo @Divisor nvarchar(3),@Numero numeric(18,0) OUTPUT
AS
BEGIN
 SET NOCOUNT ON;
 IF @@TRANCOUNT=0 BEGIN RAISERROR('Reserve dentro de la transacción de guardado.',16,1); RETURN; END;
 IF(SELECT COUNT(*)FROM dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) WHERE DivCor=@Divisor)<>1 BEGIN RAISERROR('Correlativo inexistente o ambiguo.',16,1); RETURN; END;
 DECLARE @Reserva TABLE(Numero numeric(18,0));
 UPDATE dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) SET Correl=Correl+1 OUTPUT inserted.Correl INTO @Reserva WHERE DivCor=@Divisor;
 SELECT @Numero=Numero FROM @Reserva;
END
GO
CREATE PROCEDURE [dbo].[InsertQuerie]
	@Asiento varchar(15),
	@Cliente varchar(300),
	@Tipo varchar(5),
	@Documento varchar(20),
	@Codcli numeric(18,0),
	@Recargo money,
	@Territorio varchar(20),
	@Dato1 tinyint,
	@Dato2 tinyint,
	@Dato3 tinyint,
	@Anl1 tinyint,
	@Anl2 tinyint,
	@LibSi tinyint,
	@LibNo tinyint,
	@ModClient nvarchar(250),
	@ModRucDni nvarchar(50),
	@ModDirecc nvarchar(250),
	@ModTelefo nvarchar(50),
	@Tarifa1 tinyint,
	@Tarifa2 tinyint,
	@Tarifa3 tinyint,
	@Tarifa4 tinyint,
	@Tarifa5 tinyint,
	@Mal1 tinyint,
	@Mal2 tinyint,
	@Mal3 tinyint,
	@Mal4 tinyint,
	@Mal5 tinyint,
	@Con1 tinyint,
	@Con2 tinyint,
	@Con3 tinyint,
	@ModClient2 nvarchar(250),
	@ModRucDni2 nvarchar(50),
	@ModDirecc2 nvarchar(250),
	@ModTelefo2 nvarchar(50),
	@Servicio nvarchar(100),
	@Guias nvarchar(150),
	@Adicional nvarchar(150),
	@Descuento money,
	@Responsable nvarchar(100),
	@Encargado nvarchar(50),
	@TipoQuerie nvarchar(50),
	@Origen nvarchar(10),
	@Caso nvarchar(150),
	@Problema nvarchar(150),
	@Inicio datetime,
	@Visto datetime,
	@CheDescuento as tinyint
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY

	
	INSERT INTO MDbQuerie
           (Asient,FecSol,CliDoc,TipDoc,NumDoc,CodCli,FecEmi,MonSol,DivSol,RcSol,SolTer,SolDa1,SolDa2,SolDa3,SoAnl1,SoAnl2,Solib1,Solib2,SolCli,SolRuc,SolDir,SOlTel,
            SolTa1,SolTa2,SolTa3,SolTa4,SolTa5,SolMa1,SolMa2,SolMa3,SolMa4,SolMa5,SoCon1,SoCon2,SoCon3,CheDes,SoCli2,SoRuc2,SoDir2,SoTel2,Servic,SoGuia,SolAdi,SolDes,SolRes,SolEnc,TipQue,ORIGEN,TipCas,Proble,FecIni,FecVis,IdFactura,IdNotaCredito)
		Select 
			Asiento=@Asiento,
			Emision=Getdate(),
			Cliente=@Cliente,
			Tipo=(Select CodIte From int_tipdoc Where(TipDoc=@Tipo)),
			NumDoc=@Documento,
			CodCli=@Codcli,
			DocEmi=(Select top 1 FecEmi From MdbDocument Where(NumDoc=@Documento)),
			Monto=(Select top 1 ImpFac From MdbDocument Where(NumDoc=@Documento)),
			Divisa=(Select top 1 DivFac From MdbDocument Where(NumDoc=@Documento)),
			Recargo=@Recargo,
			SolTer=(Select item From IntTerritorios Where(Descrip=@Territorio)),
			Dato1=@Dato1,
			Dato2=@Dato2,
			Dato3=@Dato3,
			Anl1=@Anl1,
			Anl2=@Anl2,
			SoLib1=@LibSi,
			SoLib2=@LibNo,
			ModCliente=@ModClient,
			ModRuc=@ModRucDni,
			ModDirecc=@ModDirecc,
			ModTelef=@ModTelefo,
			Tarifa1=@Tarifa1,
			Tarifa2=@Tarifa2,
			Tarifa3=@Tarifa3,
			Tarifa4=@Tarifa4,
			Tarifa5=@Tarifa5,
			Mal1=@Mal1,
			Mal2=@Mal2,
			Mal3=@Mal3,
			Mal4=@Mal4,
			Mal5=@Mal5,
			Con1=@Con1,
			Con2=@Con2,
			Con3=@Con3,
			CheDes=@CheDescuento,
			ModCliente2=@ModClient2,
			ModRuc2=@ModRucDni2,
			ModDirecc2=@ModDirecc2,
			ModTelef2=@ModTelefo2,
			Servicio=(Select item from IntServFac Where(Descrip=@Servicio)),
			SoGuia=@Guias,
			Adicional=@Adicional,
			Descuento=@Descuento,
			SolRes=(Select Item From IntRespQuerie Where(Descrip=@Responsable)),
			Encargado=Case upper(@Encargado) When 'GAF' then 1 when 'CONTABILIDAD' then 2 else 0 end,
			TipQuerie=Case @TipoQuerie When 'ANL ANL' then 1
									when 'ANL NC' then 2
									when 'DES NC' then 3
									when 'DES 100' then 4 else 1 end,
			Origen = Case @Origen When 'JPA' then 1
									when 'CAC' then 2
									when 'GGU' then 3 end,
			TipCas = (Select Item From Intinforquer Where(Descrip=@Caso)),
			Problem = (Select Item From Intinforquer Where(Descrip=@Problema)),
			Inicio=@Inicio,
			Vis=@Visto,
 IdFactura=(SELECT Item FROM MdbDocument WHERE NumDoc=@Documento),
 IdNotaCredito=(SELECT Item FROM MdbNCredito WHERE Numdoc=@Documento)

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.ViewUsers
AS
SELECT     dbo.MdbUsers.NomUsu + ' ' + dbo.MdbUsers.ApeUsu AS Usuario, dbo.MdbUsers.Usuari AS Login, ISNULL(dbo.IntNiveles.Descrip, 'PERSONALIZADO') 
                      AS Acceso
FROM         dbo.MdbUsers LEFT OUTER JOIN
                      dbo.IntNiveles ON dbo.MdbUsers.Access = dbo.IntNiveles.item
GO
CREATE VIEW dbo.ViewDatGen
AS
SELECT     'Cargo' AS Origen, DesCar
FROM         dbo.IntCargosEnterprice
UNION
SELECT     'Refer' AS Origen, UPPER(DesRef) AS Expr1
FROM         dbo.IntReferEnterprice
UNION
SELECT     'Divisa' AS Origen, UPPER(Moneda) AS Expr1
FROM         dbo.Int_Moneda
UNION
SELECT     'CtaSol' AS Modal, Infor1
FROM         dbo.Int_InfoSys
WHERE     (ModAct = 'CTACTESOLES')
UNION
SELECT     'CtaDol' AS Modal, Infor1
FROM         dbo.Int_InfoSys AS Int_InfoSys_2
WHERE     (ModAct = 'CTACTEDOLARES')
UNION
SELECT     'Moneda' AS Modal, Infor1
FROM         dbo.Int_InfoSys AS Int_InfoSys_1
WHERE     (ModAct = 'ANEXOCARTCOB')
UNION
SELECT     'Bancos' AS Modal, NomBank
FROM         dbo.Int_Bancos
UNION
SELECT     'Cheque' AS Modal, Estado
FROM         dbo.Int_EstadCheq
UNION
SELECT     'Modal' AS Modal, DesMod
FROM         dbo.IntRefeModalDevol
UNION
SELECT     'Motivo' AS Modal, DesMot
FROM         dbo.IntMotivoDevol
UNION
SELECT     'TipDoc' AS Modal, TipDoc
FROM         dbo.Int_TipDoc
UNION
SELECT     'RespQue' AS Modal, Descrip
FROM         dbo.IntRespQuerie
UNION
SELECT     'Servicios' AS Modal, Descrip
FROM         dbo.IntServFac
UNION
SELECT     'Territorio' AS Modal, Descrip
FROM         dbo.IntTerritorios
UNION
SELECT     'Niveles' AS Modal, Descrip
FROM         dbo.IntNiveles
GO
CREATE PROCEDURE dbo.GuardarCatalogo @Catalogo tinyint,@Codigo int,@Version binary(8),@Nombre nvarchar(150),@Abreviatura nvarchar(2)=NULL AS
BEGIN
SET NOCOUNT ON;SET XACT_ABORT ON;
IF @@TRANCOUNT<>0 BEGIN RAISERROR('El mantenimiento requiere una operación independiente.',16,1);RETURN;END;
BEGIN TRY
BEGIN TRAN;
IF @Catalogo NOT BETWEEN 1 AND 5 OR @Catalogo IS NULL OR @Codigo IS NULL OR @Codigo<0 RAISERROR('Catálogo o código inválido.',16,1);
SET @Nombre=LTRIM(RTRIM(@Nombre));SET @Abreviatura=UPPER(LTRIM(RTRIM(@Abreviatura)));
IF NULLIF(@Nombre,N'') IS NULL RAISERROR('Ingrese la descripción.',16,1);
DECLARE @Nuevo int,@Maximo int;
IF @Catalogo=1 BEGIN
IF LEN(@Nombre)>150 RAISERROR('La descripción admite 150 caracteres.',16,1);
SELECT @Maximo=ISNULL(MAX(Item),0)FROM dbo.IntTerritorios WITH(UPDLOCK,HOLDLOCK);
IF EXISTS(SELECT 1 FROM dbo.IntTerritorios WHERE Descrip=@Nombre AND Item<>@Codigo)RAISERROR('La descripción ya existe, incluso si está eliminada.',16,1);
IF @Codigo<>0 AND NOT EXISTS(SELECT 1 FROM dbo.IntTerritorios WHERE Item=@Codigo AND VersionFila=@Version AND Active=1)RAISERROR('El registro cambió o fue eliminado. Recargue.',16,1);
IF @Codigo=0 BEGIN
INSERT dbo.IntTerritorios(Descrip,Active)VALUES(@Nombre,1);SET @Codigo=CONVERT(int,SCOPE_IDENTITY());
END ELSE UPDATE dbo.IntTerritorios SET Descrip=@Nombre WHERE Item=@Codigo;
END
IF @Catalogo=2 BEGIN
IF LEN(@Nombre)>50 RAISERROR('La descripción admite 50 caracteres.',16,1);
SELECT @Maximo=ISNULL(MAX(Item),0)FROM dbo.NivelesCliente WITH(UPDLOCK,HOLDLOCK);
IF EXISTS(SELECT 1 FROM dbo.NivelesCliente WHERE Nombre=@Nombre AND Item<>@Codigo)RAISERROR('La descripción ya existe, incluso si está eliminada.',16,1);
IF @Codigo<>0 AND NOT EXISTS(SELECT 1 FROM dbo.NivelesCliente WHERE Item=@Codigo AND VersionFila=@Version AND Activo=1)RAISERROR('El registro cambió o fue eliminado. Recargue.',16,1);
IF @Codigo=0 BEGIN
IF @Maximo>=255 RAISERROR('Se alcanzó el límite de códigos de este catálogo.',16,1);
SET @Nuevo=@Maximo+1;
INSERT dbo.NivelesCliente(Item,Nombre,Activo)VALUES(@Nuevo,@Nombre,1);SET @Codigo=@Nuevo;
END ELSE UPDATE dbo.NivelesCliente SET Nombre=@Nombre WHERE Item=@Codigo;
END
IF @Catalogo=3 BEGIN
IF LEN(@Nombre)>50 RAISERROR('La descripción admite 50 caracteres.',16,1);
SELECT @Maximo=ISNULL(MAX(Item),0)FROM dbo.TiposCobranza WITH(UPDLOCK,HOLDLOCK);
IF EXISTS(SELECT 1 FROM dbo.TiposCobranza WHERE Nombre=@Nombre AND Item<>@Codigo)RAISERROR('La descripción ya existe, incluso si está eliminada.',16,1);
IF @Codigo<>0 AND NOT EXISTS(SELECT 1 FROM dbo.TiposCobranza WHERE Item=@Codigo AND VersionFila=@Version AND Activo=1)RAISERROR('El registro cambió o fue eliminado. Recargue.',16,1);
IF @Codigo=0 BEGIN
IF @Maximo>=255 RAISERROR('Se alcanzó el límite de códigos de este catálogo.',16,1);
SET @Nuevo=@Maximo+1;
INSERT dbo.TiposCobranza(Item,Nombre,Activo)VALUES(@Nuevo,@Nombre,1);SET @Codigo=@Nuevo;
END ELSE UPDATE dbo.TiposCobranza SET Nombre=@Nombre WHERE Item=@Codigo;
END
IF @Catalogo=4 BEGIN
IF LEN(@Nombre)>50 RAISERROR('La descripción admite 50 caracteres.',16,1);
SELECT @Maximo=ISNULL(MAX(CodEst),0)FROM dbo.Int_Moneda WITH(UPDLOCK,HOLDLOCK);
IF EXISTS(SELECT 1 FROM dbo.Int_Moneda WHERE Moneda=@Nombre AND CodEst<>@Codigo)RAISERROR('La descripción ya existe, incluso si está eliminada.',16,1);
IF @Codigo<>0 AND NOT EXISTS(SELECT 1 FROM dbo.Int_Moneda WHERE CodEst=@Codigo AND VersionFila=@Version AND ActEst=1)RAISERROR('El registro cambió o fue eliminado. Recargue.',16,1);
IF @Codigo=0 BEGIN
IF @Maximo>=255 RAISERROR('Se alcanzó el límite de códigos de este catálogo.',16,1);
SET @Nuevo=@Maximo+1;
INSERT dbo.Int_Moneda(CodEst,Moneda,ActEst)VALUES(@Nuevo,@Nombre,1);SET @Codigo=@Nuevo;
END ELSE UPDATE dbo.Int_Moneda SET Moneda=@Nombre WHERE CodEst=@Codigo;
END
IF @Catalogo=5 BEGIN
IF LEN(@Nombre)>50 RAISERROR('La descripción admite 50 caracteres.',16,1);
SELECT @Maximo=ISNULL(MAX(CodIte),0)FROM dbo.Int_TipDoc WITH(UPDLOCK,HOLDLOCK);
IF EXISTS(SELECT 1 FROM dbo.Int_TipDoc WHERE NamVar=@Nombre AND CodIte<>@Codigo)RAISERROR('La descripción ya existe, incluso si está eliminada.',16,1);
IF @Codigo<>0 AND NOT EXISTS(SELECT 1 FROM dbo.Int_TipDoc WHERE CodIte=@Codigo AND VersionFila=@Version AND Activo=1)RAISERROR('El registro cambió o fue eliminado. Recargue.',16,1);
IF LEN(ISNULL(@Abreviatura,N''))<>2 OR @Abreviatura COLLATE Latin1_General_BIN LIKE '%[^A-Z0-9]%' RAISERROR('Indique una abreviatura de dos letras o números.',16,1);
IF EXISTS(SELECT 1 FROM dbo.Int_TipDoc WHERE TipDoc=@Abreviatura AND CodIte<>@Codigo)RAISERROR('La abreviatura ya existe.',16,1);
IF @Codigo<>0 AND EXISTS(SELECT 1 FROM dbo.Int_TipDoc WHERE CodIte=@Codigo AND TipDoc<>@Abreviatura)RAISERROR('La abreviatura de un documento existente se conserva.',16,1);
IF @Codigo=0 BEGIN
IF @Maximo>=255 RAISERROR('Se alcanzó el límite de códigos de este catálogo.',16,1);
SET @Nuevo=@Maximo+1;
INSERT dbo.Int_TipDoc(CodIte,NamVar,Activo,TipDoc)VALUES(@Nuevo,@Nombre,1,@Abreviatura);SET @Codigo=@Nuevo;
END ELSE UPDATE dbo.Int_TipDoc SET NamVar=@Nombre WHERE CodIte=@Codigo;
END
COMMIT;SELECT Codigo=@Codigo;
END TRY
BEGIN CATCH
IF XACT_STATE()<>0 ROLLBACK;DECLARE @Mensaje nvarchar(2048);SET @Mensaje=ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);
END CATCH
END
GO
CREATE VIEW dbo.InformeRecojos AS
SELECT r.Asient Codigo,r.FecEnv FechaRegistro,c.CLIENT Cliente,r.ModGes Modalidad,r.Destino Destinatarios,r.EstadoGestion Estado,r.FechaCierre,r.MotivoCierre,(SELECT COUNT(*)FROM dbo.DocumentosCobranzas d WHERE d.IdRecojo=r.Item) Documentos,(SELECT TOP(1)e.Estado FROM dbo.EnviosCobranzas e WHERE e.IdRecojo=r.Item ORDER BY e.Fecha DESC,e.Item DESC) UltimoEnvio FROM dbo.DbRecojos r LEFT JOIN dbo.DbClientes c ON c.Item=r.IdCliente;
GO
CREATE PROCEDURE dbo.ConsultarPermisosPerfil @Actor int,@Item int AS
BEGIN
 SET NOCOUNT ON;
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE Item=@Actor AND Access=N'ADMINISTRADOR'AND ActPas=1 AND SitUsu=N'Activo')BEGIN RAISERROR('Solo un administrador activo puede gestionar perfiles.',16,1);RETURN;END;
 SELECT m.Item,m.Nombre,m.IdPadre,m.Vigente,CONVERT(bit,CASE WHEN m.Item IN(1,2)THEN 1 ELSE ISNULL(p.Acceso,0)END)AS Acceso,
 (SELECT VersionFila FROM dbo.PerfilesAcceso WHERE Item=@Item)AS VersionPerfil
 FROM dbo.Int_ModulosUsuario m LEFT JOIN dbo.PermisosPerfil p ON p.IdModulo=m.Item AND p.IdPerfil=@Item ORDER BY m.Item;
END
GO
CREATE VIEW dbo.ViewInfoTerNull
AS
SELECT     dbo.IntTerritorios.Descrip, dbo.IntInfoTerMail.DirTer, dbo.IntTerritorios.Item
FROM         dbo.IntInfoTerMail RIGHT OUTER JOIN
                      dbo.IntTerritorios ON dbo.IntInfoTerMail.Territ = dbo.IntTerritorios.Item
WHERE     (dbo.IntInfoTerMail.DirTer IS NULL)
GO
CREATE PROCEDURE [dbo].[VerifAccesStatusFac]
	@Usuari nvarchar(10)

AS
BEGIN
 SET NOCOUNT ON;
	declare @Access int
	Set @Access = isnull((Select Access From MdbUsers Where(Usuari=@Usuari)),0)
	
	
	if @Access=0
		BEGIN
			Select 
				ITEM,
				TitEst,
				Status,
				Acceso=CONVERT(int,isnull((SELECT dbo.IntAccess.Access FROM dbo.MdbUsers INNER JOIN
								dbo.IntAccess ON dbo.MdbUsers.ITEM = dbo.IntAccess.Usuari Where(dbo.IntAccess.CodMod=convert(int,dbo.Int_StatusFac.IteAcc)) and (dbo.MdbUsers.ITEM=(Select Item From MdbUsers Where(Usuari=@Usuari)))),0))
			From dbo.Int_StatusFac
		END
	ELSE
		if @Access=1
			BEGIN
				Select 
				ITEM,
				TitEst,
				Status,
				Acceso=1
			From dbo.Int_StatusFac
			END
		ELSE
		
			BEGIN
				Select 
					ITEM,
					TitEst,
					Status,
					Acceso=CONVERT(int,isnull((SELECT dbo.IntAccesPlant.Access
								FROM dbo.MdbUsers INNER JOIN dbo.IntNiveles ON dbo.MdbUsers.Access = dbo.IntNiveles.item INNER JOIN dbo.IntAccesPlant ON dbo.IntNiveles.Planti = dbo.IntAccesPlant.Planti Where(dbo.IntAccesPlant.CodMod=convert(int,dbo.Int_StatusFac.IteAcc)) and (dbo.MdbUsers.ITEM=(Select Item From MdbUsers Where(Usuari=@Usuari)))),0))
				From dbo.Int_StatusFac
			END
END
GO
CREATE PROCEDURE [dbo].[InsertNCredito]
	@Documento nvarchar(30),
	@Emision datetime,
	@Codigo Numeric(18,0),
	@Cliente nvarchar(300),
	@Divisa nvarchar(10),
	@Total Money,
	@Situacion nvarchar(50),
	@Referencia nvarchar(30)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY

	
	Insert into MdbNCredito (NumDoc,Emision,CodCli,Client,Divisa,Total,SitDoc,Refere)
	Select 
		Documento='NC-' + @Documento,
		Emision=@Emision,
		Codigo=@Codigo,
		Cliente=@Cliente,
		Divisa=(Select CodEst From Int_Moneda Where(Moneda=@Divisa)),
		SubTotal=@Total,
		SitDoc=Case @Situacion when 'PENDIENTE' THEN 1 WHEN 'ANULADO' THEN 0 ELSE 2 END,
		Referen=@Referencia

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.ViewStoryRP
AS
SELECT     dbo.MdbHisRP.Herencia, dbo.Int_StatusRP.Status, dbo.MdbHisRP.Info, dbo.MdbHisRP.Fecha
FROM         dbo.MdbHisRP INNER JOIN
                      dbo.Int_StatusRP ON dbo.MdbHisRP.Status = dbo.Int_StatusRP.CodEst
WHERE     (dbo.MdbHisRP.Active = 1)
GO
CREATE VIEW dbo.ViewClientes
AS
SELECT     dbo.MDbClientes.Asient AS Codigo, dbo.MDbClientes.Client AS Empresa, dbo.MDbClientes.RucDni, dbo.MDbClientes.Direcc AS Direccion, 
                      dbo.MDbClientes.Distri AS Distrito, dbo.MDbClientes.Telef1 AS Telefono1, dbo.MDbClientes.Telef2 AS Telefono2, dbo.MDbClientes.Telef3 AS Telefono3,
                       ISNULL(dbo.IntTerritorios.Descrip, '00') AS Territorio, CASE dbo.MDbClientes.CliTnt WHEN 0 THEN 'NO' WHEN 1 THEN 'SI' END AS ClienteTNT, 
                      CASE dbo.MDbClientes.CliTop WHEN 1 THEN 'SI' WHEN 0 THEN 'NO' END AS ClienteTOP, dbo.MDbClientes.Remark, 
                      CASE dbo.MDbClientes.CliImp WHEN 0 THEN 'NO' WHEN 1 THEN 'SI' END AS ClienteImport, dbo.MDbClientes.Contac AS Contacto, 
                      dbo.MDbClientes.ContacEnt AS Contacto_Entrega, dbo.MDbClientes.DireccEnt AS Direccion_Entrega, dbo.MDbClientes.DistriEnt AS Distrito_Entrega, 
                      dbo.MDbClientes.TelefoEnt AS Telefono_Entrega, dbo.MDbClientes.Recarg AS Recargo_Cliente, ISNULL(dbo.Int_TipDoc.TipDoc, '') AS Comprobante, 
                      ISNULL(dbo.Int_Moneda.Moneda, '') AS Moneda
FROM         dbo.MDbClientes LEFT OUTER JOIN
                      dbo.IntTerritorios ON dbo.MDbClientes.Territ = dbo.IntTerritorios.Item LEFT OUTER JOIN
                      dbo.Int_TipDoc ON dbo.MDbClientes.TipCom = dbo.Int_TipDoc.CodIte LEFT OUTER JOIN
                      dbo.Int_Moneda ON dbo.MDbClientes.Divisa = dbo.Int_Moneda.CodEst
GO
CREATE PROCEDURE [dbo].[IntSaveNewComentRP]
	@Codigo numeric(18,0),
	@Comentario varchar(500)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY

	declare @Exist tinyint
	Set @Exist=(Select Count(Item) FROM IntComentaryRP WITH(UPDLOCK,HOLDLOCK) WHERE(CodCli=@Codigo))
	If @Exist=0
		BEGIN
			Insert Into IntComentaryRP (CodCli,Coment)
				Select 
					Codi=@Codigo,
					Com=@Comentario
		END
	Else
		BEGIN
			Update IntComentaryRP Set Coment=@Comentario Where(CodCli=@Codigo)
		END

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE [dbo].[AlterGuiaRP]
	@Guia numeric(18,0),
	@Cliente nvarchar(300),
	@Codigo numeric(18,0),
	@PesoV money,
	@PesoF money,
	@Origen nvarchar(50),
	@Producto nvarchar(10),
	@Divisa  nvarchar(10),
	@SubTot money,
	@PorsRc money,
	@Rc money,
	@Igv money,
	@Total money,
	@Infor nvarchar(max),
 @Registro int=NULL
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY

 IF @Registro IS NULL BEGIN
 IF(SELECT COUNT(*)FROM dbo.MdbDiarRp WHERE NroGui=@Guia)<>1 RAISERROR('Seleccione el registro RP por su identificador; guía inexistente o ambigua.',16,1);
 SELECT @Registro=Item FROM dbo.MdbDiarRp WHERE NroGui=@Guia;
 END;
 IF NOT EXISTS(SELECT 1 FROM dbo.MdbDiarRp WHERE Item=@Registro AND NroGui=@Guia) RAISERROR('Registro RP y guía no coinciden.',16,1);
	Declare @Prod int, @Div int
	Set @Prod=ISNULL((SELECT Item FROM IntProdGuias WHERE Product=@Producto),(SELECT Item FROM IntProdGuias WHERE Product='NNE'))
	Set @Div=(Select CodEst From Int_Moneda Where(Moneda=@Divisa))
	Update MdbDiarRp Set Client=@Cliente,CodCli=@Codigo,PesFis=@PesoF,PesVol=@PesoV,Origen=@Origen,Produc=@Prod,
			Divisa=@Div,Neto=@SubTot,MonRc=@Rc,RcPor=@PorsRc,MonIGV=@Igv,Total=@Total,Info=@Infor WHERE Item=@Registro

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE dbo.ObtenerAccesosUsuario @IdUsuario int AS
BEGIN
 SET NOCOUNT ON;
 SELECT m.CodMod,CONVERT(tinyint,CASE WHEN m.Vigente=0 THEN 0 WHEN u.Access=N'ADMINISTRADOR'THEN 1 WHEN u.IdPerfilAcceso IS NULL THEN ISNULL(a.Access,0)ELSE ISNULL(p.Acceso,0)END)AS Access
 FROM dbo.Int_Usuario u CROSS JOIN dbo.Int_ModulosUsuario m
 LEFT JOIN dbo.Int_Access a ON a.IdUsuario=u.Item AND a.CodMod=m.CodMod
 LEFT JOIN dbo.PermisosPerfil p ON p.IdPerfil=u.IdPerfilAcceso AND p.IdModulo=m.Item
 WHERE u.Item=@IdUsuario
 UNION ALL SELECT a.CodMod,a.Access FROM dbo.Int_Access a WHERE a.IdUsuario=@IdUsuario AND NOT EXISTS(SELECT 1 FROM dbo.Int_ModulosUsuario m WHERE m.CodMod=a.CodMod);
END
GO
CREATE VIEW dbo.ViewHistory
AS
SELECT     dbo.Int_TipDoc.TipDoc AS Tipo, dbo.MDBHistory.NumDoc AS Documento, dbo.Int_StatusFac.Status, dbo.MDBHistory.FecEmi AS Emision, 
                      dbo.MdbUsers.Usuari AS Usuario, dbo.MDBHistory.Active
FROM         dbo.MDBHistory LEFT OUTER JOIN
                      dbo.Int_TipDoc ON dbo.MDBHistory.TipDoc = dbo.Int_TipDoc.CodIte LEFT OUTER JOIN
                      dbo.Int_StatusFac ON dbo.MDBHistory.Status = dbo.Int_StatusFac.ITEM LEFT OUTER JOIN
                      dbo.MdbUsers ON dbo.MDBHistory.Usuari = dbo.MdbUsers.ITEM
GO
CREATE VIEW dbo.DocumentosGestion
AS
SELECT t.TipDoc Tipo,d.Item IdDocumento,d.NumDoc Documento,d.FecEmi Emision,d.FecVen Vencimiento,
 d.CodCli CodigoCliente,d.NomCli Cliente,d.IdCliente,
 COALESCE(m.Moneda,d.DivFac) Divisa,d.SubFac Subtotal,d.RecFac Recargo,d.IgvFac IGV,d.ImpFac Importe,d.ImpSal Saldo,
 d.SitDoc Situacion,CONVERT(bit,CASE WHEN UPPER(LTRIM(RTRIM(d.SitDoc))) IN('ANL','ANU','ANULADO','ANULADA','3') THEN 1 ELSE 0 END) Anulado,
 d.IdEstado,d.FecEst FechaEstado,ISNULL(s.Status,N'Sin gestión') Estado,d.Courier,
 ref.NumDoc Referencia,d.VersionFila,
 ISNULL(c.ContacEnt,N'') ContactoEntrega,ISNULL(c.DireccEnt,N'') DireccionEntrega,
 ISNULL(di.Nombre,N'') DistritoEntrega,ISNULL(c.TelefoEnt,N'') TelefonoEntrega
FROM dbo.DBEmiFac d JOIN dbo.Int_TipDoc t ON t.CodIte=d.TipDoc
LEFT JOIN dbo.Int_Moneda m ON CONVERT(nvarchar(50),m.CodEst)=d.DivFac
LEFT JOIN dbo.Int_StatusFac s ON s.ITEM=d.IdEstado
LEFT JOIN dbo.DbClientes c ON c.Item=d.IdCliente LEFT JOIN dbo.Distritos di ON di.Item=c.IdDistritoEntrega
LEFT JOIN dbo.DBEmiFac ref ON ref.Item=d.IdDocumentoReferencia
UNION ALL
SELECT N'NC',n.ITEM,n.NumDoc,n.FecEmi,n.FecVen,n.CodCli,ISNULL(n.NomCli,c.CLIENT),n.IdCliente,
 COALESCE(m.Moneda,n.DivMon),n.Tarifa,n.Recargo,n.Igv,ISNULL(n.Monto,0),NULL,n.SitDoc,
 CONVERT(bit,CASE WHEN UPPER(LTRIM(RTRIM(n.SitDoc))) IN('ANL','ANU','ANULADO','ANULADA','3') THEN 1 ELSE 0 END),
 CONVERT(int,n.Estado),n.FecEst,ISNULL(s.Status,N'Sin gestión'),NULL,COALESCE(ref.NumDoc,n.NumRef),n.VersionFila,
 ISNULL(c.ContacEnt,N''),ISNULL(c.DireccEnt,N''),ISNULL(di.Nombre,N''),ISNULL(c.TelefoEnt,N'')
FROM dbo.DbNCredito n LEFT JOIN dbo.Int_StatusFac s ON s.ITEM=n.Estado
LEFT JOIN dbo.Int_Moneda m ON CONVERT(nvarchar(50),m.CodEst)=n.DivMon
LEFT JOIN dbo.DbClientes c ON c.Item=n.IdCliente LEFT JOIN dbo.Distritos di ON di.Item=c.IdDistritoEntrega
LEFT JOIN dbo.DBEmiFac ref ON ref.Item=n.IdDocumentoReferencia;
GO
CREATE PROCEDURE dbo.CobReservarNumero @Codigo varchar(3),@Numero nvarchar(10) OUTPUT
AS
BEGIN
 SET NOCOUNT ON;IF @@TRANCOUNT=0 BEGIN RAISERROR('La reserva requiere una transacción.',16,1);RETURN;END;
 DECLARE @N int;SELECT @N=Numero FROM dbo.SeriesCobranzas WITH(UPDLOCK,HOLDLOCK)WHERE Codigo=@Codigo;
 IF @N IS NULL OR @N>=9999999 BEGIN RAISERROR('Serie de cobranzas inexistente o agotada.',16,1);RETURN;END;
 UPDATE dbo.SeriesCobranzas SET Numero=Numero+1 WHERE Codigo=@Codigo;SET @Numero=@Codigo+RIGHT('0000000'+CONVERT(varchar(7),@N+1),7);
END
GO
CREATE VIEW dbo.ViewRecojos
AS
SELECT     dbo.MDbRecojos.Origen, dbo.MDbRecojos.FecEnv AS Emision, dbo.MDbRecojos.Destino, dbo.IntStaRecojos.Status, 
                      dbo.MDbRecojos.FecSta AS Fecha, dbo.MDbRecojos.Descrip AS Descripcion
FROM         dbo.MDbRecojos INNER JOIN
                      dbo.IntStaRecojos ON dbo.MDbRecojos.Status = dbo.IntStaRecojos.Item
GO
CREATE VIEW [dbo].[ViewCartas]
AS
SELECT     dbo.MdbCartas.Item, CONVERT(NVARCHAR(50), YEAR(dbo.MdbCartas.FecEmi)) + '-' + RIGHT('000000' + CONVERT(NVARCHAR(50), 
                      dbo.MdbCartas.CodCar), 6) AS Codigo, dbo.MdbCartas.FecEmi AS Emision, dbo.MdbCartas.CliCar AS Cliente, dbo.MdbCartas.CodCli, 
                      dbo.MdbCartas.DirCar AS Direccion, dbo.MdbCartas.Contac AS Contacto, dbo.IntCargosEnterprice.DesCar AS Cargo, 
                      dbo.IntReferEnterprice.DesRef AS Referencia, dbo.Int_Moneda.Moneda AS Divisa, dbo.MdbCartas.MontCar AS Monto, 
                      (CASE sgdaCar WHEN 0 THEN 'Primera' ELSE 'Segunda' END) AS Carta, ISNULL(dbo.IntStatusCart.Status, '') AS Status, 
                      dbo.MdbCartas.FecSta AS Fecha, dbo.MdbCartas.Descrip, dbo.MdbCartas.CodCar
FROM         dbo.MdbCartas INNER JOIN
                      dbo.IntCargosEnterprice ON dbo.MdbCartas.Cargo = dbo.IntCargosEnterprice.CodCar INNER JOIN
                      dbo.IntReferEnterprice ON dbo.MdbCartas.Refere = dbo.IntReferEnterprice.CodRef INNER JOIN
                      dbo.Int_Moneda ON dbo.MdbCartas.Divisa = dbo.Int_Moneda.CodEst LEFT OUTER JOIN
                      dbo.IntStatusCart ON dbo.MdbCartas.Status = dbo.IntStatusCart.Item
WHERE     (dbo.MdbCartas.Active = 1)
GO
CREATE VIEW dbo.InformeClientes AS
SELECT c.Asient Codigo,c.CLIENT Cliente,c.RUCDNI RUC,c.DIRECC Direccion,c.TELEF1 Telefono,c.Contac Contacto,t.Descrip Territorio,CASE WHEN c.EsLocal=1 THEN N'Local'ELSE N'Importado'END Procedencia,CASE WHEN c.Activo=1 THEN N'Activo'ELSE N'Baja'END Estado,c.DireccEnt DireccionEntrega,c.ContacEnt ContactoEntrega,c.TelefoEnt TelefonoEntrega,c.Recarg PorcentajeRecargo FROM dbo.DbClientes c LEFT JOIN dbo.IntTerritorios t ON t.Item=c.IdTerritorio;
GO
CREATE VIEW dbo.InformeCompilatorioRP AS
SELECT c.Asient Codigo,c.NroGui Guia,c.Client Cliente,c.PuGui FechaEnvio,c.Origen,c.TipGui Producto,c.NroPla Cartilla,c.FecPla FechaCartilla,c.FINAL Estado,c.PesFis PesoFisico,c.PesVol PesoVolumen,c.MonNet Neto,c.MonRc Recargo,c.MonIgv IGV,c.MonTot Total,c.Moneda,c.NroDoc Documento,c.NroDua DUA,CONVERT(nvarchar(1000),c.Info) Observaciones FROM dbo.DbsCompRp c WHERE NULLIF(c.FINAL,'')IS NOT NULL;
GO
CREATE PROCEDURE [dbo].[InsertServiceLSLA]
	@TipoServ nvarchar(50),
	@Cantidad nvarchar(80),
	@Tipo nvarchar(500)
	
	


AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
DECLARE @Bloqueo numeric(18,0); SELECT @Bloqueo=Correl FROM dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) WHERE DivCor IN('CLS');
IF(SELECT COUNT(*)FROM dbo.Int_Correl WHERE DivCor='CLS')<>1 RAISERROR('Correlativo inexistente o ambiguo.',16,1);IF NOT EXISTS(SELECT 1 FROM dbo.IntInfoSys WHERE ModAct='CODCLI' AND ISNUMERIC(Infor1)=1) RAISERROR('Configure el cliente de intercambio antes de generar envíos.',16,1);

	Declare @Servicio numeric(18,0),@CodClient numeric(18,0)

	Set @Servicio=(Select Corre=Correl+1 From Int_Correl Where(DivCor='CLS'))
	Set @CodClient=(Select Infor1 From IntInfoSys Where(modact='CODCLI'))
	Insert Into IntSaveActualizar (Consulta)
	Select 
		Contecto='insert into MdbServices (CodSer,NroBol,NomCli,CodCli,RucDni,DirCli,DisCli,TipSer,CanSer,Tipo)
				  Select 
					CodSer='''+convert(nvarchar(50),@Servicio)+''',
					NroBol='''+convert(nvarchar(50),@Servicio)+''',
					NomCli=(Select NomCli From MdbCostumer Where(CodCli='''+convert(nvarchar(50),@CodClient)+''')),
					CodCli='''+convert(nvarchar(50),@CodClient)+''',
					RucDni=(Select RucDni From MdbCostumer Where(CodCli='''+convert(nvarchar(50),@CodClient)+''')),
					DirCli=(Select DirCli From MdbCostumer Where(CodCli='''+convert(nvarchar(50),@CodClient)+''')),
					DisCli=(Select DisCli From MdbCostumer Where(CodCli='''+convert(nvarchar(50),@CodClient)+''')),
					TipSer=(Select CodTip From IntTipServ Where(DesTip='''+REPLACE(@TipoServ,'''','''''')+''')),
					CanSer='''+REPLACE(@Cantidad,'''','''''')+''', 
					Tipo=(Select CodPro From MdbProduct Where(DesPro='''+REPLACE(@Tipo,'''','''''')+'''))'
	
	Update Int_Correl Set Correl=Correl+1 Where(DivCor='CLS')
SELECT Servicio=@Servicio;

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE dbo.CobTerminarEnvio @Envio uniqueidentifier,@Estado nvarchar(12),@Detalle nvarchar(1000)
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 IF @Estado NOT IN('Enviado','Error','Incierto')RAISERROR('Resultado de transporte inválido.',16,1);
 BEGIN TRANSACTION;
 UPDATE dbo.EnviosCobranzas SET Estado=@Estado,Detalle=@Detalle,Fin=GETDATE()WHERE Item=@Envio AND Estado='Enviando';
 IF @@ROWCOUNT<>1 RAISERROR('El intento ya fue resuelto o no fue reclamado.',16,1);
 INSERT dbo.HistorialCobranzas(IdUsuario,Modulo,Accion,Detalle,IdRecojo,IdCarta)SELECT e.IdUsuario,CASE WHEN e.IdRecojo IS NULL THEN 15 ELSE 14 END,'CORREO',@Estado+': '+LEFT(ISNULL(@Detalle,''),980),e.IdRecojo,c.IdCarta FROM dbo.EnviosCobranzas e LEFT JOIN dbo.EmisionesCarta c ON c.Item=e.IdEmision WHERE e.Item=@Envio;
 COMMIT;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE VIEW dbo.ViewDiarRP
AS
SELECT     dbo.MdbDiarRp.Item, CONVERT(nvarchar(20), dbo.MdbDiarRp.NroGui) AS Guia, dbo.MdbDiarRp.Client AS Cliente, dbo.MdbDiarRp.CodCli AS Codigo, 
                      ISNULL(dbo.MDbClientes.Territ, '00') AS Territorio, dbo.MdbDiarRp.PesFis AS PesoFis, dbo.MdbDiarRp.PesVol AS PesoVol, dbo.MdbDiarRp.Neto, 
                      dbo.MdbDiarRp.MonRC AS Rc, dbo.MdbDiarRp.RcPor AS Porsentaje, dbo.MdbDiarRp.MonIGV AS IGV, dbo.MdbDiarRp.Total, 
                      dbo.Int_Moneda.Moneda AS Divisa, dbo.MdbDiarRp.PuGui AS Fecha, dbo.MdbDiarRp.Origen, dbo.IntProdGuias.Product AS Producto, 
                      ISNULL(dbo.Int_StatusRP.Status, '') AS Status, CASE dbo.MdbDiarRp.Salida WHEN 1 THEN 'SI' ELSE 'NO' END AS Salida, 
                      dbo.MdbDiarRp.Info AS Informacion, dbo.MdbDiarRp.Reten AS Retencion, dbo.MdbDiarRp.Servic AS Servicios
FROM         dbo.MdbDiarRp INNER JOIN
                      dbo.Int_Moneda ON dbo.MdbDiarRp.Divisa = dbo.Int_Moneda.CodEst INNER JOIN
                      dbo.IntProdGuias ON dbo.MdbDiarRp.Produc = dbo.IntProdGuias.Item LEFT OUTER JOIN
                      dbo.MDbClientes ON dbo.MdbDiarRp.CodCli = dbo.MDbClientes.Asient LEFT OUTER JOIN
                      dbo.Int_StatusRP ON dbo.MdbDiarRp.Status = dbo.Int_StatusRP.CodEst
GO
CREATE PROCEDURE dbo.EliminarCliente @Item int,@Version binary(8)
AS
BEGIN
 SET NOCOUNT ON;
 UPDATE dbo.DbClientes SET Activo=0 WHERE Item=@Item AND Activo=1 AND VersionFila=@Version AND EsLocal=1;
 IF @@ROWCOUNT<>1 RAISERROR('Solo puede dar de baja un cliente local activo sin cambios concurrentes.',16,1);
END
GO
CREATE FUNCTION dbo.OrigenesInforme()
RETURNS TABLE
AS RETURN(
 SELECT 1 Origen,N'Documentos emitidos' Nombre,N'InformeDocumentos' Vista,11 Permiso
 UNION ALL SELECT 2,N'Clientes',N'InformeClientes',20
 UNION ALL SELECT 3,N'Guías de importación',N'InformeGuias',9
 UNION ALL SELECT 4,N'Diario RP',N'InformeDiarioRP',9
 UNION ALL SELECT 5,N'Compilatorio RP',N'InformeCompilatorioRP',10
 UNION ALL SELECT 6,N'Queries',N'InformeQueries',6
 UNION ALL SELECT 7,N'Casos de quejas',N'InformeQuejas',12
 UNION ALL SELECT 8,N'Cheques rechazados',N'InformeCheques',13
 UNION ALL SELECT 9,N'Órdenes de recojo',N'InformeRecojos',14
 UNION ALL SELECT 10,N'Cartas de cobranza',N'InformeCartas',15
 UNION ALL SELECT 11,N'Devoluciones',N'InformeDevoluciones',16
);
GO
CREATE PROCEDURE [dbo].[ProSaveDevol]
	@Usuario nvarchar(10),
	@Import nvarchar(50),
	@CodCli numeric(18,0),
	@Guia  numeric(18,0),
	@Cancelado nvarchar(150),
	@Abonos nvarchar(150),
	@Motivo nvarchar(50),
	@Factura nvarchar(150),
	@Modalid nvarchar(150),
	@Banco nvarchar(150),
	@CtaCte nvarchar(50),
	@Monto money,
	@Divisa nvarchar(10),
	@Contac nvarchar(150),
	@Cliente nvarchar(300)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
DECLARE @Bloqueo numeric(18,0); SELECT @Bloqueo=Correl FROM dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) WHERE DivCor IN('CDV');
IF(SELECT COUNT(*)FROM dbo.Int_Correl WHERE DivCor='CDV')<>1 RAISERROR('Correlativo inexistente o ambiguo.',16,1);
		Insert into MdbDevol (NroCar,FecCar,Usuari,ImpExp,Client,CodCli,NroGui,Cancel,Abonos,Motivo,Factura,Modali,Banco,CtaCte,Monto,Divisa,Contac,Status,FecSta,Descrip)
			Select top 1
				Codigo=(Select Correlativo=Correl+1 from Int_Correl Where(DivCor='CDV')),
				Emision=getdate(),
				Usuario=(Select Item From MdbUsers Where(Usuari=@Usuario)),
				Import=(Select CodMod From IntRefeModalDevol Where(DesMod=@Import)),
				Cliente=@Cliente,
				CodCli=@codcli,
				Guias=@Guia,
				Cancelad=@Cancelado,
				Abonos=@Abonos,
				Motivo=(Select CodMot From IntMotivoDevol Where(DesMot=@Motivo)),
				Factura=@Factura,
				Modalida=case @Modalid when 'TRANSFERENCIA' then 1 when 'DEPOSITO' then 2 when 'OTROS' then 3 else 0 end,
				Banco=isnull((Select CodBank From Int_Bancos Where(NomBank=@Banco)),0),
				Cuenta=@CtaCte,
				MonCan=@Monto,
				Divisa=(Select CodEst From Int_Moneda Where(Moneda=@Divisa)),
				Contacto=@Contac,
				Status=1,
				FecSta=getdate(),
				Descrip=''
			
			Update Int_Correl Set Correl=correl+1 Where(DivCor='CDV')

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE [dbo].[InsertDocumentSaveLSLA]
	@Documento nvarchar(50),
	@Producto nvarchar(80),
	@Empresa nvarchar(500),
	@Direccion nvarchar(500),
	@Distrito nvarchar(150),
	@Contacto nvarchar(150),
	@Informe nvarchar(150),
	@Descripcion nvarchar(150)
	

,
 @ServicioAsignado numeric(18,0)=NULL
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
DECLARE @Bloqueo numeric(18,0); SELECT @Bloqueo=Correl FROM dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) WHERE DivCor IN('CGL');
IF(SELECT COUNT(*)FROM dbo.Int_Correl WHERE DivCor='CGL')<>1 RAISERROR('Correlativo inexistente o ambiguo.',16,1);IF NOT EXISTS(SELECT 1 FROM dbo.IntInfoSys WHERE ModAct='CODCLI' AND ISNUMERIC(Infor1)=1) RAISERROR('Configure el cliente de intercambio antes de generar envíos.',16,1);
IF @ServicioAsignado IS NULL OR @ServicioAsignado<=0 RAISERROR('Indique el servicio devuelto por InsertServiceLSLA.',16,1);

	Declare @Emision datetime,@Codigo numeric(18,0),@Servicio numeric(18,0),@CodClient numeric(18,0)
	set @Emision=GetDate()
	Set @Codigo=(Select Corre=Correl+1 From Int_Correl Where(DivCor='CGL'))
	SET @Servicio=@ServicioAsignado
	Set @CodClient=(Select Infor1 From IntInfoSys Where(modact='CODCLI'))
	--print @CodClient
	Insert Into IntSaveActualizar (Consulta)
	Select 
		Contecto='insert into MdbDocument (FecEmi,NroCod,CodServ,Docmto,CoProd,CodCli,NomEmp,DirEmp,DisEmp,ConEmp,DesEmp,Inform,Active)
				  Select 
					FecEmi='''+CONVERT(nvarchar(23),@Emision,126)+''',
					NroCod='''+convert(nvarchar(50),@Codigo)+''',
					CoServ='''+convert(nvarchar(50),@Servicio)+''',
					Docmto='''+REPLACE(@Documento,'''','''''')+''',
					CoProd=(Select CodPro From MdbProduct Where(DesPro='''+REPLACE(@Producto,'''','''''')+''')),
					CodCli='''+convert(nvarchar(50),@CodClient)+''',
					NomEmp='''+REPLACE(@Empresa,'''','''''')+''',
					DirEmp='''+REPLACE(@Direccion,'''','''''')+''',
					DisEmp=(Select CodDist From MdbDistritos Where(NomDis='''+REPLACE(@Distrito,'''','''''')+''')),
					ConEmp='''+REPLACE(@Contacto,'''','''''')+''',
					DesEmp='''+REPLACE(@Descripcion,'''','''''')+''',
					Inform='''+REPLACE(@Informe,'''','''''')+''',
					Active=0'
	
	Update Int_Correl Set Correl=Correl+1 Where(DivCor='CGL')

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.ViewPrealertRpAct
AS
SELECT     dbo.IntPrealertRP.NroGui AS Guia, dbo.MdbDiarRp.Item AS Remplazo, dbo.IntPrealertRP.Cliente, dbo.IntPrealertRP.CodCli AS Codigo, 
                      dbo.IntPrealertRP.Item
FROM         dbo.IntPrealertRP INNER JOIN
                      dbo.MdbDiarRp ON dbo.IntPrealertRP.NroGui = dbo.MdbDiarRp.NroGui
WHERE     (dbo.IntPrealertRP.Estado = 1)
GO
CREATE PROCEDURE dbo.GuardarUsuario @Actor int,@Accion tinyint,@Item int=0,@Version binary(8)=NULL,@Usuario nvarchar(50)=NULL,@Nombre nvarchar(50)=NULL,@Apellido nvarchar(50)=NULL,@Correo nvarchar(100)=NULL,@Perfil tinyint=1,@Clave varchar(32)=NULL,@Permisos varchar(40)=NULL,@IdPerfilAcceso int=NULL,@VersionPerfil binary(8)=NULL AS
BEGIN
 SET NOCOUNT ON;SET XACT_ABORT ON;
 IF @@TRANCOUNT<>0 BEGIN RAISERROR('Se requiere una operación de usuario independiente.',16,1);RETURN;END;
 BEGIN TRY
 BEGIN TRAN;
 -- El bloqueo común protege también las bajas o degradaciones simultáneas de administradores.
 DECLARE @PerfilOriginal int;
 DECLARE @Administradores int,@UsuarioOriginal nvarchar(50),@EraAdministrador bit,@Numero int,@Codigo char(6),@Acceso tinyint;
 SELECT @Administradores=COUNT(*)FROM dbo.Int_Usuario WITH(UPDLOCK,HOLDLOCK)WHERE Access=N'ADMINISTRADOR'AND ActPas=1 AND SitUsu=N'Activo';
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE Item=@Actor AND Access=N'ADMINISTRADOR'AND ActPas=1 AND SitUsu=N'Activo')RAISERROR('Solo un administrador activo puede gestionar usuarios.',16,1);
 IF @Accion IS NULL OR @Accion NOT BETWEEN 1 AND 4 RAISERROR('Operación de usuario inválida.',16,1);
 IF @Item IS NULL OR @Item<0 RAISERROR('Identificador de cuenta inválido.',16,1);
 IF @Accion=1 AND @Item<>0 RAISERROR('Alta con identificador inválido.',16,1);
 IF @Accion<>1 BEGIN
  SELECT @PerfilOriginal=IdPerfilAcceso,@UsuarioOriginal=Usuari,@EraAdministrador=CASE WHEN Access=N'ADMINISTRADOR'THEN 1 ELSE 0 END FROM dbo.Int_Usuario WHERE Item=@Item AND ActPas=1 AND SitUsu=N'Activo'AND VersionFila=@Version;
  IF @UsuarioOriginal IS NULL RAISERROR('El usuario cambió o fue eliminado. Recargue.',16,1);
  IF @Item=@Actor AND(@Accion=4 OR(@Accion=2 AND @Perfil<>0))RAISERROR('No puede dar de baja su cuenta ni quitarse la administración durante su sesión.',16,1);
  IF @EraAdministrador=1 AND(@Accion=4 OR(@Accion=2 AND @Perfil<>0))AND @Administradores<=1 RAISERROR('Debe conservar al menos un administrador activo.',16,1);
 END;
 IF @Accion IN(1,2)BEGIN
  SET @Nombre=LTRIM(RTRIM(@Nombre));SET @Apellido=LTRIM(RTRIM(@Apellido));SET @Correo=LTRIM(RTRIM(ISNULL(@Correo,N'')));
  IF NULLIF(@Nombre,N'')IS NULL OR NULLIF(@Apellido,N'')IS NULL OR @Perfil IS NULL OR @Perfil NOT IN(0,1)RAISERROR('Complete nombres, apellidos y perfil.',16,1);
  IF @Perfil=0 BEGIN SET @Permisos=REPLICATE('1',40);SET @IdPerfilAcceso=NULL;END;
  IF @Perfil=1 AND @IdPerfilAcceso IS NOT NULL BEGIN
   IF NOT EXISTS(SELECT 1 FROM dbo.PerfilesAcceso WITH(UPDLOCK,HOLDLOCK)WHERE Item=@IdPerfilAcceso AND(Activo=1 OR Item=@PerfilOriginal)AND VersionFila=@VersionPerfil)RAISERROR('El perfil no está disponible o cambió. Recargue.',16,1);
   SET @Permisos='';DECLARE @N int;SET @N=1;
   WHILE @N<=40 BEGIN
    SET @Permisos=@Permisos+CONVERT(char(1),ISNULL((SELECT CONVERT(int,Acceso)FROM dbo.PermisosPerfil WHERE IdPerfil=@IdPerfilAcceso AND IdModulo=@N),0));
    SET @N=@N+1;
   END;
  END;
  IF @Permisos IS NULL OR LEN(@Permisos)<>40 OR @Permisos LIKE '%[^01]%' COLLATE Latin1_General_BIN RAISERROR('El listado de permisos está incompleto.',16,1);
  -- Archivo y Salir siempre están disponibles para poder cerrar la sesión.
  SET @Permisos=STUFF(@Permisos,1,2,'11');
  IF @Perfil=1 AND(SUBSTRING(@Permisos,27,1)='1'OR SUBSTRING(@Permisos,33,1)='1'OR SUBSTRING(@Permisos,34,1)='1'OR SUBSTRING(@Permisos,35,1)='1')RAISERROR('La administración de usuarios corresponde al administrador.',16,1);
  IF EXISTS(SELECT 1 FROM dbo.Int_ModulosUsuario WHERE Vigente=1 AND IdPadre IS NOT NULL AND SUBSTRING(@Permisos,Item,1)='1'AND SUBSTRING(@Permisos,IdPadre,1)<>'1')RAISERROR('Active también el grupo del permiso seleccionado.',16,1);
 END;
 IF @Accion IN(1,3)AND(@Clave IS NULL OR DATALENGTH(@Clave)<>32 OR @Clave COLLATE Latin1_General_BIN LIKE '%[^0-9a-f]%')RAISERROR('Clave MD5 inválida.',16,1);
 IF @Accion=1 BEGIN
  SET @Usuario=UPPER(LTRIM(RTRIM(@Usuario)));
  IF @Usuario IS NULL OR LEN(@Usuario)NOT BETWEEN 1 AND 5 OR @Usuario COLLATE Latin1_General_BIN LIKE '%[^A-Z0-9]%'RAISERROR('Usuario: de una a cinco letras o números.',16,1);
  IF EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE Usuari=@Usuario)RAISERROR('El usuario ya existe, incluso si fue eliminado.',16,1);
  INSERT dbo.Int_Usuario(NomUsu,ApeUsu,Usuari,Paswor,Access,SitUsu,Email,ActPas,IdPerfilAcceso)VALUES(@Nombre,@Apellido,@Usuario,@Clave,CASE WHEN @Perfil=0 THEN N'ADMINISTRADOR'ELSE N'OPERADOR'END,N'Activo',@Correo,1,@IdPerfilAcceso);
  SET @Item=SCOPE_IDENTITY();
 END;
 IF @Accion=2 BEGIN
  IF @Usuario IS NOT NULL AND @Usuario<>@UsuarioOriginal RAISERROR('El nombre de acceso de una cuenta existente se conserva.',16,1);
  UPDATE dbo.Int_Usuario SET IdPerfilAcceso=@IdPerfilAcceso,NomUsu=@Nombre,ApeUsu=@Apellido,Email=@Correo,Access=CASE WHEN @Perfil=0 THEN N'ADMINISTRADOR'ELSE N'OPERADOR'END WHERE Item=@Item;
 END;
 IF @Accion IN(1,2)BEGIN
  IF @Accion=1 SET @UsuarioOriginal=@Usuario;
  SET @Numero=1;
  WHILE @Numero<=40 BEGIN
   SELECT @Codigo=CodMod FROM dbo.Int_ModulosUsuario WHERE Item=@Numero;
   SET @Acceso=CONVERT(tinyint,SUBSTRING(@Permisos,@Numero,1));
   UPDATE dbo.Int_Access SET Access=@Acceso WHERE IdUsuario=@Item AND CodMod=@Codigo;
   IF @@ROWCOUNT=0 INSERT dbo.Int_Access(Usuari,CodMod,Access,IdUsuario)VALUES(@UsuarioOriginal,@Codigo,@Acceso,@Item);
   SET @Numero=@Numero+1;
  END;
 END;
 IF @Accion=3 UPDATE dbo.Int_Usuario SET Paswor=@Clave WHERE Item=@Item;
 -- La baja es exclusiva de la cuenta. No borra permisos ni documentos históricos.
 IF @Accion=4 UPDATE dbo.Int_Usuario SET ActPas=0,SitUsu=N'Baja'WHERE Item=@Item;
 COMMIT;
 SELECT Item=@Item;
 END TRY
 BEGIN CATCH
  IF XACT_STATE()<>0 ROLLBACK;
  DECLARE @Mensaje nvarchar(2048);SET @Mensaje=ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);
 END CATCH
END
GO
CREATE VIEW dbo.ViewBusAvanz
AS
SELECT     'Ubicado en el diario General Asignado' AS Origen, NroGui AS Guia, Client AS Cliente
FROM         dbo.MdbDiarRp
WHERE     (CodCli <> 0)
UNION
SELECT     'Ubicado en el diario General por Asignar' AS Origen, NroGui, Client
FROM         dbo.MdbDiarRp AS MdbDiarRp_1
WHERE     (CodCli IS NULL OR CodCli = 0)
UNION
SELECT     'Ubicado en el Compilatorio de Rps' AS Origen, NroGui, Client
FROM         dbo.MdbCompRp
GO
CREATE VIEW dbo.ViewReemplace
AS
SELECT     dbo.IntReplaceRecojos.VarData AS Data, ISNULL(dbo.IntVarReplace.Variable, '') AS Remplazo, dbo.IntReplaceRecojos.Origen
FROM         dbo.IntReplaceRecojos LEFT OUTER JOIN
                      dbo.IntVarReplace ON dbo.IntReplaceRecojos.Replace = dbo.IntVarReplace.Item
GO
CREATE PROCEDURE [dbo].[InsertRecojos]
	@Origen varchar(50),
	@Mensaje varchar(2000),
	@Destino varchar(300)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
DECLARE @Bloqueo numeric(18,0); SELECT @Bloqueo=Correl FROM dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) WHERE DivCor IN('RCJ');
IF(SELECT COUNT(*)FROM dbo.Int_Correl WHERE DivCor='RCJ')<>1 RAISERROR('Correlativo inexistente o ambiguo.',16,1);
		Insert into MdbRecojos (Asient,Origen,FecEnv,Mensaj,Destino,Status,FecSta,Descrip)
			Select 
				Asiento=(Select Correl=Correl+1 From Int_Correl Where(DivCor='RCJ')),
				Origen=@Origen,
				FecEnv=getdate(),
				Mensaj=@Mensaje,
				Destino=@Destino,
				Status=(Select Item From IntStaRecojos Where(Status='Enviado')),
				FecSta=getdate(),
				Descrip=''
				
		Update Int_Correl Set Correl=Correl+1 Where(DivCor='RCJ')

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.ViewUsuarios
AS
SELECT     dbo.MdbUsers.NomUsu AS Nombre, dbo.MdbUsers.ApeUsu AS Apellido, dbo.MdbUsers.Usuari AS Login, ISNULL(dbo.IntNiveles.Descrip, 
                      'PERSONALIZADO') AS Acceso, CASE dbo.MdbUsers.ActPas WHEN 1 THEN 'Activo' ELSE 'Inactivo' END AS ActivePas, 
                      CASE dbo.MdbUsers.SitUsu WHEN 1 THEN 'Activo' ELSE 'Inactivo' END AS ActiveUser, 
                      CASE dbo.MdbUsers.Origen WHEN 1 THEN 'Centro Aereo' WHEN 2 THEN 'Javier Prado' ELSE 'No Especificado' END AS Origen, 
                      ISNULL(dbo.MdbUsers.USUARIO, '') AS UserMail, ISNULL(dbo.MdbUsers.PASWORD, '') AS PasMail
FROM         dbo.MdbUsers LEFT OUTER JOIN
                      dbo.IntNiveles ON dbo.MdbUsers.Access = dbo.IntNiveles.item
GO
CREATE VIEW dbo.ViewDatOrigen
AS
SELECT     dbo.IntMailOrigen.Origen, ISNULL(dbo.IntInfoMail.Grupos, '') AS Grupos, ISNULL(dbo.IntInfoBody.Body, '') AS Body
FROM         dbo.IntMailOrigen LEFT OUTER JOIN
                      dbo.IntInfoMail ON dbo.IntMailOrigen.Grupos = dbo.IntInfoMail.Item INNER JOIN
                      dbo.IntInfoBody ON dbo.IntMailOrigen.Mensaj = dbo.IntInfoBody.Item
GO
CREATE PROCEDURE dbo.EliminarCatalogo @Catalogo tinyint,@Codigo int,@Version binary(8) AS
BEGIN
SET NOCOUNT ON;
DECLARE @Filas int;SET @Filas=0;
IF @Catalogo=1 BEGIN UPDATE dbo.IntTerritorios SET Active=0 WHERE Item=@Codigo AND VersionFila=@Version AND Active=1;SET @Filas=@@ROWCOUNT;END;
IF @Catalogo=2 BEGIN UPDATE dbo.NivelesCliente SET Activo=0 WHERE Item=@Codigo AND VersionFila=@Version AND Activo=1;SET @Filas=@@ROWCOUNT;END;
IF @Catalogo=3 BEGIN UPDATE dbo.TiposCobranza SET Activo=0 WHERE Item=@Codigo AND VersionFila=@Version AND Activo=1;SET @Filas=@@ROWCOUNT;END;
IF @Catalogo=4 BEGIN UPDATE dbo.Int_Moneda SET ActEst=0 WHERE CodEst=@Codigo AND VersionFila=@Version AND ActEst=1;SET @Filas=@@ROWCOUNT;END;
IF @Catalogo=5 BEGIN UPDATE dbo.Int_TipDoc SET Activo=0 WHERE CodIte=@Codigo AND VersionFila=@Version AND Activo=1;SET @Filas=@@ROWCOUNT;END;
IF @Filas<>1 RAISERROR('El registro cambió o ya está eliminado. Recargue.',16,1);
END
GO
CREATE VIEW dbo.InformeDevoluciones AS
SELECT d.NroCar Codigo,d.FecCar FechaRegistro,d.Client Cliente,d.ImpExp Servicio,d.Motivo,d.Transf Forma,d.CtaCte Cuenta,d.Monto,d.Moneda,d.Contac Contacto,d.EstadoGestion Estado,d.FechaCierre,d.MotivoCierre,(SELECT COUNT(*)FROM dbo.DocumentosCobranzas x WHERE x.IdDevolucion=d.Item) Documentos,(SELECT COUNT(*)FROM dbo.GuiasDevolucion g WHERE g.IdDevolucion=d.Item) Guias FROM dbo.DbDevol d;
GO
CREATE VIEW dbo.InformeCartas AS
SELECT c.Correl Codigo,c.FecCar FechaRegistro,c.CliCar Cliente,c.RefeCar Referencia,c.DivCar Moneda,c.MontCar Monto,c.ContCar Contacto,c.CargCar Cargo,c.EstadoGestion Estado,c.FechaCierre,c.MotivoCierre,(SELECT COUNT(*)FROM dbo.EmisionesCarta e WHERE e.IdCarta=c.Item) Emisiones,(SELECT COUNT(*)FROM dbo.DocumentosCobranzas d WHERE d.IdCarta=c.Item) Documentos FROM dbo.DbsCartas c;
GO
CREATE PROCEDURE [dbo].[ProSaveCarta]
	@CodCli numeric(18,0),
	@Contac nvarchar(150),
	@Cargo nvarchar(150),
	@Referen nvarchar(150),
	@Divisa nvarchar(150),
	@Monto money

AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
DECLARE @Bloqueo numeric(18,0); SELECT @Bloqueo=Correl FROM dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) WHERE DivCor IN('CPC');
IF(SELECT COUNT(*)FROM dbo.Int_Correl WHERE DivCor='CPC')<>1 RAISERROR('Correlativo inexistente o ambiguo.',16,1);
	Insert into MdbCartas (CodCar,FecEmi,CliCar,CodCli,DirCar,Contac,Cargo,Refere,Divisa,MontCar,Status)
		Select top 1
			Codigo=(Select Correlativo=Correl+1 from Int_Correl Where(DivCor='CPC')),
			Emision=getdate(),
			Cliente=(Select top 1 Client FROM MDbClientes Where(Asient=@CodCli)),
			Codigo=@codcli,
			Direccion=(Select top 1 Direcc FROM MDbClientes Where(Asient=@CodCli)),
			Contacto=@Contac,
			Cargo=(Select CodCar From IntCargosEnterprice Where(DesCar=@Cargo)),
			Referencia=(Select CodRef From IntReferEnterprice Where(DesRef=@Referen)),
			Moneda=(Select CodEst From Int_Moneda Where(Moneda=@Divisa)),
			Monto=@Monto,
			Status=1
		
	Update Int_Correl Set Correl=correl+1 Where(DivCor='CPC')

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE [dbo].[InsertClient]
	@Codigo numeric(18,0),
	@Cliente nvarchar(255),
    @Ruc nvarchar(50),
    @Direccion nvarchar(300),
    @Distrito nvarchar(50),
    @Telefono nvarchar(50),
    @Territorio nvarchar(50),
    @Contacto nvarchar(100),
    @ContancEnt nvarchar(100),
    @DireccEnt nvarchar(255),
    @DistriEnt nvarchar(50),
    @TelefonEnt nvarchar(50),
    @TasaRC Money,
    @Comprobante nvarchar(50),
    @Divisa nvarchar(10),
    @TipoClient nvarchar(50),
    @Cobranza nvarchar(20)


AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY


	Declare @BusCod tinyint
	Set @BusCod=(Select count(Asient) FROM MdbClientes WITH(UPDLOCK,HOLDLOCK) WHERE(Asient=@Codigo))
	
	If @BusCod=0
		BEGIN
			Insert Into MdbClientes (Asient,Client,RucDni,Direcc,Distri,Telef1,Telef2,Telef3,Territ,CliTnt,CliTop,Remark,CliImp,Contac,ContacEnt,DireccEnt,DistriEnt,TelefoEnt,Recarg,TipCom,Divisa)
				Select 
					Asiento=@Codigo,
					Cliente=@Cliente,
					RucDni=@Ruc,
					Direccion=@Direccion,
					Disrito=@Distrito,
					Telefono=@Telefono,
					Telefono2='',
					Telefono3='',
					Territorio=ISNULL((Select Item From IntTerritorios Where(Descrip=RIGHT('00' + replace(@Territorio,'TERRITORIO',''),2))),(Select Item From IntTerritorios Where(Descrip='00'))), 
					Conbranza =Case @Cobranza When 'TNT' THEN 1 ELSE 0 END,
					TipoClient=Case @TipoClient When 'TOP' THEN 1 ELSE 0 END,
					Remark='',
					Import='',
					Contacto=@Contacto,
					ContancTOEnt=@ContancEnt,
					DireccEnt=@DireccEnt,
					DistriEnt=@DistriEnt,
					TelefonEnt=@TelefonEnt,
					Rc=@TasaRC,
					Comprobante=Case @Comprobante WHEN 'FACTURA' THEN (Select CodIte From Int_TipDoc Where(TipDoc='FT')) 
												WHEN 'BOLETA' THEN (Select CodIte From Int_TipDoc Where(TipDoc='BV')) END,
					Divisa=(Select CodEst From Int_Moneda Where(Moneda=@Divisa))
		END
	ELSE
		BEGIN
			DECLARE @TERRIT INT,@TipoCli INT,@COBRAN INT,@COMPROBANT INT,@DIV INT
			
			SET @TERRIT=ISNULL((Select Item From IntTerritorios Where(Descrip=RIGHT('00' + replace(@Territorio,'TERRITORIO',''),2))),(Select Item From IntTerritorios Where(Descrip='00')))
			SET @TipoCli=Case @TipoClient When 'TOP' THEN 1 ELSE 0 END
			SET @COBRAN=Case @Cobranza When 'TNT' THEN 1 ELSE 0 END
			SET @COMPROBANT=Case @Comprobante WHEN 'FACTURA' THEN (Select CodIte From Int_TipDoc Where(TipDoc='FT')) 
												WHEN 'BOLETA' THEN (Select CodIte From Int_TipDoc Where(TipDoc='BV')) END
			SET @DIV=(Select CodEst From Int_Moneda Where(Moneda=@Divisa))
			
			UPDATE MDBCLIENTES SET 
				Client=@Cliente,
				RucDni=@Ruc,
				Direcc=@Direccion,
				Distri=@Distrito,
				Telef1=@Telefono,
				Territ=@TERRIT,
				CliTnt=@COBRAN,
				CliTop=@TipoCli,
				Contac=@Contacto,
				ContacEnt=@ContancEnt,
				DireccEnt=@DireccEnt,
				DistriEnt=@DistriEnt,
				TelefoEnt=@TelefonEnt,
				Recarg=@TasaRC,
				TipCom=@COMPROBANT,
				Divisa=@DIV
			WHERE(Asient=@Codigo)
				
		END
	


 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE dbo.GuardarPerfilAcceso @Actor int,@Item int=0,@Version binary(8)=NULL,@Nombre nvarchar(50)=NULL,@Permisos varchar(40)=NULL,@Eliminar bit=0 AS
BEGIN
 SET NOCOUNT ON;SET XACT_ABORT ON;
 IF @@TRANCOUNT<>0 BEGIN RAISERROR('Se requiere una operación de perfil independiente.',16,1);RETURN;END;
 BEGIN TRY
 BEGIN TRAN;
 DECLARE @Administradores int,@Numero int,@Acceso bit;
 -- Mismo orden de bloqueo que Usuarios para coordinar asignaciones y modificaciones.
 SELECT @Administradores=COUNT(*)FROM dbo.Int_Usuario WITH(UPDLOCK,HOLDLOCK)WHERE Access=N'ADMINISTRADOR'AND ActPas=1 AND SitUsu=N'Activo';
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE Item=@Actor AND Access=N'ADMINISTRADOR'AND ActPas=1 AND SitUsu=N'Activo')RAISERROR('Solo un administrador activo puede gestionar perfiles.',16,1);
 IF @Item IS NULL OR @Item<0 OR @Eliminar IS NULL RAISERROR('Perfil inválido.',16,1);
 IF @Item<>0 AND NOT EXISTS(SELECT 1 FROM dbo.PerfilesAcceso WITH(UPDLOCK,HOLDLOCK)WHERE Item=@Item AND Activo=1 AND VersionFila=@Version)RAISERROR('El perfil cambió o fue eliminado. Recargue.',16,1);
 IF @Eliminar=1 BEGIN
  IF @Item=0 RAISERROR('Seleccione un perfil existente.',16,1);
  -- Sin cambios de usuarios ni de permisos relacionados: solo baja del perfil.
  UPDATE dbo.PerfilesAcceso SET Activo=0 WHERE Item=@Item;
 END ELSE BEGIN
  SET @Nombre=LTRIM(RTRIM(@Nombre));
  IF NULLIF(@Nombre,N'')IS NULL RAISERROR('Ingrese el nombre del perfil.',16,1);
  IF UPPER(@Nombre)IN(N'PERSONALIZADO',N'ADMINISTRADOR',N'ACCESO COMPLETO')RAISERROR('Nombre reservado para tipos de acceso del sistema.',16,1);
  IF EXISTS(SELECT 1 FROM dbo.PerfilesAcceso WITH(UPDLOCK,HOLDLOCK)WHERE Nombre=@Nombre AND Item<>@Item)RAISERROR('El nombre de perfil ya existe, incluso si fue eliminado.',16,1);
  IF @Permisos IS NULL OR LEN(@Permisos)<>40 OR @Permisos LIKE '%[^01]%' COLLATE Latin1_General_BIN RAISERROR('Listado de permisos incompleto.',16,1);
  SET @Permisos=STUFF(@Permisos,1,2,'11');
  IF SUBSTRING(@Permisos,27,1)='1'OR SUBSTRING(@Permisos,33,1)='1'OR SUBSTRING(@Permisos,34,1)='1'RAISERROR('La gestión de perfiles y usuarios corresponde al administrador.',16,1);
  IF EXISTS(SELECT 1 FROM dbo.Int_ModulosUsuario WHERE Vigente=1 AND IdPadre IS NOT NULL AND SUBSTRING(@Permisos,Item,1)='1'AND SUBSTRING(@Permisos,IdPadre,1)<>'1')RAISERROR('Active también el grupo del permiso seleccionado.',16,1);
  IF @Item=0 BEGIN INSERT dbo.PerfilesAcceso(Nombre)VALUES(@Nombre);SET @Item=SCOPE_IDENTITY();END
  ELSE BEGIN
   UPDATE dbo.PerfilesAcceso SET Nombre=@Nombre WHERE Item=@Item;
   -- Una edición de usuario basada en los accesos anteriores debe recargarse.
   UPDATE dbo.Int_Usuario SET IdPerfilAcceso=IdPerfilAcceso WHERE IdPerfilAcceso=@Item;
  END;
  SET @Numero=1;
  WHILE @Numero<=40 BEGIN
   SET @Acceso=CONVERT(bit,SUBSTRING(@Permisos,@Numero,1));
   IF EXISTS(SELECT 1 FROM dbo.Int_ModulosUsuario WHERE Item=@Numero AND Vigente=0)SET @Acceso=0;
   UPDATE dbo.PermisosPerfil SET Acceso=@Acceso WHERE IdPerfil=@Item AND IdModulo=@Numero;
   IF @@ROWCOUNT=0 INSERT dbo.PermisosPerfil(IdPerfil,IdModulo,Acceso)VALUES(@Item,@Numero,@Acceso);
   SET @Numero=@Numero+1;
  END;
 END;
 COMMIT;SELECT Item=@Item;
 END TRY
 BEGIN CATCH
  IF XACT_STATE()<>0 ROLLBACK;
  DECLARE @Mensaje nvarchar(2048);SET @Mensaje=ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);
 END CATCH
END
GO
CREATE VIEW dbo.ViewMailOrigen
AS
SELECT     dbo.IntMailOrigen.Origen, dbo.IntMailOrigen.Modal, ISNULL(dbo.IntInfoMail.Destin, '') AS Destinatario, ISNULL(dbo.IntInfoBody.Asunto, '') AS Asunto, 
                      ISNULL(dbo.IntInfoBody.Mensaj, '') AS Mensaje
FROM         dbo.IntMailOrigen LEFT OUTER JOIN
                      dbo.IntInfoMail ON dbo.IntMailOrigen.Grupos = dbo.IntInfoMail.Item LEFT OUTER JOIN
                      dbo.IntInfoBody ON dbo.IntMailOrigen.Mensaj = dbo.IntInfoBody.Item
WHERE     (dbo.IntMailOrigen.Active = 1)
GO
CREATE PROCEDURE dbo.ConsultarUsuarios @Actor int,@IncluyeEliminados bit=0,@Filtro nvarchar(100)=N'',@Item int=0 AS
BEGIN
 SET NOCOUNT ON;
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE Item=@Actor AND Access=N'ADMINISTRADOR'AND ActPas=1 AND SitUsu=N'Activo')BEGIN RAISERROR('Solo un administrador activo puede gestionar usuarios.',16,1);RETURN;END;
 DECLARE @Busqueda nvarchar(400);SET @Busqueda=REPLACE(REPLACE(REPLACE(REPLACE(@Filtro,N'\',N'\\'),N'%',N'\%'),N'_',N'\_'),N'[',N'\[');
 SELECT u.Item,Usuari,NomUsu,ApeUsu,Email,Access,u.IdPerfilAcceso,p.VersionFila AS VersionPerfil,CASE WHEN Access=N'ADMINISTRADOR'THEN N'Acceso completo'ELSE ISNULL(p.Nombre,N'Personalizado')END AS PerfilAcceso,CONVERT(bit,CASE WHEN SitUsu=N'Activo'AND ActPas=1 THEN 1 ELSE 0 END)AS Activo,u.VersionFila
 FROM dbo.Int_Usuario u LEFT JOIN dbo.PerfilesAcceso p ON u.IdPerfilAcceso=p.Item WHERE(@IncluyeEliminados=1 OR(SitUsu=N'Activo'AND ActPas=1))AND(@Item=0 OR u.Item=@Item)AND(Usuari LIKE N'%'+@Busqueda+N'%' ESCAPE N'\' OR NomUsu+N' '+ApeUsu LIKE N'%'+@Busqueda+N'%' ESCAPE N'\')ORDER BY Usuari,u.Item;
END
GO
CREATE PROCEDURE dbo.RecalcularEstadoCargo @Tipo nvarchar(2),@IdDocumento int
AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Estado int,@Legado tinyint,@Fecha datetime;
 SELECT @Estado=IdEstado,@Legado=Estado,@Fecha=FechaEstado FROM dbo.EstadosInicialesCargos WHERE Tipo=@Tipo AND IdDocumento=@IdDocumento;
 SELECT TOP(1) @Estado=h.IdEstado,@Legado=CONVERT(tinyint,h.IdEstado),@Fecha=h.FecEst
 FROM dbo.DbHisFac h WHERE h.Activo=1 AND h.IdEstado IS NOT NULL AND ((@Tipo='NC' AND h.IdNotaCredito=@IdDocumento) OR(@Tipo<>'NC' AND h.IdFactura=@IdDocumento))
 ORDER BY h.FecEst DESC,h.ITEM DESC;
 IF @Tipo='NC' UPDATE dbo.DbNCredito SET Estado=COALESCE(CONVERT(tinyint,@Estado),@Legado),FecEst=@Fecha WHERE ITEM=@IdDocumento;
 ELSE UPDATE dbo.DBEmiFac SET IdEstado=@Estado,Estado=@Legado,FecEst=@Fecha WHERE Item=@IdDocumento;
END
GO
CREATE PROCEDURE [dbo].[InsertPago]
	@Documento nvarchar(30),
	@Importe Money,
	@Divisa nvarchar(10),
	@Modal nvarchar(10)

AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY

	
	Insert into MdbAsientCancel (NumDoc,Monto,Divisa,TipCam,Tipo)
	Select 
		Documento=@Documento,
		Importe=@Importe,
		Divisa=(Select CodEst From Int_Moneda Where(Moneda=@Divisa)),
		TipCam=(Select Infor1 From IntInfoSys Where(ModAct='TIPCAM')),
		Modal=(Case @Modal WHEN 'ANULACION' THEN 1
							WHEN 'CANCELACION' THEN 2 END)

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.InformeCheques AS
SELECT q.Asient Codigo,q.FecCar FechaRegistro,q.Client Cliente,q.NroRUC RUC,q.NroChe Cheque,q.FecGir FechaGiro,b.NomBank Banco,q.Divisa Moneda,q.Monto,q.Gastos,q.MonCan Total,q.EstadoGestion Estado,q.FechaCierre,q.MotivoCierre FROM dbo.DbCheques q LEFT JOIN dbo.Int_Bancos b ON b.CodBank=q.Banco;
GO
CREATE VIEW dbo.ViewUltimosStatus AS
 SELECT d.NumDoc,ISNULL(s.Status,'') AS Status,h.FecEmi AS Fecha,ISNULL(u.Usuari,'') AS Usuario
 FROM(SELECT NumDoc FROM dbo.MdbDocument UNION SELECT Numdoc FROM dbo.MdbNCredito) d
 OUTER APPLY(SELECT TOP 1 Status,FecEmi,Usuari FROM dbo.MDBHistory WHERE NumDoc=d.NumDoc AND Active=1 ORDER BY FecEmi DESC,ITEM DESC)h
 LEFT JOIN dbo.Int_StatusFac s ON h.Status=s.ITEM LEFT JOIN dbo.MdbUsers u ON h.Usuari=u.ITEM
GO
CREATE VIEW dbo.ViewCompRP
AS
SELECT     dbo.MdbCompRp.Item, dbo.MdbCompRp.NroPla AS Planilla, dbo.MdbCompRp.FecPla AS Fecha, CONVERT(nvarchar(20), dbo.MdbCompRp.NroGui) 
                      AS Guia, dbo.MdbCompRp.Client AS Cliente, dbo.MdbCompRp.CodCli AS Codigo, dbo.MdbCompRp.PesFis AS PesoFis, 
                      dbo.MdbCompRp.PesVol AS PesoVol, dbo.MdbCompRp.Neto, dbo.MdbCompRp.MonRC AS Recargo, dbo.MdbCompRp.RcPor AS Porsentaje, 
                      dbo.MdbCompRp.MonIGV AS IGV, dbo.MdbCompRp.Total, dbo.Int_Moneda.Moneda, dbo.MdbCompRp.PuGui AS PickUP, dbo.MdbCompRp.Origen, 
                      dbo.IntProdGuias.Product AS Producto, ISNULL(dbo.Int_StatusRP.Status, '') AS Status, dbo.MdbCompRp.Info AS Informacion, 
                      dbo.MdbCompRp.Servic AS Servicios, dbo.MdbCompRp.InterMen
FROM         dbo.MdbCompRp INNER JOIN
                      dbo.Int_Moneda ON dbo.MdbCompRp.Divisa = dbo.Int_Moneda.CodEst INNER JOIN
                      dbo.IntProdGuias ON dbo.MdbCompRp.Produc = dbo.IntProdGuias.Item LEFT OUTER JOIN
                      dbo.Int_StatusRP ON dbo.MdbCompRp.Status = dbo.Int_StatusRP.CodEst
GO
CREATE PROCEDURE dbo.ConsultarCatalogo @Catalogo tinyint,@IncluyeEliminados bit=0,@Filtro nvarchar(150)=N'' AS
BEGIN
SET NOCOUNT ON;
DECLARE @Busqueda nvarchar(600);
SET @Busqueda=REPLACE(REPLACE(REPLACE(REPLACE(@Filtro,N'\',N'\\'),N'%',N'\%'),N'_',N'\_'),N'[',N'\[');
IF @Catalogo=1 SELECT Item AS Codigo,Descrip AS Nombre,CONVERT(nvarchar(2),N'') AS Abreviatura,CONVERT(bit,Active) AS Activo,VersionFila FROM dbo.IntTerritorios WHERE(@IncluyeEliminados=1 OR Active=1)AND Descrip LIKE N'%'+@Busqueda+N'%' ESCAPE N'\' ORDER BY Descrip,Item;
IF @Catalogo=2 SELECT Item AS Codigo,Nombre AS Nombre,CONVERT(nvarchar(2),N'') AS Abreviatura,CONVERT(bit,Activo) AS Activo,VersionFila FROM dbo.NivelesCliente WHERE(@IncluyeEliminados=1 OR Activo=1)AND Nombre LIKE N'%'+@Busqueda+N'%' ESCAPE N'\' ORDER BY Nombre,Item;
IF @Catalogo=3 SELECT Item AS Codigo,Nombre AS Nombre,CONVERT(nvarchar(2),N'') AS Abreviatura,CONVERT(bit,Activo) AS Activo,VersionFila FROM dbo.TiposCobranza WHERE(@IncluyeEliminados=1 OR Activo=1)AND Nombre LIKE N'%'+@Busqueda+N'%' ESCAPE N'\' ORDER BY Nombre,Item;
IF @Catalogo=4 SELECT CodEst AS Codigo,Moneda AS Nombre,CONVERT(nvarchar(2),N'') AS Abreviatura,CONVERT(bit,ActEst) AS Activo,VersionFila FROM dbo.Int_Moneda WHERE(@IncluyeEliminados=1 OR ActEst=1)AND Moneda LIKE N'%'+@Busqueda+N'%' ESCAPE N'\' ORDER BY Moneda,CodEst;
IF @Catalogo=5 SELECT CodIte AS Codigo,NamVar AS Nombre,TipDoc AS Abreviatura,CONVERT(bit,Activo) AS Activo,VersionFila FROM dbo.Int_TipDoc WHERE(@IncluyeEliminados=1 OR Activo=1)AND NamVar LIKE N'%'+@Busqueda+N'%' ESCAPE N'\' ORDER BY NamVar,CodIte;
IF @Catalogo NOT BETWEEN 1 AND 5 OR @Catalogo IS NULL RAISERROR('Catálogo inválido.',16,1);
END
GO
CREATE VIEW dbo.InformeQueries AS
SELECT q.Asient Codigo,q.FecSol FechaSolicitud,q.CliDoc Cliente,q.NumDoc Documento,q.TipDoc TipoDocumento,q.TipQue TipoSolicitud,q.TipCas Caso,q.MonSol ImporteSolicitado,q.DivSol Moneda,q.Estado,q.SitDoc Situacion,q.FechaCierre,q.Resolucion,q.MontoAplicado FROM dbo.DbQuerie q;
GO
CREATE PROCEDURE [dbo].[ProdDepur_DiarComp]
	@Item numeric(18,0)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY


 IF NOT EXISTS(SELECT 1 FROM dbo.MdbDiarRp WITH(UPDLOCK,HOLDLOCK) WHERE Item=@Item) RAISERROR('Registro RP inexistente o ya transferido.',16,1);
 DECLARE @Planilla numeric(18,0),@Compilado int;
 SET @Planilla=0;
 INSERT dbo.MdbCompRp(NroPla,NroGui,Client,CodCli,PesFis,PesVol,Neto,MonRc,RcPor,MonIgv,Total,Divisa,PuGui,Origen,Produc,Status,Info,Servic,InterMen,IdGuia)
 SELECT @Planilla,NroGui,Client,CodCli,PesFis,PesVol,Neto,MonRc,RcPor,MonIgv,Total,Divisa,PuGui,Origen,Produc,Status,Info,Servic,N'FINALIZADO POR DEPURACION',IdGuia FROM dbo.MdbDiarRp WHERE Item=@Item;
 SET @Compilado=SCOPE_IDENTITY();
 UPDATE dbo.MdbHisRP SET IdDiario=NULL,IdCompilado=@Compilado WHERE IdDiario=@Item;
 DELETE dbo.MdbDiarRp WHERE Item=@Item;

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE [dbo].[AlterOrigenMail]
	@Origen varchar(50),
	@Grupos varchar(50),
	@Body varchar(50),
	@Modal tinyint
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY

	If @Modal=1
		begin
			Insert into IntMailOrigen (Origen,Modal,Grupos,Mensaj)
				Select 
					Origen=@Origen,
					MOdal=2,
					Grupos=(Select Item From IntInfoMail Where(Grupos=@Grupos)),
					Mensaje=(Select Item From IntInfoBody Where(Body=@Body))
		end
	Else
		begin
			Declare @VarGru int,  @VarMen int
				Set @VarGru=(Select Item From IntInfoMail Where(Grupos=@Grupos))
				Set @VarMen=(Select Item From IntInfoBody Where(Body=@Body))
			Update IntMailOrigen Set Grupos=@VarGru,Mensaj=@VarMen Where(Origen=@Origen)
		end

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE dbo.ConsultarPerfilesAcceso @Actor int,@IncluyeEliminados bit=0,@Filtro nvarchar(50)=N'',@Item int=0 AS
BEGIN
 SET NOCOUNT ON;
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE Item=@Actor AND Access=N'ADMINISTRADOR'AND ActPas=1 AND SitUsu=N'Activo')BEGIN RAISERROR('Solo un administrador activo puede gestionar perfiles.',16,1);RETURN;END;
 DECLARE @Busqueda nvarchar(200);SET @Busqueda=REPLACE(REPLACE(REPLACE(REPLACE(@Filtro,N'\',N'\\'),N'%',N'\%'),N'_',N'\_'),N'[',N'\[');
 SELECT p.Item,p.Nombre,p.Activo,p.VersionFila,(SELECT COUNT(*)FROM dbo.Int_Usuario u WHERE u.IdPerfilAcceso=p.Item)AS UsuariosVinculados
 FROM dbo.PerfilesAcceso p WHERE(@IncluyeEliminados=1 OR p.Activo=1)AND(@Item=0 OR p.Item=@Item)AND p.Nombre LIKE N'%'+@Busqueda+N'%' ESCAPE N'\'ORDER BY p.Nombre,p.Item;
END
GO
CREATE VIEW dbo.ViewInfoTer
AS
SELECT     'Territorio ' + dbo.IntTerritorios.Descrip AS Territorio, dbo.IntInfoTerMail.DirTer AS Direccion, dbo.IntInfoTerMail.Item
FROM         dbo.IntInfoTerMail INNER JOIN
                      dbo.IntTerritorios ON dbo.IntInfoTerMail.Territ = dbo.IntTerritorios.Item
WHERE     (dbo.IntInfoTerMail.Active = 1)
GO
CREATE PROCEDURE [dbo].[InsertHistoryRP]
	@Item numeric(18,0),
	@Status varchar(10),
	@Info varchar(1000)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY


 DECLARE @EstadoId tinyint;
 SELECT @EstadoId=CodEst FROM dbo.Int_StatusRP WHERE Status=@Status;
 IF @EstadoId IS NULL RAISERROR('Estado RP inexistente.',16,1);
 UPDATE dbo.MdbDiarRp SET Status=@EstadoId WHERE Item=@Item;
 IF @@ROWCOUNT<>1 RAISERROR('Seleccione un registro del diario RP.',16,1);
 INSERT dbo.MdbHisRP(Herencia,Status,Info,Fecha,IdDiario) VALUES(@Item,@EstadoId,@Info,GETDATE(),@Item);

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.InformeQuejas AS
SELECT q.Asient Codigo,c.CLIENT Cliente,c.RUCDNI RUC,q.DescripcionCaso Caso,q.FecCas FechaApertura,q.FecFin FechaCierre,CASE WHEN q.FecFin IS NULL THEN N'Abierto'ELSE N'Cerrado'END Estado,g.CodigoEstado UltimoEstado,g.FecSta FechaUltimaGestion,(SELECT COUNT(*)FROM dbo.DbQuejas_Ges x WHERE x.IdCaso=q.Item) Gestiones FROM dbo.DbQuejas_Cas q LEFT JOIN dbo.DbClientes c ON c.Item=q.IdCliente OUTER APPLY(SELECT TOP(1)e.CodigoEstado,e.FecSta FROM dbo.DbQuejas_Ges e WHERE e.IdCaso=q.Item ORDER BY e.Item DESC)g;
GO
CREATE VIEW [dbo].[ViewCheques]
AS
SELECT     dbo.MdbCheques.Item, CONVERT(NVARCHAR(50), YEAR(dbo.MdbCheques.FecCar)) + '-' + RIGHT('000000' + CONVERT(NVARCHAR(50), 
                      dbo.MdbCheques.CodChe), 6) AS Codigo, dbo.MdbCheques.FecCar AS Emision, dbo.MdbCheques.Client AS Cliente, 
                      dbo.MdbCheques.RucCli AS RucDni, dbo.MdbCheques.DirCli AS Direccion, dbo.MdbCheques.Contac AS Contacto, 
                      dbo.IntCargosEnterprice.DesCar AS Cargo, dbo.MdbCheques.CodCli, dbo.Int_Bancos.NomBank AS Banco, dbo.MdbCheques.FecGir AS FechaGiro, 
                      dbo.MdbCheques.NroChe AS Cheque, dbo.Int_EstadCheq.Estado, dbo.Int_Moneda.Moneda, dbo.MdbCheques.Gastos, 
                      dbo.MdbCheques.MonCan AS Monto, dbo.MdbCheques.Gastos + dbo.MdbCheques.MonCan AS Total, ISNULL(dbo.IntStatusCart.Status, '') AS Status, 
                      dbo.MdbCheques.FecSta AS Fecha, dbo.MdbCheques.Descrip AS Descripcion, dbo.MdbCheques.CodChe
FROM         dbo.IntStatusCart INNER JOIN
                      dbo.Int_Bancos INNER JOIN
                      dbo.MdbCheques ON dbo.Int_Bancos.CodBank = dbo.MdbCheques.Banco INNER JOIN
                      dbo.Int_EstadCheq ON dbo.MdbCheques.Motivo = dbo.Int_EstadCheq.CodEst ON dbo.IntStatusCart.Item = dbo.MdbCheques.Status INNER JOIN
                      dbo.IntCargosEnterprice ON dbo.MdbCheques.Cargo = dbo.IntCargosEnterprice.CodCar INNER JOIN
                      dbo.Int_Moneda ON dbo.MdbCheques.Divisa = dbo.Int_Moneda.CodEst
GO
CREATE PROCEDURE [dbo].[InsertStatusFac]
	@Tipo nvarchar(10),
	@Documento nvarchar(30),
	@Status nvarchar(150),
	@Usuario nvarchar(10),
	@Correl Numeric(18,0)

AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
IF NOT EXISTS(SELECT 1 FROM dbo.MdbDocument WHERE NumDoc=@Documento) AND NOT EXISTS(SELECT 1 FROM dbo.MdbNCredito WHERE Numdoc=@Documento) RAISERROR('Documento inexistente.',16,1);
IF NOT EXISTS(SELECT 1 FROM dbo.Int_StatusFac WHERE Status=@Status) OR NOT EXISTS(SELECT 1 FROM dbo.Int_TipDoc WHERE TipDoc=@Tipo) OR NOT EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE Usuari=@Usuario AND SitUsu=1) RAISERROR('Revise tipo, estado y usuario del evento.',16,1);

	
	Insert into MDBHistory (TipDoc,NumDoc,Status,FecEmi,Usuari,Correl,IdFactura,IdNotaCredito)
		Select
			Tipo=(Select codIte From Int_TipDoc Where(TipDoc=@Tipo)),
			Documento=@Documento,
			Status=(Select Item From Int_StatusFac Where(Status=@Status)),
			FecEmi=getdate(),
			Usuari=(Select Item From MdbUsers Where(Usuari=@Usuario)),
			Correl=@Correl,
 IdFactura=(SELECT Item FROM MdbDocument WHERE NumDoc=@Documento),
 IdNotaCredito=(SELECT Item FROM MdbNCredito WHERE Numdoc=@Documento)

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE dbo.ValidarActorInforme @Actor int,@Origen int=0
AS
BEGIN SET NOCOUNT ON;BEGIN TRY
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE Item=@Actor AND ActPas=1 AND SitUsu=N'Activo')RAISERROR('Usuario inactivo.',16,1);
 CREATE TABLE #P(CodMod nvarchar(6),Acceso int);INSERT #P EXEC dbo.ObtenerAccesosUsuario @Actor;
 IF NOT EXISTS(SELECT 1 FROM #P p JOIN dbo.Int_ModulosUsuario m ON m.CodMod=p.CodMod WHERE m.Item=40 AND m.Vigente=1 AND p.Acceso=1)RAISERROR('No tiene permiso para Informes personalizados.',16,1);
 DECLARE @Modulo int;
 IF @Origen<>0 BEGIN
  SELECT @Modulo=Permiso FROM dbo.OrigenesInforme()WHERE Origen=@Origen;
  IF @Modulo IS NULL RAISERROR('Origen de informe inválido.',16,1);
  WHILE @Modulo IS NOT NULL BEGIN
   IF NOT EXISTS(SELECT 1 FROM #P p JOIN dbo.Int_ModulosUsuario m ON m.CodMod=p.CodMod WHERE m.Item=@Modulo AND m.Vigente=1 AND p.Acceso=1)RAISERROR('No tiene permiso para consultar este origen.',16,1);
   SELECT @Modulo=IdPadre FROM dbo.Int_ModulosUsuario WHERE Item=@Modulo;
  END;
 END;RETURN 0;
 END TRY BEGIN CATCH DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);RETURN 1;END CATCH;
END
GO
CREATE PROCEDURE [dbo].[InsertGuias]
	@Guia Numeric(18,0),
	@Origen nvarchar(10),
	@Fecha datetime,
	@Producto nvarchar(10),
	@Cliente nvarchar(300),
	@Tipo nvarchar(5)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY


 IF @Guia IS NULL OR @Guia<=0 OR @Fecha IS NULL RAISERROR('Guía y fecha obligatorias.',16,1);
 DECLARE @ProductoId int,@GuiaId int,@Registro numeric(18,0);
 SELECT @ProductoId=Item FROM dbo.IntProdGuias WHERE Product=@Producto;
 IF @ProductoId IS NULL SELECT @ProductoId=Item FROM dbo.IntProdGuias WHERE Product='NNE';
 IF @ProductoId IS NULL RAISERROR('Producto inexistente.',16,1);
 SELECT @GuiaId=Item FROM dbo.MdbGuias WITH(UPDLOCK,HOLDLOCK) WHERE Guia=@Guia AND Origen=@Origen AND Fecha=@Fecha AND Produc=@ProductoId AND Client=@Cliente AND Tipo=@Tipo;
 IF @GuiaId IS NULL BEGIN
  INSERT dbo.MdbGuias(Guia,Origen,Fecha,Produc,Client,Tipo)VALUES(@Guia,@Origen,@Fecha,@ProductoId,@Cliente,@Tipo);
  SET @GuiaId=SCOPE_IDENTITY();
 END;
 IF @Tipo='R' AND NOT EXISTS(SELECT 1 FROM dbo.MdbDiarRp WITH(UPDLOCK,HOLDLOCK) WHERE IdGuia=@GuiaId) AND NOT EXISTS(SELECT 1 FROM dbo.MdbCompRp WITH(UPDLOCK,HOLDLOCK) WHERE IdGuia=@GuiaId)
 BEGIN
  EXEC dbo.ReservarCorrelativo 'CGP',@Registro OUTPUT;
  INSERT dbo.MdbDiarRp(Item,NroGui,Client,PuGui,Origen,Produc,Salida,Divisa,RcPor,IdGuia)
  SELECT @Registro,@Guia,@Cliente,@Fecha,ISNULL((SELECT Origen FROM dbo.int_IATA WHERE codigo=@Origen),@Origen),@ProductoId,CASE WHEN @Producto IN('81','82','15D')THEN 1 ELSE 0 END,(SELECT CodEst FROM dbo.Int_Moneda WHERE Moneda='S/.'),ISNULL((SELECT Infor1 FROM dbo.IntInfoSys WHERE ModAct='RCPOR'),0),@GuiaId;
 END;

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW [dbo].[ViewDocument]
AS
SELECT     dbo.Int_TipDoc.TipDoc AS Tipo, dbo.MdbDocument.NumDoc AS Documento, dbo.MdbDocument.FecEmi AS Emision, 
                      dbo.MdbDocument.FecVen AS Vencimiento, dbo.MdbDocument.CodCli AS Codigo, dbo.MdbDocument.NomCli AS Cliente, 
                      ISNULL(dbo.MDbClientes.ContacEnt, '') AS ContactoEnt, ISNULL(dbo.MDbClientes.DireccEnt, '') AS DireccionEnt, ISNULL(dbo.MDbClientes.DistriEnt, '') 
                      AS DistritoEnt, ISNULL(dbo.MDbClientes.TelefoEnt, '') AS TelefonoEnt, dbo.Int_Moneda.Moneda AS Divisa, dbo.MdbDocument.SubFac AS SubTotal, 
                      dbo.MdbDocument.RecFac AS Recargo, dbo.MdbDocument.PorsRC AS Porsentaje, dbo.MdbDocument.GasOtr AS Otros, 
                      dbo.MdbDocument.IgvFac AS IGV, dbo.MdbDocument.ImpFac AS Importe, dbo.MdbDocument.ImpFac - ISNULL
                          ((SELECT     Monto
                              FROM         dbo.ViewCancelaciones
                              WHERE     (Documento = dbo.MdbDocument.NumDoc)), 0) AS Saldo, dbo.ViewUltimosStatus.Status, dbo.ViewUltimosStatus.Fecha, 
                      CASE dbo.MdbDocument.SitDoc WHEN 1 THEN 'Activo' ELSE 'Inactivo' END AS Estado, '' AS Referencia
FROM         dbo.MdbDocument LEFT OUTER JOIN
                      dbo.Int_TipDoc ON dbo.MdbDocument.TipDoc = dbo.Int_TipDoc.CodIte LEFT OUTER JOIN
                      dbo.Int_Moneda ON dbo.MdbDocument.DivFac = dbo.Int_Moneda.CodEst LEFT OUTER JOIN
                      dbo.ViewUltimosStatus ON dbo.MdbDocument.NumDoc = dbo.ViewUltimosStatus.NumDoc LEFT OUTER JOIN
                      dbo.MDbClientes ON dbo.MdbDocument.CodCli = dbo.MDbClientes.Asient
UNION
SELECT     'NC' AS Tipo, ViewUltimosStatus_1.NumDoc AS Documento, dbo.MdbNCredito.Emision, dbo.MdbNCredito.Emision AS Vencimiento, 
                      dbo.MdbNCredito.CodCli AS Codigo, dbo.MdbNCredito.Client AS Cliente, '' AS ContactoEnt, '' AS DireccionEnt, '' AS DistritoEnt, '' AS TelefonoEnt, 
                      Int_Moneda_1.Moneda, 0 AS Subtotal, 0 AS Recargo, 0 AS Porsentaje, 0 AS Otros, 0 AS Igv, dbo.MdbNCredito.Total AS Importe, 
                      dbo.MdbNCredito.Total AS Saldo, ViewUltimosStatus_1.Status, ViewUltimosStatus_1.Fecha, 
                      CASE dbo.MdbNCredito.SitDoc WHEN 1 THEN 'Activo' ELSE 'Inactivo' END AS Estado, '' AS Referencia
FROM         dbo.MdbNCredito LEFT OUTER JOIN
                      dbo.ViewUltimosStatus AS ViewUltimosStatus_1 ON dbo.MdbNCredito.Numdoc = ViewUltimosStatus_1.NumDoc LEFT OUTER JOIN
                      dbo.Int_Moneda AS Int_Moneda_1 ON dbo.MdbNCredito.Divisa = Int_Moneda_1.CodEst
GO
CREATE PROCEDURE [dbo].[ProdTraspas_DiarComp]
	@Item numeric(18,0), @Planilla numeric(18,0)=NULL
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY
DECLARE @Bloqueo numeric(18,0); SELECT @Bloqueo=Correl FROM dbo.Int_Correl WITH(UPDLOCK,HOLDLOCK) WHERE DivCor IN('CRP');
IF(SELECT COUNT(*)FROM dbo.Int_Correl WHERE DivCor='CRP')<>1 RAISERROR('Correlativo inexistente o ambiguo.',16,1);

 IF NOT EXISTS(SELECT 1 FROM dbo.MdbDiarRp WITH(UPDLOCK,HOLDLOCK) WHERE Item=@Item) RAISERROR('Registro RP inexistente o ya transferido.',16,1);
 DECLARE @Compilado int;
 IF @Planilla IS NULL EXEC dbo.ReservarCorrelativo 'CRP',@Planilla OUTPUT;
 ELSE IF @TransaccionInicial=0 OR @Planilla<=0 OR @Planilla>(SELECT Correl FROM dbo.Int_Correl WHERE DivCor='CRP') RAISERROR('Reserve la planilla dentro de la transacción del lote.',16,1);
 INSERT dbo.MdbCompRp(NroPla,NroGui,Client,CodCli,PesFis,PesVol,Neto,MonRc,RcPor,MonIgv,Total,Divisa,PuGui,Origen,Produc,Status,Info,Servic,InterMen,IdGuia)
 SELECT @Planilla,NroGui,Client,CodCli,PesFis,PesVol,Neto,MonRc,RcPor,MonIgv,Total,Divisa,PuGui,Origen,Produc,Status,Info,Servic,InterMen,IdGuia FROM dbo.MdbDiarRp WHERE Item=@Item;
 SET @Compilado=SCOPE_IDENTITY();
 UPDATE dbo.MdbHisRP SET IdDiario=NULL,IdCompilado=@Compilado WHERE IdDiario=@Item;
 DELETE dbo.MdbDiarRp WHERE Item=@Item;

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE [dbo].[Correlativo]
	@Divisor varchar(3),
	@Modal varchar(10)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY


 IF @Modal NOT IN('Consultar','Guardar','Reservar') RAISERROR('Modalidad de correlativo inválida.',16,1);
 DECLARE @Numero numeric(18,0);
 IF @Modal='Consultar' BEGIN
  IF(SELECT COUNT(*)FROM dbo.Int_Correl WHERE DivCor=@Divisor)<>1 RAISERROR('Correlativo inexistente o ambiguo.',16,1);
  SELECT @Numero=Correl+1 FROM dbo.Int_Correl WHERE DivCor=@Divisor;
 END ELSE EXEC dbo.ReservarCorrelativo @Divisor,@Numero OUTPUT;
 IF @Numero>999999 RAISERROR('Correlativo excede seis dígitos.',16,1);
 -- Consultar es una vista previa. Reservar devuelve el número ya consumido dentro de la transacción del guardado.
 IF @Modal<>'Guardar' SELECT Correl=CASE WHEN @Divisor='CAC' THEN 'CAC-' ELSE '' END+RIGHT('000000'+CONVERT(varchar(18),@Numero),6);

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE VIEW dbo.ViewDocumentQuerie
AS
SELECT     dbo.Int_TipDoc.TipDoc AS Tipo, dbo.MdbDocument.NumDoc AS Documento, dbo.MdbDocument.NomCli AS Cliente, 
                      dbo.MdbDocument.CodCli AS Codigo, dbo.MdbDocument.FecEmi AS Emision, dbo.MdbDocument.FecVen AS Vencimiento, 
                      Int_Moneda_1.Moneda AS Divisa, dbo.MdbDocument.SubFac AS SubTotal, dbo.MdbDocument.RecFac AS Recargo, 
                      dbo.MdbDocument.PorsRC AS Porsentaje, dbo.MdbDocument.GasOtr AS Gastos, dbo.MdbDocument.IgvFac AS IGV, 
                      dbo.MdbDocument.ImpFac AS Importe, dbo.MdbDocument.ImpFac - ISNULL
                          ((SELECT     Monto
                              FROM         dbo.ViewCancelaciones
                              WHERE     (Documento = dbo.MdbDocument.NumDoc)), 0) AS Saldo, ISNULL(dbo.MDbClientes.Client, '') AS ClienteAsig, 
                      ISNULL(dbo.MDbClientes.RucDni, '') AS RucDni, ISNULL(dbo.MDbClientes.Direcc, '') AS Direccion, ISNULL(dbo.MDbClientes.Distri, '') AS Distrito, 
                      ISNULL(dbo.IntTerritorios.Descrip, '') AS Territorio, ISNULL(dbo.Int_Moneda.Moneda, '') AS Moneda, ISNULL(dbo.MDbClientes.Recarg, '0.00') 
                      AS RecargoAsig, ISNULL(Int_TipDoc_1.TipDoc, '') AS TipoCli, ISNULL(Int_TipDoc_1.NamVar, '') AS DesTipCli,
                          (SELECT     COUNT(ITEM) AS Expr1
                            FROM          dbo.MDbQuerie
                            WHERE      (NumDoc = dbo.MdbDocument.NumDoc) AND (SitDoc = 1)) AS Querie,
                          (SELECT     COUNT(NroGui) AS Expr1
                            FROM          dbo.DbServGui
                            WHERE      (NumDoc = dbo.MdbDocument.NumDoc)) AS Guias
FROM         dbo.MdbDocument LEFT OUTER JOIN
                      dbo.Int_TipDoc ON dbo.MdbDocument.TipDoc = dbo.Int_TipDoc.CodIte LEFT OUTER JOIN
                      dbo.MDbClientes ON dbo.MdbDocument.CodCli = dbo.MDbClientes.Asient LEFT OUTER JOIN
                      dbo.Int_Moneda ON dbo.MDbClientes.Divisa = dbo.Int_Moneda.CodEst LEFT OUTER JOIN
                      dbo.Int_Moneda AS Int_Moneda_1 ON dbo.MdbDocument.DivFac = Int_Moneda_1.CodEst LEFT OUTER JOIN
                      dbo.Int_TipDoc AS Int_TipDoc_1 ON dbo.MDbClientes.TipCom = Int_TipDoc_1.CodIte LEFT OUTER JOIN
                      dbo.IntTerritorios ON dbo.MDbClientes.Territ = dbo.IntTerritorios.Item
GO
CREATE PROCEDURE dbo.VerificarActorOperacion @Actor int,@Modulo int
AS
BEGIN
 SET NOCOUNT ON;
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE ITEM=@Actor AND ActPas=1 AND SitUsu='Activo') BEGIN RAISERROR('Usuario inactivo.',16,1);RETURN;END;
 CREATE TABLE #Accesos(CodMod nvarchar(6),Acceso int);
 INSERT #Accesos EXEC dbo.ObtenerAccesosUsuario @Actor;
 IF NOT EXISTS(SELECT 1 FROM #Accesos a JOIN dbo.Int_ModulosUsuario m ON m.CodMod=a.CodMod WHERE m.Item=@Modulo AND a.Acceso=1) RAISERROR('No tiene permiso para esta operación.',16,1);
 IF @Modulo IN(9,10) AND NOT EXISTS(SELECT 1 FROM #Accesos a JOIN dbo.Int_ModulosUsuario m ON m.CodMod=a.CodMod WHERE m.Item=8 AND a.Acceso=1) RAISERROR('No tiene permiso para Control RP.',16,1);
 IF NOT EXISTS(SELECT 1 FROM #Accesos a JOIN dbo.Int_ModulosUsuario m ON m.CodMod=a.CodMod WHERE m.Item=CASE WHEN @Modulo IN(19,20,21) THEN 18 ELSE 3 END AND a.Acceso=1) RAISERROR('No tiene permiso al grupo.',16,1);
END
GO
CREATE VIEW dbo.InformeDocumentos AS
SELECT Tipo+' '+Documento Codigo,Tipo,Documento,Emision,Vencimiento,CodigoCliente,Cliente,Divisa Moneda,Subtotal,Recargo,IGV,Importe,Saldo,Situacion,Estado,FechaEstado,Referencia,ContactoEntrega,DireccionEntrega,DistritoEntrega,TelefonoEntrega FROM dbo.DocumentosGestion;
GO
CREATE PROCEDURE [dbo].[VerifAccess]
	@Usuari nvarchar(10)

AS
BEGIN
 SET NOCOUNT ON;
	declare @Access int
	Set @Access = isnull((Select Access From MdbUsers Where(Usuari=@Usuari)),0)
	
	
	if @Access=0
		BEGIN
			Select 
				dbo.ViewPlantillaPri.Destino,
				Codigo='A'+RIGHT('000000' + CONVERT(NVARCHAR(10),dbo.ViewPlantillaPri.Codigo),6),
				Acceso=CONVERT(int,isnull((SELECT dbo.IntAccess.Access FROM dbo.MdbUsers INNER JOIN
								dbo.IntAccess ON dbo.MdbUsers.ITEM = dbo.IntAccess.Usuari Where(dbo.IntAccess.CodMod=dbo.ViewPlantillaPri.Codigo) and (dbo.MdbUsers.ITEM=(Select Item From MdbUsers Where(Usuari=@Usuari)))),0)),
				Origen
			From dbo.ViewPlantillaPri
		END
	ELSE
		if @Access=1
			BEGIN
				Select 
					dbo.ViewPlantillaPri.Destino,
					Codigo='A'+RIGHT('000000' + CONVERT(NVARCHAR(10),dbo.ViewPlantillaPri.Codigo),6),
					Acceso=1,
					Origen
				From dbo.ViewPlantillaPri
			END
		ELSE
		
			BEGIN
				Select 
					dbo.ViewPlantillaPri.Destino,
					Codigo='A'+RIGHT('000000' + CONVERT(NVARCHAR(10),dbo.ViewPlantillaPri.Codigo),6),
					Acceso=CONVERT(int,isnull((SELECT dbo.IntAccesPlant.Access
								FROM dbo.MdbUsers INNER JOIN dbo.IntNiveles ON dbo.MdbUsers.Access = dbo.IntNiveles.item INNER JOIN dbo.IntAccesPlant ON dbo.IntNiveles.Planti = dbo.IntAccesPlant.Planti Where(dbo.IntAccesPlant.CodMod=dbo.ViewPlantillaPri.Codigo) and (dbo.MdbUsers.ITEM=(Select Item From MdbUsers Where(Usuari=@Usuari)))),0)),
					Origen
				From dbo.ViewPlantillaPri
			END
END
GO
CREATE PROCEDURE dbo.VerificarActorQuery @Actor int,@Modulo int
AS
BEGIN
 SET NOCOUNT ON;
 BEGIN TRY
 IF @Modulo NOT IN(5,6,7) RAISERROR('Módulo de Queries inválido.',16,1);
 EXEC dbo.VerificarActorOperacion @Actor,@Modulo;
 CREATE TABLE #PermisosQuery(CodMod nvarchar(6),Acceso int);
 INSERT #PermisosQuery EXEC dbo.ObtenerAccesosUsuario @Actor;
 IF NOT EXISTS(SELECT 1 FROM #PermisosQuery p JOIN dbo.Int_ModulosUsuario m ON m.CodMod=p.CodMod WHERE m.Item=4 AND p.Acceso=1) RAISERROR('No tiene acceso al grupo Solicitud de Queries.',16,1);
 RETURN 0;
 END TRY BEGIN CATCH DECLARE @Error nvarchar(2048);SET @Error=ERROR_MESSAGE();RAISERROR('%s',16,1,@Error);RETURN 1;END CATCH;
END
GO
CREATE PROCEDURE dbo.ConsultarDocumentosGestion @Actor int,@Tipo nvarchar(2)=N'',@Buscar nvarchar(100)=N'',@Exacto bit=0,@Modulo int=11
AS
BEGIN
 SET NOCOUNT ON;
 IF @Modulo NOT IN(11,19) BEGIN RAISERROR('Módulo inválido.',16,1);RETURN;END;
 BEGIN TRY EXEC dbo.VerificarActorOperacion @Actor,@Modulo; END TRY
 BEGIN CATCH DECLARE @ErrorAcceso nvarchar(2048);SET @ErrorAcceso=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorAcceso);RETURN;END CATCH;
 SELECT TOP(200) d.*,
 ISNULL(STUFF((SELECT N', '+g.NroGui FROM dbo.DbServGui g WHERE (d.Tipo='NC' AND g.IdNotaCredito=d.IdDocumento) OR(d.Tipo<>'NC' AND g.IdFactura=d.IdDocumento) ORDER BY g.Linea,g.Item FOR XML PATH(''),TYPE).value('.','nvarchar(max)'),1,2,N''),N'') Guias,
 ISNULL(STUFF((SELECT N', '+g.Servic FROM dbo.DbServGui g WHERE (d.Tipo='NC' AND g.IdNotaCredito=d.IdDocumento) OR(d.Tipo<>'NC' AND g.IdFactura=d.IdDocumento) ORDER BY g.Linea,g.Item FOR XML PATH(''),TYPE).value('.','nvarchar(max)'),1,2,N''),N'') Servicios
 FROM dbo.DocumentosGestion d
 WHERE (@Tipo=N'' OR d.Tipo=@Tipo) AND
 ((@Exacto=1 AND d.Documento=@Buscar) OR (@Exacto=0 AND (@Buscar=N'' OR CHARINDEX(@Buscar,d.Documento)>0 OR CHARINDEX(@Buscar,d.Cliente)>0 OR CHARINDEX(@Buscar,d.CodigoCliente)>0)))
 ORDER BY d.Emision DESC,d.Documento,d.Tipo;
END
GO
CREATE PROCEDURE dbo.ConsultarCamposInforme @Actor int,@Origen int
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.ValidarActorInforme @Actor,@Origen;IF @P<>0 RETURN;
 DECLARE @Vista sysname;SELECT @Vista=Vista FROM dbo.OrigenesInforme()WHERE Origen=@Origen;
 SELECT c.name Campo,TYPE_NAME(c.system_type_id) Tipo FROM sys.columns c WHERE c.object_id=OBJECT_ID(N'dbo.'+@Vista)ORDER BY c.column_id;
END
GO
CREATE PROCEDURE dbo.ImportarInformacion @Actor int,@Archivo nvarchar(255),@Datos xml,@Confirmar bit=0
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 IF @@TRANCOUNT<>0 BEGIN RAISERROR('La carga requiere su propia transacción.',16,1);RETURN;END;
 DECLARE @Hoja nvarchar(30),@NumeroFila int,@Fila xml,@Mensaje nvarchar(2048),@Total int,@CampoLargo nvarchar(128);
 DECLARE @Longitudes TABLE(Hoja nvarchar(30),Campo nvarchar(30),Maximo int,PRIMARY KEY(Hoja,Campo));
 INSERT @Longitudes VALUES
 (N'Clientes',N'Asient',20),(N'Clientes',N'CLIENT',255),(N'Clientes',N'RUCDNI',20),(N'Clientes',N'DIRECC',255),(N'Clientes',N'TELEF1',50),
 (N'Clientes',N'Contac',100),(N'Clientes',N'ContacEnt',100),(N'Clientes',N'DireccEnt',255),(N'Clientes',N'TelefoEnt',50),
 (N'Guias',N'NumeroGuia',10),(N'Guias',N'Origen',5),(N'Guias',N'Consignatario',255),(N'Guias',N'Gestion',10),(N'Guias',N'Producto',5),
 (N'Documentos',N'Tipo',2),(N'Documentos',N'Documento',14),(N'Documentos',N'Vencimiento',30),(N'Documentos',N'Cliente',20),(N'Documentos',N'Situacion',15),(N'Documentos',N'Referencia',14),
 (N'Servicios',N'Documento',14),(N'Servicios',N'NumeroGuia',10),(N'Servicios',N'Servicio',50);
 BEGIN TRY
 BEGIN TRAN;
 EXEC dbo.VerificarActorOperacion @Actor,21;
 IF @Datos IS NULL OR @Datos.exist('/carga/r')=0 RAISERROR('La carga está vacía.',16,1);
 IF @Datos.exist('/carga/r[not(@hoja="Clientes" or @hoja="Guias" or @hoja="Documentos" or @hoja="Servicios")]')=1 RAISERROR('Hoja desconocida.',16,1);
 DECLARE Filas CURSOR LOCAL FAST_FORWARD FOR SELECT n.query('.'),n.value('@hoja','nvarchar(30)'),n.value('@fila','int') FROM @Datos.nodes('/carga/r')x(n) ORDER BY CASE n.value('@hoja','nvarchar(30)')WHEN 'Clientes'THEN 1 WHEN 'Guias'THEN 2 WHEN 'Documentos'THEN 3 ELSE 4 END,n.value('@fila','int');
 OPEN Filas;FETCH NEXT FROM Filas INTO @Fila,@Hoja,@NumeroFila;
 SET @Total=0;
 WHILE @@FETCH_STATUS=0 BEGIN
 SET @Fila=@Fila.query('/r');
 -- Validar el valor completo antes de convertirlo o buscar sus referencias.
 SET @CampoLargo=NULL;
 SELECT TOP(1) @CampoLargo=a.value('local-name(.)','nvarchar(128)')
 FROM @Fila.nodes('/r/@*')x(a) LEFT JOIN @Longitudes l ON l.Hoja=@Hoja AND l.Campo=a.value('local-name(.)','nvarchar(128)')
 WHERE DATALENGTH(a.value('.','nvarchar(max)'))>2*ISNULL(l.Maximo,4000)
 ORDER BY a.value('local-name(.)','nvarchar(128)');
 IF @CampoLargo IS NOT NULL BEGIN SET @Mensaje=N'Campo '+@CampoLargo+N' demasiado largo. No se recorta ni se cambia su referencia.';RAISERROR('%s',16,1,@Mensaje);END;
 IF @Hoja='Clientes' BEGIN
 DECLARE @ClienteItem int,@VersionCliente binary(8);
SET @ClienteItem=NULL;SET @VersionCliente=NULL;SELECT @ClienteItem=Item,@VersionCliente=VersionFila FROM dbo.DbClientes WITH(UPDLOCK,HOLDLOCK)WHERE Asient=@Fila.value('(/r/@Asient)[1]','nvarchar(20)');
IF @ClienteItem IS NULL SET @ClienteItem=0;
DECLARE @Asient nvarchar(20);
SET @Asient=NULLIF(@Fila.value('(/r/@Asient)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@Asient)[1]','nvarchar(4000)'))>20 RAISERROR('Campo Asient demasiado largo.',16,1);
DECLARE @CLIENT nvarchar(255);
SET @CLIENT=NULLIF(@Fila.value('(/r/@CLIENT)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@CLIENT)[1]','nvarchar(4000)'))>255 RAISERROR('Campo CLIENT demasiado largo.',16,1);
DECLARE @RUCDNI nvarchar(20);
SET @RUCDNI=NULLIF(@Fila.value('(/r/@RUCDNI)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@RUCDNI)[1]','nvarchar(4000)'))>20 RAISERROR('Campo RUCDNI demasiado largo.',16,1);
DECLARE @DIRECC nvarchar(255);
SET @DIRECC=NULLIF(@Fila.value('(/r/@DIRECC)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@DIRECC)[1]','nvarchar(4000)'))>255 RAISERROR('Campo DIRECC demasiado largo.',16,1);
DECLARE @TELEF1 nvarchar(50);
SET @TELEF1=NULLIF(@Fila.value('(/r/@TELEF1)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@TELEF1)[1]','nvarchar(4000)'))>50 RAISERROR('Campo TELEF1 demasiado largo.',16,1);
DECLARE @IdTerritorio int;
SET @IdTerritorio=NULLIF(@Fila.value('(/r/@IdTerritorio)[1]','nvarchar(4000)'),'');
DECLARE @Contac nvarchar(100);
SET @Contac=NULLIF(@Fila.value('(/r/@Contac)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@Contac)[1]','nvarchar(4000)'))>100 RAISERROR('Campo Contac demasiado largo.',16,1);
DECLARE @ContacEnt nvarchar(100);
SET @ContacEnt=NULLIF(@Fila.value('(/r/@ContacEnt)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@ContacEnt)[1]','nvarchar(4000)'))>100 RAISERROR('Campo ContacEnt demasiado largo.',16,1);
DECLARE @DireccEnt nvarchar(255);
SET @DireccEnt=NULLIF(@Fila.value('(/r/@DireccEnt)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@DireccEnt)[1]','nvarchar(4000)'))>255 RAISERROR('Campo DireccEnt demasiado largo.',16,1);
DECLARE @TelefoEnt nvarchar(50);
SET @TelefoEnt=NULLIF(@Fila.value('(/r/@TelefoEnt)[1]','nvarchar(4000)'),'');
IF LEN(@Fila.value('(/r/@TelefoEnt)[1]','nvarchar(4000)'))>50 RAISERROR('Campo TelefoEnt demasiado largo.',16,1);
DECLARE @Recarg money;
SET @Recarg=NULLIF(@Fila.value('(/r/@Recarg)[1]','nvarchar(4000)'),'');
DECLARE @IdComprobante tinyint;
SET @IdComprobante=NULLIF(@Fila.value('(/r/@IdComprobante)[1]','nvarchar(4000)'),'');
DECLARE @IdMoneda tinyint;
SET @IdMoneda=NULLIF(@Fila.value('(/r/@IdMoneda)[1]','nvarchar(4000)'),'');
DECLARE @IdNivel tinyint;
SET @IdNivel=NULLIF(@Fila.value('(/r/@IdNivel)[1]','nvarchar(4000)'),'');
DECLARE @IdCobranza tinyint;
SET @IdCobranza=NULLIF(@Fila.value('(/r/@IdCobranza)[1]','nvarchar(4000)'),'');
DECLARE @EsInternacional bit;
SET @EsInternacional=NULLIF(@Fila.value('(/r/@EsInternacional)[1]','nvarchar(4000)'),'');
DECLARE @IdPais int;
SET @IdPais=NULLIF(@Fila.value('(/r/@IdPais)[1]','nvarchar(4000)'),'');
DECLARE @IdDistrito int;
SET @IdDistrito=NULLIF(@Fila.value('(/r/@IdDistrito)[1]','nvarchar(4000)'),'');
DECLARE @IdDistritoEntrega int;
SET @IdDistritoEntrega=NULLIF(@Fila.value('(/r/@IdDistritoEntrega)[1]','nvarchar(4000)'),'');
DECLARE @IdTipoPersona tinyint;
SET @IdTipoPersona=NULLIF(@Fila.value('(/r/@IdTipoPersona)[1]','nvarchar(4000)'),'');
DECLARE @IdTipoIdentificacion tinyint;
SET @IdTipoIdentificacion=NULLIF(@Fila.value('(/r/@IdTipoIdentificacion)[1]','nvarchar(4000)'),'');
DECLARE @ClienteGuardado TABLE(Item int);
INSERT @ClienteGuardado EXEC dbo.GuardarCliente @Item=@ClienteItem,@Version=@VersionCliente,@Asient=@Asient,@CLIENT=@CLIENT,@RUCDNI=@RUCDNI,@DIRECC=@DIRECC,@TELEF1=@TELEF1,@IdTerritorio=@IdTerritorio,@Contac=@Contac,@ContacEnt=@ContacEnt,@DireccEnt=@DireccEnt,@TelefoEnt=@TelefoEnt,@Recarg=@Recarg,@IdComprobante=@IdComprobante,@IdMoneda=@IdMoneda,@IdNivel=@IdNivel,@IdCobranza=@IdCobranza,@EsInternacional=@EsInternacional,@IdPais=@IdPais,@IdDistrito=@IdDistrito,@IdDistritoEntrega=@IdDistritoEntrega,@IdTipoPersona=@IdTipoPersona,@IdTipoIdentificacion=@IdTipoIdentificacion,@Importando=1;

 END;
 IF @Hoja='Guias' BEGIN
 DECLARE @Guia nvarchar(10),@Origen nvarchar(5),@Fecha datetime,@Producto nvarchar(5),@Consignatario nvarchar(255),@Gestion nvarchar(10),@GuiaItem int,@Asiento numeric(18,0);
 IF LEN(@Fila.value('(/r/@NumeroGuia)[1]','nvarchar(4000)'))>10 OR LEN(@Fila.value('(/r/@Consignatario)[1]','nvarchar(4000)'))>255 RAISERROR('Guía o consignatario demasiado largos.',16,1);
 SELECT @Guia=@Fila.value('(/r/@NumeroGuia)[1]','nvarchar(10)'),@Origen=@Fila.value('(/r/@Origen)[1]','nvarchar(5)'),@Fecha=@Fila.value('(/r/@FechaEnvio)[1]','datetime'),@Producto=@Fila.value('(/r/@Producto)[1]','nvarchar(5)'),@Consignatario=@Fila.value('(/r/@Consignatario)[1]','nvarchar(255)'),@Gestion=@Fila.value('(/r/@Gestion)[1]','nvarchar(10)');
 IF LEN(@Fila.value('(/r/@Origen)[1]','nvarchar(4000)'))>5 OR LEN(@Fila.value('(/r/@Producto)[1]','nvarchar(4000)'))>5 OR LEN(@Fila.value('(/r/@Gestion)[1]','nvarchar(4000)'))>10 RAISERROR('Origen, producto o gestión demasiado largos.',16,1);
 IF @Guia='' OR @Consignatario='' OR @Gestion NOT IN('RP','Sender') OR @Producto NOT IN('15D','15N','48N') RAISERROR('Revise guía, consignatario, gestión y producto.',16,1);
 IF @Origen IN('LIM','PEO','PEI') OR NOT EXISTS(SELECT 1 FROM dbo.int_IATA WHERE codigo=@Origen) RAISERROR('Origen local, ambiguo o no catalogado: revise el alcance de importación.',16,1);
 SET @GuiaItem=NULL;SELECT @GuiaItem=Item FROM dbo.DbGuias WITH(UPDLOCK,HOLDLOCK) WHERE NroGui=@Guia;
 IF @GuiaItem IS NOT NULL AND EXISTS(SELECT 1 FROM dbo.DbGuias WHERE Item=@GuiaItem AND(Origen<>@Origen OR PickUP<>@Fecha OR Produc<>@Producto OR Client<>@Consignatario OR Tipo<>@Gestion)) RAISERROR('La guía ya existe con otros datos. No se duplica ni reinicia su gestión.',16,1);
 IF @GuiaItem IS NULL BEGIN
 INSERT dbo.DbGuias(NroGui,Origen,PickUP,Produc,Client,Tipo)VALUES(@Guia,@Origen,@Fecha,@Producto,@Consignatario,@Gestion);SET @GuiaItem=SCOPE_IDENTITY();
 IF @Gestion='RP' BEGIN
 EXEC dbo.ReservarCorrelativo 'CGP',@Asiento OUTPUT;
 DECLARE @AsientoTexto nvarchar(20);SET @AsientoTexto=RIGHT('0000000000'+CONVERT(varchar(20),@Asiento),10);
 IF @Asiento>9999999999 RAISERROR('Correlativo RP agotado.',16,1);
 INSERT dbo.DbsDiarRp(Asient,NroGui,Client,PuGui,Origen,TipGui,Salida,Reten,MonNet,MonRc,MonIgv,MonTot,Moneda,RcPor,IdGuia)VALUES(@AsientoTexto,@Guia,@Consignatario,@Fecha,@Origen,@Producto,CASE WHEN @Producto='15D'THEN 1 ELSE 0 END,0,0,0,0,0,1,0,@GuiaItem);
 INSERT dbo.DbsCompRp(Asient,NroGui,Client,PuGui,Origen,TipGui,MonNet,MonRc,MonIgv,MonTot,Moneda,IdGuia)VALUES(@AsientoTexto,@Guia,@Consignatario,@Fecha,@Origen,@Producto,0,0,0,0,1,@GuiaItem);
 END;END;END;
 IF @Hoja='Documentos' BEGIN
 DECLARE @Tipo nvarchar(2),@Documento nvarchar(14),@Emision datetime,@Vencimiento datetime,@Codigo nvarchar(20),@Moneda tinyint,@Subtotal money,@Recargo money,@Igv money,@Importe money,@Situacion nvarchar(15),@Referencia nvarchar(14),@IdReferencia int,@IdCliente int,@NombreCliente nvarchar(255),@IdTipo tinyint;
 IF LEN(@Fila.value('(/r/@Documento)[1]','nvarchar(4000)'))>14 RAISERROR('Documento excede 14 caracteres.',16,1);
 SELECT @Tipo=@Fila.value('(/r/@Tipo)[1]','nvarchar(2)'),@Documento=@Fila.value('(/r/@Documento)[1]','nvarchar(14)'),@Emision=@Fila.value('(/r/@Emision)[1]','datetime'),@Vencimiento=NULLIF(@Fila.value('(/r/@Vencimiento)[1]','nvarchar(30)'),''),@Codigo=@Fila.value('(/r/@Cliente)[1]','nvarchar(20)'),@Moneda=@Fila.value('(/r/@Moneda)[1]','tinyint'),@Subtotal=@Fila.value('(/r/@Subtotal)[1]','money'),@Recargo=@Fila.value('(/r/@Recargo)[1]','money'),@Igv=@Fila.value('(/r/@IGV)[1]','money'),@Importe=@Fila.value('(/r/@Total)[1]','money'),@Situacion=@Fila.value('(/r/@Situacion)[1]','nvarchar(15)'),@Referencia=@Fila.value('(/r/@Referencia)[1]','nvarchar(14)');
 IF @Tipo NOT IN('FT','BV','NC','ND') OR @Documento NOT LIKE @Tipo+'-%' OR LEN(@Documento)<=3 OR @Importe<0 OR @Subtotal<0 OR @Recargo<0 OR @Igv<0 OR @Situacion NOT IN('ACT','ANL') OR @Vencimiento<@Emision RAISERROR('Tipo, número, fechas, situación o importes inválidos.',16,1);
 SET @IdCliente=NULL;SET @IdTipo=NULL;SET @IdReferencia=NULL;
 SELECT @IdCliente=Item,@NombreCliente=CLIENT FROM dbo.DbClientes WITH(HOLDLOCK)WHERE Asient=@Codigo AND Activo=1;
 SELECT @IdTipo=CodIte FROM dbo.Int_TipDoc WHERE TipDoc=@Tipo AND Activo=1;
 IF @IdCliente IS NULL OR @IdTipo IS NULL OR NOT EXISTS(SELECT 1 FROM dbo.Int_Moneda WHERE CodEst=@Moneda AND ActEst=1) RAISERROR('Cliente, moneda o tipo de documento inexistente o inactivo.',16,1);
 IF @Tipo IN('NC','ND')BEGIN
 SELECT @IdReferencia=Item FROM dbo.DBEmiFac WITH(HOLDLOCK)WHERE NumDoc=@Referencia AND IdCliente=@IdCliente AND TipDoc IN(1,2);
 IF @IdReferencia IS NULL RAISERROR('La nota requiere factura/boleta de referencia del mismo cliente.',16,1);
 END ELSE IF @Referencia<>'' RAISERROR('Solo las notas llevan referencia.',16,1);
 IF @Tipo='NC'BEGIN
 IF EXISTS(SELECT 1 FROM dbo.DbNCredito WITH(UPDLOCK,HOLDLOCK)WHERE NumDoc=@Documento)BEGIN
 IF NOT EXISTS(SELECT 1 FROM dbo.DbNCredito WHERE NumDoc=@Documento AND IdCliente=@IdCliente AND FecEmi=@Emision AND Monto=@Importe AND Tarifa=@Subtotal AND Recargo=@Recargo AND Igv=@Igv AND ISNULL(FecVen,'19000101')=ISNULL(@Vencimiento,'19000101') AND DivMon=CONVERT(nvarchar(50),@Moneda) AND SitDoc=@Situacion AND IdDocumentoReferencia=@IdReferencia)RAISERROR('Nota existente diferente: requiere revisión, no se sobrescribe.',16,1);
 END ELSE INSERT dbo.DbNCredito(NumDoc,FecEmi,CodCli,DivMon,Monto,NumRef,NomCli,Tarifa,SitDoc,IdCliente,IdDocumentoReferencia,Recargo,Igv,FecVen)VALUES(@Documento,@Emision,@Codigo,CONVERT(nvarchar(50),@Moneda),@Importe,@Referencia,@NombreCliente,@Subtotal,@Situacion,@IdCliente,@IdReferencia,@Recargo,@Igv,@Vencimiento);
 END ELSE BEGIN
 IF EXISTS(SELECT 1 FROM dbo.DBEmiFac WITH(UPDLOCK,HOLDLOCK)WHERE NumDoc=@Documento)BEGIN
 IF NOT EXISTS(SELECT 1 FROM dbo.DBEmiFac WHERE NumDoc=@Documento AND IdCliente=@IdCliente AND TipDoc=@IdTipo AND FecEmi=@Emision AND ISNULL(FecVen,'19000101')=ISNULL(@Vencimiento,'19000101') AND ImpFac=@Importe AND SubFac=@Subtotal AND RecFac=@Recargo AND IgvFac=@Igv AND DivFac=CONVERT(nvarchar(50),@Moneda) AND SitDoc=@Situacion AND ISNULL(IdDocumentoReferencia,0)=ISNULL(@IdReferencia,0))RAISERROR('Documento existente diferente: requiere revisión, no se sobrescribe.',16,1);
 END ELSE INSERT dbo.DBEmiFac(TipDoc,NumDoc,FecEmi,FecVen,CodCli,NomCli,DivFac,SubFac,RecFac,IgvFac,ImpFac,ImpSal,SitDoc,IdCliente,IdDocumentoReferencia)VALUES(@IdTipo,@Documento,@Emision,@Vencimiento,@Codigo,@NombreCliente,CONVERT(nvarchar(50),@Moneda),@Subtotal,@Recargo,@Igv,@Importe,@Importe,@Situacion,@IdCliente,@IdReferencia);
 END;END;
 IF @Hoja='Servicios' BEGIN
 DECLARE @DocServicio nvarchar(14),@Linea int,@NumeroGuia nvarchar(10),@Servicio nvarchar(50),@IdFactura int,@IdNC int,@IdG int,@TipoServicio tinyint;
 IF LEN(@Fila.value('(/r/@Servicio)[1]','nvarchar(4000)'))>50 RAISERROR('Servicio demasiado largo.',16,1);
 SELECT @DocServicio=@Fila.value('(/r/@Documento)[1]','nvarchar(14)'),@Linea=@Fila.value('(/r/@Linea)[1]','int'),@NumeroGuia=@Fila.value('(/r/@NumeroGuia)[1]','nvarchar(10)'),@Servicio=@Fila.value('(/r/@Servicio)[1]','nvarchar(50)');
 SET @IdFactura=NULL;SET @IdNC=NULL;SET @IdG=NULL;SET @TipoServicio=NULL;
 SELECT @IdFactura=Item,@TipoServicio=TipDoc FROM dbo.DBEmiFac WITH(HOLDLOCK)WHERE NumDoc=@DocServicio;
 SELECT @IdNC=ITEM FROM dbo.DbNCredito WITH(HOLDLOCK)WHERE NumDoc=@DocServicio;
 IF @IdNC IS NOT NULL SELECT @TipoServicio=CodIte FROM dbo.Int_TipDoc WHERE TipDoc='NC';
 SELECT @IdG=Item FROM dbo.DbGuias WITH(HOLDLOCK)WHERE NroGui=@NumeroGuia;
 IF @Linea<=0 OR @IdG IS NULL OR(@IdFactura IS NULL AND @IdNC IS NULL)OR @Servicio='' RAISERROR('Línea, documento, guía o servicio inválidos.',16,1);
 IF EXISTS(SELECT 1 FROM dbo.DbServGui WITH(UPDLOCK,HOLDLOCK)WHERE NumDoc=@DocServicio AND Linea=@Linea)BEGIN
 IF NOT EXISTS(SELECT 1 FROM dbo.DbServGui WHERE NumDoc=@DocServicio AND Linea=@Linea AND IdGuia=@IdG AND Servic=@Servicio)RAISERROR('Línea existente diferente. No se sobrescribe.',16,1);
 END ELSE INSERT dbo.DbServGui(TipDoc,NumDoc,NroGui,Servic,IdFactura,IdNotaCredito,IdGuia,Linea)VALUES(@TipoServicio,@DocServicio,@NumeroGuia,@Servicio,@IdFactura,@IdNC,@IdG,@Linea);
 END;
 SET @Total=@Total+1;
 FETCH NEXT FROM Filas INTO @Fila,@Hoja,@NumeroFila;
 END;CLOSE Filas;DEALLOCATE Filas;
 IF @Confirmar=1 BEGIN INSERT dbo.CargasInformacion(IdUsuario,Archivo,Filas)VALUES(@Actor,@Archivo,@Total);COMMIT;END ELSE ROLLBACK;
 SELECT Filas=@Total,Confirmado=@Confirmar;
 END TRY BEGIN CATCH
 IF XACT_STATE()<>0 ROLLBACK;
 SET @Mensaje=ISNULL(@Hoja,'Carga')+' fila '+ISNULL(CONVERT(varchar(12),@NumeroFila),'?')+': '+ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);
 END CATCH
END
GO
CREATE PROCEDURE dbo.CobValidarActor @Actor int,@Modulo int
AS
BEGIN
 SET NOCOUNT ON;BEGIN TRY
 IF @Modulo NOT BETWEEN 12 AND 16 RAISERROR('Módulo de cobranzas inválido.',16,1);
 EXEC dbo.VerificarActorOperacion @Actor,@Modulo;RETURN 0;
 END TRY BEGIN CATCH DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);RETURN 1;END CATCH;
END
GO
CREATE PROCEDURE dbo.ConsultarEstadosCargos @Actor int
AS
BEGIN SET NOCOUNT ON;BEGIN TRY EXEC dbo.VerificarActorOperacion @Actor,19; END TRY
 BEGIN CATCH DECLARE @ErrorAcceso nvarchar(2048);SET @ErrorAcceso=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorAcceso);RETURN;END CATCH;
 SELECT ITEM,TitEst,Status,ActCor,Correl FROM dbo.Int_StatusFac WHERE ActEst=1 ORDER BY ITEM; END
GO
CREATE PROCEDURE dbo.ConsultarDetalleDocumento @Actor int,@Tipo nvarchar(2),@IdDocumento int,@Seccion int,@Modulo int=11
AS
BEGIN
 SET NOCOUNT ON;
 IF @Modulo NOT IN(11,19) BEGIN RAISERROR('Módulo inválido.',16,1);RETURN;END;
 BEGIN TRY EXEC dbo.VerificarActorOperacion @Actor,@Modulo; END TRY
 BEGIN CATCH DECLARE @ErrorAcceso nvarchar(2048);SET @ErrorAcceso=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorAcceso);RETURN;END CATCH;
 IF NOT EXISTS(SELECT 1 FROM dbo.DocumentosGestion WHERE Tipo=@Tipo AND IdDocumento=@IdDocumento) BEGIN RAISERROR('Documento no encontrado.',16,1);RETURN;END;
 IF @Seccion=1
 SELECT g.NroGui Guia,g.Servic Servicio,gu.Origen,gu.PickUP FechaEnvio,gu.Produc Producto,gu.Tipo Gestion
 FROM dbo.DbServGui g LEFT JOIN dbo.DbGuias gu ON gu.Item=g.IdGuia
 WHERE (@Tipo='NC' AND g.IdNotaCredito=@IdDocumento) OR(@Tipo<>'NC' AND g.IdFactura=@IdDocumento) ORDER BY g.Linea,g.Item;
 ELSE IF @Seccion=2
 SELECT h.FecEst Fecha,COALESCE(s.Status,h.Estado) Estado,ISNULL(u.Usuari,h.NomUsu) Usuario,
 h.Activo,r.Numero Registro FROM dbo.DbHisFac h LEFT JOIN dbo.Int_StatusFac s ON s.ITEM=h.IdEstado
 LEFT JOIN dbo.Int_Usuario u ON u.ITEM=h.IdUsuario LEFT JOIN dbo.RegistrosCargos r ON r.Item=h.IdRegistroCargo
 WHERE (@Tipo='NC' AND h.IdNotaCredito=@IdDocumento) OR(@Tipo<>'NC' AND h.IdFactura=@IdDocumento) ORDER BY h.FecEst DESC,h.ITEM DESC;
 ELSE IF @Seccion=3
 SELECT Fecha,Courier Responsable FROM dbo.DbHisCourier WHERE (@Tipo='NC' AND IdNotaCredito=@IdDocumento) OR(@Tipo<>'NC' AND IdFactura=@IdDocumento) ORDER BY Fecha DESC,ITEM DESC;
 ELSE IF @Seccion=4
 SELECT d.Tipo,d.Documento,d.Emision,d.Divisa,d.Importe,d.Situacion,d.Referencia
 FROM dbo.DocumentosGestion actual JOIN dbo.DocumentosGestion d
 ON (actual.Tipo<>'NC' AND d.Referencia=actual.Documento) OR (d.Documento=actual.Referencia)
 WHERE actual.Tipo=@Tipo AND actual.IdDocumento=@IdDocumento ORDER BY d.Emision,d.Documento;
 ELSE RAISERROR('Sección inválida.',16,1);
END
GO
CREATE PROCEDURE dbo.ConsultarPizarraRP @Actor int,@Compilatorio bit=0,@Filtro nvarchar(100)=N'',@Grupo int=-1,@SoloSalida bit=0
AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Modulo int;SET @Modulo=CASE WHEN @Compilatorio=1 THEN 10 ELSE 9 END;
 EXEC dbo.VerificarActorOperacion @Actor,@Modulo;
 IF @Compilatorio=0
 SELECT d.ITEM,d.Asient,d.NroGui,d.Client,d.IdCliente,d.PuGui,d.Origen,d.TipGui,d.Estado,ISNULL(e.Status,'')EstadoGestion,d.Salida,d.Reten,d.PesFis,d.PesVol,d.MonNet,d.MonRc,d.RcPor,d.MonIgv,d.MonTot,d.Moneda,CONVERT(nvarchar(1000),d.Info)Info,d.VersionFila
 FROM dbo.DbsDiarRp d LEFT JOIN dbo.Int_StatusRP e ON e.CodEst=d.Estado
 WHERE(@Filtro='' OR CHARINDEX(@Filtro,d.NroGui)>0 OR CHARINDEX(@Filtro,d.Client)>0)
 AND(@Grupo=-2 OR(@Grupo=0 AND d.IdCliente IS NULL)OR(@Grupo=-1 AND d.IdCliente IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.Int_DivRp g WHERE g.IdCliente=d.IdCliente))OR(@Grupo>0 AND d.IdCliente=@Grupo AND EXISTS(SELECT 1 FROM dbo.Int_DivRp g WHERE g.IdCliente=@Grupo)))
 AND(@SoloSalida=0 OR d.Salida=1)ORDER BY d.PuGui,d.ITEM;
 ELSE
 SELECT c.ITEM,c.Asient,c.NroGui,c.Client,c.IdCliente,c.PuGui,c.Origen,c.TipGui,c.NroPla,c.FecPla,c.FINAL,c.PesFis,c.PesVol,c.MonNet,c.MonRc,c.MonIgv,c.MonTot,c.Moneda,c.NroDoc,c.NroDua,CONVERT(nvarchar(1000),c.Info)Info,c.VersionFila
 FROM dbo.DbsCompRp c WHERE NULLIF(c.FINAL,'')IS NOT NULL AND(@Filtro='' OR CHARINDEX(@Filtro,c.NroGui)>0 OR CHARINDEX(@Filtro,c.Client)>0)ORDER BY c.FecPla DESC,c.ITEM DESC;
END
GO
CREATE PROCEDURE dbo.BuscarClientesGrupoRP @Actor int,@Nombre nvarchar(100)=N''
AS
BEGIN
 SET NOCOUNT ON;
 EXEC dbo.VerificarActorOperacion @Actor,9;
 SELECT TOP 100 c.Item,c.Asient,c.CLIENT,c.RUCDNI,
 CASE WHEN EXISTS(SELECT 1 FROM dbo.Int_DivRp g WHERE g.IdCliente=c.Item)THEN N'Sí' ELSE N'No' END EnGrupo
 FROM dbo.DbClientes c WHERE c.Activo=1 AND(@Nombre='' OR CHARINDEX(@Nombre,c.CLIENT)>0)
 ORDER BY c.CLIENT,c.Item;
END
GO
CREATE PROCEDURE [dbo].[Valid_LoginClev]
	@Usuario varchar(50),
	@Pasword varchar(150)
AS
BEGIN
 SET NOCOUNT ON;
 IF @Usuario IS NULL OR LEN(@Usuario)>5 RETURN;

 IF EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE Usuari=@Usuario AND Paswor=@Pasword AND ActPas=1 AND SitUsu=1)
 EXEC dbo.VerifAccess @Usuari=@Usuario;
END
GO
CREATE PROCEDURE [dbo].[InsertGuiasManual]
	@Guia Numeric(18,0),
	@Origen nvarchar(10),
	@Fecha datetime,
	@Producto nvarchar(10),
	@Codigo Numeric(18,0),
	@Modal tinyint
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @TransaccionInicial int,@PuntoOperacionClever varchar(32);SET @TransaccionInicial=@@TRANCOUNT;SET @PuntoOperacionClever=REPLACE(CONVERT(varchar(36),NEWID()),'-','');
 IF @TransaccionInicial=0 BEGIN TRANSACTION; ELSE SAVE TRANSACTION @PuntoOperacionClever;
 BEGIN TRY


 IF NOT EXISTS(SELECT 1 FROM dbo.MDbClientes WHERE Asient=@Codigo) RAISERROR('Cliente inexistente.',16,1);
 DECLARE @ClienteNombre nvarchar(300);
 SELECT @ClienteNombre=Client FROM dbo.MDbClientes WHERE Asient=@Codigo;
 -- La asociación completa se comprueba en la rutina común; ambos modos producen una RP utilizable.
 IF @Modal NOT IN(1,2) RAISERROR('Modalidad de ingreso de guía inválida.',16,1);
 EXEC dbo.InsertGuias @Guia,@Origen,@Fecha,@Producto,@ClienteNombre,'R';
 UPDATE d SET CodCli=@Codigo FROM dbo.MdbDiarRp d JOIN dbo.MdbGuias g ON d.IdGuia=g.Item WHERE g.Guia=@Guia AND g.Origen=@Origen AND g.Fecha=@Fecha AND g.Client=@ClienteNombre;

 IF @TransaccionInicial=0 COMMIT TRANSACTION;
 END TRY
 BEGIN CATCH
  IF @TransaccionInicial=0 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
  ELSE IF @TransaccionInicial>0 AND XACT_STATE()=1 ROLLBACK TRANSACTION @PuntoOperacionClever;
  DECLARE @ErrorOperacionClever nvarchar(2048);SET @ErrorOperacionClever=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorOperacionClever);
 END CATCH
END
GO
CREATE PROCEDURE dbo.ConsultarOrigenesInforme @Actor int
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.ValidarActorInforme @Actor;IF @P<>0 RETURN;
 CREATE TABLE #P(CodMod nvarchar(6),Acceso int);INSERT #P EXEC dbo.ObtenerAccesosUsuario @Actor;
 SELECT m.Item,m.IdPadre,CASE WHEN m.Vigente=1 THEN ISNULL(p.Acceso,0)ELSE 0 END Acceso INTO #M FROM dbo.Int_ModulosUsuario m LEFT JOIN #P p ON p.CodMod=m.CodMod;
 ;WITH Accesos AS(
 SELECT o.Origen,m.Item,m.IdPadre,m.Acceso FROM dbo.OrigenesInforme()o JOIN #M m ON m.Item=o.Permiso
 UNION ALL SELECT a.Origen,m.Item,m.IdPadre,m.Acceso FROM Accesos a JOIN #M m ON m.Item=a.IdPadre
 )SELECT o.Origen,o.Nombre FROM dbo.OrigenesInforme()o JOIN(SELECT Origen FROM Accesos GROUP BY Origen HAVING MIN(Acceso)=1)a ON a.Origen=o.Origen ORDER BY o.Origen;
END
GO
CREATE PROCEDURE dbo.CambiarVigenciaCargo @Actor int,@Registro int,@Version binary(8),@Activo bit
AS
BEGIN
 SET NOCOUNT ON;SET XACT_ABORT ON;DECLARE @Mensaje nvarchar(2048);
 BEGIN TRY
 EXEC dbo.VerificarActorOperacion @Actor,19;
 BEGIN TRANSACTION;
 DECLARE @Serie int;SELECT @Serie=Numero FROM dbo.SeriesCargos WITH(UPDLOCK,HOLDLOCK) WHERE Serie='CRS';
 IF NOT EXISTS(SELECT 1 FROM dbo.RegistrosCargos WITH(UPDLOCK,HOLDLOCK) WHERE Item=@Registro AND VersionFila=@Version AND Activo<>@Activo) RAISERROR('El registro cambió o ya tiene esa situación. Recargue.',16,1);
 UPDATE dbo.RegistrosCargos SET Activo=@Activo WHERE Item=@Registro;
 INSERT dbo.EventosRegistrosCargos(IdRegistro,IdUsuario,Activo)VALUES(@Registro,@Actor,@Activo);
 UPDATE dbo.DbHisFac SET Activo=@Activo WHERE IdRegistroCargo=@Registro;
 DECLARE @Tipo nvarchar(2),@Id int;
 DECLARE Documentos CURSOR LOCAL FAST_FORWARD FOR SELECT Tipo,IdDocumento FROM dbo.DetallesCargos WHERE IdRegistro=@Registro ORDER BY Tipo,IdDocumento;
 OPEN Documentos;FETCH NEXT FROM Documentos INTO @Tipo,@Id;
 WHILE @@FETCH_STATUS=0 BEGIN EXEC dbo.RecalcularEstadoCargo @Tipo,@Id;FETCH NEXT FROM Documentos INTO @Tipo,@Id;END;
 CLOSE Documentos;DEALLOCATE Documentos;
 COMMIT;SELECT Item,Activo,VersionFila FROM dbo.RegistrosCargos WHERE Item=@Registro;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;SET @Mensaje=ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);END CATCH;
END
GO
CREATE PROCEDURE dbo.GuardarPizarraRP @Actor int,@Accion nvarchar(20),@Datos xml
AS
BEGIN
 SET NOCOUNT ON;SET XACT_ABORT ON;
 IF @@TRANCOUNT<>0 BEGIN RAISERROR('La operación requiere una transacción propia.',16,1);RETURN;END;
 DECLARE @Modulo int,@Mensaje nvarchar(2048),@Item int,@Version binary(8),@IdGuia int,@Asient nvarchar(20),@Planilla numeric(18,0),@Cantidad int;
 BEGIN TRY
 BEGIN TRAN;
 SET @Modulo=CASE WHEN @Accion='Reactivar' THEN 10 ELSE 9 END;
 EXEC dbo.VerificarActorOperacion @Actor,@Modulo;
 IF @Accion NOT IN('Guardar','Salida','Evento','Confirmar','Cerrar','Reactivar')RAISERROR('Acción desconocida.',16,1);
 CREATE TABLE #Elegidas(Item int PRIMARY KEY,Version binary(8));
 INSERT #Elegidas SELECT n.value('@Item','int'),CONVERT(binary(8),n.value('@Version','varchar(18)'),1) FROM @Datos.nodes('/operacion/r')x(n);
 SET @Cantidad=@@ROWCOUNT;
 IF @Cantidad=0 RAISERROR('Seleccione al menos una guía.',16,1);
 IF @Accion NOT IN('Confirmar','Cerrar') AND @Cantidad<>1 RAISERROR('La acción requiere una sola guía.',16,1);
 IF @Accion='Reactivar'BEGIN
 SELECT @Item=c.ITEM,@Version=c.VersionFila,@IdGuia=c.IdGuia,@Asient=c.Asient FROM dbo.DbsCompRp c WITH(UPDLOCK,HOLDLOCK)JOIN #Elegidas x ON x.Item=c.ITEM AND x.Version=c.VersionFila WHERE NULLIF(c.FINAL,'')IS NOT NULL;
 IF @Item IS NULL RAISERROR('El registro cambió o no está confirmado.',16,1);
 IF EXISTS(SELECT 1 FROM dbo.DbsDiarRp WITH(UPDLOCK,HOLDLOCK)WHERE IdGuia=@IdGuia OR Asient=@Asient)RAISERROR('La guía ya está en el Diario.',16,1);
 INSERT dbo.DbsDiarRp(Asient,NroGui,Client,CodCli,IdCliente,PesFis,PesVol,MonNet,MonRc,MonIgv,MonTot,Moneda,PuGui,Origen,TipGui,Info,Salida,Reten,IdGuia)
 SELECT Asient,NroGui,Client,CodCli,IdCliente,PesFis,PesVol,MonNet,MonRc,MonIgv,MonTot,Moneda,PuGui,Origen,TipGui,Info,0,0,IdGuia FROM dbo.DbsCompRp WHERE ITEM=@Item;
 DECLARE @NuevoDiario int;SET @NuevoDiario=SCOPE_IDENTITY();
 UPDATE dbo.DbsHisRP SET IdDiario=@NuevoDiario,IdCompilado=NULL WHERE IdCompilado=@Item;
 UPDATE dbo.DbsCompRp SET NroPla=NULL,FecPla=NULL,FINAL=NULL WHERE ITEM=@Item;
 INSERT dbo.DbsHisRP(HERENCIA,STATUS,INFO,FECHA,IdDiario,IdCompilado)VALUES(@Asient,'SYS','Reactivado por usuario '+CONVERT(varchar(12),@Actor),GETDATE(),@NuevoDiario,NULL);
 END ELSE BEGIN
 IF(SELECT COUNT(*)FROM dbo.DbsDiarRp d WITH(UPDLOCK,HOLDLOCK)JOIN #Elegidas x ON x.Item=d.ITEM AND x.Version=d.VersionFila)<>@Cantidad RAISERROR('Una guía cambió o salió del Diario. Recargue.',16,1);
 SELECT @Item=d.ITEM,@Asient=d.Asient FROM dbo.DbsDiarRp d JOIN #Elegidas x ON x.Item=d.ITEM;
 IF @Accion='Salida' BEGIN
 DECLARE @Salida bit;SET @Salida=@Datos.value('(/operacion/@Salida)[1]','bit');
 UPDATE dbo.DbsDiarRp SET Salida=@Salida WHERE ITEM=@Item;
 INSERT dbo.DbsHisRP(HERENCIA,STATUS,INFO,FECHA,IdDiario)VALUES(@Asient,'SYS',CASE WHEN @Salida=1 THEN 'Salida activada' ELSE 'Salida desactivada' END,GETDATE(),@Item);
 END;
 IF @Accion='Evento' BEGIN
 DECLARE @Estado int,@Comentario nvarchar(1000),@CodigoEstado nvarchar(5);
 SELECT @Estado=@Datos.value('(/operacion/@Estado)[1]','int'),@Comentario=@Datos.value('(/operacion/@Comentario)[1]','nvarchar(1000)');
 SELECT @CodigoEstado=Status FROM dbo.Int_StatusRP WHERE CodEst=@Estado AND ActEst=1;
 IF @CodigoEstado IS NULL OR NULLIF(LTRIM(RTRIM(@Comentario)),'')IS NULL RAISERROR('Seleccione estado activo y escriba el evento.',16,1);
 UPDATE dbo.DbsDiarRp SET Estado=@Estado WHERE ITEM=@Item;
 INSERT dbo.DbsHisRP(HERENCIA,STATUS,INFO,FECHA,IdDiario)VALUES(@Asient,@CodigoEstado,@Comentario,GETDATE(),@Item);
 END;
 IF @Accion='Guardar'BEGIN
 DECLARE @Cliente int,@PF decimal(18,4),@PV decimal(18,4),@Neto money,@Rc money,@Igv money,@Total money,@Moneda tinyint,@Info nvarchar(1000);
 SELECT @Cliente=@Datos.value('(/operacion/@Cliente)[1]','int'),@PF=@Datos.value('(/operacion/@PesoFisico)[1]','decimal(18,4)'),@PV=@Datos.value('(/operacion/@PesoVolumen)[1]','decimal(18,4)'),@Neto=@Datos.value('(/operacion/@Neto)[1]','money'),@Rc=@Datos.value('(/operacion/@Recargo)[1]','money'),@Igv=@Datos.value('(/operacion/@IGV)[1]','money'),@Total=@Datos.value('(/operacion/@Total)[1]','money'),@Moneda=@Datos.value('(/operacion/@Moneda)[1]','tinyint'),@Info=@Datos.value('(/operacion/@Info)[1]','nvarchar(1000)');
 IF @PF<0 OR @PV<0 OR @Neto<0 OR @Rc<0 OR @Igv<0 OR @Total<0 OR @Total<>@Neto+@Rc+@Igv RAISERROR('Pesos/importes negativos o total diferente de neto + recargo + IGV.',16,1);
 IF NOT EXISTS(SELECT 1 FROM dbo.DbClientes WHERE Item=@Cliente AND Activo=1)OR NOT EXISTS(SELECT 1 FROM dbo.Int_Moneda WHERE CodEst=@Moneda AND ActEst=1)RAISERROR('Cliente o moneda inválidos.',16,1);
 UPDATE d SET Client=c.CLIENT,CodCli=c.Asient,IdCliente=c.Item,PesFis=@PF,PesVol=@PV,MonNet=@Neto,MonRc=@Rc,MonIgv=@Igv,MonTot=@Total,Moneda=@Moneda,Info=@Info FROM dbo.DbsDiarRp d CROSS JOIN dbo.DbClientes c WHERE d.ITEM=@Item AND c.Item=@Cliente;
 END;
 IF @Accion IN('Confirmar','Cerrar') BEGIN
 IF(SELECT COUNT(*) FROM dbo.DbsCompRp c WITH(UPDLOCK,HOLDLOCK)JOIN dbo.DbsDiarRp d ON d.IdGuia=c.IdGuia JOIN #Elegidas x ON x.Item=d.ITEM WHERE NULLIF(c.FINAL,'')IS NULL)<>@Cantidad RAISERROR('Falta el registro pendiente del compilatorio o ya fue confirmado.',16,1);
 IF @Accion='Confirmar' AND EXISTS(SELECT 1 FROM dbo.DbsDiarRp d JOIN #Elegidas x ON x.Item=d.ITEM WHERE ISNULL(d.Salida,0)<>1 OR d.IdCliente IS NULL OR d.MonTot<=0 OR d.Moneda IS NULL)RAISERROR('Para generar cartilla: salida activa, cliente, moneda y total positivo.',16,1);
 IF @Accion='Cerrar' AND EXISTS(SELECT 1 FROM dbo.DbsDiarRp d JOIN #Elegidas x ON x.Item=d.ITEM LEFT JOIN dbo.Int_StatusRP s ON s.CodEst=d.Estado WHERE ISNULL(s.Status,'')NOT IN('COT','OTR','TNT'))RAISERROR('Solo se cierran guías con estado COT, OTR o TNT.',16,1);
 IF @Accion='Confirmar' BEGIN EXEC dbo.ReservarCorrelativo 'CRP',@Planilla OUTPUT;INSERT dbo.CartillasRP(Numero,IdUsuario)VALUES(@Planilla,@Actor);END;
 UPDATE c SET NroPla=CASE WHEN @Accion='Cerrar'THEN 'Finalizado' ELSE CONVERT(nvarchar(20),@Planilla)END,FecPla=GETDATE(),Client=d.Client,CodCli=d.CodCli,IdCliente=d.IdCliente,PesFis=d.PesFis,PesVol=d.PesVol,MonNet=d.MonNet,MonRc=d.MonRc,MonIgv=d.MonIgv,MonTot=d.MonTot,Moneda=d.Moneda,Info=d.Info,FINAL=CASE WHEN @Accion='Cerrar'THEN 'Finalizado' ELSE 'Facturado' END FROM dbo.DbsCompRp c JOIN dbo.DbsDiarRp d ON d.IdGuia=c.IdGuia JOIN #Elegidas x ON x.Item=d.ITEM;
 IF @Accion='Confirmar' INSERT dbo.DetallesCartillaRP(Planilla,IdGuia,NroGui,Cliente,Territorio,PesFis,PesVol,Neto,Recargo,IGV,Total,Moneda,FechaEnvio,Producto)
 SELECT @Planilla,d.IdGuia,d.NroGui,d.Client,t.Descrip,d.PesFis,d.PesVol,d.MonNet,d.MonRc,d.MonIgv,d.MonTot,m.Moneda,d.PuGui,d.TipGui FROM dbo.DbsDiarRp d JOIN #Elegidas x ON x.Item=d.ITEM LEFT JOIN dbo.DbClientes cl ON cl.Item=d.IdCliente LEFT JOIN dbo.IntTerritorios t ON t.Item=cl.IdTerritorio LEFT JOIN dbo.Int_Moneda m ON m.CodEst=d.Moneda;
 UPDATE h SET IdDiario=NULL,IdCompilado=c.ITEM FROM dbo.DbsHisRP h JOIN dbo.DbsDiarRp d ON d.ITEM=h.IdDiario JOIN #Elegidas x ON x.Item=d.ITEM JOIN dbo.DbsCompRp c ON c.IdGuia=d.IdGuia;
 INSERT dbo.DbsHisRP(HERENCIA,STATUS,INFO,FECHA,IdCompilado) SELECT d.Asient,'SYS',CASE WHEN @Accion='Cerrar'THEN 'Cierre de gestión' ELSE 'Cartilla '+CONVERT(nvarchar(20),@Planilla)+' preparada para facturación externa' END,GETDATE(),c.ITEM FROM dbo.DbsDiarRp d JOIN #Elegidas x ON x.Item=d.ITEM JOIN dbo.DbsCompRp c ON c.IdGuia=d.IdGuia;
 DELETE d FROM dbo.DbsDiarRp d JOIN #Elegidas x ON x.Item=d.ITEM;
 END;
 END;
 COMMIT;
 SELECT Planilla=@Planilla,Registros=@Cantidad;
 END TRY BEGIN CATCH
 IF XACT_STATE()<>0 ROLLBACK;
 SET @Mensaje=ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);
 END CATCH
END
GO
CREATE PROCEDURE dbo.RegistrarCargos @Actor int,@IdEstado int,@Fecha datetime,@Documentos xml
AS
BEGIN
 SET NOCOUNT ON;SET XACT_ABORT ON;
 DECLARE @Mensaje nvarchar(2048);
 BEGIN TRY
 EXEC dbo.VerificarActorOperacion @Actor,19;
 DECLARE @Lote TABLE(Tipo nvarchar(2),Id int,Version binary(8),PRIMARY KEY(Tipo,Id));
 INSERT @Lote SELECT n.value('@Tipo','nvarchar(2)'),n.value('@Id','int'),CONVERT(binary(8),n.value('@Version','varchar(18)'),1) FROM @Documentos.nodes('/documentos/d')t(n);
 IF NOT EXISTS(SELECT 1 FROM @Lote) OR (SELECT COUNT(*)FROM @Lote)>200 RAISERROR('Ingrese entre 1 y 200 documentos.',16,1);
 IF @Fecha IS NULL OR @Fecha<'19000101' OR @Fecha>=DATEADD(day,1,CONVERT(date,GETDATE())) RAISERROR('Fecha de gestión inválida.',16,1);
 BEGIN TRANSACTION;
 DECLARE @Numero int,@Registro int,@Titulo nvarchar(100),@Correlativo int,@ActCor tinyint;
 SELECT @Numero=Numero+1 FROM dbo.SeriesCargos WITH(UPDLOCK,HOLDLOCK) WHERE Serie='CRS';
 SELECT @Titulo=TitEst,@ActCor=ActCor,@Correlativo=CONVERT(int,ISNULL(Correl,0))+1 FROM dbo.Int_StatusFac WITH(UPDLOCK,HOLDLOCK) WHERE ITEM=@IdEstado AND ActEst=1;
 IF @Titulo IS NULL OR @Numero IS NULL RAISERROR('Estado o serie no disponibles.',16,1);
 -- Bloqueo de documentos y comprobación de la versión cargada en cada banda.
 DECLARE @Bloqueo int;
 SELECT @Bloqueo=d.Item FROM dbo.DBEmiFac d WITH(UPDLOCK,HOLDLOCK) JOIN @Lote l ON l.Tipo<>'NC' AND d.Item=l.Id;
 SELECT @Bloqueo=n.ITEM FROM dbo.DbNCredito n WITH(UPDLOCK,HOLDLOCK) JOIN @Lote l ON l.Tipo='NC' AND n.ITEM=l.Id;
 IF EXISTS(SELECT 1 FROM @Lote l LEFT JOIN dbo.DocumentosGestion d ON d.Tipo=l.Tipo AND d.IdDocumento=l.Id
 WHERE d.IdDocumento IS NULL OR d.VersionFila<>l.Version OR d.Anulado=1) RAISERROR('Un documento cambió, está anulado o no existe. Recargue la banda.',16,1);
 IF EXISTS(SELECT 1 FROM @Lote l JOIN dbo.DocumentosGestion d ON d.Tipo=l.Tipo AND d.IdDocumento=l.Id WHERE @Fecha<CONVERT(date,d.Emision)) RAISERROR('La gestión no puede ser anterior a la emisión.',16,1);
 UPDATE dbo.SeriesCargos SET Numero=@Numero WHERE Serie='CRS';
 IF @ActCor=1 UPDATE dbo.Int_StatusFac SET Correl=@Correlativo WHERE ITEM=@IdEstado;
 ELSE SET @Correlativo=NULL;
 INSERT dbo.RegistrosCargos(Numero,FechaEstado,IdEstado,IdUsuario,Titulo,Correlativo)VALUES(@Numero,@Fecha,@IdEstado,@Actor,@Titulo,@Correlativo);
 SET @Registro=SCOPE_IDENTITY();
 INSERT dbo.EventosRegistrosCargos(IdRegistro,IdUsuario,Activo)VALUES(@Registro,@Actor,1);
 INSERT dbo.EstadosInicialesCargos(Tipo,IdDocumento,IdEstado,Estado,FechaEstado)
 SELECT d.Tipo,d.IdDocumento,d.IdEstado,CASE WHEN d.Tipo='NC' THEN n.Estado ELSE f.Estado END,d.FechaEstado
 FROM @Lote l JOIN dbo.DocumentosGestion d ON d.Tipo=l.Tipo AND d.IdDocumento=l.Id
 LEFT JOIN dbo.DBEmiFac f ON d.Tipo<>'NC' AND f.Item=d.IdDocumento LEFT JOIN dbo.DbNCredito n ON d.Tipo='NC' AND n.ITEM=d.IdDocumento
 WHERE NOT EXISTS(SELECT 1 FROM dbo.EstadosInicialesCargos e WITH(UPDLOCK,HOLDLOCK) WHERE e.Tipo=d.Tipo AND e.IdDocumento=d.IdDocumento);
 INSERT dbo.DetallesCargos(IdRegistro,Tipo,IdDocumento,IdFactura,IdNotaCredito,Documento,Emision,Cliente,Divisa,Importe,Situacion,Guias,Servicios,ContactoEntrega,DireccionEntrega,DistritoEntrega,TelefonoEntrega)
 SELECT @Registro,d.Tipo,d.IdDocumento,CASE WHEN d.Tipo<>'NC' THEN d.IdDocumento END,CASE WHEN d.Tipo='NC' THEN d.IdDocumento END,d.Documento,d.Emision,d.Cliente,d.Divisa,d.Importe,d.Situacion,
 ISNULL(STUFF((SELECT N', '+g.NroGui FROM dbo.DbServGui g WHERE(d.Tipo='NC' AND g.IdNotaCredito=d.IdDocumento)OR(d.Tipo<>'NC' AND g.IdFactura=d.IdDocumento)ORDER BY g.Linea,g.Item FOR XML PATH(''),TYPE).value('.','nvarchar(max)'),1,2,N''),N''),
 ISNULL(STUFF((SELECT N', '+g.Servic FROM dbo.DbServGui g WHERE(d.Tipo='NC' AND g.IdNotaCredito=d.IdDocumento)OR(d.Tipo<>'NC' AND g.IdFactura=d.IdDocumento)ORDER BY g.Linea,g.Item FOR XML PATH(''),TYPE).value('.','nvarchar(max)'),1,2,N''),N''),
 d.ContactoEntrega,d.DireccionEntrega,d.DistritoEntrega,d.TelefonoEntrega
 FROM @Lote l JOIN dbo.DocumentosGestion d ON d.Tipo=l.Tipo AND d.IdDocumento=l.Id;
 INSERT dbo.DbHisFac(TipDoc,NumDoc,Estado,FecEst,NomUsu,IdUsuario,IdEstado,IdFactura,IdNotaCredito,IdRegistroCargo,Activo)
 SELECT t.CodIte,d.Documento,s.Status,@Fecha,u.Usuari,@Actor,@IdEstado,d.IdFactura,d.IdNotaCredito,@Registro,1
 FROM dbo.DetallesCargos d JOIN dbo.Int_TipDoc t ON t.TipDoc=d.Tipo JOIN dbo.Int_StatusFac s ON s.ITEM=@IdEstado JOIN dbo.Int_Usuario u ON u.ITEM=@Actor WHERE d.IdRegistro=@Registro;
 DECLARE @Tipo nvarchar(2),@Id int;
 DECLARE Documentos CURSOR LOCAL FAST_FORWARD FOR SELECT Tipo,Id FROM @Lote ORDER BY Tipo,Id;
 OPEN Documentos;FETCH NEXT FROM Documentos INTO @Tipo,@Id;
 WHILE @@FETCH_STATUS=0 BEGIN EXEC dbo.RecalcularEstadoCargo @Tipo,@Id;FETCH NEXT FROM Documentos INTO @Tipo,@Id;END;
 CLOSE Documentos;DEALLOCATE Documentos;
 COMMIT;
 SELECT Item,Numero,Correlativo,VersionFila FROM dbo.RegistrosCargos WHERE Item=@Registro;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;SET @Mensaje=ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);END CATCH;
END
GO
CREATE PROCEDURE dbo.GuardarGruposRP @Actor int,@Cambios xml
AS
BEGIN
 SET NOCOUNT ON;SET XACT_ABORT ON;
 IF @@TRANCOUNT<>0 BEGIN RAISERROR('La operación requiere su propia transacción.',16,1);RETURN;END;
 DECLARE @Mensaje nvarchar(2048),@Agregados int,@Quitados int;
 BEGIN TRY
 BEGIN TRAN;
 EXEC dbo.VerificarActorOperacion @Actor,9;
 DECLARE @Bloqueo int;
 SELECT @Bloqueo=Item FROM dbo.Int_DivRp WITH(UPDLOCK,HOLDLOCK)WHERE CodCli='0000000000';
 IF @Bloqueo IS NULL RAISERROR('Falta el grupo general.',16,1);
 CREATE TABLE #Agregar(IdCliente int PRIMARY KEY);
 CREATE TABLE #Quitar(Item int PRIMARY KEY,Version binary(8)NOT NULL);
 INSERT #Agregar SELECT n.value('@Cliente','int')FROM @Cambios.nodes('/cambios/agregar')x(n);
 INSERT #Quitar SELECT n.value('@Item','int'),CONVERT(binary(8),n.value('@Version','varchar(18)'),1)FROM @Cambios.nodes('/cambios/quitar')x(n);
 IF NOT EXISTS(SELECT 1 FROM #Agregar)AND NOT EXISTS(SELECT 1 FROM #Quitar)RAISERROR('No hay cambios para guardar.',16,1);
 IF(SELECT COUNT(*)FROM dbo.Int_DivRp g JOIN #Quitar q ON q.Item=g.Item AND q.Version=g.VersionFila WHERE g.CodCli<>'0000000000')<>(SELECT COUNT(*)FROM #Quitar)RAISERROR('Un grupo cambió o ya fue eliminado. Recargue antes de guardar.',16,1);
 IF EXISTS(SELECT 1 FROM #Agregar a LEFT JOIN dbo.DbClientes c WITH(HOLDLOCK)ON c.Item=a.IdCliente WHERE c.Item IS NULL OR c.Activo<>1)RAISERROR('Seleccione un cliente existente y activo.',16,1);
 DELETE g FROM dbo.Int_DivRp g JOIN #Quitar q ON q.Item=g.Item;
 SET @Quitados=@@ROWCOUNT;
 IF EXISTS(SELECT 1 FROM #Agregar a JOIN dbo.Int_DivRp g ON g.IdCliente=a.IdCliente)RAISERROR('El cliente ya tiene un grupo. Recargue la lista.',16,1);
 INSERT dbo.Int_DivRp(NomCli,CodCli,DesCli,IdCliente)SELECT c.CLIENT,c.Asient,N'Grupo separado del Diario RP',c.Item FROM dbo.DbClientes c JOIN #Agregar a ON a.IdCliente=c.Item;
 SET @Agregados=@@ROWCOUNT;
 COMMIT;
 SELECT Agregados=@Agregados,Quitados=@Quitados;
 END TRY BEGIN CATCH
 IF XACT_STATE()<>0 ROLLBACK;
 SET @Mensaje=ERROR_MESSAGE();RAISERROR('%s',16,1,@Mensaje);
 END CATCH
END
GO
CREATE PROCEDURE dbo.GenerarInforme @Actor int,@Origen int,@Solicitud xml
AS
BEGIN SET NOCOUNT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.ValidarActorInforme @Actor,@Origen;IF @P<>0 RETURN;
 IF @Solicitud IS NULL OR @Solicitud.value('count(/informe)','int')<>1 RAISERROR('Defina las columnas del informe.',16,1);
 IF @Solicitud.exist('/informe/*[not(self::c) and not(self::f)]')=1 RAISERROR('Estructura de informe inválida.',16,1);
 IF @Solicitud.value('count(/informe/c)','int')NOT BETWEEN 1 AND 50 OR @Solicitud.value('count(/informe/f)','int')>4 RAISERROR('Seleccione columnas y hasta cuatro filtros.',16,1);
 IF @Solicitud.exist('/informe/*[string-length(@Campo)>128 or string-length(@Valor)>3900]')=1 RAISERROR('Campo o valor demasiado largo.',16,1);
 DECLARE @Vista sysname,@Select nvarchar(max),@Where nvarchar(max),@SQL nvarchar(max),@Campo sysname,@Tipo sysname,@Op int,@Dia bit,@O bit,@I int,@N int,@Cond nvarchar(max),@Nombre nvarchar(20),@Valor nvarchar(4000),@Fecha datetime;
 SELECT @Vista=Vista FROM dbo.OrigenesInforme()WHERE Origen=@Origen;
 DECLARE @C TABLE(Orden int PRIMARY KEY,Campo sysname UNIQUE);
 INSERT @C SELECT x.n.value('@Orden','int'),x.n.value('@Campo','nvarchar(128)')FROM @Solicitud.nodes('/informe/c')x(n);
 IF EXISTS(SELECT 1 FROM @C c LEFT JOIN sys.columns m ON m.object_id=OBJECT_ID(N'dbo.'+@Vista)AND m.name=c.Campo WHERE m.column_id IS NULL)RAISERROR('Columna ajena al origen seleccionado.',16,1);
 IF(SELECT MIN(Orden)FROM @C)<>1 OR(SELECT MAX(Orden)FROM @C)<>(SELECT COUNT(*)FROM @C)RAISERROR('Orden de columnas inválido.',16,1);
 SET @Select=N'';SET @I=1;SET @N=(SELECT COUNT(*)FROM @C);
 WHILE @I<=@N BEGIN SELECT @Campo=Campo FROM @C WHERE Orden=@I;SET @Select=@Select+CASE WHEN @I=1 THEN N''ELSE N','END+QUOTENAME(@Campo);SET @I=@I+1;END;
 DECLARE @F TABLE(Orden int PRIMARY KEY,Campo sysname,Operador int,Valor nvarchar(4000),Dia bit,O bit);
 INSERT @F SELECT x.n.value('@Orden','int'),x.n.value('@Campo','nvarchar(128)'),x.n.value('@Operador','int'),x.n.value('@Valor','nvarchar(4000)'),x.n.value('@Dia','bit'),x.n.value('@O','bit')FROM @Solicitud.nodes('/informe/f')x(n);
 IF EXISTS(SELECT 1 FROM @F f LEFT JOIN sys.columns c ON c.object_id=OBJECT_ID(N'dbo.'+@Vista)AND c.name=f.Campo WHERE c.column_id IS NULL OR f.Operador NOT BETWEEN 1 AND 12)RAISERROR('Campo o condición de filtro inválidos.',16,1);
 IF EXISTS(SELECT 1 FROM @F)AND((SELECT MIN(Orden)FROM @F)<>1 OR(SELECT MAX(Orden)FROM @F)<>(SELECT COUNT(*)FROM @F))RAISERROR('Orden de filtros inválido.',16,1);
 DECLARE @V1 nvarchar(4000),@V2 nvarchar(4000),@V3 nvarchar(4000),@V4 nvarchar(4000),@AnteriorO bit;
 SET @Where=N'';SET @I=1;SET @N=(SELECT COUNT(*)FROM @F);
 WHILE @I<=@N BEGIN
  SELECT @Campo=f.Campo,@Op=f.Operador,@Valor=f.Valor,@Dia=f.Dia,@O=f.O,@Tipo=TYPE_NAME(c.system_type_id)FROM @F f JOIN sys.columns c ON c.object_id=OBJECT_ID(N'dbo.'+@Vista)AND c.name=f.Campo WHERE f.Orden=@I;
  SET @Nombre=N'@V'+CONVERT(nvarchar(1),@I);
  IF @Op>=7 BEGIN
   DECLARE @Patron nvarchar(max);SET @Patron=REPLACE(REPLACE(REPLACE(CONVERT(nvarchar(max),@Valor),N'[',N'[[]'),N'%',N'[%]'),N'_',N'[_]');
   IF LEN(@Patron)>3998 RAISERROR('Patrón de filtro demasiado largo.',16,1);SET @Valor=@Patron;
   IF @Op IN(9,10,11,12)SET @Valor=N'%'+@Valor;
   IF @Op IN(7,8,11,12)SET @Valor=@Valor+N'%';
   SET @Cond=N'CONVERT(nvarchar(4000),'+QUOTENAME(@Campo)+N')'+CASE WHEN @Op%2=0 THEN N' NOT LIKE 'ELSE N' LIKE 'END+@Nombre;
  END ELSE BEGIN
   SET @Cond=QUOTENAME(@Campo)+CASE @Op WHEN 1 THEN N'='WHEN 2 THEN N'<>'WHEN 3 THEN N'>'WHEN 4 THEN N'>='WHEN 5 THEN N'<'ELSE N'<='END;
   IF @Tipo IN('date','datetime','smalldatetime','datetime2')BEGIN
    IF @Valor NOT LIKE '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]T[0-9][0-9]:[0-9][0-9]:[0-9][0-9]'RAISERROR('Fecha de filtro inválida.',16,1);
    SET @Fecha=CONVERT(datetime,@Valor,126);
    IF @Dia=1 BEGIN
     IF CONVERT(date,@Fecha)='99991231' BEGIN
      SET @Cond=CASE WHEN @Op IN(1,4)THEN QUOTENAME(@Campo)+N'>=CONVERT(datetime,'+@Nombre+N',126)'WHEN @Op IN(2,5)THEN QUOTENAME(@Campo)+N'<CONVERT(datetime,'+@Nombre+N',126)'WHEN @Op=3 THEN N'1=0'ELSE QUOTENAME(@Campo)+N' IS NOT NULL'END;
     END ELSE SET @Cond=CASE @Op
      WHEN 1 THEN N'('+QUOTENAME(@Campo)+N'>=CONVERT(datetime,'+@Nombre+N',126) AND '+QUOTENAME(@Campo)+N'<DATEADD(day,1,CONVERT(datetime,'+@Nombre+N',126)))'
      WHEN 2 THEN N'('+QUOTENAME(@Campo)+N'<CONVERT(datetime,'+@Nombre+N',126) OR '+QUOTENAME(@Campo)+N'>=DATEADD(day,1,CONVERT(datetime,'+@Nombre+N',126)))'
      WHEN 3 THEN QUOTENAME(@Campo)+N'>=DATEADD(day,1,CONVERT(datetime,'+@Nombre+N',126))'
      WHEN 6 THEN QUOTENAME(@Campo)+N'<DATEADD(day,1,CONVERT(datetime,'+@Nombre+N',126))'
      ELSE @Cond+N'CONVERT(datetime,'+@Nombre+N',126)'END;
    END ELSE SET @Cond=@Cond+N'CONVERT(datetime,'+@Nombre+N',126)';
   END ELSE IF @Tipo IN('int','bigint','smallint','tinyint','numeric','decimal','money','smallmoney','real','float','bit')BEGIN
    IF NULLIF(@Valor,N'')IS NULL OR ISNUMERIC(@Valor)=0 RAISERROR('Valor numérico inválido.',16,1);
    -- Contrato decimal explícito: nunca redondear un límite introducido.
    DECLARE @Digitos nvarchar(4000),@Punto int;SET @Digitos=@Valor;
    IF LEFT(@Digitos,1)IN(N'+',N'-')SET @Digitos=SUBSTRING(@Digitos,2,4000);
    IF @Digitos=N'' OR @Digitos COLLATE Latin1_General_BIN LIKE N'%[^0-9.]%' OR LEN(@Digitos)-LEN(REPLACE(@Digitos,N'.',N''))>1 OR REPLACE(@Digitos,N'.',N'')=N'' RAISERROR('Valor numérico inválido. Use notación decimal.',16,1);
    SET @Punto=CHARINDEX(N'.',@Digitos);
    IF @Punto>0 AND LEN(@Digitos)-@Punto>6 RAISERROR('El filtro numérico admite hasta seis decimales. Revise el valor; no se redondeó.',16,1);
    SET @Fecha=NULL;DECLARE @Numero decimal(28,6);SET @Numero=CONVERT(decimal(28,6),@Valor);
    SET @Cond=@Cond+N'CONVERT(decimal(28,6),'+@Nombre+N')';
   END ELSE SET @Cond=@Cond+@Nombre;
  END;
  IF @I=1 SET @Where=@Cond;ELSE SET @Where=N'('+@Where+CASE WHEN @AnteriorO=1 THEN N' OR 'ELSE N' AND 'END+@Cond+N')';
  IF @I=1 SET @V1=@Valor;ELSE IF @I=2 SET @V2=@Valor;ELSE IF @I=3 SET @V3=@Valor;ELSE SET @V4=@Valor;
  SET @AnteriorO=@O;SET @I=@I+1;
 END;
 SET @SQL=N'SELECT TOP(10001) '+@Select+N' FROM dbo.'+QUOTENAME(@Vista)+CASE WHEN @Where=N''THEN N''ELSE N' WHERE '+@Where END+N' ORDER BY Codigo';
 EXEC sp_executesql @SQL,N'@V1 nvarchar(4000),@V2 nvarchar(4000),@V3 nvarchar(4000),@V4 nvarchar(4000)',@V1,@V2,@V3,@V4;
 END TRY BEGIN CATCH DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.ConsultarDetalleCargo @Actor int,@Registro int
AS
BEGIN SET NOCOUNT ON;BEGIN TRY EXEC dbo.VerificarActorOperacion @Actor,19; END TRY
 BEGIN CATCH DECLARE @ErrorAcceso nvarchar(2048);SET @ErrorAcceso=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorAcceso);RETURN;END CATCH;
 SELECT d.*,r.FechaEstado,r.Titulo,r.Numero,r.Correlativo,r.Activo FROM dbo.DetallesCargos d JOIN dbo.RegistrosCargos r ON r.Item=d.IdRegistro WHERE r.Item=@Registro ORDER BY d.Item;END
GO
CREATE PROCEDURE dbo.ConsultarGruposRP @Actor int,@IncluyeGeneral bit=0
AS
BEGIN
 SET NOCOUNT ON;
 EXEC dbo.VerificarActorOperacion @Actor,9;
 SELECT g.Item,g.CodCli,g.IdCliente,COALESCE(c.CLIENT,g.NomCli)Nombre,g.VersionFila,
 (SELECT COUNT(*)FROM dbo.DbsDiarRp d WHERE d.IdCliente=g.IdCliente)Guias
 FROM dbo.Int_DivRp g LEFT JOIN dbo.DbClientes c ON c.Item=g.IdCliente
 WHERE @IncluyeGeneral=1 OR g.CodCli<>'0000000000' ORDER BY CASE WHEN g.CodCli='0000000000'THEN 0 ELSE 1 END,Nombre,g.Item;
END
GO
CREATE PROCEDURE dbo.ConsultarRegistrosCargos @Actor int,@Numero int=0
AS
BEGIN SET NOCOUNT ON;BEGIN TRY EXEC dbo.VerificarActorOperacion @Actor,19; END TRY
 BEGIN CATCH DECLARE @ErrorAcceso nvarchar(2048);SET @ErrorAcceso=ERROR_MESSAGE();RAISERROR('%s',16,1,@ErrorAcceso);RETURN;END CATCH;
 SELECT TOP(50) r.Item,r.Numero,r.FechaEstado,r.Titulo,r.Correlativo,r.Activo,u.Usuari Usuario,r.VersionFila,
 (SELECT COUNT(*) FROM dbo.DetallesCargos d WHERE d.IdRegistro=r.Item) Documentos
 FROM dbo.RegistrosCargos r JOIN dbo.Int_Usuario u ON u.ITEM=r.IdUsuario WHERE @Numero=0 OR r.Numero=@Numero ORDER BY r.Numero DESC;END
GO
CREATE PROCEDURE dbo.AnularQuery @Actor int,@Item int,@Version binary(8),@Motivo nvarchar(1000)
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;
 BEGIN TRY DECLARE @Permiso int;EXEC @Permiso=dbo.VerificarActorQuery @Actor,6;IF @Permiso<>0 RETURN;BEGIN TRANSACTION;
 IF NULLIF(LTRIM(RTRIM(@Motivo)),'')IS NULL RAISERROR('Escriba el motivo de anulación de la solicitud.',16,1);
 UPDATE dbo.DbQuerie SET SitDoc='ANL',FechaAnulacion=GETDATE(),IdAnulador=@Actor WHERE ITEM=@Item AND VersionFila=@Version AND Estado='Abierto' AND SitDoc='ACT';
 IF @@ROWCOUNT<>1 RAISERROR('La solicitud cambió, se cerró o ya está anulada. Recargue.',16,1);
 INSERT dbo.HistorialQueries(IdQuery,IdUsuario,Accion,Detalle)VALUES(@Item,@Actor,'ANULAR',@Motivo);COMMIT;SELECT ITEM,Asient,VersionFila FROM dbo.DbQuerie WHERE ITEM=@Item;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @Error nvarchar(2048);SET @Error=ERROR_MESSAGE();RAISERROR('%s',16,1,@Error);END CATCH;
END
GO
CREATE PROCEDURE dbo.CobGuardarQueja @Actor int,@Clave uniqueidentifier,@Solicitud xml
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,12;IF @P<>0 RETURN;
 DECLARE @Caso int,@Cliente int,@Tipo int,@Estado int,@V binary(8),@Contacto nvarchar(50),@Telefono nvarchar(20),@Texto nvarchar(4000),@Usuario nvarchar(50),@Num nvarchar(10),@Raiz nvarchar(10),@Codigo nvarchar(20),@Descripcion nvarchar(150),@Id int,@CodigoCliente nvarchar(20),@TipoDescripcion nvarchar(150),@Nuevo bit;
 IF @Clave IS NULL OR @Solicitud.value('count(/gestion)','int')<>1 RAISERROR('Solicitud de gestión inválida.',16,1);
 SELECT @Caso=@Solicitud.value('(/gestion/@Caso)[1]','int'),@Cliente=@Solicitud.value('(/gestion/@Cliente)[1]','int'),@Tipo=@Solicitud.value('(/gestion/@Tipo)[1]','int'),@Estado=@Solicitud.value('(/gestion/@Estado)[1]','int'),@Contacto=@Solicitud.value('(/gestion/@Contacto)[1]','nvarchar(50)'),@Telefono=@Solicitud.value('(/gestion/@Telefono)[1]','nvarchar(20)'),@Texto=@Solicitud.value('(/gestion/@Texto)[1]','nvarchar(4000)');
 IF NULLIF(LTRIM(RTRIM(@Contacto)),'')IS NULL OR NULLIF(LTRIM(RTRIM(@Texto)),'')IS NULL RAISERROR('Complete contacto y descripción de la gestión.',16,1);
 SET @Nuevo=CASE WHEN @Caso=0 THEN 1 ELSE 0 END;
 BEGIN TRANSACTION;
 SELECT @Id=COALESCE(IdCaso,IdGestion)FROM dbo.ClavesCobranzas WITH(UPDLOCK,HOLDLOCK)WHERE Clave=@Clave;
 IF @Id IS NOT NULL BEGIN IF NOT EXISTS(SELECT 1 FROM dbo.ClavesCobranzas WHERE Clave=@Clave AND Modulo=12 AND IdUsuario=@Actor AND CONVERT(nvarchar(max),Solicitud) COLLATE Latin1_General_BIN2=CONVERT(nvarchar(max),@Solicitud) COLLATE Latin1_General_BIN2)RAISERROR('La clave ya identifica otra operación.',16,1);
 SELECT @Caso=COALESCE(k.IdCaso,g.IdCaso)FROM dbo.ClavesCobranzas k LEFT JOIN dbo.DbQuejas_Ges g ON g.ITEM=k.IdGestion WHERE k.Clave=@Clave;SELECT @Raiz=Asient FROM dbo.DbQuejas_Cas WHERE ITEM=@Caso;
 COMMIT;SELECT @Id Item,@Caso IdCaso,@Raiz Numero;RETURN;END;
 SELECT @Usuario=Usuari FROM dbo.Int_Usuario WHERE ITEM=@Actor;
 SELECT @Codigo=Status,@Descripcion=DesSta FROM dbo.Int_StatusCas WITH(HOLDLOCK)WHERE CodSta=@Estado AND ActEst=1;
 IF @Codigo IS NULL RAISERROR('Seleccione un estado activo.',16,1);
 IF @Caso=0 BEGIN
 SELECT @CodigoCliente=Asient FROM dbo.DbClientes WITH(HOLDLOCK)WHERE Item=@Cliente AND Activo=1;
 SELECT @TipoDescripcion=NomCas FROM dbo.Int_Casos WITH(HOLDLOCK)WHERE CodCas=@Tipo AND ActEst=1;
 IF @CodigoCliente IS NULL OR @TipoDescripcion IS NULL RAISERROR('Seleccione cliente y concepto activos.',16,1);
 IF NOT EXISTS(SELECT 1 FROM dbo.DbQuejas_Cli WITH(UPDLOCK,HOLDLOCK)WHERE IdCliente=@Cliente)INSERT dbo.DbQuejas_Cli(CodCli,Estado,CliTop,Usuario,FecApe,ContCas,DataPas,FecData,IdCliente)VALUES(@CodigoCliente,'Bueno',0,@Usuario,GETDATE(),0,0,GETDATE(),@Cliente);
 EXEC dbo.CobReservarNumero 'CAS',@Raiz OUTPUT;
 INSERT dbo.DbQuejas_Cas(Asient,CodCli,TipCas,FecCas,Estado,Usuari,DataPas,FecData,IdCliente,IdUsuario,DescripcionCaso)VALUES(@Raiz,@CodigoCliente,@Tipo,GETDATE(),@Codigo,@Usuario,0,GETDATE(),@Cliente,@Actor,@TipoDescripcion);SET @Caso=SCOPE_IDENTITY();SET @Id=@Caso;
 END ELSE BEGIN
 SET @V=CONVERT(binary(8),@Solicitud.value('(/gestion/@Version)[1]','varchar(18)'),1);
 SELECT @Raiz=Asient,@Cliente=IdCliente FROM dbo.DbQuejas_Cas WITH(UPDLOCK,HOLDLOCK)WHERE ITEM=@Caso AND VersionFila=@V AND FecFin IS NULL;
 IF @Raiz IS NULL RAISERROR('El caso cambió o ya fue cerrado. Recargue.',16,1);
 END;
 EXEC dbo.CobReservarNumero 'GES',@Num OUTPUT;
 INSERT dbo.DbQuejas_Ges(Asient,Raiz,Status,FecSta,Contac,Telef1,Gestion,Usuari,DataPas,FecData,IdUsuario,IdCaso,CodigoEstado,DescripcionEstado)VALUES(@Num,@Raiz,@Estado,GETDATE(),@Contacto,@Telefono,@Texto,@Usuario,0,GETDATE(),@Actor,@Caso,@Codigo,@Descripcion);
 IF @Id IS NULL SET @Id=SCOPE_IDENTITY();
 UPDATE dbo.DbQuejas_Cas SET Estado=@Codigo,FecFin=CASE WHEN @Codigo='CCR'THEN GETDATE()ELSE NULL END,UsuFin=CASE WHEN @Codigo='CCR'THEN @Usuario ELSE NULL END WHERE ITEM=@Caso;
 UPDATE dbo.DbQuejas_Cli SET ContCas=(SELECT COUNT(*)FROM dbo.DbQuejas_Cas WHERE IdCliente=@Cliente)WHERE IdCliente=@Cliente;
 INSERT dbo.ClavesCobranzas(Clave,Modulo,Accion,IdUsuario,Solicitud,IdCaso,IdGestion)VALUES(@Clave,12,'GESTION',@Actor,@Solicitud,CASE WHEN @Nuevo=1 THEN @Caso ELSE NULL END,CASE WHEN @Nuevo=0 THEN @Id ELSE NULL END);
 INSERT dbo.HistorialCobranzas(IdUsuario,Modulo,Accion,Detalle,IdCaso)VALUES(@Actor,12,CASE WHEN @Codigo='CCR'THEN 'CERRAR'ELSE 'GESTION'END,@Descripcion,@Caso);
 COMMIT;SELECT @Id Item,@Caso IdCaso,@Raiz Numero;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.CobSolicitarEnvio @Actor int,@Modulo int,@Item int,@Emision tinyint=1,@Clave uniqueidentifier
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF @Modulo NOT IN(14,15)OR @Clave IS NULL RAISERROR('Origen de correo inválido.',16,1);
 BEGIN TRANSACTION;DECLARE @Para nvarchar(1000),@Asunto nvarchar(255),@Cuerpo nvarchar(max),@IdEmision int,@Numero nvarchar(50),@Correo uniqueidentifier,@Cliente int;
 IF @Modulo=14 BEGIN SELECT @Para=Destino,@Cuerpo=CONVERT(nvarchar(max),Mensaj),@Numero=Asient FROM dbo.DbRecojos WITH(UPDLOCK,HOLDLOCK)WHERE Item=@Item AND EstadoGestion='Abierto';SET @Asunto='Orden de recojo '+ISNULL(@Numero,'');END
 ELSE BEGIN
 SELECT @Numero=Correl,@Cliente=IdCliente FROM dbo.DbsCartas WITH(UPDLOCK,HOLDLOCK)WHERE Item=@Item AND EstadoGestion='Abierto';
 SELECT @IdEmision=Item FROM dbo.EmisionesCarta WHERE IdCarta=@Item AND Numero=@Emision;
 SELECT @Para=COALESCE(NULLIF(t.Destinatarios,''),cfg.DestinoCartas)FROM dbo.ConfiguracionCobranzas cfg LEFT JOIN dbo.DbClientes c ON c.Item=@Cliente LEFT JOIN dbo.CorreosTerritorioCobranza t ON t.IdTerritorio=c.IdTerritorio WHERE cfg.Item=1;
 SELECT @Cuerpo=N'Se ha emitido la carta de cobranza '+@Numero+N' ('+CONVERT(nvarchar(3),@Emision)+N'ª emisión).'+CHAR(13)+CHAR(10)+N'Cliente: '+CliCar+CHAR(13)+CHAR(10)+N'Referencia: '+RefeCar+CHAR(13)+CHAR(10)+N'Importe: '+DivCar+N' '+CONVERT(nvarchar(50),MontCar)+CHAR(13)+CHAR(10)+N'Contacto: '+ISNULL(ContCar,'') FROM dbo.DbsCartas WHERE Item=@Item;
 SET @Asunto=N'Carta de cobranza '+ISNULL(@Numero,'');
 IF @IdEmision IS NULL RAISERROR('La emisión de la carta no existe.',16,1);
 END;
 IF @Numero IS NULL OR NULLIF(LTRIM(RTRIM(@Para)),'')IS NULL RAISERROR('Expediente cerrado o destinatarios sin configurar.',16,1);
 IF EXISTS(SELECT 1 FROM dbo.EnviosCobranzas WHERE Item=@Clave AND(IdUsuario<>@Actor OR ISNULL(IdRecojo,0)<>CASE WHEN @Modulo=14 THEN @Item ELSE 0 END OR ISNULL(IdEmision,0)<>ISNULL(@IdEmision,0)))RAISERROR('La clave identifica otro envío.',16,1);
 SELECT @Correo=Item FROM dbo.EnviosCobranzas WHERE Item=@Clave;
 IF @Correo IS NULL SELECT @Correo=Item FROM dbo.EnviosCobranzas WHERE Estado IN('Pendiente','Enviando')AND((@Modulo=14 AND IdRecojo=@Item)OR(@Modulo=15 AND IdEmision=@IdEmision));
 IF @Correo IS NULL BEGIN SET @Correo=@Clave;INSERT dbo.EnviosCobranzas(Item,IdRecojo,IdEmision,IdUsuario,Destinatario,Asunto,Cuerpo,Estado)VALUES(@Correo,CASE WHEN @Modulo=14 THEN @Item END,@IdEmision,@Actor,@Para,@Asunto,@Cuerpo,'Pendiente');END;
 COMMIT;SELECT Item,Estado FROM dbo.EnviosCobranzas WHERE Item=@Correo;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.DetalleQuery @Actor int,@Modulo int=6,@Item int,@Seccion int=0
AS
BEGIN
 SET NOCOUNT ON;DECLARE @Acceso int;EXEC @Acceso=dbo.VerificarActorQuery @Actor,@Modulo;IF @Acceso<>0 RETURN;
 IF @Seccion=0 SELECT q.*,t.TipDoc TipoDocumento,ISNULL(u.Usuari,q.ORIGEN) Creador,n.NumDoc NotaResolucion,n.Monto ImporteNota,c.Usuari CerradoPor FROM dbo.DbQuerie q LEFT JOIN dbo.Int_TipDoc t ON t.CodIte=q.TipDoc LEFT JOIN dbo.Int_Usuario u ON u.ITEM=q.IdCreador LEFT JOIN dbo.DbNCredito n ON n.ITEM=q.IdNotaResolucion LEFT JOIN dbo.Int_Usuario c ON c.ITEM=q.IdCerrador WHERE q.ITEM=@Item;
 ELSE IF @Seccion=1 SELECT h.Fecha,h.Accion,u.Usuari Usuario,h.Detalle FROM dbo.HistorialQueries h JOIN dbo.Int_Usuario u ON u.ITEM=h.IdUsuario WHERE h.IdQuery=@Item ORDER BY h.Item;
 ELSE IF @Seccion=2 SELECT NumeroDocumento Documento,Cliente,Emision,Divisa,Importe FROM dbo.DocumentosResolucionQuery WHERE IdQuery=@Item ORDER BY IdDocumento;
 ELSE IF @Seccion=3 SELECT NumeroGuia FROM dbo.GuiasQuery WHERE IdQuery=@Item ORDER BY NumeroGuia;
 ELSE RAISERROR('Sección inválida.',16,1);
END
GO
CREATE PROCEDURE dbo.CobConsultar @Actor int,@Modulo int,@Seccion int=0,@Item int=0,@Cliente int=0,@Buscar nvarchar(100)='',@Estado nvarchar(10)='',@Desde datetime=NULL,@Hasta datetime=NULL
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF @Desde>@Hasta BEGIN RAISERROR('Rango de fechas inválido.',16,1);RETURN;END;
 IF @Seccion=3 BEGIN
 SELECT * FROM dbo.DocumentosCobranzas WHERE(@Modulo=14 AND IdRecojo=@Item)OR(@Modulo=15 AND IdCarta=@Item)OR(@Modulo=16 AND IdDevolucion=@Item)ORDER BY Item;RETURN;END;
 IF @Seccion=4 BEGIN
 SELECT h.Fecha,u.Usuari Usuario,h.Accion,h.Detalle FROM dbo.HistorialCobranzas h JOIN dbo.Int_Usuario u ON u.ITEM=h.IdUsuario WHERE h.Modulo=@Modulo AND((@Modulo=12 AND h.IdCaso=@Item)OR(@Modulo=13 AND h.IdCheque=@Item)OR(@Modulo=14 AND h.IdRecojo=@Item)OR(@Modulo=15 AND h.IdCarta=@Item)OR(@Modulo=16 AND h.IdDevolucion=@Item)) ORDER BY h.Item;RETURN;END;
 IF @Seccion=5 AND @Modulo=15 BEGIN SELECT * FROM dbo.EmisionesCarta WHERE IdCarta=@Item ORDER BY Numero;RETURN;END;
 IF @Seccion=6 AND @Modulo=16 BEGIN SELECT * FROM dbo.GuiasDevolucion WHERE IdDevolucion=@Item ORDER BY Numero;RETURN;END;
 IF @Seccion=7 AND @Modulo IN(14,15) BEGIN
 SELECT e.Item,e.Fecha,e.Destinatario,e.Asunto,e.Estado,e.Inicio,e.Fin,e.Detalle FROM dbo.EnviosCobranzas e LEFT JOIN dbo.EmisionesCarta c ON c.Item=e.IdEmision WHERE(@Modulo=14 AND e.IdRecojo=@Item)OR(@Modulo=15 AND c.IdCarta=@Item)ORDER BY e.Fecha DESC;RETURN;END;
 IF @Modulo=12 BEGIN
 IF @Seccion=0 SELECT TOP(200)c.Item IdCliente,c.Asient,c.CLIENT,c.RUCDNI,q.FecApe,q.Usuario,(SELECT COUNT(*)FROM dbo.DbQuejas_Cas x WHERE x.IdCliente=c.Item) Casos,(SELECT COUNT(*)FROM dbo.DbQuejas_Cas x WHERE x.IdCliente=c.Item AND x.FecFin IS NULL) Abiertos,q.VersionFila FROM dbo.DbQuejas_Cli q JOIN dbo.DbClientes c ON c.Item=q.IdCliente WHERE(@Cliente=0 OR c.Item=@Cliente)AND(@Buscar=''OR CHARINDEX(@Buscar,c.CLIENT)>0 OR CHARINDEX(@Buscar,c.Asient)>0 OR CHARINDEX(@Buscar,c.RUCDNI)>0)ORDER BY c.CLIENT;
 ELSE IF @Seccion=1 SELECT TOP(200)q.*,c.CLIENT Cliente,c.RUCDNI RUCCliente,c.DIRECC DireccionCliente FROM dbo.DbQuejas_Cas q JOIN dbo.DbClientes c ON c.Item=q.IdCliente WHERE(@Cliente=0 OR q.IdCliente=@Cliente)AND(@Item=0 OR q.ITEM=@Item)AND(@Estado='' OR(@Estado='Abierto'AND q.FecFin IS NULL)OR(@Estado='Cerrado'AND q.FecFin IS NOT NULL))ORDER BY q.ITEM DESC;
 ELSE IF @Seccion=2 SELECT g.*,u.Usuari Usuario FROM dbo.DbQuejas_Ges g LEFT JOIN dbo.Int_Usuario u ON u.ITEM=g.IdUsuario WHERE g.IdCaso=@Item ORDER BY g.ITEM;
 ELSE RAISERROR('Sección de quejas inválida.',16,1);RETURN;END;
 IF @Seccion<>0 BEGIN RAISERROR('Sección de cobranza inválida.',16,1);RETURN;END;
 IF @Modulo=13 SELECT TOP(200)q.*,b.NomBank,m.Estado MotivoNombre,u.Usuari Creador FROM dbo.DbCheques q LEFT JOIN dbo.Int_Bancos b ON b.CodBank=q.Banco LEFT JOIN dbo.Int_EstadCheq m ON m.CodEst=q.Motivo LEFT JOIN dbo.Int_Usuario u ON u.ITEM=q.IdCreador WHERE(@Item=0 OR q.ITEM=@Item)AND(@Estado=''OR q.EstadoGestion=@Estado)AND(@Desde IS NULL OR q.FecCar>=@Desde)AND(@Hasta IS NULL OR q.FecCar<DATEADD(day,1,CONVERT(date,@Hasta)))AND(@Buscar=''OR CHARINDEX(@Buscar,q.Client)>0 OR CHARINDEX(@Buscar,q.NroChe)>0 OR CHARINDEX(@Buscar,q.Asient)>0)ORDER BY q.ITEM DESC;
 ELSE IF @Modulo=14 SELECT TOP(200)q.*,c.CLIENT Cliente,u.Usuari Creador,(SELECT TOP(1)e.Estado FROM dbo.EnviosCobranzas e WHERE e.IdRecojo=q.Item ORDER BY e.Fecha DESC)EstadoEnvio FROM dbo.DbRecojos q LEFT JOIN dbo.DbClientes c ON c.Item=q.IdCliente LEFT JOIN dbo.Int_Usuario u ON u.ITEM=q.IdCreador WHERE(@Item=0 OR q.Item=@Item)AND(@Estado=''OR q.EstadoGestion=@Estado)AND(@Desde IS NULL OR q.FecEnv>=@Desde)AND(@Hasta IS NULL OR q.FecEnv<DATEADD(day,1,CONVERT(date,@Hasta)))AND(@Buscar=''OR CHARINDEX(@Buscar,c.CLIENT)>0 OR CHARINDEX(@Buscar,q.Asient)>0)ORDER BY q.Item DESC;
 ELSE IF @Modulo=15 SELECT TOP(200)q.*,u.Usuari Creador,(SELECT COUNT(*)FROM dbo.EmisionesCarta e WHERE e.IdCarta=q.Item)Emisiones FROM dbo.DbsCartas q LEFT JOIN dbo.Int_Usuario u ON u.ITEM=q.IdCreador WHERE(@Item=0 OR q.Item=@Item)AND(@Estado=''OR q.EstadoGestion=@Estado)AND(@Desde IS NULL OR q.FecCar>=@Desde)AND(@Hasta IS NULL OR q.FecCar<DATEADD(day,1,CONVERT(date,@Hasta)))AND(@Buscar=''OR CHARINDEX(@Buscar,q.CliCar)>0 OR CHARINDEX(@Buscar,q.Correl)>0)ORDER BY q.Item DESC;
 ELSE IF @Modulo=16 SELECT TOP(200)q.*,u.Usuari Creador FROM dbo.DbDevol q LEFT JOIN dbo.Int_Usuario u ON u.ITEM=q.IdCreador WHERE(@Item=0 OR q.Item=@Item)AND(@Estado=''OR q.EstadoGestion=@Estado)AND(@Desde IS NULL OR q.FecCar>=@Desde)AND(@Hasta IS NULL OR q.FecCar<DATEADD(day,1,CONVERT(date,@Hasta)))AND(@Buscar=''OR CHARINDEX(@Buscar,q.Client)>0 OR CHARINDEX(@Buscar,q.NroCar)>0)ORDER BY q.Item DESC;
END
GO
CREATE PROCEDURE dbo.CobGuardarConfiguracion @Actor int,@Modulo int,@Version binary(8),@Datos xml,@CambiarClave bit=0,@ClaveProtegida varbinary(max)=NULL
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE ITEM=@Actor AND Access='ADMINISTRADOR'AND ActPas=1 AND SitUsu='Activo')RAISERROR('La configuración institucional y SMTP requiere administrador.',16,1);
 DECLARE @Host nvarchar(255),@Puerto int,@TLS bit;SELECT @Host=@Datos.value('(/config/@Host)[1]','nvarchar(255)'),@Puerto=@Datos.value('(/config/@Puerto)[1]','int'),@TLS=@Datos.value('(/config/@TLS)[1]','bit');
 IF @Puerto NOT BETWEEN 1 AND 65535 RAISERROR('Puerto SMTP inválido.',16,1);
 IF @TLS=0 AND @Host NOT IN('localhost','127.0.0.1','::1','')RAISERROR('Use STARTTLS para servidores externos.',16,1);
 IF @CambiarClave=1 AND(@ClaveProtegida IS NULL OR DATALENGTH(@ClaveProtegida)<20)RAISERROR('La clave protegida no es válida.',16,1);
 BEGIN TRANSACTION;
 UPDATE dbo.ConfiguracionCobranzas SET Empresa=@Datos.value('(/config/@Empresa)[1]','nvarchar(150)'),CuentaSoles=@Datos.value('(/config/@Soles)[1]','nvarchar(150)'),CuentaDolares=@Datos.value('(/config/@Dolares)[1]','nvarchar(150)'),Contacto=@Datos.value('(/config/@Contacto)[1]','nvarchar(150)'),CorreoConfirmacion=@Datos.value('(/config/@Correo)[1]','nvarchar(300)'),Anexo=@Datos.value('(/config/@Anexo)[1]','nvarchar(50)'),DestinoCheques=@Datos.value('(/config/@Cheques)[1]','nvarchar(1000)'),DestinoFacturas=@Datos.value('(/config/@Facturas)[1]','nvarchar(1000)'),DestinoCartas=@Datos.value('(/config/@Cartas)[1]','nvarchar(1000)'),HostSMTP=@Host,PuertoSMTP=@Puerto,UsarTLS=@TLS,UsuarioSMTP=@Datos.value('(/config/@Usuario)[1]','nvarchar(255)'),RemitenteSMTP=@Datos.value('(/config/@Remitente)[1]','nvarchar(255)'),ClaveSMTP=CASE WHEN @CambiarClave=1 THEN @ClaveProtegida ELSE ClaveSMTP END,EquipoSMTP=CASE WHEN @CambiarClave=1 THEN @Datos.value('(/config/@Equipo)[1]','nvarchar(128)')ELSE EquipoSMTP END WHERE Item=1 AND VersionFila=@Version;
 IF @@ROWCOUNT<>1 RAISERROR('La configuración cambió. Recargue antes de guardar.',16,1);
 COMMIT;SELECT VersionFila FROM dbo.ConfiguracionCobranzas WHERE Item=1;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.GuardarOtroCasoQuery @Actor int,@Item int=0,@Version binary(8)=NULL,@Descripcion nvarchar(150)=NULL,@Eliminar bit=0
AS
BEGIN
 SET NOCOUNT ON;SET XACT_ABORT ON;
 BEGIN TRY
 DECLARE @Permiso int;EXEC @Permiso=dbo.VerificarActorQuery @Actor,5;IF @Permiso<>0 RETURN;BEGIN TRANSACTION;
 IF @Item=0 BEGIN
 IF @Eliminar=1 OR NULLIF(LTRIM(RTRIM(@Descripcion)),'') IS NULL RAISERROR('Ingrese la descripción del caso.',16,1);
 INSERT dbo.OtrosCasosQuery(Descripcion)VALUES(LTRIM(RTRIM(@Descripcion)));SET @Item=SCOPE_IDENTITY();
 END ELSE BEGIN
 IF @Eliminar<>1 RAISERROR('Operación de caso no válida.',16,1);
 UPDATE dbo.OtrosCasosQuery SET Activo=0 WHERE Item=@Item AND VersionFila=@Version AND Activo=1;
 IF @@ROWCOUNT<>1 RAISERROR('El caso cambió o fue eliminado. Recargue.',16,1);
 END;COMMIT;SELECT Item,Descripcion,Activo,VersionFila FROM dbo.OtrosCasosQuery WHERE Item=@Item;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @Error nvarchar(2048);SET @Error=ERROR_MESSAGE();RAISERROR('%s',16,1,@Error);END CATCH;
END
GO
CREATE PROCEDURE dbo.CobResolverEnvio @Actor int,@Modulo int,@Envio uniqueidentifier,@Version binary(8),@Motivo nvarchar(1000)
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF @Modulo NOT IN(14,15)OR NULLIF(LTRIM(RTRIM(@Motivo)),'')IS NULL RAISERROR('Indique la verificación realizada antes de liberar el intento.',16,1);
 BEGIN TRANSACTION;
 UPDATE dbo.EnviosCobranzas SET Estado='Incierto',Fin=GETDATE(),Detalle=N'Revisión manual: '+LEFT(@Motivo,980)
 WHERE Item=@Envio AND VersionFila=@Version AND Estado='Enviando'AND Inicio<=DATEADD(second,-120,GETDATE())
 AND((@Modulo=14 AND IdRecojo IS NOT NULL)OR(@Modulo=15 AND IdEmision IS NOT NULL));
 IF @@ROWCOUNT<>1 RAISERROR('El intento cambió, terminó o todavía está dentro del tiempo de transporte. Recargue.',16,1);
 INSERT dbo.HistorialCobranzas(IdUsuario,Modulo,Accion,Detalle,IdRecojo,IdCarta)
 SELECT @Actor,@Modulo,'CORREO',N'Incierto por revisión manual: '+LEFT(@Motivo,960),e.IdRecojo,c.IdCarta
 FROM dbo.EnviosCobranzas e LEFT JOIN dbo.EmisionesCarta c ON c.Item=e.IdEmision WHERE e.Item=@Envio;
 COMMIT;SELECT @Envio Item;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.CobCatalogos @Actor int,@Modulo int,@Catalogo int
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF @Catalogo=1 SELECT CodBank Item,NomBank Descripcion FROM dbo.Int_Bancos WHERE ActEst=1 ORDER BY NomBank;
 ELSE IF @Catalogo=2 SELECT CodEst Item,Estado Descripcion FROM dbo.Int_EstadCheq WHERE ActEst=1 ORDER BY Estado;
 ELSE IF @Catalogo=3 SELECT CodEst Item,Moneda Descripcion FROM dbo.Int_Moneda WHERE ActEst=1 ORDER BY CodEst;
 ELSE IF @Catalogo=4 SELECT CodCas Item,NomCas Descripcion,ActEst Activo,VersionFila FROM dbo.Int_Casos ORDER BY NomCas;
 ELSE IF @Catalogo=5 SELECT CodSta Item,Status Codigo,DesSta Descripcion,ActEst Activo,VersionFila FROM dbo.Int_StatusCas ORDER BY CodSta;
 ELSE IF @Catalogo=6 SELECT Item,Descrip Descripcion FROM dbo.IntTerritorios WHERE Active=1 ORDER BY Descrip;
 ELSE IF @Catalogo=7 SELECT Item,Infor1 Descripcion FROM dbo.Int_InfoSys WHERE ModAct='MOTCARDEV' ORDER BY Infor1;
 ELSE RAISERROR('Catálogo de cobranzas inválido.',16,1);
END
GO
CREATE PROCEDURE dbo.BuscarNotasQuery @Actor int,@Query int,@Buscar nvarchar(50)=''
AS
BEGIN
 SET NOCOUNT ON;DECLARE @Acceso int;EXEC @Acceso=dbo.VerificarActorQuery @Actor,7;IF @Acceso<>0 RETURN;
 SELECT TOP(100) d.* FROM dbo.DocumentosGestion d JOIN dbo.DbNCredito n ON d.Tipo='NC' AND n.ITEM=d.IdDocumento
 JOIN dbo.DbQuerie q ON q.ITEM=@Query AND q.SitDoc='ACT' AND q.Estado='Abierto'
 WHERE n.IdDocumentoReferencia=q.IdFactura AND n.IdCliente=q.IdCliente AND d.Divisa=q.DivSol AND d.Anulado=0
 AND NOT EXISTS(SELECT 1 FROM dbo.DbQuerie x WHERE x.IdNotaResolucion=n.ITEM)
 AND(@Buscar='' OR CHARINDEX(@Buscar,d.Documento)>0) ORDER BY d.Emision DESC,d.Documento;
END
GO
-- Uso exclusivo del auxiliar de transporte: reclamación atómica, sin transacción durante la red.
CREATE PROCEDURE dbo.CobReclamarEnvio @Envio uniqueidentifier
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 BEGIN TRANSACTION;DECLARE @Actor int,@Modulo int,@P int;
 SELECT @Actor=IdUsuario,@Modulo=CASE WHEN IdRecojo IS NULL THEN 15 ELSE 14 END FROM dbo.EnviosCobranzas WITH(UPDLOCK,HOLDLOCK)WHERE Item=@Envio AND Estado='Pendiente';
 IF @Actor IS NULL BEGIN COMMIT;RETURN;END;
 EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 BEGIN ROLLBACK;RETURN;END;
 IF EXISTS(SELECT 1 FROM dbo.EnviosCobranzas e LEFT JOIN dbo.DbRecojos r ON r.Item=e.IdRecojo LEFT JOIN dbo.EmisionesCarta ec ON ec.Item=e.IdEmision LEFT JOIN dbo.DbsCartas c ON c.Item=ec.IdCarta WHERE e.Item=@Envio AND COALESCE(r.EstadoGestion,c.EstadoGestion)<>'Abierto')RAISERROR('El expediente se cerró antes del envío.',16,1);
 UPDATE dbo.EnviosCobranzas SET Estado='Enviando',Inicio=GETDATE()WHERE Item=@Envio;COMMIT;
 SELECT e.Item,e.Destinatario,e.Asunto,e.Cuerpo,c.HostSMTP,c.PuertoSMTP,c.UsarTLS,c.UsuarioSMTP,c.RemitenteSMTP,c.ClaveSMTP,c.EquipoSMTP FROM dbo.EnviosCobranzas e CROSS JOIN dbo.ConfiguracionCobranzas c WHERE e.Item=@Envio AND c.Item=1;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.CobSegundaCarta @Actor int,@Item int,@Version binary(8)
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,15;IF @P<>0 RETURN;
 BEGIN TRANSACTION;DECLARE @Datos xml,@Cfg xml;
 SELECT @Datos=Solicitud FROM dbo.DbsCartas WITH(UPDLOCK,HOLDLOCK)WHERE Item=@Item AND VersionFila=@Version AND EstadoGestion='Abierto' AND RefeCar='Facturas pendientes de Pago';
 IF @Datos IS NULL RAISERROR('La carta cambió, no está abierta o no corresponde a facturas pendientes.',16,1);
 IF EXISTS(SELECT 1 FROM dbo.EmisionesCarta WHERE IdCarta=@Item AND Numero=2)RAISERROR('La segunda carta ya fue emitida. Use reimprimir.',16,1);
 SET @Cfg=(SELECT Empresa,CuentaSoles,CuentaDolares,Contacto,CorreoConfirmacion,Anexo FROM dbo.ConfiguracionCobranzas WHERE Item=1 FOR XML PATH('empresa'),TYPE);
 SET @Datos.modify('delete /registro/empresa');SET @Datos.modify('insert sql:variable("@Cfg") as last into (/registro)[1]');
 INSERT dbo.EmisionesCarta(IdCarta,Numero,IdUsuario,Modelo,Datos)VALUES(@Item,2,@Actor,3,@Datos);
 UPDATE dbo.DbsCartas SET SgdaCar=1 WHERE Item=@Item;
 INSERT dbo.HistorialCobranzas(IdUsuario,Modulo,Accion,Detalle,IdCarta)VALUES(@Actor,15,'SEGUNDA CARTA','Segunda emisión registrada.',@Item);
 COMMIT;SELECT @Item Item;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.BuscarDocumentoQuery @Actor int,@Modulo int=5,@Tipo nvarchar(2),@Numero nvarchar(50)
AS
BEGIN
 SET NOCOUNT ON;DECLARE @Acceso int;EXEC @Acceso=dbo.VerificarActorQuery @Actor,@Modulo;IF @Acceso<>0 RETURN;
 SELECT d.*,c.RUCDNI,c.DIRECC,di.Nombre Distrito,t.Descrip Territorio,
 ISNULL(STUFF((SELECT N', '+s.NroGui FROM dbo.DbServGui s WHERE (d.Tipo='NC' AND s.IdNotaCredito=d.IdDocumento)OR(d.Tipo<>'NC' AND s.IdFactura=d.IdDocumento) ORDER BY s.Linea,s.Item FOR XML PATH(''),TYPE).value('.','nvarchar(max)'),1,2,''),'') Guias
 FROM dbo.DocumentosGestion d LEFT JOIN dbo.DbClientes c ON c.Item=d.IdCliente LEFT JOIN dbo.Distritos di ON di.Item=c.IdDistrito LEFT JOIN dbo.IntTerritorios t ON t.Item=c.IdTerritorio
 WHERE d.Tipo=@Tipo AND d.Documento=@Numero;
END
GO
CREATE PROCEDURE dbo.CobEstadoEnvio @Actor int,@Modulo int,@Envio uniqueidentifier
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 SELECT Item,Estado,Detalle,Inicio,Fin FROM dbo.EnviosCobranzas WHERE Item=@Envio AND((@Modulo=14 AND IdRecojo IS NOT NULL)OR(@Modulo=15 AND IdEmision IS NOT NULL));
END
GO
CREATE PROCEDURE dbo.FinalizarQuery @Actor int,@Item int,@Version binary(8),@IdNota int=0,@VersionNota binary(8)=NULL,@Monto money=0,@Resolucion nvarchar(1000),@Documentos xml
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;
 BEGIN TRY
 DECLARE @Permiso int;EXEC @Permiso=dbo.VerificarActorQuery @Actor,7;IF @Permiso<>0 RETURN;
 IF NULLIF(LTRIM(RTRIM(@Resolucion)),'')IS NULL RAISERROR('Describa la solución confirmada por el sistema externo.',16,1);
 DECLARE @Docs TABLE(Id int PRIMARY KEY,Version binary(8));
 INSERT @Docs SELECT n.value('@Id','int'),CONVERT(binary(8),n.value('@Version','varchar(18)'),1) FROM @Documentos.nodes('/documentos/d')x(n);
 IF(SELECT COUNT(*)FROM @Docs)>200 RAISERROR('Demasiados documentos de resolución.',16,1);
 BEGIN TRANSACTION;
 DECLARE @Clase nvarchar(20),@IdFactura int,@Cliente int,@Divisa nvarchar(50),@Total money,@Pedido money,@Numero nvarchar(10);
 SELECT @Clase=TipQue,@IdFactura=IdFactura,@Cliente=IdCliente,@Divisa=DivSol,@Total=MonSol,@Pedido=SolDes,@Numero=Asient FROM dbo.DbQuerie WITH(UPDLOCK,HOLDLOCK)WHERE ITEM=@Item AND VersionFila=@Version AND SitDoc='ACT' AND Estado='Abierto';
 IF @Clase IS NULL OR @Clase NOT IN('ANL ANL','ANL NC','DES NC','DES 100') OR @IdFactura IS NULL RAISERROR('La solicitud cambió, no está abierta o carece de factura vinculada.',16,1);
 IF @Clase='ANL ANL' BEGIN IF @IdNota<>0 OR @Monto<>0 RAISERROR('La anulación directa no lleva nota de crédito aplicada.',16,1);END
 ELSE BEGIN
 DECLARE @ImporteNC money;
 SELECT @ImporteNC=n.Monto FROM dbo.DbNCredito n WITH(UPDLOCK,HOLDLOCK)LEFT JOIN dbo.Int_Moneda m ON CONVERT(nvarchar(50),m.CodEst)=n.DivMon
 WHERE n.ITEM=@IdNota AND n.VersionFila=@VersionNota AND n.IdCliente=@Cliente AND n.IdDocumentoReferencia=@IdFactura AND COALESCE(m.Moneda,n.DivMon)=@Divisa AND UPPER(n.SitDoc)IN('ACT','ACTIVO');
 IF @ImporteNC IS NULL OR @Monto<=0 OR @Monto>@ImporteNC OR @Monto>@Pedido RAISERROR('Nota, cliente, referencia, moneda, versión o monto aplicado incompatibles.',16,1);
 IF @Clase IN('ANL NC','DES 100') AND @Monto<>@Total RAISERROR('La resolución total requiere cubrir el importe completo del documento.',16,1);
 IF EXISTS(SELECT 1 FROM dbo.DbQuerie WITH(UPDLOCK,HOLDLOCK)WHERE IdNotaResolucion=@IdNota)RAISERROR('La nota ya resolvió otra solicitud.',16,1);
 END;
 DECLARE @Bloqueo int;SELECT @Bloqueo=d.Item FROM dbo.DBEmiFac d WITH(UPDLOCK,HOLDLOCK)JOIN @Docs x ON d.Item=x.Id;
 IF EXISTS(SELECT 1 FROM @Docs x LEFT JOIN dbo.DBEmiFac d ON d.Item=x.Id LEFT JOIN dbo.Int_TipDoc t ON t.CodIte=d.TipDoc LEFT JOIN dbo.Int_Moneda m ON CONVERT(nvarchar(50),m.CodEst)=d.DivFac WHERE d.Item IS NULL OR d.Item=@IdFactura OR d.IdCliente<>@Cliente OR COALESCE(m.Moneda,d.DivFac)<>@Divisa OR t.TipDoc NOT IN('FT','BV') OR UPPER(d.SitDoc)NOT IN('ACT','ACTIVO') OR d.VersionFila<>x.Version) RAISERROR('Un documento de reemplazo cambió o no corresponde al cliente y moneda.',16,1);
 INSERT dbo.DocumentosResolucionQuery SELECT @Item,d.Item,d.NumDoc,d.NomCli,d.FecEmi,COALESCE(m.Moneda,d.DivFac),d.ImpFac FROM @Docs x JOIN dbo.DBEmiFac d ON d.Item=x.Id LEFT JOIN dbo.Int_Moneda m ON CONVERT(nvarchar(50),m.CodEst)=d.DivFac;
 UPDATE dbo.DbQuerie SET Estado='Cerrado',FechaCierre=GETDATE(),IdCerrador=@Actor,Resolucion=@Resolucion,MontoAplicado=@Monto,IdNotaResolucion=NULLIF(@IdNota,0) WHERE ITEM=@Item;
 INSERT dbo.HistorialQueries(IdQuery,IdUsuario,Accion,Detalle)VALUES(@Item,@Actor,'CERRAR',@Resolucion);
 COMMIT;SELECT ITEM,Asient,VersionFila FROM dbo.DbQuerie WHERE ITEM=@Item;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @Error nvarchar(2048);SET @Error=ERROR_MESSAGE();RAISERROR('%s',16,1,@Error);END CATCH;
END
GO
CREATE PROCEDURE dbo.CobGuardarRegistro @Actor int,@Modulo int,@Clave uniqueidentifier,@Solicitud xml
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF @Modulo NOT BETWEEN 13 AND 16 OR @Clave IS NULL OR @Solicitud.value('count(/registro)','int')<>1 RAISERROR('Solicitud de cobranza inválida.',16,1);
 DECLARE @Cliente int,@Version binary(8),@Moneda int,@Divisa nvarchar(50),@Nombre nvarchar(255),@Codigo nvarchar(20),@Direccion nvarchar(255),@Ruc nvarchar(20),@Contacto nvarchar(50),@Cargo nvarchar(50),@Banco int,@Motivo int,@MotivoNombre nvarchar(150),@BancoNombre nvarchar(150),@Monto money,@Gastos money,@Fecha datetime,@Giro datetime,@Cheque nvarchar(50),@Modalidad nvarchar(20),@Referencia nvarchar(50),@Manual bit,@SinDoc bit,@Numero nvarchar(10),@Serie varchar(3),@Id int,@Usuario nvarchar(50),@Foto xml,@Cfg xml,@Guias nvarchar(100),@IdEmision int;
 SELECT @Cliente=@Solicitud.value('(/registro/@Cliente)[1]','int'),@Version=CONVERT(binary(8),@Solicitud.value('(/registro/@VersionCliente)[1]','varchar(18)'),1),@Moneda=@Solicitud.value('(/registro/@Moneda)[1]','int'),@Contacto=@Solicitud.value('(/registro/@Contacto)[1]','nvarchar(50)'),@Cargo=@Solicitud.value('(/registro/@Cargo)[1]','nvarchar(50)'),@Banco=@Solicitud.value('(/registro/@Banco)[1]','int'),@Motivo=@Solicitud.value('(/registro/@Motivo)[1]','int'),@Monto=@Solicitud.value('(/registro/@Monto)[1]','money'),@Gastos=@Solicitud.value('(/registro/@Gastos)[1]','money'),@Fecha=@Solicitud.value('(/registro/@Fecha)[1]','datetime'),@Manual=@Solicitud.value('(/registro/@Manual)[1]','bit'),@SinDoc=@Solicitud.value('(/registro/@SinDocumentos)[1]','bit'),@Modalidad=@Solicitud.value('(/registro/@Modalidad)[1]','nvarchar(20)'),@Referencia=@Solicitud.value('(/registro/@Referencia)[1]','nvarchar(50)');
 IF @Fecha<'19000101' OR @Fecha>=DATEADD(day,1,CONVERT(date,GETDATE())) OR @Fecha IS NULL RAISERROR('Fecha de gestión inválida.',16,1);
 IF @Modulo<>14 AND(@Monto<=0 OR @Monto IS NULL OR @Gastos<0)RAISERROR('Ingrese un importe positivo y gastos no negativos.',16,1);
 IF NULLIF(LTRIM(RTRIM(@Contacto)),'')IS NULL RAISERROR('Ingrese el contacto responsable.',16,1);
 BEGIN TRANSACTION;
 SELECT @Id=COALESCE(IdCheque,IdRecojo,IdCarta,IdDevolucion)FROM dbo.ClavesCobranzas WITH(UPDLOCK,HOLDLOCK)WHERE Clave=@Clave;
 IF @Id IS NOT NULL BEGIN
 IF NOT EXISTS(SELECT 1 FROM dbo.ClavesCobranzas WHERE Clave=@Clave AND Modulo=@Modulo AND IdUsuario=@Actor AND CONVERT(nvarchar(max),Solicitud) COLLATE Latin1_General_BIN2=CONVERT(nvarchar(max),@Solicitud) COLLATE Latin1_General_BIN2)RAISERROR('La clave ya identifica otros datos.',16,1);
 COMMIT;SELECT @Id Item;RETURN;END;
 SELECT @Usuario=Usuari FROM dbo.Int_Usuario WHERE ITEM=@Actor;
 IF @Manual=1 BEGIN
 IF @Modulo<>16 OR @Cliente<>0 OR @SinDoc<>1 RAISERROR('El destinatario manual se admite solo en devolución sin documentos.',16,1);
 SET @Nombre=@Solicitud.value('(/registro/@NombreManual)[1]','nvarchar(255)');SET @Codigo='';SET @Direccion='';SET @Ruc='';
 IF NULLIF(LTRIM(RTRIM(@Nombre)),'')IS NULL RAISERROR('Ingrese el destinatario manual.',16,1);
 END ELSE BEGIN
 SELECT @Nombre=CLIENT,@Codigo=Asient,@Direccion=DIRECC,@Ruc=RUCDNI FROM dbo.DbClientes WITH(HOLDLOCK)WHERE Item=@Cliente AND VersionFila=@Version AND Activo=1;
 IF @Nombre IS NULL RAISERROR('El cliente cambió o no está activo. Vuelva a seleccionarlo con F9.',16,1);
 END;
 SELECT @Divisa=Moneda FROM dbo.Int_Moneda WITH(HOLDLOCK)WHERE CodEst=@Moneda AND ActEst=1;
 IF @Divisa IS NULL RAISERROR('Seleccione una moneda activa.',16,1);
 IF @Banco<>0 BEGIN SELECT @BancoNombre=NomBank FROM dbo.Int_Bancos WITH(HOLDLOCK)WHERE CodBank=@Banco AND ActEst=1;IF @BancoNombre IS NULL RAISERROR('Seleccione un banco activo.',16,1);END;
 DECLARE @Docs TABLE(Id int PRIMARY KEY,Version binary(8));
 INSERT @Docs SELECT n.value('@Id','int'),CONVERT(binary(8),n.value('@Version','varchar(18)'),1) FROM @Solicitud.nodes('/registro/d')x(n);
 IF(SELECT COUNT(*)FROM @Docs)>200 OR(@Modulo=16 AND(SELECT COUNT(*)FROM @Docs)>6)RAISERROR('Demasiados documentos asociados.',16,1);
 DECLARE @Bloqueo int;SELECT @Bloqueo=d.Item FROM dbo.DBEmiFac d WITH(HOLDLOCK)JOIN @Docs x ON d.Item=x.Id;
 IF EXISTS(SELECT 1 FROM @Docs x LEFT JOIN dbo.DocumentosGestion d ON d.IdDocumento=x.Id AND d.Tipo<>'NC' WHERE d.IdDocumento IS NULL OR d.IdCliente<>@Cliente OR d.Divisa<>@Divisa OR d.VersionFila<>x.Version OR UPPER(d.Situacion)NOT IN('ACT','ACTIVO')OR d.Tipo NOT IN('FT','BV','ND')OR(@Modulo=14 AND d.Tipo NOT IN('FT','BV')))RAISERROR('Un documento cambió, está anulado o pertenece a otro cliente, moneda o tipo.',16,1);
 IF @Modulo=13 AND EXISTS(SELECT 1 FROM @Docs)RAISERROR('Este registro de cheque no admite documentos adjuntos.',16,1);
 IF @Modulo=14 BEGIN
 IF @Modalidad NOT IN('CHEQUE','FACTURA')OR(@Modalidad='FACTURA'AND NOT EXISTS(SELECT 1 FROM @Docs))OR(@Modalidad='CHEQUE'AND EXISTS(SELECT 1 FROM @Docs))RAISERROR('Revise modalidad y documentos del recojo.',16,1);
 IF NULLIF(LTRIM(RTRIM(@Solicitud.value('(/registro/@Mensaje)[1]','nvarchar(4000)'))),'')IS NULL OR NULLIF(@Solicitud.value('(/registro/@Destino)[1]','nvarchar(1000)'),'')IS NULL RAISERROR('Complete destinatarios y mensaje del recojo.',16,1);
 END;
 IF @Modulo=15 BEGIN
 IF @Referencia NOT IN(N'Cheque Rechazado',N'Facturas pendientes de Pago')RAISERROR('Referencia de carta inválida.',16,1);
 IF NULLIF(LTRIM(RTRIM(@Cargo)),'')IS NULL RAISERROR('Ingrese el cargo del destinatario.',16,1);
 END;
 IF @Modulo=16 BEGIN
 IF @Manual=1 AND EXISTS(SELECT 1 FROM @Docs)RAISERROR('Un destinatario manual no puede asociar facturas de un cliente.',16,1);
 IF @SinDoc=0 AND NOT EXISTS(SELECT 1 FROM @Docs)RAISERROR('Asocie documentos o indique devolución sin documentos.',16,1);
 IF @SinDoc=1 AND EXISTS(SELECT 1 FROM @Docs)RAISERROR('La devolución sin documentos debe tener la lista vacía.',16,1);
 IF @Modalidad NOT IN('TRANSFERENCIA','CHEQUE','PARA DEPOSITAR')OR(@Modalidad<>'CHEQUE'AND @Banco=0)RAISERROR('Seleccione forma de devolución y banco.',16,1);
 IF @Modalidad<>'CHEQUE' AND NULLIF(@Solicitud.value('(/registro/@Cuenta)[1]','nvarchar(50)'),'')IS NULL RAISERROR('Ingrese la cuenta bancaria para transferencia o depósito.',16,1);
 SET @MotivoNombre=@Solicitud.value('(/registro/@MotivoDevol)[1]','nvarchar(100)');
 IF NULLIF(LTRIM(RTRIM(@MotivoNombre)),'')IS NULL RAISERROR('Seleccione el motivo de devolución.',16,1);
 END;
 IF @Modulo=13 BEGIN
 SELECT @MotivoNombre=Estado FROM dbo.Int_EstadCheq WITH(HOLDLOCK)WHERE CodEst=@Motivo AND ActEst=1;
 SET @Giro=@Solicitud.value('(/registro/@FechaGiro)[1]','datetime');SET @Cheque=LTRIM(RTRIM(@Solicitud.value('(/registro/@Cheque)[1]','nvarchar(50)')));
 IF @Banco=0 OR @MotivoNombre IS NULL OR NULLIF(@Cheque,'')IS NULL OR @Giro<'19000101'OR @Giro IS NULL RAISERROR('Revise banco, motivo, cheque y fecha de giro.',16,1);
 IF EXISTS(SELECT 1 FROM dbo.DbCheques WITH(UPDLOCK,HOLDLOCK)WHERE IdCliente=@Cliente AND Banco=@Banco AND NroChe=@Cheque AND FecGir=@Giro AND SitDoc='ACT')RAISERROR('El cheque ya está registrado para este cliente, banco y fecha.',16,1);
 END;
 DECLARE @G TABLE(Numero nvarchar(10) PRIMARY KEY);
 INSERT @G SELECT n.value('@Numero','nvarchar(10)')FROM @Solicitud.nodes('/registro/g')x(n);
 IF(SELECT COUNT(*)FROM @G)>6 OR(@Modulo<>16 AND EXISTS(SELECT 1 FROM @G))RAISERROR('Solo Devoluciones admite hasta seis guías.',16,1);
 IF EXISTS(SELECT 1 FROM @G x LEFT JOIN dbo.DbGuias g WITH(HOLDLOCK)ON g.NroGui=x.Numero WHERE g.Item IS NULL)RAISERROR('Una guía no existe.',16,1);
 IF EXISTS(SELECT 1 FROM @G x WHERE EXISTS(SELECT 1 FROM @Docs)AND NOT EXISTS(SELECT 1 FROM dbo.DbServGui s JOIN @Docs d ON s.IdFactura=d.Id WHERE s.NroGui=x.Numero))RAISERROR('Una guía no corresponde a los documentos seleccionados.',16,1);
 SET @Guias=ISNULL(STUFF((SELECT ','+Numero FROM @G ORDER BY Numero FOR XML PATH(''),TYPE).value('.','nvarchar(100)'),1,1,''),'');
 SET @Cfg=(SELECT Empresa,CuentaSoles,CuentaDolares,Contacto,CorreoConfirmacion,Anexo FROM dbo.ConfiguracionCobranzas WHERE Item=1 FOR XML PATH('empresa'),TYPE);
 SET @Foto=@Solicitud;
 SET @Foto.modify('insert sql:variable("@Cfg") as last into (/registro)[1]');
 SET @Foto.modify('insert attribute ClienteNombre {sql:variable("@Nombre")} into (/registro)[1]');
 SET @Foto.modify('insert attribute CodigoCliente {sql:variable("@Codigo")} into (/registro)[1]');
 SET @Foto.modify('insert attribute ClienteDireccion {sql:variable("@Direccion")} into (/registro)[1]');
 SET @Foto.modify('insert attribute ClienteRuc {sql:variable("@Ruc")} into (/registro)[1]');
 SET @Foto.modify('insert attribute Divisa {sql:variable("@Divisa")} into (/registro)[1]');
 IF @BancoNombre IS NOT NULL SET @Foto.modify('insert attribute BancoNombre {sql:variable("@BancoNombre")} into (/registro)[1]');
 IF @MotivoNombre IS NOT NULL SET @Foto.modify('insert attribute MotivoNombre {sql:variable("@MotivoNombre")} into (/registro)[1]');
 SET @Serie=CASE @Modulo WHEN 13 THEN 'CCR'WHEN 14 THEN 'RCJ'WHEN 15 THEN 'CRT'ELSE 'CDV'END;
 EXEC dbo.CobReservarNumero @Serie,@Numero OUTPUT;
 IF @Modulo=13 BEGIN
 INSERT dbo.DbCheques(Asient,Client,CodCli,Direcc,Contac,Cargo,Banco,FecGir,NroChe,Monto,Divisa,Motivo,Gastos,MonCan,Estado,SitDoc,FecCar,NroRUC,IdCliente,IdCreador,Solicitud)VALUES(@Numero,@Nombre,@Codigo,@Direccion,@Contacto,@Cargo,@Banco,@Giro,@Cheque,@Monto,@Divisa,@Motivo,@Gastos,@Monto+@Gastos,'Pendiente','ACT',@Fecha,@Ruc,@Cliente,@Actor,@Foto);SET @Id=SCOPE_IDENTITY();END
 ELSE IF @Modulo=14 BEGIN
 INSERT dbo.DbRecojos(Asient,FecEnv,ModGes,CodCli,Mensaj,Estado,Destino,IdCliente,IdCreador,Solicitud)VALUES(@Numero,@Fecha,@Modalidad,@Codigo,@Solicitud.value('(/registro/@Mensaje)[1]','nvarchar(4000)'),1,@Solicitud.value('(/registro/@Destino)[1]','nvarchar(1000)'),@Cliente,@Actor,@Foto);SET @Id=SCOPE_IDENTITY();END
 ELSE IF @Modulo=15 BEGIN
 INSERT dbo.DbsCartas(Correl,FecCar,CliCar,CodCli,DirCar,ContCar,CargCar,RefeCar,DivCar,MontCar,SgdaCar,IdCliente,IdCreador,Solicitud)VALUES(@Numero,@Fecha,@Nombre,@Codigo,@Solicitud.value('(/registro/@Direccion)[1]','nvarchar(255)'),@Contacto,@Cargo,@Referencia,@Divisa,@Monto,0,@Cliente,@Actor,@Foto);SET @Id=SCOPE_IDENTITY();
 INSERT dbo.EmisionesCarta(IdCarta,Numero,IdUsuario,Modelo,Datos)VALUES(@Id,1,@Actor,CASE WHEN @Referencia='Cheque Rechazado'THEN 1 ELSE 2 END,@Foto);SET @IdEmision=SCOPE_IDENTITY();END
 ELSE BEGIN
 INSERT dbo.DbDevol(NroCar,FecCar,Usuari,ImpExp,Client,CodCli,NroGui,Cancel,Abonos,Motivo,Transf,Otros,ParDep,CtaCte,Monto,Moneda,Contac,Estado,Final,IdCliente,IdUsuario,IdCreador,Solicitud)VALUES(@Numero,@Fecha,@Usuario,'IMPORTACION',@Nombre,@Codigo,@Guias,@Solicitud.value('(/registro/@Cancelado)[1]','nvarchar(100)'),@Solicitud.value('(/registro/@Abonos)[1]','nvarchar(100)'),@MotivoNombre,CASE WHEN @Modalidad='TRANSFERENCIA'THEN ISNULL(@BancoNombre,'')ELSE ''END,CASE WHEN @Modalidad='CHEQUE'THEN 'CHEQUE'ELSE ''END,CASE WHEN @Modalidad='PARA DEPOSITAR'THEN ISNULL(@BancoNombre,'')ELSE ''END,@Solicitud.value('(/registro/@Cuenta)[1]','nvarchar(50)'),@Monto,@Divisa,@Contacto,1,1,NULLIF(@Cliente,0),@Actor,@Actor,@Foto);SET @Id=SCOPE_IDENTITY();
 INSERT dbo.GuiasDevolucion SELECT @Id,g.Item,g.NroGui FROM @G x JOIN dbo.DbGuias g ON g.NroGui=x.Numero;END;
 INSERT dbo.DocumentosCobranzas(IdRecojo,IdCarta,IdDevolucion,IdFactura,Tipo,Documento,Emision,Cliente,Divisa,Importe)SELECT CASE WHEN @Modulo=14 THEN @Id END,CASE WHEN @Modulo=15 THEN @Id END,CASE WHEN @Modulo=16 THEN @Id END,d.IdDocumento,d.Tipo,d.Documento,d.Emision,d.Cliente,d.Divisa,d.Importe FROM @Docs x JOIN dbo.DocumentosGestion d ON d.IdDocumento=x.Id AND d.Tipo<>'NC';
 INSERT dbo.ClavesCobranzas(Clave,Modulo,Accion,IdUsuario,Solicitud,IdCheque,IdRecojo,IdCarta,IdDevolucion)VALUES(@Clave,@Modulo,'CREAR',@Actor,@Solicitud,CASE WHEN @Modulo=13 THEN @Id END,CASE WHEN @Modulo=14 THEN @Id END,CASE WHEN @Modulo=15 THEN @Id END,CASE WHEN @Modulo=16 THEN @Id END);
 INSERT dbo.HistorialCobranzas(IdUsuario,Modulo,Accion,Detalle,IdCheque,IdRecojo,IdCarta,IdDevolucion)VALUES(@Actor,@Modulo,'CREAR',@Numero,CASE WHEN @Modulo=13 THEN @Id END,CASE WHEN @Modulo=14 THEN @Id END,CASE WHEN @Modulo=15 THEN @Id END,CASE WHEN @Modulo=16 THEN @Id END);
 COMMIT;SELECT @Id Item,@Numero Numero,@IdEmision IdEmision;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.GuardarSolicitudQuery @Actor int,@Clave uniqueidentifier,@Solicitud xml
AS
BEGIN
 SET NOCOUNT ON;SET XACT_ABORT ON;
 BEGIN TRY
 DECLARE @Permiso int;EXEC @Permiso=dbo.VerificarActorQuery @Actor,5;IF @Permiso<>0 RETURN;
 IF @Solicitud IS NULL OR @Solicitud.value('count(/solicitud)','int')<>1 RAISERROR('La solicitud debe contener un solo expediente.',16,1);
 IF @Solicitud.value('(/solicitud/@Formato)[1]','int') NOT IN(1,2) RAISERROR('Formato de solicitud inválido.',16,1);
 IF @Clave IS NULL RAISERROR('Falta la clave de la solicitud.',16,1);
 DECLARE @Id int,@Doc int,@Version binary(8),@Tipo nvarchar(2),@Clase nvarchar(20),@Caso nvarchar(100),@Fecha datetime,@Inicio datetime,@Visto datetime,@Monto money,@Responsable int,@Servicio int,@Recargo money,@Otro int;
 SELECT @Doc=@Solicitud.value('(/solicitud/@Documento)[1]','int'),@Version=CONVERT(binary(8),@Solicitud.value('(/solicitud/@Version)[1]','varchar(18)'),1),@Clase=@Solicitud.value('(/solicitud/@Clase)[1]','nvarchar(20)'),@Caso=@Solicitud.value('(/solicitud/@Caso)[1]','nvarchar(100)'),@Fecha=@Solicitud.value('(/solicitud/@Fecha)[1]','datetime'),@Inicio=@Solicitud.value('(/solicitud/@Inicio)[1]','datetime'),@Visto=@Solicitud.value('(/solicitud/@Visto)[1]','datetime'),@Monto=@Solicitud.value('(/solicitud/@Monto)[1]','money'),@Responsable=@Solicitud.value('(/solicitud/@Responsable)[1]','int'),@Servicio=@Solicitud.value('(/solicitud/@Servicio)[1]','int'),@Recargo=@Solicitud.value('(/solicitud/@Recargo)[1]','money'),@Otro=@Solicitud.value('(/solicitud/@Otro)[1]','int');
 IF @Clase NOT IN('ANL ANL','ANL NC','DES NC','DES 100') OR @Caso NOT IN(N'Error en Datos del Cliente',N'Tarifas Erradas',N'Mal Servicio',N'Condiciones de Facturacion',N'Errores Otros - Especificar...') RAISERROR('Solicitud o caso inválido.',16,1);
 IF @Fecha IS NULL OR @Inicio IS NULL OR @Visto IS NULL OR @Fecha>=DATEADD(day,1,CONVERT(date,GETDATE())) OR @Inicio>@Fecha OR @Visto>@Fecha OR @Inicio< '19000101' OR @Visto<@Inicio RAISERROR('Revise las fechas de la solicitud y aprobación.',16,1);
 IF @Recargo<0 OR @Recargo>100 OR LEN(@Solicitud.value('(/solicitud/@Adicional)[1]','nvarchar(4000)'))>255 OR LEN(@Solicitud.value('(/solicitud/@ClientePropuesto)[1]','nvarchar(4000)'))>255 OR LEN(@Solicitud.value('(/solicitud/@RucPropuesto)[1]','nvarchar(4000)'))>11 OR LEN(@Solicitud.value('(/solicitud/@DireccionPropuesta)[1]','nvarchar(4000)'))>255 OR LEN(@Solicitud.value('(/solicitud/@TelefonoPropuesto)[1]','nvarchar(4000)'))>20 RAISERROR('Recargo o datos propuestos inválidos.',16,1);
 IF @Solicitud.exist('/solicitud[@SiLibGuia!="0" and @SiLibGuia!="1" or @NoLibGuia!="0" and @NoLibGuia!="1"]')=1 OR @Solicitud.value('(/solicitud/@SiLibGuia)[1]','int')+@Solicitud.value('(/solicitud/@NoLibGuia)[1]','int')<>1 RAISERROR('Seleccione una instrucción de liberación de guía.',16,1);
 IF @Caso='Error en Datos del Cliente' AND @Solicitud.exist('/solicitud[@ChkMal1="1" or @ChkMal2="1" or @ChkMal3="1" or @CambioCliente="1"]')=0 RAISERROR('Seleccione un motivo de datos.',16,1);
 IF @Caso='Tarifas Erradas' AND @Solicitud.exist('/solicitud[@ChkTar1="1" or @ChkTar2="1" or @ChkTar3="1" or @ChkTar4="1" or @ChkTar5="1"]')=0 RAISERROR('Seleccione un motivo de tarifa.',16,1);
 IF @Caso='Mal Servicio' AND @Solicitud.exist('/solicitud[@ChkServ1="1" or @ChkServ2="1" or @ChkServ3="1" or @ChkServ4="1" or @ChkServ5="1"]')=0 RAISERROR('Seleccione un motivo de servicio.',16,1);
 IF @Caso='Condiciones de Facturacion' AND @Solicitud.exist('/solicitud[@ChkCon1="1" or @ChkCon2="1" or @ChkCon3="1"]')=0 RAISERROR('Seleccione una condición de facturación.',16,1);
 IF @Caso='Errores Otros - Especificar...' AND @Otro=0 AND @Solicitud.exist('/solicitud[@ChkDescuento="1"]')=0 RAISERROR('Seleccione un caso u otro descuento.',16,1);
 IF @Caso='Error en Datos del Cliente' AND NULLIF(@Solicitud.value('(/solicitud/@ClientePropuesto)[1]','nvarchar(255)'),'') IS NULL RAISERROR('Indique los datos propuestos del cliente.',16,1);
 BEGIN TRANSACTION;
 DECLARE @Numero int,@Asiento nvarchar(10),@Usuario nvarchar(50),@Total money,@Divisa nvarchar(50),@Nombre nvarchar(255),@Codigo nvarchar(50),@IdCliente int,@Emision datetime,@IdTipo tinyint,@NombreResponsable nvarchar(150),@NombreServicio nvarchar(150);
 SELECT @Numero=Numero FROM dbo.SeriesQueries WITH(UPDLOCK,HOLDLOCK) WHERE Serie='Q';
 SELECT @Id=ITEM FROM dbo.DbQuerie WITH(UPDLOCK,HOLDLOCK)WHERE ClaveSolicitud=@Clave;
 IF @Id IS NOT NULL BEGIN
 IF NOT EXISTS(SELECT 1 FROM dbo.DbQuerie WHERE ITEM=@Id AND IdCreador=@Actor AND CONVERT(nvarchar(max),SolicitudOriginal) COLLATE Latin1_General_BIN2=CONVERT(nvarchar(max),@Solicitud) COLLATE Latin1_General_BIN2) RAISERROR('La clave de solicitud ya identifica otros datos.',16,1);
 COMMIT;SELECT ITEM,Asient,VersionFila FROM dbo.DbQuerie WHERE ITEM=@Id;RETURN;END;
 SELECT @Tipo=t.TipDoc,@Total=d.ImpFac,@Divisa=COALESCE(m.Moneda,d.DivFac),@Nombre=d.NomCli,@Codigo=d.CodCli,@IdCliente=d.IdCliente,@Emision=d.FecEmi,@IdTipo=d.TipDoc
 FROM dbo.DBEmiFac d WITH(UPDLOCK,HOLDLOCK)JOIN dbo.Int_TipDoc t ON t.CodIte=d.TipDoc LEFT JOIN dbo.Int_Moneda m ON CONVERT(nvarchar(50),m.CodEst)=d.DivFac
 WHERE d.Item=@Doc AND d.VersionFila=@Version AND UPPER(d.SitDoc)IN('ACT','ACTIVO') AND t.TipDoc IN('FT','BV');
 IF @Tipo IS NULL OR @Total<=0 OR @IdCliente IS NULL RAISERROR('La factura/boleta cambió, no está activa o no tiene cliente. Recargue.',16,1);
 IF @Fecha<CONVERT(date,@Emision) RAISERROR('La solicitud no puede preceder a la emisión.',16,1);
 IF @Clase='DES NC' AND(@Monto<=0 OR @Monto>=@Total) RAISERROR('El descuento parcial debe ser mayor que cero y menor que el total.',16,1);
 IF @Clase<>'DES NC' SET @Monto=@Total;
 SELECT @NombreResponsable=Descrip FROM dbo.IntRespQuerie WITH(HOLDLOCK)WHERE Item=@Responsable AND Active=1;
 SELECT @NombreServicio=Descrip FROM dbo.IntServFac WITH(HOLDLOCK)WHERE Item=@Servicio AND Active=1;
 IF @NombreResponsable IS NULL OR @NombreServicio IS NULL RAISERROR('Seleccione responsable y servicio activos.',16,1);
 IF @Otro<>0 AND NOT EXISTS(SELECT 1 FROM dbo.OtrosCasosQuery WITH(HOLDLOCK)WHERE Item=@Otro AND Activo=1) RAISERROR('El caso Otro no está disponible.',16,1);
 IF @Solicitud.value('count(/solicitud/g)','int')>6 OR EXISTS(SELECT Numero FROM(SELECT n.value('@Numero','nvarchar(10)') Numero FROM @Solicitud.nodes('/solicitud/g')x(n))g GROUP BY Numero HAVING COUNT(*)>1) RAISERROR('Las guías deben ser distintas; máximo seis.',16,1);
 IF EXISTS(SELECT 1 FROM @Solicitud.nodes('/solicitud/g')x(n)WHERE NOT EXISTS(SELECT 1 FROM dbo.DbServGui s JOIN dbo.DbGuias g ON g.Item=s.IdGuia WHERE s.IdFactura=@Doc AND g.NroGui=n.value('@Numero','nvarchar(10)'))) RAISERROR('Una guía no pertenece al documento seleccionado.',16,1);
 IF @Numero>=999999999 RAISERROR('Serie de Queries agotada.',16,1);SET @Numero=@Numero+1;UPDATE dbo.SeriesQueries SET Numero=@Numero WHERE Serie='Q';SET @Asiento='Q'+RIGHT('000000000'+CONVERT(varchar(9),@Numero),9);
 SELECT @Usuario=Usuari FROM dbo.Int_Usuario WHERE ITEM=@Actor;
 INSERT dbo.DbQuerie(Asient,FecSol,CliDoc,TipDoc,NumDoc,CodCli,FecEmi,MonSol,DivSol,RcSol,SolTer,SoAnl1,SoAnl2,Solib1,Solib2,SolCli,SolRuc,SolDir,SolTel,SolDes,SolRes,SolEnc,SitDoc,TipQue,Estado,ORIGEN,TipCas,Proble,FecIni,FecVis,IdCliente,IdFactura,IdCreador,IdResponsable,IdServicio,ServicioSolicitud,OpcionesSolicitud,SolicitudOriginal,ClaveSolicitud,SolAdi)
 VALUES(@Asiento,@Fecha,@Nombre,@IdTipo,(SELECT NumDoc FROM dbo.DBEmiFac WHERE Item=@Doc),@Codigo,@Emision,@Total,@Divisa,@Recargo,@Solicitud.value('(/solicitud/@Territorio)[1]','nvarchar(150)'),CASE WHEN @Clase='ANL NC'THEN 1 ELSE 0 END,CASE WHEN @Clase='ANL ANL'THEN 1 ELSE 0 END,@Solicitud.value('(/solicitud/@SiLibGuia)[1]','tinyint'),@Solicitud.value('(/solicitud/@NoLibGuia)[1]','tinyint'),@Solicitud.value('(/solicitud/@ClientePropuesto)[1]','nvarchar(255)'),@Solicitud.value('(/solicitud/@RucPropuesto)[1]','nvarchar(11)'),@Solicitud.value('(/solicitud/@DireccionPropuesta)[1]','nvarchar(255)'),@Solicitud.value('(/solicitud/@TelefonoPropuesto)[1]','nvarchar(20)'),@Monto,@NombreResponsable,@Usuario,'ACT',@Clase,'Abierto',@Usuario,@Caso,COALESCE((SELECT Descripcion FROM dbo.OtrosCasosQuery WHERE Item=@Otro),@Caso),@Inicio,@Visto,@IdCliente,@Doc,@Actor,@Responsable,@Servicio,@NombreServicio,@Solicitud,@Solicitud,@Clave,@Solicitud.value('(/solicitud/@Adicional)[1]','nvarchar(255)'));
 SET @Id=SCOPE_IDENTITY();
 INSERT dbo.GuiasQuery SELECT @Id,g.Item,g.NroGui FROM @Solicitud.nodes('/solicitud/g')x(n)JOIN dbo.DbGuias g ON g.NroGui=n.value('@Numero','nvarchar(10)');
 INSERT dbo.HistorialQueries(IdQuery,IdUsuario,Accion,Detalle)VALUES(@Id,@Actor,'CREAR',@Clase+' / '+@Caso);
 COMMIT;SELECT ITEM,Asient,VersionFila FROM dbo.DbQuerie WHERE ITEM=@Id;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @Error nvarchar(2048);SET @Error=ERROR_MESSAGE();RAISERROR('%s',16,1,@Error);END CATCH;
END
GO
CREATE PROCEDURE dbo.CobGuardarCatalogo @Actor int,@Tipo int,@Item int=0,@Version binary(8)=NULL,@Codigo nvarchar(20)='',@Descripcion nvarchar(150),@Baja bit=0
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int,@Modulo int;SET @Modulo=CASE WHEN @Tipo=7 THEN 16 ELSE 12 END;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF @Tipo NOT IN(4,5,7)RAISERROR('Catálogo inválido.',16,1);
 IF NULLIF(LTRIM(RTRIM(@Descripcion)),'')IS NULL OR LEN(@Descripcion)>50 RAISERROR('La descripción debe tener entre 1 y 50 caracteres.',16,1);
 BEGIN TRANSACTION;
 IF @Tipo=4 BEGIN
 IF @Item=0 BEGIN IF @Baja=1 RAISERROR('Caso inexistente.',16,1);INSERT dbo.Int_Casos(NomCas,ActEst)VALUES(@Descripcion,1);SET @Item=SCOPE_IDENTITY();END
 ELSE BEGIN UPDATE dbo.Int_Casos SET ActEst=CASE WHEN @Baja=1 THEN 0 ELSE 1 END WHERE CodCas=@Item AND VersionFila=@Version;IF @@ROWCOUNT<>1 RAISERROR('El concepto cambió. Recargue.',16,1);END;END
 ELSE IF @Tipo=5 BEGIN
 IF @Item=0 BEGIN IF @Baja=1 OR NULLIF(@Codigo,'')IS NULL OR LEN(@Codigo)>20 RAISERROR('Ingrese un código válido.',16,1);INSERT dbo.Int_StatusCas(Status,DesSta,ActEst)VALUES(@Codigo,@Descripcion,1);SET @Item=SCOPE_IDENTITY();END
 ELSE BEGIN IF EXISTS(SELECT 1 FROM dbo.Int_StatusCas WHERE CodSta=@Item AND Status='CCR')AND @Baja=1 RAISERROR('El estado de cierre CCR debe permanecer activo.',16,1);UPDATE dbo.Int_StatusCas SET ActEst=CASE WHEN @Baja=1 THEN 0 ELSE 1 END WHERE CodSta=@Item AND VersionFila=@Version;IF @@ROWCOUNT<>1 RAISERROR('El estado cambió. Recargue.',16,1);END;END
 ELSE BEGIN IF @Item<>0 OR @Baja=1 RAISERROR('Operación de motivo inválida.',16,1);IF EXISTS(SELECT 1 FROM dbo.Int_InfoSys WHERE ModAct='MOTCARDEV' AND Infor1=@Descripcion)RAISERROR('El motivo ya existe.',16,1);INSERT dbo.Int_InfoSys(ModAct,Infor1)VALUES('MOTCARDEV',@Descripcion);SET @Item=SCOPE_IDENTITY();END;
 COMMIT;SELECT @Item Item;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.CobCorreosTerritorio @Actor int,@Modulo int,@Guardar bit=0,@IdTerritorio int=0,@Version binary(8)=NULL,@Correos nvarchar(1000)=''
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF @Guardar=1 BEGIN
 IF NOT EXISTS(SELECT 1 FROM dbo.Int_Usuario WHERE ITEM=@Actor AND Access='ADMINISTRADOR'AND ActPas=1 AND SitUsu='Activo')RAISERROR('Solo un administrador puede guardar destinatarios.',16,1);
 BEGIN TRANSACTION;
 IF NOT EXISTS(SELECT 1 FROM dbo.IntTerritorios WHERE Item=@IdTerritorio AND Active=1)RAISERROR('Territorio inexistente o inactivo.',16,1);
 IF EXISTS(SELECT 1 FROM dbo.CorreosTerritorioCobranza WITH(UPDLOCK,HOLDLOCK)WHERE IdTerritorio=@IdTerritorio)BEGIN
 UPDATE dbo.CorreosTerritorioCobranza SET Destinatarios=@Correos WHERE IdTerritorio=@IdTerritorio AND VersionFila=@Version;IF @@ROWCOUNT<>1 RAISERROR('Los destinatarios cambiaron. Recargue.',16,1);END
 ELSE BEGIN IF @Version IS NOT NULL RAISERROR('Los destinatarios cambiaron. Recargue.',16,1);INSERT dbo.CorreosTerritorioCobranza(IdTerritorio,Destinatarios)VALUES(@IdTerritorio,@Correos);END;
 COMMIT;END;
 SELECT t.Item,t.Descrip Territorio,ISNULL(c.Destinatarios,'')Destinatarios,c.VersionFila FROM dbo.IntTerritorios t LEFT JOIN dbo.CorreosTerritorioCobranza c ON c.IdTerritorio=t.Item WHERE t.Active=1 ORDER BY t.Descrip;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
CREATE PROCEDURE dbo.ConsultarQueries @Actor int,@Modulo int=6,@Buscar nvarchar(100)='',@Estado nvarchar(10)='',@Desde datetime=NULL,@Hasta datetime=NULL,@Usuario int=0
AS
BEGIN
 SET NOCOUNT ON;DECLARE @Acceso int;EXEC @Acceso=dbo.VerificarActorQuery @Actor,@Modulo;IF @Acceso<>0 RETURN;
 IF @Desde>@Hasta BEGIN RAISERROR('Rango de fechas inválido.',16,1);RETURN;END;
 SELECT TOP(200) q.*,t.TipDoc TipoDocumento,ISNULL(u.Usuari,q.ORIGEN) Creador,n.NumDoc NotaResolucion
 FROM dbo.DbQuerie q LEFT JOIN dbo.Int_TipDoc t ON t.CodIte=q.TipDoc LEFT JOIN dbo.Int_Usuario u ON u.ITEM=q.IdCreador LEFT JOIN dbo.DbNCredito n ON n.ITEM=q.IdNotaResolucion
 WHERE (@Buscar='' OR CHARINDEX(@Buscar,q.Asient)>0 OR CHARINDEX(@Buscar,q.NumDoc)>0 OR CHARINDEX(@Buscar,q.CliDoc)>0)
 AND (@Estado='' OR (@Estado='Anuladas' AND q.SitDoc='ANL') OR (@Estado=q.Estado AND q.SitDoc='ACT'))
 AND(@Desde IS NULL OR q.FecSol>=@Desde)AND(@Hasta IS NULL OR q.FecSol<DATEADD(day,1,CONVERT(date,@Hasta)))AND(@Usuario=0 OR q.IdCreador=@Usuario)
 ORDER BY q.FecSol DESC,q.ITEM DESC;
END
GO
CREATE PROCEDURE dbo.CobDocumento @Actor int,@Modulo int,@Tipo nvarchar(2),@Numero nvarchar(50)
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 SELECT * FROM dbo.DocumentosGestion WHERE Tipo=@Tipo AND Documento=@Numero AND Tipo IN('FT','BV','ND');
END
GO
CREATE PROCEDURE dbo.CobLeerConfiguracion @Actor int,@Modulo int
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 SELECT Item,Empresa,CuentaSoles,CuentaDolares,Contacto,CorreoConfirmacion,Anexo,DestinoCheques,DestinoFacturas,DestinoCartas,HostSMTP,PuertoSMTP,UsarTLS,UsuarioSMTP,RemitenteSMTP,EquipoSMTP,CONVERT(bit,CASE WHEN ClaveSMTP IS NULL THEN 0 ELSE 1 END)TieneClave,VersionFila FROM dbo.ConfiguracionCobranzas WHERE Item=1;
END
GO
CREATE PROCEDURE dbo.ConsultarCatalogosQuery @Actor int,@Modulo int=5,@Catalogo int
AS
BEGIN SET NOCOUNT ON;DECLARE @Acceso int;EXEC @Acceso=dbo.VerificarActorQuery @Actor,@Modulo;IF @Acceso<>0 RETURN;
 IF @Catalogo=1 SELECT Item,Descrip Descripcion FROM dbo.IntRespQuerie WHERE Active=1 ORDER BY Descrip;
 ELSE IF @Catalogo=2 SELECT Item,Descrip Descripcion FROM dbo.IntServFac WHERE Active=1 ORDER BY Descrip;
 ELSE IF @Catalogo=3 SELECT Item,Descripcion,VersionFila FROM dbo.OtrosCasosQuery WHERE Activo=1 ORDER BY Descripcion;
 ELSE IF @Catalogo=4 SELECT ITEM Item,Usuari Descripcion FROM dbo.Int_Usuario ORDER BY Usuari;
 ELSE RAISERROR('Catálogo inválido.',16,1);
END
GO
CREATE PROCEDURE dbo.CobEnvioEnCurso @Actor int,@Modulo int,@Item int,@Emision tinyint=1
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 SELECT e.Item,e.Estado,e.VersionFila,DATEDIFF(second,e.Inicio,GETDATE()) Segundos
 FROM dbo.EnviosCobranzas e LEFT JOIN dbo.EmisionesCarta c ON c.Item=e.IdEmision
 WHERE e.Estado='Enviando'AND((@Modulo=14 AND e.IdRecojo=@Item)OR(@Modulo=15 AND c.IdCarta=@Item AND c.Numero=@Emision));
END
GO
CREATE PROCEDURE dbo.CobBuscarClientes @Actor int,@Modulo int,@Buscar nvarchar(100)='',@Item int=0
AS
BEGIN SET NOCOUNT ON;DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 SELECT TOP(100)c.Item,c.Asient,c.CLIENT,c.RUCDNI,c.DIRECC,c.TELEF1,c.Contac,c.IdTerritorio,c.VersionFila FROM dbo.DbClientes c
 WHERE c.Activo=1 AND(@Item=0 OR c.Item=@Item)AND(@Buscar=''OR CHARINDEX(@Buscar,c.CLIENT)>0 OR CHARINDEX(@Buscar,c.Asient)>0 OR CHARINDEX(@Buscar,c.RUCDNI)>0)ORDER BY c.CLIENT,c.Item;
END
GO
CREATE PROCEDURE dbo.ServiciosDocumentoQuery @Actor int,@Modulo int=5,@Documento int
AS
BEGIN
 SET NOCOUNT ON;DECLARE @Acceso int;EXEC @Acceso=dbo.VerificarActorQuery @Actor,@Modulo;IF @Acceso<>0 RETURN;
 SELECT s.NroGui Guia,s.Servic Servicio FROM dbo.DbServGui s WHERE s.IdFactura=@Documento ORDER BY s.Linea,s.Item;
END
GO
CREATE PROCEDURE dbo.CobCerrarRegistro @Actor int,@Modulo int,@Item int,@Version binary(8),@Anular bit=0,@Motivo nvarchar(1000)
AS
BEGIN SET NOCOUNT ON;SET XACT_ABORT ON;BEGIN TRY
 DECLARE @P int;EXEC @P=dbo.CobValidarActor @Actor,@Modulo;IF @P<>0 RETURN;
 IF @Modulo NOT BETWEEN 13 AND 16 OR NULLIF(LTRIM(RTRIM(@Motivo)),'')IS NULL RAISERROR('Indique el motivo o resultado de la gestión.',16,1);
 DECLARE @Estado nvarchar(10);SET @Estado=CASE WHEN @Anular=1 THEN 'Anulado'ELSE 'Cerrado'END;BEGIN TRANSACTION;
 IF @Modulo=13 UPDATE dbo.DbCheques SET EstadoGestion=@Estado,FechaCierre=GETDATE(),IdCerrador=@Actor,MotivoCierre=@Motivo,Estado=CASE WHEN @Anular=1 THEN 'Anulado'ELSE 'Finalizado'END,SitDoc=CASE WHEN @Anular=1 THEN 'ANL'ELSE 'ACT'END WHERE ITEM=@Item AND VersionFila=@Version AND EstadoGestion='Abierto';
 ELSE IF @Modulo=14 UPDATE dbo.DbRecojos SET EstadoGestion=@Estado,FechaCierre=GETDATE(),IdCerrador=@Actor,MotivoCierre=@Motivo,Estado=CASE WHEN @Anular=1 THEN 2 ELSE 0 END WHERE Item=@Item AND VersionFila=@Version AND EstadoGestion='Abierto';
 ELSE IF @Modulo=15 UPDATE dbo.DbsCartas SET EstadoGestion=@Estado,FechaCierre=GETDATE(),IdCerrador=@Actor,MotivoCierre=@Motivo WHERE Item=@Item AND VersionFila=@Version AND EstadoGestion='Abierto';
 ELSE UPDATE dbo.DbDevol SET EstadoGestion=@Estado,FechaCierre=GETDATE(),IdCerrador=@Actor,MotivoCierre=@Motivo,Final=0,Estado=CASE WHEN @Anular=1 THEN 0 ELSE 1 END WHERE Item=@Item AND VersionFila=@Version AND EstadoGestion='Abierto';
 IF @@ROWCOUNT<>1 RAISERROR('El expediente cambió o ya no está abierto. Recargue.',16,1);
 -- Cancelar únicamente trabajos que todavía no comenzaron el transporte.
 UPDATE e SET Estado='Error',Fin=GETDATE(),Detalle=N'Cancelado: el expediente se cerró antes de iniciar el envío.'
 FROM dbo.EnviosCobranzas e LEFT JOIN dbo.EmisionesCarta c ON c.Item=e.IdEmision
 WHERE e.Estado='Pendiente'AND((@Modulo=14 AND e.IdRecojo=@Item)OR(@Modulo=15 AND c.IdCarta=@Item));
 INSERT dbo.HistorialCobranzas(IdUsuario,Modulo,Accion,Detalle,IdCheque,IdRecojo,IdCarta,IdDevolucion)VALUES(@Actor,@Modulo,CASE WHEN @Anular=1 THEN 'ANULAR'ELSE 'CERRAR'END,@Motivo,CASE WHEN @Modulo=13 THEN @Item END,CASE WHEN @Modulo=14 THEN @Item END,CASE WHEN @Modulo=15 THEN @Item END,CASE WHEN @Modulo=16 THEN @Item END);
 COMMIT;SELECT @Item Item;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK;DECLARE @E nvarchar(2048);SET @E=ERROR_MESSAGE();RAISERROR('%s',16,1,@E);END CATCH;
END
GO
-- FIN_ARCHIVO_SQL
