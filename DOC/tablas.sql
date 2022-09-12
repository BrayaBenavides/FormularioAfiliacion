CREATE DATABASE Colsubsidio
GO
 USE Colsubsidio

 
 CREATE TABLE Empleador 
 (
  Id INT PRIMARY KEY identity(1,1) NOT NULL,
  TrabajadorId INT  NOT NULL,
  TipoBeneficiario INT  NOT NULL,
  Tipo_Novedad INT NULL,
  FechaRadicacion DATE  NOT NULL,
  TipoIdentificacionEmpresa INT NOT NULL,
  Nit INT NULL,
  DigitoVerificacion INT NULL,
  Nombre_RazonSocial VARCHAR(50) NULL,
  Sector INT NULL,
  Sucursal VARCHAR(50) NULL, 
  Direccion VARCHAR(50) NULL,
  Departamento VARCHAR(50) NULL,
  Telefono VARCHAR(50) NULL
 );
 GO

 
 CREATE TABLE Trabajador
 (
  TrabajadorId INT PRIMARY KEY NOT NULL,
  Id INT identity(1,1) NOT NULL,
  TipoBeneficiario INT NOT NULL,
  trabajadorTipoId INT NOT NULL,
  Nombre1 VARCHAR(50) NOT NULL,
  Nombre2 VARCHAR(50) NULL,
  Apellido1 VARCHAR(50) NOT NULL,
  Apellido2 VARCHAR(50) NULL,
  FechaNacimiento DATE NOT NULL,
  IdGenero INT NOT NULL,
  IdNacionalidad INT NULL, 
  IdEstadoCivil INT NOT NULL,
  IdNivelOcupacion INT NOT NULL,
  IdNivelEducativo INT NULL,
  FechaIngresoEmpresa DATE NOT NULL,
  HorasMes INT NOT NULL,
  Trabajador INT NULL,
  SalarioBasico INT NOT NULL,
  Celular VARCHAR(50) NOT NULL,
  EPS VARCHAR(50) NULL,
  AFP VARCHAR(50) NULL,
  Direccion_V VARCHAR(50) NOT NULL,
  Ciudad_V INT NULL,
  Dpto_V INT NULL,
  Zona_V INT NOT NULL,
  Telefono_V VARCHAR(50) NULL,
  Direccion_T VARCHAR(50) NULL,
  Ciudad_T INT NULL,
  Dpto_T INT  NULL,
  Zona_T INT NULL,
  Telefono_T VARCHAR(50) NULL,
  Email VARCHAR(50) NOT NULL
 );
 GO

 CREATE TABLE Conyuge 
 (
  Id INT PRIMARY KEY identity(1,1) NOT NULL,
  TrabajadorId INT NOT NULL,
  TipoBeneficiario INT NOT NULL,
  BeneficiarioTipoId INT NOT NULL,
  BeneficiarioId INT NOT NULL,
  Nombre1 VARCHAR(50) NOT NULL,
  Nombre2 VARCHAR(50) NULL,
  Apellido1 VARCHAR(50) NOT NULL,
  Apellido2 VARCHAR(50) NULL,
  FechaNacimiento DATE NOT NULL,
  IdGenero INT NOT NULL,
  FechaIngresoEmpresa DATE NULL,
  SalarioBasico INT NULL,
  TrabajaConyugue VARCHAR(50) NULL,
  RazonSocial VARCHAR(50) NULL,
  Nit INT NULL,
  RecibeSub VARCHAR(50) NULL,
  CajaSub VARCHAR(50) NULL
 );

 GO
 CREATE TABLE Beneficiario
 (
 Id INT PRIMARY KEY identity(1,1) not NULL,
 TrabajadorId INT  NULL,
 TipoBeneficiario INT  NULL,
 BeneficiarioTipoId INT  NULL,
 BeneficiarioId INT  NULL,
 Nombre1 VARCHAR(50)  NULL,
 Nombre2 VARCHAR(50) NULL,
 Apellido1 VARCHAR(50)  NULL,
 Apellido2 VARCHAR(50) NULL,
 IdParentesco INT  NULL,
 FechaNacimiento DATE  NULL,
 IdGenero INT  NULL
 );
 
 
 

 GO 
 CREATE TABLE Nacionalidad
 (
 Id INT PRIMARY KEY identity(1,1),
 Nomenclatura VARCHAR(50)  NULL,
 Descripcion VARCHAR(50)  NULL,
 Pais VARCHAR(50)  NULL
 );
 GO 
  
 CREATE TABLE Ciudades
 (
 Ciudad_Id INT PRIMARY KEY NOT NULL,
 Nombre VARCHAR(50)  NULL,
 Departamento VARCHAR(50)  NULL,
 EpsSucursalId VARCHAR(50)  NULL,
 CodDepartamento INT  NULL,
 CodMunicipio INT  NULL
 );
 GO 

 CREATE TABLE Empresas
 (
 Id INT PRIMARY KEY identity(1,1) NOT NULL,
 Nit INT NOT NULL,
 Tipo INT NULL,
 Razon VARCHAR(50) NULL,
 DV INT NULL,
 );
 GO


 
 ------------- llaves foraneas ----------
ALTER TABLE Empleador 
ADD FOREIGN KEY (TrabajadorId) REFERENCES Trabajador(TrabajadorId); 
go
 ALTER TABLE Conyuge
ADD FOREIGN KEY (TrabajadorId) REFERENCES Trabajador(TrabajadorId); 
go
ALTER TABLE Beneficiario
ADD FOREIGN KEY (TrabajadorId) REFERENCES Trabajador(TrabajadorId); 
go
