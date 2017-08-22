CREATE TABLE [dbo].[Persona]
(
	[Id] INT NOT NULL PRIMARY KEY, 
    [Nombre] VARCHAR(50) NULL, 
    [Rut] VARCHAR(50) NULL, 
    [Comuna] VARCHAR(50) NULL, 
    [EsPersona] BIT NULL, 
    [Direccion] VARBINARY(50) NULL
)
