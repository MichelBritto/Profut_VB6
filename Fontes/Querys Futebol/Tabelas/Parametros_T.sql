CREATE TABLE Parametros_T
    (
      ID_IN INT IDENTITY(1, 1) ,
      Parametro_IN INT NOT NULL,
      Descricao_VC VARCHAR(1024) NOT NULL,
      Valor_IN INT NOT NULL,
      Valor_VC VARCHAR(MAX) NULL ,
      Valor_VB VARBINARY(MAX) NULL
    )
	