CREATE TABLE dbo.Cargo_T
    (
      ID_IN INT IDENTITY(1, 1)
                NOT NULL ,
      Descricao_VC VARCHAR(128) NOT NULL ,
      Ativo_BT BIT NULL
    )