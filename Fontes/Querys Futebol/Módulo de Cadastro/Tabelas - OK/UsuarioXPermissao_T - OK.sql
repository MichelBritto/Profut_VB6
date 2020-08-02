CREATE TABLE dbo.UsuarioXPermissao_T
    (
      ID_IN INT IDENTITY(1, 1)
                NOT NULL ,
      Usuario_IN INT NULL ,
      Permissao_IN INT NULL ,
      Status_BT BIT NULL
    )
