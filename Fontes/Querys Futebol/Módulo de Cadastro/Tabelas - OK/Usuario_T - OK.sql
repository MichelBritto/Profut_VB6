CREATE TABLE dbo.Usuario_T
    (
      ID_IN INT IDENTITY(1, 1)
                NOT NULL ,
      Login_VC VARCHAR(128) NOT NULL ,
      Cargo_IN INT NOT NULL ,
      Nome_VC VARCHAR(128) NULL ,
      Telefone_VC VARCHAR(128) NULL ,
      Email_VC VARCHAR(128) NULL ,
      clube_IN INT NULL
    )
