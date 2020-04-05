IF ( SELECT COUNT(*)
     FROM   SYSOBJECTS
     WHERE  NAME LIKE 'Permissao_T'
            AND xtype = 'U'
   ) = 0 
    BEGIN
        CREATE TABLE Permissao_T
            (
              ID_IN INT IDENTITY(1, 1) ,
              Permissao_IN INT NOT NULL ,
              Descricao_VC VARCHAR(256) NOT NULL ,
              Status_BT BIT NOT NULL CONSTRAINT PK_Permissao_T PRIMARY KEY ( Permissao_IN )
            )
    END

GO
IF ( SELECT COUNT(*)
     FROM   SYSOBJECTS
     WHERE  NAME LIKE 'Usuario_T'
            AND xtype = 'U'
   ) = 0 
    BEGIN
        CREATE TABLE Usuario_T
            (
              ID_IN INT IDENTITY(1, 1) ,
              Login_VC VARCHAR(128) NOT NULL ,
              Cargo_IN INT NOT NULL ,
              Nome_VC VARCHAR(128) ,
              Telefone_VC VARCHAR(128) ,
              Email_VC VARCHAR(128) ,
              CONSTRAINT PK_Usuario_T PRIMARY KEY ( ID_IN )
            )
    END

GO
IF ( SELECT COUNT(*)
     FROM   SYSOBJECTS
     WHERE  NAME LIKE 'Cargo_T'
            AND xtype = 'U'
   ) = 0 
    BEGIN
        CREATE TABLE Cargo_T
            (
              ID_IN INT IDENTITY(1, 1) ,
              Descricao_VC VARCHAR(128) NOT NULL ,
              CONSTRAINT PK_Cargo_T PRIMARY KEY ( ID_IN )
            )
    END

GO
IF ( SELECT COUNT(*)
     FROM   SYSOBJECTS
     WHERE  NAME LIKE 'CargoXPermissao_T'
            AND xtype = 'U'
   ) = 0 
    BEGIN
        CREATE TABLE CargoXPermissao_T
            (
              ID_IN INT IDENTITY(1, 1) ,
              Permissao_IN INT NOT NULL ,
              Cargo_IN INT NOT NULL CONSTRAINT PK_CargoXPermissao_T PRIMARY KEY ( ID_IN ) ,
              CONSTRAINT FK_CargoXPermissao_T_Permissao_T FOREIGN KEY ( Permissao_IN ) REFERENCES dbo.Permissao_T ( Permissao_IN ) ,
              CONSTRAINT FK_CargoXPermissao_T_Cargo_T FOREIGN KEY ( Cargo_IN ) REFERENCES Cargo_T ( ID_IN )
            )
    END

GO

--------------------------------------------CRIANDO PERMISSÕES BASICAS INICIAIS---------------------------------------

INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 1 , -- Permissao_IN - int
          'Permitir criar competições. 1 - Permite/0 - Não permite' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 2 , -- Permissao_IN - int
          'Permitir adicionar novos clubes ao sistema. 1 - Permite/0 - Não permite' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 3 , -- Permissao_IN - int
          'Permitir adicionar clubes a um campeonato. 1 - Permite/0 - Não permite' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 4 , -- Permissao_IN - int
          'Permite adicionar novos jogadores. 1 - Permite/0 - Não permite' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )

INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 5 , -- Permissao_IN - int
          'Permite visualizar clubes e jogadores de todos os clubes. 1 - Permite/0 - Não permite' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
--------------------------------------------------------------------------------------------------------------------------
        
--------------------------------------------CARGOS BASICOS INICIAIS-------------------------------------------------------

INSERT  INTO dbo.Cargo_T
        ( Descricao_VC )
VALUES  ( 'Alta Direção'  -- Descricao_VC - varchar(128)
          )
INSERT  INTO dbo.Cargo_T
        ( Descricao_VC )
VALUES  ( 'Gerencia'  -- Descricao_VC - varchar(128)
          )          
INSERT  INTO dbo.Cargo_T
        ( Descricao_VC )
VALUES  ( 'Responsável de Clube'  -- Descricao_VC - varchar(128)
          )

---------------------------------------------------------------------------------------------------------------------------
-------------------------------------------PERMISSÕES DOS CARGOS BASICOS INICIAIS------------------------------------------        
----------------------Alta direção------------------------
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 1, -- Permissao_IN - int
          1  -- Cargo_IN - int
          )
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 2, -- Permissao_IN - int
          1  -- Cargo_IN - int
          )
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 3, -- Permissao_IN - int
          1  -- Cargo_IN - int
          )
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 4, -- Permissao_IN - int
          1  -- Cargo_IN - int
          )
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 5, -- Permissao_IN - int
          1  -- Cargo_IN - int
          )
--------------Gerencia--------------------------------------          
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 2, -- Permissao_IN - int
          2  -- Cargo_IN - int
          )
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 3, -- Permissao_IN - int
          1  -- Cargo_IN - int
          )
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 4, -- Permissao_IN - int
          1  -- Cargo_IN - int
          )
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 5, -- Permissao_IN - int
          1  -- Cargo_IN - int
          )
------------------Responsável---------------------------------
INSERT  INTO dbo.CargoXPermissao_T
        ( Permissao_IN, Cargo_IN )
VALUES  ( 4, -- Permissao_IN - int
          3  -- Cargo_IN - int
          )
    
      
          