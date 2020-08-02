CREATE TABLE PosicaoAtleta_T
    (
      ID_IN INT IDENTITY(1, 1) ,
      Descricao_VC VARCHAR(1024) ,
      Ativo_BT BIT
    )
    
GO

INSERT  INTO dbo.PosicaoAtleta_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Goleiro', -- Descricao_VC - varchar(1024)
          1  -- Ativo_BT - bit
          )
INSERT  INTO dbo.PosicaoAtleta_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Lateral direito', -- Descricao_VC - varchar(1024)
          1  -- Ativo_BT - bit
          )
INSERT  INTO dbo.PosicaoAtleta_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Lateral esquerdo', -- Descricao_VC - varchar(1024)
          1  -- Ativo_BT - bit
          )
INSERT  INTO dbo.PosicaoAtleta_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Zagueiro', -- Descricao_VC - varchar(1024)
          1  -- Ativo_BT - bit
          )
INSERT  INTO dbo.PosicaoAtleta_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Volante', -- Descricao_VC - varchar(1024)
          1  -- Ativo_BT - bit
          )
INSERT  INTO dbo.PosicaoAtleta_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Meia', -- Descricao_VC - varchar(1024)
          1  -- Ativo_BT - bit
          )
INSERT  INTO dbo.PosicaoAtleta_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Atacante', -- Descricao_VC - varchar(1024)
          1  -- Ativo_BT - bit
          )          
