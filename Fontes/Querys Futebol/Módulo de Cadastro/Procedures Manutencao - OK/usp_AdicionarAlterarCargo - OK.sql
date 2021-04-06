CREATE PROCEDURE dbo.usp_AdicionarAlterarCargo
    (
      @Descricao_VC VARCHAR(1024) ,
      @Ativo_BT BIT ,
      @Cargo_IN INT = NULL
    )
AS 
    IF ISNULL(@Cargo_IN, 0) = 0 
        BEGIN
            IF ( SELECT COUNT(*)
                 FROM   dbo.Cargo_T
                 WHERE  Descricao_VC = @Descricao_VC
               ) = 0 
                BEGIN
					--DESSA FORMA GARANTO QUE NAO SERÁ ADICIONADO NENHUM DUPLICADO
                    INSERT  INTO dbo.Cargo_T
                            ( Descricao_VC, Ativo_BT )
                    VALUES  ( @Descricao_VC, -- Descricao_VC - varchar(128)
                              @Ativo_BT  -- Ativo_BT - bit
                              )
                END
        END
    ELSE 
        BEGIN
            UPDATE  dbo.Cargo_T
            SET     Ativo_BT = @Ativo_BT ,
                    Descricao_VC = @Descricao_VC
            WHERE   ID_IN = @Cargo_IN
        END
GO

INSERT INTO dbo.Cargo_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Administrador', -- Descricao_VC - varchar(128)
          1  -- Ativo_BT - bit
          )
INSERT INTO dbo.Cargo_T
        ( Descricao_VC, Ativo_BT )
VALUES  ( 'Responsável de Clube', -- Descricao_VC - varchar(128)
          1  -- Ativo_BT - bit
          )          
