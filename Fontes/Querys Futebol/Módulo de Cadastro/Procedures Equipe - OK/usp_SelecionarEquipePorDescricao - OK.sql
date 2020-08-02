ALTER PROCEDURE dbo.usp_SelecionarEquipePorDescricao
    (
      @Descricao_VC VARCHAR(1024)
    )
AS 
    SET @Descricao_VC = '%' + @Descricao_VC + '%'

    SELECT  *
    FROM    dbo.EQUIPE_T
    WHERE   NOME_VC LIKE @Descricao_VC
    
GO
