CREATE FUNCTION FNC_VALIDARCARTEOGORIAJOGADOR ( @Jogador_IN INT )
RETURNS INT
    BEGIN
        DECLARE @IdadeJogador_IN INT ,
            @CartegoriaRetorno_IN INT
	
        SET @IdadeJogador_IN = DATEDIFF(YEAR, ( SELECT  DATANASCIMENTO_DT
                                                FROM    dbo.JOGADOR_T
                                                WHERE   ID_JOGADOR_IN = @Jogador_IN
                                              ), GETDATE())
	
        SET @CartegoriaRetorno_IN = ( SELECT TOP 1
                                                ID_IN
                                      FROM      dbo.CARTEGORIA_T
                                      WHERE     IDADEMAXIMA_IN = @IdadeJogador_IN
                                    )
	
        RETURN @CartegoriaRetorno_IN
    END