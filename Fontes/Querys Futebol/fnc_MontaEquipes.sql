CREATE FUNCTION dbo.fnc_MontaEquipes
    (
      @Equipes_VC VARCHAR(MAX)
    )
RETURNS @Empresas_T TABLE
    (
      EQUIPE_IN INT NOT NULL
    )
AS 
    BEGIN
        DECLARE @Empresa_VC VARCHAR(200)
        DECLARE @Letra_VC VARCHAR(200)
        DECLARE @Contador_SI SMALLINT
        DECLARE @TamanhoString_SI SMALLINT	
	
        IF @Equipes_VC IS NULL 
            BEGIN
                INSERT  INTO @Empresas_T
                        ( EQUIPE_IN )
                        SELECT  ID_IN
                        FROM    dbo.EQUIPE_T
                RETURN
            END

        SET @Equipes_VC = LTRIM(RTRIM(@Equipes_VC))
        SET @Empresa_VC = ''
	
        SET @TamanhoString_SI = LEN(@Equipes_VC)
        SET @Contador_SI = 0

        WHILE @Contador_SI <= @TamanhoString_SI 
            BEGIN		

                SET @Letra_VC = CAST(SUBSTRING(@Equipes_VC, @Contador_SI, 1) AS CHAR)		
                IF @Letra_VC = ',' 
                    BEGIN
                        INSERT  INTO @Empresas_T
                                ( EQUIPE_IN )
                        VALUES  ( CAST(@Empresa_VC AS INT) )
                        SET @Empresa_VC = ''
                    
                    END
                ELSE 
                    BEGIN				
                
                        SET @Empresa_VC = @Empresa_VC + LTRIM(RTRIM(@Letra_VC))
                    END
		
                SET @Contador_SI = @Contador_SI + 1
            END
        RETURN	
    END
GO
