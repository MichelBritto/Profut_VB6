IF ( SELECT COUNT(*)
     FROM   SYSOBJECTS
     WHERE  xtype = 'U'
            AND name = 'CARTEGORIA_T'
   ) = 0 
    BEGIN
        CREATE TABLE CARTEGORIA_T
            (
              ID_IN INT IDENTITY(1,1) ,
              DESCRICAO_VC VARCHAR(128) NOT NULL ,
              IDADEMAXIMA_IN INT NOT NULL ,
              CONSTRAINT PK_CARTEGORIA_T_IN PRIMARY KEY ( ID_IN )
            )
    END

GO


INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 7', -- DESCRICAO_VC - varchar(128)
          7  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 8', -- DESCRICAO_VC - varchar(128)
          8  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 9', -- DESCRICAO_VC - varchar(128)
          9  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 10', -- DESCRICAO_VC - varchar(128)
          10  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 11', -- DESCRICAO_VC - varchar(128)
          11  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 12', -- DESCRICAO_VC - varchar(128)
          12  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 13', -- DESCRICAO_VC - varchar(128)
          13  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 14', -- DESCRICAO_VC - varchar(128)
          14  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 15', -- DESCRICAO_VC - varchar(128)
          15  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 16', -- DESCRICAO_VC - varchar(128)
          16  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 17', -- DESCRICAO_VC - varchar(128)
          17  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 18', -- DESCRICAO_VC - varchar(128)
          18  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 19', -- DESCRICAO_VC - varchar(128)
          19  -- IDADEMAXIMA_IN - int
          )
INSERT  INTO dbo.CARTEGORIA_T
        ( DESCRICAO_VC, IDADEMAXIMA_IN )
VALUES  ( 'SUB 20', -- DESCRICAO_VC - varchar(128)
          20  -- IDADEMAXIMA_IN - int
          )
