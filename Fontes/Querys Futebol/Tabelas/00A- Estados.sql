-- Create Table --

CREATE TABLE ESTADOS_T
    (
      ID_IN INT IDENTITY(1, 1) ,
      CODIGO_UF_IN INT ,
      NOME_VC VARCHAR(128) ,
      UF_CH CHAR(2),
    )

-- Insert Data --

INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 12, 'Acre', 'AC' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 27, 'Alagoas', 'AL' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 16, 'Amapá', 'AP' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 13, 'Amazonas', 'AM' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 29, 'Bahia', 'BA' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 23, 'Ceará', 'CE' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 53, 'Distrito Federal', 'DF' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 32, 'Espírito Santo', 'ES' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 52, 'Goiás', 'GO' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 21, 'Maranhão', 'MA' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 51, 'Mato Grosso', 'MT' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 50, 'Mato Grosso do Sul', 'MS' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 31, 'Minas Gerais', 'MG' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 15, 'Pará', 'PA' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 25, 'Paraíba', 'PB' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 41, 'Paraná', 'PR' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 26, 'Pernambuco', 'PE' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 22, 'Piauí', 'PI' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 33, 'Rio de Janeiro', 'RJ' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 24, 'Rio Grande do Norte', 'RN' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 43, 'Rio Grande do Sul', 'RS' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 11, 'Rondônia', 'RO' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 14, 'Roraima', 'RR' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 42, 'Santa Catarina', 'SC' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 35, 'São Paulo', 'SP' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 28, 'Sergipe', 'SE' )
INSERT  INTO ESTADOS_T
        ( CODIGO_UF_IN, NOME_VC, UF_CH )
VALUES  ( 17, 'Tocantins', 'TO' )