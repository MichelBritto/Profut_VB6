-- Create Table --

CREATE TABLE ESTADOS_T (
    ID_IN       INT IDENTITY(1,1),
    CODIGO_UF_IN INT ,
    NOME_VC     VARCHAR(128) ,
    UF_CH       CHAR(2),
)

-- Insert Data --

Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (12, 'Acre', 'AC')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (27, 'Alagoas', 'AL')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (16, 'Amapá', 'AP')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (13, 'Amazonas', 'AM')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (29, 'Bahia', 'BA')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (23, 'Ceará', 'CE')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (53, 'Distrito Federal', 'DF')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (32, 'Espírito Santo', 'ES')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (52, 'Goiás', 'GO')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (21, 'Maranhão', 'MA')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (51, 'Mato Grosso', 'MT')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (50, 'Mato Grosso do Sul', 'MS')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (31, 'Minas Gerais', 'MG')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (15, 'Pará', 'PA')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (25, 'Paraíba', 'PB')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (41, 'Paraná', 'PR')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (26, 'Pernambuco', 'PE')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (22, 'Piauí', 'PI')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (33, 'Rio de Janeiro', 'RJ')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (24, 'Rio Grande do Norte', 'RN')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (43, 'Rio Grande do Sul', 'RS')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (11, 'Rondônia', 'RO')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (14, 'Roraima', 'RR')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (42, 'Santa Catarina', 'SC')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (35, 'São Paulo', 'SP')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (28, 'Sergipe', 'SE')
Insert into ESTADOS_T (CODIGO_UF_IN, NOME_VC, UF_CH) values (17, 'Tocantins', 'TO')