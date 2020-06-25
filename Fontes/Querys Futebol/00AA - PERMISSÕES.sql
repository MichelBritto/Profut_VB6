--PERMISSOES CADASTRO DE JOGADOR

----ABRIR TELA
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 1 , -- Permissao_IN - int
          'Acessar Cadastro de Jogador' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        

----INCLUIR NOVOS JOGADORES
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 2 , -- Permissao_IN - int
          'Incluir novos atletas no sistema' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        
----ALTERAR JOGADORES
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 3 , -- Permissao_IN - int
          'Alterar informações de atletas no sistema' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )      

----IMPRIMIR FICHA/CARTEIRINHA

INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 4 , -- Permissao_IN - int
          'Imprimir ficha/carteirinha de atleta' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        

--PERMISSOES CADASTRO DE EQUIPE

----ABRIR TELA
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 5 , -- Permissao_IN - int
          'Acessar Cadastro de Equipe' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        
----INCLUIR NOVAS EQUIPES
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 6 , -- Permissao_IN - int
          'Incluir novas equipes no sistema' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )

--PERMISSOES RELATÓRIO DE JOGADOR


----ABRIR TELA
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 7 , -- Permissao_IN - int
          'Acessar o Relatório de Jogadores' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        
--PERMISSOES MANUTENÇÃO

----ABRIR TELA USUÁRIOS
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 8 , -- Permissao_IN - int
          'Acessar tela de manutenção de usuários' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        
----ADICIONAR/ALTERAR USUÁRIO
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 9 , -- Permissao_IN - int
          'Adicionar/Alterar usuários no sistema' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        
----ABRIR TELA DE PERMISSÃO
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 10 , -- Permissao_IN - int
          'Acesso a tela de Permissões' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        
----ABRIR TELA/ADICIONAR NOVOS CARGOS
INSERT  INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 11 , -- Permissao_IN - int
          'Adicionar/Alterar Cargos no sistema ' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        
--PERMISSÃO DE VISUALIZAÇÃO DE JOGADORES E ADMINISTRAÇÃO DE EQUIPES

----PERMITIR VISUALIZAR TODOS OS JOGADORES
INSERT INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 12 , -- Permissao_IN - int
          'Permitir visualizar todos os jogadores' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )

----PERMITIR VISUALIZAR TODOS OS CLUBES
INSERT INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 13 , -- Permissao_IN - int
          'Permitir visualizar todos os clubes' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
-----PERMITIR ALTERAR CADASTRO DE EQUIPE
INSERT INTO dbo.Permissao_T
        ( Permissao_IN ,
          Descricao_VC ,
          Status_BT
        )
VALUES  ( 14 , -- Permissao_IN - int
          'Permitir alterar cadastro de equipe' , -- Descricao_VC - varchar(256)
          1  -- Status_BT - bit
        )
        

        
