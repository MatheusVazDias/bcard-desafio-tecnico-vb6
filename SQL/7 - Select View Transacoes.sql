DECLARE @DataInicial DATETIME2 = '2026-08-01 00:00:00';
DECLARE @DataFinal   DATETIME2 = '2026-09-01 00:00:00';

SELECT *
FROM dbo.vw_TransacoesFinanceiras
WHERE Data_Transacao >= @DataInicial
  AND Data_Transacao < @DataFinal;

/*Escolhi utilizar o filtro assim devido ao campo ser datahora
 dessa forma não é preciso filtrar as transações selecionando hora.
  porém o where pode ser feito com bwtween @DataInicial and @DataFinal*/