CREATE OR ALTER VIEW dbo.vw_TransacoesFinanceiras
AS
SELECT
    T.Id_Transacao,
    T.Numero_Cartao,
    T.Valor_Transacao,
    T.Data_Transacao,
    T.Descricao,
    T.Status_Transacao,
    dbo.fn_CategorizarValor(T.Valor_Transacao) AS Categoria
FROM dbo.Transacao AS T;