CREATE OR ALTER FUNCTION dbo.fn_TransacoesCategorizadas
(
    @Data_Inicial DATETIME2,
    @Data_Final DATETIME2
)
RETURNS TABLE
AS
RETURN
(
    SELECT
        T.Id_Transacao,
        T.Numero_Cartao,
        T.Valor_Transacao,
        T.Data_Transacao,
        T.Descricao,
        T.Status_Transacao,
        dbo.fn_CategorizarValor(T.Valor_Transacao) AS Categoria
    FROM dbo.Transacao AS T
    WHERE T.Data_Transacao >= @Data_Inicial
      AND T.Data_Transacao < @Data_Final
);