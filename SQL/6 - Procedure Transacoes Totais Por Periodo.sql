CREATE OR ALTER PROCEDURE dbo.usp_Transacao_TotalPorPeriodo
(
    @Data_Inicial DATETIME2,
    @Data_Final DATETIME2,
    @Status_Transacao VARCHAR(10) = NULL
)
AS
BEGIN
    SET NOCOUNT ON;

    IF @Data_Inicial >= @Data_Final
    BEGIN
        THROW 50001, 
              'A data inicial deve ser menor que a data final.', 
              1;
    END;

    SELECT
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,
        COUNT(*) AS Quantidade_Transacoes,
        Status_Transacao
    FROM dbo.Transacao
    WHERE Data_Transacao >= @Data_Inicial
      AND Data_Transacao < @Data_Final
      AND (
            @Status_Transacao IS NULL
            OR Status_Transacao = @Status_Transacao
          )
    GROUP BY
        Numero_Cartao,
        Status_Transacao;
END;