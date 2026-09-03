CREATE OR ALTER FUNCTION dbo.fn_CategorizarValor
(
    @Valor DECIMAL(18,2)
)
RETURNS VARCHAR(20)
AS
BEGIN

    RETURN
        CASE
            WHEN @Valor > 2000 THEN 'Premium'
            WHEN @Valor >= 1000 THEN 'Alta'
            WHEN @Valor >= 500 THEN 'Média'
            ELSE 'Baixa'
        END;

END;