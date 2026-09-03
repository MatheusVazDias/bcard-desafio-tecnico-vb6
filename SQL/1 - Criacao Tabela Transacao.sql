CREATE TABLE dbo.Transacao
(
    Id_Transacao BIGINT IDENTITY(1,1) NOT NULL,
    Numero_Cartao CHAR(16) NOT NULL,
    Valor_Transacao DECIMAL(18,2) NOT NULL,
    Data_Transacao DATETIME2 NOT NULL,
    Descricao VARCHAR(255) NULL,
    Status_Transacao VARCHAR(10) NOT NULL,

    CONSTRAINT PK_Transacao
        PRIMARY KEY CLUSTERED (Id_Transacao),

    CONSTRAINT CK_Transacao_Valor
        CHECK (Valor_Transacao > 0),

    CONSTRAINT CK_Transacao_Cartao
        CHECK (LEN(Numero_Cartao) = 16
           AND Numero_Cartao NOT LIKE '%[^0-9]%'),

    CONSTRAINT CK_Transacao_Status
        CHECK (Status_Transacao IN
              ('Aprovada', 'Pendente', 'Cancelada'))
);