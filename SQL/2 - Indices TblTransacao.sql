CREATE INDEX IX_Transacao_Data
ON dbo.Transacao (Data_Transacao);

CREATE INDEX IX_Transacao_Cartao_Data
ON dbo.Transacao (Numero_Cartao, Data_Transacao);

CREATE INDEX IX_Transacao_Status_Data
ON dbo.Transacao (Status_Transacao, Data_Transacao);