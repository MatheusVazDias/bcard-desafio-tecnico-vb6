INSERT INTO dbo.Transacao
(
    Numero_Cartao,
    Valor_Transacao,
    Data_Transacao,
    Descricao,
    Status_Transacao
)
VALUES
-- Cartão 1
('4111111111111111', 150.00, '2026-08-01 09:15:00',
 'Compra supermercado', 'Aprovada'),

('4111111111111111', 750.50, '2026-08-02 14:30:00',
 'Compra eletrônicos', 'Aprovada'),

('4111111111111111', 1250.00, '2026-08-05 18:45:00',
 'Compra notebook', 'Pendente'),

('4111111111111111', 2500.00, '2026-08-10 10:20:00',
 'Compra televisão', 'Aprovada'),

('4111111111111111', 320.75, '2026-08-15 12:10:00',
 'Restaurante', 'Cancelada'),

-- Cartão 2
('5555555555554444', 450.00, '2026-08-03 08:30:00',
 'Posto de combustível', 'Aprovada'),

('5555555555554444', 980.00, '2026-08-07 16:00:00',
 'Compra celular', 'Pendente'),

('5555555555554444', 1750.00, '2026-08-12 11:40:00',
 'Compra notebook', 'Aprovada'),

('5555555555554444', 3200.00, '2026-08-18 20:15:00',
 'Viagem', 'Pendente'),

('5555555555554444', 85.90, '2026-08-22 13:25:00',
 'Farmácia', 'Cancelada'),

-- Cartão 3
('3782822463100050', 200.00, '2026-08-04 10:00:00',
 'Supermercado', 'Aprovada'),

('3782822463100050', 500.00, '2026-08-06 15:20:00',
 'Restaurante', 'Pendente'),

('3782822463100050', 1000.00, '2026-08-09 19:30:00',
 'Compra roupas', 'Aprovada'),

('3782822463100050', 2000.00, '2026-08-14 09:45:00',
 'Compra móveis', 'Aprovada'),

('3782822463100050', 4500.00, '2026-08-20 17:10:00',
 'Pacote de viagem', 'Cancelada'),

-- Cartão 4
('6011111111111117', 49.90, '2026-08-05 07:50:00',
 'Café', 'Aprovada'),

('6011111111111117', 499.99, '2026-08-08 13:15:00',
 'Supermercado', 'Aprovada'),

('6011111111111117', 999.99, '2026-08-13 16:40:00',
 'Eletrônicos', 'Pendente'),

('6011111111111117', 1999.99, '2026-08-17 21:00:00',
 'Móveis', 'Cancelada'),

('6011111111111117', 5000.00, '2026-08-25 10:30:00',
 'Compra internacional', 'Aprovada'),

-- Setembro para testar filtro de período
('4111111111111111', 650.00, '2026-09-01 09:00:00',
 'Supermercado', 'Aprovada'),

('5555555555554444', 1500.00, '2026-09-01 14:20:00',
 'Eletrônicos', 'Pendente'),

('3782822463100050', 2750.00, '2026-09-02 11:15:00',
 'Viagem', 'Aprovada');
GO