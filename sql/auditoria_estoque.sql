-- Auditoria de Estoque em SQL
-- Mesmas análises do script Python, escritas em SQL padrão (testado no SQLite).
-- Para rodar: sqlite3 estoque.db < sql/auditoria_estoque.sql

-- 1. Tabela da contagem de inventário
DROP TABLE IF EXISTS contagem_estoque;
CREATE TABLE contagem_estoque (
    codigo          VARCHAR(10) PRIMARY KEY,
    produto         VARCHAR(100) NOT NULL,
    categoria       VARCHAR(50),
    lote            VARCHAR(10),
    validade        DATE,
    estoque_minimo  INT,
    qtd_sistema     INT,   -- quantidade registrada no ERP
    qtd_fisica      INT,   -- quantidade contada na prateleira
    custo_unitario  DECIMAL(10, 2)
);

INSERT INTO contagem_estoque VALUES
('MED001', 'Dipirona 500mg (cx 10 comp)',      'Analgésico',      'L2401', '2027-08-15',  40, 120, 118,  4.90),
('MED002', 'Paracetamol 750mg (cx 20 comp)',   'Analgésico',      'L2402', '2027-03-10',  30,  85,  85,  6.20),
('MED003', 'Amoxicilina 500mg (cx 21 caps)',   'Antibiótico',     'L2318', '2026-12-20',  25,  32,  28, 18.50),
('MED004', 'Omeprazol 20mg (cx 28 caps)',      'Gastrointestinal','L2405', '2027-11-30',  30,  60,  60,  9.80),
('MED005', 'Soro Fisiológico 0,9% 500ml',      'Hidratação',      'L2420', '2027-05-05', 100, 240, 226,  5.40),
('MED006', 'Losartana 50mg (cx 30 comp)',      'Cardiovascular',  'L2390', '2027-01-18',  35,  22,  22,  7.30),
('MED007', 'Insulina NPH 10ml',                'Hormônio',        'L2412', '2026-12-05',  15,  18,  17, 29.90),
('MED008', 'Ceftriaxona 1g (frasco)',          'Antibiótico',     'L2399', '2027-02-28',  20,  14,  14, 22.00),
('MAT001', 'Luva de Procedimento M (cx 100)',  'Material',        'L2433', '2029-06-01',  50, 180, 165, 32.00),
('MAT002', 'Seringa 10ml (unid)',              'Material',        'L2441', '2030-01-01', 300, 950, 950,  0.85),
('MAT003', 'Gaze Estéril (pct 10)',            'Material',        'L2425', '2028-09-12', 150, 410, 398,  2.10),
('MAT004', 'Esparadrapo 10cm x 4,5m',          'Material',        'L2410', '2028-04-20',  40,  35,  35,  8.70),
('MED009', 'Dexametasona 4mg/ml (amp)',        'Corticoide',      'L2387', '2027-01-05',  30,  48,  52,  3.60),
('MED010', 'Ondansetrona 8mg (amp)',           'Antiemético',     'L2402', '2027-07-22',  25,  40,  40,  6.90),
('MED011', 'Heparina 5000UI (frasco)',         'Anticoagulante',  'L2379', '2026-12-28',  20,  26,  24, 41.50),
('MED012', 'Captopril 25mg (cx 30 comp)',      'Cardiovascular',  'L2422', '2027-09-14',  30,  70,  70,  3.20);


-- 2. Divergências entre sistema e contagem física, com impacto financeiro
SELECT
    codigo,
    produto,
    qtd_sistema,
    qtd_fisica,
    qtd_fisica - qtd_sistema                              AS divergencia,
    ROUND((qtd_fisica - qtd_sistema) * custo_unitario, 2) AS impacto_reais,
    CASE
        WHEN qtd_fisica < qtd_sistema THEN 'Falta'
        WHEN qtd_fisica > qtd_sistema THEN 'Sobra'
    END                                                   AS situacao
FROM contagem_estoque
WHERE qtd_fisica <> qtd_sistema
ORDER BY impacto_reais;


-- 3. Acurácia do inventário por categoria
SELECT
    categoria,
    COUNT(*)                                                         AS itens,
    SUM(CASE WHEN qtd_fisica = qtd_sistema THEN 1 ELSE 0 END)        AS itens_ok,
    ROUND(100.0 * SUM(CASE WHEN qtd_fisica = qtd_sistema THEN 1 ELSE 0 END)
          / COUNT(*), 1)                                             AS acuracia_pct,
    ROUND(SUM((qtd_fisica - qtd_sistema) * custo_unitario), 2)       AS impacto_reais
FROM contagem_estoque
GROUP BY categoria
ORDER BY acuracia_pct;


-- 4. Estoque crítico: itens abaixo do mínimo e quanto repor
SELECT
    codigo,
    produto,
    estoque_minimo,
    qtd_fisica,
    estoque_minimo - qtd_fisica AS sugestao_reposicao
FROM contagem_estoque
WHERE qtd_fisica < estoque_minimo
ORDER BY sugestao_reposicao DESC;


-- 5. Lotes que vencem nos próximos 60 dias (a partir de 19/11/2026)
SELECT
    codigo,
    produto,
    lote,
    validade,
    CAST(julianday(validade) - julianday('2026-11-19') AS INT) AS dias_para_vencer
FROM contagem_estoque
WHERE julianday(validade) - julianday('2026-11-19') <= 60
ORDER BY dias_para_vencer;
