-- Relatórios de atendimento ao cliente em SQL.
-- Cria uma tabela de chamados, insere dados de exemplo e calcula
-- o desempenho por agente (tempo médio de resposta e satisfação).

-- 1. Criação da tabela de chamados de atendimento
CREATE TABLE chamados_suporte (
    id_chamado INT PRIMARY KEY,
    id_cliente INT NOT NULL,
    agente VARCHAR(100),
    categoria VARCHAR(50),
    status VARCHAR(20),
    tempo_resposta_minutos INT,
    nota_satisfacao INT,
    data_abertura DATETIME
);

-- 2. Inserção de dados de exemplo
INSERT INTO chamados_suporte (id_chamado, id_cliente, agente, categoria, status, tempo_resposta_minutos, nota_satisfacao, data_abertura) VALUES
(101, 1001, 'Ana Silva', 'Dúvida Técnica', 'Fechado', 15, 5, '2026-10-01 09:30:00'),
(102, 1002, 'Carlos Souza', 'Financeiro', 'Fechado', 45, 4, '2026-10-01 10:15:00'),
(103, 1003, 'Ana Silva', 'Erro de Sistema', 'Em Aberto', NULL, NULL, '2026-10-01 11:00:00'),
(104, 1004, 'Beatriz Lima', 'Dúvida Técnica', 'Fechado', 20, 5, '2026-10-02 14:20:00'),
(105, 1005, 'Carlos Souza', 'Erro de Sistema', 'Fechado', 90, 2, '2026-10-02 16:05:00');

-- 3. Desempenho por agente (somente chamados fechados)
SELECT
    agente,
    COUNT(id_chamado) AS total_chamados,
    AVG(tempo_resposta_minutos) AS tempo_medio_resposta_min,
    AVG(nota_satisfacao) AS media_satisfacao
FROM chamados_suporte
WHERE status = 'Fechado'
GROUP BY agente
ORDER BY media_satisfacao DESC;
