-- ============================================================
-- BPA/APAC — Schema Supabase
-- Cole este SQL no Supabase → SQL Editor → Run
-- ============================================================

-- Tabela de relatórios gerados
CREATE TABLE IF NOT EXISTS relatorios (
    id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    criado_em     TIMESTAMPTZ DEFAULT NOW(),
    tipo          TEXT NOT NULL,          -- 'BPA' | 'APAC' | 'CONFERENCIA'
    nome_arquivo  TEXT,
    estabelecimento TEXT,
    cnes          TEXT,
    competencia   TEXT,
    n_pacientes   INTEGER DEFAULT 0,
    n_procedimentos INTEGER DEFAULT 0,
    n_datas       INTEGER DEFAULT 0,
    n_divergencias INTEGER DEFAULT 0,
    n_pac_div     INTEGER DEFAULT 0,
    status        TEXT DEFAULT 'concluido',
    divergencias_json TEXT              -- JSON com resumo dos tipos
);

-- Tabela de divergências (conferência PA vs PDF)
CREATE TABLE IF NOT EXISTS divergencias (
    id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    relatorio_id  UUID REFERENCES relatorios(id) ON DELETE CASCADE,
    criado_em     TIMESTAMPTZ DEFAULT NOW(),
    nome_paciente TEXT,
    cod_procedimento TEXT,
    desc_procedimento TEXT,
    qtd_pdf       INTEGER DEFAULT 0,
    qtd_mai       INTEGER DEFAULT 0,
    diferenca     INTEGER DEFAULT 0,
    tipo          TEXT   -- 'QTD DIFERENTE' | 'SOMENTE NO PDF' | 'SOMENTE NO MAI' etc.
);

-- Índices para performance
CREATE INDEX IF NOT EXISTS idx_relatorios_tipo      ON relatorios(tipo);
CREATE INDEX IF NOT EXISTS idx_relatorios_criado_em ON relatorios(criado_em DESC);
CREATE INDEX IF NOT EXISTS idx_divergencias_rel_id  ON divergencias(relatorio_id);
CREATE INDEX IF NOT EXISTS idx_divergencias_paciente ON divergencias(nome_paciente);

-- View para dashboard
CREATE OR REPLACE VIEW dashboard_resumo AS
SELECT
    DATE_TRUNC('month', criado_em) AS mes,
    tipo,
    COUNT(*)                        AS total_relatorios,
    SUM(n_pacientes)                AS total_pacientes,
    SUM(n_procedimentos)            AS total_procedimentos,
    SUM(n_divergencias)             AS total_divergencias
FROM relatorios
GROUP BY 1, 2
ORDER BY 1 DESC, 2;

-- RLS (Row Level Security) — descomente se quiser autenticação por usuário
-- ALTER TABLE relatorios  ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE divergencias ENABLE ROW LEVEL SECURITY;
-- CREATE POLICY "acesso_publico" ON relatorios  FOR ALL USING (true);
-- CREATE POLICY "acesso_publico" ON divergencias FOR ALL USING (true);


-- ============================================================
-- AUTENTICAÇÃO — Tabela de usuários
-- ============================================================

CREATE TABLE IF NOT EXISTS usuarios (
    id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    criado_em    TIMESTAMPTZ DEFAULT NOW(),
    email        TEXT UNIQUE NOT NULL,
    nome         TEXT NOT NULL,
    senha_hash   TEXT NOT NULL,
    ativo        BOOLEAN DEFAULT TRUE,
    perfil       TEXT DEFAULT 'usuario'   -- 'admin' | 'usuario'
);

-- Inserir usuário admin padrão (senha: admin123 — troque após o primeiro login)
-- Gere o hash com: python3 -c "from werkzeug.security import generate_password_hash; print(generate_password_hash('admin123'))"
-- Substitua o hash abaixo pelo gerado:
INSERT INTO usuarios (email, nome, senha_hash, perfil)
VALUES (
    'admin@seuestabelecimento.com.br',
    'Administrador',
    'pbkdf2:sha256:600000$placeholder$hash',   -- <-- SUBSTITUIR
    'admin'
) ON CONFLICT (email) DO NOTHING;
