"""
Módulo de integração com Supabase.
Configurar variáveis de ambiente:
  SUPABASE_URL  = https://xxxx.supabase.co
  SUPABASE_KEY  = sua_anon_key
"""

import os
import json
from datetime import datetime

# Carrega .env em desenvolvimento local
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')

_client = None


def get_client():
    global _client
    if _client is None and SUPABASE_URL and SUPABASE_KEY:
        try:
            from supabase import create_client
            _client = create_client(SUPABASE_URL, SUPABASE_KEY)
        except Exception as e:
            print(f"[Supabase] Erro ao conectar: {e}")
    return _client


def supabase_ativo():
    """Retorna True se Supabase está configurado."""
    return bool(SUPABASE_URL and SUPABASE_KEY)


# ─────────────────────────────────────────────────────────────
# Salvar relatório
# ─────────────────────────────────────────────────────────────

def salvar_relatorio_bpa(nome_arquivo, info, n_pac, n_dat, has_i, has_c, proc_totals):
    """Salva metadados de um relatório BPA no Supabase."""
    sb = get_client()
    if not sb:
        return None
    try:
        n_proc = sum(proc_totals.values()) if proc_totals else 0
        data = {
            'tipo':             'BPA',
            'nome_arquivo':     nome_arquivo,
            'estabelecimento':  info.get('nome_cnes', ''),
            'cnes':             info.get('cnes', ''),
            'competencia':      info.get('competencia', ''),
            'n_pacientes':      n_pac,
            'n_procedimentos':  n_proc,
            'n_datas':          n_dat,
            'status':           'concluido',
        }
        res = sb.table('relatorios').insert(data).execute()
        return res.data[0]['id'] if res.data else None
    except Exception as e:
        print(f"[Supabase] Erro ao salvar BPA: {e}")
        return None


def salvar_relatorio_apac(nome_arquivo, info, n_apac, proc_totals):
    """Salva metadados de um relatório APAC no Supabase."""
    sb = get_client()
    if not sb:
        return None
    try:
        n_proc = sum(proc_totals.values()) if proc_totals else 0
        data = {
            'tipo':             'APAC',
            'nome_arquivo':     nome_arquivo,
            'estabelecimento':  info.get('nome', ''),
            'cnes':             info.get('cnes', ''),
            'competencia':      info.get('comp', ''),
            'n_pacientes':      n_apac,
            'n_procedimentos':  n_proc,
            'status':           'concluido',
        }
        res = sb.table('relatorios').insert(data).execute()
        return res.data[0]['id'] if res.data else None
    except Exception as e:
        print(f"[Supabase] Erro ao salvar APAC: {e}")
        return None


def salvar_conferencia(nome_pa, nome_pdf, info, divergencias, tipos):
    """Salva conferência e divergências no Supabase."""
    sb = get_client()
    if not sb:
        return None
    try:
        from collections import Counter
        n_pac_div = len(set(d['nome'] for d in divergencias))
        rel_data = {
            'tipo':             'CONFERENCIA',
            'nome_arquivo':     nome_pa,
            'estabelecimento':  info.get('nome_cnes', info.get('nome', '')),
            'cnes':             info.get('cnes', ''),
            'competencia':      info.get('competencia', info.get('comp', '')),
            'n_divergencias':   len(divergencias),
            'n_pac_div':        n_pac_div,
            'divergencias_json': json.dumps(dict(tipos), ensure_ascii=False),
            'status':           'concluido',
        }
        res = sb.table('relatorios').insert(rel_data).execute()
        if not res.data:
            return None
        rel_id = res.data[0]['id']

        # Salva divergências em lotes de 500
        batch = []
        for d in divergencias:
            batch.append({
                'relatorio_id':      rel_id,
                'nome_paciente':     d['nome'],
                'cod_procedimento':  d['proc'],
                'desc_procedimento': d.get('desc', ''),
                'qtd_pdf':           d['qtd_pdf'],
                'qtd_mai':           d['qtd_mai'],
                'diferenca':         d['qtd_mai'] - d['qtd_pdf'],
                'tipo':              d['tipo'],
            })
            if len(batch) >= 500:
                sb.table('divergencias').insert(batch).execute()
                batch = []
        if batch:
            sb.table('divergencias').insert(batch).execute()

        return rel_id
    except Exception as e:
        print(f"[Supabase] Erro ao salvar conferência: {e}")
        return None


# ─────────────────────────────────────────────────────────────
# Consultas para o dashboard
# ─────────────────────────────────────────────────────────────

def listar_relatorios(limite=50):
    sb = get_client()
    if not sb:
        return []
    try:
        res = (sb.table('relatorios')
                 .select('*')
                 .order('criado_em', desc=True)
                 .limit(limite)
                 .execute())
        return res.data or []
    except Exception as e:
        print(f"[Supabase] Erro ao listar: {e}")
        return []


def buscar_divergencias(relatorio_id):
    sb = get_client()
    if not sb:
        return []
    try:
        res = (sb.table('divergencias')
                 .select('*')
                 .eq('relatorio_id', relatorio_id)
                 .order('nome_paciente')
                 .execute())
        return res.data or []
    except Exception as e:
        print(f"[Supabase] Erro ao buscar divergências: {e}")
        return []


def stats_dashboard():
    sb = get_client()
    if not sb:
        return {}
    try:
        res = sb.table('relatorios').select('tipo', count='exact').execute()
        total = res.count or 0

        bpa   = sb.table('relatorios').select('*', count='exact').eq('tipo','BPA').execute().count or 0
        apac  = sb.table('relatorios').select('*', count='exact').eq('tipo','APAC').execute().count or 0
        conf  = sb.table('relatorios').select('*', count='exact').eq('tipo','CONFERENCIA').execute().count or 0

        return {'total': total, 'bpa': bpa, 'apac': apac, 'conferencia': conf}
    except Exception as e:
        print(f"[Supabase] Erro stats: {e}")
        return {}
