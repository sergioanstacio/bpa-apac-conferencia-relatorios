"""
Módulo de autenticação via Supabase (tabela usuarios).
Usa werkzeug para hash de senha e flask-login para sessão.
"""

from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin
import supabase_db as db


class Usuario(UserMixin):
    def __init__(self, data: dict):
        self.id     = data['id']
        self.email  = data['email']
        self.nome   = data['nome']
        self.perfil = data.get('perfil', 'usuario')
        self.ativo  = data.get('ativo', True)

    def is_active(self):
        return self.ativo

    def is_admin(self):
        return self.perfil == 'admin'


# ─────────────────────────────────────────────────────────────
# Operações de autenticação
# ─────────────────────────────────────────────────────────────

def buscar_usuario_por_id(uid: str):
    sb = db.get_client()
    if not sb:
        return None
    try:
        res = sb.table('usuarios').select('*').eq('id', uid).single().execute()
        return Usuario(res.data) if res.data else None
    except Exception:
        return None


def buscar_usuario_por_email(email: str):
    sb = db.get_client()
    if not sb:
        return None
    try:
        res = sb.table('usuarios').select('*').eq('email', email.lower().strip()).execute()
        return Usuario(res.data[0]) if res.data else None
    except Exception:
        return None


def autenticar(email: str, senha: str):
    """
    Verifica email + senha.
    Retorna Usuario se válido, None se inválido.
    """
    sb = db.get_client()
    if not sb:
        return None
    try:
        res = sb.table('usuarios') \
                .select('*') \
                .eq('email', email.lower().strip()) \
                .eq('ativo', True) \
                .execute()
        if not res.data:
            return None
        u = res.data[0]
        if check_password_hash(u['senha_hash'], senha):
            return Usuario(u)
        return None
    except Exception as e:
        print(f"[Auth] Erro ao autenticar: {e}")
        return None


def criar_usuario(email: str, nome: str, senha: str, perfil: str = 'usuario'):
    sb = db.get_client()
    if not sb:
        return None, "Supabase não configurado"
    try:
        # Verifica se já existe
        existe = sb.table('usuarios').select('id').eq('email', email.lower()).execute()
        if existe.data:
            return None, "E-mail já cadastrado"
        hash_senha = generate_password_hash(senha)
        res = sb.table('usuarios').insert({
            'email':      email.lower().strip(),
            'nome':       nome.strip(),
            'senha_hash': hash_senha,
            'perfil':     perfil,
        }).execute()
        return Usuario(res.data[0]), None
    except Exception as e:
        return None, str(e)


def alterar_senha(uid: str, senha_atual: str, nova_senha: str):
    sb = db.get_client()
    if not sb:
        return False, "Supabase não configurado"
    try:
        res = sb.table('usuarios').select('senha_hash').eq('id', uid).single().execute()
        if not res.data:
            return False, "Usuário não encontrado"
        if not check_password_hash(res.data['senha_hash'], senha_atual):
            return False, "Senha atual incorreta"
        novo_hash = generate_password_hash(nova_senha)
        sb.table('usuarios').update({'senha_hash': novo_hash}).eq('id', uid).execute()
        return True, "Senha alterada com sucesso"
    except Exception as e:
        return False, str(e)


def listar_usuarios():
    sb = db.get_client()
    if not sb:
        return []
    try:
        res = sb.table('usuarios').select('id,email,nome,perfil,ativo,criado_em') \
                .order('criado_em', desc=True).execute()
        return res.data or []
    except Exception:
        return []


def toggle_ativo(uid: str, ativo: bool):
    sb = db.get_client()
    if not sb:
        return False
    try:
        sb.table('usuarios').update({'ativo': ativo}).eq('id', uid).execute()
        return True
    except Exception:
        return False


def gerar_hash(senha: str) -> str:
    """Utilitário para gerar hash de senha via CLI."""
    return generate_password_hash(senha)
