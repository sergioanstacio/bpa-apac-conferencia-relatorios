import os, uuid, threading, json
from flask import Flask, request, jsonify, send_file, render_template, redirect, url_for, session, flash
from flask_login import LoginManager, login_user, logout_user, login_required, current_user

from gerar_relatorio import (
    gerar_excel, gerar_excel_apac, detectar_tipo_arquivo,
    ler_arquivo, carregar_procedimentos, baixar_sigtap,
    ler_arquivo_apac, conferir_pa_vs_pdf, ler_pdf_faturamento
)
import supabase_db as db
import auth

# ── App setup ────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'bpa-sia-sus-secret-key-2025')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
jobs = {}

# ── Flask-Login ──────────────────────────────────────────────
login_manager = LoginManager(app)
login_manager.login_view = 'login_page'
login_manager.login_message = 'Faça login para acessar o sistema.'

@login_manager.user_loader
def load_user(uid):
    return auth.buscar_usuario_por_id(uid)

# ─────────────────────────────────────────────────────────────
# AUTH ROUTES
# ─────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET'])
def login_page():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    return render_template('login.html', erro=None, email='')


@app.route('/login', methods=['POST'])
def login_post():
    email = request.form.get('email', '').strip()
    senha = request.form.get('senha', '')

    if not db.supabase_ativo():
        # Modo offline: aceita admin/admin123 para desenvolvimento local
        if email == 'admin' and senha == 'admin123':
            from auth import Usuario
            u = Usuario({'id':'local','email':'admin','nome':'Admin Local',
                         'perfil':'admin','ativo':True})
            login_user(u)
            return redirect(url_for('index'))
        return render_template('login.html',
                               erro='Supabase não configurado. Use admin / admin123 localmente.',
                               email=email)

    usuario = auth.autenticar(email, senha)
    if usuario:
        login_user(usuario, remember=True)
        next_page = request.args.get('next', url_for('index'))
        return redirect(next_page)

    return render_template('login.html',
                           erro='E-mail ou senha incorretos.', email=email)


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login_page'))


# ─────────────────────────────────────────────────────────────
# ADMIN ROUTES
# ─────────────────────────────────────────────────────────────

def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_admin():
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated


@app.route('/admin/usuarios')
@login_required
@admin_required
def admin_usuarios():
    usuarios = auth.listar_usuarios()
    return render_template('admin_usuarios.html', usuarios=usuarios,
                           current_user=current_user, msg=None, ok=True)


@app.route('/admin/usuarios/criar', methods=['POST'])
@login_required
@admin_required
def admin_criar_usuario():
    nome  = request.form.get('nome','').strip()
    email = request.form.get('email','').strip()
    senha = request.form.get('senha','')
    perfil= request.form.get('perfil','usuario')
    u, err = auth.criar_usuario(email, nome, senha, perfil)
    usuarios = auth.listar_usuarios()
    return render_template('admin_usuarios.html', usuarios=usuarios,
                           current_user=current_user,
                           msg=err or f'Usuário {nome} criado com sucesso!',
                           ok=err is None)


@app.route('/admin/usuarios/<uid>/desativar', methods=['POST'])
@login_required
@admin_required
def admin_desativar(uid):
    auth.toggle_ativo(uid, False)
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/usuarios/<uid>/ativar', methods=['POST'])
@login_required
@admin_required
def admin_ativar(uid):
    auth.toggle_ativo(uid, True)
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/minha-senha', methods=['POST'])
@login_required
def admin_minha_senha():
    ok, msg = auth.alterar_senha(
        current_user.id,
        request.form.get('senha_atual',''),
        request.form.get('nova_senha',''))
    usuarios = auth.listar_usuarios() if current_user.is_admin() else []
    return render_template('admin_usuarios.html', usuarios=usuarios,
                           current_user=current_user, msg=msg, ok=ok)


# ─────────────────────────────────────────────────────────────
# MAIN ROUTES
# ─────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def index():
    return render_template('index.html',
                           supabase_ativo=db.supabase_ativo(),
                           usuario=current_user)


@app.route('/historico')
@login_required
def historico():
    return jsonify(db.listar_relatorios(100))


@app.route('/historico/<rel_id>/divergencias')
@login_required
def historico_divergencias(rel_id):
    return jsonify(db.buscar_divergencias(rel_id))


@app.route('/stats')
@login_required
def stats():
    return jsonify(db.stats_dashboard())


@app.route('/upload', methods=['POST'])
@login_required
def upload():
    modo   = request.form.get('modo', 'bpa')
    job_id = str(uuid.uuid4())

    if modo == 'conferencia':
        f_pa  = request.files.get('file_pa')
        f_pdf = request.files.get('file_pdf')
        if not f_pa or not f_pdf:
            return jsonify({'error': 'Envie os dois arquivos'}), 400
        pa_path  = os.path.join(UPLOAD_DIR, f'{job_id}_pa{os.path.splitext(f_pa.filename)[1]}')
        pdf_path = os.path.join(UPLOAD_DIR, f'{job_id}.pdf')
        out_path = os.path.join(OUTPUT_DIR, f'{job_id}_conferencia.xlsx')
        f_pa.save(pa_path); f_pdf.save(pdf_path)
        nome_pa = f_pa.filename
        jobs[job_id] = {'status':'running','message':'Iniciando…','file':out_path,
                        'nome':nome_pa,'modo':'conferencia'}

        def run_conf():
            try:
                def cb(msg): jobs[job_id]['message'] = msg
                saida, n_div, n_pac, tipos = conferir_pa_vs_pdf(
                    pa_path, pdf_path, out_path, callback=cb)
                cb("Salvando no Supabase…")
                info_bpa, bpa_i, _, _, _, _ = ler_arquivo(pa_path)
                pdf_data = ler_pdf_faturamento(pdf_path)
                import collections as _c, re, unicodedata
                def norm(s):
                    s = unicodedata.normalize('NFD', s)
                    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
                    return re.sub(r'\s+', ' ', s.upper().strip())
                mai_data = _c.defaultdict(lambda: _c.defaultdict(int))
                for nome, datas in bpa_i.items():
                    for dp in datas.values():
                        for proc, qtd in dp.items(): mai_data[norm(nome)][proc] += qtd
                pdf_by_prefix = {}
                for pn in pdf_data:
                    if pn[:30] not in pdf_by_prefix: pdf_by_prefix[pn[:30]] = pn
                mai_to_pdf = {}
                for mn in mai_data:
                    if mn in pdf_data: mai_to_pdf[mn] = mn
                    elif mn in pdf_by_prefix: mai_to_pdf[mn] = pdf_by_prefix[mn]
                    else:
                        for pn in pdf_data:
                            if pn.startswith(mn) or mn.startswith(pn[:len(mn)]):
                                mai_to_pdf[mn] = pn; break
                divs = []
                for mn, pn in mai_to_pdf.items():
                    pp = pdf_data[pn]; pm = mai_data[mn]
                    for proc in set(pp)|set(pm):
                        qp=pp.get(proc,0); qm=pm.get(proc,0)
                        if qp != qm:
                            divs.append({'nome':mn,'proc':proc,'desc':'','qtd_pdf':qp,'qtd_mai':qm,
                                         'tipo':'QTD DIFERENTE' if qp>0 and qm>0
                                                else('SOMENTE NO PDF' if qm==0 else 'SOMENTE NO MAI')})
                rel_id = db.salvar_conferencia(nome_pa, pdf_path, info_bpa, divs, tipos)
                jobs[job_id].update({'status':'done','message':'Concluído!',
                                     'n_div':n_div,'n_pac':n_pac,'tipos':dict(tipos),'rel_id':rel_id})
            except Exception as e:
                jobs[job_id].update({'status':'error','message':str(e)})
        threading.Thread(target=run_conf, daemon=True).start()

    else:
        f = request.files.get('file')
        if not f: return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        ext     = os.path.splitext(f.filename)[1]
        in_path = os.path.join(UPLOAD_DIR, f'{job_id}{ext}')
        out_xl  = os.path.join(OUTPUT_DIR, f'{job_id}_relatorio.xlsx')
        f.save(in_path)
        tipo_det = detectar_tipo_arquivo(in_path)
        fmt      = request.form.get('fmt', 'excel')
        jobs[job_id] = {'status':'running','message':'Iniciando…','file_xl':out_xl,
                        'file':out_xl,'nome':f.filename,'modo':tipo_det.lower(),'fmt':fmt}

        def run_rel():
            try:
                def cb(msg): jobs[job_id]['message'] = msg
                import collections as _c
                if tipo_det == 'APAC':
                    saida, n = gerar_excel_apac(in_path, out_xl, callback=cb)
                    info_a, ac, ap = ler_arquivo_apac(in_path)
                    pt = _c.defaultdict(int)
                    for procs in ap.values():
                        for proc,_,qtd,_ in procs: pt[proc] += qtd
                    db.salvar_relatorio_apac(f.filename, info_a, n, pt)
                    jobs[job_id].update({'status':'done','message':'Concluído!','n_apac':n,'tipo':'APAC'})
                else:
                    saida, n_pac, n_dat, has_i, has_c = gerar_excel(in_path, out_xl, callback=cb)
                    info_b, bi, bc, dates, _, _ = ler_arquivo(in_path)
                    pt = _c.defaultdict(int)
                    for d in bi.values():
                        for dp in d.values():
                            for p,q in dp.items(): pt[p] += q
                    for procs in bc.values():
                        for p,q in procs.items(): pt[p] += q
                    db.salvar_relatorio_bpa(f.filename, info_b, n_pac, n_dat, has_i, has_c, pt)
                    jobs[job_id].update({'status':'done','message':'Concluído!',
                                         'n_pac':n_pac,'n_dat':n_dat,'has_i':has_i,'has_c':has_c,'tipo':'BPA'})
            except Exception as e:
                jobs[job_id].update({'status':'error','message':str(e)})
        threading.Thread(target=run_rel, daemon=True).start()

    return jsonify({'job_id': job_id})


@app.route('/status/<job_id>')
@login_required
def status(job_id):
    job = jobs.get(job_id, {'status':'error','message':'Job não encontrado'})
    return jsonify({**job, 'job_id': job_id})


@app.route('/download/<job_id>')
@login_required
def download(job_id):
    job = jobs.get(job_id)
    if not job: return 'Job não encontrado', 404
    path = job.get('file_xl') or job.get('file','')
    if not os.path.exists(path): return 'Arquivo não encontrado', 404
    base   = os.path.splitext(job.get('nome','relatorio'))[0]
    sufixo = '_conferencia.xlsx' if job.get('modo')=='conferencia' else '_relatorio.xlsx'
    return send_file(path, as_attachment=True, download_name=f'{base}{sufixo}')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
