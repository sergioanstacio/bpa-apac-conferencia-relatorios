# 📋 BPA/APAC — Gerador de Relatórios SIA/SUS

Sistema web para importar arquivos de produção do SIA/SUS (BPA e APAC) e gerar relatórios Excel/PDF com conferência de faturamento.

## Funcionalidades

- **BPA** — Relatório por paciente → data → procedimento com valor SIGTAP
- **APAC** — Relatório por paciente com procedimentos e totais
- **Conferência** — Compara arquivo PA/MAI com PDF de faturamento e aponta divergências
- Exportação em **Excel** e/ou **PDF**
- Valores baseados na tabela **SIGTAP/DATASUS** oficial

---

## 🚀 Como publicar no GitHub e rodar online (grátis)

### Passo 1 — Criar repositório no GitHub

1. Acesse [github.com](https://github.com) e faça login
2. Clique em **"New repository"**
3. Nome: `bpa-relatorios` (ou outro de sua escolha)
4. Marque **Public** ou **Private**
5. Clique em **"Create repository"**

### Passo 2 — Enviar os arquivos

Pelo terminal (após instalar Git):
```bash
git init
git add .
git commit -m "Versão inicial"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/bpa-relatorios.git
git push -u origin main
```

Ou arraste os arquivos diretamente pela interface web do GitHub.

---

## 🌐 Hospedagem gratuita no Render.com

1. Acesse [render.com](https://render.com) e crie conta (gratuito)
2. Clique em **"New Web Service"**
3. Conecte seu repositório GitHub
4. Configure:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --workers 2 --timeout 300`
5. Clique em **"Create Web Service"**

Em 2-3 minutos seu sistema estará online em uma URL como:
`https://bpa-relatorios.onrender.com`

---

## 🏃 Rodar localmente

```bash
# Instalar dependências
pip install -r requirements.txt

# Iniciar servidor
python app.py
```

Acesse: **http://localhost:5000**

---

## Estrutura do projeto

```
bpa-relatorios/
├── app.py                 # Servidor Flask (web)
├── gerar_relatorio.py     # Lógica principal de geração
├── requirements.txt       # Dependências Python
├── Procfile               # Para Render/Heroku
├── sigtap_cache.zip       # Tabela SIGTAP pré-carregada
├── templates/
│   └── index.html         # Interface web
├── uploads/               # Arquivos enviados (criado automaticamente)
└── outputs/               # Relatórios gerados (criado automaticamente)
```

## Extensões de arquivo suportadas

| Tipo | Extensões |
|------|-----------|
| BPA  | .MAI .JAN .FEV .MAR .ABR .JUN .JUL .AGO .SET .OUT .NOV .DEZ |
| APAC | .ABR .JAN .FEV .MAR .MAI .JUN .JUL .AGO .SET .OUT .NOV .DEZ |

---

Desenvolvido com Python · Flask · openpyxl · reportlab · pdfplumber  
Tabela SIGTAP/DATASUS · SIA/SUS
