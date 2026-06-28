# Gerador de Fluxograma Batch — PRODAM

Ferramenta online para gerar fluxogramas de rotinas batch automaticamente
a partir dos documentos de Característica (.doc/.docx) e Script SQL (.sql).

**Zero instalação. Acessa pelo navegador.**

---

## Como subir no Streamlit Cloud (faça uma vez só)

### 1. Criar repositório no GitHub

1. Acesse [github.com](https://github.com) e faça login
2. Clique no **+** → **New repository**
3. Preencha:
   - **Nome:** `fluxograma-batch` (ou qualquer nome)
   - **Visibility:** Private (recomendado)
4. Clique em **Create repository**

### 2. Subir os arquivos

Na tela do repositório:

1. Clique em **uploading an existing file**
2. Arraste os 2 arquivos:
   - `app.py`
   - `requirements.txt`
3. Clique em **Commit changes**

### 3. Conectar ao Streamlit Cloud

1. Acesse [share.streamlit.io](https://share.streamlit.io)
2. Faça login com sua conta GitHub
3. Clique em **New app**
4. Selecione:
   - **Repository:** seu-usuario/fluxograma-batch
   - **Branch:** main
   - **Main file path:** app.py
5. Clique em **Deploy!**
6. Aguarde 2-3 minutos

### 4. Usar

O Streamlit gera um link tipo:

```
https://seu-usuario-fluxograma-batch.streamlit.app
```

Compartilhe esse link com a equipe. Qualquer pessoa acessa pelo navegador, sem login.

---

## Como a equipe usa

1. Abre o link no navegador
2. Faz upload da **Característica** (.doc ou .docx)
3. Faz upload do **Script SQL** (.sql)
4. Clica em **▶▶ Analisar e gerar fluxograma**
5. O programa lê os documentos, identifica as etapas, e gera o fluxograma
6. Clica em **Baixar SVG** para salvar

---

## O que o programa extrai automaticamente

Do documento de **Característica**:
- Código da rotina (ex: SH07681B)
- Sistema e subsistema
- Denominação do programa
- Horário de execução
- Etapas (SQLPLUS, COMPACTAR, TRANSFERIR, etc.)
- Arquivos de entrada e saída
- Dados de SFTP (IP, porta, pasta, usuário)

Do **Script SQL**:
- Tabelas Oracle utilizadas (FROM, JOIN)
- Arquivo de saída (SPOOL)

---

## Formatos suportados

| Arquivo | Formatos |
|---------|----------|
| Característica | .doc (Word 97-2003), .docx, .txt |
| Script SQL | .sql, .txt |
| Saída | SVG (imagem vetorial) |

---

## Dúvidas

**Erro ao ler arquivo .doc**
O programa extrai texto automaticamente de arquivos .doc binários.
Se falhar, salve o arquivo como .docx no Word e tente novamente.

**O fluxograma não identificou todas as etapas**
O parser procura o padrão "ETAPA XX" no documento. Verifique se o
documento segue o formato padrão da Característica PRODAM.

**Posso usar sem internet?**
Não. O app roda no Streamlit Cloud (nuvem). Precisa de internet.
