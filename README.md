# 📊 Dashboard Operacional — TERCEIROS

Dashboard React para visualização dos dados operacionais da planilha TERCEIROS.xlsx, com sincronização via Google Apps Script.

---

## 📁 Estrutura do projeto

```
dashboard-terceiros/
├── src/
│   ├── App.jsx          # Dashboard principal (todas as abas)
│   ├── main.jsx         # Entrada React
│   └── index.css        # Tailwind CSS
├── CodigoPlanilha.gs    # Script para Google Apps Script (API da planilha)
├── index.html
├── vite.config.js
├── tailwind.config.js
├── postcss.config.js
└── package.json
```

---

## 🚀 Como rodar localmente

```bash
# 1. Instalar dependências
npm install

# 2. Rodar em desenvolvimento
npm run dev

# 3. Acessar no navegador
# http://localhost:5173
```

---

## 🔌 Sincronização com a planilha (Google Apps Script)

O arquivo `CodigoPlanilha.gs` cria uma API JSON a partir da planilha Google Sheets.

### Passos para ativar:

1. Abra sua planilha no Google Sheets
2. Vá em **Extensões → Apps Script**
3. Apague o código padrão e cole o conteúdo de `CodigoPlanilha.gs`
4. Clique em **Executar `testar`** para verificar se as abas são lidas corretamente
5. Clique em **Implantar → Nova implantação**
   - Tipo: `Aplicativo da Web`
   - Executar como: `Eu`
   - Acesso: `Qualquer pessoa`
6. Copie a URL gerada

### Conectar o dashboard à API:

No arquivo `src/App.jsx`, localize a linha com `APPS_SCRIPT_URL` e substitua pela URL copiada:

```js
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/SEU_ID_AQUI/exec';
```

Com isso, o dashboard buscará os dados automaticamente e atualizará a cada **5 minutos**.

---

## 📦 Build para produção

```bash
npm run build
```

Os arquivos finais estarão na pasta `dist/` — prontos para hospedar no **GitHub Pages**, **Vercel** ou **Netlify** (todos gratuitos).

### Deploy no GitHub Pages (gratuito):

```bash
# Instalar gh-pages
npm install --save-dev gh-pages

# Adicionar ao package.json:
# "homepage": "https://SEU_USUARIO.github.io/dashboard-terceiros",
# "scripts": { "deploy": "gh-pages -d dist" }

# Build + deploy
npm run build
npm run deploy
```

---

## 📋 Abas do dashboard

| Aba | Origem na planilha |
|---|---|
| 📊 SM Matriz vs Filial | SM MATRIZ + SM FILIAL |
| 👷 Ajudantes PX | Ajudantes PX |
| 🚫 Lista Negra | Motoristas Lista Negra |
| ❌ 3º Cancelado | 3º CANCELADO |
| ⚠️ 3º Reprovado | 3º REPROVADO |
| 📋 3º RDO | 3º RDO |
| 🚨 Pronta Resposta | PRONTA RESPOSTA |

---

## 🛠 Tecnologias

- [React 18](https://react.dev/)
- [Vite](https://vitejs.dev/)
- [Tailwind CSS](https://tailwindcss.com/)
- [Recharts](https://recharts.org/)
- [Google Apps Script](https://developers.google.com/apps-script)
