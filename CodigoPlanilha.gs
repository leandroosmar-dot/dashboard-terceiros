// ═══════════════════════════════════════════════════════════════════════════════
// DASHBOARD OPERACIONAL — TERCEIROS
// Google Apps Script — API JSON para o Dashboard
//
// COMO USAR:
// 1. Na planilha, clique em Extensões → Apps Script
// 2. Apague o código padrão e cole todo este arquivo
// 3. Clique em Implantar → Nova implantação
// 4. Tipo: Aplicativo da Web
// 5. Executar como: Eu (seu email)
// 6. Quem tem acesso: Qualquer pessoa (para o dashboard acessar)
// 7. Clique em Implantar e copie a URL gerada
// 8. Cole a URL no dashboard (instrução no final deste arquivo)
// ═══════════════════════════════════════════════════════════════════════════════

// ── CONFIGURAÇÃO DAS ABAS ────────────────────────────────────────────────────
const CONFIG = {
  SM_MATRIZ:      { sheet: 'SM MATRIZ',             startRow: 2, cols: { data:'A', quantidade:'B', liberacao:'C' } },
  SM_FILIAL:      { sheet: 'SM FILIAL',             startRow: 2, cols: { data:'A', quantidade:'B', liberacao:'C' } },
  CANCELADOS:     { sheet: '3º CANCELADO',          startRow: 2, cols: { data:'A', cavalo:'B', carreta:'C', modelo:'D', motivo:'E' } },
  REPROVADOS:     { sheet: '3º REPROVADO',          startRow: 2, cols: { data:'A', cavalo:'B', carreta:'C', modelo:'D', motivo:'E' } },
  AJUDANTES:      { sheet: 'Ajudantes  PX',         startRow: 2, cols: { nome:'A', cpf:'B', expiracao:'C', telefone:'D' } },
  LISTA_NEGRA:    { sheet: 'Motoristas Lista Negra',startRow: 2, cols: { nome:'A', placa:'B', telefone:'C', motivo:'D', responsavel:'E', data:'F' } },
  RDO:            { sheet: '3º RDO',                startRow: 2, cols: { data:'A', cavalo:'B', carreta:'C', modelo:'D', nome:'E', valor:'F', vinculo:'G' } },
  PRONTA_RESPOSTA:{ sheet: 'PRONTA RESPOSTA',       startRow: 2, cols: { data:'A', cavalo:'B', carreta:'C', modelo:'D', obs:'E' } },
};

// ── UTILITÁRIOS ──────────────────────────────────────────────────────────────

/**
 * Converte número serial do Excel/Sheets para string de data dd/MM/yyyy
 */
function formatDate(value) {
  if (!value) return '';
  if (value instanceof Date) {
    const d = String(value.getDate()).padStart(2,'0');
    const m = String(value.getMonth()+1).padStart(2,'0');
    const y = value.getFullYear();
    return `${d}/${m}/${y}`;
  }
  if (typeof value === 'string' && value.match(/^\d{2}\/\d{2}\/\d{4}$/)) return value;
  return String(value);
}

/**
 * Lê uma aba e retorna array de objetos com base no CONFIG
 */
function readSheet(ss, configKey) {
  const cfg = CONFIG[configKey];
  try {
    const sheet = ss.getSheetByName(cfg.sheet);
    if (!sheet) {
      Logger.log(`Aba não encontrada: ${cfg.sheet}`);
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < cfg.startRow) return [];

    const data = sheet.getRange(cfg.startRow, 1, lastRow - cfg.startRow + 1, 10).getValues();
    const colMap = cfg.cols;
    const colIdx = {};

    // Converter letra de coluna para índice 0-based
    Object.keys(colMap).forEach(field => {
      const letter = colMap[field].toUpperCase();
      colIdx[field] = letter.charCodeAt(0) - 65; // A=0, B=1, etc.
    });

    const results = [];
    data.forEach(row => {
      // Pular linha se todos os campos estiverem vazios
      const hasData = Object.values(colIdx).some(i => row[i] !== '' && row[i] !== null && row[i] !== undefined);
      if (!hasData) return;

      const obj = {};
      Object.keys(colMap).forEach(field => {
        const idx = colIdx[field];
        let val = row[idx];

        // Formatar datas
        if (val instanceof Date) {
          val = formatDate(val);
        } else if (typeof val === 'number' && field.toLowerCase().includes('data')) {
          // Número serial de data
          const d = new Date(Math.round((val - 25569) * 86400 * 1000));
          val = formatDate(d);
        } else if (val === null || val === undefined) {
          val = '';
        } else {
          val = String(val).trim();
        }

        obj[field] = val;
      });

      // Pular se não tem campo principal preenchido
      const firstField = Object.keys(colMap)[0];
      if (!obj[firstField]) return;

      results.push(obj);
    });

    return results;
  } catch(e) {
    Logger.log(`Erro ao ler aba ${cfg.sheet}: ${e.message}`);
    return [];
  }
}

// ── ENDPOINT PRINCIPAL ────────────────────────────────────────────────────────

/**
 * Endpoint GET — retorna todos os dados como JSON
 * Parâmetro opcional: ?aba=SM_MATRIZ para buscar só uma aba
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const params = e && e.parameter ? e.parameter : {};
    const abaFiltro = params.aba ? params.aba.toUpperCase() : null;

    let resultado = {};

    if (abaFiltro && CONFIG[abaFiltro]) {
      // Retornar só a aba solicitada
      resultado[abaFiltro] = readSheet(ss, abaFiltro);
    } else {
      // Retornar todas as abas
      Object.keys(CONFIG).forEach(key => {
        resultado[key] = readSheet(ss, key);
      });
    }

    // Adicionar metadados
    resultado._meta = {
      atualizado_em: formatDate(new Date()),
      hora: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'),
      total_registros: Object.keys(resultado)
        .filter(k => k !== '_meta')
        .reduce((acc, k) => { acc[k] = resultado[k].length; return acc; }, {})
    };

    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ erro: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── ATUALIZAÇÃO AUTOMÁTICA (opcional) ────────────────────────────────────────

/**
 * Cria gatilho para rodar checkAtualizacoes a cada 5 minutos
 * Execute esta função UMA VEZ manualmente para ativar o gatilho
 */
function criarGatilho() {
  // Remove gatilhos antigos para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAtualizacoes') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Cria novo gatilho a cada 5 minutos
  ScriptApp.newTrigger('checkAtualizacoes')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Gatilho criado: atualização a cada 5 minutos');
}

/**
 * Função chamada pelo gatilho — pode ser usada para cache, log, etc.
 */
function checkAtualizacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const agora = new Date();
  Logger.log(`Verificação automática: ${formatDate(agora)} ${Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH:mm:ss')}`);

  // Opcional: registrar última atualização em uma célula da planilha
  try {
    let logSheet = ss.getSheetByName('_LOG_DASHBOARD');
    if (!logSheet) {
      logSheet = ss.insertSheet('_LOG_DASHBOARD');
      logSheet.getRange(1,1,1,3).setValues([['Data','Hora','Status']]);
    }
    logSheet.appendRow([
      formatDate(agora),
      Utilities.formatDate(agora, Session.getScriptTimeZone(), 'HH:mm:ss'),
      'OK'
    ]);
  } catch(e) {
    Logger.log('Erro no log: ' + e.message);
  }
}

// ── TESTE LOCAL ───────────────────────────────────────────────────────────────

/**
 * Teste — rode esta função no editor para verificar se tudo funciona
 * antes de implantar
 */
function testar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== TESTE DE LEITURA DAS ABAS ===');

  Object.keys(CONFIG).forEach(key => {
    const dados = readSheet(ss, key);
    Logger.log(`${key}: ${dados.length} registros`);
    if (dados.length > 0) {
      Logger.log('  Primeiro: ' + JSON.stringify(dados[0]));
      Logger.log('  Último:   ' + JSON.stringify(dados[dados.length-1]));
    }
  });

  Logger.log('\n=== TESTE CONCLUÍDO ===');
  Logger.log('Se todos os registros > 0, a API está pronta para implantar!');
}

// ═══════════════════════════════════════════════════════════════════════════════
// INSTRUÇÕES PARA CONECTAR AO DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
//
// Após implantar, você terá uma URL assim:
// https://script.google.com/macros/s/SEU_ID_AQUI/exec
//
// Para buscar todos os dados:
// https://script.google.com/macros/s/SEU_ID_AQUI/exec
//
// Para buscar só uma aba:
// https://script.google.com/macros/s/SEU_ID_AQUI/exec?aba=SM_MATRIZ
//
// No dashboard React, substitua os dados estáticos por:
//
//   const [dados, setDados] = useState(null);
//
//   useEffect(() => {
//     const URL = 'https://script.google.com/macros/s/SEU_ID_AQUI/exec';
//     fetch(URL)
//       .then(r => r.json())
//       .then(d => setDados(d))
//       .catch(err => console.error('Erro ao buscar dados:', err));
//
//     // Atualiza a cada 5 minutos
//     const interval = setInterval(() => {
//       fetch(URL).then(r => r.json()).then(d => setDados(d));
//     }, 5 * 60 * 1000);
//
//     return () => clearInterval(interval);
//   }, []);
//
//   if (!dados) return <div>Carregando...</div>;
//
// ═══════════════════════════════════════════════════════════════════════════════
