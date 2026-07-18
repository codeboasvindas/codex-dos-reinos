require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Google Sheets Auth ────────────────────────────────────────────────────
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_SHEETS_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_SHEETS_PRIVATE_KEY.replace(/\\n/g, '\n'),
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// ─── Esquema das Tabelas (Auto-provisionamento) ───────────────────────────
// Cada tabela nova ou coluna nova adicionada aqui (ou dinamicamente pelo
// frontend via payload) é criada automaticamente na planilha, sem precisar
// mexer manualmente no Google Sheets.
const TABLE_SCHEMAS = {
  Reinos: ['id_reino', 'nome', 'capital', 'governo', 'lema', 'url_emblema', 'url_img_capital', 'url_armadura_cap', 'desc_armadura_cap', 'url_armadura_geral', 'desc_armadura_geral'],
  Cidades: ['id_cidade', 'id_reino', 'nome', 'tipo', 'populacao', 'Infantaria', 'Cavalaria', 'Marinha', 'navios_guerra', 'navios_patrulha', 'obs', 'id_casa_governante'],
  Casas: ['id_casa', 'id_reino', 'nome', 'status', 'lema', 'id_suserano'],
  Personagens: ['id_personagem', 'id_casa', 'sexo', 'nome', 'titulo', 'idade', 'status_vida', 'id_pai', 'id_conjuge', 'notas', 'url_imagem'],
  Lore: ['id_lore', 'id_reino', 'categoria', 'nome', 'descricao', 'url_imagem'],
  Conflitos: ['id_conflito', 'id_reino', 'nome', 'descricao', 'escopo', 'id_reino_2', 'subtitulo', 'data_periodo'],
  Exercito: ['id_exercito', 'id_reino', 'ramo', 'nome', 'efetivo', 'comandante', 'descricao', 'url_imagem'],
  Registros: ['id_registro', 'id_reino', 'titulo', 'data_periodo', 'descricao', 'url_imagem'],
  Geografia: ['id_geografia', 'id_reino', 'nome', 'descricao', 'url_imagem'],
};

const ALLOWED_TABLES = Object.keys(TABLE_SCHEMAS);

// ─── Core Helpers ──────────────────────────────────────────────────────────

// Cache dos nomes de abas existentes na planilha, para evitar chamadas repetidas
let sheetTitlesCache = null;
async function getSheetTitles(force = false) {
  if (!sheetTitlesCache || force) {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    sheetTitlesCache = meta.data.sheets.map(s => s.properties.title);
  }
  return sheetTitlesCache;
}

// Garante que a aba exista na planilha. Se não existir, cria automaticamente
// com o cabeçalho definido em TABLE_SCHEMAS.
async function ensureSheetExists(tableName) {
  const titles = await getSheetTitles();
  const exists = titles.some(t => t.toLowerCase() === tableName.toLowerCase());
  if (exists) return;

  const headers = TABLE_SCHEMAS[tableName] || ['id'];
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: tableName } } }] },
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${tableName}'!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: [headers] },
  });
  await getSheetTitles(true);
}

// Garante que todas as tabelas conhecidas existam (chamado na subida do servidor)
async function ensureAllSheets() {
  for (const table of ALLOWED_TABLES) {
    await ensureSheetExists(table);
  }
}

// Garante que todas as colunas usadas em um payload existam no cabeçalho da aba.
// Se o frontend mandar um campo novo que ainda não existe na planilha, a coluna
// é criada automaticamente ao final do cabeçalho.
async function ensureHeaders(tableName, headers, payloadKeys) {
  const missing = payloadKeys.filter(k => k && !headers.includes(k));
  if (missing.length === 0) return headers;

  const newHeaders = [...headers, ...missing];
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${tableName}'!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: [newHeaders] },
  });
  return newHeaders;
}

// Lê uma tabela inteira e transforma as linhas em objetos baseados no cabeçalho (Linha 1)
async function getTableData(tableName) {
  await ensureSheetExists(tableName);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${tableName}'!A:Z`, // Lê da coluna A até Z
  });
  const rows = res.data.values || [];
  if (rows.length === 0) return { headers: TABLE_SCHEMAS[tableName] || [], data: [] };

  const headers = rows[0];
  const data = rows.slice(1).map((row, idx) => {
    const obj = { _rowIndex: idx + 2 }; // Linha real na planilha (1 é cabeçalho, 2 é o primeiro dado)
    headers.forEach((header, i) => {
      obj[header] = row[i] || '';
    });
    return obj;
  });

  return { headers, data };
}

// Pega o ID interno da Aba (sheetId) necessário para deletar linhas
async function getSheetId(tableName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = meta.data.sheets.find(s => s.properties.title.toLowerCase() === tableName.toLowerCase());
  return sheet ? sheet.properties.sheetId : null;
}

// Middleware para verificar se a tabela requisitada existe
function checkTable(req, res, next) {
  const table = req.params.table;
  const realName = ALLOWED_TABLES.find(t => t.toLowerCase() === table.toLowerCase());
  if (!realName) return res.status(400).json({ error: 'Tabela não permitida ou inexistente.' });
  req.tableName = realName;
  next();
}

// ─── Rotas da API ──────────────────────────────────────────────────────────

// 1. CARGA INICIAL (Puxa o banco inteiro de uma vez para o frontend)
app.get('/api/db', async (req, res) => {
  try {
    await ensureAllSheets();
    const ranges = ALLOWED_TABLES.map(t => `'${t}'!A:Z`);
    const response = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges
    });

    const db = {};
    response.data.valueRanges.forEach((sheetData, index) => {
      const tableName = ALLOWED_TABLES[index];
      const rows = sheetData.values || [];
      if (rows.length > 0) {
        const headers = rows[0];
        db[tableName.toLowerCase()] = rows.slice(1).map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = row[i] || ''; });
          return obj;
        });
      } else {
        db[tableName.toLowerCase()] = [];
      }
    });

    res.json(db);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// 2. LER UMA TABELA ESPECÍFICA (Ex: GET /api/cidades)
app.get('/api/:table', checkTable, async (req, res) => {
  try {
    const { data } = await getTableData(req.tableName);
    res.json(data);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// 3. CRIAR UM NOVO REGISTRO (Ex: POST /api/personagens)
app.post('/api/:table', checkTable, async (req, res) => {
  try {
    let { headers, data } = await getTableData(req.tableName);
    headers = await ensureHeaders(req.tableName, headers, Object.keys(req.body));

    const idField = headers[0]; // Assume que a primeira coluna é sempre o ID principal (ex: id_personagem)

    // Lógica de Auto-Incremento do ID
    let maxId = 0;
    data.forEach(item => {
      const val = parseInt(item[idField]);
      if (!isNaN(val) && val > maxId) maxId = val;
    });

    // Insere o novo ID gerado no body da requisição
    const newData = { ...req.body, [idField]: maxId + 1 };

    // Monta o array na mesma ordem das colunas da planilha
    const rowToInsert = headers.map(h => newData[h] !== undefined ? newData[h] : '');

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${req.tableName}'!A:Z`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [rowToInsert] },
    });

    res.json({ ok: true, id_gerado: maxId + 1, record: newData });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// 4. ATUALIZAR UM REGISTRO EXISTENTE (Ex: PUT /api/casas/5)
app.put('/api/:table/:id', checkTable, async (req, res) => {
  try {
    let { headers, data } = await getTableData(req.tableName);
    headers = await ensureHeaders(req.tableName, headers, Object.keys(req.body));

    const idField = headers[0];
    const targetId = req.params.id;

    // Encontra a linha onde o ID bate
    const item = data.find(d => String(d[idField]) === String(targetId));
    if (!item) return res.status(404).json({ error: 'Registro não encontrado.' });

    // Mescla os dados antigos com os novos
    const updatedRow = headers.map(h => {
      if (req.body[h] !== undefined) return req.body[h]; // Atualiza se foi enviado
      return item[h] !== undefined ? item[h] : ''; // Mantém o que já estava
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${req.tableName}'!A${item._rowIndex}:Z${item._rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] },
    });

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// 5. DELETAR UM REGISTRO (Ex: DELETE /api/lore/12)
app.delete('/api/:table/:id', checkTable, async (req, res) => {
  try {
    const { headers, data } = await getTableData(req.tableName);
    const idField = headers[0];
    const targetId = req.params.id;

    const item = data.find(d => String(d[idField]) === String(targetId));
    if (!item) return res.status(404).json({ error: 'Registro não encontrado.' });

    const sheetId = await getSheetId(req.tableName);
    const rowIndexZeroBased = item._rowIndex - 1; // API do Sheets começa em 0 para deleção estrutural

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: rowIndexZeroBased,
              endIndex: rowIndexZeroBased + 1
            }
          }
        }]
      }
    });

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// Serve o mapa local
app.get('/mapa.png', (req, res) => res.sendFile(path.join(__dirname, 'mapa.png')));

// Fallback para servir o HTML
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;

ensureAllSheets()
  .catch(e => console.error('Falha ao garantir abas da planilha na subida:', e.message))
  .finally(() => {
    app.listen(PORT, () => console.log(`Codex Relacional rodando em http://localhost:${PORT}`));
  });
