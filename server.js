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

// Tabelas permitidas para evitar acesso indevido a outras planilhas
const ALLOWED_TABLES = ['Reinos', 'Cidades', 'Casas', 'Personagens', 'Lore', 'Conflitos', 'Exercito'];

// ─── Core Helpers ──────────────────────────────────────────────────────────

// Lê uma tabela inteira e transforma as linhas em objetos baseados no cabeçalho (Linha 1)
async function getTableData(tableName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${tableName}'!A:Z`, // Lê da coluna A até Z
  });
  const rows = res.data.values || [];
  if (rows.length === 0) return { headers: [], data: [] };

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
    const { headers, data } = await getTableData(req.tableName);
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
    const { headers, data } = await getTableData(req.tableName);
    const idField = headers[0]; 
    const targetId = req.params.id;

    // Encontra a linha onde o ID bate
    const item = data.find(d => String(d[idField]) === String(targetId));
    if (!item) return res.status(404).json({ error: 'Registro não encontrado.' });

    // Mescla os dados antigos com os novos
    const updatedRow = headers.map(h => {
      if (req.body[h] !== undefined) return req.body[h]; // Atualiza se foi enviado
      return item[h]; // Mantém o que já estava
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
app.listen(PORT, () => console.log(`Codex Relacional rodando em http://localhost:${PORT}`));