/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  CONTROLE DE COMISSÕES — Web App (Apps Script)                     ║
 * ║  Planilha "Fechamentos - Davi"                                     ║
 * ╠══════════════════════════════════════════════════════════════════╣
 * ║  COMO IMPLANTAR (uma vez só):                                      ║
 * ║  1. Abra a planilha → menu Extensões → Apps Script                 ║
 * ║  2. Apague o conteúdo e cole TODO este arquivo                     ║
 * ║  3. Salve (ícone do disquete)                                      ║
 * ║  4. Implantar → Nova implantação → ⚙ Tipo: App da Web              ║
 * ║       • Executar como: Eu (seu e-mail)                             ║
 * ║       • Quem pode acessar: Qualquer pessoa                         ║
 * ║  5. Copie a URL que termina em /exec                               ║
 * ║       (o app já vem com a URL atual pré-preenchida — só troque     ║
 * ║        se gerar uma nova implantação)                              ║
 * ║                                                                    ║
 * ║  Ao ALTERAR este código depois, use:                               ║
 * ║  Implantar → Gerenciar implantações → ✏ Editar → Versão: Nova      ║
 * ║  (assim a MESMA URL /exec passa a rodar o código novo)             ║
 * ╚══════════════════════════════════════════════════════════════════╝
 *
 * AÇÕES (todas via GET, ?action=...):
 *   (sem action) / action=recebidas        → comissões com Recebido=TRUE (import + relatório semanal)
 *   action=months                          → nomes das abas (meses) que têm comissões
 *   action=list&aba=NOME                    → todas as linhas da aba, com etapas + nº da linha
 *   action=update&aba=NOME&row=N&field=F&value=V → marca/desmarca etapa ou grava data
 *   action=write&aba=NOME&nome=...&...      → adiciona uma nova linha de comissão
 */

// Deixe vazio para usar a planilha onde o script está vinculado.
// Ou cole o ID da planilha "Fechamentos - Davi" entre as aspas.
var SPREADSHEET_ID = '1gxXfryV6Mji9xIjGgktb0C_z35Rh_yc8Ia08wTeRQwM';

function getSS() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

// Chave canônica do app  ->  rótulos possíveis do cabeçalho na planilha
var FIELD_LABELS = {
  nome:            ['Nome do cliente'],
  numero:          ['Número', 'Numero'],
  data_reuniao:    ['Data da última reunião', 'Data da ultima reuniao'],
  data_fechamento: ['Data de fechamento/resposta', 'Data de fechamento'],
  closer:          ['Closer'],
  percentual:      ['Porcentagem comissão', 'Porcentagem comissao', '% comissão'],
  valor:           ['Valor'],
  comissao:        ['Comissão', 'Comissao'],
  fechado:         ['Fechado'],
  assinado:        ['Contrato', 'Assinado'],
  pago:            ['Pago'],
  recebido:        ['Recebido'],
  data_recebido:   ['Data - Recebido', 'Data-Recebido', 'Data Recebido', 'Data recebido'],
  briefing:        ['Briefing'],
  obs:             ['Obs', 'Observação', 'Observacao']
};

function norm(s) { return String(s == null ? '' : s).trim().toLowerCase(); }

// Localiza a linha de cabeçalho (que contém "Nome do cliente") e mapeia colunas (1-based)
function findHeader(sheet) {
  var maxR = Math.min(8, sheet.getLastRow() || 1);
  var maxC = sheet.getLastColumn() || 1;
  if (maxR < 1 || maxC < 1) return null;
  var grid = sheet.getRange(1, 1, maxR, maxC).getValues();
  for (var r = 0; r < grid.length; r++) {
    var row = grid[r];
    for (var c = 0; c < row.length; c++) {
      if (norm(row[c]) === 'nome do cliente') {
        var cols = {};
        for (var key in FIELD_LABELS) {
          var labels = FIELD_LABELS[key].map(norm);
          for (var cc = 0; cc < row.length; cc++) {
            if (labels.indexOf(norm(row[cc])) >= 0) { cols[key] = cc + 1; break; }
          }
        }
        return { headerRow: r + 1, cols: cols };
      }
    }
  }
  return null;
}

function isCommissionSheet(sheet) {
  var h = findHeader(sheet);
  return !!(h && h.cols.nome);
}

function toBool(v) {
  if (v === true) return true;
  if (v === false || v == null) return false;
  var s = norm(v);
  return s === 'true' || s === 'verdadeiro' || s === 'sim' || s === '1' || s === 'x' || s === '✓';
}

// Valor de exibição (mantém o estilo da planilha: "30/04", "15-05", "-")
function fmtDisplay(v) {
  if (v instanceof Date) {
    return ('0' + v.getDate()).slice(-2) + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/' + v.getFullYear();
  }
  return String(v == null ? '' : v);
}

// Ano "de apoio" para datas sem ano, extraído do nome da aba (ex: "Junho 2026", "Jan/26")
function abaYear(name) {
  var m = String(name || '').match(/(\d{4})/);
  if (m) return parseInt(m[1], 10);
  m = String(name || '').match(/(\d{2})\b/);
  if (m) return 2000 + parseInt(m[1], 10);
  return new Date().getFullYear();
}

// Converte para ISO (yyyy-mm-dd); usa o ano da aba quando a data não tiver ano
function toISO(v, fallbackYear) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  if (!s || s === '-') return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/);
  if (m) {
    var d = ('0' + m[1]).slice(-2), mo = ('0' + m[2]).slice(-2), y = m[3];
    if (!y) y = fallbackYear || new Date().getFullYear();
    if (String(y).length === 2) y = '20' + y;
    return y + '-' + mo + '-' + d;
  }
  return '';
}

function parseNum(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/R\$/g, '').replace(/\s/g, '').replace(/\./g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function parsePct(v) {
  if (typeof v === 'number') return v <= 1 ? v * 100 : v;
  var s = String(v == null ? '' : v).replace('%', '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function listMonths() {
  var out = [];
  getSS().getSheets().forEach(function (sh) {
    if (isCommissionSheet(sh)) out.push(sh.getName());
  });
  return out;
}

function listRows(abaName) {
  var sh = getSS().getSheetByName(abaName);
  if (!sh) throw new Error('Aba não encontrada: ' + abaName);
  var hdr = findHeader(sh);
  if (!hdr) throw new Error('Cabeçalho não encontrado na aba: ' + abaName);
  var cols = hdr.cols;
  var first = hdr.headerRow + 1;
  var last = sh.getLastRow();
  var items = [];
  if (last >= first) {
    var data = sh.getRange(first, 1, last - first + 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var g = function (key) { return cols[key] ? row[cols[key] - 1] : ''; };
      var nome = String(g('nome') || '').trim();
      if (!nome) continue;
      items.push({
        row: first + i,
        nome: nome,
        data_reuniao: fmtDisplay(g('data_reuniao')),
        data_fechamento: fmtDisplay(g('data_fechamento')),
        closer: String(g('closer') || ''),
        percentual: parsePct(g('percentual')),
        valor: parseNum(g('valor')),
        comissao: parseNum(g('comissao')),
        fechado: toBool(g('fechado')),
        assinado: toBool(g('assinado')),
        pago: toBool(g('pago')),
        recebido: toBool(g('recebido')),
        data_recebido: fmtDisplay(g('data_recebido')),
        obs: String(g('obs') || '')
      });
    }
  }
  return { aba: abaName, comissoes: items };
}

function updateCell(abaName, rowNum, field, value) {
  var sh = getSS().getSheetByName(abaName);
  if (!sh) throw new Error('Aba não encontrada: ' + abaName);
  var hdr = findHeader(sh);
  if (!hdr || !hdr.cols[field]) throw new Error('Coluna não encontrada para o campo: ' + field);
  var cell = sh.getRange(Number(rowNum), hdr.cols[field]);
  if (['fechado', 'assinado', 'pago', 'recebido'].indexOf(field) >= 0) {
    cell.setValue(toBool(value));
  } else if (field === 'data_recebido') {
    cell.setValue(value ? value : '-');
  } else {
    cell.setValue(value);
  }
  return { updated: true, aba: abaName, row: Number(rowNum), field: field, value: value };
}

function writeRow(p) {
  var sh = getSS().getSheetByName(p.aba);
  if (!sh) throw new Error('Aba não encontrada: ' + p.aba);
  var hdr = findHeader(sh);
  if (!hdr) throw new Error('Cabeçalho não encontrado: ' + p.aba);
  var cols = hdr.cols;
  var first = hdr.headerRow + 1;
  var last = sh.getLastRow();
  var target = last + 1;
  // reaproveita a primeira linha de dados "vazia" (sem nome), se houver
  if (last >= first && cols.nome) {
    var names = sh.getRange(first, cols.nome, last - first + 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (!String(names[i][0]).trim()) { target = first + i; break; }
    }
  }
  var set = function (key, val) { if (cols[key]) sh.getRange(target, cols[key]).setValue(val); };
  set('nome', p.nome || '');
  set('data_fechamento', p.data_fechamento || '');
  if (p.percentual) set('percentual', parseFloat(p.percentual) / 100);
  if (p.valor) set('valor', parseFloat(p.valor));
  if (cols.comissao) {
    var cc = sh.getRange(target, cols.comissao);
    if (!cc.getFormula() && p.comissao) cc.setValue(parseFloat(p.comissao));
  }
  set('fechado', toBool(p.fechado));
  set('assinado', toBool(p.assinado));
  set('pago', toBool(p.pago));
  set('recebido', toBool(p.recebido));
  set('data_recebido', p.data_recebido ? p.data_recebido : '-');
  set('obs', p.obs || '');
  return { written: true, aba: p.aba, row: target };
}

// Compatível com o webhook antigo: comissões já recebidas (todas as abas)
function recebidas() {
  var out = [];
  listMonths().forEach(function (aba) {
    try {
      var fy = abaYear(aba);
      listRows(aba).comissoes.forEach(function (c) {
        if (c.recebido) {
          var dr = toISO(c.data_recebido, fy);
          out.push({
            nome: c.nome, valor: c.valor, percentual: c.percentual, comissao: c.comissao,
            data_fechamento: toISO(c.data_fechamento, fy),
            data_recebido: dr, data: dr, obs: c.obs, aba: aba
          });
        }
      });
    } catch (e) {}
  });
  return out;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || '';
  try {
    if (action === 'months') return jsonOut({ success: true, meses: listMonths() });
    if (action === 'list')   return jsonOut(Object.assign({ success: true }, listRows(p.aba)));
    if (action === 'update') return jsonOut(Object.assign({ success: true }, updateCell(p.aba, p.row, p.field, p.value)));
    if (action === 'write')  return jsonOut(Object.assign({ success: true }, writeRow(p)));
    var rec = recebidas();
    return jsonOut({ success: true, tipo: 'recebidas', comissoes: rec, total: rec.length, timestamp: new Date().toISOString() });
  } catch (err) {
    return jsonOut({ success: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }
