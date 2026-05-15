/**
 * Apps Script - Exportar Comissões Recebidas (v2)
 * Retorna: nome, valor cheio, percentual, comissao, data_fechamento, data_recebido, obs
 *
 * PARA ATUALIZAR a implantação existente (URL continua a mesma):
 * 1. Cole este código no Apps Script
 * 2. Clique em Implantar → Gerenciar implantações
 * 3. Clique no lápis (Editar) da implantação ativa
 * 4. Em "Versão" selecione "Nova versão"
 * 5. Clique em Implantar — pronto, a URL não muda!
 */

const SHEET_GID = 1556243639;

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(s => s.getSheetId() === SHEET_GID);
    if (!sheet) return jsonResponse({ success: false, error: 'Aba não encontrada. Verifique o SHEET_GID.' });

    const data = sheet.getDataRange().getValues();

    // Localiza linha do cabeçalho
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i].map(c => String(c).toLowerCase()).some(c => c.includes('nome do cliente'))) {
        headerRow = i; break;
      }
    }
    if (headerRow < 0) return jsonResponse({ success: false, error: 'Cabeçalho não encontrado.' });

    const headers = data[headerRow].map(c => String(c).toLowerCase().trim());

    // Mapeia colunas
    const colNome         = headers.findIndex(h => h.includes('nome') && h.includes('cliente'));
    const colValor        = headers.findIndex(h => h === 'valor');
    const colPercentual   = headers.findIndex(h => h.includes('porcentagem') || h.includes('percentual'));
    const colComissao     = headers.findIndex(h => h === 'comissão' || h === 'comissao');
    const colDataFech     = headers.findIndex(h => h.includes('fechamento') || (h.includes('data') && h.includes('resposta')));
    const colPago         = headers.findIndex(h => h === 'pago');
    const colDataRecebido = headers.findIndex(h => h.includes('data') && h.includes('recebido'));
    const colObs          = headers.findIndex(h => h === 'obs' || h.includes('observa'));

    const comissoes = [];

    for (let i = headerRow + 1; i < data.length; i++) {
      const row  = data[i];
      const nome = String(row[colNome] || '').trim();
      if (!nome) continue;

      // Só inclui se checkbox "Pago" estiver marcada
      const pago = row[colPago] === true || String(row[colPago]).toUpperCase() === 'TRUE';
      if (!pago) continue;

      // Percentual: Sheets armazena 10% como 0.1 → converte para 10
      let pct = parseValor(colPercentual >= 0 ? row[colPercentual] : 0);
      if (pct > 0 && pct <= 1) pct = pct * 100;

      // Usa data de recebimento se preenchida, senão usa data de fechamento
      const dataRecebidoRaw = colDataRecebido >= 0 ? row[colDataRecebido] : null;
      const dataFechRaw     = colDataFech >= 0 ? row[colDataFech] : null;
      const dataRef         = (dataRecebidoRaw && String(dataRecebidoRaw).trim() !== '' && String(dataRecebidoRaw).trim() !== '-')
                              ? dataRecebidoRaw : dataFechRaw;

      comissoes.push({
        nome,
        valor:           parseValor(colValor    >= 0 ? row[colValor]    : 0),
        percentual:      pct,
        comissao:        parseValor(colComissao >= 0 ? row[colComissao] : 0),
        data_fechamento: formatDate(dataFechRaw),
        data_recebido:   formatDate(dataRef),
        obs:             String(colObs >= 0 ? row[colObs] : '').trim()
      });
    }

    return jsonResponse({
      success:    true,
      comissoes,
      total:      comissoes.length,
      timestamp:  new Date().toISOString()
    });

  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function parseValor(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  if (typeof raw === 'number') return raw;
  return parseFloat(String(raw).replace(/%/g, '').replace(/R\$\s*/g, '').replace(/\./g, '').replace(',', '.').trim()) || 0;
}

function formatDate(raw) {
  if (!raw || String(raw).trim() === '-' || String(raw).trim() === '') return '';
  if (raw instanceof Date) {
    const d = raw;
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  const parts = String(raw).trim().split('/');
  if (parts.length >= 2) {
    const dia = parts[0].padStart(2, '0');
    const mes = parts[1].padStart(2, '0');
    const ano = parts.length >= 3 ? parts[2] : new Date().getFullYear();
    return `${ano}-${mes}-${dia}`;
  }
  return String(raw).trim();
}

function jsonResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
