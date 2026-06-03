/**
 * DIÁRIO DE OBRA — Adições ao Google Apps Script
 * ================================================
 * Copie as funções abaixo para o seu Google Apps Script existente.
 *
 * PASSO 1 — PLANILHA
 * Crie uma nova aba chamada exatamente "Avisos" com os seguintes cabeçalhos
 * na linha 1 (uma coluna por célula, na ordem abaixo):
 *
 *   A        B            C       D          E        F          G     H      I
 *   ID | DESTINATARIO | TIPO | MENSAGEM | OBRA_ID | OBRA_NOME | DATA | LIDO | REMETENTE
 *
 * PASSO 2 — doGet
 * Adicione os novos casos dentro do switch/if que já trata as actions:
 *
 *   case 'loadAvisos':     return jsonpResponse(loadAvisos(params.destinatario), params.callback);
 *   case 'marcarLido':     return jsonpResponse(marcarLido(params.id), params.callback);
 *   case 'marcarTodosLidos': return jsonpResponse(marcarTodosLidos(params.destinatario), params.callback);
 *   case 'criarAviso':     return jsonpResponse(criarAviso(params), params.callback);
 *
 * PASSO 3 — saveOneObra
 * Ao final da função que salva uma obra, adicione a chamada:
 *   gerarAvisosAutomaticos(obraAntiga, novaObra);
 * (veja exemplo no final deste arquivo)
 */

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function getAvisosSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Avisos');
  if (!sheet) {
    sheet = ss.insertSheet('Avisos');
    sheet.appendRow(['ID','DESTINATARIO','TIPO','MENSAGEM','OBRA_ID','OBRA_NOME','DATA','LIDO','REMETENTE']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToAviso(row) {
  return {
    id:          String(row[0] || ''),
    destinatario:String(row[1] || ''),
    tipo:        String(row[2] || ''),
    mensagem:    String(row[3] || ''),
    obraId:      String(row[4] || ''),
    obraNome:    String(row[5] || ''),
    data:        String(row[6] || ''),
    lido:        row[7] === true || String(row[7]).toUpperCase() === 'TRUE',
    remetente:   String(row[8] || ''),
  };
}

// ─── loadAvisos ───────────────────────────────────────────────────────────────
/**
 * destinatario = '' → retorna todos (para ADM monitorar)
 * destinatario = 'Nome do Técnico' → retorna apenas os do técnico
 */
function loadAvisos(destinatario) {
  const sheet = getAvisosSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { avisos: [] };

  const dest = String(destinatario || '').trim();
  const avisos = [];

  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    const rowDest = String(row[1] || '').trim();
    if (dest === '' || rowDest === dest) {
      avisos.push(rowToAviso(row));
    }
  }

  // Não lidos primeiro, depois por data decrescente
  avisos.sort((a, b) => {
    if (a.lido !== b.lido) return a.lido ? 1 : -1;
    return b.data.localeCompare(a.data);
  });

  return { avisos };
}

// ─── marcarLido ───────────────────────────────────────────────────────────────
function marcarLido(id) {
  const sheet = getAvisosSheet();
  const data  = sheet.getDataRange().getValues();
  const strId = String(id || '');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === strId) {
      sheet.getRange(i + 1, 8).setValue(true); // coluna H = LIDO
      return { success: true };
    }
  }
  return { success: false, erro: 'Aviso não encontrado' };
}

// ─── marcarTodosLidos ─────────────────────────────────────────────────────────
function marcarTodosLidos(destinatario) {
  const sheet = getAvisosSheet();
  const data  = sheet.getDataRange().getValues();
  const dest  = String(destinatario || '').trim();

  for (let i = 1; i < data.length; i++) {
    const rowDest = String(data[i][1] || '').trim();
    const jaLido  = data[i][7] === true || String(data[i][7]).toUpperCase() === 'TRUE';
    if (rowDest === dest && !jaLido) {
      sheet.getRange(i + 1, 8).setValue(true);
    }
  }
  return { success: true };
}

// ─── criarAviso ───────────────────────────────────────────────────────────────
/**
 * params = { destinatario, tipo, mensagem, obraId, obraNome, remetente }
 * tipos válidos: nova_obra | nova_pendencia | mensagem_adm
 */
function criarAviso(params) {
  const sheet = getAvisosSheet();
  const id    = String(Date.now()) + String(Math.floor(Math.random() * 9999));
  const data  = new Date().toISOString();

  sheet.appendRow([
    id,
    String(params.destinatario || ''),
    String(params.tipo         || ''),
    String(params.mensagem     || ''),
    String(params.obraId       || ''),
    String(params.obraNome     || ''),
    data,
    false,
    String(params.remetente    || 'Sistema'),
  ]);

  return { success: true, id };
}

// ─── gerarAvisosAutomaticos ───────────────────────────────────────────────────
/**
 * Chame esta função no final da sua saveOneObra, passando a obra
 * como estava ANTES de salvar (obraAntiga) e a nova versão (novaObra).
 *
 * Se for uma obra nova (não existia antes), passe obraAntiga = null.
 *
 * Exemplo de uso dentro de saveOneObra:
 *
 *   function saveOneObra(params) {
 *     const novaObra  = JSON.parse(params.obra);
 *     const obraAntiga = encontrarObraExistente(novaObra.id); // sua lógica atual
 *
 *     // ... seu código existente de salvar na planilha ...
 *
 *     gerarAvisosAutomaticos(obraAntiga, novaObra);   // ← adicione esta linha
 *     return { success: true, versao: novaObra.versao + 1 };
 *   }
 */
function gerarAvisosAutomaticos(obraAntiga, novaObra) {
  if (!novaObra || !novaObra.responsavel) return;

  const novoResp = String(novaObra.responsavel).trim();
  const nome     = String(novaObra.nome || '').trim();
  const obraId   = String(novaObra.id   || '');

  // 1. Obra nova atribuída a um técnico
  if (!obraAntiga && novoResp) {
    criarAviso({
      destinatario: novoResp,
      tipo:         'nova_obra',
      mensagem:     `Você foi atribuído à obra ${nome}`,
      obraId,
      obraNome:     nome,
      remetente:    'Sistema',
    });
    return; // nova obra não tem pendências anteriores para comparar
  }

  // 2. Responsável trocado
  if (obraAntiga && String(obraAntiga.responsavel || '').trim() !== novoResp) {
    criarAviso({
      destinatario: novoResp,
      tipo:         'nova_obra',
      mensagem:     `Você foi atribuído à obra ${nome}`,
      obraId,
      obraNome:     nome,
      remetente:    'Sistema',
    });
  }

  // 3. Novas pendências adicionadas
  if (obraAntiga) {
    const idsAntigos = new Set(
      (obraAntiga.pendencias || []).map(p => String(p.id))
    );
    const novasPends = (novaObra.pendencias || []).filter(
      p => !idsAntigos.has(String(p.id))
    );
    novasPends.forEach(p => {
      criarAviso({
        destinatario: novoResp,
        tipo:         'nova_pendencia',
        mensagem:     `Nova pendência em ${nome}: ${p.texto}`,
        obraId,
        obraNome:     nome,
        remetente:    'Sistema',
      });
    });
  }
}
