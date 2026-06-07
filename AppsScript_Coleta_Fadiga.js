/**
 * COLETA FADIGA FAB — Google Apps Script
 * ────────────────────────────────────────────────────────────
 * Formato: LARGO (wide format) — 1 linha por participante,
 * colunas para os 3 momentos lado a lado.
 * Pronto para importação direta no IBM SPSS.
 *
 * COMO PUBLICAR:
 *  1. Cole este código no editor do Apps Script
 *  2. Salvar → Implantar → Nova implantação
 *  3. Tipo: Web app
 *  4. Executar como: Sua conta
 *  5. Quem tem acesso: Qualquer pessoa
 *  6. Copie a URL e cole em index.html na constante SHEETS_URL
 * ────────────────────────────────────────────────────────────
 */

// Colunas da planilha — ordem fixa para SPSS
// Cada bloco de colunas corresponde a um momento
const COLUNAS = [
  // ── Identificação ──────────────────────────────────────────
  'codigo_participante',

  // ── MOMENTO 1: Dia Sem Voo (Baseline) ──────────────────────
  'm1_timestamp_inicio',
  'm1_timestamp_fim',
  // Caracterização
  'm1_idade',
  'm1_genero',
  'm1_peso',
  'm1_altura',
  'm1_posto',
  'm1_qualificacao',
  'm1_horas_total',
  'm1_horas_c99',
  // IPAQ — dias e tempos (HH:MM)
  'm1_ipaq_caminhada_dias',
  'm1_ipaq_caminhada_tempo',
  'm1_ipaq_moderada_dias',
  'm1_ipaq_moderada_tempo',
  'm1_ipaq_vigorosa_dias',
  'm1_ipaq_vigorosa_tempo',
  'm1_ipaq_sentado_semana',
  'm1_ipaq_sentado_fds',
  // Escalas
  'm1_kss',
  'm1_sps',

  // ── MOMENTO 2: Antes do Pré-Voo ────────────────────────────
  'm2_timestamp_inicio',
  'm2_timestamp_fim',
  // Bloco A — Dados da Missão
  'm2_funcao_bordo',
  'm2_funcao_operacional',
  'm2_horario_decolagem',
  'm2_tempo_voo_previsto',
  'm2_tempo_voo_total_missao',
  'm2_autoridades',
  'm2_tipo_autoridade',
  'm2_comitiva',
  'm2_tipo_comitiva',
  // Bloco B — Estado no Momento da Coleta
  'm2_sono_horas',
  'm2_sono_qualidade',
  'm2_estresse',
  'm2_cafeina',
  'm2_cafeina_detalhe',
  'm2_deslocamento',
  // Escalas
  'm2_kss',
  'm2_sps',

  // ── MOMENTO 3: Após o Pré-Voo ──────────────────────────────
  'm3_timestamp_inicio',
  'm3_timestamp_fim',
  // NASA-TLX (escala 0–20)
  'm3_tlx_mental',
  'm3_tlx_fisica',
  'm3_tlx_temporal',
  'm3_tlx_desempenho_raw',
  'm3_tlx_desempenho',
  'm3_tlx_esforco',
  'm3_tlx_frustracao',
  // Escalas
  'm3_kss',
  'm3_sps',
];

function doPost(e) {
  try {
    console.log('🔵 doPost chamado');
    console.log('postData:', e.postData?.contents?.substring(0, 100));

    const data = JSON.parse(e.postData.contents);
    const mId = (data.momento || '').toLowerCase(); // 'm1', 'm2' ou 'm3'
    const participante = data.codigo_participante;

    console.log('Participante:', participante);
    console.log('Momento:', mId);

    if (!participante || !['m1','m2','m3'].includes(mId)) {
      console.log('❌ Dados inválidos');
      return resposta('erro', 'Dados inválidos: participante ou momento ausente.');
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Dados') || ss.insertSheet('Dados');

    // Garantir cabeçalho
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUNAS);
      sheet.getRange(1, 1, 1, COLUNAS.length)
           .setFontWeight('bold')
           .setBackground('#1a237e')
           .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // Localizar linha do participante (coluna A)
    const ultimaLinha = sheet.getLastRow();
    let linhaParticipante = -1;
    if (ultimaLinha > 1) {
      const colA = sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues();
      for (let i = 0; i < colA.length; i++) {
        if (colA[i][0] === participante) {
          linhaParticipante = i + 2; // +2: base 1 + cabeçalho
          break;
        }
      }
    }

    // Se não existe ainda, criar linha com participante
    if (linhaParticipante === -1) {
      sheet.appendRow([participante]);
      linhaParticipante = sheet.getLastRow();
    }

    // Mapear campos do payload para as colunas certas
    const prefixo = mId + '_';
    COLUNAS.forEach((col, idx) => {
      let valor = '';
      if (col === 'codigo_participante') return; // já preenchido

      // timestamps ficam sob prefixo do momento
      if (col === `${mId}_timestamp_inicio`) {
        valor = data.timestamp_inicio || '';
      } else if (col === `${mId}_timestamp_fim`) {
        valor = data.timestamp_fim || '';
      } else if (col.startsWith(prefixo)) {
        const campo = col;
        valor = (data[campo] !== undefined && data[campo] !== null) ? data[campo] : '';
      } else {
        return; // campo de outro momento: não sobrescrever
      }

      sheet.getRange(linhaParticipante, idx + 1).setValue(valor);
    });

    return resposta('ok', `Momento ${mId} de ${participante} salvo.`);

  } catch(err) {
    return resposta('erro', err.message);
  }
}

function resposta(status, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status, msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Função utilitária — rode manualmente para limpar a planilha
 * e recriar o cabeçalho (útil em testes).
 */
function resetPlanilha() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dados');
  if (sheet) sheet.clearContents();
}
