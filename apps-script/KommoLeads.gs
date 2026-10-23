/* =========================================================================
   BM · KOMMO — LEADS E PRÉ-QUALIFICAÇÃO

   Alimenta a planilha "Dashboards Kommo" com UMA LINHA POR LEAD: quando
   chegou, de onde veio, em que etapa está, o que respondeu na triagem e se
   virou reunião. É a base do painel de pré-qualificação.

   Mora no mesmo projeto do Meta Ads só por causa do token: KOMMO_TOKEN já
   está nas Propriedades do script e um token só é menos coisa para vencer.
   Fora isso é independente — outra planilha, outra aba, outro gatilho.

   O funil "Processual" fica de fora: lá são clientes, não leads.

   Reunião marcada = entrada na etapa "Enviado ao ADVBOX", a mesma regra do
   painel de campanhas (ver Kommo.gs).

   INSTALAÇÃO
   1. configurarKommoLeads()  — cria a aba e as chaves em _Config.
   2. sincronizarKommoLeadsCompleta()  — carga do histórico (retomável).
   3. criarGatilhoKommoLeads()  — mantém em dia de hora em hora.
   ========================================================================= */

var KL_ABA = 'Leads';
var KL_CAB = ['ID', 'Criado em', 'Hora', 'Funil', 'Etapa', 'Situação', 'Origem', 'Campanha',
              'Produto jurídico', 'Plano de saúde', 'Tipo do plano', 'Mensalidade',
              'Tempo de plano', 'Responsável', 'Reunião em', 'Atualizado em'];

var KL_PADRAO = {
  KOMMO_LEADS_PLANILHA: '1Pj8w7YofhBR2PixGFUET9AhG713ZUcbig3lu3rlnlQI',
  KOMMO_LEADS_INICIO: '2026-01-01',
  KOMMO_FUNIS_IGNORADOS: 'Processual',
  KOMMO_LEADS_DIAS: '7'
};

var KL_DESCRICAO = {
  KOMMO_LEADS_PLANILHA: 'ID da planilha "Dashboards Kommo", onde a aba Leads é escrita.',
  KOMMO_LEADS_INICIO: 'Primeiro dia de leads a carregar (AAAA-MM-DD).',
  KOMMO_FUNIS_IGNORADOS: 'Funis que não são leads, separados por vírgula (Processual = clientes).',
  KOMMO_LEADS_DIAS: 'Quantos dias para trás a sincronização de hora em hora reconfere, por data de atualização.'
};

/* O Apps Script corta em 6 minutos. Aos 4 o trabalho para, grava onde parou e
   um gatilho de 1 minuto retoma — igual ao histórico do Meta. */
var KL_LIMITE_MS = 4 * 60 * 1000;

/* ============================================================ COMANDOS */

function configurarKommoLeads() {
  garantirAba_(planilhaKommoLeads_(), KL_ABA, KL_CAB);
  var aba = garantirAba_(planilha_(), ABA.config, CAB.config);
  var atuais = {};
  var ult = aba.getLastRow();
  if (ult > 1) aba.getRange(2, 1, ult - 1, 2).getValues().forEach(function (l) { if (l[0]) atuais[String(l[0])] = true; });
  var novas = [];
  Object.keys(KL_PADRAO).forEach(function (k) {
    if (!atuais[k]) novas.push([k, KL_PADRAO[k], KL_DESCRICAO[k] || '']);
  });
  if (novas.length) aba.getRange(aba.getLastRow() + 1, 1, novas.length, 3).setValues(novas);
  alertar_('Aba "' + KL_ABA + '" criada na planilha Dashboards Kommo e chaves gravadas em _Config.\n\n' +
           'Agora: sincronizarKommoLeadsCompleta e depois criarGatilhoKommoLeads.');
}

function criarGatilhoKommoLeads() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sincronizarKommoLeads') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizarKommoLeads').timeBased().everyHours(1).create();
  alertar_('Gatilho de hora em hora criado para sincronizarKommoLeads.');
}

/** Histórico inteiro desde KOMMO_LEADS_INICIO. Retomável. */
function sincronizarKommoLeadsCompleta() {
  props_().setProperty('kl_cursor', JSON.stringify({ pagina: 1, gravados: 0, inicio: agoraTxt_() }));
  var aba = garantirAba_(planilhaKommoLeads_(), KL_ABA, KL_CAB);
  var ult = aba.getLastRow();
  if (ult > 1) aba.getRange(2, 1, ult - 1, KL_CAB.length).clearContent();
  continuarKommoLeads();
}

/** Retomada da carga completa; também é o gatilho de 1 minuto. */
function continuarKommoLeads() {
  var cursor = JSON.parse(props_().getProperty('kl_cursor') || 'null');
  if (!cursor) return;
  var t0 = new Date().getTime();
  var de = Math.floor((dataDeIso_(cfgKommo2_('KOMMO_LEADS_INICIO')) || new Date(2026, 0, 1)).getTime() / 1000);
  var reunioes = mapaReunioes_();
  var funis = funisValidos_();
  var usuarios = usuariosKommo_();
  var aba = garantirAba_(planilhaKommoLeads_(), KL_ABA, KL_CAB);

  while (new Date().getTime() - t0 < KL_LIMITE_MS) {
    var j = kommoGet_('/leads', { limit: 250, page: cursor.pagina, 'filter[is_deleted]': 0,
                                  'filter[created_at][from]': de });
    var ls = (j && j._embedded && j._embedded.leads) || [];
    var linhas = linhasDeLeads_(ls, funis, usuarios, reunioes);
    if (linhas.length) aba.getRange(aba.getLastRow() + 1, 1, linhas.length, KL_CAB.length).setValues(linhas);
    cursor.gravados += linhas.length;
    if (ls.length < 250) {
      props_().deleteProperty('kl_cursor');
      limparGatilhosKommoLeads_();
      gravarSync_({ kommoLeadsOk: agoraTxt_(), kommoLeadsErro: '',
                    kommoLeadsMensagem: cursor.gravados + ' leads gravados desde ' + cfgKommo2_('KOMMO_LEADS_INICIO') });
      alertar_(cursor.gravados + ' leads gravados na aba "' + KL_ABA + '".');
      return;
    }
    cursor.pagina++;
    props_().setProperty('kl_cursor', JSON.stringify(cursor));
  }
  /* Acabou o tempo desta execução: agenda a próxima e sai. */
  agendarKommoLeads_();
  gravarSync_({ kommoLeadsMensagem: 'Carga em andamento: ' + cursor.gravados + ' leads, página ' + cursor.pagina + '…' });
}

/** De hora em hora: reconfere quem foi criado OU mexido nos últimos N dias.
    Pelo updated_at, e não pelo created_at, porque o que muda num lead antigo
    é justamente a etapa — que é o dado do painel. */
function sincronizarKommoLeads() {
  var dias = Number(cfgKommo2_('KOMMO_LEADS_DIAS')) || 7;
  var desde = Math.floor((hoje_().getTime() - (dias - 1) * 86400000) / 1000);
  var inicio = dataDeIso_(cfgKommo2_('KOMMO_LEADS_INICIO')) || new Date(2026, 0, 1);
  var reunioes = mapaReunioes_();
  var funis = funisValidos_();
  var usuarios = usuariosKommo_();

  var novas = [];
  for (var pagina = 1; pagina <= 60; pagina++) {
    var j = kommoGet_('/leads', { limit: 250, page: pagina, 'filter[is_deleted]': 0,
                                  'filter[updated_at][from]': desde });
    var ls = (j && j._embedded && j._embedded.leads) || [];
    novas = novas.concat(linhasDeLeads_(ls, funis, usuarios, reunioes, inicio));
    if (ls.length < 250) break;
  }
  var total = mesclarLeads_(novas);
  gravarSync_({ kommoLeadsOk: agoraTxt_(), kommoLeadsErro: '',
                kommoLeadsMensagem: novas.length + ' leads reconferidos · ' + total + ' na aba' });
  return total;
}

function verificarKommoLeads() {
  var funis = funisValidos_();
  var nomes = [];
  for (var id in funis) nomes.push(funis[id].nome);
  var j = kommoGet_('/leads', { limit: 1, 'filter[is_deleted]': 0 });
  alertar_('Kommo respondendo.\n\nFunis considerados (Processual fica de fora):\n  ' + nomes.join('\n  ') +
           '\n\nPlanilha de destino: ' + cfgKommo2_('KOMMO_LEADS_PLANILHA') +
           '\nLeads a partir de ' + cfgKommo2_('KOMMO_LEADS_INICIO') + '.');
}

/* ============================================================== MONTAGEM */

/** Converte leads da API em linhas da aba. `inicio` corta o que for antigo
    demais (a incremental busca por updated_at e traz lead de qualquer data). */
function linhasDeLeads_(leads, funis, usuarios, reunioes, inicio) {
  var tz = Session.getScriptTimeZone();
  var agora = agoraTxt_();
  var out = [];
  for (var i = 0; i < leads.length; i++) {
    var l = leads[i];
    var funil = funis[l.pipeline_id];
    if (!funil) continue;                        // funil ignorado (Processual)
    var criado = new Date(l.created_at * 1000);
    if (inicio && criado < inicio) continue;
    var etapa = funil.etapas[l.status_id] || '';
    out.push([
      String(l.id),
      Utilities.formatDate(criado, tz, 'yyyy-MM-dd'),
      Utilities.formatDate(criado, tz, 'HH:mm'),
      funil.nome,
      etapa,
      situacaoEtapa_(etapa),
      campoLead_({ campos: l.custom_fields_values || [] }, 'Origem'),
      campoLead_({ campos: l.custom_fields_values || [] }, 'Campanha'),
      campoLead_({ campos: l.custom_fields_values || [] }, 'Produto Jurídico'),
      campoLead_({ campos: l.custom_fields_values || [] }, 'Qual seu plano de saúde?'),
      campoLead_({ campos: l.custom_fields_values || [] }, 'Qual o tipo do seu plano?'),
      campoLead_({ campos: l.custom_fields_values || [] }, 'Qual o valor da mensalidade?'),
      campoLead_({ campos: l.custom_fields_values || [] }, 'Quanto tempo você está nele?'),
      usuarios[l.responsible_user_id] || '',
      reunioes[String(l.id)] || '',
      agora
    ]);
  }
  return out;
}

/* A etapa é o estado do lead; o nome dela diz em que pé está. Classificar
   aqui (e não no painel) deixa o número conferível na própria planilha. */
function situacaoEtapa_(etapa) {
  var e = normKommo_(etapa);
  if (!e) return '';
  if (/enviado ao advbox|pesquisa de satisfacao|closed - won|venda ganha/.test(e)) return 'Reunião marcada';
  if (/sem resposta|lead perdido|inelegivel|desqualificado|inviavel|nao responderam|closed - lost|venda perdida/.test(e)) return 'Perdido';
  if (/entrada de leads|etapa de leads de entrada|aguardando forms|1º contato|1o contato|comentarios|direct|ja esta no wpp/.test(e)) return 'Entrada';
  return 'Em qualificação';
}

/** Funis que contam como lead, com o nome de cada etapa. */
function funisValidos_() {
  var fora = String(cfgKommo2_('KOMMO_FUNIS_IGNORADOS') || '').split(',')
    .map(function (s) { return normKommo_(s); }).filter(Boolean);
  var j = kommoGet_('/leads/pipelines', { limit: 250 });
  var ps = (j && j._embedded && j._embedded.pipelines) || [];
  var out = {};
  ps.forEach(function (p) {
    if (fora.indexOf(normKommo_(p.name)) >= 0) return;
    var etapas = {};
    ((p._embedded && p._embedded.statuses) || []).forEach(function (s) { etapas[s.id] = s.name; });
    out[p.id] = { nome: p.name, etapas: etapas };
  });
  return out;
}

function usuariosKommo_() {
  var j = kommoGet_('/users', { limit: 250 });
  var out = {};
  (((j || {})._embedded || {}).users || []).forEach(function (u) { out[u.id] = u.name; });
  return out;
}

/** leadId -> data da PRIMEIRA entrada em "Enviado ao ADVBOX". */
function mapaReunioes_() {
  var ini = dataDeIso_(cfgKommo2_('KOMMO_LEADS_INICIO')) || new Date(2026, 0, 1);
  var primeiro = primeirasEntradasAdvbox_(ini, new Date());
  var tz = Session.getScriptTimeZone();
  var out = {};
  for (var id in primeiro) out[id] = Utilities.formatDate(new Date(primeiro[id].quando * 1000), tz, 'yyyy-MM-dd');
  return out;
}

/* ============================================================== GRAVAÇÃO */

/** Substitui as linhas dos leads reconferidos e mantém o resto. */
function mesclarLeads_(novas) {
  var aba = garantirAba_(planilhaKommoLeads_(), KL_ABA, KL_CAB);
  var ult = aba.getLastRow();
  var antigas = ult > 1 ? aba.getRange(2, 1, ult - 1, KL_CAB.length).getValues() : [];
  var porId = {};
  antigas.forEach(function (l) { if (l[0]) porId[String(l[0])] = l; });
  novas.forEach(function (l) { porId[String(l[0])] = l; });
  var todas = [];
  for (var k in porId) todas.push(porId[k]);
  todas.sort(function (a, b) {
    var da = dataTxt_(a[1]), db = dataTxt_(b[1]);
    return da < db ? -1 : da > db ? 1 : String(a[2]).localeCompare(String(b[2]));
  });
  reescrever_(aba, KL_CAB, todas);
  return todas.length;
}

function planilhaKommoLeads_() {
  return SpreadsheetApp.openById(cfgKommo2_('KOMMO_LEADS_PLANILHA'));
}

function agendarKommoLeads_() {
  limparGatilhosKommoLeads_();
  ScriptApp.newTrigger('continuarKommoLeads').timeBased().after(60 * 1000).create();
}

function limparGatilhosKommoLeads_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'continuarKommoLeads') ScriptApp.deleteTrigger(t);
  });
}

/* cfg_ cai no PADRAO_CONFIG do Meta, que não conhece estas chaves; aqui a
   reserva é a tabela deste arquivo. */
function cfgKommo2_(chave) {
  var v = cfg_(chave);
  return v === undefined || v === null || String(v) === '' ? KL_PADRAO[chave] : cfgTexto_(v);
}
