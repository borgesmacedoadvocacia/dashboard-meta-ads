/* =========================================================================
   CAMPANHAS META ADS → PLANILHA
   Borges Macedo Advocacia

   O que este script faz
   ---------------------
   Busca no Meta (Graph API / Marketing API) os resultados das campanhas das
   contas de anuncio configuradas e grava nesta planilha, em abas proprias:

     Insights diários  uma linha por campanha POR DIA (gasto, impressoes,
                       alcance, cliques, leads...). E daqui que o painel monta
                       qualquer recorte de periodo - mes atual, 30 dias, ano,
                       periodo escolhido - sem consultar o Meta a cada clique.
     Janelas           campanha x janela (7 dias, 30 dias, mes atual) com as
                       metricas que NAO somam entre dias: alcance, frequencia,
                       CPM, CTR, CPC. Somar alcance diario da alcance errado.
     Anúncios          anuncio x janela (7 e 30 dias): e onde um gestor de
                       trafego ve qual criativo esta ganhando e qual cansou.
     Campanhas         orcamento diario/vitalicio e status de cada campanha
                       (somando os conjuntos quando o orcamento e por conjunto).
     Mapeamento        ORIGEM do Fluxo de Clientes -> nome da campanha. O
                       painel casa sozinho pelo nome; esta aba serve para
                       corrigir a mao o que ele errar.
     _Config           parametros editaveis (contas, data de inicio, tipos
                       de acao que contam como lead...).
     _Sync             estado da ultima sincronizacao, lido pelo painel.

   Onde fica o token
   -----------------
   NUNCA no codigo. Em Configuracoes do projeto -> Propriedades do script,
   crie a propriedade META_ACCESS_TOKEN com o token do Usuario do Sistema
   (permissao ads_read, validade "nunca expira"). Este arquivo pode ser
   versionado sem expor nada.

   Primeira vez
   ------------
   1. Cole este arquivo em Codigo.gs e salve.
   2. Propriedades do script -> META_ACCESS_TOKEN = <token>.
   3. Execute configurarTudo() uma vez (autorize quando pedir).
   4. Recarregue a planilha: aparece o menu "BM · Meta Ads".
   5. Menu -> Testar conexão com o Meta. Depois -> Sincronizar completa.
   6. Implantar -> Nova implantacao -> App da Web (executar como Eu, acesso
      Qualquer pessoa). A URL vai no botao Atualizar do painel.
   ========================================================================= */

var API = 'https://graph.facebook.com/v21.0/';

var ABA = {
  insights:   'Insights diários',
  janelas:    'Janelas',
  anuncios:   'Anúncios',
  campanhas:  'Campanhas',
  mapeamento: 'Mapeamento',
  config:     '_Config',
  sync:       '_Sync'
};

/* Valores iniciais de _Config. Depois de criados, quem manda e a aba. */
var PADRAO_CONFIG = {
  META_AD_ACCOUNTS: '106339469723880,731383791781450',
  DATA_INICIO: '2025-01-01',
  DIAS_INCREMENTAL: '7',
  TOKEN_WEBAPP: 'bm-metaads',
  /* Tipos de acao do Meta que contam como lead, por familia. O painel mostra
     as tres separadas e soma as que LEAD_TOTAL indicar. Conferir uma vez
     contra a coluna "Resultados" do Gerenciador de Anuncios. */
  TIPOS_LEAD_FORMULARIO: 'lead',
  TIPOS_LEAD_CONVERSA:   'onsite_conversion.messaging_conversation_started_7d',
  TIPOS_LEAD_PIXEL:      'offsite_conversion.fb_pixel_lead',
  LEAD_TOTAL: 'formulario+conversa'
};

var CAB = {
  insights: ['Data', 'Conta', 'ID da campanha', 'Campanha', 'Gasto', 'Impressões', 'Alcance',
             'Cliques', 'Cliques no link', 'Visualizações da página', 'Leads (formulário)',
             'Conversas iniciadas', 'Leads (pixel)', 'Leads (total)', 'Ações (JSON)', 'Atualizado em'],
  janelas:  ['Janela', 'Conta', 'ID da campanha', 'Campanha', 'Início', 'Fim', 'Gasto', 'Impressões',
             'Alcance', 'Frequência', 'CPM', 'CTR (%)', 'CPC', 'Cliques', 'Cliques no link',
             'CTR do link (%)', 'Visualizações da página', 'Leads (formulário)', 'Conversas iniciadas',
             'Leads (pixel)', 'Leads (total)', 'Custo por lead', 'Atualizado em'],
  anuncios: ['Janela', 'Conta', 'ID da campanha', 'Campanha', 'ID do conjunto', 'Conjunto',
             'ID do anúncio', 'Anúncio', 'Gasto', 'Impressões', 'Alcance', 'Frequência', 'CPM',
             'CTR (%)', 'CPC', 'Cliques no link', 'CTR do link (%)', 'Visualizações da página',
             'Leads (formulário)', 'Conversas iniciadas', 'Leads (pixel)', 'Leads (total)',
             'Custo por lead', 'Atualizado em'],
  campanhas: ['Conta', 'ID da campanha', 'Campanha', 'Status', 'Status efetivo', 'Objetivo',
              'Orçamento diário', 'Orçamento vitalício', 'Origem do orçamento', 'Conjuntos ativos',
              'Início', 'Fim', 'Atualizado em'],
  mapeamento: ['Origem (Fluxo de Clientes)', 'Campanha (nome exato no Meta)', 'Observação'],
  config: ['Chave', 'Valor', 'Descrição'],
  sync: ['Chave', 'Valor']
};

var JANELAS = [
  { chave: '7d',  preset: 'last_7d' },
  { chave: '30d', preset: 'last_30d' },
  { chave: 'mes', preset: 'this_month' }
];

var CAMPOS_INSIGHT = 'campaign_id,campaign_name,spend,impressions,reach,frequency,clicks,' +
                     'inline_link_clicks,cpm,cpc,ctr,inline_link_click_ctr,actions,cost_per_action_type';

/* ============================================================ MENU / SETUP */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('BM · Meta Ads')
    .addItem('Sincronizar incremental (últimos dias)', 'sincronizarIncremental')
    .addItem('Sincronizar completa (histórico)', 'sincronizarCompleta')
    .addItem('Atualizar só campanhas e anúncios', 'atualizarCampanhasEAnuncios')
    .addSeparator()
    .addItem('Testar conexão com o Meta', 'testarConexao')
    .addItem('Configurar tudo (1ª vez)', 'configurarTudo')
    .addItem('Recriar gatilho de hora em hora', 'criarGatilhoHorario')
    .addItem('Cancelar sincronização', 'cancelarSincronizacao')
    .addToUi();
}

function configurarTudo() {
  var ss = planilha_();
  Object.keys(ABA).forEach(function (k) { garantirAba_(ss, ABA[k], CAB[k]); });
  garantirConfig_(ss);
  /* A aba padrao vazia que vem com toda planilha nova so atrapalha. */
  var vazia = ss.getSheetByName('Página1');
  if (vazia && vazia.getLastRow() <= 1 && ss.getSheets().length > 1) ss.deleteSheet(vazia);
  criarGatilhoHorario();
  gravarSync_({ status: 'ocioso', etapa: 'configurado',
                mensagem: 'Abas criadas. Coloque META_ACCESS_TOKEN nas Propriedades do script e rode a sincronização completa.' });
  alertar_('Pronto. Abas criadas e gatilho de hora em hora ativo.\n\n' +
    'Agora: Configurações do projeto → Propriedades do script → META_ACCESS_TOKEN.\n' +
    'Depois: testarConexao → sincronizarCompleta (pelo editor ou pelo menu da planilha).');
}

function criarGatilhoHorario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sincronizarIncremental') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizarIncremental').timeBased().everyHours(1).create();
}

function cancelarSincronizacao() {
  limparGatilhosContinuacao_();
  props_().deleteProperty('cursor');
  gravarSync_({ status: 'ocioso', etapa: 'cancelada', mensagem: 'Sincronização cancelada pelo usuário.' });
}

function testarConexao() {
  var linhas = [];
  contas_().forEach(function (conta) {
    try {
      var r = chamarMeta_('act_' + conta, { fields: 'name,currency,account_status,timezone_name' });
      linhas.push('✔ ' + conta + ' — ' + r.name + ' · ' + r.currency + ' · fuso ' + r.timezone_name +
                  ' · status ' + r.account_status);
    } catch (e) {
      linhas.push('✘ ' + conta + ' — ' + e.message);
    }
  });
  var texto = linhas.join('\n');
  gravarSync_({ etapa: 'teste', mensagem: texto.slice(0, 400) });
  alertar_(texto);
  return texto;
}

/* ======================================================= SINCRONIZAÇÕES */

/** Ultimos N dias (DIAS_INCREMENTAL). Roda de hora em hora pelo gatilho. */
function sincronizarIncremental() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var dias = Number(cfg_('DIAS_INCREMENTAL')) || 7;
    var fim = hoje_();
    var ini = new Date(fim.getTime() - (dias - 1) * 86400000);
    gravarSync_({ status: 'rodando', etapa: 'incremental', progresso: 0,
                  mensagem: 'Buscando ' + iso_(ini) + ' a ' + iso_(fim) + '…' });
    var total = 0;
    contas_().forEach(function (conta) {
      total += sincronizarPeriodo_(conta, ini, fim);
    });
    atualizarCampanhas_();
    atualizarJanelasEAnuncios_();
    gravarSync_({ status: 'ocioso', etapa: 'incremental', progresso: 100,
                  ultimaOk: agoraTxt_(), ultimoErro: '',
                  mensagem: 'Incremental concluída: ' + total + ' linhas de campanha×dia gravadas.' });
  } catch (e) {
    gravarSync_({ status: 'erro', ultimoErro: agoraTxt_() + ' — ' + e.message, mensagem: 'Falhou: ' + e.message });
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** Historico inteiro desde DATA_INICIO, mes a mes. Retomavel: o Apps Script
    corta a execucao em 6 minutos, entao o progresso fica num cursor e um
    gatilho de 1 minuto chama continuarSincronizacao(). */
function sincronizarCompleta() {
  props_().setProperty('cursor', JSON.stringify({ conta: 0, mes: cfg_('DATA_INICIO').slice(0, 7), gravadas: 0 }));
  limparGatilhosContinuacao_();
  gravarSync_({ status: 'rodando', etapa: 'completa', progresso: 0, mensagem: 'Iniciando histórico…' });
  continuarSincronizacao();
}

function continuarSincronizacao() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  var t0 = Date.now();
  try {
    limparGatilhosContinuacao_();
    var cur = JSON.parse(props_().getProperty('cursor') || 'null');
    if (!cur) return;
    var contas = contas_();
    var mesFinal = iso_(hoje_()).slice(0, 7);
    var totalMeses = mesesEntre_(cfg_('DATA_INICIO').slice(0, 7), mesFinal) * contas.length;

    while (cur.conta < contas.length) {
      while (cur.mes <= mesFinal) {
        var ini = new Date(cur.mes + '-01T00:00:00');
        var fim = new Date(ini.getFullYear(), ini.getMonth() + 1, 0);
        if (fim > hoje_()) fim = hoje_();
        cur.gravadas += sincronizarPeriodo_(contas[cur.conta], ini, fim);
        cur.mes = proximoMes_(cur.mes);
        props_().setProperty('cursor', JSON.stringify(cur));
        var feitos = cur.conta * mesesEntre_(cfg_('DATA_INICIO').slice(0, 7), mesFinal) +
                     mesesEntre_(cfg_('DATA_INICIO').slice(0, 7), cur.mes) - 1;
        gravarSync_({ status: 'rodando', etapa: 'completa',
                      progresso: Math.min(99, Math.round(feitos / Math.max(1, totalMeses) * 100)),
                      mensagem: 'Conta ' + contas[cur.conta] + ' · até ' + cur.mes + ' · ' + cur.gravadas + ' linhas' });
        /* Perto do limite de 6 min: agenda a continuacao e sai limpo. */
        if (Date.now() - t0 > 4.5 * 60 * 1000) {
          ScriptApp.newTrigger('continuarSincronizacao').timeBased().after(60 * 1000).create();
          return;
        }
      }
      cur.conta++;
      cur.mes = cfg_('DATA_INICIO').slice(0, 7);
      props_().setProperty('cursor', JSON.stringify(cur));
    }
    atualizarCampanhas_();
    atualizarJanelasEAnuncios_();
    props_().deleteProperty('cursor');
    gravarSync_({ status: 'ocioso', etapa: 'completa', progresso: 100, ultimaOk: agoraTxt_(), ultimoErro: '',
                  mensagem: 'Histórico concluído: ' + cur.gravadas + ' linhas de campanha×dia.' });
  } catch (e) {
    gravarSync_({ status: 'erro', ultimoErro: agoraTxt_() + ' — ' + e.message, mensagem: 'Falhou: ' + e.message });
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function atualizarCampanhasEAnuncios() {
  atualizarCampanhas_();
  atualizarJanelasEAnuncios_();
  gravarSync_({ etapa: 'campanhas', ultimaOk: agoraTxt_(), mensagem: 'Campanhas, janelas e anúncios atualizados.' });
}

/* ------------------------------------------------ insights diários */

/** Busca campanha x dia no periodo e grava (substituindo o que ja havia
    para a mesma conta e mesmas datas). Devolve quantas linhas gravou. */
function sincronizarPeriodo_(conta, ini, fim) {
  var itens = listarTudo_('act_' + conta + '/insights', {
    level: 'campaign',
    time_increment: 1,
    time_range: JSON.stringify({ since: iso_(ini), until: iso_(fim) }),
    fields: CAMPOS_INSIGHT,
    limit: 500
  });
  var tipos = tiposLead_();
  var agora = agoraTxt_();
  var linhas = itens.map(function (it) {
    var a = acoes_(it.actions);
    var leads = contarLeads_(a, tipos);
    return [
      it.date_start, String(conta), it.campaign_id, it.campaign_name,
      num_(it.spend), num_(it.impressions), num_(it.reach),
      num_(it.clicks), num_(it.inline_link_clicks), a['landing_page_view'] || 0,
      leads.formulario, leads.conversa, leads.pixel, leads.total,
      JSON.stringify(it.actions || []), agora
    ];
  });
  gravarInsights_(String(conta), iso_(ini), iso_(fim), linhas);
  return linhas.length;
}

/** Substitui, na aba de insights, as linhas da conta dentro do intervalo. */
function gravarInsights_(conta, since, until, novas) {
  var ss = planilha_();
  var aba = garantirAba_(ss, ABA.insights, CAB.insights);
  var ult = aba.getLastRow();
  var antigas = ult > 1 ? aba.getRange(2, 1, ult - 1, CAB.insights.length).getValues() : [];
  var mantidas = antigas.filter(function (l) {
    var d = dataTxt_(l[0]);
    var mesma = String(l[1]) === conta && d >= since && d <= until;
    return !mesma && String(l[0] || '').trim();
  });
  var todas = mantidas.concat(novas);
  todas.sort(function (a, b) {
    var da = dataTxt_(a[0]), db = dataTxt_(b[0]);
    return da < db ? -1 : da > db ? 1 : String(a[3]).localeCompare(String(b[3]));
  });
  if (ult > 1) aba.getRange(2, 1, ult - 1, CAB.insights.length).clearContent();
  if (todas.length) aba.getRange(2, 1, todas.length, CAB.insights.length).setValues(todas);
}

/* ------------------------------------------------- campanhas / orçamento */

function atualizarCampanhas_() {
  var ss = planilha_();
  var aba = garantirAba_(ss, ABA.campanhas, CAB.campanhas);
  var agora = agoraTxt_();
  var linhas = [];
  contas_().forEach(function (conta) {
    var campanhas = listarTudo_('act_' + conta + '/campaigns', {
      fields: 'id,name,status,effective_status,objective,daily_budget,lifetime_budget,start_time,stop_time',
      limit: 200
    });
    /* Quando o orcamento e por conjunto (ABO), a campanha vem sem orcamento:
       soma-se o dos conjuntos ativos, senao o painel mostraria zero. */
    var conjuntos = listarTudo_('act_' + conta + '/adsets', {
      fields: 'id,campaign_id,daily_budget,lifetime_budget,effective_status',
      limit: 500
    });
    var porCampanha = {};
    conjuntos.forEach(function (c) {
      var k = c.campaign_id;
      (porCampanha[k] = porCampanha[k] || { diario: 0, vitalicio: 0, ativos: 0 });
      if (c.effective_status === 'ACTIVE') {
        porCampanha[k].diario += centavos_(c.daily_budget);
        porCampanha[k].vitalicio += centavos_(c.lifetime_budget);
        porCampanha[k].ativos++;
      }
    });
    campanhas.forEach(function (c) {
      var propria = centavos_(c.daily_budget) || centavos_(c.lifetime_budget);
      var dosConjuntos = porCampanha[c.id] || { diario: 0, vitalicio: 0, ativos: 0 };
      linhas.push([
        String(conta), c.id, c.name, c.status, c.effective_status, c.objective,
        propria ? centavos_(c.daily_budget) : dosConjuntos.diario,
        propria ? centavos_(c.lifetime_budget) : dosConjuntos.vitalicio,
        propria ? 'campanha' : 'conjuntos',
        dosConjuntos.ativos,
        dataTxt_(c.start_time), dataTxt_(c.stop_time), agora
      ]);
    });
  });
  reescrever_(aba, CAB.campanhas, linhas);
}

/* ------------------------------------------- janelas e anúncios (gestor) */

function atualizarJanelasEAnuncios_() {
  var ss = planilha_();
  var abaJ = garantirAba_(ss, ABA.janelas, CAB.janelas);
  var abaA = garantirAba_(ss, ABA.anuncios, CAB.anuncios);
  var tipos = tiposLead_();
  var agora = agoraTxt_();
  var lj = [], la = [];

  contas_().forEach(function (conta) {
    JANELAS.forEach(function (j) {
      listarTudo_('act_' + conta + '/insights', {
        level: 'campaign', date_preset: j.preset, fields: CAMPOS_INSIGHT, limit: 500
      }).forEach(function (it) {
        var a = acoes_(it.actions), leads = contarLeads_(a, tipos);
        var gasto = num_(it.spend);
        lj.push([
          j.chave, String(conta), it.campaign_id, it.campaign_name, it.date_start, it.date_stop,
          gasto, num_(it.impressions), num_(it.reach), num_(it.frequency), num_(it.cpm),
          num_(it.ctr), num_(it.cpc), num_(it.clicks), num_(it.inline_link_clicks),
          num_(it.inline_link_click_ctr), a['landing_page_view'] || 0,
          leads.formulario, leads.conversa, leads.pixel, leads.total,
          leads.total ? gasto / leads.total : '', agora
        ]);
      });
      if (j.chave === 'mes') return;   // anuncio so em 7 e 30 dias
      listarTudo_('act_' + conta + '/insights', {
        level: 'ad', date_preset: j.preset, limit: 500,
        fields: 'ad_id,ad_name,adset_id,adset_name,' + CAMPOS_INSIGHT
      }).forEach(function (it) {
        var a = acoes_(it.actions), leads = contarLeads_(a, tipos);
        var gasto = num_(it.spend);
        la.push([
          j.chave, String(conta), it.campaign_id, it.campaign_name, it.adset_id, it.adset_name,
          it.ad_id, it.ad_name, gasto, num_(it.impressions), num_(it.reach), num_(it.frequency),
          num_(it.cpm), num_(it.ctr), num_(it.cpc), num_(it.inline_link_clicks),
          num_(it.inline_link_click_ctr), a['landing_page_view'] || 0,
          leads.formulario, leads.conversa, leads.pixel, leads.total,
          leads.total ? gasto / leads.total : '', agora
        ]);
      });
    });
  });
  reescrever_(abaJ, CAB.janelas, lj);
  reescrever_(abaA, CAB.anuncios, la);
}

/* =================================================================== META */

/** Uma chamada ao Graph API com repeticao em erro transitorio. */
function chamarMeta_(caminho, params) {
  var token = PropertiesService.getScriptProperties().getProperty('META_ACCESS_TOKEN');
  if (!token) throw new Error('META_ACCESS_TOKEN não está nas Propriedades do script.');
  var q = Object.keys(params || {}).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  });
  q.push('access_token=' + encodeURIComponent(token));
  var url = caminho.indexOf('http') === 0 ? caminho : API + caminho + '?' + q.join('&');

  var ultimo = null;
  for (var tent = 1; tent <= 4; tent++) {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var codigo = resp.getResponseCode();
    var corpo;
    try { corpo = JSON.parse(resp.getContentText()); } catch (e) { corpo = null; }
    if (codigo === 200 && corpo && !corpo.error) return corpo;
    var erro = (corpo && corpo.error) || { message: 'HTTP ' + codigo, code: codigo };
    ultimo = erro;
    /* Limite de taxa (4, 17, 32, 613) e 5xx melhoram esperando; o resto nao. */
    var transitorio = codigo >= 500 || [4, 17, 32, 613].indexOf(Number(erro.code)) >= 0;
    if (!transitorio) break;
    Utilities.sleep(1500 * tent * tent);
  }
  throw new Error('Meta: ' + (ultimo.message || JSON.stringify(ultimo)) +
                  (ultimo.code ? ' (código ' + ultimo.code + ')' : ''));
}

/** Segue o paging.next ate o fim. */
function listarTudo_(caminho, params) {
  var out = [];
  var pagina = chamarMeta_(caminho, params);
  var guarda = 0;
  while (pagina) {
    (pagina.data || []).forEach(function (x) { out.push(x); });
    var prox = pagina.paging && pagina.paging.next;
    if (!prox || ++guarda > 200) break;
    Utilities.sleep(250);
    pagina = chamarMeta_(prox, null);
  }
  return out;
}

/** actions [{action_type, value}] -> { tipo: valor } */
function acoes_(lista) {
  var m = {};
  (lista || []).forEach(function (a) { m[a.action_type] = num_(a.value); });
  return m;
}

function tiposLead_() {
  var lista = function (k) { return String(cfg_(k) || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean); };
  return { formulario: lista('TIPOS_LEAD_FORMULARIO'), conversa: lista('TIPOS_LEAD_CONVERSA'),
           pixel: lista('TIPOS_LEAD_PIXEL'), total: String(cfg_('LEAD_TOTAL') || 'formulario+conversa') };
}

function contarLeads_(a, tipos) {
  var soma = function (ts) { return ts.reduce(function (s, t) { return s + (a[t] || 0); }, 0); };
  var r = { formulario: soma(tipos.formulario), conversa: soma(tipos.conversa), pixel: soma(tipos.pixel) };
  r.total = tipos.total.split('+').reduce(function (s, k) { return s + (r[k.trim()] || 0); }, 0);
  return r;
}

/* ================================================================ WEB APP */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var acao = p.acao || 'status';
  var saida;
  try {
    if (acao === 'atualizar') {
      if (String(p.token || '') !== String(cfg_('TOKEN_WEBAPP'))) {
        saida = { ok: false, erro: 'Token inválido.' };
      } else {
        var st = lerSync_();
        if (st.status === 'rodando') {
          saida = { ok: true, jaRodando: true, sync: st };
        } else {
          /* Dispara e responde na hora: a busca leva mais que o navegador espera. */
          ScriptApp.newTrigger('sincronizarIncremental').timeBased().after(1000).create();
          gravarSync_({ status: 'rodando', etapa: 'incremental', progresso: 0, mensagem: 'Disparada pelo painel…' });
          saida = { ok: true, iniciada: true };
        }
      }
    } else {
      saida = { ok: true, sync: lerSync_() };
    }
  } catch (err) {
    saida = { ok: false, erro: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================ UTILITÁRIOS */

/* Com o script criado fora da planilha nao ha interface: o aviso vai para o
   Registro de execucao. */
function alertar_(texto) {
  try { SpreadsheetApp.getUi().alert(texto); } catch (e) { Logger.log(texto); }
}

/* ID da planilha "Campanhas Meta Ads". Nao e segredo (a planilha e compartilhada
   por link); fica aqui para o script funcionar tambem quando criado FORA da
   planilha, no script.google.com — caso em que getActiveSpreadsheet() e nulo
   e o menu da planilha nao existe (as funcoes rodam pelo editor). */
var PLANILHA_ID = '1vwULTReUh2vUoVE4W9TuAKo1inOLg7ju0oIv_tntPvg';

var PLANILHA_CACHE_ = null;
function planilha_() {
  if (PLANILHA_CACHE_) return PLANILHA_CACHE_;
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) {
    var id = PropertiesService.getScriptProperties().getProperty('PLANILHA_ID') || PLANILHA_ID;
    ss = SpreadsheetApp.openById(id);
  }
  PLANILHA_CACHE_ = ss;
  return ss;
}
function props_() { return PropertiesService.getScriptProperties(); }

function garantirAba_(ss, nome, cab) {
  var aba = ss.getSheetByName(nome);
  if (!aba) aba = ss.insertSheet(nome);
  if (cab && aba.getLastRow() === 0) {
    aba.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

function reescrever_(aba, cab, linhas) {
  var ult = aba.getLastRow();
  if (ult > 1) aba.getRange(2, 1, ult - 1, cab.length).clearContent();
  if (linhas.length) aba.getRange(2, 1, linhas.length, cab.length).setValues(linhas);
}

var DESCRICAO_CONFIG = {
  META_AD_ACCOUNTS: 'IDs das contas de anúncio, separados por vírgula (sem "act_").',
  DATA_INICIO: 'Primeiro dia do histórico (AAAA-MM-DD).',
  DIAS_INCREMENTAL: 'Quantos dias para trás a sincronização de hora em hora refaz.',
  TOKEN_WEBAPP: 'Token que o painel envia ao Web App para pedir atualização.',
  TIPOS_LEAD_FORMULARIO: 'Tipos de ação do Meta contados como lead de formulário.',
  TIPOS_LEAD_CONVERSA: 'Tipos de ação contados como conversa iniciada (WhatsApp/Messenger).',
  TIPOS_LEAD_PIXEL: 'Tipos de ação contados como lead do pixel (site).',
  LEAD_TOTAL: 'Quais famílias somam no "Leads (total)": formulario, conversa, pixel, unidas por +.'
};

function garantirConfig_(ss) {
  var aba = garantirAba_(ss, ABA.config, CAB.config);
  var atuais = {};
  var ult = aba.getLastRow();
  if (ult > 1) aba.getRange(2, 1, ult - 1, 2).getValues().forEach(function (l) { if (l[0]) atuais[String(l[0])] = true; });
  var novas = [];
  Object.keys(PADRAO_CONFIG).forEach(function (k) {
    if (!atuais[k]) novas.push([k, PADRAO_CONFIG[k], DESCRICAO_CONFIG[k] || '']);
  });
  if (novas.length) aba.getRange(aba.getLastRow() + 1, 1, novas.length, 3).setValues(novas);
  aba.autoResizeColumns(1, 3);
}

function cfg_(chave) {
  var aba = planilha_().getSheetByName(ABA.config);
  if (aba) {
    var ult = aba.getLastRow();
    if (ult > 1) {
      var v = aba.getRange(2, 1, ult - 1, 2).getValues();
      for (var i = 0; i < v.length; i++) if (String(v[i][0]) === chave) return String(v[i][1]);
    }
  }
  return PADRAO_CONFIG[chave];
}

function contas_() {
  return String(cfg_('META_AD_ACCOUNTS') || '').split(',')
    .map(function (s) { return s.trim().replace(/^act_/, ''); }).filter(Boolean);
}

function gravarSync_(campos) {
  var aba = garantirAba_(planilha_(), ABA.sync, CAB.sync);
  var atual = lerSync_();
  Object.keys(campos).forEach(function (k) { atual[k] = campos[k]; });
  atual.atualizadoEm = agoraTxt_();
  var linhas = Object.keys(atual).map(function (k) { return [k, atual[k]]; });
  var ult = aba.getLastRow();
  if (ult > 1) aba.getRange(2, 1, ult - 1, 2).clearContent();
  aba.getRange(2, 1, linhas.length, 2).setValues(linhas);
}

function lerSync_() {
  var aba = planilha_().getSheetByName(ABA.sync);
  var o = {};
  if (!aba) return o;
  var ult = aba.getLastRow();
  if (ult > 1) aba.getRange(2, 1, ult - 1, 2).getValues().forEach(function (l) { if (l[0]) o[String(l[0])] = l[1]; });
  return o;
}

function limparGatilhosContinuacao_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'continuarSincronizacao') ScriptApp.deleteTrigger(t);
  });
}

function num_(v) { var n = Number(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) ? n : 0; }
/* Orcamentos chegam em centavos, como texto. */
function centavos_(v) { return v ? num_(v) / 100 : 0; }

function hoje_() { var d = new Date(); d.setHours(0, 0, 0, 0); return d; }
function iso_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function agoraTxt_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); }
function dataTxt_(v) {
  if (!v) return '';
  if (v instanceof Date) return iso_(v);
  return String(v).slice(0, 10);
}
function proximoMes_(am) {
  var a = Number(am.slice(0, 4)), m = Number(am.slice(5, 7));
  m++; if (m > 12) { m = 1; a++; }
  return a + '-' + (m < 10 ? '0' : '') + m;
}
function mesesEntre_(de, ate) {
  var a1 = Number(de.slice(0, 4)), m1 = Number(de.slice(5, 7));
  var a2 = Number(ate.slice(0, 4)), m2 = Number(ate.slice(5, 7));
  return (a2 - a1) * 12 + (m2 - m1) + 1;
}
