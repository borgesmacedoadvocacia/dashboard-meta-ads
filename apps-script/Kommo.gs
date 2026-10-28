/* =========================================================================
   BM · Meta Ads — REUNIÕES MARCADAS (Kommo)

   O painel precisa saber quantas reuniões cada campanha gerou. No Kommo isso
   é a entrada do lead na etapa "Enviado ao ADVBOX" — existe nos dois funis
   comerciais ([API] e [SP] CRM Comercial) e é o momento em que a reunião foi
   de fato marcada.

   Conta-se a ENTRADA na etapa, não quem está parado nela: o lead que já
   avançou para "Pesquisa de Satisfação" ou foi ganho continua tendo marcado
   reunião. Por isso a leitura é pela API de eventos (lead_status_changed
   filtrada por value_after), e não pela lista de leads da etapa.

   LEADS ELEGÍVEIS seguem a mesma regra, com outra lista de etapas: conta o
   lead na PRIMEIRA vez que entrou em qualquer etapa de "Elegíveis" em
   diante (KOMMO_ETAPAS_ELEGIVEL). Quem pulou direto para a qualificação ou
   foi perdido depois dela também passou pela elegibilidade; Inelegível e
   "Sem resposta: ELEGIBILIDADE" ficam de fora. Vão para a aba "Elegíveis".

   A campanha vem do campo personalizado "Campanha" do lead, que o Kommo
   recebe do Meta com o NOME EXATO da campanha — é o que permite casar com a
   aba Campanhas sem adivinhação. Lead sem esse campo preenchido (≈20% em
   2026) vira reunião sem campanha, contada no total e apontada no painel.

   INSTALAÇÃO
   1. Propriedades do script → KOMMO_TOKEN (chave de longa duração do Kommo).
   2. Rodar configurarKommo() uma vez (cria a aba e as chaves em _Config).
   3. Rodar verificarKommo() para conferir token e etapas.
   4. Rodar sincronizarKommoCompleta() para o histórico.
   A partir daí a incremental de hora em hora mantém a aba em dia sozinha.
   ========================================================================= */

var KOMMO_ABA = 'Reuniões';
var KOMMO_ABA_ELEGIVEIS = 'Elegíveis';
var KOMMO_ABA_QUALIFICADOS = 'Qualificados';
var KOMMO_CAB = ['Data', 'Hora', 'Funil', 'ID do lead', 'Lead', 'Campanha', 'Origem',
                 'Produto jurídico', 'Criado em', 'Atualizado em'];

var KOMMO_PADRAO = {
  KOMMO_SUBDOMINIO: 'bmadvocacia',
  KOMMO_ETAPA_REUNIAO: 'Enviado ao ADVBOX',
  KOMMO_ETAPAS_ELEGIVEL: 'Elegíveis; Análise de Qualificação; Respondeu FUP Qualificação; ' +
    'Aguardando Documentação; Respondeu FUP Documentação; Análise de Viabilidade; Agendar Reunião; ' +
    'Enviado ao ADVBOX; Pesquisa de Satisfação; Desqualificado; Inviável (Não tem abusividade); ' +
    'Sem resposta: QUALIFICAÇÃO; Sem resposta: AGUARDANDO DOCUMENTO; Sem resposta: AGENDAR REUNIÃO; ' +
    'Closed - won; Venda ganha',
  /* Um degrau acima: qualificado é quem chegou à documentação. Sai da lista
     de elegível o trecho da qualificação em si (Elegíveis, Análise de
     Qualificação, Respondeu FUP Qualificação) e quem morreu lá
     (Desqualificado, Sem resposta: QUALIFICAÇÃO). Entram as etapas de FUP -
     Documentação, porque o lead pode ir da qualificação direto para o funil
     de follow up e ainda assim ter sido qualificado. */
  KOMMO_ETAPAS_QUALIFICADO: 'Aguardando Documentação; Respondeu FUP Documentação; ' +
    'Análise de Viabilidade; Agendar Reunião; Enviado ao ADVBOX; Pesquisa de Satisfação; ' +
    'Inviável (Não tem abusividade); Sem resposta: AGUARDANDO DOCUMENTO; ' +
    'Sem resposta: AGENDAR REUNIÃO; Closed - won; Venda ganha; ' +
    'FUP 1 - Documentação; FUP 2 - Documentação; FUP 3 - Documentação; FUP 4 - Documentação; ' +
    'FUP 5 - Documentação; FUP 6 - Documentação; FUP 7 - Documentação',
  KOMMO_CAMPO_CAMPANHA: 'Campanha'
};

var KOMMO_DESCRICAO = {
  KOMMO_SUBDOMINIO: 'Subdomínio do Kommo (https://SUBDOMÍNIO.kommo.com).',
  KOMMO_ETAPA_REUNIAO: 'Nome da etapa que marca a reunião. Vale em todos os funis que a tiverem.',
  KOMMO_ETAPAS_ELEGIVEL: 'Etapas, separadas por ponto e vírgula, que contam o lead como elegível (de "Elegíveis" em diante).',
  KOMMO_ETAPAS_QUALIFICADO: 'Etapas, separadas por ponto e vírgula, que contam o lead como qualificado (de "Aguardando Documentação" em diante).',
  KOMMO_CAMPO_CAMPANHA: 'Campo personalizado do lead que traz o nome da campanha do Meta.'
};

/* Lote de leads por chamada. O UrlFetchApp corta a URL em 2 KB e cada id
   custa ~30 caracteres já codificado (filter%5Bid%5D%5B%5D=29245944&), então
   100 estourava o limite: 40 deixa a URL em ~1,3 KB e cobre o histórico
   inteiro (≈1.050 leads) em 27 chamadas. */
var KOMMO_LOTE = 40;

/* ============================================================ COMANDOS */

function configurarKommo() {
  var ss = planilha_();
  garantirAba_(ss, KOMMO_ABA, KOMMO_CAB);
  garantirAba_(ss, KOMMO_ABA_ELEGIVEIS, KOMMO_CAB);
  garantirAba_(ss, KOMMO_ABA_QUALIFICADOS, KOMMO_CAB);
  var aba = garantirAba_(ss, ABA.config, CAB.config);
  var atuais = {};
  var ult = aba.getLastRow();
  if (ult > 1) aba.getRange(2, 1, ult - 1, 2).getValues().forEach(function (l) { if (l[0]) atuais[String(l[0])] = true; });
  var novas = [];
  Object.keys(KOMMO_PADRAO).forEach(function (k) {
    if (!atuais[k]) novas.push([k, KOMMO_PADRAO[k], KOMMO_DESCRICAO[k] || '']);
  });
  if (novas.length) aba.getRange(aba.getLastRow() + 1, 1, novas.length, 3).setValues(novas);
  alertar_('Aba "' + KOMMO_ABA + '" e chaves do Kommo criadas.\n\n' +
    'Agora: Configurações do projeto → Propriedades do script → KOMMO_TOKEN.\n' +
    'Depois: verificarKommo e sincronizarKommoCompleta.');
}

/** Diz se o token vale, quais etapas achou e quantas reuniões há hoje. */
function verificarKommo() {
  var etapas = etapasReuniao_();
  if (!etapas.length) {
    alertar_('Token OK, mas nenhum funil tem a etapa "' + cfgKommo_('KOMMO_ETAPA_REUNIAO') + '".\n' +
      'Confira o nome em _Config → KOMMO_ETAPA_REUNIAO.');
    return;
  }
  var ini = new Date(hoje_().getTime() - 29 * 86400000);
  var achados = coletarReunioes_(ini, new Date());
  var comCampanha = 0;
  achados.forEach(function (l) { if (l[5]) comCampanha++; });
  alertar_('Kommo respondendo.\n\n' +
    'Etapas encontradas:\n  ' + etapas.map(function (e) { return e.funil + ' → ' + e.etapa; }).join('\n  ') + '\n\n' +
    'Últimos 30 dias: ' + achados.length + ' reunião(ões), ' + comCampanha + ' com campanha identificada.');
}

/** Últimos DIAS_INCREMENTAL dias. Chamada pela sincronização de hora em hora. */
function sincronizarKommo() {
  var dias = Number(cfg_('DIAS_INCREMENTAL')) || 35;
  var ini = new Date(hoje_().getTime() - (dias - 1) * 86400000);
  var fim = new Date();
  var novas = coletarReunioes_(ini, fim);
  var total = mesclarReunioes_(novas, ini, fim, false);
  var novasEleg = coletarElegiveis_(ini, fim);
  var totalEleg = mesclarEntradas_(KOMMO_ABA_ELEGIVEIS, novasEleg, ini, fim, false);
  var novasQual = coletarQualificados_(ini, fim);
  var totalQual = mesclarEntradas_(KOMMO_ABA_QUALIFICADOS, novasQual, ini, fim, false);
  gravarSync_({ kommoOk: agoraTxt_(), kommoErro: '',
                kommoMensagem: novas.length + ' reunião(ões), ' + novasEleg.length + ' elegível(is) e ' +
                  novasQual.length + ' qualificado(s) na janela · ' + total + ' reuniões, ' +
                  totalEleg + ' elegíveis e ' + totalQual + ' qualificados nas abas' });
  return total;
}

/** Histórico inteiro de reuniões desde DATA_INICIO. Refaz a aba do zero.
    Os elegíveis têm comando próprio: são muito mais eventos e não cabem
    junto nos 6 minutos de uma execução. */
function sincronizarKommoCompleta() {
  var ini = dataDeIso_(cfgTexto_(cfg_('DATA_INICIO'))) || new Date(2025, 0, 1);
  var fim = new Date();
  gravarSync_({ kommoMensagem: 'Buscando reuniões desde ' + iso_(ini) + '…' });
  var novas = coletarReunioes_(ini, fim);
  var total = mesclarReunioes_(novas, ini, fim, true);
  gravarSync_({ kommoOk: agoraTxt_(), kommoErro: '',
                kommoMensagem: 'Histórico de reuniões: ' + total + ' desde ' + iso_(ini) });
  alertar_(total + ' reunião(ões) gravadas na aba "' + KOMMO_ABA + '", desde ' + iso_(ini) + '.\n\n' +
    'Para os leads elegíveis: BM · Meta Ads → Elegíveis do Kommo (histórico).');
  return total;
}

/** Histórico de elegíveis, um mês por vez. Cada mês é gravado assim que
    termina e a execução para antes do limite de 6 minutos, guardando onde
    parou: basta rodar de novo até o aviso dizer que acabou. */
function sincronizarElegiveisCompleta() {
  return historicoMensal_(KOMMO_ABA_ELEGIVEIS, coletarElegiveis_, 'KOMMO_ELEGIVEIS_CURSOR',
    'elegível(is)', 'Elegíveis do Kommo (histórico)');
}

/** Histórico de qualificados, pelo mesmo caminho dos elegíveis. */
function sincronizarQualificadosCompleta() {
  return historicoMensal_(KOMMO_ABA_QUALIFICADOS, coletarQualificados_, 'KOMMO_QUALIFICADOS_CURSOR',
    'qualificado(s)', 'Qualificados do Kommo (histórico)');
}

/** Carga de histórico mês a mês, retomável: grava cada mês assim que termina
    e para antes dos 6 minutos guardando onde parou. Serve para qualquer aba
    de entradas — o que muda é a lista de etapas de cada coletor. */
function historicoMensal_(nomeAba, coletor, chaveCursor, rotulo, comandoMenu) {
  /* A incremental de hora em hora escreve nas mesmas abas. Sem o mesmo lock
     que ela usa, as duas leem a aba, mesclam e reescrevem — e a última a
     gravar apaga as linhas da outra. */
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    alertar_('A sincronização de hora em hora está rodando agora e escreve nas mesmas abas.\n' +
             'Espere um minuto e rode "' + comandoMenu + '" de novo.');
    return;
  }
  try {
    return historicoMensalSemLock_(nomeAba, coletor, chaveCursor, rotulo, comandoMenu);
  } finally {
    lock.releaseLock();
  }
}

function historicoMensalSemLock_(nomeAba, coletor, chaveCursor, rotulo, comandoMenu) {
  var props = props_();
  var inicioCfg = dataDeIso_(cfgTexto_(cfg_('DATA_INICIO'))) || new Date(2025, 0, 1);
  var cursor = dataDeIso_(props.getProperty(chaveCursor));
  if (!cursor) {
    reescrever_(garantirAba_(planilha_(), nomeAba, KOMMO_CAB), KOMMO_CAB, []);
    cursor = inicioCfg;
  }
  var t0 = Date.now(), agora = new Date();
  while (cursor < agora) {
    if (Date.now() - t0 > 4 * 60 * 1000) {
      props.setProperty(chaveCursor, iso_(cursor));
      alertar_(nomeAba + ' gravados até ' + iso_(new Date(cursor.getTime() - 86400000)) +
        '. Rode "' + comandoMenu + '" de novo para continuar.');
      return;
    }
    var fimMes = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
    var fim = fimMes < agora ? new Date(fimMes.getTime() - 1000) : agora;
    gravarSync_({ kommoMensagem: 'Buscando ' + nomeAba + ' de ' + iso_(cursor) + ' a ' + iso_(fim) + '…' });
    mesclarEntradas_(nomeAba, coletor(cursor, fim), cursor, fim, false);
    cursor = fimMes;
  }
  props.deleteProperty(chaveCursor);
  var ult = planilha_().getSheetByName(nomeAba).getLastRow();
  var total = Math.max(ult - 1, 0);
  gravarSync_({ kommoOk: agoraTxt_(), kommoErro: '',
                kommoMensagem: 'Histórico de ' + nomeAba + ': ' + total + ' desde ' + iso_(inicioCfg) });
  alertar_(total + ' lead(s) ' + rotulo + ' gravados na aba "' + nomeAba + '", desde ' + iso_(inicioCfg) + '.');
  return total;
}

/* ============================================================== COLETA */

/** Toda etapa com o nome configurado, em qualquer funil. Resolver por NOME
    (e não por ID fixo) é o que mantém isto de pé quando alguém recria o
    funil ou duplica a etapa num funil novo. */
function etapasReuniao_() {
  return etapasPorNome_([cfgKommo_('KOMMO_ETAPA_REUNIAO')]);
}

/** Etapas de "Elegíveis" em diante, em todos os funis que as tiverem. */
function etapasElegivel_() {
  return etapasPorNome_(String(cfgKommo_('KOMMO_ETAPAS_ELEGIVEL')).split(';'));
}

/** Etapas de "Aguardando Documentação" em diante. */
function etapasQualificado_() {
  return etapasPorNome_(String(cfgKommo_('KOMMO_ETAPAS_QUALIFICADO')).split(';'));
}

function etapasPorNome_(nomes) {
  var alvos = {};
  nomes.forEach(function (n) { if (normKommo_(n)) alvos[normKommo_(n)] = true; });
  var j = kommoGet_('/leads/pipelines', { limit: 250 });
  var funis = j && j._embedded && j._embedded.pipelines ? j._embedded.pipelines : [];
  var out = [];
  funis.forEach(function (p) {
    var sts = (p._embedded && p._embedded.statuses) || [];
    sts.forEach(function (s) {
      if (alvos[normKommo_(s.name)]) out.push({ funil: p.name, funilId: p.id, etapa: s.name, etapaId: s.id });
    });
  });
  return out;
}

/** leadId -> { quando, funil } da PRIMEIRA entrada na etapa, entre ini e fim.
    Vive sozinho porque o painel de pré-qualificação (KommoLeads.gs) precisa
    exatamente deste mapa, e duas definições de "reunião marcada" seria uma a
    mais do que o escritório aguenta conferir. */
function primeirasEntradasAdvbox_(ini, fim) {
  var etapas = etapasReuniao_();
  if (!etapas.length) throw new Error('Nenhum funil tem a etapa "' + cfgKommo_('KOMMO_ETAPA_REUNIAO') + '".');
  return primeirasEntradas_(etapas, ini, fim);
}

/** leadId -> { quando, funil } da PRIMEIRA entrada em qualquer uma das etapas. */
function primeirasEntradas_(etapas, ini, fim) {
  var de = Math.floor(ini.getTime() / 1000), ate = Math.floor(fim.getTime() / 1000);
  var primeiro = {};
  etapas.forEach(function (et) {
    for (var pagina = 1; pagina <= 80; pagina++) {
      var j = kommoGet_('/events', {
        limit: 250, page: pagina,
        'filter[type]': 'lead_status_changed',
        'filter[created_at][from]': de,
        'filter[created_at][to]': ate,
        'filter[value_after][leads_statuses][0][pipeline_id]': et.funilId,
        'filter[value_after][leads_statuses][0][status_id]': et.etapaId
      });
      var ev = j && j._embedded && j._embedded.events ? j._embedded.events : [];
      for (var i = 0; i < ev.length; i++) {
        var id = String(ev[i].entity_id), q = ev[i].created_at;
        /* Lead que vai e volta da etapa gera vários eventos: vale o primeiro,
           senão a mesma reunião mudaria de dia a cada arrastada no funil. */
        if (!primeiro[id] || q < primeiro[id].quando) primeiro[id] = { quando: q, funil: et.funil };
      }
      if (ev.length < 250) break;
    }
  });
  return primeiro;
}

/** Linhas prontas da aba, para as reuniões marcadas entre ini e fim. */
function coletarReunioes_(ini, fim) {
  return linhasDeEntradas_(primeirasEntradasAdvbox_(ini, fim));
}

/** Linhas prontas da aba Elegíveis, para quem virou elegível entre ini e fim. */
function coletarElegiveis_(ini, fim) {
  var etapas = etapasElegivel_();
  if (!etapas.length) throw new Error('Nenhum funil tem as etapas de KOMMO_ETAPAS_ELEGIVEL.');
  return linhasDeEntradas_(primeirasEntradas_(etapas, ini, fim));
}

/** Idem para os qualificados: primeira entrada numa etapa de documentação
    em diante. Todo qualificado passou pela elegibilidade antes — se a conta
    der qualificados acima de elegíveis, é sinal de lista desalinhada. */
function coletarQualificados_(ini, fim) {
  var etapas = etapasQualificado_();
  if (!etapas.length) throw new Error('Nenhum funil tem as etapas de KOMMO_ETAPAS_QUALIFICADO.');
  return linhasDeEntradas_(primeirasEntradas_(etapas, ini, fim));
}

function linhasDeEntradas_(primeiro) {
  var ids = Object.keys(primeiro);
  var leads = buscarLeads_(ids);
  var campoCamp = cfgKommo_('KOMMO_CAMPO_CAMPANHA');
  var agora = agoraTxt_();
  var tz = Session.getScriptTimeZone();
  return ids.map(function (id) {
    var e = primeiro[id], l = leads[id] || {};
    var d = new Date(e.quando * 1000);
    return [
      Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
      Utilities.formatDate(d, tz, 'HH:mm'),
      e.funil, id, l.nome || '',
      campoLead_(l, campoCamp), campoLead_(l, 'Origem'), campoLead_(l, 'Produto Jurídico'),
      l.criado ? Utilities.formatDate(new Date(l.criado * 1000), tz, 'yyyy-MM-dd') : '',
      agora
    ];
  }).sort(function (a, b) {
    return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : String(a[1]).localeCompare(String(b[1]));
  });
}

/** Leads por id, em lotes. Lead apagado simplesmente não volta — a reunião
    continua contada, só sem nome nem campanha. */
function buscarLeads_(ids) {
  var out = {};
  for (var i = 0; i < ids.length; i += KOMMO_LOTE) {
    var lote = ids.slice(i, i + KOMMO_LOTE);
    var j = kommoGet_('/leads', { limit: 250, 'filter[id][]': lote });
    var ls = j && j._embedded && j._embedded.leads ? j._embedded.leads : [];
    ls.forEach(function (l) {
      out[String(l.id)] = { nome: l.name, criado: l.created_at, campos: l.custom_fields_values || [] };
    });
  }
  return out;
}

function campoLead_(lead, nome) {
  var campos = lead.campos || [];
  var alvo = normKommo_(nome);
  for (var i = 0; i < campos.length; i++) {
    if (normKommo_(campos[i].field_name) !== alvo) continue;
    var vs = campos[i].values || [];
    var v = vs.length ? vs[0].value : '';
    /* O Kommo grava "-" quando o fluxo não soube a campanha: é tão vazio
       quanto vazio, e deixar passar criaria uma campanha chamada "-". */
    return v === '-' ? '' : String(v == null ? '' : v);
  }
  return '';
}

/* ============================================================== GRAVAÇÃO */

function mesclarReunioes_(novas, ini, fim, completa) {
  return mesclarEntradas_(KOMMO_ABA, novas, ini, fim, completa);
}

function mesclarEntradas_(nomeAba, novas, ini, fim, completa) {
  var aba = garantirAba_(planilha_(), nomeAba, KOMMO_CAB);
  var ult = aba.getLastRow();
  var antigas = ult > 1 ? aba.getRange(2, 1, ult - 1, KOMMO_CAB.length).getValues() : [];
  var deTxt = iso_(ini), ateTxt = iso_(fim);
  var porLead = {};

  if (!completa) antigas.forEach(function (l) {
    var id = String(l[3] || '').trim();
    if (!id) return;
    var d = dataTxt_(l[0]);
    /* Fora da janela a linha antiga é a única fonte e fica. Dentro da janela
       ela só fica se o Kommo confirmar de novo — é assim que lead apagado
       some da aba em vez de virar reunião fantasma. */
    if (d < deTxt || d > ateTxt) porLead[id] = l;
  });

  novas.forEach(function (l) {
    var id = String(l[3]);
    var ja = porLead[id];
    /* Um lead é uma reunião só, na primeira vez que entrou na etapa. */
    if (ja && dataTxt_(ja[0]) <= String(l[0])) return;
    porLead[id] = l;
  });

  var todas = Object.keys(porLead).map(function (k) { return porLead[k]; });
  todas.sort(function (a, b) {
    var da = dataTxt_(a[0]), db = dataTxt_(b[0]);
    return da < db ? -1 : da > db ? 1 : String(a[1]).localeCompare(String(b[1]));
  });
  reescrever_(aba, KOMMO_CAB, todas);
  return todas.length;
}

/* ================================================================ API */

function kommoGet_(caminho, params) {
  var token = props_().getProperty('KOMMO_TOKEN');
  if (!token) throw new Error('Falta KOMMO_TOKEN nas Propriedades do script.');
  var url = 'https://' + cfgKommo_('KOMMO_SUBDOMINIO') + '.kommo.com/api/v4' + caminho + queryKommo_(params);
  var opcoes = { method: 'get', muteHttpExceptions: true,
                 headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } };
  for (var tentativa = 1; tentativa <= 4; tentativa++) {
    var r;
    /* "Endereço não disponível" e afins chegam como exceção, não como código
       HTTP: são quedas de rede e merecem a mesma nova tentativa que um 5xx. */
    try { r = UrlFetchApp.fetch(url, opcoes); }
    catch (e) {
      if (tentativa === 4) throw e;
      Utilities.sleep(2000 * tentativa); continue;
    }
    var c = r.getResponseCode();
    if (c === 200) {
      try { return JSON.parse(r.getContentText()); }
      catch (e) { throw new Error('Kommo devolveu resposta inválida em ' + caminho); }
    }
    /* 204 = filtro sem nenhum resultado. É resposta boa, não erro. */
    if (c === 204) return null;
    if (c === 401 || c === 403) throw new Error('Kommo recusou o token (HTTP ' + c + '). Gere outro e atualize KOMMO_TOKEN.');
    if (c === 429 || c >= 500) { Utilities.sleep(1500 * tentativa); continue; }
    throw new Error('Kommo HTTP ' + c + ' em ' + caminho + ': ' + r.getContentText().slice(0, 200));
  }
  throw new Error('Kommo não respondeu depois de 4 tentativas em ' + caminho);
}

/** O Kommo filtra por chaves com colchetes aninhados; cada uma vai inteira,
    percent-encoded. Valor em array repete a chave (filter[id][]=1&...=2). */
function queryKommo_(params) {
  var qs = [];
  Object.keys(params || {}).forEach(function (k) {
    var v = params[k];
    if (v === null || v === undefined || v === '') return;
    if (Object.prototype.toString.call(v) === '[object Array]') {
      v.forEach(function (x) { qs.push(encodeURIComponent(k) + '=' + encodeURIComponent(x)); });
    } else {
      qs.push(encodeURIComponent(k) + '=' + encodeURIComponent(v));
    }
  });
  return qs.length ? '?' + qs.join('&') : '';
}

/* ============================================================ AUXILIARES */

function cfgKommo_(chave) {
  var v = cfg_(chave);
  return v === undefined || v === null || String(v) === '' ? KOMMO_PADRAO[chave] : cfgTexto_(v);
}

function dataDeIso_(s) {
  var m = String(s == null ? '' : s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

/** Minúsculas sem acento. Escrito à mão porque normalize('NFD') não existe
    no runtime antigo do Apps Script e este arquivo precisa rodar nos dois. */
function normKommo_(s) {
  return String(s == null ? '' : s).toLowerCase()
    .replace(/[áàâãä]/g, 'a').replace(/[éèêë]/g, 'e').replace(/[íìîï]/g, 'i')
    .replace(/[óòôõö]/g, 'o').replace(/[úùûü]/g, 'u').replace(/[ç]/g, 'c')
    .replace(/\s+/g, ' ').trim();
}
