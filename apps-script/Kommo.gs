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
var KOMMO_CAB = ['Data', 'Hora', 'Funil', 'ID do lead', 'Lead', 'Campanha', 'Origem',
                 'Produto jurídico', 'Criado em', 'Atualizado em'];

var KOMMO_PADRAO = {
  KOMMO_SUBDOMINIO: 'bmadvocacia',
  KOMMO_ETAPA_REUNIAO: 'Enviado ao ADVBOX',
  KOMMO_CAMPO_CAMPANHA: 'Campanha'
};

var KOMMO_DESCRICAO = {
  KOMMO_SUBDOMINIO: 'Subdomínio do Kommo (https://SUBDOMÍNIO.kommo.com).',
  KOMMO_ETAPA_REUNIAO: 'Nome da etapa que marca a reunião. Vale em todos os funis que a tiverem.',
  KOMMO_CAMPO_CAMPANHA: 'Campo personalizado do lead que traz o nome da campanha do Meta.'
};

/* Lote de leads por chamada. 100 deixa a URL em ~1,5 KB — longe do limite —
   e cobre o histórico inteiro (≈1.000 leads) em 10 chamadas. */
var KOMMO_LOTE = 100;

/* ============================================================ COMANDOS */

function configurarKommo() {
  var ss = planilha_();
  garantirAba_(ss, KOMMO_ABA, KOMMO_CAB);
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
  gravarSync_({ kommoOk: agoraTxt_(), kommoErro: '',
                kommoMensagem: novas.length + ' reunião(ões) na janela · ' + total + ' na aba' });
  return total;
}

/** Histórico inteiro desde DATA_INICIO. Refaz a aba do zero. */
function sincronizarKommoCompleta() {
  var ini = dataDeIso_(cfgTexto_(cfg_('DATA_INICIO'))) || new Date(2025, 0, 1);
  var fim = new Date();
  gravarSync_({ kommoMensagem: 'Buscando reuniões desde ' + iso_(ini) + '…' });
  var novas = coletarReunioes_(ini, fim);
  var total = mesclarReunioes_(novas, ini, fim, true);
  gravarSync_({ kommoOk: agoraTxt_(), kommoErro: '',
                kommoMensagem: 'Histórico de reuniões: ' + total + ' desde ' + iso_(ini) });
  alertar_(total + ' reunião(ões) gravadas na aba "' + KOMMO_ABA + '", desde ' + iso_(ini) + '.');
  return total;
}

/* ============================================================== COLETA */

/** Toda etapa com o nome configurado, em qualquer funil. Resolver por NOME
    (e não por ID fixo) é o que mantém isto de pé quando alguém recria o
    funil ou duplica a etapa num funil novo. */
function etapasReuniao_() {
  var alvo = normKommo_(cfgKommo_('KOMMO_ETAPA_REUNIAO'));
  var j = kommoGet_('/leads/pipelines', { limit: 250 });
  var funis = j && j._embedded && j._embedded.pipelines ? j._embedded.pipelines : [];
  var out = [];
  funis.forEach(function (p) {
    var sts = (p._embedded && p._embedded.statuses) || [];
    sts.forEach(function (s) {
      if (normKommo_(s.name) === alvo) out.push({ funil: p.name, funilId: p.id, etapa: s.name, etapaId: s.id });
    });
  });
  return out;
}

/** Linhas prontas da aba, para as reuniões marcadas entre ini e fim. */
function coletarReunioes_(ini, fim) {
  var etapas = etapasReuniao_();
  if (!etapas.length) throw new Error('Nenhum funil tem a etapa "' + cfgKommo_('KOMMO_ETAPA_REUNIAO') + '".');

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
  var aba = garantirAba_(planilha_(), KOMMO_ABA, KOMMO_CAB);
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
    var r = UrlFetchApp.fetch(url, opcoes);
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
