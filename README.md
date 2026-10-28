# Dashboard — Campanhas Meta Ads

Painel de resultado das campanhas do Meta Ads do escritório Borges Macedo Advocacia,
alimentado em tempo real pela planilha **"Campanhas Meta Ads"** (que o Apps Script
preenche a partir da Marketing API do Meta) e pelas planilhas **"Fluxo de Clientes"** de cada ano (2026 e 2025, aba DADOS), de onde vêm clientes e retorno financeiro.

**Acesso exclusivo das lideranças.** Os perfis Administração e Lideranças têm cofre
neste painel; a credencial do perfil Equipe não decifra nada.

## Como os dados chegam

```
Meta Ads (Graph API) ──► Apps Script (gatilho de hora em hora) ──► planilha Campanhas Meta Ads
Fluxo de Clientes › DADOS ──────────────────────────────────────────────┐
                                                                        ▼
                                                                  este painel
```

O Apps Script (`apps-script/Codigo.gs`, vinculado à planilha) grava:

| Aba | Conteúdo |
|---|---|
| Insights diários | uma linha por campanha por dia — gasto, impressões, alcance, cliques, leads por família (formulário, conversas, pixel) e ações brutas em JSON |
| Janelas | campanha × janela (7 dias, 30 dias, mês atual) com as métricas que não somam entre dias: alcance, frequência, CPM, CTR, CPC |
| Anúncios | anúncio × janela (7 e 30 dias) — é onde se vê qual criativo ganha e qual cansou |
| Campanhas | orçamento diário/vitalício (somando conjuntos quando o orçamento é por conjunto) e status |
| Mapeamento | ORIGEM do Fluxo de Clientes → nome exato da campanha; o que estiver aqui vence o casamento automático |
| Reuniões | uma linha por lead que entrou na etapa **Enviado ao ADVBOX** do Kommo — data, lead, campanha, origem e produto |
| Elegíveis | idem, para quem entrou numa etapa de **Elegíveis** em diante |
| Qualificados | idem, para quem chegou de **Aguardando Documentação** em diante |
| _Config | contas, data de início, tipos de ação que contam como lead, subdomínio e etapa do Kommo |
| _Sync | estado da última sincronização, mostrado no topo do painel |

O token do Meta **não fica no código**: vive em *Configurações do projeto → Propriedades
do script → `META_ACCESS_TOKEN`*. O do Kommo, idem, em `KOMMO_TOKEN`.

## Reuniões marcadas (Kommo)

`apps-script/Kommo.gs` traz para a planilha as reuniões marcadas. Reunião marcada = lead que
**entrou na etapa `Enviado ao ADVBOX`**, etapa que existe nos dois funis comerciais
(`[API]` e `[SP] CRM Comercial`).

- Conta-se a **entrada** na etapa, lida pela API de eventos (`lead_status_changed` filtrada por
  `value_after`), e não quem está parado nela: lead que já avançou para *Pesquisa de Satisfação*
  ou foi ganho marcou reunião do mesmo jeito. A data da reunião é a da mudança de etapa.
- Um lead é **uma reunião só**, na primeira vez que entrou — quem vai e volta no funil não vira
  duas.
- A campanha vem do campo personalizado **`Campanha`** do lead, que o Kommo recebe do Meta com o
  nome exato da campanha; por isso quase tudo casa por nome idêntico, e a semelhança (a mesma do
  Fluxo de Clientes) só entra quando a campanha foi renomeada depois. Lead com o campo vazio ou
  `-` (≈20% em 2026, entrada manual e indicação) fica fora da tabela e aparece em
  *Pontos de atenção*.
- Primeira instalação: `KOMMO_TOKEN` nas Propriedades do script → `configurarKommo` →
  `verificarKommo` → **BM · Meta Ads → Reuniões do Kommo (histórico)**. Depois disso a
  incremental de hora em hora mantém a aba sozinha; se o Kommo falhar, a sincronização do Meta
  segue normalmente e o erro fica em `_Sync` (`kommoErro`).
- O histórico do Kommo começa em **junho de 2025** — antes disso a etapa não era usada.
- **Elegíveis e qualificados** seguem a mesma leitura de eventos, cada um com sua lista de etapas
  em `_Config` (`KOMMO_ETAPAS_ELEGIVEL`, `KOMMO_ETAPAS_QUALIFICADO`): não é "a etapa é X", é
  "o lead passou do ponto X", contado na primeira entrada. Por isso quem está em
  *Sem resposta: QUALIFICAÇÃO* já conta como elegível, e quem foi para *Inviável* já conta como
  qualificado. Todo qualificado passou antes pela elegibilidade — qualificados acima de elegíveis
  é sinal de lista desalinhada.
- O histórico dessas duas abas é carregado **mês a mês**: rode
  **Elegíveis / Qualificados do Kommo (histórico)** no menu até o aviso dizer que terminou. Cada
  execução grava o que deu tempo e guarda onde parou.
- A aba **Mapeamento** vale também para o campo `Campanha` do Kommo: quando a automação do
  WhatsApp continua carimbando o nome de uma campanha antiga, aponte o nome velho para a que
  está no ar e as reuniões passam a cair na linha certa.

## O que é exibido

**Resultados das campanhas** — por período (mês atual por padrão; últimos 30 dias, últimos
3 meses, este ano, todo o período, mês escolhido ou período livre): nome, orçamento diário e mensal, gasto,
leads, custo por lead, **leads elegíveis e custo por elegível**, **qualificados**,
**reuniões marcadas, conversão lead → reunião e custo por reunião**, clientes,
retorno financeiro, conversão lead → cliente, CAC e ROI. O filtro **Campanhas: só as ativas** corta
antes do total, dos cartões e dos gráficos — não só da tabela — e o recorte diz quantas ficaram de fora.
Retorno = **honorários iniciais** dos fechamentos com **pagamento confirmado**, pela **data do
fechamento**, cujo campo ORIGEM casa com a campanha. Fórmulas: CPL = gasto ÷ leads ·
Conversão = clientes ÷ leads · CAC = gasto ÷ clientes · ROI = retorno ÷ gasto.

**Ordenação.** Toda coluna das tabelas de campanhas, de mídia, de anúncios e de mapeamento
ordena: clique no título, clique de novo para inverter. Número começa do maior, texto começa
do A, e linha sem valor (`—`) fica sempre no fim, nos dois sentidos. A coluna escolhida é
guardada no navegador, por tabela, e sobrevive à atualização automática. O total do rodapé,
os cartões e os gráficos não dependem da ordem — o gráfico "gasto × retorno" continua
mostrando as 8 que mais gastam. Nos anúncios a ordenação vem **antes** do corte em 200
linhas: ordenar por CPL traz os 200 daquele CPL, não os 200 que mais gastam.

**Visão do gestor de tráfego** — por janela fechada: entrega do orçamento, CPM, CTR do link,
página vista ÷ clique, CPL, frequência e alcance por campanha; anúncio a anúncio com
diagnóstico; e indicações geradas por regras (fadiga de criativo, CTR baixo, CPM subindo,
CPL piorando, cliques que não viram página vista, entrega abaixo do orçamento, gasto sem lead,
dependência de um único anúncio).

## Estrutura

```
index.html              painel (HTML + CSS + JS, sem build)
apps-script/Codigo.gs   coletor Meta → planilha (colar no Apps Script da planilha)
apps-script/Kommo.gs    coletor do Kommo → abas Reuniões, Elegíveis e Qualificados
logo.png / favicon.png
```
