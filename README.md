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
| _Config | contas, data de início, tipos de ação que contam como lead |
| _Sync | estado da última sincronização, mostrado no topo do painel |

O token do Meta **não fica no código**: vive em *Configurações do projeto → Propriedades
do script → `META_ACCESS_TOKEN`*.

## O que é exibido

**Resultados das campanhas** — por período (mês atual por padrão; últimos 30 dias, últimos
3 meses, este ano, todo o período, mês escolhido ou período livre): nome, orçamento diário e mensal, gasto,
leads, custo por lead, clientes, retorno financeiro, conversão lead → cliente, CAC e ROI.
Retorno = **honorários iniciais** dos fechamentos com **pagamento confirmado**, pela **data do
fechamento**, cujo campo ORIGEM casa com a campanha. Fórmulas: CPL = gasto ÷ leads ·
Conversão = clientes ÷ leads · CAC = gasto ÷ clientes · ROI = retorno ÷ gasto.

**Visão do gestor de tráfego** — por janela fechada: entrega do orçamento, CPM, CTR do link,
página vista ÷ clique, CPL, frequência e alcance por campanha; anúncio a anúncio com
diagnóstico; e indicações geradas por regras (fadiga de criativo, CTR baixo, CPM subindo,
CPL piorando, cliques que não viram página vista, entrega abaixo do orçamento, gasto sem lead,
dependência de um único anúncio).

## Estrutura

```
index.html              painel (HTML + CSS + JS, sem build)
apps-script/Codigo.gs   coletor Meta → planilha (colar no Apps Script da planilha)
logo.png / favicon.png
```
