# Relatórios Veritatem: instruções completas

Relatório padronizado para os clientes da Veritatem. O conteúdo (métricas, funil, gráficos,
tabelas de campanhas, anúncios, públicos e regiões) tem como base o que o Reportei entregava,
mas o visual é próprio da Veritatem: capa grafite com o símbolo da marca, acento bronze,
fontes Outfit e Manrope (em `assets/fontes`, já embutidas no PDF) e seções numeradas.
Não altere cores, fontes ou estrutura por conta própria; o padrão é igual para todos os clientes. **Somente números e gráficos: não escrever análise,
comentários ou plano de ação no PDF.**

## Três modelos

| Modelo | Quando | Arquivos |
|---|---|---|
| Semanal Meta Ads | cliente só com Meta | exports do Meta |
| Semanal Meta + Google | cliente com as duas plataformas | exports do Meta e do Google |
| Mensal | fechamento do mês | mesmos exports, período do mês, `--tipo mensal` |

A seção de cada plataforma aparece automaticamente quando há arquivo dela. No relatório
combinado, o Google vem primeiro e o Meta depois (como no Reportei).

## Modo conectado (padrão): sem exportar nada

Quando a conexão do Meta Ads (`MCP Meta`) estiver disponível, o pedido pode ser só
"relatório <cliente> de <data> a <data>". Você puxa tudo pela conexão:

1. **Conta:** `ads_get_ad_accounts`; use o `ad_account_id` cujo nome bate com o cliente
   (se houver mais de uma conta parecida, pergunte qual). Se a conta vier com
   `is_ads_mcp_enabled: false`, peça o export do Gerenciador e siga o fluxo de arquivos.
2. **Dados:** chame `ads_get_ad_entities` com `time_range` `{"since":"AAAA-MM-DD","until":"AAAA-MM-DD"}`
   e `include_additional_context: false`. Salve cada resposta, do jeito que veio, em
   `<pasta>/<nome>.json` (se vier `next_cursor`, busque as outras páginas e junte as listas):
   - `conta`: level `ad_account`, fields `amount_spent, reach, impressions, frequency, clicks, link_click, page_engagement, post_reaction, post_save, post_shares, comment, video_thruplay_watched_actions, lead, landing_page_view, omni_purchase`
   - `campanhas`: level `campaign`, fields `name, results` + os mesmos de `conta`; filtering `amount_spent GREATER_THAN 0`
   - `anuncios`: level `ad`, fields `name, adset_name, creative_id, amount_spent, results, reach, impressions, frequency, clicks, link_click`; mesmo filtro; sort `amount_spent_descending`
   - `dia`: level `ad_account`, fields `amount_spent, impressions, reach, clicks, link_click`, `time_increment: "1"`
   - `idade`, `genero`, `regiao`: level `ad_account`, breakdowns `["age"]`, `["gender"]`, `["region"]`, fields `amount_spent, impressions, reach, clicks`
   - `dispositivo`, `hora`: breakdowns `["device_platform"]` e `["hourly_stats_aggregated_by_audience_time_zone"]`, fields `amount_spent, impressions, reach`
3. **Converter:** `python3 scripts/meta_conexao.py --entrada <pasta> --saida <pasta_csv>`.
   A conta.json vira a linha de resumo, então alcance e frequência saem exatos.
4. **Prévias dos anúncios (padrão):** `ads_get_creatives` com `creative_ids` (os `creative_id` dos
   anúncios que entram no relatório) e fields `id, object_type, image_url, image_hash, video_id,
   thumbnail_url, child_attachments`. Para imagem grande: o `image_hash` do criativo ou do primeiro item
   de `child_attachments` (carrossel) em `ads_get_ad_images` (fields `hash, url`). Vídeos ficam com
   `thumbnail_url` (miniatura). Monte um JSON `{creative_id: {"url": ..., "grande": true|false}}`.
   As imagens ficam nos servidores do Facebook, que este ambiente não baixa; por isso o relatório
   com prévias sai como **HTML para abrir no navegador** (o navegador carrega as imagens) e salvar
   como PDF: rode o gerador com `--previas-links <json>`. Diga à pessoa: abrir no Chrome, Ctrl+P,
   "Salvar como PDF". Os links das imagens valem por alguns dias; peça para salvar o PDF logo.
5. **Google Ads (manual, por enquanto):** o Google **não** é puxado automaticamente. Se o cliente tem
   Google Ads (relatório Meta + Google), já no início peça à pessoa os exports do Google do mesmo
   período, com esta mensagem (ajuste as datas):

   > Para a parte do Google Ads, me manda os exports de **dd/mm a dd/mm** (CSV ou Excel), com as colunas
   > Custo, Impr., Cliques, Conversões, Todas as conv., % de impr. na parte superior e % de impr. no topo absoluto:
   > 1. **Campanhas** (obrigatório)
   > 2. **Grupos de anúncios**
   > 3. **Palavras-chave**
   > 4. **Campanhas segmentado por Dia**
   >
   > No Google Ads: abra cada tela, ajuste o período, clique no ícone de download e escolha CSV ou Excel.

   Enquanto espera, já puxe e prepare o Meta. Com os arquivos do Google em mãos, passe-os junto com os
   CSVs do Meta em `--arquivos`. Dispositivo do Google só entra se a pessoa mandar.
   *(Futuro, ainda desativado: o script em `google-ads-script/` grava planilhas "Veritatem Google Ads - <conta>"
   no Drive, convertidas por `scripts/google_planilha.py`. Só use se a pessoa disser que o script está ativo.)*
6. **Gerar:** `gerar_relatorio.py` com os CSVs do Meta e do Google em `--arquivos` e `--previas-links`. Sem prévias (ou se a pessoa pedir PDF direto), omita `--previas-links` e o PDF sai pronto.

## Mensagem para o cliente (WhatsApp)

Todo relatório gera também a mensagem padrão da Veritatem: vem no campo `mensagem` do JSON impresso
pelo gerador (e salva em `<nome do relatório>_mensagem.txt`). Depois de entregar o relatório, **cole a
mensagem na resposta, dentro de um bloco de código**, exatamente como veio, para a pessoa copiar e
mandar no grupo do cliente. Não reescreva nem acrescente números: ela usa os mesmos dados do relatório.
Formato: saudação conforme a hora, período, bloco do Google (impressões, cliques, conversões e valor
investido) e bloco do Meta (impressões, alcance, cliques e valor investido, mais conversas/leads/compras
quando a conta tiver). Negrito com `*` e linhas com `>` (citação do WhatsApp).

## Fluxo com arquivos exportados

1. Confirme três informações (pergunte só o que não estiver claro na mensagem ou nos arquivos):
   nome do cliente, período (dd/mm/aaaa a dd/mm/aaaa) e se é semanal ou mensal.
   As colunas "Início/Término dos relatórios" do export do Meta costumam trazer o período.
2. Rode o gerador com todos os arquivos enviados (ele identifica cada um pelas colunas):

```bash
pip install pandas openpyxl matplotlib playwright --break-system-packages -q 2>/dev/null
python3 scripts/gerar_relatorio.py \
  --cliente "Nome do Cliente" --tipo semanal \
  --inicio 02/10/2026 --fim 08/10/2026 \
  --arquivos /caminho/arquivo1.csv /caminho/arquivo2.xlsx ... \
  --saida /mnt/user-data/outputs
```

   (Use o caminho completo de `scripts/gerar_relatorio.py` dentro da pasta desta skill.)
   Se o Chromium do Playwright não estiver instalado, rode `playwright install chromium`;
   se não der, o script usa WeasyPrint (`pip install weasyprint`).

3. Leia o JSON impresso no final:
   - `secoes`: quais arquivos foram reconhecidos. Se algum arquivo enviado não apareceu, avise.
   - `avisos`: repasse cada aviso para quem pediu, em linguagem simples.
4. Converta a primeira parte do PDF em imagem e confira visualmente antes de entregar
   (`pdftoppm -r 60 -png arquivo.pdf previa`).
5. Entregue o PDF com uma linha de contexto. Nome do arquivo:
   `Veritatem_<Cliente>_<Semanal|Mensal>_<inicio>_a_<fim>.pdf` (o script já gera assim).

## Regras por cliente

`references/clientes.json` guarda ajustes fixos de cada cliente (o gerador aplica sozinho pelo nome do
cliente). Hoje: **Armenza** sem alcance por dispositivo, sem impressões/alcance por horário e sem regiões
no Meta. Quando alguém pedir um ajuste "como padrão" para um cliente, edite esse arquivo (e, se tiver acesso,
envie para o repositório). Para esconder algo só uma vez: `--ocultar meta:hora,meta:regiao`.

## Regras

- **Nunca inventar ou estimar números.** Tudo vem dos exports. Se faltar algo, a seção fica de fora
  e você avisa o que faltou.
- Não escrever análise nem texto opinativo no PDF. O único texto fixo é o glossário de métricas.
- Comparativo com período anterior **não** faz parte do padrão. Se pedirem, peça os exports do
  período anterior e entregue à parte uma tabela (métrica, período anterior, período atual, variação %).
- Os cards de destaque do Meta mostram automaticamente os resultados principais da conta (conversas,
  leads, compras, visitas ao perfil, visualizações da página). Para forçar, use
  `--destaques conversas,leads` (opções: conversas, leads, compras, visitas_perfil, lpv).
- Logo: só a da Veritatem (pasta `assets/`). Não usar logo do cliente.
- Se o export de campanhas do Meta vier sem a linha de resumo, o alcance total é aproximado;
  o script avisa. Se houver conexão com o Meta Ads disponível, puxe o alcance da conta no período
  (nível conta, campo `reach`) e passe `--alcance-total <número>`. Senão, peça para reexportar
  marcando "incluir linha de resumo".
- Exports em inglês ou português, CSV ou XLSX, com vírgula ou ponto decimal: todos funcionam.
- Performance Max não tem "% de impressões na parte superior"; o card vira "Custo por conversão".
- Se a pessoa mandar prints em vez de arquivos, explique que o padrão precisa dos exports
  (guia em `references/exportacao.md`) e mande o resumo do que exportar.

## Como exportar

Ver `references/exportacao.md`. Mande esse guia quando alguém perguntar o que exportar.
