# Relatórios Veritatem

Código central dos relatórios de tráfego pago da Veritatem (Meta Ads e Google Ads).
A skill "veritatem-relatorios" instalada no Claude baixa este repositório a cada relatório,
então qualquer alteração feita aqui vale na hora para a equipe inteira, sem reinstalar nada.

- `INSTRUCOES.md`: passo a passo que o Claude segue para gerar o relatório
- `scripts/`: geração do PDF/HTML e conversão dos dados do Meta e do Google
- `assets/`: logos e fontes da Veritatem
- `google-ads-script/`: script do Google Ads que alimenta as planilhas no Drive
- `references/`: o que exportar quando a conta não tem conexão direta
