# Gestão Comercial (PWA)

App de gestão para pequenos comércios. Funciona no aparelho do dono, offline, sem servidor, login ou mensalidade.

## Como entregar a um comerciante
1. Publique esta pasta em qualquer hospedagem estática com HTTPS (GitHub Pages, Netlify, Cloudflare Pages: todas têm plano gratuito).
2. O comerciante abre o endereço no celular ou computador, usa "Instalar app" / "Adicionar à tela inicial" e faz a configuração inicial.
3. Cada aparelho tem seus próprios dados. Nada é enviado a lugar nenhum.

Para testar localmente: `python3 -m http.server 8000` e abra `http://localhost:8000`.

## Cuidados com os dados (explique ao comerciante)
- Os dados ficam só no aparelho. Limpar os dados do navegador, desinstalar o app ou trocar de celular APAGA tudo.
- Faça o backup em Planilha e backup > Salvar backup, de preferência toda semana, e guarde o arquivo no Google Drive da empresa.
- Para recuperar (ou passar para outro aparelho): Restaurar backup.
- A planilha .xlsx serve para consultar e organizar as vendas. Ela NÃO restaura o app; só o arquivo de backup restaura.

## Atualizar o app
Depois de mudar qualquer arquivo, aumente o número da versão `V` em `sw.js` (ex.: gc-v3 para gc-v4), senão os aparelhos continuam com a versão antiga em cache.

## Arquivos
index.html, app.js (lógica), styles.css, sw.js (offline), manifest.webmanifest e icon.svg (padrões), xlsx.min.js (SheetJS, gera o Excel).
