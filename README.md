# Portal de Solicitação de Dados do CAF

Portal completo com dicionário PF/PJ, justificativas obrigatórias, formulário, termos individuais, máscara de telefone e Excel.

## Publicar

1. Crie um repositório público chamado `formulario` na conta `caf-cod`.
2. Envie o conteúdo desta pasta para a raiz, incluindo `.github/workflows/pages.yml`.
3. Em Settings > Pages > Build and deployment > Source, selecione **GitHub Actions**.
4. Aguarde o fluxo **Publicar portal no GitHub Pages** concluir na aba Actions.

Endereço previsto: https://caf-cod.github.io/formulario/.

## Executar localmente

Requer Node.js 22.13 ou superior.

```sh
npm install
npm run dev
```

`app/page.tsx` contém as telas e o e-mail de destino (`destinationEmail`). `app/data/caf-data.json` contém o dicionário. `vite.config.ts` configura o caminho `/formulario/`.

Os dados são processados no navegador. Os PDFs assinados e a planilha devem ser anexados manualmente ao e-mail.
