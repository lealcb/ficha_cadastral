# Ficha Cadastral - Casas Comigo

Aplicativo estático para o cliente preencher a ficha cadastral de locação e baixar a versão preenchida em Word (DOCX) e PDF.

## Privacidade

Os dados são processados apenas no navegador do usuário. Não há banco de dados, servidor de aplicação, cookies ou ferramenta de análise.

## Regras do formulário

- RG do locatário é opcional.
- A seção de cônjuge só é habilitada quando marcada.
- RG do cônjuge também é opcional.
- Os demais dados cadastrais principais são obrigatórios.

## Desenvolvimento local

```bash
npm install
npm run dev
```

## Publicação

O workflow em `.github/workflows/deploy-pages.yml` publica o site no GitHub Pages a cada push na branch `main`.
