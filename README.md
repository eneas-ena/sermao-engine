# Infográficos de Sermões - Vercel

Aplicativo pessoal para analisar sermões em DOCX e gerar:

- infográfico PNG;
- PDF 16:9;
- PowerPoint editável;
- projeto JSON reutilizável.

## Publicar no GitHub

1. Crie um repositório privado no GitHub.
2. Extraia o ZIP e envie todo o conteúdo desta pasta para o repositório.
3. Confirme que `package.json`, `requirements.txt` e `vercel.json` estão na raiz.

## Publicar no Vercel

1. No painel do Vercel, escolha `Add New > Project`.
2. Importe o repositório criado no GitHub.
3. Em `Framework Preset`, use `Other` se o Vercel não escolher automaticamente.
4. Não informe Build Command nem Output Directory.
5. Clique em `Deploy`.

O aplicativo não usa banco de dados nem exige variáveis de ambiente. Os arquivos
são gerados somente para download e não ficam armazenados no servidor.

## Observação sobre arquivos grandes

O Vercel limita o corpo de cada requisição. O aplicativo reduz automaticamente a
imagem de capa antes do envio. Sermões em DOCX normalmente ficam bem abaixo do limite.
