# Controle de Atividades (PWA)

App instalável para registrar FEZ / NÃO FEZ por turma, com a planilha Google Sheets como backend.

- Raiz: o PWA (publicado na Vercel como site estático).
- `apps-script/`: backend (Code.gs). Implante como App da Web, "Executar como: Eu", acesso "Qualquer pessoa".
  Troque `TOKEN_API` por uma senha sua. No primeiro uso o PWA pede o endereço `/exec` e o token.
