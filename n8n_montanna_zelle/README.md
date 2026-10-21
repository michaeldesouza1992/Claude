# Montanna Zelle – Webhook de SMS (fase de teste)

Objetivo: receber no n8n o texto dos SMS de Zelle enviados pelo iPhone.
O Monday **ainda não** é atualizado nesta fase.

## 1. Importar o workflow no n8n
1. n8n → **Create Workflow** → menu `⋯` → **Import from File**.
2. Escolha `montanna-zelle-webhook.json`.
3. Abra o nó **Webhook** e confira:
   - HTTP Method: `POST`
   - Path: `montanna-zelle`
   - Authentication: `Header Auth`
   - Respond: `Immediately`
4. Em **Credential for Header Auth** → **Create New**:
   - Name: `X-Montanna-Token`
   - Value: um token longo e aleatório (guarde-o; não compartilhe).
5. **Save**. O nó mostra a **Test URL** e a **Production URL**.

## 2. Testar sem o iPhone (opcional)
Clique **Listen for test event** no nó e, em outro dispositivo:

```bash
curl -X POST "<TEST_URL>" \
  -H "X-Montanna-Token: <SEU_TOKEN>" \
  -H "Content-Type: application/json" \
  -d '{"sender":"Zelle","message":"You received $100.00 from JOHN DOE"}'
```

A Test URL só funciona enquanto o n8n está em "Listen for test event".

## 3. Atalho no iPhone (app Atalhos)
1. **Automação** → **+** → **Mensagem**.
2. "Mensagem contém": `Zelle` (ou o remetente do banco). Marque **Executar imediatamente**.
3. Ação **Obter Conteúdo do URL**:
   - URL: Test URL (depois troque pela Production URL)
   - Método: `POST`
   - Cabeçalhos: `X-Montanna-Token` = seu token
   - Corpo da requisição: `JSON`
     - `message` → variável **Mensagem** (conteúdo)
     - `sender` → variável **Remetente**
4. Salve. Com o n8n ouvindo, peça um Zelle de teste (ou envie a si mesmo um SMS com "Zelle").

## 4. Próximo passo (depois do teste)
Ativar o workflow (usar Production URL) e adicionar nós para extrair
valor/nome e atualizar o Monday.
