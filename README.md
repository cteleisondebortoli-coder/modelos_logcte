/-/ Automação para extrair em Excel todos os modelos de uma transportadora por CNPJ

Este repositório contém uma automação na nuvem (GitHub Actions) projetada para extrair dados detalhados de Modelos CTe da plataforma LogCTe.

O robô acessa o sistema web da LogCTe, pesquisa por um CNPJ específico, lista todos os modelos ativos/inativos encontrados e acessa cada um para extrair mais de 25 campos de configuração e regras. No fim, exporta tudo em um arquivo .xlsx para download.

---

/-/ 🔐 Configuração de Usuário e Senha pelo Secrets:
Para que o bot faça login na LogCTe sem expor sua senha para o público no código, você deve registrar as senhas no "Cofre" do repositório:

1. Na página do seu repositório, clique na aba **Settings** (Configurações).
2. No menu esquerdo, vá em **Secrets and variables** e depois em **Actions**.
3. Crie dois "New repository secrets":
   * **Nome:** `LOGCTE_EMAIL` | **Secret:** `(seu-email-da-logcte)`
   * **Nome:** `LOGCTE_PASSWORD` | **Secret:** `(sua-senha-da-logcte)`

---

/-/ 🚀 Como Executar
1. Navegue até a aba **Actions** na parte superior do repositório.
2. Na barra esquerda, clique em **Extração LogCTe**.
3. Do lado direito, clique no botão cinza **Run workflow**.
4. Irá aparecer um campo solicitando o **CNPJ**. Cole o CNPJ desejado, apenas números.
5. Clique no botão verde **Run workflow**.

---

/-/ 📊 Como Baixar a Planilha Excel
1. Após a execução terminar com um ✅ verde na aba **Actions**, clique nela para abrir o resumo.
2. Desça até a parte inferior da página, na seção **Artifacts**.
3. Clique no arquivo gerado (ex: `Resultado-LogCTe-[CNPJ].zip`).
4. Ele baixará um ZIP que contém a sua planilha `.xlsx` com todos os dados preenchidos e links prontos para acesso direto.
