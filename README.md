# Agenda de Entregas Futuras · Armazém Santa Fé

Controle das vendas com faturamento antecipado (CFOP 5922) e das notas de entrega (CFOP 5117), com prazo de 120 dias, relatórios, conferência com o relatório do sistema e acesso por usuário e senha.

- **Site:** este repositório (GitHub Pages). Não contém nenhum dado de cliente.
- **Dados:** Google Drive do administrador, pasta **Agenda de Entregas – Armazém Santa Fé**
  - planilha **Agenda de Entregas – Banco** (abas `usuarios`, `vendas`, `sessoes`, `log`)
  - pasta **XMLs** (arquivos das notas enviados pela agenda)
- **Servidor:** `apps-script/Code.gs`, instalado como App da Web dentro da planilha.

## Perfis
| Perfil | O que faz |
|---|---|
| Administrador | tudo, inclusive criar usuários, trocar senhas, desativar acessos e importar a agenda antiga |
| Operador | registra vendas e entregas, edita e consulta |
| Consulta | só visualiza agenda, relatórios e conferência |

As senhas ficam guardadas com hash e sal (nunca em texto). Depois de 5 senhas erradas o usuário espera 10 minutos. A sessão dura 12 horas.

## Instalação (uma vez)
1. Abra a planilha **Agenda de Entregas – Banco** no Drive → **Extensões → Apps Script**.
2. Apague o conteúdo do arquivo `Código.gs`, cole o conteúdo de [`apps-script/Code.gs`](apps-script/Code.gs) e salve.
3. No menu de funções escolha **instalar** e clique **Executar** → autorize com a sua conta Google.
4. **Implantar → Nova implantação → tipo App da Web** · Executar como: **Eu** · Quem pode acessar: **Qualquer pessoa** → **Implantar** → copie o link que termina em `/exec`.
5. Cole o link em `config.js` (ou envie para quem cuida do repositório).
6. Abra o site, crie o **administrador** no primeiro acesso e, na aba **Administrador**, crie os outros usuários e importe a agenda antiga.

## Atualizar o servidor
Se `Code.gs` mudar: cole o novo código no Apps Script → **Implantar → Gerenciar implantações → editar (lápis) → Versão: Nova versão → Implantar**. O link `/exec` continua o mesmo.
