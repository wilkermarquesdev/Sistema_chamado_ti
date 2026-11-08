
# 🛠️ Sistema de Gestão de Chamados, Estoque e Compras

Sistema web integrado focado na gestão ágil de ordens de serviço (O.S.) de TI, controle inteligente de inventário e fluxo automatizado de reposição de suprimentos.

---

## 🚀 Tecnologias Utilizadas

Este projeto foi construído utilizando uma stack leve, moderna e de baixo custo de manutenção:

* **Back-end & Lógica:** Google Apps Script (`JavaScript`)
* **Front-end:** HTML5, CSS3, JavaScript e Tailwind CSS
* **Banco de Dados / Armazenamento:** Google Sheets (Planilhas Google)
* **Comunicações & IA:** 
  * Integração com WhatsApp via **WAHA** para notificações e fluxo ágil.
  * Inteligência Artificial **Google Gemini (Gemini 2.5 Flash)** para automações e análises.

---

## 📁 Estrutura do Projeto

* `appsscript.json` - Configurações e manifesto do projeto no Apps Script.
* `Código.gs` - Módulo central contendo as rotas (`doGet`), manipulação de planilhas e formatações globais.
* `Index.html` - Interface principal do usuário (Front-end).
* `WAHA.gs` - Módulo de integração e envio de mensagens via WhatsApp.
* `Inventario.gs` - Módulo de controle de estoque, alertas de ruptura e gestão de suprimentos.

---

## ⚙️ Funcionalidades Principais

1. **Abertura e Acompanhamento de Chamados:** Registro centralizado de ocorrências de TI com controle de status e histórico de atendimento.
2. **Controle de Estoque Dinâmico:** Monitoramento de peças e insumos com verificação automatizada de níveis críticos.
3. **Notificações Automatizadas:** Envio de alertas de movimentações e chamados via WhatsApp.
4. **Assistência Inteligente:** Uso integrado de IA para otimização de processos e triagem de demandas.

---

## 👤 Autor

Desenvolvido por **Wilker José**  
* 📧 E-mail: wilker.marques.dev@gmail.com  
* 📱 WhatsApp: (62) 99421-5080
