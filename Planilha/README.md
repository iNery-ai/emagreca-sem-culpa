# Manual da Planilha: Gestão de Tráfego Pago e Validação de Produtos Low Ticket

Esta planilha foi projetada especificamente para permitir um acompanhamento diário ágil da performance de produtos digitais *low ticket* (baixo custo), identificando gargalos no funil de vendas e apoiando tomadas de decisão rápidas de escala ou otimização.

---

## 🚀 Como Preencher Diariamente (Em menos de 2 minutos)

Para manter a planilha 100% atualizada, sua rotina operacional diária consiste em apenas **dois passos simples**:

1. **Passo 1: Registro de Vendas (Aba `Registro de Vendas`)**
   - Sempre que houver vendas, registre-as inserindo a **Data**, o **Valor da Venda** (ex: R$ 47,00), o **Produto**, o **Criativo Responsável** (selecione no dropdown) e a **Campanha**.
   - *Dica:* Você pode colar as linhas diretamente do relatório de vendas exportado da sua plataforma de pagamento (Hotmart, Kiwify, Monetizze, etc.).

2. **Passo 2: Controle Diário de Tráfego (Aba `Controle Diario`)**
   - Adicione uma nova linha com a data de hoje.
   - Acesse o gerenciador de anúncios (Meta Ads, Google Ads, TikTok Ads, etc.) e preencha as colunas manuais:
     - **Valor Investido**
     - **Impressões**
     - **Alcance**
     - **Cliques no Link**
     - **Visualizações da Página** (do pixel da Landing Page)
     - **Finalizações de Compra** (Checkout Iniciado)
   - As colunas de **Compras** e **Faturamento** nesta aba, assim como todas as métricas de desempenho (**CPC, CTR, CPA, ROAS, Conversão de Página e Checkout**), são calculadas **automaticamente** a partir do cruzamento de dados com a aba `Registro de Vendas`.
   - Adicione uma breve anotação na coluna **Observações** (ex: "Subiu orçamento", "Novo Criativo B", etc.) para registrar alterações importantes.

---

## 📊 Estrutura das 7 Abas

### 1. Dashboard Executivo (Aba Automática)
* Onde você acompanha a saúde macro do negócio.
* **Métricas Gerais:** Investimento total, Faturamento total, Lucro bruto, ROAS médio, CPA médio, Ticket médio, Conversão geral de funil e muito mais.
* **Metas de Performance:** No canto direito, você pode definir sua **Meta ROAS** (padrão: `2.0`) e **Meta CPA Máximo** (padrão: `R$ 20,00`).
* **Alertas Visuais:** O painel indica automaticamente se a operação está saudável:
  - 🟢 **OK:** Meta de ROAS atingida ou CPA abaixo da meta.
  - 🟡 **Atenção:** ROAS entre 1.0 e a Meta, ou CPA levemente acima do esperado.
  - 🔴 **Ação:** ROAS abaixo de 1.0 (prejuízo bruto) ou CPA crítico.
* **Gráficos Integrados:** Visualização da evolução de gastos vs. faturamento, quantidade diária de compras e tendência de CPA vs. ROAS diário.

### 2. Controle Diario (Aba Principal)
* Histórico diário consolidado.
* Contém regras de formatação condicional que destacam o ROAS do dia e sinalizam se o CPA diário está subindo ou acima da média histórica.
* *Painel Congelado:* O cabeçalho fica travado na parte superior para facilitar a navegação em históricos longos.

### 3. Controle de Criativos (Aba Semiacompanhada)
* Cadastro e análise de cada anúncio.
* **Ranking de Criativos (Tabela Auxiliar):** Identifica dinamicamente os **3 melhores criativos** da operação com base no ROAS e CPA individuais acumulados nas vendas.
* **Classificação Automática:** Classifica cada criativo como `Excelente`, `Bom`, `Regular` ou `Ruim`.
* **Dropdowns Dinâmicos:** A lista de criativos cadastrada aqui alimenta automaticamente o menu de seleção da aba `Registro de Vendas`.

### 4. Analise de Funil (Aba Automática)
* Análise detalhada de etapas de conversão e taxas de carregamento da página.
* **Diagnósticos Rápidos:** A coluna à direita avisa se há problemas específicos:
  - 🔴 **Criativo:** CTR abaixo de 1% (baixa atratividade dos anúncios).
  - 🔴 **Carregamento:** Perda superior a 30% entre os cliques no link e as visualizações da página (página de vendas pesada ou lenta).
  - 🔴 **Landing Page:** Conversão da página para o checkout inferior a 5% (promessa fraca ou design ruim).
  - 🔴 **Checkout/Oferta:** Conversão de checkout para compra inferior a 5% (oferta fraca, falta de checkout transparente ou Pix).

### 5. Historico de Escala (Aba Semiacompanhada)
* Registro de alterações de orçamento diário nas campanhas.
* Permite avaliar o impacto percentual do aumento de verba nas campanhas e analisar a resposta pós-escala (`Positivo`, `Neutro`, `Negativo`).
* Contém um gráfico de linhas do tempo mostrando a evolução do orçamento.

### 6. Registro de Vendas (Aba Principal)
* Registro bruto de transações para controle fino de ticket médio, faturamento por criativo e quantidade de vendas reais.

### 7. Insights Automaticos (Aba Automática)
* Diagnóstico completo gerado por fórmulas inteligentes que lêem toda a planilha e oferecem recomendações operacionais práticas em tempo real:
  - Análise da meta de CPA e ROAS.
  - Otimizações para taxa de checkout e de página.
  - Detecção automática de fadiga de criativos (CPA subindo por 3 dias seguidos).
  - Indicação em linguagem natural do Criativo Destaque atual da operação.

---

## 🛠️ Compatibilidade
A planilha foi construída com fórmulas padrão e formatação condicional nativa, sendo **totalmente compatível** com:
* Microsoft Excel (Desktop e Web)
* Google Planilhas (Google Sheets)
