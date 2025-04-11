# Planilha para Cálculo do Simples Nacional com Importação de Notas Fiscais

## Descrição

Esta planilha foi desenvolvida para automatizar o cálculo do Simples Nacional, com funcionalidades para importar dados de **Notas Fiscais Eletrônicas (DANFE)** e **Notas Fiscais de Serviços Eletrônicas (NFSe)** da Prefeitura de Belo Horizonte. O processo utiliza macros VBA para extrair informações relevantes dos arquivos XML, consolidar os dados e preparar relatórios necessários para o cálculo do Simples Nacional.

---

## Funcionalidades

1. **Importação de Notas Fiscais**:
   - Suporte para arquivos XML de DANFE (NF-e) e NFSe.
   - Identificação automática do tipo de nota fiscal.
   - Extração de dados como número da nota, data de emissão, CNPJ da prestadora, CFOP, valor total, ISS retido, e status de cancelamento.

2. **Consolidação de Dados**:
   - Relatório por filial (CNPJ) e mês.
   - Totais de faturamento bruto, devoluções, retenções de ISS, e faturamento líquido.
   - Identificação da primeira e última nota fiscal do período.
   - Gestão de notas fiscais faltantes e canceladas.

3. **Cálculo Automático do Simples Nacional**:
   - Integração com a planilha "Calc_Simples" para calcular os tributos mensais baseados nos dados consolidados.

---

## Pré-requisitos

1. Microsoft Excel com suporte a macros habilitado.
2. Arquivos XML das notas fiscais:
   - DANFE (NF-e): Deve seguir o padrão da Secretaria da Fazenda.
   - NFSe: Deve seguir o padrão da Prefeitura de Belo Horizonte.

3. Planilhas com tabelas estruturadas:
   - **"Lancamentos"**: Para armazenar as notas fiscais importadas.
   - **"Resumo"**: Para consolidar dados por filial.
   - **"Calc_Simples"**: Para o cálculo mensal do Simples Nacional.

---

## Como Usar

1. **Importar Notas Fiscais**:
   - Abra a planilha no Excel.
   - Vá até a aba "Lancamentos".
   - Clique no botão "Processar Notas Fiscais" para selecionar os arquivos XML.
   - Revise os dados importados na tabela "TabelaDados".

2. **Consolidar Dados**:
   - Na aba "Resumo", clique no botão "Consolidar Dados".
   - Revise os relatórios gerados por filial e por mês.

3. **Calcular o Simples Nacional**:
   - Verifique os dados na aba "Calc_Simples".
   - Siga os passos do cálculo, ajustando conforme necessário.

---

## Estrutura do Código VBA

O projeto utiliza duas macros principais:

### 1. **ProcessarMultiplasNotasFiscaisParaTabela**

- Importa e processa múltiplos arquivos XML.
- Identifica o tipo de nota fiscal (DANFE ou NFSe).
- Adiciona os dados relevantes à tabela "TabelaDados".

### 2. **ConsolidarDados**

- Processa os dados importados para gerar relatórios consolidados.
- Atualiza informações de faturamento bruto, líquido, retenções e devoluções.
- Organiza os dados por filial e por mês.

### Funções Auxiliares

- `AjustarFormatoData`: Converte datas do formato ISO para o formato regional.
- Validações de CFOP para diferenciar devoluções e vendas.

---

## Observações

- **Configuração da Planilha**: Certifique-se de que as tabelas estruturadas estão corretamente nomeadas:
  - Tabela de origem: `TabelaDados`
  - Tabela de consolidação por filial: `TabelaConsolidada`
  - Tabela de consolidação mensal: `TabelaMensal`

- **Notas Canceladas**: São marcadas automaticamente e excluídas dos cálculos financeiros.

- **Formatos de Data**: As datas são formatadas para o padrão regional (DD/MM/AAAA).

---

## Atualizações Futuras

- Suporte para outras cidades além de Belo Horizonte.
- Melhorias na interface do usuário para maior usabilidade.
- Integração com sistemas fiscais online.

---
