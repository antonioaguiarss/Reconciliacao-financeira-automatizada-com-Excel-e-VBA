# Reconciliação Financeira Automatizada — Excel + VBA

Projeto de portfólio que automatiza o cruzamento de duas fontes fictícias:
- **Extrato**: transações bancárias/pagamentos efetuados.
- **Notas_Fiscais**: registro de notas fiscais do contas a pagar.

A macro compara nota fiscal, fornecedor, moeda e valor; sinaliza divergências e cria a aba `Reconciliação` com o resultado de cada cruzamento.

> **Importante:** dados fictícios para demonstração. A macro é um exemplo educacional, não substitui controles internos, revisão contábil ou validação em ambiente de homologação. Não a utilize em produção sem testes com cópias dos arquivos e aprovação da área responsável.

## Estrutura

```text
reconciliacao_financeira_vba/
├── Reconciliação_Financeira_Exemplo.xlsx
├── dados/
│   ├── extrato.csv
│   ├── notas_fiscais.csv
│   └── reconciliacao_exemplo.csv
├── vba/
│   └── modReconciliacao.bas
└── README.md
```

## Como executar no Excel

1. Abra `Reconciliação_Financeira_Exemplo.xlsx` no Excel desktop.
2. Salve uma cópia como **Pasta de Trabalho Habilitada para Macro do Excel (`.xlsm`)**.
3. Pressione `Alt + F11` para abrir o Editor do Visual Basic.
4. No menu do editor, escolha **File > Import File...** e selecione `vba/modReconciliacao.bas`.
5. Confirme que as abas se chamam exatamente `Extrato` e `Notas_Fiscais`.
6. Volte ao Excel e pressione `Alt + F8`.
7. Selecione `ReconciliarPagamentos` e clique em **Executar**.
8. Consulte a aba `Reconciliação`, filtre a coluna `Resultado` e investigue os itens divergentes.

Dependendo das configurações de segurança da empresa, pode ser necessário habilitar macros somente para arquivos confiáveis. Nunca habilite macros de fontes desconhecidas.

## Colunas esperadas

**Extrato**
- `ID_Transacao`
- `Data_Pagamento`
- `Fornecedor`
- `Nota_Fiscal`
- `Valor`
- `Moeda`
- `Status_Banco`

**Notas_Fiscais**
- `Nota_Fiscal`
- `Data_Emissao`
- `Fornecedor`
- `Valor_Nota`
- `Moeda`
- `Referencia_Servico`
- `Status_Nota`

## Regras da macro

1. Procura primeiro uma correspondência exata de número de nota + fornecedor + moeda.
2. Compara o valor com tolerância de R$ 0,01.
3. Se não encontrar a nota exata, tenta localizar uma única candidata por fornecedor + moeda + valor, marcando-a para revisão em vez de declarar conciliação.
4. Sinaliza pagamentos estornados, notas estornadas/pendentes, pagamentos sem nota e notas sem pagamento no extrato.
5. Evita associar a mesma linha de nota a mais de uma transação.
6. Gera uma aba de relatório e mantém as abas de origem sem alterações.

## Status possíveis

- `CONCILIADO`: nota, fornecedor, moeda e valor conferem.
- `DIVERGÊNCIA DE VALOR`: nota/fornecedor/moeda conferem, mas o valor não.
- `REVISAR: NOTA DIFERENTE`: fornecedor, moeda e valor batem, mas o número da nota difere.
- `REVISAR: MÚLTIPLAS CANDIDATAS`: mais de uma nota pode corresponder.
- `SEM NOTA CORRESPONDENTE`: transação não encontrou nota compatível.
- `SEM PAGAMENTO NO EXTRATO`: nota não foi associada a uma transação.
- Status de revisão para estornos e notas pendentes.

## Limitações conhecidas

- O algoritmo é baseado em regras e não identifica sozinho a verdade contábil de uma transação.
- A correspondência alternativa por fornecedor/moeda/valor é marcada para revisão.
- Não cobre automaticamente pagamentos parciais, agrupamento de várias notas em uma transferência, retenções tributárias, taxas bancárias, câmbio, adiantamentos ou liquidações em lote.
- A macro utiliza a função `Val` com substituição simples de separador decimal; valide os formatos numéricos e regionais antes de adaptar para arquivos reais.
- O procedimento presume cabeçalhos e nomes de abas conforme descritos neste README.

## História para entrevista

**Problema:** o cruzamento manual entre extrato e notas fiscais é repetitivo, sujeito a erro e difícil de auditar.

**Solução:** uma macro VBA percorre as transações, procura correspondências com regras explícitas, sinaliza divergências e gera um relatório filtrável.

**Valor de negócio:** reduz trabalho repetitivo, padroniza critérios de conferência e direciona a atenção da equipe para exceções — sem automatizar indevidamente a decisão final.

**Próximas melhorias:** botão na planilha, resumo de indicadores, log de execução, tolerâncias configuráveis, validação de múltiplas moedas e tratamento de pagamentos parciais.
