🧾 Reconciliação Financeira Automatizada — Excel + VBA

Macro em VBA que cruza automaticamente os pagamentos feitos (extrato bancário) com as notas fiscais do contas a pagar, aponta o que bate, o que diverge e o que ficou sem par, e reúne tudo numa aba de relatório.

⚠️ Projeto educacional. Os dados são fictícios. A macro não substitui controles internos, revisão contábil nem testes em ambiente de homologação. Não use em produção sem testar com cópias dos arquivos e sem aprovação da área responsável.

Por que este projeto existe

Conferir manualmente, linha por linha, se cada pagamento tem uma nota fiscal correspondente (e vice-versa) é lento e sujeito a erro. Esta macro faz o cruzamento em segundos e deixa para a pessoa apenas os casos que realmente precisam de análise.

Como funciona, em resumo

A macro lê duas abas e compara as linhas entre elas:

Aba	O que contém
Extrato	Transações bancárias e pagamentos efetuados
Notas_Fiscais	Notas fiscais registradas no contas a pagar

Para cada pagamento, ela procura a nota correspondente comparando número da nota, fornecedor, moeda e valor. O resultado de cada cruzamento vai para uma nova aba, Reconciliação. As abas de origem não são alteradas.

Regras de cruzamento

A macro segue esta ordem:

Busca exata: procura uma nota com o mesmo número, fornecedor e moeda do pagamento.
Conferência do valor: se encontrou, compara os valores aceitando uma diferença de até 0,01 (para absorver arredondamentos).
Busca alternativa: se não achou a nota exata, procura uma única nota com mesmo fornecedor, moeda e valor. Quando encontra, não declara conciliado: marca para revisão, porque o número da nota não confere.
Situações especiais: sinaliza pagamentos estornados, notas estornadas ou pendentes, pagamentos sem nota e notas sem pagamento.
Sem uso duplicado: uma mesma nota nunca é associada a mais de um pagamento.
Status do resultado
Status	Significado	O que fazer
CONCILIADO	Nota, fornecedor, moeda e valor conferem	Nada
DIVERGÊNCIA DE VALOR	Nota, fornecedor e moeda conferem, mas o valor não	Verificar o valor pago e o valor da nota
REVISAR: NOTA DIFERENTE	Fornecedor, moeda e valor batem, mas o número da nota é outro	Confirmar se é a mesma operação lançada com outro número
REVISAR: MÚLTIPLAS CANDIDATAS	Mais de uma nota poderia corresponder ao pagamento	Escolher manualmente a nota certa
SEM NOTA CORRESPONDENTE	O pagamento não encontrou nenhuma nota compatível	Localizar a nota ou investigar o pagamento
SEM PAGAMENTO NO EXTRATO	A nota não foi associada a nenhum pagamento	Verificar se a nota está em aberto
Status de revisão de estornos e pendências	Pagamento estornado, ou nota estornada ou pendente	Confirmar a situação com a área responsável
Como executar no Excel
Abra Reconciliação_Financeira_Exemplo.xlsx no Excel para desktop.
Salve uma cópia como Pasta de Trabalho Habilitada para Macro do Excel (.xlsm).
Pressione Alt + F11 para abrir o Editor do Visual Basic.
No menu do editor, clique em File > Import File... e selecione vba/modReconciliacao.bas.
Confira se as abas se chamam exatamente Extrato e Notas_Fiscais.
Volte ao Excel e pressione Alt + F8.
Selecione ReconciliarPagamentos e clique em Executar.
Abra a aba Reconciliação, filtre a coluna Resultado e analise os itens que não estão como CONCILIADO.

🔒 Dependendo das configurações de segurança da empresa, pode ser preciso habilitar macros apenas para arquivos confiáveis. Nunca habilite macros de fontes desconhecidas.

Colunas esperadas

Os nomes dos cabeçalhos precisam ser exatamente estes:

Aba Extrato	Aba Notas_Fiscais
ID_Transacao	Nota_Fiscal
Data_Pagamento	Data_Emissao
Fornecedor	Fornecedor
Nota_Fiscal	Valor_Nota
Valor	Moeda
Moeda	Referencia_Servico
Status_Banco	Status_Nota
Estrutura do projeto
reconciliacao_financeira_vba/
├── Reconciliação_Financeira_Exemplo.xlsx   # pasta de trabalho com as abas de exemplo
├── dados/
│   ├── extrato.csv                         # pagamentos fictícios
│   ├── notas_fiscais.csv                   # notas fiscais fictícias
│   └── reconciliacao_exemplo.csv           # exemplo do resultado gerado
├── vba/
│   └── modReconciliacao.bas                # código da macro
└── README.md

Limitações conhecidas
A macro segue regras fixas. Ela aponta onde os dados batem ou não, mas não decide sozinha qual é a situação contábil correta de uma transação.
A busca alternativa (por fornecedor, moeda e valor) sempre gera um status de revisão, nunca uma conciliação automática.
Não trata automaticamente pagamentos parciais, uma transferência que paga várias notas, retenções de impostos, taxas bancárias, variação cambial, adiantamentos ou pagamentos em lote.
Os valores são convertidos com a função Val e uma troca simples do separador decimal (vírgula por ponto). Antes de adaptar para arquivos reais, confira o formato dos números e as configurações regionais do Excel.
A macro presume que os nomes das abas e dos cabeçalhos sejam exatamente os listados acima. Qualquer diferença (acento, espaço, maiúscula) pode impedir a execução.


Antonio Aguiar
