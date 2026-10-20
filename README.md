# Romaneio Automático (VBA)

Macros em VBA que geram os romaneios de carga a partir de um arquivo de pedidos. Para cada pedido, cria uma aba a partir do modelo `template`, preenche carga, pedido e itens vindos do banco, pinta palavras-chave, calcula o peso das portas e confere os totais contra a origem.

## Módulos

| Módulo | Função |
|---|---|
| `modRomaneio.bas` | Ponto de entrada (`ProcessarPedidos`). Lê o arquivo de origem, extrai pedidos e cria uma aba por pedido |
| `modPosImport.bas` | Pós-processamento de cada aba: substituição, pintura, peso, conferência e limpeza |
| `modLucasBD.bas` | Copia itens da fonte para o romaneio e traduz siglas (acabamentos, larguras, espessuras) |
| `modEsperaBanco.bas` | Aguarda a fonte de dados (QueryTables/ListObjects) ficar pronta, com timeout de 20 s |
| `modColorirItens.bas` | Pinta termos e sequências nos itens |
| `modPesoPortas.bas` | Calcula o peso das portas por tipo (Doorblanks, Colmeia, Lisa) e tamanho |
| `modSequencia.bas` | Escreve a sequência e a cor correspondente em cada aba |
| `modConferenciaTotais.bas` | Compara o total da origem com o total da aba e monta o relatório de divergências |
| `modConversao.bas` | Conversão segura de números, textos e sequências |
| `modAjusteFonte.bas` | Reduz a fonte da célula do pedido até caber |
| `modAppState.bas` | Modo rápido: desliga atualização de tela e cálculo, com contador de profundidade |
| `modTestes.bas` | Testes manuais (`ExecutarTestes`) das conversões, substituições e cálculos |

## Fluxo de execução

```mermaid
flowchart TD
    A([ProcessarPedidos]) --> B[Selecionar arquivo de origem]
    B --> C{Arquivo já aberto?}
    C -->|Não| D[Abrir somente leitura]
    C -->|Sim| E[Usar workbook aberto]
    D --> E
    E --> F[Identificar aba de origem]
    F --> G{Aba encontrada?}
    G -->|Não| X1[Erro: origem não identificada]
    G -->|Sim| H[Ler colunas de carga, pedido, total e sequência\nconfiguradas em TEMPLATE T1/V1]
    H --> I{Há pedidos?}
    I -->|Não| X2[Erro: nenhum pedido extraído]
    I -->|Sim| J{Total da origem > 0?}
    J -->|Não| X3[Erro: total zerado, conferir coluna]
    J -->|Sim| K[Criar uma aba por pedido a partir de template\nPreencher carga P5, pedido P3 e sequência]
    K --> L[Pós-import em cada aba visível, exceto template]

    L --> M[Esperar fonte do banco, até 20 s]
    M --> N{Fonte pronta?}
    N -->|Não| S1[Aba ignorada]
    N -->|Sim| O[Limpar área B12:M53 e copiar itens\ncom tradução de siglas]
    O --> P[Recalcular e calcular peso das portas]
    P --> Q[Pintar termos e sequências]
    Q --> R[Registrar total da aba para conferência]
    R --> T[Ajustar fonte do pedido e remover linhas vazias]
    T --> U{Mais abas?}
    S1 --> U
    U -->|Sim| L
    U -->|Não| V{Divergência de totais?}
    V -->|Sim| W[Exibir relatório de conferência]
    V -->|Não| Y([Concluído])
    W --> Y
```

## Pré-requisitos

- Excel for Windows com macros habilitadas.
- Planilha modelo com aba `template` e as células de parâmetros:
  - `P3`: número do pedido
  - `P5`: número da carga
  - `TEMPLATE!T1`: coluna da sequência na origem (`0` desativa)
  - `TEMPLATE!V1`: coluna do total na origem (`0` desativa a validação de total)
- Fonte de dados do ERP configurada na aba de destino (conexão ODBC/Query). A conexão **não** fica no código; veja a seção de segurança.

## Instalação

1. Abra o arquivo modelo `.xlsb` ou `.xlsm`.
2. Alt+F11 → Arquivo → Importar arquivo, e selecione os `.bas` de `macros/`.
3. Salve o arquivo como `.xlsm`.

## Segurança

- Não versione o `.xlsb`: ele contém a string de conexão com o ERP, o nome do usuário do banco e dados de clientes.
- Não coloque usuário, senha ou caminhos de rede nos `.bas`. Use um arquivo de configuração local (`config.local.*`, já ignorado pelo `.gitignore`) ou configure a conexão só no Excel.
- Mantenha o repositório privado enquanto houver referências à estrutura do banco.

## Observações técnicas

- `modLucasBD.Substituir` apaga e reescreve `B12:M53`. Não rode sobre aba que tenha dados digitados manualmente.
- `POS_EXCLUIR_LINHA_INTEIRA = True` em `modPosImport` exclui linhas inteiras. Reexecutar pode deslocar as linhas da fonte `W11:W56`.
- `modEsperaBanco` aguarda até 20 s. Se o banco estiver lento, a aba é ignorada, e o aviso aparece no log do Debug, não em caixa de mensagem.
- `modTestes.ExecutarTestes` deve ser rodado após qualquer mudança em conversão, substituição ou peso.
