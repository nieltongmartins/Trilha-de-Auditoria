# Correção de integridade de fórmulas matriciais

## Caminho investigado

Na versão 3.1 do `openpyxl`, uma célula que contém `<f t="array" ref="...">`
é carregada com `cell.value` do tipo `ArrayFormula`. A classe expõe `text`
(texto da fórmula, incluindo `=`) e `ref` (intervalo matricial). A outra classe
especial suportada nessa versão, `DataTableFormula`, expõe `ref` e os metadados
`ca`, `dt2D`, `dtr`, `r1`, `r2`, `del1` e `del2`, mas não texto de fórmula.

O leitor XML rápido identifica `array` e `dataTable` e encaminha o arquivo para
`_read_openpyxl`. Antes desta correção, `_read_openpyxl` copiava `cell.value`
diretamente para o snapshot. Assim, o objeto chegava a `compare_snapshots`, era
comparado por identidade, entrava em `CellChange` e, em `_serialize`, era
convertido por `str()`. O SQLite recebia o texto com endereço de memória e o
relatório apenas exportava esse mesmo texto. Portanto, a falha surgia antes do
relatório e podia produzir falsos positivos entre leituras equivalentes.

## Normalização adotada

`normalize_formula_value` é a normalização central. Ela é aplicada na entrada
do fallback openpyxl, defensivamente no comparador e antes da serialização. Uma
fórmula comum e valores que não são fórmulas permanecem inalterados.

As representações canônicas são um prefixo legível seguido de JSON compacto,
com chaves ordenadas:

```text
ARRAYFORMULA|{"ref":"D2:D100","text":"=SE(...)"}
DATATABLEFORMULA|{"ca":false,"del1":false,"del2":false,"dt2D":false,"dtr":false,"r1":null,"r2":null,"ref":"D2:D100"}
```

Isso preserva os campos semanticamente relevantes sem depender da identidade
ou do endereço de memória do objeto Python. A mesma representação passa a ser
usada pela comparação, pelo SQLite e pelo relatório.

## Histórico existente

Na investigação do repositório não havia arquivo SQLite versionado ou local
contendo `ArrayFormula object at`. Nenhum dado histórico é apagado ou migrado;
caso esse padrão seja encontrado em uma base de produção, uma eventual
migração deve ser tratada separadamente.
