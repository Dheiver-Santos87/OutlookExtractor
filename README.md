# OutlookExtractor

O código acima serve para extrair os anexos de todos os emails presentes na pasta de entrada do Outlook e salvá-los em uma pasta de destino.

A classe OutlookExtractor possui um método __init__() que cria uma conexão com o Outlook e seleciona a pasta de entrada (que é a pasta "Caixa de entrada" por padrão).

O método principal da classe é o extract_attachments(), que realiza as seguintes etapas:

1. Armazena a quantidade de emails na pasta de entrada para exibir ao final do processo quantos emails foram processados.
2. Cria a pasta de destino (que é a pasta "Anexos" dentro de "Downloads") caso ela não exista.
3. Inicializa a contagem de emails processados.
4. Percorre cada email da pasta de entrada.
5. Por cada email, percorre cada anexo e obtém a extensão do arquivo.
6. Cria uma pasta de destino de acordo com a extensão do arquivo (por exemplo, a pasta "csv" para arquivos .csv) caso ela não exista.
7. Salva o anexo na pasta de destino correspondente.
8. Incrementa a contagem de emails processados.
9. Exibe a quantidade total de emails processados.
Por fim, é criada uma instância da classe OutlookExtractor e o método extract_attachments() é chamado para iniciar o processo de extração de anexos.
