# Verificador_de_Planilhas

O programa tem como objetivo realizar comparação entre 2 planilhas subsequentes com o intuito de se obter as atualizações entre as 2 versões dos mesmos dados.

O mesmo foi realizado utilizando a linguagem de programação Python, juntamente com as bibliotecas “openpyxl” para manipulação de planilhas, e “time” para realização de uma pausa após o término da execução do programa

                            Como utilizar

      Informações Preliminares
      
•	As planilhas devem estar no formato “.xlsx”.

•	As planilhas utilizadas devem apresentar a mesma quantidade e organização de colunas entre si.

•	As planilhas devem estar no mesmo diretório (pasta) do programa.

•	A tabela com os dados da planilha deve se iniciar na célula “A1”, sendo permitida a presença de cabeçalho.

•	As planilhas devem apresentar 1 (uma) coluna com identificador único. Exemplo: CIAD para Aeródromos, CPF para Pessoas Físicas, CNPJ para Pessoas Jurídicas, etc.

•	A coluna com o identificador único deve estar no intervalo de “A” até “ZZZ”.

•	O arquivo de saída será criado no mesmo diretório (pasta) do programa, portando recomenda-se que, quando solicitado, não seja informado nome igual ao de outro arquivo já existente no diretório (pasta)


      Execução

Ao executar o programa será aberto o promp de comando, neste momento será solicitado o nome da planilha mais atualizada e, na sequência, a planilha desatualizada para realização das análises.

Para identificação dos dados, será solicitada a coluna na planilha onde se apresenta o identificador único.

Na sequência será perguntado se as planilhas apresentam cabeçalho, sendo necessário digitar 1 para caso negativo (os dados se iniciam na primeira linha da planilha) ou 2 para caso positivo (os dados se iniciam na segunda linha da planilha).

Por fim será solicitado o nome do arquivo resultante da análise das atualizações dos dados, este arquivo será salvo no formato “.xlsx”

Após o processamento, o tempo de processamento depende da quantidade de dados presentes nas planilhas, é criado o arquivo resultante no mesmo diretório do programa e será apresentada a mensagem “Programa Finalizado” no promp de comando. Após 3,5 segundos, o promp de comando será encerrado.

O arquivo gerado é composto pelos dados que apresentaram discrepância entre as duas planilhas, é possível observar que foi criada uma nova coluna denominada “Status” com informações referentes a tais diferenças, sendo limitado aos valores “Desatualizado”, “Atualizado”, “Adicionado” e “Excluído”. Também é possível observar que as linhas em que a coluna Status apresentam “Desatualizado” e “Atualizado”  encontram-se em sequência e se referem ao mesmo dado, sendo as discrepâncias apresentadas com coloração diferenciada, para chamar a atenção do usuário.
