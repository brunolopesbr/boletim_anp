# boletim_anp
Boletim de Produção da ANP

Esse repositório foi criado por Bruno Magalhães Lopes, como parte de seu trabalho de conclusão de curso no MBA em Data Science pela USP/Esalqp.

Os arquivos com o conteúdo das tabelas divulgadas pela ANP entre outubro de 2017 e a presente data estão no diretório [/csv](https://github.com/brunolopesbr/boletim_anp/tree/main/csv). As tabelas descontinuadas (incluindo as tabelas com dados do período da Covid estão em [/csv_descontinuado](https://github.com/brunolopesbr/boletim_anp/tree/main/csv_descontinuado).

Para baixar novos arquivos do site da ANP, executar o arquivo R Markdown  do arquivo [download_novos.Rmd](https://github.com/brunolopesbr/boletim_anp/blob/main/download_novos.Rmd). Para incorporar novas planilhas à base de dados, usar o arquivo [Importar_novos_arquivos.Rmd](https://github.com/brunolopesbr/boletim_anp/blob/main/Importar_novos_arquivos.Rmd). Eles dependem das funções personalizadas que são definidas no arquivo [funcoes_data_e_extrair.Rmd](https://github.com/brunolopesbr/boletim_anp/blob/main/Importar_novos_arquivos.Rmd).

Para baixar todos os arquivos do Boletim Mensal de Produção, o arquivo [download.Rmd](https://github.com/brunolopesbr/boletim_anp/blob/main/download.Rmd) está disponível. Para importar, limpar e produzir tabelas com os dados mensais de produção, recorrer ao arquivo [base_dados_atual.Rmd](https://github.com/brunolopesbr/boletim_anp/blob/main/download.Rmd).

A fonte original dos dados do [Boletim de Produção](https://www.gov.br/anp/pt-br/centrais-de-conteudo/publicacoes/boletins-anp/boletins/boletim-mensal-da-producao-de-petroleo-e-gas-natural) é a Agência Nacional do Petróleo, Gás Natural e Biocombustíveis.
