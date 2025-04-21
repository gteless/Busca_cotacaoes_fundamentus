# Busca de Cotações Fundamentus

Este projeto automatiza o processo de consulta de cotações de ações no site [Fundamentus](https://www.fundamentus.com.br) e organiza os resultados em um arquivo Excel. O código é desenvolvido para ser utilizado como um robô de automação, simplificando a obtenção de dados financeiros de ações.

## Funcionalidades

- O código lê um arquivo Excel na pasta de entrada (`INPUT`).
- Processa as cotações de ações especificadas no arquivo.
- Para cada ação, consulta a página correspondente no [Fundamentus](https://www.fundamentus.com.br) e captura os dados financeiros, como:
  - Cotação
  - Mínimo 52 semanas
  - Máximo 52 semanas
  - Valor de mercado
  - Número de ações
  - P/L (Preço/Lucro)
  - LPA (Lucro por ação)
  - P/VP (Preço/Valor Patrimonial)
  - VPA (Valor Patrimonial por ação)
  - ROE (Retorno sobre o patrimônio líquido)
- Após a captura dos dados, os resultados são salvos no arquivo Excel e movidos para a pasta de finalizados (`FINALIZADO`).
- Todo o processo é logado, sendo possível acompanhar o progresso e identificar qualquer erro que ocorra.

## Estrutura de Pastas

![image](https://github.com/user-attachments/assets/2fcb76f2-8518-405c-85b9-97be4879bb4e)

O projeto cria as seguintes pastas para organizar o processamento dos arquivos:

1. **LOG**: Pasta onde são gerados logs detalhados do que foi feito durante o processo. Um arquivo de log será criado com a data e hora de execução, contendo informações detalhadas sobre o andamento do processamento.
2. **INPUT**: Pasta onde o usuário deve colocar os arquivos Excel a serem processados.
3. **PROCESSAMENTO**: Pasta onde os arquivos são temporariamente armazenados enquanto são processados.
4. **FINALIZADO**: Pasta onde os arquivos processados são movidos após a finalização do processo.


## Exemplo de um arquivo de log gerado:

2025-04-17 23:59:28 - =========================== INICIO - Valida Arquivo ================================ 

2025-04-17 23:59:28 - Validando se já não contem arquivo na pasta PROCESSAMENTO 

2025-04-17 23:59:28 - Arquivo encontrado na pasta PROCESSAMENTO: C:\Users\gabri\BPA001 - BuscaCotacoes\3. PROCESSAMENTO\Pasta1.xlsx 

2025-04-17 23:59:28 - Capturando driver 

2025-04-17 23:59:28 - Pegando o nome apenas para o driver .xlsx 

2025-04-17 23:59:28 - Definindo connection string 

2025-04-17 23:59:29 - Pegando worksheet 

2025-04-17 23:59:29 - Query executada: SELECT * FROM [DADOS$] WHERE [Status] IS NULL 

2025-04-17 23:59:29 - Abrindo navegador na página - (https://www.fundamentus.com.br/index.php) 

2025-04-17 23:59:34 - Consultado o Ticker POMO3 no site: https://www.fundamentus.com.br/detalhes.php?papel=POMO3 

2025-04-17 23:59:34 - Página carregada com sucesso, iniciando a captura dos dados 

2025-04-17 23:59:34 - Dados capturados: Cotação: 5,09 | Min 52 sem: 4,17 | Max 52 sem: 6,87 | Valor de mercado: 5.783.620.000 | Nro. Ações: 1.136.270.000 | Dia: 2,62% | P/L: 4,82 | LPA: 1,06 | P/VP: 1,44 | VPA: 3,54 | ROE: 29,8% ... 

2025-04-17 23:59:35 - Matando sessão do driver no fim do processo 

2025-04-17 23:59:36 - Movendo arquivo para finalizado


## Instruções de Uso

1. Coloque o arquivo Excel na pasta `INPUT`. Este arquivo deve conter uma lista de tickers (códigos de ações) a serem consultados no site Fundamentus.
   
2. O código irá ler o arquivo, processar as cotações para cada ticker e gerar um novo arquivo Excel na pasta `FINALIZADO`.

3. O arquivo de log será gerado na pasta `LOG`, contendo informações detalhadas sobre cada etapa do processo.

4. Após a execução, o arquivo processado será movido para a pasta `FINALIZADO`.
