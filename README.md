# Projeto-de-Formatura
Este repositório contém o código, desenvolvido em Python, utilizado para controlar e analisar a rede EPRI CKT7 e IEEE 13 Barras com geração fovoltaica e armazenamento no Projeto de Formatura de minha autoria, apresentado em 2022, para a conclusão do curso de Engenharia Elétrica com Ênfase em Energia e Automação na POLI/USP. São utilizados alguns algoritmos desenvolvidos por Paulo Radatz, devidamente indicados. Eles podem ser baixados pelo link: https://github.com/PauloRadatz/tutorials_files.

Os parâmetros de entrada do programa são: 

parâmetros do sistema fotovoltaico (curva de irradiação, curva eficiência do inversor, curva da temperatura ambiente, tensão e potência nominais do módulo fotovoltaico); níveis de penetração (lista com 4 cenários de penetração FV e de armazenamento residencial e comercial); níveis limites de sobretensão e subtensão; número de simulações para cada nível penetração.

O programa segue, simplificadamente, a lógica abaixo:

Primeiramente, as cargas são classificadas em residenciais ou comerciais/industriais, através do algoritmo de machine learning KNN (é necessário ter sido treinado 1 vez), ou manualmente. Depois, são criados monitores para cada carga, obtendo os perfis de potência diários delas.  Integra-se o perfil de potência das cargas para a obtenção da energia diária delas. São obtidos, então, a energia diária e a potência nominal total (soma das energias e potências nominais de cada carga) para as classes residenciais e comercial/industrial, o que é reportado no arquivo "LoadsConsumption\_report.txt" gerado.

Para cada cenário de penetração FV e de armazenamento, então, são feitas "n" simulações. Em cada simulação, as cargas de cada classe consumidora são sorteadas e módulos fotovoltaicos são acopladas a elas, até que a penetração FV definida para cada classe seja atendida. Uma parcela dessas cargas com sistemas FV são sorteadas novamente para acoplamento de sistemas de armazenamento, de modo a ser atingida a penetração de armazenamento. Para esta configuração de rede são obtidas, então, as perdas técnicas e o número de transgressões de tensão.

São obtidas, então, a média e o desvio-padrão das perdas e do número de transgressões de tensão das "n" simulações, para cada cenário. Esses resultados são reportados em "StatisticalAnalysis\_report.txt" e por meio de gráficos, onde é possível comparar as perdas e transgressões de tensão para cada cenário. 

Instruções para utilização do programa:
  -Escrever "clear all" no início do arquivo "master.DSS", para limpar a memória de outros arquivos DSS (sem este procedimento, houve erros ao redefinir o PVSystem em alguns          algoritmos);
  -Escrever "Redirect PVSYstem.DSS" e "Redirect StorageFleet.DSS" no arquivo "master.DSS", pois várias funções deste trabalho redefinem o PVSystem com este nome;
  -Definir um EnergyMeter chamado "m1" para o alimentador da linha (os valores dos registradores dele são usados diversas vezes nos algoritmos).

Fabio Andrade Zagotto
fabio.zagotto@gmail.com
