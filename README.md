# PyGOM
Scripts em Python para automatização de diversas tarefas administrativas no sistema GOMNET e da empresa EMA. Atualmente conta com diversos módulos criados, além de outros ainda em construção.

*Obs.: A princípio não há interface GUI. Todos os módulos funcionam através de linha de comando.*

## Módulos - GOMNET - Selenium
 - Upload/Download de arquivos .dwg, .pdf, .zip;
 - Vinculação de supervisor à SOB;
 - Realização de vistoria da SOB;
 - Programação da SOB - com opção de lançamento de baremos;
 - Consulta de status da SOB;
 - Inserção de data *apta a fiscalizar* no Gommobile;
 - Consulta de SOB's energizadas;
 - Extração de material para arquivo em .xlsx - com opção de enviar por e-mail (Outlook);
 - Verificação de arquivos (.dwg, .pdf) anexados com base em palavras chaves, por exemplo: "AS_BUILT", "SGD", "DESENHO_LOCADO", etc.

## Módulos - GOMNET - BeautifulSoup + Requests
- Extração de status da SOB - opção de única ou várias (através de arquivo .xlsx);
- Download de arquivos .dwg, .pdf, .zip - utilizando Aria2c.  

## Módulos EMA
- Trabalho no arquivo *Pacote de Obras* (.xlsx) utilizando a biblioteca Openpyxl:
	1. Filtrando SOB's a programar;
	2. Fazendo uma varredura na pasta "Pacote de Obras" do sistema procurando por pastas previamente criadas;
	3. Criando pastas com a nomenclatura:
		- Número da SOB;
		- Data de vencimento.
	4. Inserindo dentro de cada pasta um arquivo *Rascunho.xlsx* modelo para utilização posterior pelo Locador e Programador.
	5. Criando um arquivo .txt contendo as novas SOB's para download dos arquivos .dwg no sistema GOMNET.
- Trabalho no arquivo *Programação Atualizada.xlsx*:
	1. Fazendo uma varredura nas colunas SOB e Data;
	2. Criando arquivos .xlsx de Planejamento com a nomenclatura:
		- Data;
		- Sob;
		- Nº da etapa.
- Trabalha no arquivo *Programação Atualizada.xlsx*:
	1. Fazendo uma varredura nas colunas Base, Sob, Data, Hora Inicial, Hora Final, Info Status, Prioridade, Estrutura, Observação, Apoio, Tipo de SGD, Nº de Dispositivo, Hora Início Desligamento, Hora Final Desligamento, Alimentador, Logradouro, Bairro, Município e Descrição do Serviço;
	2. Salvando as informações da linha atual em suas respectivas variáveis;
	3. Realizando um laço de repetição utilizando a função *zip()* e inserindo cada valor em seu respectivo campo no arquivo *"[Data][Sob][Nº da Etapa].xlsx"*.
