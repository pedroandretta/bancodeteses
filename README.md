# bancodeteses
Instrumentos para a coleta e tratamento dos dados no Banco de Teses &amp; Dissertações da CAPES

Observações:
Os arquivos e métodos de coletas foram sistematizados utilizando o navegador Mozilla Firefox e Pacote Office Home and Student 2010.
O arquivo "coletador_codigo_fonte.docm" e "coletador_páginas_web.xlsm" foram criados em meados de fevereiro de 2017. Suas macros foram configuradas para o Painel de resultados do Portal de Teses e Dissertações da CAPES da época.
Para utilizar os arquivos  "coletador_codigo_fonte.docm" e "coletador_páginas_web.xlsm" o arquivo deve estar habilitado para edicão e com suas macros também habilitadas.
O arquivo "extrair_bancodetesescapes.iim" funciona no IMACRO.
O arquivo "extrair_bancodetesescapes.iim" e "unir arquivos txt.txt" podem substituir a função do "coletador_codigo_fonte.docm".
Uma vez utilizado o "extrair_bancodetesescapes.iim" e "unir arquivos txt.txt", deve-se isolar e TRATAR as string "href=" ", para que todas direcionem para a URL do registro SUCUPIRA, feito isso, basta copiar os endereços no "coletador_páginas_web.xlsm".
O arquivo "coletador_páginas_web.xlsm" possui duas macros: "adds" e "ReplicaDados". Deve-se usar primeiro a "adds" e em seguida "ReplicaDados".
    A macro "adds" deve ser constantemente editada, alterando o valor de 4 na expressão "For x = 1 To 4". O valor 4 deve corresponder ao mesmo numero de linhas utilizados na coluna A e F do documento.
    A macro "adds" copiará todo conteúdo html apresentado nas linhas A/F em novas planilhas, no mesmo arquivo. O computador deverá estar conectado a internet. Não será possível utilizar o excel enquanto o processo é "rodado", do contrário dará erro.
    A macro "ReplicaDados" irá unificar todas as planilhas geradas pela macro "adds".
Depois de coletados os dados do Portal da CAPES, deve-se seguir para os arquivo “sucupira - arrumar dados - tcc.xlsx”.
	Copie as LINHAS utilizadas no "coletador_páginas_web.xlsm" para “sucupira - arrumar dados - tcc.xlsx”.
	O arquivo “sucupira - arrumar dados - tcc.xlsx”, calculará valores textuais e irá “limpar os excessos de informações”. Ver linha 211.
		Sempre será preciso arrastar o jogo de fórmulas do arquivo “sucupira - arrumar dados - tcc.xlsx”, para a direita, ou remover as colunas não empregadas.
	Uma vez com os dados “limpos”, copie e cole especial, na planilha seguinte.
	Na planilha seguinte, copie novamente e cole especial – transpor para a próxima coluna.
Esses procedimentos valem para “sucupira - arrumar dados - ppgs.xlsx”, “sucupira - arrumar dados - docentes.xlsx”, “sucupira - arrumar dados - externo.xlsx”, desde que a origem dos dados sejam diretos do SUCUPIRA.

Dica:
Use o Open Refine para trarar dados de autor e título.
Ao coletar registros do Portal de Teses e Dissertações da CAPES, por vezes aparecem duplicatas. Certifique-se de tratar os dados antes de autor e título antes de remover duplicatas.
Use o Pajek “createpajek.exe” para após empregar o “elaborador de rede.xlsx”.
