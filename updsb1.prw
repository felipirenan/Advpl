#INCLUDE 'totvs.ch'
#include 'topconn.ch'
#Include "RwMake.CH"
#include "tbiconn.ch"

#DEFINE ENTER CHR(13)+CHR(10)

/*/{Protheus.doc} updsb1
Atualiza o cadastro de produto de acordo com o arquvio .csv 
@type function
@version  12.1.33
@author Renan Saran
@since 09/04/2022
/*/
User Function updsb1()
	local cTitulo := "[updsb1.prw] Atualiza Produtos(SB1)" //Define o tÃ­tulo da janela que aparecerÃ¡ na tela.
	local aTexto := {}	//Define um array contendo cada uma das linhas de texto que serÃ£o exibidas no corpo da tela.
	local aButtons := {} //Array com as opÃ§Ãµes disponÃ­veis para o usuÃ¡rio. SerÃ£o botÃµes do tipo SBUTTON() com a seguinte estrutura: { nTipo, lEnable, { | | AÃ§Ã£o() } }.
	local lOk := .F.	//variavel de controle, indica se o processe selecionado deve proseeguir(.T.) ou ser interrompido(.F.).
	Private _cArquivo  //recebe o caminho do arquivo
	Private _aMotsSel		:=	{}
	Private  _aSels	:=	{}
	Private _aMots		:=	{}





//inicia a inclusao dos botoes no arary que sera passado para classe FormBatch
	aAdd(aButtons, { 1, .T., {|| lOk := .T., FechaBatch() }} ) //1 = Ok
	aAdd(aButtons, { 14,.T., {|| _cArquivo := BuscaArquvio()}} ) //14 = Abrir
	aAdd(aButtons, { 6, .T., {|| CriaPergunta()}} ) //6 = imprimir
	aAdd(aButtons, { 2, .T., {|| lOk := .F., FechaBatch() }} ) //2 = Cancelar
	aAdd(aButtons, { 17, .T.,{|| MonChek()}} ) //17 = Filtrar

//inicia a inclusao dos textos no array que sera passado para classe FormBatch, esses textos serao apresentado no corpo da tela.
	AADD(aTexto, "Essa rotina tem como objetivo atualizar o cadastro de produtos.")
	AADD(aTexto, "Conforme arquivo de dados no formato .CSV")


	Formbatch(cTitulo, aTexto, aButtons) //"cria" tela pricipal da rotina

	if lok .and. !EMPTY(_cArquivo) //Se selecionado o botao Ok
		FWmsgRun(, { |oSay| GravaRegistro(oSay)},"Por Favor, Aguarde.","Procesando Dados...")
	elseif loK .and. EMPTY(_cArquivo)
		MsgAlert("Para prosseguir � necessario selecionar o arquivo.","Selecione o Arquivo")
	endif



RETURN


/*/{Protheus.doc} GravaRegistro
inicia o processo de atualizacao dos registros de acordo com o arquivo .csv
@type function
@version 12.1.33
@author Renan Sarn
@since 09/04/2022
/*/
Static Function GravaRegistro(oSay)


	Local aCab        := {} //recebe os campos e valores que devem ser alterados.
	local cConteudo //variavel que ira receber TODO o conteudo do arquivo .cvs.
	local lCampos := .T.
	local aCampos := {} //array que ira receber APENAS os  CAMPOS.
	local aDados := {} //array que ira receber APENAS o VALOR  dos campos.
	local cErro := ""
	local nPosCod

	//Variaveis de controle
	local cCampo := "" //varaivel que ira receber o campo(atual) que deve sofrer a alteraÃ§Ã£o.
	local nLinha //varaivel de controle do for, para andamento do array por linha.
	local nColuna //variavel de controle do for, para andemento do array por coluna.
	local nManipulador := FT_FUse(_cArquivo)

	Private lMsErroAuto := .F. //variavel para controle do erro

	If nManipulador <> -1 //Verifica se foi possivel abrir o arquivo -1 = nao foi.

		//-------------------------------------------------------------------------------------------------------------\\
		//--------------inicia o processo de tranferencia das informacoes do arquivo para o "advpl" -------------------\\.
		//-------------------------------------------------------------------------------------------------------------\\

		FT_FGoTop() //posiciona na primeira linha do arquivo.

		While !FT_FEOF()

			// oSay:SetText ('Gerando Template em Planilha...')
			// ProcessMessages()


			cConteudo := FT_FReadLn() //passa TODO o conteudo do arquivo para a variavel.

			If lCampos //verifica se e a primeria linha(1 linha = linha dos campos).
				aCampos := STRTOKARR2( cConteudo, ";", .T. ) // passa os CAMPOS para o array.
				lCampos := .F.
			else
				if !empty(cConteudo)
					AADD(aDados, STRTOKARR2( cConteudo, ";", .T.)) //passa o conteudo dos campo para o array .T. = valores vazios devem ser retornados.
				endif
			endIf

			FT_FSKIP() //pula para a proxima linha.

		endDo

		//-------------------------------------------------------------------------------------------------------------\\
		//--------------inicia o processo de gravacao(atualizao) das registros na tabela(SB1) -------------------\\.
		//-------------------------------------------------------------------------------------------------------------\\


		nPosCod	:= Ascan(aCampos, "B1_COD") //pega o posicao do campo B1_COD no array aCampos, assim evita que ele tenha que estar sempre na primeira posicao.

		for nLinha := 1 to len(aDados)


			oSay:SetText("Gravando Registros..." ) // ALTERA O TEXTO CORRETO
			ProcessMessage() // FOR�A O DESCONGELAMENTO DO SMARTCLIENT

			aCab:= {{aCampos[nPosCod], aDados[nLinha][nPosCod], Nil}} //adicona o codigo do produto que deve sofrer a alteracao

			for nColuna :=1 to len(aCampos)

				if nPosCod <> nColuna //verifica se a posicao atual nao e a mesma do campo B1_COD, assim nao grava no arrqy o codigo do produto que deve ser alterado novamente.

					cCampo := "SB1->" +UPPER(Alltrim((aCampos[nColuna]))) // recebe o campo atual, sera utilozado para verificar o tipo (N/D/C/L) do campo.

					//adiciona no array aCab, os campos e valores que devem ser alterados.
					if ValType(&cCampo) == "N" //Numerico.
						aadd( aCab, {aCampos[nColuna], val(aDados[nLinha][nColuna]), Nil})
					elseif ValType(&cCampo) == "D" //data.
						aadd( aCab, {aCampos[nColuna], alltrim(stod(aDados[nLinha][nColuna])), Nil})
					else
						aadd( aCab, {aCampos[nColuna], UPPER(alltrim(aDados[nLinha][nColuna])), Nil})
					endif

				endif

			next nColuna


			MSExecAuto({|x,y| MATA010(x,y)},aCab,4) // 4 = alteracao

			conout("Acabou de executar a op��o Alterar na rotina autom�tica do Cadastro de Complemento do Produto")

			//-- Retorno de erro na execu��o da rotina
			If lMsErroAuto
				conout("erro ao incluir o produto")
				cErro:=MostraErro()
			Else
				conout("Inclu�do com sucesso")
				ProcessMessage()
			Endif

		next nLinha

		if empty(cErro)
			MsgInfo("Produto(s) alterado(s) com sucesso!","Sucesso")
		endif

	else

		MsgAlert("Nao foi possivel abrir o arquivo CSV")
		ConOut( "Erro ao ler CSV - FERROR " + str(FError(),4) )

	endif


return


/*/{Protheus.doc} BuscaArquvio
Faz a busca do arquivo 
@type function
@version 12.1.33
@author Renan Saran
@since 09/04/2022
/*/
Static Function BuscaArquvio()

	Local cMascara   := "*CSV|*.csv" //define o filtro dos arquivos.
	Local cTitulo    := "Escolha o arquivo" //Titulo da janela de seleï¿½ï¿½o de arquivos.
	Local nMascPad   := 1
	Local cDirIni    := "c:\" //diretorio que sera inciado ao abrir a janela de seleï¿½ï¿½o de arquivo.
	Local lSalvar    := .T. //*.T. = Salvar || .F. = Abrir
	Local nOpcoes    := GETF_LOCALHARD + GETF_NETWORKDRIVE  //Mascara de bits para escolher as opções de visualização do Objeto. Disco Local + Rede.
	Local lArvore    := .F. //exibe diretorio do servidor .T. exibe .F. nao exibe.


return cGetFile(cMascara,cTitulo,nMascPad,cDirIni,lSalvar,nOpcoes,lArvore) //abre janela para seleï¿½ï¿½o de arquvivos.


/*/{Protheus.doc} GeraTemplateCsv
Gera templete padrao, no qual o usuario deve-se basear para criar o arquivo que se rpa importado
@type function
@version 12.1.33
@author Felipi Marques/Renan Saran
@since 09/04/2022
/*/

Static function GeraTemplateCsv()

	local cWorkSheet := "Template"
	local cTabela := "Layout SB1"
	Local oExcel := FWMsExcelEx():New()
	Local cArquivos :=GetTempPath()+'template.xml'

	oExcel:AddworkSheet(cWorkSheet)
	oExcel:AddTable (cWorkSheet,cTabela)

	oExcel:AddColumn(cWorkSheet,cTabela,"Col1",1,1)
	oExcel:AddColumn(cWorkSheet,cTabela,"Col2",1,1)


	oExcel:AddRow(cWorkSheet,cTabela,{"011",012})

	oExcel:SetCelBold(.F.)
	oExcel:SetCelFont('Arial')
	oExcel:SetCelItalic(.F.)
	oExcel:SetCelUnderLine(.F.)
	oExcel:SetCelSizeFont(10)

	oExcel:Activate()
	oExcel:GetXMLFile(cArquivos)

	oExcel:= MsExcel():New()
	oExcel:WorkBooks:Open(cArquivos)
	oExcel:SetVisible(.T.)
	oExcel:Destroy()




return

	// local CGETCAR := ';'
	// Local aAreaX3		:= SX3->(GetArea())
	// Local cConteud	:= ""
	// Local cCaminho	:= GetTempPath()
	// Local cArquivo	:= GetNextAlias()
	// Local cExtensao	:= ""
	// Local cAlias        := getNextAlias()
	// Local cQry	        := ""
	// Local y             :=1
	// local nX
	// local nTamArray
	// Local aStruSQL      := {}
	// local tmp := getTempPath()
	// local cCamSel := ""
	// local targetDir:= tFileDialog( "All files (*.*) | All Text files (*.csv) ",;
		// 	'Selecionar diret�rio',, tmp, .F., GETF_RETDIRECTORY  )

	// cConteud  := ""
	// cExtensao := ".csv"

	// nTamArray := len(_aMotsSel)

	// for nX := 1 to len(_aMotsSel)
	// 	if nX == 1
	// 		cCamSel +=  " , " + _aMotsSel[nX][1]
	// 	elseif nX == nTamArray
	// 		cCamSel += " , " +  _aMotsSel[nX][1]
	// 	else
	// 		cCamSel += " , " + _aMotsSel[nX][1]
	// 	endif
	// next


	// //Monta a query de sele��o
	// cQry := " SELECT  B1_COD " + cCamSel
	// cQry += " FROM   " + RETSQLTAB("SB1")
	// cQry += " WHERE "+RETSQLCOND("SB1") + " and "
	// cQry += " B1_COD between " + MV_PAR01 + " and " + MV_PAR02 + " and "
	// cQry += " B1_MSBLQL='2'
	// cQry += " ORDER BY B1_COD  DESC "

	// //Compatibiliza a query de acordo com o banco utilizado
	// cQry := ChangeQuery(cQry)

	// //Abre um alias com a query informada.
	// MPSysOpenQuery(cQry, cAlias)



	// aStruSQL := (cAlias)->(DbStruct())

	// dbSelectArea(cAlias)
	// (cAlias)->(dbGoTop())

	// // Carrega campos da da tabela
	// For y:=1 To Len(aStruSQL)
	// 	cConteud += AllTrim(aStruSQL[y][1] )   +cGetCar
	// Next y

	// // Retira a ultima ',' do da variavel
	// cConteud := SubStr(cConteud,1,Len(cConteud)-1)


	// cConteud += ENTER

	// // Prepara Exporta��o
	// While !(cAlias)->(EoF())

	// 	oSay:SetText ('Gerando Template em Planilha...')
	// 	ProcessMessages()

	// 	For y:=1 To Len(aStruSQL)
	// 		if aStruSQL[y][2] == "N"
	// 			cConteud += AllTrim(cValToChar( &("(cAlias)->"+aStruSQL[y][1])))   +cGetCar
	// 		else
	// 			cConteud +=  + AllTrim( &("(cAlias)->"+aStruSQL[y][1]) )  +cGetCar
	// 		endif
	// 	Next y
	// 	cValToChar(cConteud) += ENTER
	// 	(cAlias)->(dbSkip())
	// EndDo

	// If(Select(cAlias) > 0,(cAlias)->(dbCloseArea()),"")

	// 	// Ajusta o endere�o da pasta
	// 	cCaminho := targetDir + "\"

	// 	//Gera o arquivo
	// 	MemoWrite(cCaminho+cArquivo+cExtensao, cConteud)

	// 	//Tentando abrir o arquivo
	// 	nRet := ShellExecute("open", "excel.exe",cArquivo+cExtensao, cCaminho, 1)

	// 	//Se houver algum erro
	// 	If nRet <= 32
	// 		MsgStop("N�o foi poss�vel abrir o arquivo <b>"+cCaminho+cArquivo+cExtensao+"</b>!", "Aten��o")
	// 	EndIf

	// 	RestArea(aAreaX3)



	// 	RETURN

/*/{Protheus.doc} CriaPergunta
Cria as perguntas/parametros que serao utilizados na geracao do template
@type function
@version 12.1.33
@author Renan Saran
@since 13/4/2022
/*/
Static function CriaPergunta()

	local aPergs := {}

	aAdd(aPergs, {1, "Produto De ",  Space(TamSX3('B1_COD')[01]),  "", ".T.", "SB1", ".T.", 80,  .F.})
	aAdd(aPergs, {1, "Produto Ate ", Space(TamSX3('B1_COD')[01]),  "", ".T.", "SB1", ".T.", 80,  .T.})


	if parambox(aPergs, "Informe os Parametros")
		GeraTemplateCsv()	//FWmsgRun(, { |oSay| GeraTemplateCsv(oSay)},"Por Favor, Aguarde.","Procesando Dados")
	endif

return



Static Function MonChek()

	Local _aArea			:=	GetArea()
	Local _oBtnCancel
	Local _oBtnOK
	Local noLstBoxMot 		:=	1
	Local _oSayTitulo
	Private _oOk    		:=	LoadBitmap( GetResources(), "LBOK" )
	Private _oNo     		:=	LoadBitmap( GetResources(), "LBNO" )
	Private _lClose			:=	.F.
	Private loChkMarca 		:=	.F.
	Private _oChkMarca
	Private _oLstBoxMot
	Private _oDlgMot
	Public _aMots			:=	{}

	Default _lComisM		:=	.F.

	//Busca motivos do cadastro
	_aMots := GetCamps(_aMots)

	//Preenche o CHECK do 'Marca Todos'
	ChkAllMark()

	DEFINE MSDIALOG _oDlgMot TITLE "Selecione os Campos " FROM 000, 000  TO 480, 800 COLORS 0, 16777215 PIXEL

	@ 005, 006 SAY _oSayTitulo PROMPT "Marque os CAMPOS que deseja alterar" SIZE 143, 007 OF _oDlgMot COLORS 0, 16777215 PIXEL
	//@ 005, 332 CHECKBOX _oChkBSAtivos VAR loChkBSAtivos PROMPT "Somente ATIVOS" SIZE 053, 008 ON CLICK FilAtiv(_lComisM, loChkBSAtivos) OF _oDlgMot COLORS 0, 16777215 PIXEL

	@ 029, 003 LISTBOX _oLstBoxMot Var noLstBoxMot FIELDS HEADER "","Campo","Titulo","Descricao" FIELDSIZES 10,30,140,30,20 SIZE 392, 186 On DBLCLICK MarcaItem() OF _oDlgMot PIXEL
	_oLstBoxMot:SetArray(_aMots)
	_oLstBoxMot:bLine:={ ||{iif(_aMots[_oLstBoxMot:nAT,1],_oOk,_oNo), _aMots[_oLstBoxMot:nAT,2], _aMots[_oLstBoxMot:nAT,3], _aMots[_oLstBoxMot:nAT,4]}}

	@ 017, 006 CHECKBOX _oChkMarca VAR loChkMarca PROMPT "Marca/Desmarca Todos" SIZE 080, 008 ON CLICK Marca(loChkMarca) OF _oDlgMot COLORS 0, 16777215 PIXEL

	@ 220, 357 BUTTON _oBtnOK PROMPT "OK" ACTION CloseDlg(.T.) SIZE 037, 012 OF _oDlgMot PIXEL
	@ 220, 311 BUTTON _oBtnCancel PROMPT "Cancelar" ACTION CloseDlg(.F.) SIZE 037, 012 OF _oDlgMot PIXEL



	ACTIVATE MSDIALOG _oDlgMot CENTERED VALID _lClose



	RestArea(_aArea)

return .F.

Static function CloseDlg(lOk)

	Default _lOK	:=	lok

	//Usuario clicou no botao OK
	If _lOK

		_aMotsSel	:=	FillSele()
		If !Empty(_aMotsSel)
			_lClose	:=	.T.
			_oDlgMot:End()
		Else
			_lClose	:=	.F.
			MsgAlert("Selecione pelo menos UM CAMPO antes de confirmar!", "ATENCAO - updsb1.prw")
		Endif


	Else
		_aMotsSel	:=	{}
		_lClose		:=	.T.
		_oDlgMot:End()
	Endif




return

Static Function MarcaItem()

	//Inverte o Item clicado
	_aMots[_oLstBoxMot:nAT,1] := !_aMots[_oLstBoxMot:nAT,1]

	//Atualiza CHECK BOX "Marca/Desmarca Todos"
	ChkAllMark()

Return Nil

Return

Static Function FillSele()

	//Local _aSels	:=	{}
	Local _nAT		:=	1
	Local _nCols	:=	2
	_aSels:={}

	For _nAT := 1 To Len(_aMots)

		IF _aMots[_nAT][1]
			//Inclui nova linha
			Aadd(_aSels, {})

			//Pula primeira coluna que indica quais estao selecionados
			For _nCols := 2 To Len(_aMots[_nAT])
				Aadd(_aSels[Len(_aSels)], _aMots[_nAT][_nCols])
			Next

		Endif

	Next

Return _aSels


Static Function Marca(_lMarcado)

	Local _nAT		:=	1

	Default _lMarcado	:=	.F.

	For _nAT := 1 To Len(_aMots)
		_aMots[_nAT][1]	:=	_lMarcado
	Next

	_oLstBoxMot:Refresh()

Return Nil

Static Function ChkAllMark()

	Local _nAT		:=	1
	Local _lAll		:=	.T.

	//Verifica se tudo esta marcado
	For _nAT := 1 To Len(_aMots)
		If !_aMots[_nAT][1]
			_lAll	:=	.F.

			Exit
		Endif
	Next

	//Atualiza CHECK 'Marca Todos'
	loChkMarca := _lAll

	If _oChkMarca <> Nil
		_oChkMarca:Refresh()
	Endif

Return Nil


Static Function GetCamps(_aMots)

	Local _aArea		:=	GetArea()
	Local _cAlias		:=	GetNextAlias()
	Local cQuery		:=	""
	local nX
	Local _lMarcado		:=	.F.
	Default _aMots := _aMots



	cQuery := " Select X3_CAMPO, X3_TITULO, X3_DESCRIC " + ENTER
	cQuery += " FROM SX3010 " + ENTER
	cQuery += " WHERE X3_ARQUIVO = 'SB1' AND X3_VISUAL <> 'V' AND " + ENTER
	cQuery += " X3_CAMPO <> 'B1_OBS' AND X3_CAMPO <> 'B1_VM_PROC' "
	cQuery += " ORDER BY X3_CAMPO " + ENTER


	DbUseArea (.T., "TOPCONN", TcGenQry (NIL,NIL,cQuery), _cAlias, .T., .T.)

	DbSelectArea(_cAlias)

	for nX := 1 to len(_aMots)
		IF _aMots[nX][1] = .T.
			return _aMots
		endif

	next nX

	While (_cAlias)->(!Eof())


		Aadd(_aMots, {_lMarcado, (_cAlias)->X3_CAMPO, Alltrim((_cAlias)->X3_TITULO), (_cAlias)->X3_DESCRIC})

		(_cAlias)->(DbSkip())

	EndDo

	(_cAlias)->(DbCloseArea())

	RestArea(_aArea)

Return _aMots

