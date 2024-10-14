#INCLUDE 'TOTVS.CH'
User Function IMPSE1(nOpc)
	Local oProcess := Nil
	Private lApenasBaixas := Nil
	Private lBaixar := Nil
	Private TIMESTAMP := FWTimeStamp(2)
	Private ID := MD5(TIMESTAMP, 2)
	Private aErros := {}

	Default nOpc := 3
	// MsAguarde({|| LERCSV(@lApenasBaixas, @lBaixar, nOpc, @aErros)}, "Aguarde...", "Processando Registros...")
	oProcess := MsNewProcess():New({|| LERCSV(oProcess, @lApenasBaixas, @lBaixar, nOpc, @aErros)}, "Processando...", "Aguarde...", .T.)
	oProcess:Activate()

	// U_GeraLog(aErros)
Return

/*/{Protheus.doc} LERCSV
    Tem o objetivo de ler todas as linhas de um arquivo selecionado e retornar o conteudo de cada linha dentro de uma posição num Array
    @type  Function
    @autor Saulo Maciel
    @since 21/09/2020
    @version 0.0.1
    @param cFile, Caractere, Nome do arquivo que será lido
    @return aDados, Array, Lista dos dados do arquivos dentro de um array
    @example
    Exemplo de Mascara para a função cGetFile: "Todos os arquivos|*.*|Arquivos de Texto|*.txt|Arquivos XML|*.xml|CSV Separado por Virgula|*.csv|"
    @see (https://tdn.totvs.com/display/framework/FWFileReader; https://tdn.totvs.com/display/tec/cGetFile)
    /*/
Static Function LERCSV(oProcess, lApenasBaixas, lBaixar, nOpc, aErros)
	Local oFile
	Local cFile := Space(50) // Aumentado o espaço para 255 caracteres
	Local aResps := {}
	Local aPergs := {}
	Local nQtdReg := 0
	Local aDados := {}
	Local nAtual := 0
	Local nLinha := 1
	Local nTam := 0
	Local cDir := ''
	Local aFiles := {}
	Local nIx := 0

	Private aCabec := {}

	// Monta os parâmetros para o Parambox
	aAdd(aPergs, {6, "Escolha um arquivo CSV", cFile, "", "", "", 50, .F., "CSV (separado por virgula) (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"})
	aAdd(aPergs,{3,"Inclue Titulos",1,{"Sim","Não"},50,"",.F.})
	aAdd(aPergs,{3,"Realiza Baixa",1,{"Sim","Não"},50,"",.F.})

	// Adiciona as coordenadas de posição e tamanho da janela do Parambox
	// If ParamBox(aPergs, 'Leitura de arquivos CSV', @aResps, , , , , , , , , , 800, 600, .T., .T.)
	If Parambox(aPergs, "Leitura de arquivos CSV", aResps,,,,,,,,.T.,.T.)
		cFile := Alltrim(aResps[1])
		If Empty(cFile)
			Alert("Nenhum arquivo selecionado", "Erro")
			Return {}
		EndIf

		If aResps[2] == 2 .AND. aResps[3] == 1
			lApenasBaixas := .T.
			lBaixar := .T.
		ElseIf aResps[3] == 1
			lBaixar := .T.
		ElseIf aResps[2] == 2 .AND. aResps[3] == 2
			FwAlertError("Você escolheu não fazer nada")
			Return
		Else
			lApenasBaixas := .F.
			lBaixar := .F.
		EndIf

		If FWAlertNoYes("Ao continuar, todos os arquivos '.csv' da mesma pasta serao processados! Seguir em frente?", "Processamento em massa")
			nTam := Len(cFile) - Len(ExtractFile(cFile))
			cDir := Left(cFile, nTam)
			aFiles := Directory(cDir+"*.csv")

			oProcess:SetRegua1(Len(aFiles))

			For nIx := 1 To Len(aFiles)
				cFile := cDir+aFiles[nIx,1]
				oFile := FWFileReader():New(cFile)
				oProcess:IncRegua1("Arquivo: " + ExtractFile(cFile))

				nLinha := 1
				nAtual := 0
				If (oFile:Open())
					aDados := oFile:GetAllLines()
					oProcess:SetRegua2(Len(aDados))
					nQtdReg := Len(aDados)
					While nLinha <= nQtdReg
						nAtual++
						oProcess:IncRegua2("Incluindo registro " + cValToChar(nAtual) + " de " + cValToChar(nQtdReg) + "...")
						If nLinha == 1
							aCabec := SEPARA(aDados[nLinha], ";")
							nLinha++
							Loop
						EndIf

						Executa(nOpc, aDados[nLinha], aCabec, @aErros)
						nLinha++
					EndDo
					oFile:Close()
					FreeObj(oFile)
				Else
					Alert("Não foi possível abrir o arquivo: "+cFile, "Erro")
					Return {}
				EndIf
			Next
		EndIf
	Else
		FwAlertError("Operação cancelada pelo usuário")
	Endif
Return

Static Function Executa(nOpc, aCAR, aCabec, aErros)
	Local aArea := GetArea()
	Local aAreaSE1 := SE1->(GetArea())
	Local aAjustes := {}
	Local lExiste := .F.
	Local aExiste := {}
	Local nCampo := 0
	Local cCampo := ""
	Local aLinha := {}

	// Declaração de variáveis para posições
	Private nPosFilial := nPosPrefixo := nPosNum := nPosParcela := nPosTipo := nPosNaturez := nPosFornece := nPosLoja := ;
		nPosCCusto := nPosEmissao := nPosVencto := nPosVencRea := nPosValor := nPosHist := nPosDtBx := ;
		nPosAutBanco := nPosAutAgencia := nPosAutConta := 0

	Default nOpc := 3

	// Atribuição das posições em uma única passagem sobre o cabeçalho
	For nCampo := 1 To LEN(aCabec)
		cCampo := ALLTRIM(aCabec[nCampo])
		Do Case
		Case cCampo == "E1_FILIAL"
			nPosFilial := nCampo
		Case cCampo == "E1_PREFIXO"
			nPosPrefixo := nCampo
		Case cCampo == "E1_NUM"
			nPosNum := nCampo
		Case cCampo == "E1_PARCELA"
			nPosParcela := nCampo
		Case cCampo == "E1_TIPO"
			nPosTipo := nCampo
		Case cCampo == "E1_NATUREZ"
			nPosNaturez := nCampo
		Case cCampo == "E1_CLIENTE"
			nPosFornece := nCampo
		Case cCampo == "E1_LOJA"
			nPosLoja := nCampo
		Case cCampo == "E1_CCUSTO"
			nPosCCusto := nCampo
		Case cCampo == "E1_EMISSAO"
			nPosEmissao := nCampo
		Case cCampo == "E1_VENCTO"
			nPosVencto := nCampo
		Case cCampo == "E1_VENCREA"
			nPosVencRea := nCampo
		Case cCampo == "E1_VALOR"
			nPosValor := nCampo
		Case cCampo == "E1_HIST"
			nPosHist := nCampo
		Case cCampo == "DT_BAIXA"
			nPosDtBx := nCampo
		Case cCampo == "AUTBANCO"
			nPosAutBanco := nCampo
		Case cCampo == "AUTAGENCIA"
			nPosAutAgencia := nCampo
		Case cCampo == "AUTCONTA"
			nPosAutConta := nCampo
		EndCase
	Next

	If !lApenasBaixas
		aLinha := Separa(aCAR, ";")
		If nPosNum > 0
			aExiste := ExisteReg(aLinha, @aErros, @aAjustes, aCAR)
			lExiste := aExiste[1]
			If !lExiste
				ExecAuto(aLinha, nOpc, @aErros, aCAR)
			Else
				If nOpc == 3
					IncLog(@aErros, aLinha, "AJUDA:FA050NUM   Número título já existe para este forne-cedor.",,aCAR, "Falha - Inclusao")
				ElseIf nOpc == 5
					ExecAuto(aLinha, nOpc, @aErros, aCar)
				EndIf
			EndIf
		EndIf
	EndIf

	If lBaixar
		aLinha := Separa(aCAR, ";")
		BaixaSE1(nOpc, aLinha, @aErros, aCAR)
	EndIf

	If Len(aAjustes) > 0
		Ajustar(aAjustes)
	EndIf

	SE1->(RestArea(aAreaSE1))
	RestArea(aArea)
Return

Static Function ExecAuto(aReg, nOpc, aErros, cLinha)
	Local aArrayI := {}
	Local nPosErro := 0
	Local cPrefixo := GetValue(aReg, nPosPrefixo, "S")
	Local cNum := GetValue(aReg, nPosNum, "S")
	Local cParcela := GetValue(aReg, nPosParcela, "S")
	Local cTipo := GetValue(aReg, nPosTipo, "S")
	Local cNaturez := ALLTRIM(GetValue(aReg, nPosNaturez, "S"))
	Local cFornece := GetValue(aReg, nPosFornece, "S")
	Local cLoja := GetValue(aReg, nPosLoja, "S")
	Local cCCusto := GetValue(aReg, nPosCCusto, "S")
	Local dEmissao := GetValue(aReg, nPosEmissao, "D")
	Local dVencto := GetValue(aReg, nPosVencto, "D")
	Local dVencRea := DataValida(dVencto)
	Local nValor := GetValue(aReg, nPosValor, "N")
	Local cHist := GetValue(aReg, nPosHist, "S")
	Local aRet := {}

	Private lMsErroAuto := .F.
	Private lAutoErrNoFile := .T.

	cHist := IIF(Empty(cHist), "IMPORTADO VIA CSV: ", cHist)
	cHist += ": " + TIMESTAMP

	aArrayI := { ;
		{ "E1_PREFIXO"  , cPrefixo  , NIL },;
		{ "E1_NUM"      , cNum      , NIL },;
		{ "E1_PARCELA"  , cParcela  , NIL },;
		{ "E1_TIPO"     , cTipo     , NIL },;
		{ "E1_NATUREZ"  , cNaturez  , NIL },;
		{ "E1_CLIENTE"  , cFornece  , NIL },;
		{ "E1_LOJA"     , cLoja     , NIL },;
		{ "E1_CCUSTO"   , cCCusto   , NIL },;
		{ "E1_EMISSAO"  , dEmissao  , NIL },;
		{ "E1_VENCTO"   , dVencto   , NIL },;
		{ "E1_VENCREA"  , dVencRea  , NIL },;
		{ "E1_VALOR"    , nValor    , NIL },;
		{ "E1_HIST"     , cHist     , NIL } }

	BEGIN TRANSACTION
		MsExecAuto({|x,y| FINA040(x,y)}, aArrayI, nOpc)


		If lMsErroAuto
			aRet := GetAutoGRLog()
			nPosErro := aScan(aRet, {|x| "Inval" $ Alltrim(x)})
			If nPosErro > 0
				IncLog(@aErros, aReg, aRet[1], aRet[nPosErro], cLinha, "Falha - Inclusao")
			Else
				IncLog(@aErros, aReg, aRet[1],,cLinha, "Falha - Inclusao")
			EndIf
		Else
			If nOpc == 3
				IncLog(@aErros, aReg, "Titulo incluido com sucesso!", , cLinha, "Sucesso - Inclusao")
			ElseIf nOpc == 5
				IncLog(@aErros, aReg, "Titulo excluido com sucesso!", , cLinha, "Sucesso - Inclusao")
			EndIf
		EndIf
	END TRANSACTION
Return
Static Function BaixaSE1(nOpc, aReg, aErros, cLinha)
	// Declarações de variáveis locais
	Local lRet := .T.
	Local cHistBaixa := "BAIXA VIA IMPORTACAO: " + TIMESTAMP
	Local aBaixa := {}
	Local cAliasSE1 := GetNextAlias()
	Local cPrefixo := GetValue(aReg, nPosPrefixo, "S")
	Local cNum := GetValue(aReg, nPosNum, "S")
	Local cParcela := GetValue(aReg, nPosParcela, "S")
	Local cTipo := GetValue(aReg, nPosTipo, "S")
	Local cFornece := GetValue(aReg, nPosFornece, "S")
	Local cLoja := GetValue(aReg, nPosLoja, "S")
	Local nValor := GetValue(aReg, nPosValor, "N")
	Local cHist := GetValue(aReg, nPosHist, "S") + ": "+ TIMESTAMP
	Local dDtBaixa := GetValue(aReg, nPosDtBx, "D")
	Local cAutBanco := GetValue(aReg, nPosAutBanco, "S")
	Local cAutAgencia := GetValue(aReg, nPosAutAgencia, "S")
	Local cAutConta := GetValue(aReg, nPosAutConta, "S")
	Local aBanco := {}
	Local dDataBkp := dDatabase

	Private lMsErroAuto := .F.
	Private lAutoErrNoFile := .T.

	// Operação a ser realizada (3 = Baixa, 5 = Cancelamento, 6 = Exclusão)
	Default nOpc := 3

	// Valor a ser baixado
	Default nValor := 0
	Default nSeqBx := 1

	// Consulta SQL para obter o registro
	BeginSql Alias cAliasSE1
        SELECT
            SE1.E1_FILIAL,
            SE1.E1_PREFIXO,
            SE1.E1_NUM,
            SE1.E1_PARCELA,
            SE1.E1_TIPO,
            SE1.E1_CLIENTE,
            SE1.E1_LOJA,
            SE1.E1_SALDO,
            SE1.E1_VALOR
        FROM %Table:SE1% AS SE1
        WHERE
            SE1.%NotDel%
			AND SE1.E1_FILIAL = %exp:FwXFilial('SE1')%
			AND Se1.E1_PREFIXO = %exp:cPrefixo%
            AND SE1.E1_NUM = %exp:cNum%
            AND SE1.E1_PARCELA = %exp:cParcela%
			AND SE1.E1_TIPO = %exp:cTipo%
            AND SE1.E1_CLIENTE = %exp:cFornece%
            AND SE1.E1_LOJA = %exp:cLoja%
	EndSql

	// Criando diretório, se não existir
	MakeDir('C:/Temp/', Nil, .F.)
	// Escrevendo a última query no memo
	MemoWrite('C:/Temp/query.sql', GetLastQuery()[2])


	aBanco := GetBanco(cAutBanco, cAutAgencia, cAutConta)
	dDatabase := dDtBaixa

	// Verificação se o registro foi encontrado
	BEGIN TRANSACTION
		If Len(aBanco) > 0
			If (cAliasSE1)->(!Eof())
				// Verifica o tipo de operação a ser realizada
				If nOpc == 3 // Operação de Baixa
					If lRet := (nValor + (cAliasSE1)->E1_SALDO) > 0
						nValor := If(nValor > 0, nValor, (cAliasSE1)->E1_SALDO)
					EndIf
				ElseIf (cAliasSE1)->E1_VALOR >= (cAliasSE1)->E1_SALDO
					nValor := 0
				EndIf

				If lRet
					// Adiciona os dados da baixa ao array
					AAdd(aBaixa, {"E1_FILIAL", (cAliasSE1)->E1_FILIAL, nil})
					AAdd(aBaixa, {"E1_PREFIXO", (cAliasSE1)->E1_PREFIXO, nil})
					AAdd(aBaixa, {"E1_NUM", (cAliasSE1)->E1_NUM, nil})
					AAdd(aBaixa, {"E1_PARCELA", (cAliasSE1)->E1_PARCELA, nil})
					AAdd(aBaixa, {"E1_TIPO", (cAliasSE1)->E1_TIPO, nil})
					AAdd(aBaixa, {"E1_CLIENTE", (cAliasSE1)->E1_CLIENTE, nil})
					AAdd(aBaixa, {"E1_LOJA", (cAliasSE1)->E1_LOJA, nil})
					AAdd(aBaixa, {"AUTMOTBX", "NOR", nil})
					AAdd(aBaixa, {"AUTBANCO", aBanco[1], nil})
					AAdd(aBaixa, {"AUTAGENCIA", aBanco[2], nil})
					AAdd(aBaixa, {"AUTCONTA", aBanco[3], nil})
					AAdd(aBaixa, {"AUTDTBAIXA", dDtBaixa, nil})
					AAdd(aBaixa, {"AUTDTCREDITO", dDtBaixa, nil})
					AAdd(aBaixa, {"AUTHIST", IIF(Empty(cHist), cHistBaixa, cHist), nil})
					AAdd(aBaixa, {"AUTVLRPG", nValor, nil})

					//Pergunte da rotina
					AcessaPerg("FINA070", .F.)

					// Chama a rotina de baixa manual (FINA070)


					If nOpc == 5
						MsExecauto({|x,y,z,v| FINA070(x,y,z,v)}, aBaixa, nOpc, .F., nSeqBx)
					Else
						MsExecAuto({|x, y| FINA070(x, y)}, aBaixa, nOpc)
					EndIf

					// Verifica se houve erro na rotina de baixa
					If lMsErroAuto
						aRet := GetAutoGRLog()
						nPosErro := aScan(aRet, {|x| "Inval" $ Alltrim(x)})
						If Len(aRet) > 0
							If Empty(nPosErro)
								nPosErro := 1
							EndIf
							IncLog(@aErros, aReg, "Não Baixado", aRet[nPosErro], cLinha, "Falha - Baixa")
						ElseIf nValor > (cAliasSE1)->E1_SALDO
							IncLog(@aErros, aReg, "Não Baixado", "Tentativa de baixar valor maior do que saldo do titulo", cLinha)
						Else
							IncLog(@aErros, aReg, "Não Baixado", "Erro não catalogado", cLinha, "Falha - Baixa")
						EndIf
					Else
						// Mensagem de sucesso conforme a operação realizada
						If nOpc == 3
							IncLog(@aErros, aReg, "", "Baixa efetuada com sucesso!", cLinha, "Sucesso - Baixa")
						EndIf
					EndIf
				Else
					IncLog(@aErros, aReg, "", "O título não possui saldo a pagar em aberto", cLinha, "Falha - Baixa")
				EndIf
			Else
				IncLog(@aErros, aReg, "", "O título a pagar não foi localizado", cLinha, "Falha - Baixa")
			EndIf
		Else
			IncLog(@aErros, aReg, cAutBanco+"/"+cAutAgencia+"/"+cAutConta, "Banco não Localizado", cLinha, "Falha - Baixa")
		EndIf
		(cAliasSE1)->(DbCloseArea())

		dDatabase := dDataBkp
	END TRANSACTION
Return

Static Function UPDHIST(nOpc, aReg, aErros)
	// Declarações de variáveis locais
	Local cHistBaixa := "TITULO IMPORTADO #001"
	Local cAliasSE1 := GetNextAlias()
	Local cPrefixo := GetValue(aReg, nPosPrefixo, "S")
	Local cNum := GetValue(aReg, nPosNum, "S")
	Local cParcela := GetValue(aReg, nPosParcela, "S")
	Local cTipo := GetValue(aReg, nPosTipo, "S")
	Local cFornece := GetValue(aReg, nPosFornece, "S")
	Local cLoja := GetValue(aReg, nPosLoja, "S")
	Local nValor := GetValue(aReg, nPosValor, "N")

	Private lMsErroAuto := .F.
	Private lAutoErrNoFile := .T.

	// Operação a ser realizada (3 = Baixa, 5 = Cancelamento, 6 = Exclusão)
	Default nOpc := 3

	// Valor a ser baixado
	Default nValor := 0
	Default nSeqBx := 1

	// Consulta SQL para obter o registro
	BeginSql Alias cAliasSE1
        SELECT
            SE1.R_E_C_N_O_ REG
        FROM %Table:SE1% AS SE1
        WHERE
            SE1.%NotDel%
            AND SE1.E1_PREFIXO = %exp:cPrefixo%
            AND SE1.E1_NUM = %exp:cNum%
            AND SE1.E1_PARCELA = %exp:cParcela%
            AND SE1.E1_TIPO = %exp:cTipo%
            AND SE1.E1_CLIENTE = %exp:cFornece%
            AND SE1.E1_LOJA = %exp:cLoja%
	EndSql

	// Verificação se o registro foi encontrado

	While (cAliasSE1)->(!Eof())
		SE1->(DbGoTo((cAliasSE1)->REG))
		Reclock("SE1",.F.)
		REPLACE E1_HIST WITH cHistBaixa
		MsUnlock()
		(cAliasSE1)->(DbSkip())
	EndDo
	(cAliasSE1)->(DbCloseArea())

Return

// Função auxiliar para obter valor de aReg com verificação de posição e tipo
Static Function GetValue(aReg, nPos, cTipo)
	Local uRet
	If nPos > 0
		uRet := aReg[nPos]
		Do Case
		Case cTipo == "S"
			uRet := aReg[nPos]
		Case cTipo == "N"
			uRet := STRTRAN(aReg[nPos], ",","")
			uRet := STRTRAN(uRet, ".","")
			uRet := GetDToVal(uRet)
			uRet /= 100
		Case cTipo == "D"
			uRet := CTOD(aReg[nPos])
		EndCase
	Else
		Do Case
		Case cTipo == "S"
			uRet := ""
		Case cTipo == "N"
			uRet := 0
		Case cTipo == "D"
			uRet := CTOD("")
		EndCase
	EndIf
Return uRet


Static Function GetBanco(cAutBanco, cAutAgencia, cAutConta)
	Local aRet := {}
	Local cAliasSA6 := GetNextAlias()

	BEGINSQL ALIAS cAliasSA6
		SELECT
			SA6.A6_COD, SA6.A6_AGENCIA, SA6.A6_NUMCON
		FROM %Table:SA6% AS SA6
		WHERE
			SA6.%NotDel%
			AND A6_COD = %exp:cAutBanco%
			AND A6_AGENCIA = %exp:cAutAgencia%
			AND A6_NUMCON = %exp:cAutConta%
	ENDSQL

	If (cAliasSA6)->(!Eof())
		aRet := {(cAliasSA6)->A6_COD, (cAliasSA6)->A6_AGENCIA, (cAliasSA6)->A6_NUMCON}
	EndIf
	(cAliasSA6)->(DbCloseArea())

Return aRet


Static Function ExisteReg(aReg, aErros, aAjustes, cLinha)
	Local aRet := .F.
	Local cAlias := GetNextAlias()
	Local cNum := GetValue(aReg, nPosNum, "S")
	Local cParcela := GetValue(aReg, nPosParcela, "S")
	Local cFornece := GetValue(aReg, nPosFornece, "S")
	Local cLoja := GetValue(aReg, nPosLoja, "S")
	// Local nValor := GetValue(aReg, nPosValor, "N")
	Local dDtBaixa := GetValue(aReg, nPosDtBx, "D")
	Local dDtSE1 := STOD("")
	Local aArea := GetArea()
	Local aAreaSE1 := SE1->(GetArea())
	Local nSE5Reg := 0

	BEGINSQL ALIAS cAlias
		SELECT
			SE1.R_E_C_N_O_ REG,
			SE1.E1_BAIXA BAIXA,
			SE1.E1_NUM NUM
		FROM %Table:SE1% AS SE1
		WHERE
			SE1.%NotDel%
			AND SE1.E1_NUM = %exp:cNum%
			AND SE1.E1_PARCELA = %exp:cParcela%
			AND SE1.E1_CLIENTE = %exp:cFornece%
			AND SE1.E1_LOJA = %exp:cLoja%
	ENDSQL

	If !(cAlias)->(Eof())
		dDtSE1 := STOD((cAlias)->BAIXA)
		If dDtBaixa != dDtSE1 .AND. !Empty(dDtSE1)
			aRet := {.T., (cAlias)->REG}

			SE1->(DbGoTo((cAlias)->REG))

			IncLog(@aErros, aReg, "Data de baixa errada na base de dados!", "Planilha: "+ DTOC(dDtBaixa) +" Protheus: " + DTOC(dDtSE1), cLinha, "Ajuste")

			nSE5Reg := GetSe5(dDtSE1, @aErros)
			If nSE5Reg > 0
				If !Empty(dDtBaixa) .AND. !Empty(dDtSE1)
					aAdd(aAjustes, {(cAlias)->REG, nSE5Reg, dDtBaixa})
				EndIf
			EndIf
		Else
			aRet := {.T., (cAlias)->REG}
		EndIf
	Else
		aRet := {.F., ""}
	EndIf
	(cAlias)->(DbCloseArea())

	SE1->(RestArea(aAreaSE1))
	RestArea(aArea)
Return aRet

User Function LOGIMPSE1()

	Local cTitAb01  := "Log de Importação"
	Local cArquivo  := GetTempPath()+CriaTrab(NIL,.F.)+'.xml'
	Local cAba01    := "Log"
	Local aColAba01 := {}
	Local nAtual := 0
	Local cAlias := GetNextAlias()
	Local aResps := {}
	Local aPergs := {}
	Local oFWMsExcel := Nil
	Local ID := GetLastId()
	Local cWhere := ''

	AADD( aColAba01, {"FILIAL"			,"C" } )
	AADD( aColAba01, {"ID"			    ,"C" } )
	AADD( aColAba01, {"PREFIXO"			,"C" } )
	AADD( aColAba01, {"NUM. TITULO" 	,"C" } )
	AADD( aColAba01, {"PARCELA"			,"C" } )
	AADD( aColAba01, {"TIPO"			,"C" } )
	AADD( aColAba01, {"FORNECEDOR"		,"C" } )
	AADD( aColAba01, {"LOJA"			,"C" } )
	AADD( aColAba01, {"VALOR"			,"M" } )
	AADD( aColAba01, {"EMISSAO"			,"D" } )
	AADD( aColAba01, {"VENCIMENTO"		,"D" } )
	AADD( aColAba01, {"VENC. REAL"		,"D" } )
	AADD( aColAba01, {"DATA/HORA IMP"	,"C" } )
	AADD( aColAba01, {"TIPO MOV."		,"C" } )
	AADD( aColAba01, {"AJUDA"			,"C" } )
	AADD( aColAba01, {"LINHA COMPLETA"	,"C" } )


	oFWMsExcel := FWMSExcelEx():New()
	oFWMsExcel:AddworkSheet(cAba01)    //adiciona a aba
	oFWMsExcel:AddTable(cAba01,cTitAb01)

	For nAtual := 1 to ( Len(aColAba01) ) //cria as colunas da aba
		If aColAba01[nAtual,2] == "C"
			oFWMsExcel:AddColumn(cAba01,cTitAb01,aColAba01[nAtual,1],1,1)
		ElseIf aColAba01[nAtual,2] == "D"
			oFWMsExcel:AddColumn(cAba01,cTitAb01,aColAba01[nAtual,1],1,4)
		ElseIf aColAba01[nAtual,2] == "N"
			oFWMsExcel:AddColumn(cAba01,cTitAb01,aColAba01[nAtual,1],1,2,.T.)
		ElseIf aColAba01[nAtual,2] == "M"
			oFWMsExcel:AddColumn(cAba01,cTitAb01,aColAba01[nAtual,1],1,3,.T.)
		EndIf
	Next

	aAdd(aPergs,{3,"Apenas ultimo processamento?",1,{"Sim","Não"},50,"",.F.})
	If Parambox(aPergs, "Leitura de arquivos CSV", aResps,,,,,,,,.T.,.T.)
		If aResps[1] == 1  //Apenas o ultimo processamento
			cWhere := "%"
			cWhere += "ZPX.ZPX_ID = '"+ID+"'"
			cWhere += "%"
		Else
			cWhere := "%"
			cWhere += "ZPX.ZPX_ID != ''"
			cWhere += "%"
		EndIf
	EndIf

	BEGINSQL ALIAS cAlias
		SELECT
			ZPX.*,
			ISNULL(CAST(CAST(ZPX.ZPX_HELP AS VARBINARY(8000)) AS VARCHAR(8000)),'') AS ZPX_HELP,
			ISNULL(CAST(CAST(ZPX.ZPX_LINHA AS VARBINARY(8000)) AS VARCHAR(8000)),'') AS ZPX_LINHA
		FROM %Table:ZPX% AS ZPX
		WHERE
			ZPX.%NotDel%
			AND %Exp:cWhere%
	ENDSQL

	While !(cAlias)->(Eof())
		aMensErros := {;
			(cAlias)->ZPX_FILIAL,;
			(cAlias)->ZPX_ID,;
			(cAlias)->ZPX_PREFIX,;
			(cAlias)->ZPX_NUM,;
			(cAlias)->ZPX_PARCEL,;
			(cAlias)->ZPX_TIPO,;
			(cAlias)->ZPX_FORNEC,;
			(cAlias)->ZPX_LOJA,;
			(cAlias)->ZPX_VALOR,;
			STOD((cAlias)->ZPX_EMISSA),;
			STOD((cAlias)->ZPX_VENC),;
			STOD((cAlias)->ZPX_VENCR),;
			(cAlias)->ZPX_HORA,;
			(cAlias)->ZPX_TPMV,;
			(cAlias)->ZPX_HELP,;
			(cAlias)->ZPX_LINHA}

		oFWMsExcel:AddRow(cAba01,cTitAb01,aMensErros)
		(cAlias)->(DbSkip())
	EndDo
	(cAlias)->(DbCloseArea())

	If oFWMsExcel:nRows > 0
		oFWMsExcel:Activate()
		oFWMsExcel:GetXMLFile(cArquivo)
		oExcel := MsExcel():New()
		oExcel:WorkBooks:Open(cArquivo)
		oExcel:SetVisible(.T.)
		oExcel:Destroy()
	EndIf

Return
Static Function IncLog(aErros, aReg, cMsgErro, cCpoErro, cLinha, cTpMv)
	Local cAlias := GetNextAlias()
	Local cPrefix := aReg[nPosPrefixo]
	Local cNum := aReg[nPosNum]
	Local cParcela := aReg[nPosParcela]
	Local cTipo := aReg[nPosTipo]
	Local cFornece := aReg[nPosFornece]
	Local cLojaFor := aReg[nPosLoja]
	Local nValor := GetValue(aReg, nPosValor, 'N')
	Local dEmissao := GetValue(aReg, nPosEmissao, 'D')
	Local dVencto := GetValue(aReg, nPosVencto, 'D')
	Local dVencRea := GetValue(aReg, nPosVencRea, 'D')

	Default cCpoErro := ""
	Default cLinha := ""

	BEGINSQL ALIAS cAlias
		SELECT
			ZPX.R_E_C_N_O_ REG
		FROM %Table:ZPX% AS ZPX
		WHERE
			ZPX.%NotDel%
			AND ZPX.ZPX_ID = %exp:ID%
			AND ZPX.ZPX_PREFIX = %exp:cPrefix%
			AND ZPX.ZPX_NUM = %exp:cNum%
			AND ZPX.ZPX_PARCEL = %exp:cParcela%
			AND ZPX.ZPX_TIPO = %exp:cTipo%
			AND ZPX.ZPX_FORNEC = %exp:cFornece%
			AND ZPX.ZPX_LOJA = %exp:cLojaFor%
			AND ZPX.ZPX_TPMV = %exp:cTpMv%
	ENDSQL

	If (cAlias)->(Eof())
		If Reclock("ZPX", .T.)
			REPLACE ZPX_ID WITH ID
			REPLACE ZPX_PREFIX WITH cPrefix
			REPLACE ZPX_NUM WITH cNum
			REPLACE ZPX_PARCEL WITH cParcela
			REPLACE ZPX_TIPO WITH cTipo
			REPLACE ZPX_FORNEC WITH cFornece
			REPLACE ZPX_LOJA WITH cLojaFor
			REPLACE ZPX_VALOR WITH nValor
			REPLACE ZPX_EMISSA WITH dEmissao
			REPLACE ZPX_VENC WITH dVencto
			REPLACE ZPX_VENCR WITH dVencRea
			REPLACE ZPX_HORA WITH TIMESTAMP
			REPLACE ZPX_TPMV WITH cTpMv
			REPLACE ZPX_HELP WITH cMsgErro + "/" + cCpoErro
			REPLACE ZPX_LINHA WITH cLinha
			ZPX->(MsUnlock())
		EndIf
	EndIf
	(cAlias)->(DbCloseArea())
Return

Static Function GetSe5(dDtSE1, aErros)
	Local nRet := 0
	Local cAlias := GetNextAlias()

	BEGINSQL ALIAS cAlias
		SELECT
			SE5.R_E_C_N_O_ REG
		FROM %Table:SE5% AS SE5
		WHERE
			SE5.%NotDel%
			AND SE5.E5_FILIAL = %xFilial:SE5%
			AND SE5.E5_TIPODOC = 'BA'
			AND SE5.E5_PREFIXO = %exp:SE1->E1_PREFIXO%
			AND SE5.E5_NUMERO = %exp:SE1->E1_NUM%
			AND SE5.E5_PARCELA = %exp:SE1->E1_PARCELA%
			AND SE5.E5_TIPO = %exp:SE1->E1_TIPO%
			// AND SE5.E5_DATA = %exp:DTOS(dDtSE1)%
			AND SE5.E5_CLIFOR = %exp:SE1->E1_CLIENTE%
			AND SE5.E5_LOJA = %exp:SE1->E1_LOJA%
	ENDSQL

	If !(cAlias)->(Eof())
		nRet := (cAlias)->REG
	EndIf
	(cAlias)->(DbCloseArea())

Return nRet


Static Function Ajustar(aAjustes)
	Local lContinua := Nil
	Local nIx := 0

	lContinua := FWAlertNoYes("Deseja ajustar a data das baixas?", "Ajuste Baixas")

	If lContinua
		For nIx := 1 To Len(aAjustes)
			dDtBaixa := aAjustes[nIx, 3]
			SE1->(DbGoTo(aAjustes[nIx, 1]))
			Reclock("SE1")
			REPLACE E1_BAIXA WITH dDtBaixa
			REPLACE E1_MOVIMEN WITH dDtBaixa
			MsUnlock()

			SE5->(DbGoTo(aAjustes[nIx, 2]))
			Reclock("SE5")
			REPLACE E5_DATA WITH dDtBaixa
			REPLACE E5_DTDIGIT WITH dDtBaixa
			REPLACE E5_DTDISPO WITH dDtBaixa
			MsUnlock()
		Next
	EndIf
Return

Static Function GetLastId()
	Local cLastId := ''
	Local cAlias := GetNextAlias()

	BEGINSQL ALIAS cAlias
		SELECT TOP 1
			ZPX.ZPX_ID ID
		FROM %Table:ZPX% AS ZPX
		WHERE
			ZPX.%NotDel%
		ORDER BY
			ZPX_HORA DESC
	ENDSQL

	If !(cAlias)->(Eof())
		cLastId := (cAlias)->ID
	EndIf
	(cAlias)->(DbCloseArea())

Return cLastId
