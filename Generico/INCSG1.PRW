#Include 'Protheus.ch'
#Include 'TBICONN.ch'

User Function INCSG1(cRev)
	Processa({|| fProcessa(cRev)}, "Aguarde...", "Executando rotina...")
Return

Static Function fProcessa(cRev)
	Local aCab := {}
	Local aGets := {}
	Local aItens := {}
	Local cPrdPai := ""
	Local cComp := ""
	Local aDados := {}
	Local nIx := nPosCod := nPosComp := nPosQtd := 0
	Local aResps := {}
	Local aPergs := {}

	PRIVATE lMsErroAuto := .F.
	aDados := U_LERCSV()
	nLinhas = ContProds(aDados)

	aAdd(aPergs,{1,"Revisão",Space(3),"","","","",0,.T.}) //Input
	If Parambox (aPergs,'Cadastro de Estruturas via CSV',aResps,,,,,,,,.T.,.T.)
		If Empty(cRev)
			cRev := aResps[1]
		EndIf

		BEGIN TRANSACTION
			For nIx := 2 To Len(aDados)
				If nIx == 2
					nPosCod := AScan(aDados[nIx], {|x| ALLTRIM(x) == "G1_COD"})
					nPosComp := AScan(aDados[nIx], {|x| ALLTRIM(x) == "G1_COMP"})
					nPosQtd := AScan(aDados[nIx], {|x| ALLTRIM(x) == "G1_QUANT"})

					If nPosCod == 0 .OR. nPosComp == 0 .OR. nPosQtd == 0
						FwAlertError("Um dos seguintes cabeçalhos não foi encontrado: G1_COD, G1_COMP, G1_QUANT. Atualização não foi realizada!!!")
						Return
					EndIf
				Else
					If Empty(cPrdPai) .AND. !Empty(aDados[nIx, nPosCod])
						cPrdPai := aDados[nIx, nPosCod]
						If cPrdPai == aDados[nIx, nPosComp]
							LOOP
						EndIf
					EndIf

					If cPrdPai == aDados[nIx, nPosCod]
						cComp := aDados[nIx, nPosComp]
						nQtd := Val(REPLACE(aDados[nIx, nPosQtd],",","."))
						aCab :={{"G1_COD", cPrdPai,NIL},;
							{"G1_QUANT"                  , 1  , NIL},;
							{"NIVALT"                    , "S", NIL},;
							{"ATUREVSB1"                 , "N", NIL}}

						aadd(aGets, {"G1_COD"   , cPrdPai         , NIL})
						aadd(aGets, {"G1_COMP"  , cComp           , NIL})
						aadd(aGets, {"G1_TRT"   , Space(3)        , NIL})
						aadd(aGets, {"G1_QUANT" , nQtd            , NIL})
						aadd(aGets, {"G1_PERDA" , 0               , NIL})
						aadd(aGets, {"G1_INI"   , dDataBase       , NIL})
						aadd(aGets, {"G1_FIM"   , CTOD("31/12/49"), NIL})
						aadd(aGets, {"G1_REVINI", cRev           , NIL})
						aadd(aGets, {"G1_REVFIM", cRev           , NIL})
						aadd(aItens,aGets)
						aGets := {}
						If nIx == Len(aDados)
							Inclusao(aCab, aItens, cPrdPai, cComp, Space(3), cRev, nLinhas)
						EndIf
					Else
						Inclusao(aCab, aItens, cPrdPai, cComp, Space(3), cRev, nLinhas)
						cPrdPai := aDados[nIx, nPosCod]
						cComp := aDados[nIx, nPosComp]
						aGets := {}
						aItens := {}

						If cPrdPai == aDados[nIx, nPosComp]
							LOOP
						EndIf
					EndIf
				EndIf
			Next
		END TRANSACTION
	Else
		FwAlertError("Operação cancelada pelo usuario")
	Endif
Return

Static Function Inclusao(aCab,aItens, cPrdPai, cComp, cTrt, cRev, nLinhas)
	Local aArea := GetArea()
	Local aAreaSG1 := SG1->(GetArea())
	Local aAreaSB1 := SB1->(GetArea())
	Local aExisteStr := {}
	Local nOpc := 0
	Local nAtual := 0

	ProcRegua(nLinhas)

	SB1->(DbSetOrder(1))//B1_FILIAL + B1_COD
	If SB1->(MsSeek(xFilial("SB1")+cPrdPai))
		nAtual++
		IncProc("Incluindo estrutura do produto: " + SB1->B1_DESC)
		If RECLOCK( "SB1", .F. )
			REPLACE B1_REVATU With cRev
			MSUNLOCK()

			SG1->(DbSetOrder(1)) //G1_FILIAL + G1_COD + G1_COMP + G1_TRT
			aExisteStr := ExisteStr(cPrdPai, cComp, cTrt, cRev)
			If !aExisteStr[1]
				nOpc := aExisteStr[2]
				MSExecAuto({|x,y,z| mata200(x,y,z)},aCab,aItens,nOpc)
				If lMsErroAuto
					MostraErro()
				EndIf
			Else
				MSExecAuto({|x,y,z| mata200(x,y,z)},aCab,aItens,5) //Exclui estrutura existente
				If lMsErroAuto
					MostraErro()
				Else
					MSExecAuto({|x,y,z| mata200(x,y,z)},aCab,aItens,3) //Inclue conforme esta na planilha
					If lMsErroAuto
						MostraErro()
					EndIf
				EndIf
			EndIf
		EndIf
	EndIf



	SB1->(RestArea(aAreaSB1))
	SG1->(RestArea(aAreaSG1))
	RestArea(aArea)
Return

Static Function ExisteStr(cPrdPai, cComp, cTrt, cRev)
	Local lRet := .F.
	Local cAliasRev := GetNextAlias()
	Local cAlias := GetNextAlias()
	Local nOpc := 3

	BEGINSQL ALIAS cAliasRev
		SELECT
			SG1.*
		FROM %Table:SG1% AS SG1
		WHERE
			SG1.%NotDel%
			AND SG1.G1_COD = %exp:cPrdPai%
			AND SG1.G1_TRT = %exp:cTrt%
			AND SG1.G1_REVINI = %exp:cRev%
	ENDSQL

	If !(cAliasRev)->(Eof())
		lRet := .T.
		nOpc := 3
	EndIf
	(cAliasRev)->(DbCloseArea())

	BEGINSQL ALIAS cAlias
		SELECT
			SG1.*
		FROM %Table:SG1% AS SG1
		WHERE
			SG1.%NotDel%
			AND SG1.G1_COD = %exp:cPrdPai%
			AND SG1.G1_TRT = %exp:cTrt%
	ENDSQL

	If !(cAlias)->(Eof())
		lRet := .F.
		nOpc := 4
	EndIf
	(cAlias)->(DbCloseArea())
Return {lRet, nOpc}

Static Function ContProds(aLista)
	Local nQtd := 0
	Local nPos := 0
	Local nCont := 0
	Local aNovo := {}

	For nCont := 3 To Len(aLista)
		cAux := aLista[nCont, 1]
		If !Empty(cAux)
			nPos := aScan(aNovo,{|x| x[1] == cAux})
			If nPos == 0
				AADD( aNovo, {cAux} )
			EndIf
		EndIf
	Next
	nQtd := Len(aNovo)
Return nQtd
