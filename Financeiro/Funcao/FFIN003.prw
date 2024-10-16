#INCLUDE "Protheus.CH"
#INCLUDE "Totvs.ch"
#INCLUDE "TOPCONN.ch"
#INCLUDE "TBICONN.ch"
#INCLUDE "FWMVCDef.CH"

//--------------------------------------------------------------------------------------------------------------------
/*/{PROTHEUS.DOC} FFIN003
FUNÇÃO FFIN003- Titulos Liquidados (Legado)
@VERSION PROTHEUS 12
@SINCE 15/10/2024
/*/
User Function FFIN003()
	Local aArea := FWGetArea()
	Local oBrowse

	oBrowse := FWMBrowse():New()
	oBrowse:SetDescription("Titulos Liquidados (Legado)")
	//oBrowse:SetAmbiente(.F.)
	//oBrowse:SetWalkThru(.F.)
	//oBrowse:SetMenuDef('FFIN003')
	oBrowse:SetAlias('ZZX')
	//oBrowse:DisableDetails()
	//oBrowse:SetFixedBrowse(.T.)

	oBrowse:AddLegend("ZZX_STATUS == '  '", 'BR_BRANCO'  , 'Apto a processar')
	oBrowse:AddLegend("ZZX_STATUS == 'EI'", 'BR_VERMELHO', 'Erro na Inclusao do Titulo')
	oBrowse:AddLegend("ZZX_STATUS == 'EB'", 'BR_LARANJA' , 'Erro na Baixa do Titulo')
	oBrowse:AddLegend("ZZX_STATUS == 'TI'", 'BR_AZUL'    , 'Titulo Incluido')
	oBrowse:AddLegend("ZZX_STATUS == 'BP'", 'BR_CINZA'   , 'Titulo Baixado Parcial')
	oBrowse:AddLegend("ZZX_STATUS == 'OK'", 'BR_VERDE'   , 'Titulo Incluido e Baixado')

	oBrowse:Activate()

	FWRestArea(aArea)

Return

/*---------------------------------------------------------------------*
 | Func:  MenuDef                                                      |
 | Desc:  Criação do Menu da rotina                                    |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function MenuDef()
	Local aRot := {}
	
	ADD OPTION aRot TITLE 'Importar'   ACTION 'U_IMPFIN003'   	OPERATION 3 ACCESS 0
	ADD OPTION aRot TITLE 'Processar'  ACTION 'U_PROCES003'   	OPERATION 4 ACCESS 0
	ADD OPTION aRot TITLE 'Visualizar' ACTION 'VIEWDEF.FFIN003' OPERATION 2 ACCESS 0
	ADD OPTION aRot TITLE 'Alterar'    ACTION 'VIEWDEF.FFIN003' OPERATION 4 ACCESS 0
	ADD OPTION aRot TITLE 'Excluir'    ACTION 'U_EXCFIN003'   	OPERATION 5 ACCESS 0

Return aRot

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ModelDef()
	Local oModel    := Nil
	Local oStPai   	:= FWFormStruct( 1, 'ZZX')

	oModel := MPFormModel():New("FFIN003M" , , , )
	oModel:SetDescription(OemtoAnsi("Titulos Liquidados (Legado)") )
	oModel:AddFields('ZZXMASTER',/*cOwner*/,oStPai)
	oModel:SetPrimaryKey({})

Return oModel

/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ViewDef()
	Local oView     := Nil
	Local oModel    := FWLoadModel('FFIN003')
	Local oStPai   	:= FWFormStruct( 2, 'ZZX')
	
	oView := FWFormView():New()
	oView:SetModel(oModel)
	oView:AddField('VIEW_CAB',oStPai,'ZZXMASTER')
	oView:CreateHorizontalBox('CABEC',100)
	oView:SetOwnerView('VIEW_CAB','CABEC')
	oView:EnableTitleView('VIEW_CAB','Titulos Liquidados')

Return oView

/*---------------------------------------------------------------------*
 | Func:  IMPFIN003()                                                  |
 | Desc:  Realiza a importação do arquivo                              |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/

User Function IMPFIN003()
	Processa({|| fnImport()}, "Importando arquivo...")
Return 

Static Function fnImport()

	Local cArq 		:= tFileDialog( "Arquivos de texto (*.txt;*.csv)",,,, .F., /*GETF_MULTISELECT*/)
	Local nAtu 		:= 0
    Local nFim  	:= 0
	Local cLinha 	:= ""
	Local aRegistro := {}
	Local cNumImp   := fPrxNumZZX()
	Local cTimeHr   := FWTimeStamp(2)
	Local cHelp     := ""
	Local nDia      := 0
	Local nMes      := 0
	Local nAno      := 0

	If !File(cArq)
		FwAlertError("Não foi possível importar o arquivo, por favor tente novamente.","Importar Arquivo")
        Return
    EndIf

	FT_FUSE(cArq)
    nFim := FT_FLASTREC()
    ProcRegua(nFim)
    FT_FGOTOP()

	DBSelectArea("SA1")
	DBSelectArea("ZZX")

	While !FT_FEOF()
            
        nAtu++
        IncProc("Linha " + cValToChar(nAtu) + " de " + cValToChar(nFim) + "...")

		If nAtu == 1
			FT_FSKIP()
		EndIF 

		cLinha := FT_FREADLN()
                
        If !Empty(cLinha)
            aRegistro := {}
            aRegistro := Separa(cLinha,";",.T.)
            
			If Len(aRegistro) == 24

				aRegistro[13] := StrTran(aRegistro[13],".","")
				aRegistro[13] := StrTran(aRegistro[13],",",".")
				aRegistro[13] := Val(aRegistro[13])

				nDia := Val(IIF(At("/",aRegistro[10])+1 == RAt("/",aRegistro[10])-1, SubStr(aRegistro[10],At("/",aRegistro[10])+1,1), SubStr(aRegistro[10],At("/",aRegistro[10])+1,2)))
				nMes := Val(SubStr(aRegistro[10],1,At("/",aRegistro[10])-1))
				nAno := Val(SubStr(aRegistro[10],Rat("/",aRegistro[10])+1))
				aRegistro[10] := CToD(StrZero(nDia,2)+"/"+StrZero(nMes,2)+"/"+StrZero(nAno,4))

				nDia := Val(IIF(At("/",aRegistro[11])+1 == RAt("/",aRegistro[11])-1, SubStr(aRegistro[11],At("/",aRegistro[11])+1,1), SubStr(aRegistro[11],At("/",aRegistro[11])+1,2)))
				nMes := Val(SubStr(aRegistro[11],1,At("/",aRegistro[11])-1))
				nAno := Val(SubStr(aRegistro[11],Rat("/",aRegistro[11])+1))
				aRegistro[11] := CToD(StrZero(nDia,2)+"/"+StrZero(nMes,2)+"/"+StrZero(nAno,4))

				nDia := Val(IIF(At("/",aRegistro[12])+1 == RAt("/",aRegistro[12])-1, SubStr(aRegistro[12],At("/",aRegistro[12])+1,1), SubStr(aRegistro[12],At("/",aRegistro[12])+1,2)))
				nMes := Val(SubStr(aRegistro[12],1,At("/",aRegistro[12])-1))
				nAno := Val(SubStr(aRegistro[12],Rat("/",aRegistro[12])+1))
				aRegistro[12] := CToD(StrZero(nDia,2)+"/"+StrZero(nMes,2)+"/"+StrZero(nAno,4))
				
				nDia := Val(IIF(At("/",aRegistro[15])+1 == RAt("/",aRegistro[15])-1, SubStr(aRegistro[15],At("/",aRegistro[15])+1,1), SubStr(aRegistro[15],At("/",aRegistro[15])+1,2)))
				nMes := Val(SubStr(aRegistro[15],1,At("/",aRegistro[15])-1))
				nAno := Val(SubStr(aRegistro[15],Rat("/",aRegistro[15])+1))
				aRegistro[15] := CToD(StrZero(nDia,2)+"/"+StrZero(nMes,2)+"/"+StrZero(nAno,4))
				
				If SA1->(MSSeek(xFilial("SA1") + aRegistro[07] ))
					If AllTrim(aRegistro[08]) <> SA1->A1_LOJA
						aRegistro[09] := aRegistro[08]
						aRegistro[08] := SA1->A1_LOJA
					EndIF 
				EndIF 

				IF ! ZZX->(MSSeek(xFilial("ZZX") + Pad(aRegistro[01],FWTamSX3("ZZX_FILMOV")[1]) + ;
												   Pad(aRegistro[02],FWTamSX3("ZZX_PREFIX")[1]) + ;
												   Pad(aRegistro[03],FWTamSX3("ZZX_NUM")[1]   ) + ;
												   Pad(aRegistro[04],FWTamSX3("ZZX_PARCEL")[1]) + ;
												   Pad(aRegistro[05],FWTamSX3("ZZX_TIPO")[1]  ) ))

					RecLock("ZZX",.T.)
						ZZX->ZZX_FILIAL := xFilial("ZZX")
						ZZX->ZZX_FILMOV := Pad(aRegistro[01],FWTamSX3("ZZX_FILMOV")[1])
						ZZX->ZZX_PREFIX := Pad(aRegistro[02],FWTamSX3("ZZX_PREFIX")[1])
						ZZX->ZZX_NUM    := Pad(aRegistro[03],FWTamSX3("ZZX_NUM")[1]   )
						ZZX->ZZX_PARCEL := Pad(aRegistro[04],FWTamSX3("ZZX_PARCEL")[1])
						ZZX->ZZX_TIPO   := Pad(aRegistro[05],FWTamSX3("ZZX_TIPO")[1]  )
						ZZX->ZZX_NATURE := Pad(aRegistro[06],FWTamSX3("ZZX_NATURE")[1])
						ZZX->ZZX_CLIENT := Pad(aRegistro[07],FWTamSX3("ZZX_CLIENT")[1])
						ZZX->ZZX_LOJA   := Pad(aRegistro[08],FWTamSX3("ZZX_LOJA")[1]  )
						ZZX->ZZX_LJOLD  := Pad(aRegistro[09],FWTamSX3("ZZX_LJOLD")[1] )
						ZZX->ZZX_EMISSA := aRegistro[10]
						ZZX->ZZX_VENCTO := aRegistro[11]
						ZZX->ZZX_VENCRE := aRegistro[12]
						ZZX->ZZX_VALOR  := aRegistro[13]
						ZZX->ZZX_VEND1  := Pad(aRegistro[14],FWTamSX3("ZZX_VEND1")[1] )
						ZZX->ZZX_BAIXA  := aRegistro[15]
						ZZX->ZZX_CCUSTO := Pad(aRegistro[16],FWTamSX3("ZZX_CCUSTO")[1])
						ZZX->ZZX_NUMBCO := Pad(aRegistro[17],FWTamSX3("ZZX_NUMBCO")[1])
						ZZX->ZZX_NSUTEF := Pad(aRegistro[18],FWTamSX3("ZZX_NSUTEF")[1])
						ZZX->ZZX_CARTAU := Pad(aRegistro[19],FWTamSX3("ZZX_CARTAU")[1])
						ZZX->ZZX_HIST   := Pad(aRegistro[20],FWTamSX3("ZZX_HIST")[1]  )
						ZZX->ZZX_BANCO  := Pad(aRegistro[21],FWTamSX3("ZZX_BANCO")[1] )
						ZZX->ZZX_AGENCI := Pad(aRegistro[22],FWTamSX3("ZZX_AGENCI")[1])
						ZZX->ZZX_CONTA  := Pad(aRegistro[23],FWTamSX3("ZZX_CONTA")[1] )
						ZZX->ZZX_DVCTA  := Pad(aRegistro[24],FWTamSX3("ZZX_DVCTA")[1] )
						ZZX->ZZX_LINARQ := cValToChar(nAtu)
						ZZX->ZZX_TOTLIN := cValToChar(nFim)
						ZZX->ZZX_DTIMP  := dDataBase
						ZZX->ZZX_USRIMP := cUserName
						ZZX->ZZX_ARQUIV := cArq
						ZZX->ZZX_IMPORT := cNumImp
					ZZX->(MsUnlock())	
				Else
					cHelp := "Linha " + cValToChar(nAtu) + "- O titulo ja existe na tabela ZZX"
					fnGrvLog(aRegistro,cHelp,cLinha,cTimeHr,"IMPORTACAO")
				EndIF 
				
			Else
				cHelp := "Linha " + cValToChar(nAtu) + "- O registro possui menos colunas que o previsto"
				fnGrvLog(aRegistro,cHelp,cLinha,cTimeHr,"IMPORTACAO")
			EndIF 

        EndIF 
	FT_FSKIP()
    EndDo

    FT_FUSE()
Return

/*---------------------------------------------------------------------*
 | Func:  fPrxNumZZX                                                   |
 | Desc:  Retorna o próximo número para a tabela ZZX                   |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/

Static Function fPrxNumZZX()

	Local cRet := StrZero(0,FWTamSX3("ZZX_IMPORT")[1])
	Local cQry := ''
	Local cAli := GetNextAlias()

	cQry+= " SELECT DISTINCT MAX(ZZX_IMPORT) ZZX_IMPORT "
	cQry+= " FROM " + RetSqlTab('ZZX')
	cQry+= " WHERE " + RetSqlCond('ZZX')
	cQry:= ChangeQuery(cQry)
	If Select(cAli) <> 0
		(cAli)->(DbCloseArea())
	EndIf
	dbUseArea(.T.,'TOPCONN', TCGenQry(,,cQry),cAli, .F., .T.)

	If (cAli)->(!Eof())
		cRet:= (cAli)->ZZX_IMPORT
	EndIf

	If Select(cAli) <> 0
		(cAli)->(DbCloseArea())
	EndIf

	cRet:= Soma1(cRet)

Return cRet

/*---------------------------------------------------------------------*
 | Func:  fnGrvLog                                                     |
 | Desc:  Grava LOG na ZPX                                             |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function fnGrvLog(aRegistro,cHelp,cLinha,cTimeHr,cTpMov)
	Reclock("ZPX", .T.)
		REPLACE ZPX_PREFIX 	WITH Pad(aRegistro[02],FWTamSX3("ZPX_PREFIX")[1])
		REPLACE ZPX_NUM 	WITH Pad(aRegistro[03],FWTamSX3("ZPX_NUM")[1]   )
		REPLACE ZPX_PARCEL 	WITH Pad(aRegistro[04],FWTamSX3("ZPX_PARCEL")[1])
		REPLACE ZPX_TIPO	WITH Pad(aRegistro[05],FWTamSX3("ZPX_TIPO")[1]  )
		REPLACE ZPX_FORNEC 	WITH Pad(aRegistro[07],FWTamSX3("ZPX_FORNEC")[1])
		REPLACE ZPX_LOJA 	WITH Pad(aRegistro[08],FWTamSX3("ZPX_LOJA")[1]  )
		REPLACE ZPX_HELP 	WITH cHelp
		REPLACE ZPX_LINHA 	WITH cLinha
		REPLACE ZPX_HORA 	WITH cTimeHr
		REPLACE ZPX_TPMV 	WITH cTpMov
	ZPX->(MsUnlock())
Return

/*---------------------------------------------------------------------*
 | Func:  EXCFIN003()                                                  |
 | Desc:  Exclui registros da tabela ZZX                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/

User Function EXCFIN003()
	Local oPanel   := Nil
	Local oDialog  := Nil
	Local cLote    := Space(FWTamSX3("ZZX_IMPORT")[1])

	IF FWAlertYesNo("Deseja realizar exclusão em LOTE ?","Exclusão de Registros")
		oDialog := FWDialogModal():New()
		oDialog:SetBackground( .T. ) 
		oDialog:SetTitle( 'Exclusão de Registros em LOTE' )
		oDialog:SetSize( 100, 150 )
		oDialog:EnableFormBar( .T. )
		oDialog:SetCloseButton( .T. )
		oDialog:SetEscClose( .T. )
		oDialog:CreateDialog()
		oDialog:CreateFormBar()
		oDialog:addCloseButton(Nil, "Fechar")
		oDialog:addCloseButton(Nil, "Confirmar")
		oPanel := oDialog:GetPanelMain()
		oTSay  := TSay():New(10,5,{|| "LOTE: "},oPanel,,,,,,.T.,,,50,70,,,,,,.T.)
        oCombo := TComboBox():New(29,28,{|u|iif(PCount()>0,cLote:=u,cLote)},aTab,100,20,oDlg,,{||},,,,.T.,,,,,,,,,'cLote')
		oDialog:Activate()
	Else
		FWExecView("Exclusão","FFIN003",5,,{|| .T.},,,)
	EndIF 

Return

/*---------------------------------------------------------------------*
 | Func:  PROCES003()                                                  |
 | Desc:  Realiza a inclusão e baixa dos títulos registrados na ZZX.   |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/

User Function PROCES003()
	Processa({|| fnProcess()}, "Processando registros...")
Return 

Static Function fnProcess()
	Local aArea    := FWGetArea()
	Local aAreaSE1 := SE1->(FWGetArea())
	Local aPergs   := {}
	Local aSelOpc  := {"1=Inclusao e Baixa", "2=Apenas Inclusao", "3=Apenas Baixa"}
	Local nSelect  := 1
	Local cLoteDe  := Space(FWTamSX3("ZZX_IMPORT")[01])
	Local cLoteAt  := Space(FWTamSX3("ZZX_IMPORT")[01])
	Local cPrefDe  := Space(FWTamSX3("ZZX_PREFIX")[01])
	Local cPrefAt  := Space(FWTamSX3("ZZX_PREFIX")[01])
	Local cNumDe   := Space(FWTamSX3("ZZX_IMPORT")[01])
	Local cNumAt   := Space(FWTamSX3("ZZX_IMPORT")[01])
	Local cClieDe  := Space(FWTamSX3("ZZX_CLIENT")[01])
	Local cClieAt  := Space(FWTamSX3("ZZX_CLIENT")[01])
	Local cQry 	   := ''
	Local nAtual   := 0
	Local nFim     := 0

	Private _cAlias := GetNextAlias()

	aAdd(aPergs, {1, "Lote de    :", cLoteDe,  "", ".T.", "", ".T.", 80,  .F.})
	aAdd(aPergs, {1, "Lote ate   :", cLoteAt,  "", ".T.", "", ".T.", 80,  .T.})
	aAdd(aPergs, {1, "Prefixo de :", cPrefDe,  "", ".T.", "", ".T.", 80,  .F.})
	aAdd(aPergs, {1, "Prefixo ate:", cPrefAt,  "", ".T.", "", ".T.", 80,  .T.})
	aAdd(aPergs, {1, "Numero de  :", cNumDe ,  "", ".T.", "", ".T.", 80,  .F.})
	aAdd(aPergs, {1, "Numero ate :", cNumAt ,  "", ".T.", "", ".T.", 80,  .T.})
	aAdd(aPergs, {1, "Cliente de :", cClieDe,  "", ".T.", "SA1", ".T.", 80,  .F.})
	aAdd(aPergs, {1, "Cliente ate:", cClieAt,  "", ".T.", "SA1", ".T.", 80,  .T.})
	aAdd(aPergs, {2, "Processar:", nSelect, aSelOpc, 80, ".T.", .T.})

   	If !ParamBox(aPergs ,"Informe os dados")
    	FWAlertWarning("Processo cancelado pelo usuário.","HELPFINF003")
		Return		
	EndIF

	IF ValType(MV_PAR09) == "C"
		MV_PAR09 := Val(MV_PAR09)
	EndIF

	cQry := " SELECT * "
	cQry += " FROM " + RetSqlTab('ZZX')
	cQry += " WHERE D_E_L_E_T_ <> '*' "
	cQry += " 	AND ZZX_IMPORT BETWEEN '" + MV_PAR01 + "' AND '" + MV_PAR02 + "' "
	cQry += " 	AND ZZX_PREFIX BETWEEN '" + MV_PAR03 + "' AND '" + MV_PAR04 + "' "
	cQry += " 	AND ZZX_NUM    BETWEEN '" + MV_PAR05 + "' AND '" + MV_PAR06 + "' "
	cQry += " 	AND ZZX_CLIENT BETWEEN '" + MV_PAR07 + "' AND '" + MV_PAR08 + "' "
	Do Case
		Case MV_PAR09 == 1
			cQry += " AND ZZX_STATUS IN ('','EI','TI') "
		Case  MV_PAR09 == 2
			cQry += " AND ZZX_STATUS IN ('','EI') "
		Case  MV_PAR09 == 3
			cQry += " AND ZZX_STATUS IN ('TI','EB') "
	End Case 
	cQry := ChangeQuery(cQry)
	If Select(_cAlias) <> 0
		(_cAlias)->(DbCloseArea())
	EndIf
	dbUseArea(.T.,'TOPCONN', TCGenQry(,,cQry),_cAlias, .F., .T.)
	Count To nFim
	ProcRegua(nFim)

	(_cAlias)->(DbGoTop())

	DBSelectArea("SE1")
	SE1->(dbSetOrder(2))

	While (_cAlias)->(!Eof())

		nAtual++
        IncProc("Registro " + cValToChar(nAtual) + " de " + cValToChar(nFim) + "...")

		IF MV_PAR09 == 1 .Or. MV_PAR09 == 2
			IF ! SE1->(MSSeek( Pad((_cAlias)->ZZX_FILMOV, FWTamSX3("E1_FILIAL")[01]) + ;
							   Pad((_cAlias)->ZZX_CLIENT, FWTamSX3("E1_CLIENTE")[01]) + ;
							   Pad((_cAlias)->ZZX_LOJA  , FWTamSX3("E1_LOJA")[01]) + ;
							   Pad((_cAlias)->ZZX_PREFIX, FWTamSX3("E1_PREFIXO")[01]) + ;
							   Pad((_cAlias)->ZZX_NUM   , FWTamSX3("E1_NUM")[01]) + ;
							   Pad((_cAlias)->ZZX_PARCEL, FWTamSX3("E1_PARCELA")[01]) + ;
							   Pad((_cAlias)->ZZX_TIPO  , FWTamSX3("E1_TIPO")[01]) ))
				
				fnIncTit() //Inclui o Titulo
			Else
				IF Empty(SE1->E1_SALDO)
					DBGoTo((_cAlias)->R_E_C_N_O_)
					RecLock("ZZX",.F.)
						ZZX_USRPRO := cUserName
						ZZX_DTPROS := dDataBase
						ZZX_ERRBX  := ""
						ZZX_STATUS := 'OK'
					ZZX->(MsUnlock())	
				ElseIF SE1->E1_SALDO < SE1->E1_VALOR
					DBGoTo((_cAlias)->R_E_C_N_O_)
					RecLock("ZZX",.F.)
						ZZX_USRPRO := cUserName
						ZZX_DTPROS := dDataBase
						ZZX_ERRBX  := ""
						ZZX_STATUS := 'BP'
					ZZX->(MsUnlock())
				EndIF 				
			EndIF 
		ElseIF MV_PAR09 == 3
			fnBXTit() //Baixa o Titulo
		EndIF 
	
	(_cAlias)->(DbSkip())
	End

	If Select(_cAlias) <> 0
		(_cAlias)->(DbCloseArea())
	EndIf

	FWRestArea(aAreaSE1)
	FWRestArea(aArea)
Return

/*--------------------------------------------------------------------------*
 | Func:  fnIncTit()                                                        |
 | Desc:  Faz a inclusão/baixa dos registrados da ZZX via ExecAuto FINA040  |
 | Obs.:  /                                                                 |
 *-------------------------------------------------------------------------*/

Static Function fnIncTit()
	Local nOpc := 3
	Local nCount := 0
	Local aTitInc := {}
	Local aErroAuto := {}
	Local cLogErro := ""
	Local cFilAux := cFilAnt 

	Private lMsErroAuto := .F.
	Private lAutoErrNoFile := .T.

	cFilAnt := (_cAlias)->ZZX_FILMOV

	aAdd(aTitInc,{ "E1_PREFIXO"  , (_cAlias)->ZZX_PREFIX  		, NIL })
	aAdd(aTitInc,{ "E1_NUM"      , (_cAlias)->ZZX_NUM     		, NIL })
	aAdd(aTitInc,{ "E1_PARCELA"  , (_cAlias)->ZZX_PARCEL  		, NIL })
	aAdd(aTitInc,{ "E1_TIPO"     , (_cAlias)->ZZX_TIPO    		, NIL })
	aAdd(aTitInc,{ "E1_NATUREZ"  , (_cAlias)->ZZX_NATURE  		, NIL })
	aAdd(aTitInc,{ "E1_CLIENTE"  , (_cAlias)->ZZX_CLIENT  		, NIL })
	aAdd(aTitInc,{ "E1_LOJA"     , (_cAlias)->ZZX_LOJA    		, NIL })
	aAdd(aTitInc,{ "E1_CCUSTO"   , (_cAlias)->ZZX_CCUSTO  		, NIL })
	aAdd(aTitInc,{ "E1_EMISSAO"  , SToD((_cAlias)->ZZX_EMISSA)  , NIL })
	aAdd(aTitInc,{ "E1_VENCTO"   , SToD((_cAlias)->ZZX_VENCTO)  , NIL })
	aAdd(aTitInc,{ "E1_VENCREA"  , SToD((_cAlias)->ZZX_VENCRE)  , NIL })
	aAdd(aTitInc,{ "E1_VALOR"    , (_cAlias)->ZZX_VALOR  		, NIL })
	aAdd(aTitInc,{ "E1_HIST"     , (_cAlias)->ZZX_HIST  		, NIL })
	aAdd(aTitInc,{ "E1_NUMBCO"   , (_cAlias)->ZZX_NUMBCO  		, NIL })
	aAdd(aTitInc,{ "E1_NSUTEF"   , (_cAlias)->ZZX_NSUTEF  		, NIL })
	aAdd(aTitInc,{ "E1_CARTAU"   , (_cAlias)->ZZX_CARTAU  		, NIL })

	BEGIN TRANSACTION
		MsExecAuto({|x,y| FINA040(x,y)}, aTitInc, nOpc)
		
		If lMsErroAuto
			cLogErro := ""
			aErroAuto := GetAutoGRLog()
			
			For nCount := 1 To Len(aErroAuto)
				cLogErro += StrTran(StrTran(aErroAuto[nCount], "<", ""), "-", "") + " "
			Next nCount
			
			DBSelectArea('ZZX')
			DBGoTo((_cAlias)->R_E_C_N_O_)
			RecLock("ZZX",.F.)
				ZZX_USRPRO := cUserName
				ZZX_DTPROS := dDataBase
				ZZX_ERRINC := cLogErro
				ZZX_STATUS := 'EI'
			ZZX->(MsUnlock())
		Else
			DBSelectArea('ZZX')
			DBGoTo((_cAlias)->R_E_C_N_O_)
			RecLock("ZZX",.F.)
				ZZX_USRPRO := cUserName
				ZZX_DTPROS := dDataBase
				ZZX_ERRINC := ""
				ZZX_STATUS := 'TI'
			ZZX->(MsUnlock())
			IF MV_PAR09 == 1
				fnBXTit() //Baixa o Titulo
			EndIF 
		EndIf
	END TRANSACTION

	cFilAnt := cFilAux

Return

/*--------------------------------------------------------------------------*
 | Func:  fnBXTit()                                                         |
 | Desc:  Faz a baixa do Titulo incluído anteriormente pelo ExecAuto FINA040|
 | Obs.:  /                                                                 |
 *-------------------------------------------------------------------------*/

Static Function fnBXTit()
	Local nOpc := 3
	Local nCount := 0
	Local aBXTit := {}
	Local aErroAuto := {}
	Local cLogErro := ""
	Local cHistBx := IIF(Empty((_cAlias)->ZZX_HIST), cHistBx, (_cAlias)->ZZX_HIST + " - " + FWTimeStamp(2) )
	Local cFilAux := cFilAnt
	Local dDataBkp := dDatabase
	Local dDtBaixa := SToD((_cAlias)->ZZX_BAIXA)
	Local cStatusBx := ""

	Private lMsErroAuto := .F.
	Private lAutoErrNoFile := .T.

	cFilAnt := (_cAlias)->ZZX_FILMOV
	dDatabase := dDtBaixa

	AAdd(aBXTit, {"E1_FILIAL" 	, Pad((_cAlias)->ZZX_FILMOV,FWTamSX3("E1_FILIAL")[01])	, Nil })
	AAdd(aBXTit, {"E1_PREFIXO"	, Pad((_cAlias)->ZZX_PREFIX,FWTamSX3("E1_PREFIXO")[01]) , Nil })
	AAdd(aBXTit, {"E1_NUM"    	, Pad((_cAlias)->ZZX_NUM   ,FWTamSX3("E1_NUM")[01])		, Nil })
	AAdd(aBXTit, {"E1_PARCELA"	, Pad((_cAlias)->ZZX_PARCEL,FWTamSX3("E1_PARCELA")[01])	, Nil })
	AAdd(aBXTit, {"E1_TIPO"   	, Pad((_cAlias)->ZZX_TIPO  ,FWTamSX3("E1_TIPO")[01])	, Nil })
	AAdd(aBXTit, {"E1_CLIENTE"	, Pad((_cAlias)->ZZX_CLIENT,FWTamSX3("E1_CLIENTE")[01])	, Nil })
	AAdd(aBXTit, {"E1_LOJA"   	, Pad((_cAlias)->ZZX_LOJA  ,FWTamSX3("E1_LOJA")[01])	, Nil })
	AAdd(aBXTit, {"AUTMOTBX"  	, "NOR"													, Nil })
	AAdd(aBXTit, {"AUTBANCO"  	, Pad((_cAlias)->ZZX_BANCO ,FWTamSX3("ZZX_BANCO")[01])	, Nil })
	AAdd(aBXTit, {"AUTAGENCIA"	, Pad((_cAlias)->ZZX_AGENCI,FWTamSX3("ZZX_AGENCI")[01])	, Nil })
	AAdd(aBXTit, {"AUTCONTA"  	, Pad((_cAlias)->ZZX_CONTA ,FWTamSX3("ZZX_CONTA")[01])	, Nil })
	AAdd(aBXTit, {"AUTDTBAIXA"	, dDtBaixa												, Nil })
	AAdd(aBXTit, {"AUTDTCREDITO", dDtBaixa												, Nil })
	AAdd(aBXTit, {"AUTHIST"     , cHistBx												, Nil })
	AAdd(aBXTit, {"AUTVLRPG"    , (_cAlias)->ZZX_VALOR									, Nil })

	AcessaPerg("FINA070", .F.)

	BEGIN TRANSACTION
		If nOpc == 5
			MsExecauto({|x,y,z,v| FINA070(x,y,z,v)}, aBXTit, nOpc, .F., nSeqBx)
		Else
			MsExecAuto({|x, y| FINA070(x, y)}, aBXTit, nOpc)
		EndIf
		
		If lMsErroAuto
			cLogErro := ""
			aErroAuto := GetAutoGRLog()
			
			For nCount := 1 To Len(aErroAuto)
				cLogErro += StrTran(StrTran(aErroAuto[nCount], "<", ""), "-", "") + " "
			Next nCount
			
			DBSelectArea('ZZX')
			DBGoTo((_cAlias)->R_E_C_N_O_)
			RecLock("ZZX",.F.)
				ZZX_USRPRO := cUserName
				ZZX_DTPROS := dDataBkp
				ZZX_ERRBX  := cLogErro
				ZZX_STATUS := 'EB'
			ZZX->(MsUnlock())
		Else
			IF Empty(SE1->E1_SALDO)
				cStatusBx := "OK"	
			ElseIF SE1->E1_SALDO < SE1->E1_VALOR
				cStatusBx := 'BP'
			EndIF
			DBSelectArea('ZZX')
			DBGoTo((_cAlias)->R_E_C_N_O_)
			RecLock("ZZX",.F.)
				ZZX_USRPRO := cUserName
				ZZX_DTPROS := dDataBkp
				ZZX_ERRBX  := ""
				ZZX_STATUS := cStatusBx
			ZZX->(MsUnlock())
		EndIf
	END TRANSACTION

	dDatabase := dDataBkp
	cFilAnt := cFilAux

Return
