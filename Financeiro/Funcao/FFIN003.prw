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
	oBrowse:AddLegend("ZZX_STATUS == 'BX'", 'BR_CINZA'   , 'Titulo Baixado')
	oBrowse:AddLegend("ZZX_STATUS == 'OK'", 'BR_VERDE'   , 'Titulo Incluido/Baixado')

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
	Processa({|| xProcessa()}, "Importando arquivo...")
Return 

Static Function xProcessa()

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
