#INCLUDE "Protheus.CH"
#INCLUDE "Totvs.ch"
#INCLUDE "TOPCONN.ch"
#INCLUDE "TBICONN.ch"
#INCLUDE "FWMVCDef.CH"

//--------------------------------------------------------------------------------------------------------------------
/*/{PROTHEUS.DOC} FFIN003
FUNÇÃO FFIN003- Titulos Liquidados (Legado)
@VERSION PROTHEUS 12
@SINCE 14/10/2024
/*/
User Function FFIN003()
	Local oBrowse

	oBrowse := FWMBrowse():New()
	oBrowse:SetDescription("Titulos Liquidados (Legado)")
	oBrowse:SetAmbiente(.F.)
	oBrowse:SetWalkThru(.F.)
	oBrowse:SetMenuDef('FFIN003')
	oBrowse:SetAlias('ZZX')
	oBrowse:DisableDetails()
	oBrowse:SetFixedBrowse(.T.)

	oBrowse:AddLegend("Empty(ZZX_STATUS)", 'BR_VERDE'    , 'Apto a processar')
	oBrowse:AddLegend("ZZX_STATUS == 'E'", 'BR_VERMELHO' , 'Erro na Importacao')
	oBrowse:AddLegend("ZZX_STATUS == 'I'", 'BR_AMARELO'  , 'Erro na Inclusao do Titulo')
	oBrowse:AddLegend("ZZX_STATUS == 'B'", 'BR_LARANJA'  , 'Erro na Baixa do Titulo')

	oBrowse:Activate()

Return oBrowse

/*---------------------------------------------------------------------*
 | Func:  MenuDef                                                      |
 | Desc:  Criação do Menu da rotina                                    |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function MenuDef()
	Local aRot := {}
	
	ADD OPTION aRot TITLE 'Importar'   ACTION 'U_IMPFIN003' 	OPERATION 3 ACCESS 0
	ADD OPTION aRot TITLE 'Processar'  ACTION 'U_PROCES003' 	OPERATION 4 ACCESS 0
	ADD OPTION aRot TITLE 'Visualizar' ACTION 'VIEWDEF.FFIN003' OPERATION 2 ACCESS 0
	ADD OPTION aRot TITLE 'Excluir'    ACTION 'VIEWDEF.FFIN003' OPERATION 5 ACCESS 0

Return aRot

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ModelDef()
	Local oModel    := Nil
	Local oStPai   	:= FWFormStruct( 1, 'ZZX')

	oModel := MPFormModel():New("FFIN003MV" , , , )
	oModel:SetDescription(OemtoAnsi("Titulos Liquidados (Legado)") )
	oModel:AddFields('ZZXMASTER',/*cOwner*/,oStPai)
	oModel:SetPrimaryKey( {"ZZX_FILIAL", "ZZX_FILMOV",  "ZZX_PREFIX", "ZZX_NUM", "ZZX_PARCEL", "ZZX_TIPO"})

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
	oView:AddField('VIEW_CAB',oStPai,'SZZMASTER')
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

	Local cArq 		:= tFileDialog( "Arquivo de planilha Excel (*.csv)",,,, .F., /*GETF_MULTISELECT*/)
	Local nAtu 		:= 0
    Local nFim  	:= 0
	Local cLinha 	:= ""
	Local aRegistro := {}
	Local cNumImp   := fPrxNumZZX()

	If !File(cArq)
		FwAlertError("Não foi possível importar o arquivo, por favor tente novamente.","Importar Arquivo")
        Return
    EndIf

	FT_FUSE(cArq)
    nFim := FT_FLASTREC()
    ProcRegua(nFim)
    FT_FGOTOP()

	DBSelectArea("ZZX")

	While !FT_FEOF()
            
        nAtu++
        IncProc("Lendo a linha " + cValToChar(nAtu) + " de " + cValToChar(nFim) + "...")

		cLinha := FT_FREADLN()
                
        If !Empty(cLinha)
            aRegistro := {}
            aRegistro := Separa(cLinha,";",.T.)
            
			If Len(aRegistro) == 24

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
						ZZX->ZZX_EMISSA := CToD(aRegistro[10])
						ZZX->ZZX_VENCTO := CToD(aRegistro[11])
						ZZX->ZZX_VENCRE := CToD(aRegistro[12])
						ZZX->ZZX_VALOR  := Val(aRegistro[13])
						ZZX->ZZX_VEND1  := Pad(aRegistro[14],FWTamSX3("ZZX_VEND1")[1] )
						ZZX->ZZX_BAIXA  := CToD(aRegistro[15])
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

				EndIF 
				
			Else

			EndIF 

        EndIF 
	FT_FSKIP()
    EndDo

    FT_FUSE()
Return

//====================================================================================================================\
/*/{Protheus.doc}fPrxNumZZX
  ====================================================================================================================
	@description
	Retorna o próximo número para a tabela ZZX
/*/
//===================================================================================================================\
Static Function fPrxNumZZX()

	Local cRet := StrZero(0,10)
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
