//Bibliotecas
#Include "Protheus.ch"
 
/*------------------------------------------------------------------------------------------------------*
 | P.E.:  M460FIM                                                                                       |
 | Desc:  Gravação dos dados após gerar NF de Saída                                                     |
 | Links: http://tdn.totvs.com/pages/releaseview.action?pageId=6784180                                  |
 *------------------------------------------------------------------------------------------------------*/
 
User Function M460FIM()
    Local aAreaSF2 := SF2->(GetArea())
    Local aAreaSA1 := SA1->(GetArea())
    Local aAreaSE1 := SE1->(GetArea())
    Local _cAlias  := GetNextAlias()
    Local cQry     := ""
    
    Private cTipoCart := Space(FwTamSx3("E1_TIPO")[1])
    Private cAutoriz  := Space(FwTamSx3("E1_CARTAUT")[1])
    Private cNSU      := Space(FwTamSx3("E1_NSUTEF")[1])
    
    If FWAlertYesNo("Deseja preencher Autorização/NSU do cartão ?","Dados do Cartão")
        zPreCart()
        If Empty(cTipoCart) .Or. Empty(cAutoriz) .Or. Empty(cNSU)
            FWAlertWarning("Processo de atualização do título estornado devido a um ou mais campos não está preenchido","ATENÇÃO")
            RestArea(aAreaSF2)
            RestArea(aAreaSA1)
            RestArea(aAreaSE1)
            Return
        EndIF
    Else
        RestArea(aAreaSF2)
        RestArea(aAreaSA1)
        RestArea(aAreaSE1)
        Return
    EndIF

    cQry := " Select E1_CLIENTE, E1_LOJA, E1_PREFIXO, E1_NUM, E1_PARCELA, E1_TIPO, E1_VALOR from " + RetSqlName("SE1")
    cQry += "  where D_E_L_E_T_ <> '*'"
    cQry += "    and E1_FILIAL  = '" + xFilial("SE1") + "'"
    cQry += "    and E1_CLIENTE  = '" + SF2->F2_CLIENTE + "'"
    cQry += "    and E1_LOJA  = '" + SF2->F2_LOJA + "'"
    cQry += "    and E1_PREFIXO  = '" + SF2->F2_SERIE + "'"
    cQry += "    and E1_NUM  = '" + SF2->F2_DOC + "'"
    cQry := ChangeQuery(cQry)
    dbUseArea(.T.,"TOPCONN",TCGenQry(,,cQry),_cAlias,.F.,.T.)

    While !(_cAlias)->(EOF())
    
        DBSelectArea("SE1")
        SE1->(DBSetOrder(1)) //E1_FILIAL+E1_PREFIXO+E1_NUM+E1_PARCELA+E1_TIPO
        SE1->(DBGoTop())
        If SE1->(MsSeek(xFilial("SE1")+(_cAlias)->E1_PREFIXO+(_cAlias)->E1_NUM+(_cAlias)->E1_PARCELA+(_cAlias)->E1_TIPO))
            
                RecLock("SE1",.F.)
                    SE1->E1_TIPO    := cTipoCart
                    SE1->E1_CARTAUT := cAutoriz
                    SE1->E1_NSUTEF  := cNSU
                SE1->(MsUnlock())
        EndIf  

        (_cAlias)->(DbSkip())

    EndDo

    IF Select(_cAlias) <> 0
        (_cAlias)->(DBCloseArea())
    EndIf

    RestArea(aAreaSF2)
    RestArea(aAreaSA1)
    RestArea(aAreaSE1)
    
Return

/*/{Protheus.doc} zPreCart
    Tela para preenchimento dos dados do cartão
    @type  Static Function
    @author user
    @since 02/01/2024
    @version 1.0
    @return Nil
/*/
Static Function zPreCart()
    Local oDialog, oPanel, oTSay, oCombo, oDlg
    Local aTipo := {"CC - Cartão de Crédito","CD - Cartão de Débito"}

    oDialog := FWDialogModal():New()
    oDialog:SetBackground( .T. ) 
    oDialog:SetTitle( 'Preencha os dados do Cartão abaixo:' )
    oDialog:SetSize( 120, 250 )
    oDialog:EnableFormBar( .T. )
    oDialog:SetCloseButton( .F. )
    oDialog:SetEscClose( .F. )
    oDialog:CreateDialog()
    oDialog:CreateFormBar()
    oDialog:AddCloseButton(Nil, "Confirmar")

    oPanel := oDialog:GetPanelMain()

        oTSay  := TSay():New(10,5,{|| "Tipo do cartão ? "},oPanel,,,,,,.T.,,,50,70,,,,,,.T.)
        oCombo := TComboBox():New(30,53,{|u|iif(PCount()>0,cTipoCart:=u,cTipoCart)},aTipo,80,20,oDlg,,{||},,,,.T.,,,,,,,,,'cTipoCart')

        oTSay  := TSay():New(35,5,{|| "Autorização ? "},oPanel,,,,,,.T.,,,80,70,,,,,,.T.)
        @ 33,50 MSGET cAutoriz SIZE 060,009 OF oPanel PIXEL
        
        oTSay  := TSay():New(50,5,{|| "NSU Tef ?"},oPanel,,,,,,.T.,,,80,70,,,,,,.T.)
        @ 48,50 MSGET cNSU SIZE 060,009 OF oPanel PIXEL
        
    oDialog:Activate()

    cTipoCart := Alltrim(SubStr( cTipoCart, 1, At('-', cTipoCart) - 2))

Return
