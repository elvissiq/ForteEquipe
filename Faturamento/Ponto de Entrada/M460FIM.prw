//Bibliotecas
#Include "Protheus.ch"
 
/*------------------------------------------------------------------------------------------------------*
 | P.E.:  M460FIM                                                                                       |
 | Desc:  Gravação dos dados após gerar NF de Saída                                                     |
 | Links: http://tdn.totvs.com/pages/releaseview.action?pageId=6784180                                  |
 *------------------------------------------------------------------------------------------------------*/
 
User Function M460FIM()
    Local aAreaSF2  := SF2->(FWGetArea())
    Local aAreaSA1  := SA1->(FWGetArea())
    Local aAreaSE1  := SE1->(FWGetArea())
    Local _cAlias   := GetNextAlias()
    Local cQry      := ""
    
    Private cTipoCart := Space(FwTamSx3("E1_TIPO")[1])
    Private cAutoriz  := Space(FwTamSx3("E1_CARTAUT")[1])
    Private cNSU      := Space(FwTamSx3("E1_NSUTEF")[1])
    
    If FWAlertYesNo("Deseja preencher Autorização/NSU do cartão ?","Dados do Cartão")
        zPreCart()
        If Empty(cTipoCart) .Or. Empty(cAutoriz) .Or. Empty(cNSU)
            FWAlertWarning("Processo de atualização do título estornado devido a um ou mais campos não está preenchido","ATENÇÃO")
            FWRestArea(aAreaSF2)
            FWRestArea(aAreaSA1)
            FWRestArea(aAreaSE1)
            Return
        EndIF
    Else
        FWRestArea(aAreaSF2)
        FWRestArea(aAreaSA1)
        FWRestArea(aAreaSE1)
        Return
    EndIF

    cQry := " SELECT E1_FILIAL, E1_PREFIXO, E1_NUM, E1_PARCELA, E1_TIPO, E1_CLIENTE, E1_LOJA FROM " + RetSqlName("SE1")
    cQry += " WHERE D_E_L_E_T_ <> '*'"
    cQry += "       AND E1_FILIAL  = '" + xFilial("SE1") + "'"
    cQry += "       AND E1_CLIENTE = '" + SF2->F2_CLIENTE + "'"
    cQry += "       AND E1_LOJA    = '" + SF2->F2_LOJA + "'"
    cQry += "       AND E1_PREFIXO = '" + SF2->F2_SERIE + "'"
    cQry += "       AND E1_NUM     = '" + SF2->F2_DOC + "'"
    cQry := ChangeQuery(cQry)
    IF Select(_cAlias) <> 0
        (_cAlias)->(DBCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TCGenQry(,,cQry),_cAlias,.F.,.T.)
    
    DBSelectArea("SE1")
    DBSelectArea("FK7")
    SE1->(DbSetOrder(1)) //E1_FILIAL+E1_PREFIXO+E1_NUM+E1_PARCELA+E1_TIPO
    FK7->(DbSetOrder(3)) //FK7_FILIAL+FK7_ALIAS+FK7_FILTIT+FK7_PREFIX+FK7_NUM+FK7_PARCEL+FK7_TIPO+FK7_CLIFOR+FK7_LOJA
    SE1->(DBGoTop())
    FK7->(DBGoTop())

    While !(_cAlias)->(EOF())
        If SE1->(MsSeek(xFilial("SE1")+(_cAlias)->E1_PREFIXO+(_cAlias)->E1_NUM+(_cAlias)->E1_PARCELA+(_cAlias)->E1_TIPO))
                RecLock("SE1",.F.)
                    SE1->E1_TIPO    := cTipoCart
                    SE1->E1_CARTAUT := cAutoriz
                    SE1->E1_NSUTEF  := cNSU
                SE1->(MsUnlock())
        EndIf
        If FK7->(MsSeek(xFilial("FK7")+"SE1"+(_cAlias)->E1_FILIAL+(_cAlias)->E1_PREFIXO+(_cAlias)->E1_NUM+(_cAlias)->E1_PARCELA+;
                                             (_cAlias)->E1_TIPO+(_cAlias)->E1_CLIENTE+(_cAlias)->E1_LOJA))

                RecLock("FK7",.F.)
                    FK7->FK7_CHAVE := (_cAlias)->E1_FILIAL + "|" + (_cAlias)->E1_PREFIXO + "|" + (_cAlias)->E1_NUM + "|" + ;
                                      Pad(cTipoCart,FwTamSx3("E1_TIPO")[1]) + "|" + (_cAlias)->E1_CLIENTE + "|" + (_cAlias)->E1_LOJA
                    FK7->FK7_TIPO  := cTipoCart
                FK7->(MsUnlock())
        EndIf
        (_cAlias)->(DbSkip())
    EndDo

    IF Select(_cAlias) <> 0
        (_cAlias)->(DBCloseArea())
    EndIf

    FWRestArea(aAreaSF2)
    FWRestArea(aAreaSA1)
    FWRestArea(aAreaSE1)
    
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
