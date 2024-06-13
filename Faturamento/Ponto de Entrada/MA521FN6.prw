#Include "Protheus.ch"
#include "totvs.ch"
#include "topconn.ch"
 
/*----------------------------------------------------------------------------------------------------------*
 | P.E.:  MA521FN6                                                                                          |
 | Desc:  Este PE tem como finalidade realizar o desvio da valida��o de relacionamento entre a nota e baixa |
 |        de ativos, assim � poss�vel desabilitar a valida��o deste vinculo na exclus�o de documento        |
 |        de sa�da, com isso mesmo que existe o v�nculo gravado na tabela FN6 a nota ser� exclu�da.         |
 |                                                                                                          |    
 | O retorno da fun��o deve ser l�gico, sendo:                                                              |    
 |  .T. - Desabilita valida��o de vinculo FN6.                                                              |
 |  .F. - N�o desabilita valida��o de vinculo FN6.                                                          |
 |                                                                                                          |
 | Obs.: Este PE n�o manipula informa��es de qualquer tabela, sendo sua fun��o �nica e exclusivamente para  |
 |         desabilitar a valida��o conforme descrito.                                                       |
 *---------------------------------------------------------------------------------------------------------*/
 
User Function MA521FN6()
    Local aAreaSE1 := SE1->(FWGetArea())
    Local cQry     := ""
    Local _cAlias  := GetNextAlias()
    
    cQry := " SELECT E1_FILIAL, E1_PREFIXO, E1_NUM, E1_PARCELA, E1_TIPO, E1_CLIENTE, E1_LOJA FROM " + RetSqlName("SE1")
    cQry += " WHERE D_E_L_E_T_ <> '*'"
    cQry += "       AND E1_FILIAL   = '" + xFilial("SE1") + "'"
    cQry += "       AND E1_CLIENTE  = '" + SF2->F2_CLIENTE + "'"
    cQry += "       AND E1_LOJA     = '" + SF2->F2_LOJA + "'"
    cQry += "       AND E1_PREFIXO  = '" + SF2->F2_SERIE + "'"
    cQry += "       AND E1_NUM      = '" + SF2->F2_DOC + "'"
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
                SE1->E1_TIPO := MVNOTAFIS
            SE1->(MsUnlock())
        EndIf
        If FK7->(MsSeek(xFilial("FK7")+"SE1"+(_cAlias)->E1_FILIAL+(_cAlias)->E1_PREFIXO+(_cAlias)->E1_NUM+(_cAlias)->E1_PARCELA+;
                                             (_cAlias)->E1_TIPO+(_cAlias)->E1_CLIENTE+(_cAlias)->E1_LOJA))

                RecLock("FK7",.F.)
                    FK7->FK7_CHAVE := (_cAlias)->E1_FILIAL + "|" + (_cAlias)->E1_PREFIXO + "|" + (_cAlias)->E1_NUM +;
                                      MVNOTAFIS + "|" + (_cAlias)->E1_CLIENTE + "|" + (_cAlias)->E1_LOJA
                    FK7->FK7_TIPO  := MVNOTAFIS
                FK7->(MsUnlock())
        EndIf
        (_cAlias)->(DbSkip())
    EndDo

    IF Select(_cAlias) <> 0
        (_cAlias)->(DBCloseArea())
    EndIf
    FWRestArea(aAreaSE1)
Return .F.
