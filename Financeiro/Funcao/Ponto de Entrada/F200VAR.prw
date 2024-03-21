#include "Protheus.ch"
#include "Totvs.ch"
#include "TbiConn.ch"
#include "TopConn.ch"
 
//-------------------------------------------------------------------------------
/*/{Protheus.doc} F200VAR
Manipular as informa��es (vari�veis) no retorno do Cnab a Receber (FINA200). 
 
@PARAMIXB aDados[1] = N�mero do T�tulo    | Variavel de origem: cNumTit
          aDados[2] = Data da Baixa       | Variavel de origem: dBaixa
          aDados[3] = Tipo do T�tulo      | Variavel de origem: cTipo
          aDados[4] = Nosso N�mero        | Variavel de origem: cNsNum
          aDados[5] = Valor da Despesa    | Variavel de origem: nDespes
          aDados[6] = Valor do Desconto   | Variavel de origem: nDescont
          aDados[7] = Valor do Abatimento | Variavel de origem: nAbatim
          aDados[8] = Valor Recebido      | Variavel de origem: nValRec
          aDados[9] = Juros               | Variavel de origem: nJuros
          aDados[10] = Multa              | Variavel de origem: nMulta
          aDados[11] = Outras Despesas    | Variavel de origem: nOutrDesp
          aDados[12] = Valor do Credito   | Variavel de origem: nValCc
          aDados[13] = Data do Credito    | Variavel de origem: dDataCred
          aDados[14] = Ocorr�ncia         | Variavel de origem: cOcorr
          aDados[15] = Motivo do banco    | Variavel de origem: cMotBan
          aDados[16] = Linha Inteira      | Variavel de origem: xBuffer
          aDados[17] = Data de Vencimento | Variavel de origem: dDtVc
 
/*/
//-------------------------------------------------------------------------------
 
User Function F200VAR()
 
    Local aDados   := PARAMIXB
    Local aAreaSE1 := SE1->(FWGetArea())
    Local cNumTit  := Alltrim(Paramixb[1][1])
    Local cIDCnab  := Alltrim(Paramixb[1][4])
    Local cOrgTit  := SuperGetMV("MV_XORGTIT",.F.,"IMPORT")
    Local cQry     := ""
    Local _cAlias  := GetNextAlias()
    
    cQry := " SELECT * "
	cQry += " FROM " + RetSqlName("SE1")
	cQry += " WHERE D_E_L_E_T_ = '' "
    cQry += "   AND E1_FILIAL = '" + xFilial("SE1") + "' "
    cQry += "   AND E1_SALDO  > 0 "
    cQry += "   AND E1_ORIGEM = '" + cOrgTit + "' "
    cQry += "   AND E1_IDCNAB LIKE ('" + cIDCnab + "%') "
    IF Select(_cAlias) <> 0
        (_cAlias)->(DbCloseArea())
    EndIf
    TCQuery cQry New Alias &_cAlias

    IF !(_cAlias)->(Eof())

        DBselectArea("SE1")
        SE1->(DbSetOrder(2))
        SE1->(DBGoTop())
        IF SE1->(MsSeek(xFilial("SE1")+(_cAlias)->E1_CLIENTE+(_cAlias)->E1_LOJA+(_cAlias)->E1_PREFIXO+(_cAlias)->E1_NUM+(_cAlias)->E1_PARCELA+(_cAlias)->E1_TIPO))
            
            RecLock("SE1",.F.)
                SE1->E1_IDCNAB := cNumTit
            SE1->(MSUnLock())

        EndIF 

    EndIF 

    (_cAlias)->(DbCloseArea()) 
    FWRestArea(aAreaSE1)
 
Return(aDados)
