#include "Protheus.ch"
#include "Totvs.ch"
#include "TbiConn.ch"
#include "TopConn.ch"
 
//-------------------------------------------------------------------------------
/*/{Protheus.doc} F200VAR
Manipular as informações (variáveis) no retorno do Cnab a Receber (FINA200). 
 
@PARAMIXB aDados[1] = Número do Título    | Variavel de origem: cNumTit
          aDados[2] = Data da Baixa       | Variavel de origem: dBaixa
          aDados[3] = Tipo do Título      | Variavel de origem: cTipo
          aDados[4] = Nosso Número        | Variavel de origem: cNsNum
          aDados[5] = Valor da Despesa    | Variavel de origem: nDespes
          aDados[6] = Valor do Desconto   | Variavel de origem: nDescont
          aDados[7] = Valor do Abatimento | Variavel de origem: nAbatim
          aDados[8] = Valor Recebido      | Variavel de origem: nValRec
          aDados[9] = Juros               | Variavel de origem: nJuros
          aDados[10] = Multa              | Variavel de origem: nMulta
          aDados[11] = Outras Despesas    | Variavel de origem: nOutrDesp
          aDados[12] = Valor do Credito   | Variavel de origem: nValCc
          aDados[13] = Data do Credito    | Variavel de origem: dDataCred
          aDados[14] = Ocorrência         | Variavel de origem: cOcorr
          aDados[15] = Motivo do banco    | Variavel de origem: cMotBan
          aDados[16] = Linha Inteira      | Variavel de origem: xBuffer
          aDados[17] = Data de Vencimento | Variavel de origem: dDtVc
 
/*/
//-------------------------------------------------------------------------------
 
User Function F200VAR()
 
    Local aDados   := PARAMIXB
    Local aAreaSE1 := SE1->(FWGetArea())
    //Local cNTitulo := Alltrim(Paramixb[1][1])
    Local cIdCNAB  := Alltrim(Paramixb[1][4])
    //Local nVlJuros := Paramixb[1][9]
    Local cOrgTit  := SuperGetMV("MV_XORGTIT",.F.,"IMPORT")
    Local cQry     := ""
    Local _cAlias  := GetNextAlias()

    cQry := " SELECT * "
	cQry += " FROM " + RetSqlName("SE1")
	cQry += " WHERE D_E_L_E_T_ = '' "
    cQry += "   AND E1_FILIAL = '" + xFilial("SE1") + "' "
    cQry += "   AND E1_SALDO  > 0 "
    cQry += "   AND E1_ORIGEM = '" + cOrgTit + "' "
    cQry += "   AND E1_IDCNAB LIKE ('" + cIdCNAB + "%') "
    IF Select(_cAlias) <> 0
        (_cAlias)->(DbCloseArea())
    EndIf
    TCQuery cQry New Alias &_cAlias

    DBselectArea("SE1")

    IF !(_cAlias)->(Eof())

        SE1->(DbSetOrder(2))
        SE1->(DBGoTop())
        IF SE1->(MsSeek(xFilial("SE1")+(_cAlias)->E1_CLIENTE+(_cAlias)->E1_LOJA+(_cAlias)->E1_PREFIXO+(_cAlias)->E1_NUM+(_cAlias)->E1_PARCELA+(_cAlias)->E1_TIPO))
            
            cNumTit := SE1->E1_IDCNAB
            aDados[1][1] := SE1->E1_IDCNAB
            /*
            RecLock("SE1",.F.)
                SE1->E1_IDCNAB := cNTitulo
            SE1->(MSUnLock()) 
            */
        EndIF 

    EndIF 

    (_cAlias)->(DbCloseArea()) 

    /*
    SE1->(DbSetOrder(16))
    SE1->(DBGoTop())
    IF SE1->(MsSeek(xFilial("SE1") + cNTitulo ))
        If SE1->E1_SALDO < nValRec
            nJuros += ( (nValRec - nVlJuros ) - SE1->E1_SALDO ) + nDespes
            Paramixb[1][9] += ( (nValRec - nVlJuros ) - SE1->E1_SALDO ) + nDespes
        EndIF
    EnDIF 
    */

    FWRestArea(aAreaSE1)
 
Return(aDados)
