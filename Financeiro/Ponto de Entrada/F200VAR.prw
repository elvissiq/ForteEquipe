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
    Local aAreaZZ1 := ZZ1->(FWGetArea())
    Local cNumTit  := Alltrim(Paramixb[1][1])
    Local cIDCnab  := Alltrim(Paramixb[1][4])
    Local cOrgTit  := SuperGetMV("MV_XORGTIT",.F.,"IMPORT")
    Local lInclui  := .T.
    Local cQry     := ""
    Local _cAlias  := GetNextAlias()
    Local cNomArq  := MV_PAR04

    If IsSrvUnix()
        cNomArq := Alltrim(SubSTR(MV_PAR04,Rat("/",MV_PAR04)+1))
    Else
        cNomArq := Alltrim(SubSTR(MV_PAR04,Rat("\",MV_PAR04)+1))
    EndIF

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

    //Grava o LOG do arquivo na tabela Customizada para abrir em tela posteriormente
    IF ExisteSX2("ZZ1")
        
        DBselectArea("ZZ1")
        If ZZ1->(MSSeek(xFilial("ZZ1")+Pad(Alltrim(SubSTR(MV_PAR04,Rat("\",MV_PAR04)+1)),FWTamSX3("ZZ1_ARQUIV")[1])+;
                                        Pad(aDados[1][1],FWTamSX3("ZZ1_NUMTIT")[1])+;
                                        Pad(aDados[1][4],FWTamSX3("ZZ1_NSNUM")[1])))
            lInclui := .F. //Alteração
        EndIf 

        RecLock("ZZ1",lInclui)
            ZZ1->ZZ1_FILIAL := xFilial("ZZ1")
            ZZ1->ZZ1_NUMTIT := aDados[1][1]
            ZZ1->ZZ1_DBAIXA := aDados[1][2]
            ZZ1->ZZ1_TIPO   := aDados[1][3]
            ZZ1->ZZ1_NSNUM  := aDados[1][4]
            ZZ1->ZZ1_VLDESP := aDados[1][5]
            ZZ1->ZZ1_VLDESC := aDados[1][6]
            ZZ1->ZZ1_VLABAT := aDados[1][7]
            ZZ1->ZZ1_VLREC  := aDados[1][8]
            ZZ1->ZZ1_JUROS  := aDados[1][9]
            ZZ1->ZZ1_MULTA  := aDados[1][10]
            ZZ1->ZZ1_OUTDES := aDados[1][11]
            ZZ1->ZZ1_VLCRED := aDados[1][12]
            ZZ1->ZZ1_DCRED  := aDados[1][13]
            ZZ1->ZZ1_OCORR  := aDados[1][14]
            ZZ1->ZZ1_MOTBAN := aDados[1][15]
            ZZ1->ZZ1_LINHA  := ""
            ZZ1->ZZ1_LINARQ := aDados[1][16]
            ZZ1->ZZ1_DVENC  := aDados[1][17]
            ZZ1->ZZ1_BANCO  := cBanco
            ZZ1->ZZ1_AGENCI := cAgencia
            ZZ1->ZZ1_NUMCOM := cConta
            ZZ1->ZZ1_ARQUIV := cNomArq
        ZZ1->(MSUnLock())

    EndIF 

    FWRestArea(aAreaZZ1)
    FWRestArea(aAreaSE1)
 
Return(aDados)
