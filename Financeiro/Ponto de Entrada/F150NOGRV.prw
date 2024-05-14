#include "Protheus.ch"
#include "Totvs.ch"
#include "TbiConn.ch"
#include "TopConn.ch"
 
//--------------------------------------------------------------------------------------------------------------------
/*/{Protheus.doc} F150NOGRV
    Ponto de Entrada que permite desprezar a gravação, de título com impedimento, no arquivo de remessa de cobrança.
/*/
//--------------------------------------------------------------------------------------------------------------------
 
User Function F150NOGRV()
    Local lRet    := .T.
    Local lInclui := .T.
    Local cNomArq := Alltrim(SubSTR(MV_PAR04,Rat("\",MV_PAR04)+1))

    //Grava o LOG do arquivo na tabela Customizada para abrir em tela posteriormente
    IF ExisteSX2("ZZ1")
        
        DBselectArea("ZZ1")
        If ZZ1->(MSSeek(xFilial("ZZ1")+Pad(cNomArq,FWTamSX3("ZZ1_ARQUIV")[1])+;
                                       Pad(SE1->E1_NUM,FWTamSX3("ZZ1_NUMTIT")[1])+;
                                       Pad(SE1->E1_IDCNAB,FWTamSX3("ZZ1_NSNUM")[1])))
            lInclui := .F. //Alteração
        EndIf 

        RecLock("ZZ1",lInclui)
            ZZ1->ZZ1_FILIAL := xFilial("ZZ1")
            ZZ1->ZZ1_NUMTIT := SE1->E1_NUM
            ZZ1->ZZ1_DBAIXA := SE1->E1_BAIXA
            ZZ1->ZZ1_TIPO   := SE1->E1_TIPO
            ZZ1->ZZ1_NSNUM  := SE1->E1_IDCNAB
            ZZ1->ZZ1_VLDESP := 0
            ZZ1->ZZ1_VLDESC := SE1->E1_DESCFIN
            ZZ1->ZZ1_VLABAT := ( SE1->E1_CSLL + SE1->E1_COFINS + SE1->E1_PIS + SE1->E1_IRRF )
            ZZ1->ZZ1_VLREC  := SE1->E1_VALOR
            ZZ1->ZZ1_JUROS  := SE1->E1_JUROS
            ZZ1->ZZ1_MULTA  := SE1->E1_MULTA
            ZZ1->ZZ1_OUTDES := 0
            ZZ1->ZZ1_VLCRED := SE1->E1_VALOR
            ZZ1->ZZ1_DCRED  := SE1->E1_VENCREA
            ZZ1->ZZ1_OCORR  := ""
            ZZ1->ZZ1_MOTBAN := ""
            ZZ1->ZZ1_LINHA  := Alltrim(AllToChar(nSeq+2, "", .F.))
            ZZ1->ZZ1_LINARQ := xBuffer
            ZZ1->ZZ1_DVENC  := SE1->E1_VENCREA
            ZZ1->ZZ1_BANCO  := MV_PAR05
            ZZ1->ZZ1_AGENCI := MV_PAR06
            ZZ1->ZZ1_NUMCOM := MV_PAR07
            ZZ1->ZZ1_ARQUIV := cNomArq
        ZZ1->(MSUnLock())

    EndIF

Return lRet 
