#include "Protheus.ch"
 
//--------------------------------------------------------------------------------------------------------------------------------------------
/*/{Protheus.doc} F150QTDT
    Ponto de Entrada que permite alterar o nome e extensão do arquivo de saída, gerado a partir da rotina Arquivos de Cobranças (FINA150). 
/*/
//--------------------------------------------------------------------------------------------------------------------------------------------
 
User Function F150QTDT()
    Local cNomArq := Alltrim(SubSTR(MV_PAR04,Rat("\",MV_PAR04)+1))

    U_FFIN002(cNomArq) //Tela de Log do arquivo de Retorno CNAB
    
Return
