#include "Protheus.ch"
 
//--------------------------------------------------------------------------------------------------------------------------------------------
/*/{Protheus.doc} F150QTDT
    Ponto de Entrada que permite alterar o nome e extens�o do arquivo de sa�da, gerado a partir da rotina Arquivos de Cobran�as (FINA150). 
/*/
//--------------------------------------------------------------------------------------------------------------------------------------------
 
User Function F150QTDT()
    Local cNomArq := Alltrim(SubSTR(MV_PAR04,Rat("\",MV_PAR04)+1))

    U_FFIN002(cNomArq) //Tela de Log do arquivo de Retorno CNAB
    
Return
