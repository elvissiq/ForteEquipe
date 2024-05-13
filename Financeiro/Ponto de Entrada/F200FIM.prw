#include "Protheus.ch"
 
//-------------------------------------------------------------------------------
/*/{Protheus.doc} F200FIM
    O ponto de entrada F200FIM do CNAB a receber sera executado após gravar a 
    linha de lançamento contábil no arquivo de contra-prova.  
/*/
//-------------------------------------------------------------------------------
 
User Function F200FIM()
    Local cNomArq := MV_PAR04

    If IsSrvUnix()
        cNomArq := Alltrim(SubSTR(MV_PAR04,Rat("/",MV_PAR04)+1))
    Else
        cNomArq := Alltrim(SubSTR(MV_PAR04,Rat("\",MV_PAR04)+1))
    EndIF

    U_FFIN002(cNomArq) //Tela de Log do arquivo de Retorno CNAB
    
Return
