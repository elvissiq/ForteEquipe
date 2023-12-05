#INCLUDE "Protheus.ch"
#INCLUDE "FWMVCDEF.ch"
#INCLUDE "TOTVS.ch"

/*/{Protheus.doc} FINA920A
Ponto de Entrada em MVC na rotina de Consciliação de Recebimentos
@type User Function
@author Elvis Siqueira
@since 04/12/2023
/*/

User Function FINA920A()
    Local aParam := PARAMIXB
    Local xRet := .T.
    Local oObj := ''
    Local cIdPonto := ''
    Local cIdModel := ''
    
    If aParam <> NIL

        oObj := aParam[1]
        cIdPonto := aParam[2]
        cIdModel := aParam[3]

        If cIdPonto == 'FT300ORC'

        EndIf
    
    EndIf 

Return xRet
