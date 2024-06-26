#INCLUDE "TOTVS.CH"

/*/
calcular o c�digo sequencial do cadastro de produtos
N�o utilizamos getsxenum porque o cliente quer um c�digo de 6 char
e, como o campo possui 15 char, ele gera o c�digo com 15
/*/
User Function LASTSB1()

	LOCAL cRet
	LOCAL aArea := GetArea()

	BeginSQL Alias "TSB1"
		SELECT
		Max(B1_COD) PROXCOD
		FROM
		%table:SB1% SB1
		WHERE
		SB1.B1_FILIAL = %xFilial:SB1%
		AND SB1.B1_TIPO != "MO"
		AND SB1.%notdel%
	EndSQL

	cRet := Soma1(Alltrim(TSB1->PROXCOD))
	
	TSB1->(dbCloseArea())
	
	RestArea(aArea)

Return cRet

