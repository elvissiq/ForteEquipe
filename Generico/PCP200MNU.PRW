#INCLUDE 'TOTVS.CH'
#Include 'FWMVCDEF.ch'

User Function PCP200MNU
	Local aArea := GetArea()

	//Adicionando funçao de vincular
	aadd(aRotina,{"Atualiza via Planilha","U_INCSG1()", 0 , 4, 0 , Nil})

	RestArea(aArea)
Return
