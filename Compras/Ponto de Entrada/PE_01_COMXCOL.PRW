#INCLUDE 'TOTVS.CH'

User Function COMCOLRT
    Local aArea := GetArea()

	//Adicionando fun�ao de vincular
	aadd(aRotina,{"Ler XML de notas","U_IMPPRXML()", 0 , 3, 0 , Nil})

	RestArea(aArea)
Return
