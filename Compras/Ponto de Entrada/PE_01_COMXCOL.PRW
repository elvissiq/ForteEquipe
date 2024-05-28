#INCLUDE 'TOTVS.CH'

User Function COMCOLRT
    Local aArea := GetArea()

	//Adicionando funçao de vincular
	aadd(aRotina,{"Ler XML de notas","U_IMPPRXML()", 0 , 3, 0 , Nil})

	RestArea(aArea)
Return
