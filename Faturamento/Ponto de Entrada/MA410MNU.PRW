#INCLUDE 'TOTVS.CH'

User Function MA410MNU
    Local aArea := GetArea()
     
    //Adicionando função de vincular
    aadd(aRotina,{"Imprimir Pedido","U_ZRPEDVEN", 0 , 4, 0 , Nil})
    aadd(aRotina,{"Liberar PV","MATA440", 0 , 4, 0 , Nil})
    aadd(aRotina,{"Liberar Credito/Estoque","MATA456", 0 , 4, 0 , Nil})
    aadd(aRotina,{"Nfe Sefaz","SPEDNFE", 0 , 4, 0 , Nil})
    aadd(aRotina,{"Boleto","U_BOLETO", 0 , 4, 0 , Nil})
     
    RestArea(aArea)                   
Return
