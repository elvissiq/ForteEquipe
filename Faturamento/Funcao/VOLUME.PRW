#Include "RwMake.ch"

/*/{Protheus.doc} VOLUME
Fun��o para calcular o Peso e volume no Pedido de Venda
@author TOTVS NORDESTE
@since 07/07/2023
@version 1.0
    @return Nil, Fun��o n�o tem retorno
    @example
    u_VOLUME()
    @obs N�o se pode executar fun��o MVC dentro do f�rmulas
/*/

User Function VOLUME()
	Local cTipo    := ""
	Local nVolume  := 0
	Local nVol	   := 0
	Local nConv    := 0
	Local nPosProd := Ascan(aHeader, {|x| AllTrim(x[2]) == "C6_PRODUTO"})
	Local nPosQTD  := Ascan(aHeader , { |x| Alltrim(x[2]) = "C6_QTDVEN"})
	Local lUsaConv := SuperGetMV("BO_USACONV",.F.,.F.)
	Local i, x

	If lUsaConv
		For i:=1 to len(aCols)
			cTipo := POSICIONE("SB1",1,FWxFilial("SB1")+aCols[i][nPosProd],"B1_TIPCONV")
			nConv := POSICIONE("SB1",1,FWxFilial("SB1")+aCols[i][nPosProd],"B1_CONV") 

				IF cTipo == 'M' .And. !Empty(nConv)
					nVolume	:= int(aCols[i,nPosQTD] *nConv)
					
						nVol += iif(nVolume <> aCols[i,nPosQTD] *nConv,nVolume+1,nVolume)

				ELSEIF cTipo == 'D' .And. !Empty(nConv) .And. lUsaConv
					nVolume	:= int(aCols[i,nPosQTD] /nConv)
						
						nVol += iif(nVolume <> aCols[i,nPosQTD] /nConv,nVolume+1,nVolume)

				ELSEIF Empty(nVol)

					For x:=1 to len(aCols)
						nVol += aCols[x,nPosQTD]
					Next x

				ENDIF
		Next i
	Else 
	
		For i:=1 to len(aCols)
			nVol += aCols[i,nPosQTD]
		Next i			

	EndIf 

	nVol := Round(nVol,TAMSX3("C5_VOLUME1")[2])

Return nVol