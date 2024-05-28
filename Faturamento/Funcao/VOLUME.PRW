#Include "RwMake.ch"

/*/{Protheus.doc} VOLUME
Função para calcular o Peso e volume no Pedido de Venda
@author TOTVS NORDESTE
@since 15/04/2024
@version 1.0
    @return Nil, Função não tem retorno
    @example
    u_VOLUME()
    @obs Não se pode executar função MVC dentro do fórmulas
/*/

User Function VOLUME()
	Local cTipo    := ""
	Local nVolume  := 0
	Local nVol	   := 0
	Local nPProdL  := 0
	Local nPProdB  := 0
	Local nPesoL   := 0
	Local nPesoB   := 0
	Local nConv    := 0
	Local nPosProd := Ascan(aHeader, {|x| AllTrim(x[2]) == "C6_PRODUTO"})
	Local nPosQTD  := Ascan(aHeader , { |x| Alltrim(x[2]) = "C6_QTDVEN"})
	Local lUsaConv := SuperGetMV("BO_USACONV",.F.,.F.)
	Local i, x

	If lUsaConv
		For i:=1 to len(aCols)
			cTipo := POSICIONE("SB1",1,FWxFilial("SB1")+aCols[i][nPosProd],"B1_TIPCONV")
			nConv := POSICIONE("SB1",1,FWxFilial("SB1")+aCols[i][nPosProd],"B1_CONV")
			nPProdL := POSICIONE("SB1",1,FWxFilial("SB1")+aCols[i][nPosProd],"B1_PESO")
			nPProdB := POSICIONE("SB1",1,FWxFilial("SB1")+aCols[i][nPosProd],"B1_PESBRU")

			IF cTipo == 'M' .And. !Empty(nConv)
				nVolume	:= int(aCols[i,nPosQTD] *nConv)

				nVol += iif(nVolume <> aCols[i,nPosQTD] *nConv,nVolume+1,nVolume)
				nPesoL += aCols[i,nPosQTD] * nPProdL
				nPesoB += aCols[i,nPosQTD] * nPProdB

			ELSEIF cTipo == 'D' .And. !Empty(nConv) .And. lUsaConv
				nVolume	:= int(aCols[i,nPosQTD] /nConv)

				nVol += iif(nVolume <> aCols[i,nPosQTD] /nConv,nVolume+1,nVolume)
				nPesoL += aCols[i,nPosQTD] * nPProdL
				nPesoB += aCols[i,nPosQTD] * nPProdB

			ELSEIF Empty(nVol)

				For x:=1 to len(aCols)
					nVol += aCols[x,nPosQTD]
				Next x

			ENDIF
		Next i
	Else

		For i:=1 to len(aCols)

			nPProdL := POSICIONE("SB1",1,FWxFilial("SB1")+aCols[i][nPosProd],"B1_PESO")
			nPProdB := POSICIONE("SB1",1,FWxFilial("SB1")+aCols[i][nPosProd],"B1_PESBRU")
			nPesoL += aCols[i,nPosQTD] * nPProdL
			nPesoB += aCols[i,nPosQTD] * nPProdB

			nVol += aCols[i,nPosQTD]
		Next i

	EndIf

	nVol := Round(nVol,TAMSX3("C5_VOLUME1")[2])

	M->C5_PESOL := nPesoL
	M->C5_PBRUTO := nPesoB
	M->C5_VOLUME1 := nVol

Return nVol
