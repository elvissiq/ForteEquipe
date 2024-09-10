#INCLUDE "Protheus.ch"
#INCLUDE "FWMVCDEF.ch"
#INCLUDE "TopConn.ch"
#INCLUDE "Tbiconn.ch"

Static cTitulo := "Importação - Arquivo de Conciliação"
Static __oRegiFIF := Nil

/*/{Protheus.doc} FFIN001
Leitura/Importação do arquivo de Conciliação CIELO
@type function
@author Elvis Siqueira
@since 09/09/2024
/*/
User Function FFIN001()

    Processa({|| FIN1Proc()}, "Lendo arquivo...")
    
Return

Static Function FIN1Proc()
    Local aArea := FWGetArea()
    Local oExcel
    Local aTamLin
    Local nContP,nContL
    Local cArq := ""
    Local cSeqFIF := ""
    Local nTamNUCOMP := FwTamSx3("FIF_NUCOMP")[1]
    Local nTamNSUTEF := FwTamSx3("FIF_NSUTEF")[1]
    Local nTamPARCEL := FwTamSx3("FIF_PARCEL")[1]
    Local cTPVenda := ""
    Local cNSU := ""
    Local nParcel := 0
    Local nValBru := 0
    Local nValLiq := 0
    Local nValCom := 0
    Local nTXServ := 0
    Local cQry    := ""
    Local _cAlias := GetNextAlias()
    Local dDtTEF  := CToD("//")
    Local cCarAut := ""

    cArq := tFileDialog( "Arquivo de planilha Excel (*.xlsx)",,,, .F., /*GETF_MULTISELECT*/)

    If !Empty(cArq)

        If SubSTR(cArq,RAT(".",cArq)) <> '.xlsx'
            FWAlertWarning('O arquivo não está no formato "xlsx".','Tipo do Arquivo')
            FWRestArea(aArea)
            Return
        EndIF 
        
        oExcel	:= YExcel():new(,cArq)
        
        If F914aExis(cArq)
            MsgStop("O arquivo: "+Alltrim(cArq)+", já foi importado anteriormente.")
            Return
        EndIF
        
        DBSelectArea("FVR")
        DBSelectArea("FV3")
        DBSelectArea("FIF")

        For nContP := 1 To oExcel:LenPlanAt()	        //Ler as Planilhas
            oExcel:SetPlanAt(nContP)		            //Informa qual a planilha atual
            aTamLin	:= oExcel:LinTam() 		            //Linha inicio e fim da linha
            ProcRegua(aTamLin[2])
            For nContL := aTamLin[1] to aTamLin[2]
                
                IncProc("Gravando tabela FIF, arquivo na linha " + cValToChar(nContL) + " de " + cValToChar(aTamLin[2]) + "...")
      
                If nContL > 5
                    
                    cNSU := IIF(ValType(oExcel:GetValue(nContL,22)) == "N", AsString(oExcel:GetValue(nContL,22)), oExcel:GetValue(nContL,22))
                    nParcel := IIF(ValType(oExcel:GetValue(nContL,16)) == "N", AsString(oExcel:GetValue(nContL,16)), oExcel:GetValue(nContL,16))

                    cTPVenda := ""
                    If Alltrim(oExcel:GetValue(nContL,2)) == "Venda parcelada" .OR.  Alltrim(oExcel:GetValue(nContL,2)) == "Venda crédito"
                        cTPVenda := "C"
                    ElseIF Alltrim(oExcel:GetValue(nContL,2)) == "Venda débito" .OR.  Alltrim(oExcel:GetValue(nContL,2)) == "Ajuste a débito"
                        cTPVenda := "D"
                    EndIF

                    If !Empty(cTPVenda) .AND. Alltrim(oExcel:GetValue(nContL,1)) == "Pago"
                
                        cSeqFIF := proxIdFIF()

                        If FVR->(MSSeek(xFilial("FVR")+Pad(cSeqFIF,FwTamSX3("FVR_IDPROC")[1])))
                            RecLock('FVR',.F.)
                                DbDelete()
                            FVR->(MsUnlock())
                        EndIF

                        RecLock('FVR',.T.)
                            FVR_FILIAL := xFilial("FVR") 
                            FVR_DESCLE := "1"
                            FVR_IDPROC := cSeqFIF
                            FVR_NOMARQ := Alltrim(cArq)
                            FVR_DTPROC := dDataBase
                            FVR_HRPROC := Time()
                            FVR_QTDPRO := 0
                            FVR_QTDALT := 0
                            FVR_QTDINC := 0
                            FVR_QTDLIN := 0
                            FVR_QTDTOT := 0
                            FVR_NOMUSU := USRRETNAME(__cUserId)
                            FVR_STATUS := ""
                            FVR_CODUSU := __cUserId
                            FVR_NOMADM := ""
                            FVR_CODADM := ""
                            FVR_MODPAG := ""
                        FVR->(MsUnlock())

                        If FV3->(MSSeek(xFilial("FV3")+Pad(cSeqFIF,FwTamSX3("FV3_IDPROC")[1])+Pad(AsString(nContL),FwTamSX3("FV3_LINARQ")[1])))
                            RecLock('FV3',.F.)
                                DbDelete()
                            FV3->(MsUnlock())
                        EndIF

                        RecLock('FV3',.T.)
                            FV3->FV3_FILIAL := xFilial("FV3") 
                            FV3->FV3_IDPROC := cSeqFIF
                            FV3->FV3_LINARQ := AsString(nContL)
                            FV3->FV3_NOMARQ := Alltrim(cArq)
                            FV3->FV3_DTPROC := dDataBase
                            FV3->FV3_HRPROC := Time()
                            FV3->FV3_CODEST := ""
                            FV3->FV3_NUCOMP := ""
                            FV3->FV3_MOTIVO := ""
                        FV3->(MsUnlock())
                        
                        lGrvFIF := .T.

                        If FIF->(MSSeek(xFilial("FIF")+PadL(cNSU, nTamNSUTEF, '0')+PadL(AsString(nParcel), nTamPARCEL, '0')))
                            If Empty(FIF->FIF_NUM)
                                RecLock('FIF',.F.)
                                    DbDelete()
                                FIF->(MsUnlock())
                            Else 
                                lGrvFIF := .F.
                            EndIF 
                        EndIF 

                        If lGrvFIF 

                            If ValType(oExcel:GetValue(nContL,23)) == "N"
                                nValBru := oExcel:GetValue(nContL,23)
                            Else
                                nValBru := STRTran(oExcel:GetValue(nContL,23),".","")
                                nValBru := STRTran(nValBru,",",".")
                                nValBru := Val(nValBru)
                            EndIF
                            
                            If ValType(oExcel:GetValue(nContL,25)) == "N"
                                nValLiq := oExcel:GetValue(nContL,25)
                            Else
                                nValLiq := STRTran(oExcel:GetValue(nContL,25),".","")
                                nValLiq := STRTran(nValLiq,",",".")
                                nValLiq := Val(nValLiq)
                            EndIF
                            
                            If ValType(oExcel:GetValue(nContL,24)) == "N"
                                nValCom := oExcel:GetValue(nContL,24)
                            Else
                                nValCom := STRTran(oExcel:GetValue(nContL,24),".","")
                                nValCom := STRTran(nValCom,",",".")
                                nValCom := Val(nValCom)
                            EndIF

                            If ValType(oExcel:GetValue(nContL,33)) == "N"
                                nTXServ := oExcel:GetValue(nContL,33)
                            Else
                                nTXServ := STRTran(oExcel:GetValue(nContL,33),".","")
                                nTXServ := STRTran(nTXServ,",",".")
                                nTXServ := Val(nTXServ)
                            EndIF

                            dDtTEF := CToD("//")
                            cCarAut := ""

                            cQry := " SELECT E1_EMISSAO, E1_VENCREA, E1_CARTAUT "
                            cQry += " FROM " + RetSqlName("SE1")
                            cQry += " WHERE D_E_L_E_T_ <> '*' "
                            cQry += " AND E1_FILIAL = '" + xFilial("SE1") + "' "
                            cQry += " AND E1_PARCELA = '" + PadL(IIF(ValType(oExcel:GetValue(nContL,16)) == "N", AsString(oExcel:GetValue(nContL,16)), oExcel:GetValue(nContL,16)), nTamPARCEL, '0') + "' "
                            cQry += " AND E1_NSUTEF LIKE ('%" + IIF(ValType(oExcel:GetValue(nContL,22)) == "N", AsString(oExcel:GetValue(nContL,22)), oExcel:GetValue(nContL,22)) + "%') "
                            cQry += " AND E1_CARTAUT LIKE ('%" + IIF(ValType(oExcel:GetValue(nContL,21)) == "N", AsString(oExcel:GetValue(nContL,21)), oExcel:GetValue(nContL,21)) + "%') "
                            IF Select(_cAlias) <> 0
                                (_cAlias)->(DbCloseArea())
                            EndIf
                            TCQuery cQry New Alias &_cAlias
                            IF (_cAlias)->(!Eof())
                                dDtTEF := SToD((_cAlias)->E1_EMISSAO)
                                cCarAut := AllTrim((_cAlias)->E1_CARTAUT)
                            Else
                                dDtTEF := IIF(ValType(oExcel:GetValue(nContL,13)) == "D", oExcel:GetValue(nContL,13), CToD(oExcel:GetValue(nContL,13)))
                                cCarAut := IIF(ValType(oExcel:GetValue(nContL,21)) == "N", AsString(oExcel:GetValue(nContL,21)), oExcel:GetValue(nContL,21))
                            EndIF
                            (_cAlias)->(DbCloseArea())

                            RecLock('FIF',.T.)
                                FIF->FIF_FILIAL := xFilial("FIF") 
                                FIF->FIF_TPREG  := "10"
                                FIF->FIF_INTRAN := ""
                                FIF->FIF_CODEST := IIF(ValType(oExcel:GetValue(nContL,36)) == "N", AsString(oExcel:GetValue(nContL,36)), oExcel:GetValue(nContL,36))
                                FIF->FIF_DTTEF  := dDtTEF
                                FIF->FIF_NURESU := IIF(ValType(oExcel:GetValue(nContL,32)) == "N", AsString(oExcel:GetValue(nContL,32)), oExcel:GetValue(nContL,32))
                                FIF->FIF_NUCOMP := PadL(IIF(ValType(oExcel:GetValue(nContL,32)) == "N", AsString(oExcel:GetValue(nContL,32)), oExcel:GetValue(nContL,32)), nTamNUCOMP, '0')
                                FIF->FIF_NSUTEF := PadL(IIF(ValType(oExcel:GetValue(nContL,22)) == "N", AsString(oExcel:GetValue(nContL,22)), oExcel:GetValue(nContL,22)), nTamNSUTEF, '0')
                                FIF->FIF_NUCART := IIF(ValType(oExcel:GetValue(nContL,18)) == "N", AsString(oExcel:GetValue(nContL,18)), oExcel:GetValue(nContL,18))
                                FIF->FIF_VLBRUT := nValBru
                                FIF->FIF_TOTPAR := IIF(ValType(oExcel:GetValue(nContL,17)) == "N", AsString(oExcel:GetValue(nContL,17)), oExcel:GetValue(nContL,17))
                                FIF->FIF_VLLIQ  := nValLiq
                                FIF->FIF_DTCRED := IIF(ValType(oExcel:GetValue(nContL,11)) == "D", oExcel:GetValue(nContL,11), CToD(oExcel:GetValue(nContL,11)))
                                FIF->FIF_PARCEL := PadL(IIF(ValType(oExcel:GetValue(nContL,16)) == "N", AsString(oExcel:GetValue(nContL,16)), oExcel:GetValue(nContL,16)), nTamPARCEL, '0')
                                FIF->FIF_TPPROD := cTPVenda
                                FIF->FIF_CAPTUR := "1"
                                FIF->FIF_CODRED := "340"
                                FIF->FIF_CODBCO := "341"
                                FIF->FIF_CODAGE := IIF(ValType(oExcel:GetValue(nContL,5)) == "N", AsString(oExcel:GetValue(nContL,5)), oExcel:GetValue(nContL,5))
                                FIF->FIF_NUMCC  := IIF(ValType(oExcel:GetValue(nContL,6)) == "N", AsString(oExcel:GetValue(nContL,6)), oExcel:GetValue(nContL,6))
                                FIF->FIF_VLCOM  := nValCom
                                FIF->FIF_TXSERV := nTXServ
                                FIF->FIF_CODLOJ := IIF(ValType(oExcel:GetValue(nContL,10)) == "N", AsString(oExcel:GetValue(nContL,10)), oExcel:GetValue(nContL,10))
                                FIF->FIF_CODAUT := cCarAut
                                FIF->FIF_CUPOM  := IIF(ValType(oExcel:GetValue(nContL,43)) == "N", AsString(oExcel:GetValue(nContL,43)), oExcel:GetValue(nContL,43))
                                FIF->FIF_SEQREG := PadL(AsString(nContL), 6, '0')
                                FIF->FIF_DTAJST := STOD("")
                                FIF->FIF_CODMAJ := ""
                                FIF->FIF_STATUS := "1"
                                FIF->FIF_DTBAIX := STOD("")
                                FIF->FIF_DTIMP  := dDataBase
                                FIF->FIF_USERGA := ""
                                FIF->FIF_MSIMP  := DTOS(dDataBase)
                                FIF->FIF_PREFIX := ""
                                FIF->FIF_NUM    := ""
                                FIF->FIF_PARC   := ""
                                FIF->FIF_TIPO   := ""
                                FIF->FIF_PARALF := IIF(ValType(oExcel:GetValue(nContL,16)) == "N", AsString(oExcel:GetValue(nContL,16)), oExcel:GetValue(nContL,16))
                                FIF->FIF_CODFIL := xFilial("FIF")
                                FIF->FIF_CODBAN := "" //Posicione("MDE",2,xFilial("MDE")+Pad(UPPER(oModelGrid:GetValue("BANDEIR"))),"MDE_CODIGO")
                                FIF->FIF_SEQFIF := cSeqFIF
                                FIF->FIF_DTANT  := STOD("")
                                FIF->FIF_STVEND := ""
                                FIF->FIF_DTCONV := STOD("")
                                FIF->FIF_CODJUS := ""
                                FIF->FIF_DESJUS := ""
                                FIF->FIF_DESJUT := ""
                                FIF->FIF_CODADM := ""
                                FIF->FIF_DTVEN  := STOD("")
                                FIF->FIF_DTPAG  := STOD("")
                                FIF->FIF_USUVEN := ""
                                FIF->FIF_USUPAG := ""
                                FIF->FIF_ARQVEN := ""
                                FIF->FIF_PGJUST := ""
                                FIF->FIF_PGDES1 := ""
                                FIF->FIF_ARQPAG := Alltrim(cArq)
                                FIF->FIF_NSUARQ := ""
                                FIF->FIF_PGDES2 := ""
                                FIF->FIF_IDORAJ := ""
                                FIF->FIF_MSFIL  := xFilial("FIF")
                                FIF->FIF_MODPAG := ""
                            FIF->(MsUnlock())
                        EndIf
                    EndIF
                EndIF
            Next 
        Next
        
        oExcel:Close() 

    EndIF 

    FWRestArea(aArea)

Return

//--------------------------------------------------------------------------
/*/{Protheus.doc} proxIdFIF
Retorna próxima sequencia para a tabela FIF	

@author Elvis Siqueira
@type  Static Function
@since 04/12/2023
@version 1.0
@return cSeqFIF, character, retorna próxima sequencia para a tabela FIF	
/*/
//--------------------------------------------------------------------------
Static Function proxIdFIF() As Character
	Local aArea  		As Array
	Local aAreaFIF 		As Array
	Local cQryFIF		As Character
	Local cSeqFIF		As Character
	Local cTRBFIF		As Character

	aArea  		:= FWGetArea()
	cTRBFIF		:= GetNextAlias()
	cSeqFIF 	:= StrZero(1, 6)
	aAreaFIF 	:= FIF->(FWGetArea())

	cQryFIF := " SELECT MAX(FIF_SEQFIF) MAXFIF"
	cQryFIF += " FROM " + RetSqlName("FIF")
	cQryFIF += " WHERE D_E_L_E_T_ = ''"

	cQryFIF := ChangeQuery(cQryFIF)

	DbUseArea(.T., "TOPCONN", TcGenQry(,, cQryFIF), cTRBFIF)

	If (cTRBFIF)->(!Eof()) .And. !Empty((cTRBFIF)->(MAXFIF))
		cSeqFIF := Soma1((cTRBFIF)->(MAXFIF))
	EndIf

	(cTRBFIF)->(DbCloseArea())

    FWRestArea(aAreaFIF)
	FWRestArea(aArea)
	
Return cSeqFIF

/*{Protheus.doc} F914aExis
	Verifica se já existe o registro na base

	@author Elvis Siqueira
	@since 04/12/2023
	@version 1.0
*/
Static Function F914aExis(cArquivo As Character) As Logical
	Local lRet   As Logical
	Local cQuery As Character
	
	Default cArquivo := ""
	
	//Inicializa variáveis
	cQuery     := ""
	lRet	   := .F.
	
	If __oRegiFIF == Nil
		cQuery := "SELECT FIF.R_E_C_N_O_ NREGFIF " 
		cQuery += "FROM ? FIF WHERE "		
		cQuery += "FIF.FIF_ARQPAG = ? "
		cQuery += "AND FIF.D_E_L_E_T_ = ' ' "
        cQuery := ChangeQuery(cQuery)
		__oRegiFIF := FWPreparedStatement():New(cQuery)		
	EndIf
	
	__oRegiFIF:SetNumeric(1, RetSqlName("FIF"))
	__oRegiFIF:SetString(2, cArquivo)
	cQuery := __oRegiFIF:GetFixQuery()
	
	lRet := (MpSysExecScalar(cQuery, "NREGFIF") > 0)

Return lRet
