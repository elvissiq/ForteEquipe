//Bibliotecas
#INCLUDE "Protheus.ch"
#INCLUDE "TOTVS.ch"
#INCLUDE "Topconn.ch"
 
/*------------------------------------------------------------------------------------------------------*
 | Funcao:  ATUALIAS                                                                                    |
 | Descricao:  Funcao utilizada para atualizar o nome das tabelas no banco de dados                     |
 *------------------------------------------------------------------------------------------------------*/
 
User Function ATUALIAS()
    Private  oDialog, oPanel, oTSay
    Private cAntGrp := Space(3)
    Private cNovGrp := Space(3)
    Private cExecut := FWUUIDV4()
    Private lBtOK   := .F.

    oDialog := FWDialogModal():New()
    oDialog:SetBackground( .T. ) 
    oDialog:SetTitle( 'Informe o Grupo de Empresas' )
    oDialog:SetSize( 100, 180 )
    oDialog:EnableFormBar( .T. )
    oDialog:SetCloseButton( .T. )
    oDialog:SetEscClose( .T. )
    oDialog:CreateDialog()
    oDialog:CreateFormBar()
    oDialog:addCloseButton(Nil, "Fechar")
    oDialog:addOkButton({|| fButtomOk() },'Confirmar')

    oPanel := oDialog:GetPanelMain()

        oTSay  := TSay():New(10,5,{|| "Código Antigo"},oPanel,,,,,,.T.,,,80,70,,,,,,.T.)
        @ 08,50 MSGET cAntGrp SIZE 030,009 OF oPanel PIXEL
        
        oTSay  := TSay():New(35,5,{|| "Código Novo"},oPanel,,,,,,.T.,,,80,70,,,,,,.T.)
        @ 33,50 MSGET cNovGrp SIZE 030,009 OF oPanel PIXEL
        
    oDialog:Activate()

    If lBtOK .And. !Empty(cAntGrp) .And. !Empty(cNovGrp)
        Processa({|| fProcess()}, "Atualizando...")
    ElseIF lBtOK
        FWAlertWarning('Os dois campos devem ser preenchidos com seus respectivos códigos.','Atualiza Tabelas')
        U_ATUALIAS()
    EndIF 

Return 
/*/{Protheus.doc} fButtomOk
    Botão OK 
/*/
Static Function fButtomOk()
    lBtOK := .T.
    oDialog:DeActivate()
Return 

/*/{Protheus.doc} fProcess
    Faz o processamento 
/*/
Static Function fProcess()
    Local aArea := FWGetArea()
    Local aSM0Data := FWLoadSM0()
    Local _cAliasQry := GetNextAlias()
    Local cQry := ""
    Local cQryExec := ""
    Local nAtual := 0
    Local nTotal := 0
    Local nStatus := 0
    Local lAchou := .F.
    Local nY 

    For nY := 1 To Len(aSM0Data)
        lAchou := IIF(aSM0Data[nY][1] == SubSTR(cNovGrp,1,2), .T., .F. )
        If lAchou
            Exit
        EndIF
    Next 
    
    If !lAchou
        FWAlertWarning('Atualização não será realizada, pois não foi encontrado no cadastro de Empresa o grupo ' + SubSTR(cNovGrp,1,2) + '.' + CRLF + CRLF + " Revise os códigos digitados", 'Atualiza Tabelas')
        Return
    ElseIF !(MsFile("SX2"+cAntGrp))
        FWAlertWarning('Atualização não será realizada, pois não foi encontrada a tabela SX2' + cAntGrp + '.' + CRLF + CRLF + " Revise os códigos digitados", 'Atualiza Tabelas')
        Return
    EndIF

    cQry := " SELECT * FROM SX2" + cAntGrp + ""
    cQry += " WHERE D_E_L_E_T_ <> '*'"
    cQry := ChangeQuery(cQry)
    IF Select(_cAliasQry) <> 0
        (_cAliasQry)->(DBCloseArea())
    EndIf
    dbUseArea(.T.,"TOPCONN",TCGenQry(,,cQry),_cAliasQry,.F.,.T.)
    Count To nTotal
    ProcRegua(nTotal)

    (_cAliasQry)->(DbGoTop())
    While (_cAliasQry)->(!Eof())
        
        nAtual++
        IncProc("Atualizando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")

        IF MsFile(AllTrim((_cAliasQry)->X2_ARQUIVO))
            
            cQryExec := "SELECT * INTO " + AllTrim((_cAliasQry)->X2_ARQUIVO) + "_BKP FROM " + AllTrim((_cAliasQry)->X2_ARQUIVO)
            nStatus := TCSQLExec(cQryExec)
            IF nStatus >= 0 
                
                cQryExec := "DROP TABLE " + AllTrim((_cAliasQry)->X2_ARQUIVO)
                nStatus := TCSQLExec(cQryExec)
                
                IF nStatus >= 0
                    
                    cQryExec := " UPDATE SX2"+ cAntGrp + " SET X2_ARQUIVO = '" + AllTrim((_cAliasQry)->X2_CHAVE + cNovGrp ) + "' "
                    cQryExec += " WHERE X2_CHAVE = '" + (_cAliasQry)->X2_CHAVE + "' "    
                    nStatus := TCSQLExec(cQryExec)

                    DBSelectArea((_cAliasQry)->X2_CHAVE)

                    cQryExec := "INSERT INTO " + AllTrim((_cAliasQry)->X2_ARQUIVO) + " FROM " + AllTrim((_cAliasQry)->X2_CHAVE + cAntGrp ) + "_BKP"
                    nStatus := TCSQLExec(cQryExec)
                    IF nStatus < 0
                        fGrvLog(TCSQLError(),(_cAliasQry)->X2_CHAVE,nAtual,nTotal,cQryExec)
                    EndIF 
                Else
                    fGrvLog(TCSQLError(),(_cAliasQry)->X2_CHAVE,nAtual,nTotal,cQryExec)
                EndIF 
            Else
                fGrvLog(TCSQLError(),(_cAliasQry)->X2_CHAVE,nAtual,nTotal,cQryExec)
            EndIF 
        Else
            cQryExec := " UPDATE SX2"+ cAntGrp + " SET X2_ARQUIVO = '" + AllTrim((_cAliasQry)->X2_CHAVE + cNovGrp ) + "' "
            cQryExec += " WHERE X2_CHAVE = '" + (_cAliasQry)->X2_CHAVE + "' "
            
            nStatus := TCSQLExec(cQryExec)
            IF nStatus < 0 
                fGrvLog(TCSQLError(),(_cAliasQry)->X2_CHAVE,nAtual,nTotal,cQryExec)
            EndIF
        EndIF 
    (_cAliasQry)->(DbSkip())
    EndDo

    IF Select(_cAliasQry) <> 0
        (_cAliasQry)->(DBCloseArea())
    EndIf

    FWRestArea(aArea)

Return

//====================================================================================================================\
/*/{Protheus.doc}fGrvLog
  ====================================================================================================================
	@description
	Grava o LOG na tabela XXE
/*/
//===================================================================================================================\
Static Function fGrvLog(cLog,cTabela,nAtual,nTotal,cQryExec)

    TCLink()
        cQryInsert := " INSERT INTO " + RetSqlName("XXE") 
        cQryInsert += " ( XXE_ID,   " 
        cQryInsert += " XXE_ADAPT,  "
        cQryInsert += " XXE_FILE,   " 
        cQryInsert += " XXE_LAYOUT, "
        cQryInsert += " XXE_DESC,   "
        cQryInsert += " XXE_DATE,   "
        cQryInsert += " XXE_TIME,   "
        cQryInsert += " XXE_TYPE,   "
        cQryInsert += " XXE_ERROR,  "
        cQryInsert += " XXE_USRID,  "
        cQryInsert += " XXE_USRNAM, "
        cQryInsert += " XXE_COMPLE, "
        cQryInsert += " XXE_ORIGIN, "
        cQryInsert += " XXE_IDOPER, "
        cQryInsert += " XXE_XML )   "
        cQryInsert += " VALUES (    "
        cQryInsert += " '"+XXEProx()+"',"
        cQryInsert += " '"+FunName()+"',"
        cQryInsert += " " +cExecut+ ","
        cQryInsert += " '"+cTabela+"',"
        cQryInsert += " '"+FWX2Nome(cTabela)+"',"
        cQryInsert += " '"+DToS(dDataBase)+"',"
        cQryInsert += " '"+Time()+"',"
        cQryInsert += " '2',"
        cQryInsert += " '"+cLog+"',"
        cQryInsert += " '"+__cUserID+"',"
        cQryInsert += " '"+cUserName+"',"
        cQryInsert += " '"+cLog+"',"
        cQryInsert += " '"+cValToChar(nAtual)+"-"+cValToChar(nTotal)+"',"
        cQryInsert += " '"+FWTimeStamp(1)+"',"
        cQryInsert += " '"+cQryExec+"')"
        nStatus := TCSqlExec(cQryInsert)
    TCUnlink()

Return

//====================================================================================================================\
/*/{Protheus.doc}XXEProx
  ====================================================================================================================
	@description
	Retorna o próximo número para a tabela XXE
/*/
//===================================================================================================================\
Static Function XXEProx()

	Local cRet := StrZero(0,10)
	Local cQry := ''
	Local cAli := GetNextAlias()

	cQry+= " SELECT MAX(XXE_ID) XXE_ID "
	cQry+= " FROM " + RetSqlTab('XXE')
	cQry+= " WHERE " + RetSqlCond('XXE')

	cQry:= ChangeQuery(cQry)

	If Select(cAli) <> 0
		(cAli)->(DbCloseArea())
	EndIf

	dbUseArea(.T.,'TOPCONN', TCGenQry(,,cQry),cAli, .F., .T.)

	If (cAli)->(!Eof())
		cRet:= (cAli)->XXE_ID
	EndIf

	If Select(cAli) <> 0
		(cAli)->(DbCloseArea())
	EndIf

	cRet:= Soma1(cRet)

Return ( cRet )
