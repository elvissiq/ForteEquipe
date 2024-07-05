#INCLUDE 'totvs.ch'

/*/{Protheus.doc} VALFINREC
@type function
@version 
@author TOTVS Nordeste
@since 29/11/2023
@return
/*/

User Function VALFINREC()
    Local aTab := {"SE1 - Contas a Receber","SE2 - Contas a Pagar","SA1 - Clientes"}
    Local oDialog, oPanel, oTSay, oCombo, oDlg

    Private cTabela := ""
    Private cCampo  := Space(10)

    oDialog := FWDialogModal():New()
    oDialog:SetBackground( .T. ) 
    oDialog:SetTitle( 'Selecione a tabela abaixo, e informe o campo a ser atualizado.' )
    oDialog:SetSize( 160, 250 )
    oDialog:EnableFormBar( .T. )
    oDialog:SetCloseButton( .F. )
    oDialog:SetEscClose( .F. )
    oDialog:CreateDialog()
    oDialog:CreateFormBar()
    oDialog:AddCloseButton(Nil, "Confirmar")

    oPanel := oDialog:GetPanelMain()

        oTSay  := TSay():New(10,5,{|| "Tabela: "},oPanel,,,,,,.T.,,,50,70,,,,,,.T.)
        oCombo := TComboBox():New(29,28,{|u|iif(PCount()>0,cTabela:=u,cTabela)},aTab,100,20,oDlg,,{||},,,,.T.,,,,,,,,,'cTabela')
        
        oTSay  := TSay():New(25,5,{|| "Campo: "},oPanel,,,,,,.T.,,,50,70,,,,,,.T.)
        @ 23,25 MSGET cCampo SIZE 040,009 OF oPanel PIXEL

        oTSay  := TSay():New(50,5, {|| "Observações da Rotina: "},oPanel,,,,,,.T.,,,200,50,,,,,,.T.)
        oTSay  := TSay():New(58,5, {|| "Está rotina irá utilizar como chave o indice abaixo: "},oPanel,,,,,,.T.,,,200,50,,,,,,.T.)
        oTSay  := TSay():New(70,20,{|| '1 - Títulos a Receber e Pagar (Filial + Prefixo + Número do Título + Parcela + Tipo)'},oPanel,,,,,,.T.,,,200,50,,,,,,.T.)
        oTSay  := TSay():New(78,20,{|| '2 - Clientes (Filial + Código + Loja)'},oPanel,,,,,,.T.,,,200,50,,,,,,.T.)
        oTSay  := TSay():New(90,5, {|| 'Os campos sitados acima devem constar no arquivo para que seja possível o posicionamento no registro.'},oPanel,,,,,,.T.,,,200,50,,,,,,.T.)

    oDialog:Activate()

    cTabela := SubStr( cTabela, 1, At('-', cTabela) - 2)
    cCampo  := Alltrim(UPPER(cCampo))

    If !Empty(cTabela) .AND. !Empty(cCampo)
        Processa({|| xProcessa()}, "Atualizando Registros...")
    EndIF 

Return

Static Function xProcessa()
    Local oExcel
    Local aTamLin
    Local nContP,nContL
    Local cArq 

    cArq := tFileDialog( "Arquivo de planilha Excel (*.xlsx) | Todos tipos (*.*)",,,, .F., /*GETF_MULTISELECT*/)
    oExcel	:= YExcel():new(,cArq)

    If !Empty(cArq)
        DBSelectArea(cTabela)
        &(cTabela)->(DBSetOrder(1))

        For nContP := 1 To oExcel:LenPlanAt()	        //Ler as Planilhas
            oExcel:SetPlanAt(nContP)		            //Informa qual a planilha atual
            ConOut("Planilha:"+oExcel:GetPlanAt("2"))	//Nome da Planilha
            aTamLin	:= oExcel:LinTam() 		            //Linha inicio e fim da linha
            ProcRegua(aTamLin[2])
            For nContL := aTamLin[1] To aTamLin[2]

                aTamCol	:= oExcel:ColTam(nContL)

                IncProc("Gravando alteração " + cValToChar(nContL) + " de " + cValToChar(aTamLin[2]) + "...")
                
                If aTamCol[1] > 0

                    &(cTabela)->(DBGoTop())
                    If cTabela $('SA1')
                        IF &(cTabela)->(MsSeek(Pad(oExcel:GetValue(nContL,1),TamSX3(SubSTR(cTabela,2)+"_FILIAL")[1])+;
                                        Pad(oExcel:GetValue(nContL,2),TamSX3(SubSTR(cTabela,2)+"_COD")[1])+;
                                        Pad(oExcel:GetValue(nContL,3),TamSX3(SubSTR(cTabela,2)+"_LOJA")[1])))
                            
                            RecLock(cTabela,.F.)
                                &(cTabela)->(&(cCampo)) := Alltrim(oExcel:GetValue(nContL,4))
                            &(cTabela)->(MsUnLock())
                        
                        Else 
                            FWAlertError("Linha "+cValToChar(nContL),"Não Posicionou!")
                        EndIF

                    ElseIf cTabela $('SE1/SE2')
                        IF &(cTabela)->(MsSeek(Pad(oExcel:GetValue(nContL,1),TamSX3(SubSTR(cTabela,2)+"_FILIAL")[1])+;
                                        Pad(oExcel:GetValue(nContL,2),TamSX3(SubSTR(cTabela,2)+"_PREFIXO")[1])+;
                                        Pad(oExcel:GetValue(nContL,3),TamSX3(SubSTR(cTabela,2)+"_NUM")[1])+;
                                        Pad(oExcel:GetValue(nContL,4),TamSX3(SubSTR(cTabela,2)+"_PARCELA")[1])+;
                                        Pad(oExcel:GetValue(nContL,5),TamSX3(SubSTR(cTabela,2)+"_TIPO")[1])))
                            
                            RecLock(cTabela,.F.)
                                &(cTabela)->(&(cCampo))  := Alltrim(oExcel:GetValue(nContL,6))
                            &(cTabela)->(MsUnLock())
                        
                        Else 
                            FWAlertError("Linha "+cValToChar(nContL),"Não Posicionou!")
                        EndIF
                    EndIF 

                EndIf
            
            Next
        Next 
    EndIf 

    oExcel:Close()

Return
