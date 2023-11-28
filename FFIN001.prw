#INCLUDE "Totvs.ch"

/*/{Protheus.doc} FFIN001
Leitura do arquivo de Rateio
@type function
@author Elvis Siqueira
@since 28/11/2023
/*/

User Function FFIN001()
    Local aArea := GetArea()
    Local oExcel
    Local aTamLin
    Local nContP,nContL,nContC
    Local xValor
    Local cArq
    Local aButtons := {;
                    {.F.,Nil},;         // Copiar
                    {.F.,Nil},;         // Recortar
                    {.F.,Nil},;         // Colar
                    {.F.,Nil},;         // Calculadora
                    {.F.,Nil},;         // Spool
                    {.F.,Nil},;         // Imprimir
                    {.T.,"Confirmar"},; // Confirmar
                    {.T.,"Cancelar"},;  // Cancelar
                    {.F.,Nil},;         // WalkTrhough
                    {.F.,Nil},;         // Ambiente
                    {.F.,Nil},;         // Mashup
                    {.F.,Nil},;         // Help
                    {.F.,Nil},;         // Formul�rio HTML
                    {.F.,Nil};          // ECM
                  }
    
    Private oTabTMPCab
    Private oTabTMPGrid
    Private oTabTMPTit
    Private aFieldsCab  := {}
    Private aFieldsGrid := {}
    Private aFieldsTit  := {}

    aAdd(aFieldsCab, {"DTIMPOR" ,"D",8  ,0,"Data Importa��o",""})
    aAdd(aFieldsCab, {"PERIODO" ,"C",40 ,0,"Per�odo",""})
    aAdd(aFieldsCab, {"ARQUIVO" ,"C",200,0,"Arquivo",""})
    
    oTabTMPCab:= FWTemporaryTable():New(cAliasCab)
    oTabTMPCab:SetFields(aFieldsCab)
    oTabTMPCab:Create()

	aAdd(aFieldsGrid, {"STATUS"  ,"C",10,0,"Status",""})
    aAdd(aFieldsGrid, {"TPLANC"  ,"C",20,0,"Tipo de Lan�amento",""})
    aAdd(aFieldsGrid, {"DESCRIC" ,"C",40,0,"Descri��o",""})
    aAdd(aFieldsGrid, {"BANCO"   ,"C",20,0,"Banco",""})
    aAdd(aFieldsGrid, {"AGENCIA" ,"C",7 ,0,"Ag�ncia",""})
    aAdd(aFieldsGrid, {"GRAVAME" ,"C",3,0,"Gravame",""})
    aAdd(aFieldsGrid, {"CNPJORI" ,"C",20,0,"CNPJ da institui��o origem da negocia��o",""})
    aAdd(aFieldsGrid, {"NOMEORI" ,"C",40,0,"Nome da institui��o origem da negocia��o",""})
    aAdd(aFieldsGrid, {"ESTABE"  ,"C",15,0,"Estabelecimento",""})
    aAdd(aFieldsGrid, {"DTPAGAM" ,"D",8 ,0,"Data de pagamento",""})
    aAdd(aFieldsGrid, {"DTLANCA" ,"D",8 ,0,"Data do lan�amento",""})
    aAdd(aFieldsGrid, {"DTAUTVE" ,"D",8 ,0,"Data da autoriza��o da venda",""})
    aAdd(aFieldsGrid, {"BANDEIR" ,"C",20,0,"Bandeira",""})
    aAdd(aFieldsGrid, {"FORMPAG" ,"C",25,0,"Forma de Pagamento",""})
    aAdd(aFieldsGrid, {"NPARCEL" ,"C",2 ,0,"N�mero da parcela",""})
    aAdd(aFieldsGrid, {"QPARCEL" ,"C",2 ,0,"Quantidade de parcelas",""})
    aAdd(aFieldsGrid, {"NCARTAO" ,"C",15,0,"N�mero do cart�o",""})
    aAdd(aFieldsGrid, {"CODTRAN" ,"C",25,0,"C�digo da transa��o",""})
    aAdd(aFieldsGrid, {"TID"     ,"C",25,0,"TID",""})
    aAdd(aFieldsGrid, {"CODAUTO" ,"C",10,0,"C�digo de autoriza��o",""})
    aAdd(aFieldsGrid, {"NSU"     ,"C",10,0,"NSU",""})
    aAdd(aFieldsGrid, {"VALBRUT" ,"N",16,2,"Valor bruto","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALDESC" ,"N",16,2,"Valor descontado","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALLIQ"  ,"N",16,2,"Valor l�quido","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALCOB"  ,"N",16,2,"Valor cobrado","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALPEND" ,"N",16,2,"Valor pendente","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALTOT"  ,"N",16,2,"Valor total","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"RECRAP"  ,"C",3 ,0,"Receba R�pido",""})
    aAdd(aFieldsGrid, {"TPCAPT"  ,"C",20,0,"Tipo de captura",""})
    aAdd(aFieldsGrid, {"RESOP"   ,"C",10,0,"Resumo da opera��o",""})
    aAdd(aFieldsGrid, {"TAXA"    ,"N",16,2,"Taxas (%)","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"TARIFA"  ,"N",16,2,"Tarifa","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"CODVEND" ,"C",30,0,"C�digo da venda",""})
    aAdd(aFieldsGrid, {"NUMMAQ"  ,"C",15,0,"N�mero da m�quina",""})
    aAdd(aFieldsGrid, {"PERCONS" ,"D",8 ,0,"Per�odo considerado",""})
    aAdd(aFieldsGrid, {"NUMOPER" ,"C",15,0,"N�mero da opera��o",""})
    aAdd(aFieldsGrid, {"TAXAANT" ,"N",16,2,"Taxa de antecipa��o","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"TAXAEMB" ,"N",16,2,"Taxa de embarque","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALENTR" ,"N",16,2,"Valor da entrada","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"NUMPEDI" ,"C",30,0,"N�mero do pedido",""})
    aAdd(aFieldsGrid, {"NUMNOTA" ,"C",10,0,"N�mero da nota fiscal",""})
    aAdd(aFieldsGrid, {"ID"      ,"C",20,0,"ID",""})
    aAdd(aFieldsGrid, {"DTPGCIE" ,"D",8 ,0,"Data de pagamento na conta Cielo",""})

    oTabTMPGrid:= FWTemporaryTable():New(cAliasGrid)
    oTabTMPGrid:SetFields(aFieldsGrid)
    oTabTMPGrid:Create()







    cArq := tFileDialog( "Arquivo de planilha Excel (*.xlsx) | Todos tipos (*.*)",,,, .F., /*GETF_MULTISELECT*/)
    oExcel	:= YExcel():new(,cArq)

    For nContP := 1 To oExcel:LenPlanAt()	        //Ler as Planilhas
        oExcel:SetPlanAt(nContP)		            //Informa qual a planilha atual
        ConOut("Planilha:"+oExcel:GetPlanAt("2"))	//Nome da Planilha
        aTamLin	:= oExcel:LinTam() 		            //Linha inicio e fim da linha
        For nContL := aTamLin[1] to aTamLin[2]
            aTamCol	:= oExcel:ColTam(nContL)        //Coluna inicio e fim
            If aTamCol[1] > 0                       //Se a linha tem algum valor
                For nContC:=aTamCol[1] to aTamCol[2]
                    xValor	:= oExcel:GetValue(nContL,nContC)	//Conte�do 
                    If ValType(xValor)=="O"
                        ConOut(oExcel:Ref(nContL,nContC)+"["+cValToChar(xValor:GetDate())+"]["+cValToChar(xValor:GetTime())+"]")
                    Else
                        ConOut(oExcel:Ref(nContL,nContC)+"["+cValToChar(xValor)+"]")
                    EndIf
                Next
            EndIf
        Next
    Next

    oExcel:Close()

    If !Empty(nTotal)
        FWExecView(cTitulo,"PFINF02",MODEL_OPERATION_UPDATE,,{|| .T.},,50,aButtons)
    Else 
        APMsgAlert("Nenhuma movimenta��o encontrada com os parametros informados.","Aten��o")
    EndIF 

    oTabTMPCab:Delete()
    oTabTMPGrid:Delete()
    RestArea(aArea)

Return

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Cria��o do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ModelDef()
    Local oStructCab := FWFormModelStruct():New()
    Local oStructGrd := FWFormModelStruct():New()
    Local oModel
    Local nId

    oStructCab:AddTable(cAliasCab,{},"Tabela Cabe�alho")
    For nId := 1 To Len(aFieldsCab)
        oStructCab:AddField(aFieldsCab[nId][5]; 
                        ,aFieldsCab[nId][5]; 
                        ,aFieldsCab[nId][1]; 
                        ,aFieldsCab[nId][2];
                        ,aFieldsCab[nId][3];
                        ,aFieldsCab[nId][4];
                        ,Nil,Nil,{},.F.,,.F.,.F.,.F.)
    Next nId

    oStructGrd:AddTable(cAliasGrid,{},"Tabela Grid")
    For nId := 1 To Len(aFieldsGrid)
        oStructGrd:AddField(aFieldsGrid[nId][5]; 
                        ,aFieldsGrid[nId][5]; 
                        ,aFieldsGrid[nId][1]; 
                        ,aFieldsGrid[nId][2];
                        ,aFieldsGrid[nId][3];
                        ,aFieldsGrid[nId][4];
                        ,Nil,Nil,{},.F.,,.F.,.F.,.F.)
    Next nId

    oModel := MPFormModel():New("FFIN001M", /*bPre*/, {|| zConfirma() }/*bPos*/, /*bCommit*/ , /*bCancel*/)
    oModel:AddFields("MASTER", /*cOwner*/, oStructCab)
    oModel:AddGrid("GRID", "MASTER", oStructGrd)
    oModel:SetDescription(cTitulo)
    oModel:GetModel("MASTER"):SetDescription(cTitulo)
    oModel:SetPrimaryKey({})

Return oModel

/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Desc:  Cria��o da vis�o MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
    Local oModel := FWLoadModel("PFINF02")
    Local oStructCab := FWFormViewStruct():New()
    Local oStructGrd := FWFormViewStruct():New()
    Local oView
    Local lAlt := .F.
    Local nId
    

    For nId := 1 To Len(aFieldsCab)
        oStructCab:AddField(aFieldsCab[nId][1],;   // 01 = Nome do Campo
                        StrZero(nId,2),;           // 02 = Ordem
                        aFieldsCab[nId][5],;       // 03 = T�tulo do campo
                        aFieldsCab[nId][5],;       // 04 = Descri��o do campo
                        Nil,;                      // 05 = Array com Help
                        aFieldsCab[nId][2],;       // 06 = Tipo do campo
                        aFieldsCab[nId][6],;       // 07 = Picture
                        Nil,;                      // 08 = Bloco de PictTre Var
                        Nil,;                      // 09 = Consulta F3
                        .F.,;                      // 10 = Indica se o campo � alter�vel
                        Nil,;                      // 11 = Pasta do Campo
                        Nil,;                      // 12 = Agrupamnento do campo
                        Nil,;                      // 13 = Lista de valores permitido do campo (Combo)
                        Nil,;                      // 14 = Tamanho m�ximo da op��o do combo
                        Nil,;                      // 15 = Inicializador de Browse
                        .F.,;                      // 16 = Indica se o campo � virtual (.T. ou .F.)
                        Nil,;                      // 17 = Picture Variavel
                        Nil)                       // 18 = Indica pulo de linha ap�s o campo (.T. ou .F.)
    Next nId

    For nId := 1 To Len(aFieldsGrid)
        
        If aFieldsGrid[nId][1] == "OK"
            lAlt := .T.
        EndIf 
        
        oStructGrd:AddField(aFieldsGrid[nId][1],;                   // 01 = Nome do Campo
                        StrZero(nId,2),;                            // 02 = Ordem
                        aFieldsGrid[nId][5],;                       // 03 = T�tulo do campo
                        aFieldsGrid[nId][5],;                       // 04 = Descri��o do campo
                        Nil,;                                       // 05 = Array com Help
                        aFieldsGrid[nId][2],;                       // 06 = Tipo do campo
                        aFieldsGrid[nId][6],;                       // 07 = Picture
                        Nil,;                                       // 08 = Bloco de PictTre Var
                        Nil,;                                       // 09 = Consulta F3
                        IIF(aFieldsGrid[nId][1] == "OK",.T.,.F.),;  // 10 = Indica se o campo � alter�vel
                        Nil,;                                       // 11 = Pasta do Campo
                        Nil,;                                       // 12 = Agrupamnento do campo
                        Nil,;                                       // 13 = Lista de valores permitido do campo (Combo)
                        Nil,;                                       // 14 = Tamanho m�ximo da op��o do combo
                        Nil,;                                       // 15 = Inicializador de Browse
                        .F.,;                                       // 16 = Indica se o campo � virtual (.T. ou .F.)
                        Nil,;                                       // 17 = Picture Variavel
                        Nil)                                        // 18 = Indica pulo de linha ap�s o campo (.T. ou .F.)
    Next nId

    oView := FWFormView():New()    
    oView:SetModel(oModel)
    oView:SetProgressBar(.T.)
    oView:AddField("VIEW1", oStructCab, "MASTER")
    oView:AddGrid("VIEW2" , oStructGrd,"GRID")

    oView:CreateHorizontalBox("CABEC" , 0 )
    oView:CreateHorizontalBox("GRID" , 100 )
    oView:SetOwnerView("VIEW1", "CABEC")
    oView:SetOwnerView("VIEW2", "GRID")

    oView:SetAfterViewActivate({|oView| ViewActv(oView)})

    oView:AddUserButton( 'Desmarcar Todos', 'NOTE',;
                        {|oModel| fDesmark(oView)},;
                         /*cToolTip  | Coment�rio do bot�o*/,;
                         /*nShortCut | Codigo da Tecla para cria��o de Tecla de Atalho*/,;
                         /*aOptions  | */,;
                         /*lShowBar */ .T.)
    
    oView:AddUserButton( 'Marcar Todos', 'NOTE',;
                        {|oModel| fMark(oView)},;
                         /*cToolTip  | Coment�rio do bot�o*/,;
                         /*nShortCut | Codigo da Tecla para cria��o de Tecla de Atalho*/,;
                         /*aOptions  | */,;
                         /*lShowBar */ .T.)
 
Return oView

/*---------------------------------------------------------------------*
 | Func:  ViewActv                                                     |
 | Desc:  Realiza o PUT nos campos para grava��o na tabela SE1         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ViewActv(oView)
Local oModel := FWModelActive() 
Local oModelGrid := oModel:GetModel("GRID")

    DBSelectArea(_cAlias)
    (_cAlias)->(DBGoTOP())

    While!(_cAlias)->(EOF())

        oModelGrid:AddLine()

        oModelGrid:LoadValue("OK"     , .T.   )
        oModelGrid:LoadValue("FILIAL" , (_cAlias)->E2_FILIAL        )
        oModelGrid:LoadValue("PREFIXO", (_cAlias)->E2_PREFIXO       )
        oModelGrid:LoadValue("NUM"    , (_cAlias)->E2_NUM           )
        oModelGrid:LoadValue("PARCELA", (_cAlias)->E2_PARCELA       )
        oModelGrid:LoadValue("TIPO"   , (_cAlias)->E2_TIPO          )
        oModelGrid:LoadValue("NATUREZ", (_cAlias)->E2_NATUREZ       )
        oModelGrid:LoadValue("FORNECE", (_cAlias)->E2_FORNECE       )
        oModelGrid:LoadValue("LOJA"   , (_cAlias)->E2_LOJA          )
        oModelGrid:LoadValue("NOMFOR" , (_cAlias)->E2_NOMFOR        )
        oModelGrid:LoadValue("EMISSAO", SToD((_cAlias)->E2_EMISSAO) )
        oModelGrid:LoadValue("VENCTO" , SToD((_cAlias)->E2_VENCTO)  )
        oModelGrid:LoadValue("VENCREA", SToD((_cAlias)->E2_VENCREA) )
        oModelGrid:LoadValue("VALOR"  , (_cAlias)->E2_VALOR         )
        oModelGrid:LoadValue("HIST"   , (_cAlias)->E2_HIST          )
        
        oView:Refresh('VIEW2')
    (_cAlias)->(DBSkip())
    EndDo
    
    oModelGrid:GoLine(1)
    oView:Refresh('VIEW2')
    oView:SetNoDeleteLine('VIEW2')
    oView:SetNoInsertLine('VIEW2')

Return
