#INCLUDE "Totvs.ch"

/*/{Protheus.doc} FFIN001
Leitura do arquivo de Rateio
@type function
@author Elvis Siqueira
@since 31/10/2023
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
                    {.F.,Nil},;         // Formulário HTML
                    {.F.,Nil};          // ECM
                  }
    
    Private oTabTMPCab
    Private oTabTMPGrid
    Private aFieldsCab := {}
    Private aFieldsGrid := {}

    aAdd(aFieldsCab, {"CODFOR"  ,GetSx3Cache("CV4_FILIAL","X3_TIPO"),GetSx3Cache("CV4_FILIAL","X3_TAMANHO"),GetSx3Cache("CV4_FILIAL","X3_DECIMAL"),GetSx3Cache("CV4_FILIAL","X3_TITULO"),GetSx3Cache("CV4_FILIAL","X3_PICTURE")})
    aAdd(aFieldsCab, {"LOJFOR"  ,GetSx3Cache("CV4_FILIAL","X3_TIPO"),GetSx3Cache("CV4_FILIAL","X3_TAMANHO"),GetSx3Cache("CV4_FILIAL","X3_DECIMAL"),GetSx3Cache("CV4_FILIAL","X3_TITULO"),GetSx3Cache("CV4_FILIAL","X3_PICTURE")})
    aAdd(aFieldsCab, {"NOMFOR"  ,GetSx3Cache("CV4_FILIAL","X3_TIPO"),GetSx3Cache("CV4_FILIAL","X3_TAMANHO"),GetSx3Cache("CV4_FILIAL","X3_DECIMAL"),GetSx3Cache("CV4_FILIAL","X3_TITULO"),GetSx3Cache("CV4_FILIAL","X3_PICTURE")})
    aAdd(aFieldsCab, {"PREFIXO" ,GetSx3Cache("CV4_FILIAL","X3_TIPO"),GetSx3Cache("CV4_FILIAL","X3_TAMANHO"),GetSx3Cache("CV4_FILIAL","X3_DECIMAL"),GetSx3Cache("CV4_FILIAL","X3_TITULO"),GetSx3Cache("CV4_FILIAL","X3_PICTURE")})
    aAdd(aFieldsCab, {"NUMERO"  ,GetSx3Cache("CV4_FILIAL","X3_TIPO"),GetSx3Cache("CV4_FILIAL","X3_TAMANHO"),GetSx3Cache("CV4_FILIAL","X3_DECIMAL"),GetSx3Cache("CV4_FILIAL","X3_TITULO"),GetSx3Cache("CV4_FILIAL","X3_PICTURE")})
    aAdd(aFieldsCab, {"PARCELA" ,GetSx3Cache("CV4_FILIAL","X3_TIPO"),GetSx3Cache("CV4_FILIAL","X3_TAMANHO"),GetSx3Cache("CV4_FILIAL","X3_DECIMAL"),GetSx3Cache("CV4_FILIAL","X3_TITULO"),GetSx3Cache("CV4_FILIAL","X3_PICTURE")})

    oTabTMPCab:= FWTemporaryTable():New(cAliasCab)
    oTabTMPCab:SetFields(aFieldsCab)
    oTabTMPCab:Create()

	aAdd(aFieldsGrid, {"FILIAL" ,GetSx3Cache("CV4_FILIAL","X3_TIPO"),GetSx3Cache("CV4_FILIAL","X3_TAMANHO"),GetSx3Cache("CV4_FILIAL","X3_DECIMAL"),GetSx3Cache("CV4_FILIAL","X3_TITULO"),GetSx3Cache("CV4_FILIAL","X3_PICTURE")})
    aAdd(aFieldsGrid, {"SEQUEN" ,GetSx3Cache("CV4_SEQUEN","X3_TIPO"),GetSx3Cache("CV4_SEQUEN","X3_TAMANHO"),GetSx3Cache("CV4_SEQUEN","X3_DECIMAL"),GetSx3Cache("CV4_SEQUEN","X3_TITULO"),GetSx3Cache("CV4_SEQUEN","X3_PICTURE")})
    aAdd(aFieldsGrid, {"DTSEQ"  ,GetSx3Cache("CV4_DTSEQ" ,"X3_TIPO"),GetSx3Cache("CV4_DTSEQ" ,"X3_TAMANHO"),GetSx3Cache("CV4_DTSEQ" ,"X3_DECIMAL"),GetSx3Cache("CV4_DTSEQ" ,"X3_TITULO"),GetSx3Cache("CV4_DTSEQ" ,"X3_PICTURE")})
    aAdd(aFieldsGrid, {"DEBITO" ,GetSx3Cache("CV4_DEBITO","X3_TIPO"),GetSx3Cache("CV4_DEBITO","X3_TAMANHO"),GetSx3Cache("CV4_DEBITO","X3_DECIMAL"),GetSx3Cache("CV4_DEBITO","X3_TITULO"),GetSx3Cache("CV4_DEBITO","X3_PICTURE")})
    aAdd(aFieldsGrid, {"CREDIT" ,GetSx3Cache("CV4_CREDIT","X3_TIPO"),GetSx3Cache("CV4_CREDIT","X3_TAMANHO"),GetSx3Cache("CV4_CREDIT","X3_DECIMAL"),GetSx3Cache("CV4_CREDIT","X3_TITULO"),GetSx3Cache("CV4_CREDIT","X3_PICTURE")})
    aAdd(aFieldsGrid, {"PERCEN" ,GetSx3Cache("CV4_PERCEN","X3_TIPO"),GetSx3Cache("CV4_PERCEN","X3_TAMANHO"),GetSx3Cache("CV4_PERCEN","X3_DECIMAL"),GetSx3Cache("CV4_PERCEN","X3_TITULO"),GetSx3Cache("CV4_PERCEN","X3_PICTURE")})
    aAdd(aFieldsGrid, {"VALOR"  ,GetSx3Cache("CV4_VALOR" ,"X3_TIPO"),GetSx3Cache("CV4_VALOR" ,"X3_TAMANHO"),GetSx3Cache("CV4_VALOR" ,"X3_DECIMAL"),GetSx3Cache("CV4_VALOR" ,"X3_TITULO"),GetSx3Cache("CV4_VALOR" ,"X3_PICTURE")})
    aAdd(aFieldsGrid, {"HIST"   ,GetSx3Cache("CV4_HIST"  ,"X3_TIPO"),GetSx3Cache("CV4_HIST"  ,"X3_TAMANHO"),GetSx3Cache("CV4_HIST"  ,"X3_DECIMAL"),GetSx3Cache("CV4_HIST"  ,"X3_TITULO"),GetSx3Cache("CV4_HIST"  ,"X3_PICTURE")})
    aAdd(aFieldsGrid, {"CCC"    ,GetSx3Cache("CV4_CCC"   ,"X3_TIPO"),GetSx3Cache("CV4_CCC"   ,"X3_TAMANHO"),GetSx3Cache("CV4_CCC"   ,"X3_DECIMAL"),GetSx3Cache("CV4_CCC"   ,"X3_TITULO"),GetSx3Cache("CV4_CCC"   ,"X3_PICTURE")})
    aAdd(aFieldsGrid, {"CCD"    ,GetSx3Cache("CV4_CCD"   ,"X3_TIPO"),GetSx3Cache("CV4_CCD"   ,"X3_TAMANHO"),GetSx3Cache("CV4_CCD"   ,"X3_DECIMAL"),GetSx3Cache("CV4_CCD"   ,"X3_TITULO"),GetSx3Cache("CV4_CCD"   ,"X3_PICTURE")})
    aAdd(aFieldsGrid, {"ITEMD"  ,GetSx3Cache("CV4_ITEMD" ,"X3_TIPO"),GetSx3Cache("CV4_ITEMD" ,"X3_TAMANHO"),GetSx3Cache("CV4_ITEMD" ,"X3_DECIMAL"),GetSx3Cache("CV4_ITEMD" ,"X3_TITULO"),GetSx3Cache("CV4_ITEMD" ,"X3_PICTURE")})
    aAdd(aFieldsGrid, {"ITEMC"  ,GetSx3Cache("CV4_ITEMC" ,"X3_TIPO"),GetSx3Cache("CV4_ITEMC" ,"X3_TAMANHO"),GetSx3Cache("CV4_ITEMC" ,"X3_DECIMAL"),GetSx3Cache("CV4_ITEMC" ,"X3_TITULO"),GetSx3Cache("CV4_ITEMC" ,"X3_PICTURE")})
    aAdd(aFieldsGrid, {"CLVLDB" ,GetSx3Cache("CV4_CLVLDB","X3_TIPO"),GetSx3Cache("CV4_CLVLDB","X3_TAMANHO"),GetSx3Cache("CV4_CLVLDB","X3_DECIMAL"),GetSx3Cache("CV4_CLVLDB","X3_TITULO"),GetSx3Cache("CV4_CLVLDB","X3_PICTURE")})
    aAdd(aFieldsGrid, {"CLVLDC" ,GetSx3Cache("CV4_CLVLDC","X3_TIPO"),GetSx3Cache("CV4_CLVLDC","X3_TAMANHO"),GetSx3Cache("CV4_CLVLDC","X3_DECIMAL"),GetSx3Cache("CV4_CLVLDC","X3_TITULO"),GetSx3Cache("CV4_CLVLDC","X3_PICTURE")})
    aAdd(aFieldsGrid, {"ITSEQ"  ,GetSx3Cache("CV4_ITSEQ" ,"X3_TIPO"),GetSx3Cache("CV4_ITSEQ" ,"X3_TAMANHO"),GetSx3Cache("CV4_ITSEQ" ,"X3_DECIMAL"),GetSx3Cache("CV4_ITSEQ" ,"X3_TITULO"),GetSx3Cache("CV4_ITSEQ" ,"X3_PICTURE")})
    
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
                    xValor	:= oExcel:GetValue(nContL,nContC)	//Conteúdo 
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
        APMsgAlert("Nenhuma movimentação encontrada com os parametros informados.","Atenção")
    EndIF 

    oTabTMPCab:Delete()
    oTabTMPGrid:Delete()
    RestArea(aArea)

Return

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ModelDef()
    Local oStructCab := FWFormModelStruct():New()
    Local oStructGrd := FWFormModelStruct():New()
    Local oModel
    Local nId

    oStructCab:AddTable(cAliasCab,{},"Tabela Cabeçalho")
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
 | Desc:  Criação da visão MVC                                         |
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
                        aFieldsCab[nId][5],;       // 03 = Título do campo
                        aFieldsCab[nId][5],;       // 04 = Descrição do campo
                        Nil,;                      // 05 = Array com Help
                        aFieldsCab[nId][2],;       // 06 = Tipo do campo
                        aFieldsCab[nId][6],;       // 07 = Picture
                        Nil,;                      // 08 = Bloco de PictTre Var
                        Nil,;                      // 09 = Consulta F3
                        .F.,;                      // 10 = Indica se o campo é alterável
                        Nil,;                      // 11 = Pasta do Campo
                        Nil,;                      // 12 = Agrupamnento do campo
                        Nil,;                      // 13 = Lista de valores permitido do campo (Combo)
                        Nil,;                      // 14 = Tamanho máximo da opção do combo
                        Nil,;                      // 15 = Inicializador de Browse
                        .F.,;                      // 16 = Indica se o campo é virtual (.T. ou .F.)
                        Nil,;                      // 17 = Picture Variavel
                        Nil)                       // 18 = Indica pulo de linha após o campo (.T. ou .F.)
    Next nId

    For nId := 1 To Len(aFieldsGrid)
        
        If aFieldsGrid[nId][1] == "OK"
            lAlt := .T.
        EndIf 
        
        oStructGrd:AddField(aFieldsGrid[nId][1],;                   // 01 = Nome do Campo
                        StrZero(nId,2),;                            // 02 = Ordem
                        aFieldsGrid[nId][5],;                       // 03 = Título do campo
                        aFieldsGrid[nId][5],;                       // 04 = Descrição do campo
                        Nil,;                                       // 05 = Array com Help
                        aFieldsGrid[nId][2],;                       // 06 = Tipo do campo
                        aFieldsGrid[nId][6],;                       // 07 = Picture
                        Nil,;                                       // 08 = Bloco de PictTre Var
                        Nil,;                                       // 09 = Consulta F3
                        IIF(aFieldsGrid[nId][1] == "OK",.T.,.F.),;  // 10 = Indica se o campo é alterável
                        Nil,;                                       // 11 = Pasta do Campo
                        Nil,;                                       // 12 = Agrupamnento do campo
                        Nil,;                                       // 13 = Lista de valores permitido do campo (Combo)
                        Nil,;                                       // 14 = Tamanho máximo da opção do combo
                        Nil,;                                       // 15 = Inicializador de Browse
                        .F.,;                                       // 16 = Indica se o campo é virtual (.T. ou .F.)
                        Nil,;                                       // 17 = Picture Variavel
                        Nil)                                        // 18 = Indica pulo de linha após o campo (.T. ou .F.)
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
                         /*cToolTip  | Comentário do botão*/,;
                         /*nShortCut | Codigo da Tecla para criação de Tecla de Atalho*/,;
                         /*aOptions  | */,;
                         /*lShowBar */ .T.)
    
    oView:AddUserButton( 'Marcar Todos', 'NOTE',;
                        {|oModel| fMark(oView)},;
                         /*cToolTip  | Comentário do botão*/,;
                         /*nShortCut | Codigo da Tecla para criação de Tecla de Atalho*/,;
                         /*aOptions  | */,;
                         /*lShowBar */ .T.)
 
Return oView

/*---------------------------------------------------------------------*
 | Func:  ViewActv                                                     |
 | Desc:  Realiza o PUT nos campos para gravação na tabela SE1         |
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
