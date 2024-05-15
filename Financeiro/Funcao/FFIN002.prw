#INCLUDE "Protheus.ch"
#INCLUDE "FWMVCDEF.ch"
#INCLUDE "TopConn.ch"
#INCLUDE "Tbiconn.ch"
#INCLUDE "RPTDef.ch"
#INCLUDE "FWPrintSetup.ch"

/*/{Protheus.doc} FFIN002
Tela de LOG de importação do Arquivo de Retorno CNAB de Cobrança
@type function
@author Elvis Siqueira
@since 08/05/2024
/*/
User Function FFIN002(pArq)
    Local aArea  := FWGetArea()
    Local aButtons := {;
                        {.F.,Nil},;         // Copiar
                        {.F.,Nil},;         // Recortar
                        {.F.,Nil},;         // Colar
                        {.F.,Nil},;         // Calculadora
                        {.F.,Nil},;         // Spool
                        {.F.,Nil},;         // Imprimir
                        {.F.,Nil},;         // Confirmar
                        {.T.,"Fechar"},;    // Cancelar
                        {.F.,Nil},;         // WalkTrhough
                        {.F.,Nil},;         // Ambiente
                        {.F.,Nil},;         // Mashup
                        {.F.,Nil},;         // Help
                        {.F.,Nil},;         // Formulário HTML
                        {.F.,Nil};          // ECM
                    }
    
    Private cCamposCAB := "ZZ1_ARQUIV;ZZ1_BANCO;ZZ1_AGENCI;ZZ1_NUMCOM;"
    Private cArquivo := pArq

    DBSelectArea('ZZ1')
    ZZ1->(DBGoTop())
    ZZ1->(MSSeek(xFilial('ZZ1')+Pad(cArquivo,FWTamSX3("ZZ1_ARQUIV")[1])))

    FWExecView("LOG de Importação (Retorno Cobrança)","FFIN002",MODEL_OPERATION_VIEW,,{|| .T.},,,aButtons)
    
    FWRestArea(aArea)

Return 

/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ModelDef()
    Local oModel as object
    Local oStrMaster := FWFormStruct(1, 'ZZ1', {|cCampo| Alltrim(cCampo) $ cCamposCAB})
    Local oStrGrid := FWFormStruct(1, 'ZZ1')

    oModel := MPFormModel():New('FFIN002M',/*bPre*/,/*bPost*/,/*bCommit*/,/*bCancel*/)
    oModel:AddFields('ZZ1MASTER',/*cOwner*/,oStrMaster/*bPre*/,/*bPos*/,/*bLoad*/)
    oModel:AddGrid('ZZ1GRID','ZZ1MASTER',oStrGrid,/*bLinePre*/,/*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)
    oModel:SetRelation('ZZ1GRID',{{'ZZ1_FILIAL','xFilial("ZZ1")'},;
                                  {'ZZ1_ARQUIV','ZZ1_ARQUIV'    }},;
                                   ZZ1->(IndexKey(1)))

	  oModel:SetPrimaryKey({})

    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_NUMTIT' , 'ZZ1__NUMTIT', 'COUNT', { || .T. },,'Qtd. Titulos' )
    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_VLDESP' , 'ZZ1__VLDESP', 'SUM'  , { || .T. },,'Despesas' )
    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_VLDESC' , 'ZZ1__VLDESC', 'SUM'  , { || .T. },,'Descontos' )
    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_VLABAT' , 'ZZ1__VLABAT', 'SUM'  , { || .T. },,'Abatimentos' )
    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_VLREC'  , 'ZZ1__VLREC' , 'SUM'  , { || .T. },,'Recebidos' )
    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_JUROS'  , 'ZZ1__JUROS' , 'SUM'  , { || .T. },,'Juros' )
    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_MULTA'  , 'ZZ1__MULTA' , 'SUM'  , { || .T. },,'Multas' )
    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_OUTDES' , 'ZZ1__OUTDES', 'SUM'  , { || .T. },,'Outras Desp.' )
    oModel:AddCalc( 'FFIN002CALC',  'ZZ1MASTER', 'ZZ1GRID', 'ZZ1_VLCRED' , 'ZZ1__VLCRED', 'SUM'  , { || .T. },,'Créditos' )

    //oModel:SetDescription("")

Return oModel

/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
    Local oView as object
    Local oModel as object
    Local oStrCalc as object
    Local oStrMaster := FWFormStruct(2, 'ZZ1', {|cCampo| Alltrim(cCampo) $ cCamposCAB})
    Local oStrGrid := FWFormStruct(2, 'ZZ1')

    oModel := FWLoadModel("FFIN002")
    oStrCalc := FWCalcStruct( oModel:GetModel('FFIN002CALC') )

    oView := FwFormView():New()
    oView:SetModel(oModel)
    oView:SetProgressBar(.T.)
    oView:AddField("VIEW_MASTER", oStrMaster , "ZZ1MASTER")
    oView:AddGrid("VIEW_GRID"   , oStrGrid   , "ZZ1GRID")
    oView:AddField('VIEW_CALC'  , oStrCalc   ,'FFIN002CALC')

    oView:CreateHorizontalBox("CABEC", 10 )
    oView:CreateHorizontalBox("GRID" , 70 )
    oView:CreateHorizontalBox("CALC" , 20 )

    oView:SetOwnerView('VIEW_MASTER','CABEC')
    oView:SetOwnerView('VIEW_GRID'  ,'GRID')
    oView:SetOwnerView('VIEW_CALC'  ,'CALC')

    //oView:SetAfterViewActivate({|oView| ViewActv(oView)})

    oView:AddUserButton( 'Imprimir', 'MAGIC_BMP',;
                            {|| U_FFINR002() },;
                            /*cToolTip  | Comentário do botão*/,;
                            /*nShortCut | Codigo da Tecla para criação de Tecla de Atalho*/,;
                            /*aOptions  | */,;
                            /*lShowBar */ .T.)

Return oView

/*-----------------------------------------------------------------------*
 | Func:  ViewActv                                                       |
 | Desc:  Chamada na inicialização da View para preenchimento dos campos |
 | Obs.:  /                                                              |
 *----------------------------------------------------------------------*/
Static Function ViewActv(oView)
    Local oModel := FWModelActive() 
    Local oModelMaster := oModel:GetModel("ZZ1MASTER")
    Local oModelGrid := oModel:GetModel("ZZ1GRID")
    Local nLin := 1

    DBSelectArea('ZZ1')
    ZZ1->(DBGoTop())
    ZZ1->(MSSeek(xFilial('ZZ1')+Pad(cArquivo,FWTamSX3("ZZ1_ARQUIV")[1])))

    oModelMaster:LoadValue("ZZ1_ARQUIV", ZZ1->ZZ1_ARQUIV)
    oModelMaster:LoadValue("ZZ1_BANCO" , ZZ1->ZZ1_BANCO )
    oModelMaster:LoadValue("ZZ1_AGENCI", ZZ1->ZZ1_AGENCI)
    oModelMaster:LoadValue("ZZ1_NUMCOM", ZZ1->ZZ1_NUMCOM)

    oView:Refresh("VIEW_MASTER")

    While ZZ1->(!Eof()) .AND. Alltrim(ZZ1->ZZ1_ARQUIV) == cArquivo

        nLin++

        oModelGrid:AddLine()

        oModelGrid:LoadValue("ZZ1_NUMTIT", ZZ1->ZZ1_NUMTIT)
        oModelGrid:LoadValue("ZZ1_DBAIXA", ZZ1->ZZ1_DBAIXA)
        oModelGrid:LoadValue("ZZ1_TIPO"  , ZZ1->ZZ1_TIPO  )
        oModelGrid:LoadValue("ZZ1_NSNUM" , ZZ1->ZZ1_NSNUM )
        oModelGrid:LoadValue("ZZ1_VLDESP", ZZ1->ZZ1_VLDESP)
        oModelGrid:LoadValue("ZZ1_VLDESC", ZZ1->ZZ1_VLDESC)
        oModelGrid:LoadValue("ZZ1_VLABAT", ZZ1->ZZ1_VLABAT)
        oModelGrid:LoadValue("ZZ1_VLREC" , ZZ1->ZZ1_VLREC )
        oModelGrid:LoadValue("ZZ1_JUROS" , ZZ1->ZZ1_JUROS )
        oModelGrid:LoadValue("ZZ1_MULTA" , ZZ1->ZZ1_MULTA )
        oModelGrid:LoadValue("ZZ1_OUTDES", ZZ1->ZZ1_OUTDES)
        oModelGrid:LoadValue("ZZ1_VLCRED", ZZ1->ZZ1_VLCRED)
        oModelGrid:LoadValue("ZZ1_DCRED" , ZZ1->ZZ1_DCRED )
        oModelGrid:LoadValue("ZZ1_OCORR" , ZZ1->ZZ1_OCORR )
        oModelGrid:LoadValue("ZZ1_MOTBAN", ZZ1->ZZ1_MOTBAN)
        oModelGrid:LoadValue("ZZ1_LINHA" , AllToChar(nLin))
        oModelGrid:LoadValue("ZZ1_LINARQ", ZZ1->ZZ1_LINARQ)
        oModelGrid:LoadValue("ZZ1_DVENC" , ZZ1->ZZ1_DVENC )
        oModelGrid:LoadValue("ZZ1_BANCO" , ZZ1->ZZ1_BANCO )
        oModelGrid:LoadValue("ZZ1_AGENCI", ZZ1->ZZ1_AGENCI)
        oModelGrid:LoadValue("ZZ1_NUMCOM", ZZ1->ZZ1_NUMCOM)
        oModelGrid:LoadValue("ZZ1_ARQUIV", ZZ1->ZZ1_ARQUIV)

        ZZ1->(DBSkip())
    EndDo 

    oModelGrid:GoLine(1)
    oView:Refresh("VIEW_GRID")
    oView:SetNoDeleteLine('VIEW_GRID')
    oView:SetNoInsertLine('VIEW_GRID')

Return

/*/{Protheus.doc} FFINR002
Relatório de importação do Arquivo de Retorno CNAB de Cobrança
@type function
@author Elvis Siqueira
@since 06/05/2024
/*/
User Function FFINR002()

    FwMsgRun(Nil, {|| fMontaRel() }, "Processando", "Gerando o relatório.")

Return

/*---------------------------------------------------------------------*
 | Func:  fMontaRel                                                    |
 | Desc:  Funcao principal que monta o relatório                       |
 *---------------------------------------------------------------------*/

Static Function fMontaRel()
    Local oModel     := FWModelActive() 
    Local oModelGrid := oModel:GetModel("ZZ1GRID")
    Local cNomeRel   := FunName()+"_"+FWTimeStamp(1,dDataBase,Time())
    Local cPictVal   := GetSx3Cache("E1_VALOR","X3_PICTURE")
    Local nTotTit    := 0
    Local nTotDesp   := 0
    Local nTotDesc   := 0
    Local nTotAbat   := 0
    Local nTotRece   := 0
    Local nTotJuro   := 0
    Local nTotMult   := 0
    Local nTotODes   := 0
    Local nTotCred   := 0
    Local nY

    Private oPrintRel   := Nil
    Private oFontCab    := TFont():New("Arial", , -10, , .F.)
	Private oFontCabN   := TFont():New("Arial", , -10, , .T.)
    Private oFontCab2   := TFont():New("Arial", , -16, , .T.)
    Private oFontDet    := TFont():New("Arial", , -10, , .F.)
    Private	oFontDetN   := TFont():New("Arial", , -10, , .T.)
    Private oFontRod    := TFont():New("Arial", , -06, , .F.)
    Private nLinAtu     := 0
	Private nLinFin     := 580
	Private nColIni     := 010
	Private nColFin     := 830
    Private nPagAtu     := 1
    Private nPadLeft    := 0 //Alinhamento a Esquerda
    Private nPadRight   := 1 //Alinhamento a Direita
    Private nPadCenter  := 2 //Alinhamento Centralizado

    //Posição das Colunas
    Private nPosNumT := 0010
    Private nPosNNum := 0090
    Private nPosDTBX := 0180
    Private nPosDesp := 0230
    Private nPosDesc := 0300
    Private nPosAbat := 0370
    Private nPosRece := 0440
    Private nPosJuro := 0510
    Private nPosMult := 0580
    Private nPosODes := 0650
    Private nPosCred := 0720
    Private nPosDCre := 0790

    oPrintRel := FWMSPrinter():New(cNomeRel, IMP_PDF, .F., /*cStartPath*/, .T., , @oPrintRel, , , , , .T.)
	oPrintRel:cPathPDF := GetTempPath()
	oPrintRel:SetResolution(72)
	oPrintRel:SetLandscape()
	oPrintRel:SetPaperSize(DMPAPER_A4)
	oPrintRel:SetMargin(10, 10, 10, 10)

    //Imprime o cabecalho
	fImpCab()

    For nY := 1 To oModelGrid:Length()
        
        oModelGrid:GoLine(nY)
	
        oPrintRel:SayAlign(nLinAtu, nPosNumT, oModelGrid:GetValue("ZZ1_NUMTIT")                                     , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosNNum, oModelGrid:GetValue("ZZ1_NSNUM")                                      , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosDTBX, DToC(oModelGrid:GetValue("ZZ1_DBAIXA"))                               , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosDesp, AllTrim(AllToChar(oModelGrid:GetValue("ZZ1_VLDESP"),cPictVal, .F.))   , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosDesc, AllTrim(AllToChar(oModelGrid:GetValue("ZZ1_VLDESC"),cPictVal, .F.))   , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosAbat, AllTrim(AllToChar(oModelGrid:GetValue("ZZ1_VLABAT"),cPictVal, .F.))   , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosRece, AllTrim(AllToChar(oModelGrid:GetValue("ZZ1_VLREC" ),cPictVal, .F.))   , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosJuro, AllTrim(AllToChar(oModelGrid:GetValue("ZZ1_JUROS" ),cPictVal, .F.))   , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosMult, AllTrim(AllToChar(oModelGrid:GetValue("ZZ1_MULTA" ),cPictVal, .F.))   , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosODes, AllTrim(AllToChar(oModelGrid:GetValue("ZZ1_OUTDES"),cPictVal, .F.))   , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosCred, AllTrim(AllToChar(oModelGrid:GetValue("ZZ1_VLCRED"),cPictVal, .F.))   , oFontDet, 100, 07, , nPadLeft,)
        oPrintRel:SayAlign(nLinAtu, nPosDCre, DToC(oModelGrid:GetValue("ZZ1_DCRED "))                               , oFontDet, 100, 07, , nPadLeft,)

        nTotTit++
        nTotDesp += oModelGrid:GetValue("ZZ1_VLDESP")
        nTotDesc += oModelGrid:GetValue("ZZ1_VLDESC")
        nTotAbat += oModelGrid:GetValue("ZZ1_VLABAT")
        nTotRece += oModelGrid:GetValue("ZZ1_VLREC" )
        nTotJuro += oModelGrid:GetValue("ZZ1_JUROS" )
        nTotMult += oModelGrid:GetValue("ZZ1_MULTA" )
        nTotODes += oModelGrid:GetValue("ZZ1_OUTDES")
        nTotCred += oModelGrid:GetValue("ZZ1_VLCRED")

        nLinAtu += 10

        If nLinAtu + 10 >= nLinFin
            fImpRod()
            fImpCab()
	    ElseIf nY == oModelGrid:Length()
            nLinAtu += 10

            oPrintRel:SayAlign(nLinAtu, nPosDTBX, "Qtd. Titulos: " + cValToChar(nTotTit)      , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosDTBX, "Totais: "                                  , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosDesp, AllTrim(AllToChar(nTotDesp,cPictVal, .F.))  , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosDesc, AllTrim(AllToChar(nTotDesc,cPictVal, .F.))  , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosAbat, AllTrim(AllToChar(nTotAbat,cPictVal, .F.))  , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosRece, AllTrim(AllToChar(nTotRece,cPictVal, .F.))  , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosJuro, AllTrim(AllToChar(nTotJuro,cPictVal, .F.))  , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosMult, AllTrim(AllToChar(nTotMult,cPictVal, .F.))  , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosODes, AllTrim(AllToChar(nTotODes,cPictVal, .F.))  , oFontDetN, 100, 07, , nPadLeft,)
            oPrintRel:SayAlign(nLinAtu, nPosCred, AllTrim(AllToChar(nTotCred,cPictVal, .F.))  , oFontDetN, 100, 07, , nPadLeft,)

            oModelGrid:GoLine(1)
        EndIf

    Next nY

    //Imprime o rodape
	fImpRod()

    //Gera o pdf para visualizacao
	oPrintRel:Preview()

Return 

/*---------------------------------------------------------------------*
 | Func:  fImpCab                                                      |
 | Desc:  Funcao que imprime o cabecalho                               |
 *---------------------------------------------------------------------*/

Static Function fImpCab()
	Local nLinCab    := 025
    Local cLogoEmp   := fLogoEmp()
	Local cEmpresa   := Iif(Empty(SM0->M0_NOMECOM), Alltrim(SM0->M0_NOME), Alltrim(SM0->M0_NOMECOM))
	Local cEmpTel    := Alltrim(Transform(SubStr(SM0->M0_TEL, 1, Len(SM0->M0_TEL)), "@R (99)9999-9999"))
	Local cEmpCidade := AllTrim(SM0->M0_CIDENT)+" / "+SM0->M0_ESTENT
	Local cEmpCnpj   := Alltrim(Transform(SM0->M0_CGC, "@R 99.999.999/9999-99"))
	Local cEmpCep    := Alltrim(Transform(SM0->M0_CEPENT, "@R 99999-999"))
    Local cMsgRel    := ""

    If Alltrim(ProcName(25)) == "FINA150"
        cMsgRel := "Relatório (Log de Geração CNAB Cobrança)"
    ElseIF Alltrim(ProcName(26)) == "FINA200"
        cMsgRel := "Relatório (Log de Retorno CNAB Cobrança)"
    EndIf 
	
	oPrintRel:StartPage()
	
	//Dados da Empresa
	oPrintRel:Box(nLinCab, nColIni, nLinCab + 65, nColFin )
	oPrintRel:SayBitmap(nLinCab+3,  nColIni+5, cLogoEmp, 054, 054)
	oPrintRel:SayAlign(nLinCab+3,     nColIni+65, "Empresa: " + cEmpresa,  oFontCabN, 500, 07, , nPadLeft, )
	nLinCab += 13
	oPrintRel:SayAlign(nLinCab,     nColIni+65, "CNPJ: " + cEmpCnpj,      oFontCabN, 500, 07, , nPadLeft, )
	nLinCab += 13
	oPrintRel:SayAlign(nLinCab,     nColIni+65, "Cidade: " + cEmpCidade,  oFontCabN, 500, 07, , nPadLeft, )
    oPrintRel:SayAlign(nLinCab,     nColFin-300, cMsgRel,                 oFontCab2, 500, 07, , nPadLeft, )
	nLinCab += 13
	oPrintRel:SayAlign(nLinCab,     nColIni+65, "CEP: " + cEmpCep,        oFontCabN, 500, 07, , nPadLeft, )
	nLinCab += 13
	oPrintRel:SayAlign(nLinCab,     nColIni+65, "Telefone: " + cEmpTel,   oFontCabN, 500, 07, , nPadLeft, )

	//Cabecalho com descricao das colunas
	nLinCab += 20
	
	oPrintRel:SayAlign(nLinCab, nPosNumT, "Num Titulo"  ,   oFontDetN, 100, 07, , nPadLeft,)	
	oPrintRel:SayAlign(nLinCab, nPosNNum, "Nosso Numero",   oFontDetN, 100, 07, , nPadLeft,)
    oPrintRel:SayAlign(nLinCab, nPosDTBX, "Dt Baixa"    ,   oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosDesp, "Vlr. Despesa",   oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosDesc, "Vlr. Descont",	oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosAbat, "Vlr. Abatime", 	oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosRece, "Vlr. Recebid", 	oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosJuro, "Vlr. Juros  ", 	oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosMult, "Vlr. Multa  ",   oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosODes, "Out. Despesa",   oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosCred, "Vlr. Credito",   oFontDetN, 100, 07, , nPadLeft,)
	oPrintRel:SayAlign(nLinCab, nPosDCre, "Dt. Credito ",   oFontDetN, 100, 07, , nPadLeft,)
	
	//Atualizando a linha inicial do relatório
	nLinAtu := nLinCab + 20
Return

/*---------------------------------------------------------------------*
 | Func:  fImpRod                                                      |
 | Desc:  Funcao que imprime o rodape                                  |
 *---------------------------------------------------------------------*/

Static Function fImpRod()
	Local oModel     := FWModelActive() 
    Local oModelMaster := oModel:GetModel("ZZ1MASTER")
    Local nLinRod:= nLinFin + 10
	Local cTexto := ""

	//Linha Separatória
	oPrintRel:Line(nLinRod, nColIni, nLinRod, nColFin)
	nLinRod += 3
	
	//Dados da Esquerda
	cTexto := "Arquivo: "+ Alltrim(oModelMaster:GetValue("ZZ1_ARQUIV")) +"    |    "+dToC(dDataBase)+"     "+Time()+"     "+FunName()+"     "+cUserName
	oPrintRel:SayAlign(nLinRod, nColIni,    cTexto, oFontRod, 250, 07, , nPadLeft, )
	
	//Direita
	cTexto := "Página "+cValToChar(nPagAtu)
	oPrintRel:SayAlign(nLinRod, nColFin-40, cTexto, oFontRod, 040, 07, , nPadRight, )
	
	//Finalizando a Página e somando mais um
	oPrintRel:EndPage()
	nPagAtu++
Return

/*---------------------------------------------------------------------*
 | Func:  fLogoEmp                                                     |
 | Desc:  Funcao que retorna o logo da empresa (igual a DANFE)         |
 *---------------------------------------------------------------------*/

Static Function fLogoEmp()
	Local cGrpCompany := AllTrim(FWGrpCompany())
	Local cCodEmpGrp  := AllTrim(FWCodEmp())
	Local cUnitGrp    := AllTrim(FWUnitBusiness())
	Local cFilGrp     := AllTrim(FWFilial())
	Local cLogo       := ""
	Local cCamFim     := GetTempPath()
	Local cStart      := GetSrvProfString("Startpath", "")

	//Se tiver filiais por grupo de empresas
	If !Empty(cUnitGrp)
		cDescLogo	:= cGrpCompany + cCodEmpGrp + cUnitGrp + cFilGrp
		
	//Se nao, será apenas, empresa + filial
	Else
		cDescLogo	:= cEmpAnt + cFilAnt
	EndIf
	
	//Pega a imagem
	cLogo := cStart + "LGMID" + cDescLogo + ".PNG"
	
	//Se o arquivo Nao existir, pega apenas o da empresa, desconsiderando a filial
	If !File(cLogo)
		cLogo	:= cStart + "LGMID" + cEmpAnt + ".PNG"
	EndIf
	
	//Copia para a temporaria do s.o.
	CpyS2T(cLogo, cCamFim)
	cLogo := cCamFim + StrTran(cLogo, cStart, "")
	
	//Se o arquivo Nao existir na temporaria, espera meio segundo para terminar a cópia
	If !File(cLogo)
		Sleep(500)
	EndIf
Return cLogo
