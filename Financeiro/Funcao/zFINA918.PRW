#INCLUDE "PROTHEUS.CH"
#INCLUDE "FWMVCDEF.CH"
//#INCLUDE "FINA918.CH"

Static __cFilTef := ""

/*/{Protheus.doc} FINA918
Conciliacao de Pagamentos

@author Totvs
@since 07/12/2020
/*/
User Function zFINA918()

	Local oBrowse 		As Object

	Private cCadastro	As Character
	Private cString 	As Character

	cCadastro := "Conciliacao Pagamentos" //STR0020	
	cString   := "FIF"
	__cFilTef := cFilAnt //salva a filial logada para gerar o mov. de lote no fina110.
	
	dbSelectArea(cString)
	dbSetOrder(1)

	SetKey(019, {||}) //Desabilita CTRL+S para prevenir error.log

	oBrowse := BrowseDef()
	oBrowse:Activate()

	__cFilTef := ""
Return

/*/{Protheus.doc} BrowseDef
Definicoes do Browse

@author Totvs
@since 07/12/2020

/*/
Static Function BrowseDef() As Object
    Local oBrowse As Object
	Local cFiltro As Character

    oBrowse := FwMBrowse():New()

    oBrowse:SetAlias("FIF")

    oBrowse:SetDescription("Conciliacao Pagamentos")		//"Conciliacao Pagamentos"
	// Legendas
    oBrowse:AddLegend("FIF_STATUS $ '1'", "BR_VERDE"	, 	"Não Processado"				) //#"Não Processado"
    oBrowse:AddLegend("FIF_STATUS $ '2'", "BR_VERMELHO"	, 	"Conciliado Normal"				) //#"Conciliado Normal"
    oBrowse:AddLegend("FIF_STATUS $ '3'", "BR_AMARELO"	,	"Divergente"					) //#"Divergente"
    oBrowse:AddLegend("FIF_STATUS $ '4'", "BR_AZUL"		, 	"Conciliado Manual"				) //#"Conciliado Manual"
    oBrowse:AddLegend("FIF_STATUS $ '5'", "BR_BRANCO"	, 	"Descartado"					) //#"Descartado"
    oBrowse:AddLegend("FIF_STATUS $ '6'", "BR_LARANJA"	, 	"Ant. Nao Processada"			) //#"Ant. Nao Processada"
    oBrowse:AddLegend("FIF_STATUS $ '7'", "BR_PINK"		, 	"Antecipado"	                ) //#"Antecipado"
	oBrowse:AddLegend("FIF_STATUS $ '8'", "BR_AZUL_CLARO", 	"Antecipação Conciliada Manual"	) //#"Antecipação Conciliada Manual"

	cFiltro := " FIF_STATUS  <> ' ' "
	
	oBrowse:SetFilterDefault(cFiltro)

Return oBrowse

/*/{Protheus.doc} A918VISUAL
Rotina para visualização do registro posicionado
@type Function
@author Rogério Melônio
@since 18/01/2018
@version 12
/*/
Static Function A918VISUAL()
Local nRecno := FIF->(Recno())

A918ALTERA("FIF", nRecno, 2)

Return Nil

/*/{Protheus.doc} A918EDITA
Rotina para ediçao do registro posicionado

@author Totvs
@since 07/12/2020

/*/
Static Function A918Edita()
Local nRecno := FIF->(Recno())

A918ALTERA("FIF", nRecno, 4)

Return Nil

/*/{Protheus.doc} A918ALTERA
Rotina para visualização ou alteração do registro posicionado
@type Function
@author Unknown
@since 06/08/2009
@version 12
/*/
Static Function A918ALTERA(cAlias,nReg,nOpc)                                                          		
Local aCpoEnch		:= {}
Local aAlterEnch	:= {}
Local nModelo		:= 3	// Se for diferente de 1 desabilita execucao de gatilhos estrangeiros                           
Local lF3			:= .F.	// Indica se a enchoice esta sendo criada em uma consulta F3 para utilizar variaveis de memoria 
Local lMemoria		:= .T.	// Indica se a enchoice utilizara variaveis de memoria ou os campos da tabela na edicao         
Local lColumn		:= .F.	// Indica se a apresentacao dos campos sera em forma de coluna                                  
Local caTela		:= ""	// Nome da variavel tipo "private" que a enchoice utilizara no lugar da propriedade aTela       
Local lNoFolder		:= .F.	// Indica se a enchoice nao ira utilizar as Pastas de Cadastro (SXA)                            
Local lProperty		:= .T.	// Indica se a enchoice nao utilizara as variaveis aTela e aGets, somente suas propriedades com os mesmos nomes
Local aButtons		:= {}
Local lRet			:= .F.
Local nI			:= 0
Local oDlg			// Dialog Principal
Local aSize			:= MsAdvSize()
Local aPos  		:= {}
Local oSize 		:= FWDefSize():New(.T.)

//Monto os campos da Enchoice deixando somente os campos customizados e o FIF_STATUS para alterar
dbSelectArea("SX3")
dbSetOrder(1)
If SX3->( dbSeek( cAlias ) )
	While !SX3->( Eof() ) .And. SX3->X3_ARQUIVO == cAlias
		aAdd(aCpoEnch,SX3->X3_CAMPO)
		If (Alltrim(SX3->X3_PROPRI) == "U" .Or. Alltrim(SX3->X3_CAMPO) == "FIF_STATUS") .And. nOpc <> 2
			aAdd(aAlterEnch,SX3->X3_CAMPO)
		EndIf
		SX3->( dbSkip() )
	End
EndIf

oSize:AddObject( "ENCHOICE", 100,100, .T., .T. )
oSize:Process()
aPos := {oSize:GetDimension("ENCHOICE","LININI"),oSize:GetDimension("ENCHOICE","COLINI"),oSize:GetDimension("ENCHOICE","LINEND"),oSize:GetDimension("ENCHOICE","COLEND")}

DEFINE MSDIALOG oDlg TITLE "Conciliador TEF" + " - " + If(ALTERA,"Alterar","Visualizar") FROM aSize[7],0 to aSize[6],aSize[5] Of oMainWnd PIXEL // ##"Conciliador TEF" ##"Alterar" ##"Visualizar"
	RegToMemory(cAlias, INCLUI, .F.)
	Enchoice(cAlias,,nOpc,,,,aCpoEnch,aPos,aAlterEnch,nModelo,,,,oDlg,lF3,lMemoria,lColumn,caTela,lNoFolder,lProperty)

ACTIVATE MSDIALOG oDlg CENTERED  ON INIT EnchoiceBar(oDlg, {|| lRet := .T.,oDlg:End()},{||lRet := .F.,oDlg:End()},,aButtons)

If lRet .And. nOpc <> 2 .And. Len(aAlterEnch) > 0
	RecLock(cAlias,.F.)
	For nI := 1 to Len(aAlterEnch)
		(cAlias)->&(aAlterEnch[nI]) :=  M->&(aAlterEnch[nI])
	Next nI
	(cAlias)->( MsUnLock() )
EndIf

Return

/*/{Protheus.doc}F918BxLote
Retornar a filial em que será gerada o movimento totalizador de baixa por lote

@author Sivaldo Oliveira
@since  18/02/2017
@version 12
/*/
Static Function F918BxLote()

Return __cFilTef

/*/{Protheus.doc} MenuDef
Definicoes de Menu

@author Totvs
@since 07/12/2020

/*/
Static Function MenuDef() As Array
	Local aRotina As Array

	aRotina := {}

	ADD OPTION aRotina TITLE "Pesquisar" 	ACTION 'AxPesqui'  	OPERATION 1 	ACCESS 0 //Pesquisar
	ADD OPTION aRotina TITLE "Conciliação"	ACTION "F920View"	OPERATION 2   	ACCESS 0 //"Conciliação"
	ADD OPTION aRotina TITLE "Alterar"		ACTION "A918EDITA"	OPERATION 4   	ACCESS 0 //"Alterar"		
	ADD OPTION aRotina TITLE "Visualizar"	ACTION "A918Visual"	OPERATION 2   	ACCESS 0 //"Visualizar"	
	ADD OPTION aRotina TITLE "Listagem Reg"	ACTION "FINA910C"	OPERATION 1   	ACCESS 0 //"Listagem Reg"
	ADD OPTION aRotina TITLE 'Importar'		ACTION "U_FFIN001"	OPERATION 3   	ACCESS 0 //"Importar"

Return aRotina

/*/{Protheus.doc} F918AtuTPD
Avalia se o ambiente possui os pré-requisitos para utilizar
a conciliação de pagamentos via carteiras digitais (TPD - Totvs Pagamentos Digitais)

@author Totvs
@since 06/10/2021

/*/
Static Function F918AtuTPD() As Logical

	Local lRetorno 	as Logical
	Local oPerg As Object
	Local aPergunte As Array

	lRetorno := .F.
	
	If FIF->(ColumnPos("FIF_MODPAG")) > 0 .And. FJU->(ColumnPos("FJU_DOCTEF")) > 0 .And. TamSX3("FIF_NUCOMP")[1] >= 100	
		//Verifica se existe o novo grupo de perguntas
		oPerg := FwSx1Util():New()
		oPerg:AddGroup("FINA918")
		oPerg:SearchGroup()
		aPergunte := oPerg:GetGroup("FINA918")
		If LEN(aPergunte) > 1 .AND. !Empty(aPergunte[2])
			lRetorno := .T.
			FwFreeArray(aPergunte)
		Endif
		FWFReeObj(oPerg)
	Endif

Return lRetorno

/*/{Protheus.doc} FASOFTEXP
Função que verifica o código da Software Express 
cadastrado na tabela MDE - Códigos de Retorno TEF

@return	cCodAdm, Character, retorna o código da Software Express.
@author Alberto Teixeira
@since 20/10/2022
/*/
Static Function FASOFTEXP(oSofex) As Character

	Local cQuery	As Character
	Local cCodAdm	As Character

	cQuery := " SELECT MDE.MDE_CODIGO CODADM"
	cQuery += " FROM " + RetSqlName("MDE") + " MDE "
	cQuery += " WHERE "
	cQuery += " MDE_ARQIMP <> ' ' AND "
	cQuery += " MDE_ARQPRO <> ' ' AND "
	cQuery += " MDE_ARQINC <> ' ' AND "
	cQuery += " MDE_TIPO = 'RD' AND "
	cQuery += " MDE_DESC IN ('SOFTWAREEXPRESS','SOFTWARE EXPRESS') AND "
	cQuery += " MDE.D_E_L_E_T_ = ' ' "
	cQuery := ChangeQuery(cQuery)

	oSofex := FWPreparedStatement():New(cQuery)

    cCodAdm  := MpSysExecScalar(cQuery,"CODADM")

Return cCodAdm
