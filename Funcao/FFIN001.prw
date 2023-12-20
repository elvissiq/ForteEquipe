#INCLUDE "Protheus.ch"
#INCLUDE "FWMVCDEF.ch"
#INCLUDE "TopConn.ch"
#INCLUDE "Tbiconn.ch"

Static cTitulo := "Importação - Arquivo de Conciliação"
Static __oRegiFIF := Nil

/*/{Protheus.doc} FFIN001
Leitura do arquivo de Conciliação CIELO
@type function
@author Elvis Siqueira
@since 20/12/2023
/*/
User Function FFIN001()
    Private oProcess := Nil

    oProcess := MsNewProcess():New({|| FIN1Proc()}, "Processando arquivo...", "Aguarde...", .T.)
    oProcess:Activate()
    
Return

Static Function FIN1Proc()
    Local aArea := FWGetArea()
    Local oExcel
    Local aTamLin
    Local nContP,nContL
    Local cArq := ""
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
    Private aFieldsCab  := {}
    Private aFieldsGrid := {}
    Private cAliasCab := GetNextAlias()
    Private cAliasGrid := GetNextAlias()

    aAdd(aFieldsCab, {"DTIMPOR" ,"D",8  ,0,"Data Importação",""})
    aAdd(aFieldsCab, {"PERIODO" ,"C",40 ,0,"Período",""})
    aAdd(aFieldsCab, {"ARQUIVO" ,"C",200,0,"Arquivo",""})
    
    oTabTMPCab:= FWTemporaryTable():New(cAliasCab)
    oTabTMPCab:SetFields(aFieldsCab)
    oTabTMPCab:Create()

	aAdd(aFieldsGrid, {"STATUS"  ,"C",10,0,"Status",""})
    aAdd(aFieldsGrid, {"TPLANC"  ,"C",20,0,"Tipo de Lançamento",""})
    aAdd(aFieldsGrid, {"DESCRIC" ,"C",40,0,"Descrição",""})
    aAdd(aFieldsGrid, {"BANCO"   ,"C",20,0,"Banco",""})
    aAdd(aFieldsGrid, {"AGENCIA" ,"C",7 ,0,"Agência",""})
    aAdd(aFieldsGrid, {"CONTA"   ,"C",10,0,"Conta",""})
    aAdd(aFieldsGrid, {"GRAVAME" ,"C",3,0,"Gravame",""})
    aAdd(aFieldsGrid, {"CNPJORI" ,"C",20,0,"CNPJ da instituição origem da negociação",""})
    aAdd(aFieldsGrid, {"NOMEORI" ,"C",40,0,"Nome da instituição origem da negociação",""})
    aAdd(aFieldsGrid, {"ESTABE"  ,"C",15,0,"Estabelecimento",""})
    aAdd(aFieldsGrid, {"DTPAGAM" ,"D",8 ,0,"Data de pagamento",""})
    aAdd(aFieldsGrid, {"DTLANCA" ,"D",8 ,0,"Data do lançamento",""})
    aAdd(aFieldsGrid, {"DTAUTVE" ,"D",8 ,0,"Data da autorização da venda",""})
    aAdd(aFieldsGrid, {"BANDEIR" ,"C",20,0,"Bandeira",""})
    aAdd(aFieldsGrid, {"FORMPAG" ,"C",25,0,"Forma de Pagamento",""})
    aAdd(aFieldsGrid, {"NPARCEL" ,"C",2 ,0,"Número da parcela",""})
    aAdd(aFieldsGrid, {"QPARCEL" ,"C",2 ,0,"Quantidade de parcelas",""})
    aAdd(aFieldsGrid, {"NCARTAO" ,"C",15,0,"Número do cartão",""})
    aAdd(aFieldsGrid, {"CODTRAN" ,"C",25,0,"Código da transação",""})
    aAdd(aFieldsGrid, {"TID"     ,"C",25,0,"TID",""})
    aAdd(aFieldsGrid, {"CODAUTO" ,"C",10,0,"Código de autorização",""})
    aAdd(aFieldsGrid, {"NSU"     ,"C",10,0,"NSU",""})
    aAdd(aFieldsGrid, {"VALBRUT" ,"N",16,2,"Valor bruto","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALDESC" ,"N",16,2,"Valor descontado","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALLIQ"  ,"N",16,2,"Valor líquido","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALCOB"  ,"N",16,2,"Valor cobrado","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALPEND" ,"N",16,2,"Valor pendente","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALTOT"  ,"N",16,2,"Valor total","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"RECRAP"  ,"C",3 ,0,"Receba Rápido",""})
    aAdd(aFieldsGrid, {"CANAL"   ,"C",15,0,"Canal de venda",""})
    aAdd(aFieldsGrid, {"TPCAPT"  ,"C",20,0,"Tipo de captura",""})
    aAdd(aFieldsGrid, {"RESOP"   ,"C",10,0,"Resumo da operação",""})
    aAdd(aFieldsGrid, {"TAXA"    ,"N",16,2,"Taxas (%)","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"TARIFA"  ,"N",16,2,"Tarifa","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"CODVEND" ,"C",30,0,"Código da venda",""})
    aAdd(aFieldsGrid, {"NUMMAQ"  ,"C",15,0,"Número da máquina",""})
    aAdd(aFieldsGrid, {"PERCONS" ,"D",8 ,0,"Período considerado",""})
    aAdd(aFieldsGrid, {"NUMOPER" ,"C",15,0,"Número da operação",""})
    aAdd(aFieldsGrid, {"TAXAANT" ,"N",16,2,"Taxa de antecipação","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"TAXAEMB" ,"N",16,2,"Taxa de embarque","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"VALENTR" ,"N",16,2,"Valor da entrada","@E 9,999,999,999,999.99"})
    aAdd(aFieldsGrid, {"NUMPEDI" ,"C",30,0,"Número do pedido",""})
    aAdd(aFieldsGrid, {"NUMNOTA" ,"C",10,0,"Número da nota fiscal",""})
    aAdd(aFieldsGrid, {"ID"      ,"C",20,0,"ID",""})
    aAdd(aFieldsGrid, {"DTPGCIE" ,"D",8 ,0,"Data de pagamento na conta Cielo",""})

    oTabTMPGrid:= FWTemporaryTable():New(cAliasGrid)
    oTabTMPGrid:SetFields(aFieldsGrid)
    oTabTMPGrid:Create()

    cArq := tFileDialog( "Arquivo de planilha Excel (*.xlsx)",,,, .F., /*GETF_MULTISELECT*/)
    oExcel	:= YExcel():new(,cArq)

    If !Empty(cArq)
        
        If F914aExis(cArq)
            MsgStop("O arquivo: "+Alltrim(cArq)+", já foi importado anteriormente.")
            Return
        EndIF

        DBSelectArea(cAliasCab)
        DBSelectArea(cAliasGrid)

        For nContP := 1 To oExcel:LenPlanAt()	        //Ler as Planilhas
            oExcel:SetPlanAt(nContP)		            //Informa qual a planilha atual
            aTamLin	:= oExcel:LinTam() 		            //Linha inicio e fim da linha
            oProcess:SetRegua1(aTamLin[2])
            For nContL := aTamLin[1] to aTamLin[2]
                
                oProcess:IncRegua1("Lendo o arquivo, linha " + cValToChar(nContL) + " de " + cValToChar(aTamLin[2]) + "...")
      
                If nContL == 3
                            
                    RecLock(cAliasCab,.T.)
                        (cAliasCab)->DTIMPOR := dDataBase
                        (cAliasCab)->PERIODO := oExcel:GetValue(nContL,1)
                        (cAliasCab)->ARQUIVO := cArq
                    (cAliasCab)->(MSUnLock())

                ElseIF nContL > 5

                    RecLock(cAliasGrid,.T.)

                        (cAliasGrid)->STATUS  := oExcel:GetValue(nContL,1)
                        (cAliasGrid)->TPLANC  := oExcel:GetValue(nContL,2)
                        (cAliasGrid)->DESCRIC := oExcel:GetValue(nContL,3)
                        (cAliasGrid)->BANCO   := oExcel:GetValue(nContL,4)
                        (cAliasGrid)->AGENCIA := IIF(ValType(oExcel:GetValue(nContL,5)) == "N", AsString(oExcel:GetValue(nContL,5)), oExcel:GetValue(nContL,5))
                        (cAliasGrid)->CONTA   := oExcel:GetValue(nContL,6)
                        (cAliasGrid)->GRAVAME := oExcel:GetValue(nContL,7)
                        (cAliasGrid)->CNPJORI := oExcel:GetValue(nContL,8)
                        (cAliasGrid)->NOMEORI := oExcel:GetValue(nContL,9)
                        (cAliasGrid)->ESTABE  := IIF(ValType(oExcel:GetValue(nContL,10)) == "N", AsString(oExcel:GetValue(nContL,10)), oExcel:GetValue(nContL,10))
                        (cAliasGrid)->DTPAGAM := oExcel:GetValue(nContL,11)
                        (cAliasGrid)->DTLANCA := oExcel:GetValue(nContL,12)
                        (cAliasGrid)->DTAUTVE := oExcel:GetValue(nContL,13)
                        (cAliasGrid)->BANDEIR := oExcel:GetValue(nContL,14)
                        (cAliasGrid)->FORMPAG := oExcel:GetValue(nContL,15)
                        (cAliasGrid)->NPARCEL := IIF(ValType(oExcel:GetValue(nContL,16)) == "N", AsString(oExcel:GetValue(nContL,16)), oExcel:GetValue(nContL,16))
                        (cAliasGrid)->QPARCEL := IIF(ValType(oExcel:GetValue(nContL,17)) == "N", AsString(oExcel:GetValue(nContL,17)), oExcel:GetValue(nContL,17))
                        (cAliasGrid)->NCARTAO := oExcel:GetValue(nContL,18)
                        (cAliasGrid)->CODTRAN := IIF(ValType(oExcel:GetValue(nContL,19)) == "N", AsString(oExcel:GetValue(nContL,19)), oExcel:GetValue(nContL,19))
                        (cAliasGrid)->TID     := IIF(ValType(oExcel:GetValue(nContL,20)) == "N", AsString(oExcel:GetValue(nContL,20)), oExcel:GetValue(nContL,20))
                        (cAliasGrid)->CODAUTO := IIF(ValType(oExcel:GetValue(nContL,21)) == "N", AsString(oExcel:GetValue(nContL,21)), oExcel:GetValue(nContL,21))
                        (cAliasGrid)->NSU     := IIF(ValType(oExcel:GetValue(nContL,22)) == "N", AsString(oExcel:GetValue(nContL,22)), oExcel:GetValue(nContL,22))
                        (cAliasGrid)->VALBRUT := oExcel:GetValue(nContL,23)
                        (cAliasGrid)->VALDESC := Abs(oExcel:GetValue(nContL,24))
                        (cAliasGrid)->VALLIQ  := oExcel:GetValue(nContL,25)
                        (cAliasGrid)->VALCOB  := oExcel:GetValue(nContL,26)
                        (cAliasGrid)->VALPEND := oExcel:GetValue(nContL,27)
                        (cAliasGrid)->VALTOT  := oExcel:GetValue(nContL,28)
                        (cAliasGrid)->RECRAP  := oExcel:GetValue(nContL,29)
                        (cAliasGrid)->CANAL   := oExcel:GetValue(nContL,30)
                        (cAliasGrid)->TPCAPT  := oExcel:GetValue(nContL,31)
                        (cAliasGrid)->RESOP   := IIF(ValType(oExcel:GetValue(nContL,32)) == "N", AsString(oExcel:GetValue(nContL,32)), oExcel:GetValue(nContL,32))
                        (cAliasGrid)->TAXA    := oExcel:GetValue(nContL,33)
                        (cAliasGrid)->TARIFA  := oExcel:GetValue(nContL,34)
                        (cAliasGrid)->CODVEND := IIF(ValType(oExcel:GetValue(nContL,35)) == "N", AsString(oExcel:GetValue(nContL,35)), oExcel:GetValue(nContL,35))
                        (cAliasGrid)->NUMMAQ  := IIF(ValType(oExcel:GetValue(nContL,36)) == "N", AsString(oExcel:GetValue(nContL,36)), oExcel:GetValue(nContL,36))
                        (cAliasGrid)->PERCONS := IIF(ValType(oExcel:GetValue(nContL,37)) == "D", oExcel:GetValue(nContL,37), STOD(""))
                        (cAliasGrid)->NUMOPER := IIF(ValType(oExcel:GetValue(nContL,38)) == "N", AsString(oExcel:GetValue(nContL,38)), oExcel:GetValue(nContL,38))
                        (cAliasGrid)->TAXAANT := oExcel:GetValue(nContL,39)
                        (cAliasGrid)->TAXAEMB := oExcel:GetValue(nContL,40)
                        (cAliasGrid)->VALENTR := oExcel:GetValue(nContL,41)
                        (cAliasGrid)->NUMPEDI := IIF(ValType(oExcel:GetValue(nContL,42)) == "N", AsString(oExcel:GetValue(nContL,42)), oExcel:GetValue(nContL,42))
                        (cAliasGrid)->NUMNOTA := IIF(ValType(oExcel:GetValue(nContL,43)) == "N", AsString(oExcel:GetValue(nContL,43)), oExcel:GetValue(nContL,43))
                        (cAliasGrid)->ID      := oExcel:GetValue(nContL,44)
                        (cAliasGrid)->DTPGCIE := IIF(ValType(oExcel:GetValue(nContL,45)) == "D", oExcel:GetValue(nContL,45),STOD(""))
                    
                    (cAliasGrid)->(MSUnLock())

                EndIF
            Next 
        Next

        FWExecView("","FFIN001",MODEL_OPERATION_UPDATE,,{|| .T.},,,aButtons)

    EndIF 

    oExcel:Close() 

    oTabTMPCab:Delete()
    oTabTMPGrid:Delete()
    FWRestArea(aArea)

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
    Local bCommit := {|oModel|fSave(oModel)} //MsgRun("Gravando tabelas...", "Aguarde", { |oModel|fSave(oModel)})

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

    oModel := MPFormModel():New("FFIN001M", /*bPre*/, /*bPos*/, bCommit , /*bCancel*/)
    oModel:AddFields("MASTER", /*cOwner*/, oStructCab)
    oModel:AddGrid("GRID", "MASTER", oStructGrd)
    oModel:SetDescription(cTitulo)
    //oModel:GetModel("MASTER"):SetDescription("")
    oModel:SetPrimaryKey({})

Return oModel

/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
    Local oModel := FWLoadModel("FFIN001")
    Local oStructCab := FWFormViewStruct():New()
    Local oStructGrd := FWFormViewStruct():New()
    Local oView
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
        
        oStructGrd:AddField(aFieldsGrid[nId][1],;                   // 01 = Nome do Campo
                        StrZero(nId,2),;                            // 02 = Ordem
                        aFieldsGrid[nId][5],;                       // 03 = Título do campo
                        aFieldsGrid[nId][5],;                       // 04 = Descrição do campo
                        Nil,;                                       // 05 = Array com Help
                        aFieldsGrid[nId][2],;                       // 06 = Tipo do campo
                        aFieldsGrid[nId][6],;                       // 07 = Picture
                        Nil,;                                       // 08 = Bloco de PictTre Var
                        Nil,;                                       // 09 = Consulta F3
                        .F.,;                                       // 10 = Indica se o campo é alterável
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
    //oView:SetProgressBar(.T.)
    oView:AddField("VIEW1", oStructCab, "MASTER")
    oView:AddGrid("VIEW2" , oStructGrd,"GRID")

    oView:CreateHorizontalBox("CABEC" , 10 )
    oView:CreateHorizontalBox("GRID" , 90 )
    oView:SetOwnerView("VIEW1", "CABEC")
    oView:SetOwnerView("VIEW2", "GRID")

    oView:SetAfterViewActivate({|oView| ViewActv(oView)})
 
Return oView

/*---------------------------------------------------------------------*
 | Func:  ViewActv                                                     |
 | Desc:  Realiza o PUT nos campos para gravação na tabela FIF         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
Static Function ViewActv(oView)
    Local aAreaCab := (cAliasCab)->(FWGetArea())
	Local aAreaGrid := (cAliasGrid)->(FWGetArea())
    Local oModel := FWModelActive() 
    Local oModelCab := oModel:GetModel("MASTER")
    Local oModelGrid := oModel:GetModel("GRID")
    Local nTotTab := 0
    Local nLinWher := 0

    DBSelectArea(cAliasCab)
    (cAliasCab)->(DBGoTOP())

    While!(cAliasCab)->(EOF())

        oModelCab:LoadValue("DTIMPOR"  , (cAliasCab)->DTIMPOR )
        oModelCab:LoadValue("PERIODO"  , (cAliasCab)->PERIODO )
        oModelCab:LoadValue("ARQUIVO"  , (cAliasCab)->ARQUIVO )

    oView:Refresh('VIEW1')
    (cAliasCab)->(DBSkip())
    EndDo

    DBSelectArea(cAliasGrid)
    Count To nTotTab
    oProcess:SetRegua2(nTotTab)
    (cAliasGrid)->(DBGoTOP())

    While !(cAliasGrid)->(EOF())

        nLinWher++
        
        oProcess:IncRegua2("Carregando dados em tela " + AsString(nLinWher) + " de " + AsString(nTotTab) + "...")

        If nLinWher > 1
            oModelGrid:AddLine()
        EndIF 

        oModelGrid:LoadValue("STATUS"  , (cAliasGrid)->STATUS  )
        oModelGrid:LoadValue("TPLANC"  , (cAliasGrid)->TPLANC  )
        oModelGrid:LoadValue("DESCRIC" , (cAliasGrid)->DESCRIC )
        oModelGrid:LoadValue("BANCO"   , (cAliasGrid)->BANCO   )
        oModelGrid:LoadValue("AGENCIA" , (cAliasGrid)->AGENCIA )
        oModelGrid:LoadValue("CONTA"   , (cAliasGrid)->AGENCIA )
        oModelGrid:LoadValue("GRAVAME" , (cAliasGrid)->GRAVAME )
        oModelGrid:LoadValue("CNPJORI" , (cAliasGrid)->CNPJORI )
        oModelGrid:LoadValue("NOMEORI" , (cAliasGrid)->NOMEORI )
        oModelGrid:LoadValue("ESTABE"  , (cAliasGrid)->ESTABE  )
        oModelGrid:LoadValue("DTPAGAM" , (cAliasGrid)->DTPAGAM )
        oModelGrid:LoadValue("DTLANCA" , (cAliasGrid)->DTLANCA )
        oModelGrid:LoadValue("DTAUTVE" , (cAliasGrid)->DTAUTVE )
        oModelGrid:LoadValue("BANDEIR" , (cAliasGrid)->BANDEIR )
        oModelGrid:LoadValue("FORMPAG" , (cAliasGrid)->FORMPAG )
        oModelGrid:LoadValue("NPARCEL" , (cAliasGrid)->NPARCEL )
        oModelGrid:LoadValue("QPARCEL" , (cAliasGrid)->QPARCEL )
        oModelGrid:LoadValue("NCARTAO" , (cAliasGrid)->NCARTAO )
        oModelGrid:LoadValue("CODTRAN" , (cAliasGrid)->CODTRAN )
        oModelGrid:LoadValue("TID"     , (cAliasGrid)->TID     )
        oModelGrid:LoadValue("CODAUTO" , (cAliasGrid)->CODAUTO )
        oModelGrid:LoadValue("NSU"     , (cAliasGrid)->NSU     )
        oModelGrid:LoadValue("VALBRUT" , (cAliasGrid)->VALBRUT )
        oModelGrid:LoadValue("VALDESC" , (cAliasGrid)->VALDESC )
        oModelGrid:LoadValue("VALLIQ"  , (cAliasGrid)->VALLIQ  )
        oModelGrid:LoadValue("VALCOB"  , (cAliasGrid)->VALCOB  )
        oModelGrid:LoadValue("VALPEND" , (cAliasGrid)->VALPEND )
        oModelGrid:LoadValue("VALTOT"  , (cAliasGrid)->VALTOT  )
        oModelGrid:LoadValue("RECRAP"  , (cAliasGrid)->RECRAP  )
        oModelGrid:LoadValue("CANAL"   , (cAliasGrid)->CANAL   )
        oModelGrid:LoadValue("TPCAPT"  , (cAliasGrid)->TPCAPT  )
        oModelGrid:LoadValue("RESOP"   , (cAliasGrid)->RESOP   )
        oModelGrid:LoadValue("TAXA"    , (cAliasGrid)->TAXA    )
        oModelGrid:LoadValue("TARIFA"  , (cAliasGrid)->TARIFA  )
        oModelGrid:LoadValue("CODVEND" , (cAliasGrid)->CODVEND )
        oModelGrid:LoadValue("NUMMAQ"  , (cAliasGrid)->NUMMAQ  )
        oModelGrid:LoadValue("PERCONS" , (cAliasGrid)->PERCONS )
        oModelGrid:LoadValue("NUMOPER" , (cAliasGrid)->NUMOPER )
        oModelGrid:LoadValue("TAXAANT" , (cAliasGrid)->TAXAANT )
        oModelGrid:LoadValue("TAXAEMB" , (cAliasGrid)->TAXAEMB )
        oModelGrid:LoadValue("VALENTR" , (cAliasGrid)->VALENTR )
        oModelGrid:LoadValue("NUMPEDI" , (cAliasGrid)->NUMPEDI )
        oModelGrid:LoadValue("NUMNOTA" , (cAliasGrid)->NUMNOTA )
        oModelGrid:LoadValue("ID"      , (cAliasGrid)->ID      )
        oModelGrid:LoadValue("DTPGCIE" , (cAliasGrid)->DTPGCIE )
        
        oView:Refresh('VIEW2')
    
    (cAliasGrid)->(DBSkip())
    EndDo
    
    oModelGrid:GoLine(1)
    oView:Refresh('VIEW1')
    oView:Refresh('VIEW2')
    oView:SetNoDeleteLine('VIEW2')
    oView:SetNoInsertLine('VIEW2')
    FWRestArea(aAreaCab)
	FWRestArea(aAreaGrid)

Return

/*------------------------------------------------------------------*
 | Func:  fSave                                                     |
 | Desc:  Grava dados na tabela FIF.                                |
 | Obs.:  /                                                         |
 *-----------------------------------------------------------------*/
Static Function fSave(oModel)
    Local lRet := .T.
    Local oModelCab := oModel:GetModel("MASTER")
    Local oModelGrid := oModel:GetModel("GRID")
    Local cSeqFIF := ""
    Local nTamNUCOMP := TamSx3("FIF_NUCOMP")[1]
    Local nTamNSUTEF := TamSx3("FIF_NSUTEF")[1]
    Local nTamPARCEL := TamSx3("FIF_PARCEL")[1]
    Local cTPVenda   := ""
    Local nY 
    
    For nY := 1 To oModelGrid:Length()
        oModelGrid:GoLine(nY)
        
        If !oModelGrid:IsDeleted()
            
            cTPVenda := ""
            If Alltrim(oModelGrid:GetValue("TPLANC")) == "Venda parcelada" .OR.  Alltrim(oModelGrid:GetValue("TPLANC")) == "Venda crédito"
                cTPVenda := "C"
            ElseIF Alltrim(oModelGrid:GetValue("TPLANC")) == "Venda débito" .OR.  Alltrim(oModelGrid:GetValue("TPLANC")) == "Ajuste a débito"
                cTPVenda := "D"
            EndIF 
            
            If !Empty(cTPVenda) .AND. Alltrim(oModelGrid:GetValue("STATUS")) == "Pago"
                
                cSeqFIF := proxIdFIF()

                RecLock('FVR',.T.)
                    FVR_FILIAL := FWxFilial("FVR") 
                    FVR_DESCLE := "1"
                    FVR_IDPROC := cSeqFIF
                    FVR_NOMARQ := oModelCab:GetValue("ARQUIVO")
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

                RecLock('FV3',.T.)
                    FV3->FV3_FILIAL := FWxFilial("FV3") 
                    FV3->FV3_IDPROC := cSeqFIF
                    FV3->FV3_LINARQ := AsString(nY)
                    FV3->FV3_NOMARQ := oModelCab:GetValue("ARQUIVO")
                    FV3->FV3_DTPROC := dDataBase
                    FV3->FV3_HRPROC := Time()
                    FV3->FV3_CODEST := ""
                    FV3->FV3_NUCOMP := ""
                    FV3->FV3_MOTIVO := ""
                FV3->(MsUnlock())

                RecLock('FIF',.T.)
                    FIF->FIF_FILIAL := FWxFilial("FIF") 
                    FIF->FIF_TPREG  := "10"
                    FIF->FIF_INTRAN := ""
                    FIF->FIF_CODEST := oModelGrid:GetValue("NUMMAQ")
                    FIF->FIF_DTTEF  := oModelGrid:GetValue("DTAUTVE")
                    FIF->FIF_NURESU := oModelGrid:GetValue("RESOP")
                    FIF->FIF_NUCOMP := PadL(oModelGrid:GetValue("RESOP"), nTamNUCOMP, '0')
                    FIF->FIF_NSUTEF := PadL(oModelGrid:GetValue("NSU"), nTamNSUTEF, '0')
                    FIF->FIF_NUCART := oModelGrid:GetValue("NCARTAO")
                    FIF->FIF_VLBRUT := oModelGrid:GetValue("VALBRUT")
                    FIF->FIF_TOTPAR := oModelGrid:GetValue("QPARCEL")
                    FIF->FIF_VLLIQ  := oModelGrid:GetValue("VALLIQ")
                    FIF->FIF_DTCRED := oModelGrid:GetValue("DTLANCA")
                    FIF->FIF_PARCEL := PadL(oModelGrid:GetValue("NPARCEL"), nTamPARCEL, '0')
                    FIF->FIF_TPPROD := IIF(oModelGrid:GetValue("NPARCEL") == "Venda débito","D","C")
                    FIF->FIF_CAPTUR := "1"
                    FIF->FIF_CODRED := "340"
                    FIF->FIF_CODBCO := "341"
                    FIF->FIF_CODAGE := SubSTR(oModelGrid:GetValue("AGENCIA"),1,RAT("-",oModelGrid:GetValue("AGENCIA"))-1)
                    FIF->FIF_NUMCC  := SubSTR(oModelGrid:GetValue("CONTA"),1,RAT("-",oModelGrid:GetValue("CONTA"))-1)
                    FIF->FIF_VLCOM  := oModelGrid:GetValue("VALDESC")
                    FIF->FIF_TXSERV := oModelGrid:GetValue("TAXA")
                    FIF->FIF_CODLOJ := oModelGrid:GetValue("ESTABE")
                    FIF->FIF_CODAUT := oModelGrid:GetValue("CODAUTO")
                    FIF->FIF_CUPOM  := oModelGrid:GetValue("NUMNOTA")
                    FIF->FIF_SEQREG := PadL(oModelGrid:GetLine(), 6, '0')
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
                    FIF->FIF_PARALF := RetAsc(oModelGrid:GetValue("NPARCEL"),2,.T.)
                    FIF->FIF_CODFIL := FWxFilial("FIF")
                    FIF->FIF_CODBAN := "" //Posicione("MDE",2,FWxFilial("MDE")+Pad(UPPER(oModelGrid:GetValue("BANDEIR"))),"MDE_CODIGO")
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
                    FIF->FIF_ARQPAG := oModelCab:GetValue("ARQUIVO")
                    FIF->FIF_NSUARQ := ""
                    FIF->FIF_PGDES2 := ""
                    FIF->FIF_IDORAJ := ""
                    FIF->FIF_MSFIL  := FWxFilial("FIF")
                    FIF->FIF_MODPAG := ""
                FIF->(MsUnlock())
            
            EndIF 

        EndIf 
    Next
  
  lRet := FwFormCommit(oModel)

Return(lRet)

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
