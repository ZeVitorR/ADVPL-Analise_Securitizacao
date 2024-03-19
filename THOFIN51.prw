//Bibliotecas
#Include "Protheus.ch"
#Include 'TbiConn.ch'
#Include "TopConn.ch"
#Include "RwMake.ch"
#Include "Totvs.ch"
#Include "FWMVCDef.ch"
 

/*/{Protheus.doc} User Function THOFIN51
    Função para selecionar o que será realizado: Geração de carga ou atualição do sistema
    @type  Function
    @author user
    @since 05/12/2023
    @version version
    @param param_name, param_type, param_descr
    @RETURN RETURN_var, RETURN_type, RETURN_description
    @example
    (examples)
    @see (links_or_references)
    /*/
User Function THOFIN51()
    Local cPerg        := "THOFIN51"
    Local aColunas     := {}
    Local nX
    //Janela e componentes
    Private oDlgMk
    Private oPanGd
    Private oMarkBrowse
    //Tamanho da janela
    Private aTamanho := MsAdvSize()
    Private nJanLarg := aTamanho[5]
    Private nJanAltu := aTamanho[6]

    Private aRet        := {}
    Private aArea        := GetArea()
    Private nErro := 0
    Private cTempAlias := GetNextAlias()
    Private AliasTmp, cTitulo
    Private aCampos    := {}
    Private   pos         := 0
	
    // Pergunte(cPerg,.T.)
    IF ! Pergunte(cPerg,.T.)
        RETURN
    ENDIF
    
        //se o tipo for gerar carga
    IF( mv_par01 == 1 )
        MsgInfo("PARA CONSULTAR FILIAL COM PRODUTO NOVO CADASTRE-O NA CARTEIRA", "Aviso")
        SelFil()
        // aRet := AdmGetFil()
        // WHILE Len(aRet) == 0
        //     MsgInfo("Você não selecionou nenhuma filial! Selecione a desejada", "Aviso")
        //     aRet := AdmGetFil()
        // ENDDO

        
        RestArea(aArea)
    else //se for atualizar securitizadora
        
        //Adicionando o campo da tabela temporária
        aadd(aCampos, {"PRODUTO" , "C", 20, 00})
        aadd(aCampos, {"ID_UNICO", "C", 10, 00})
        aadd(aCampos, {"ZZSECUR" , "C", 10, 00})
        aadd(aCampos, {"ZZHTSEC" , "C", 10, 00})
        aadd(aCampos, {"FILIAL"  , "C", 06, 00})
        aadd(aCampos, {"PREFIXO" , "C", 03, 00})
        aadd(aCampos, {"NUM"     , "C", 09, 00})
        aadd(aCampos, {"PARCELA" , "C", 03, 00})
        aadd(aCampos, {"RAZAO"   , "C", 50, 00})
        aadd(aCampos, {"VALOR"   , "C", 17, 00})
        aadd(aCampos, {"ACRESC"  , "C", 09, 00})
        aadd(aCampos, {"PORTADO" , "C", 05, 00})
        aadd(aCampos, {"NUMBCO"  , "C", 10, 00})
        aadd(aCampos, {"OK"      , "C", 02, 00})

        //Criando a tabela temporária
        AliasTMP   := GetNextAlias()
        oTempTable := FWTemporaryTable():New(AliasTMP)
        oTemptable:SetFields(aCampos)
        oTemptable:AddIndex("1", {"PRODUTO"} )
        oTempTable:Create()
        cTitulo := oTempTable::GetTableNameForQuery()

        FWMsgRun(, {|oSay| lerAquivo(oSay) }, "Processando", "Realizando a leitura do arquivo")
        

        IF (nErro == 0)
            For nX := 1 To Len(aCampos)    
                AAdd(aColunas,FWBrwColumn():New())
                aColunas[Len(aColunas)]:SetData( &("{||"+aCampos[nX][1]+"}") )
                aColunas[Len(aColunas)]:SetTitle(aCampos[nX][1])
                aColunas[Len(aColunas)]:SetSize(aCampos[nX][3])
                aColunas[Len(aColunas)]:SetDecimal(aCampos[nX][4])              
            Next nX       
            
            aSeek := {}
            cCampoAux := "PRODUTO"
            aAdd(aSeek,{GetSX3Cache(cCampoAux, "X3_TITULO"), {{"", GetSX3Cache(cCampoAux, "X3_TIPO"), GetSX3Cache(cCampoAux, "X3_TAMANHO"), GetSX3Cache(cCampoAux, "X3_DECIMAL"), AllTrim(GetSX3Cache(cCampoAux, "X3_TITULO")), AllTrim(GetSX3Cache(cCampoAux, "X3_PICTURE"))}} } )
            DEFINE MSDIALOG oDlgMk TITLE '' FROM 000, 000  TO nJanAltu, nJanLarg COLORS 0, 16777215 PIXEL
                //Dados
                oPanGd := tPanel():New(001, 001, '', oDlgMk, , , , RGB(000,000,000), RGB(254,254,254), (nJanLarg/2)-1,     (nJanAltu/2 - 1))
                //Criando o FWMarkBrowse
                oMarkBrowse := FWMarkBrowse():New()
                oMarkBrowse:SetAlias(AliasTMP) 
                oMarkBrowse:oBrowse:SetDBFFilter(.T.)
                oMarkBrowse:oBrowse:SetUseFilter(.T.) //Habilita a utilização do filtro no Browse
                oMarkBrowse:oBrowse:SetFixedBrowse(.T.)               
                oMarkBrowse:SetDescription("Dados do Processamento do arquivo")
                oMarkBrowse:oBrowse:SetSeek(.T.,aSeek) //Habilita a utilização da pesquisa de registros no Browse
                oMarkBrowse:SetColumns(aColunas)
                oMarkBrowse:AddButton("Executar",{|| AltDados()} ,,3,,.F.)
                oMarkBrowse:AddButton("Selecionar Tudo", {|| SelAll()},,4,,.F.)
                oMarkBrowse:SetFieldMark( 'OK' )    //Campo que será marcado/descmarcado
                oMarkBrowse:SetOwner(oPanGd)
                oMarkBrowse:SetTemporary(.T.)
                //Ativando a janela
                oMarkBrowse:Activate() 
            ACTIVATE MsDialog oDlgMk CENTERED
            oMarkBrowse:DeActivate()
        ENDIF
        
        
    ENDIF
    RestArea( aArea )
RETURN

Static Function SelFil()
    Local aArea         := GetArea()
    Local aCampos := {}
    Local oTTable := Nil
    Local aColunas := {}
    local nX
    //Janela e componentes
    Private oDlgMark
    Private oPanGrid
    Private oMaBrowse
    Private cAlTmp := GetNextAlias()
    //Tamanho da janela
    Private aTamanho := MsAdvSize()
    Private nJanLarg := aTamanho[5]
    Private nJanAltu := aTamanho[6]

    //Adiciona as colunas que serão criadas na temporária
    aAdd(aCampos, { 'CODIGO', 'C', 2, 0,''}) //Código
    aAdd(aCampos, { 'FILIAL', 'C', 6, 0,''}) //Filial
    aAdd(aCampos, { 'NOME', 'C', 50, 0,''}) //Nome
    aAdd(aCampos, { 'OK', 'C', 2, 0,''}) //Flag para marcação

    //Cria a tabela temporária
    oTTable:= FWTemporaryTable():New(cAlTmp)
    oTTable:SetFields( aCampos )
    oTTable:Create()  

    Popula()
    //Percorrendo todos os campos da estrutura e adicionando na aColuna
    For nX := 1 To Len(aCampos)-1    
        AAdd(aColunas,FWBrwColumn():New())
        aColunas[Len(aColunas)]:SetData( &("{||"+aCampos[nX][1]+"}") )
        aColunas[Len(aColunas)]:SetTitle(aCampos[nX][1])
        aColunas[Len(aColunas)]:SetSize(aCampos[nX][3])
        aColunas[Len(aColunas)]:SetDecimal(aCampos[nX][4])              
    Next nX 

    DEFINE MSDIALOG oDlgMark TITLE 'SELECIONE UMA FILIAL' FROM 000, 000  TO 600, 800 COLORS 0, 16777215 PIXEL
        //Dados
        oPanGrid := tPanel():New(001, 001, '', oDlgMark, , , , RGB(000,000,000), RGB(254,254,254), (800/2)-1,     (600/2 - 1))
        oMaBrowse:= FWMarkBrowse():New()
        oMaBrowse:SetDescription("SELEÇÃO DE FILIAL") //Titulo da Janela
        oMaBrowse:SetAlias(cAlTmp)
        oMaBrowse:AddButton("Confirma",{|| FWMsgRun(, {|oSay| GeraExcel(oSay) }, "Processando", "Processando dados")} ,,3,,.F.)
        oMaBrowse:AddButton("Selecionar Tudo", {||oMaBrowse:AllMark()},,4,,.F.)
        oMaBrowse:SetTemporary(.T.) //Indica que o Browse utiliza tabela temporária
        oMaBrowse:SetFieldMark('OK')
        oMaBrowse:SetOwner(oPanGrid)
        oMaBrowse:SetColumns(aColunas)
        oMaBrowse:Activate()

    ACTIVATE MsDialog oDlgMark CENTERED
    
    RestArea(aArea)
    
    
Return 

Static Function Popula()
    dbSelectArea("SM0")
    dbSetOrder(1)
    SM0->(dbGoTop())

    While SM0->(!EOF())
        If SM0->M0_CODIGO == '02'

            if Select("cQryFil")>0

                cQryFil->(dbCloseArea())

            EndIf

            BeginSql Alias "cQryFil"

                SELECT ZA2_FILIAL FROM %table:ZA2% ZA2
                WHERE ZA2_FILIAL = %Exp:SM0->M0_CODFIL%
                AND ZA2.%notDel%
                GROUP BY ZA2_FILIAL
                
            EndSql
            
            While cQryFil->(!EOF())

                RecLock(cAlTmp, .T.)
                    (cAlTmp)->OK := Space(2)
                    (cAlTmp)->CODIGO := SM0->M0_CODIGO
                    (cAlTmp)->FILIAL := SM0->M0_CODFIL
                    (cAlTmp)->NOME   := SM0->M0_NOMECOM
                (cAlTmp)->(MsUnlock())

                cQryFil->(dbSkip())

            EndDo

        Endif
        SM0->(dbSkip())
    Enddo
Return

Static Function GeraExcel(oSay)
    Local cAliasTemp := ""
    // Local cATemp     := ""
    Local oFWMsExcelEx
    Local cHora      := SUBSTR(Time(),1,2)+SUBSTR(Time(),4,2)+SUBSTR(Time(),7,2)
    Local cWorkSheet := "DADOS"
    Local cTable     := "THOFIN51"
    // Local dir        := GetTempPath()
    Local cArquivo   := "C:\spool\" + cTable+AllTrim(cHora)+".xml"
    Local aLinhaAux  := {}
    Local nCount     := 1
    Local n, cSecur, nC
    local cMarca     := oMaBrowse:Mark()
    Local aSec       := {}
    oSay:SetText("Iniciando processamento...")
    dbSelectArea(cAlTmp)
    (cAlTmp)->(DBGoTop())
    if !(cAlTmp)->(EOF())
        
        While (cAlTmp)->(!EOF())
            if oMaBrowse:IsMark(cMarca)
                aAdd(aRet,(cAlTmp)->FILIAL)
            ENDIF
            (cAlTmp)->(dbSkip())
        EndDo
        if( len(aRet)>0)
            data1:=strzero(year(mv_par04),4)+strzero(month(mv_par04),2)+strzero(day(mv_par04),2)
            data2:=strzero(year(mv_par05),4)+strzero(month(mv_par05),2)+strzero(day(mv_par05),2)
        
            cSecur := ALLTRIM(mv_par06)
            AADD(aSec, Separa(cSecur, ",", .T.))
            for nC := 1 to LEN(aSec[1])
                IF( nC == 1)
                    cSecur := "'"+aSec[1][nC]+"'"
                else 
                    cSecur += ",'"+aSec[1][nC]+"'"   
                ENDIF  
            next
            
            //Buscando os produtos para ver se ira para securitizadora    
            cAliasTemp       := "SELECT E1_PRODUTO, R_E_C_N_O_ , E1_ZZSECUR, E1_ZZHTSEC, E1_FILIAL, E1_PREFIXO, E1_NUM, "
            cAliasTemp       += " E1_PARCELA, E1_VENCREA, E1_RAZAO, E1_VALOR, E1_ACRESC, (E1_VALOR + E1_ACRESC) VALOR_PARCELA,E1_EMISSAO, E1_BAIXA,E1_SALDO, E1_SALDO, E1_PORTADO, E1_NUMBCO,E1_TIPO "
            cAliasTemp       += " FROM " + RetSqlName("SE1")
            cAliasTemp           += " WHERE E1_FILIAL IN ("
            for n := 1 to Len(aRet)
                IF n == 1
                    cAliasTemp  += "'"+aRet[n]+"'"
                ELSE 
                    cAliasTemp  += ", '"+aRet[n]+"'"
                ENDIF
            next n
            cAliasTemp           += " ) "
            IF !EMPTY(mv_par04)
                cAliasTemp       += "AND E1_VENCREA BETWEEN '"+data1+"' AND '"+data2+"' "
            ENDIF  
            IF mv_par03 = 1
                cAliasTemp       += " AND E1_BAIXA = '' AND E1_SALDO <> 0"
            ENDIF
            IF Empty(mv_par06)
                cAliasTemp           += " AND E1_ZZSECUR = '' "
            ELSEIF SubStr(mv_par06, 1, 3) == "ZZZ" .OR. SubStr(mv_par06, 1, 3) == "zzz"
                cAliasTemp           += " AND E1_ZZSECUR <> '' "
            ELSE
                cAliasTemp           += " AND E1_ZZSECUR IN ("+cSecur+") "
            ENDIF
            IF mv_par07 = 2
                cAliasTemp       += " AND E1_PORTADO = '' "
            ENDIF
            cAliasTemp       += " AND E1_PREFIXO <> 'ZZZ' "
            cAliasTemp       += " AND D_E_L_E_T_= '' "
            cAliasTemp       += " ORDER BY E1_FILIAL, R_E_C_N_O_, E1_PREFIXO, E1_NUM, E1_PARCELA"
            TCQuery cAliasTemp New Alias "QRY_PRO"
            
            IF QRY_PRO->(EoF())
                oSay:SetText("Processamento concluído. Finalizando...")
                MsgInfo("Não possui nenhum dado a ser exibido!", "Aviso")
            Else
                //Criando o objeto que irá gerar o conteúdo do Excel
                oFWMsExcelEx       := FWMsExcelEx():New()
                oFWMsExcelEx:AddworkSheet(cWorkSheet)         
                    //Criando a Tabela
                    oFWMsExcelEx:AddTable(cWorkSheet, cTable) 
                    //Criando Colunas
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_PRODUTO", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "ID_UNICO", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_ZZSECUR", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_ZZHTSEC", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_FILIAL", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_PREFIXO", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_NUM", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_PARCELA", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_EMISSAO", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_VENCREA", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_BAIXA", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_RAZAO", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_VALOR", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_ACRESC", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "VALOR_PARCELA", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_SALDO", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_PORTADO", 1, 1)
                    oFWMsExcelEx:AddColumn(cWorkSheet, cTable, "E1_NUMBCO", 1, 1)            
                    
                    //Percorrendo os produtos
                    WHILE !QRY_PRO->(EoF())
                        //Criando a linha
                        IF !Empty(TRIM(QRY_PRO->E1_VENCREA)) .OR. QRY_PRO->E1_VENCREA != ' '
                            cDatavr   := SUBSTR(QRY_PRO->E1_VENCREA   ,7,2)+'/'+SUBSTR(QRY_PRO->E1_VENCREA   ,5,2)+'/'+left(QRY_PRO->E1_VENCREA  ,4)
                        Else
                            datavr   := ''
                        ENDIF                
                        IF !Empty(TRIM(QRY_PRO->E1_EMISSAO)) .OR. QRY_PRO->E1_EMISSAO != ' '
                            cDataem := SUBSTR(QRY_PRO->E1_EMISSAO,7,2)+'/'+SUBSTR(QRY_PRO->E1_EMISSAO,5,2)+'/'+left(QRY_PRO->E1_EMISSAO,4)
                        Else
                            cDataem := ''
                        ENDIF
                        IF !Empty(TRIM(QRY_PRO->E1_BAIXA)) .OR. QRY_PRO->E1_BAIXA != ' '
                            cDatabx := SUBSTR(QRY_PRO->E1_BAIXA,7,2)+'/'+SUBSTR(QRY_PRO->E1_BAIXA,5,2)+'/'+left(QRY_PRO->E1_BAIXA,4)
                        Else
                            cDatabx := ''
                        ENDIF                
                        aLinhaAux        := Array(18)
                        aLinhaAux[1]     := QRY_PRO->E1_PRODUTO
                        aLinhaAux[2]     := QRY_PRO->R_E_C_N_O_
                        aLinhaAux[3]     := QRY_PRO->E1_ZZSECUR
                        aLinhaAux[4]     := QRY_PRO->E1_ZZHTSEC
                        aLinhaAux[5]     := QRY_PRO->E1_FILIAL
                        aLinhaAux[6]     := QRY_PRO->E1_PREFIXO
                        aLinhaAux[7]     := QRY_PRO->E1_NUM
                        aLinhaAux[8]     := QRY_PRO->E1_PARCELA
                        aLinhaAux[9]     := cDataem
                        aLinhaAux[10]    := cDatavr
                        aLinhaAux[11]    := cDatabx
                        aLinhaAux[12]    := QRY_PRO->E1_RAZAO
                        aLinhaAux[13]    := QRY_PRO->E1_VALOR
                        aLinhaAux[14]    := QRY_PRO->E1_ACRESC
                        aLinhaAux[15]    := QRY_PRO->VALOR_PARCELA
                        aLinhaAux[16]    := QRY_PRO->E1_SALDO
                        aLinhaAux[17]    := QRY_PRO->E1_PORTADO
                        aLinhaAux[18]    := QRY_PRO->E1_NUMBCO
                        
                        nCount           += 1
                        
                        //Adiciona a linha no Excel
                        oFWMsExcelEx:AddRow(cWorkSheet, cTable, aLinhaAux)
                        
                        QRY_PRO->(DbSkip())
                    ENDDO
                    oSay:SetText("Processamento concluído. Finalizando...")

                oFWMsExcelEx:Activate()
                IF mv_par08 == ' ' .OR. mv_par08 == NIL
                    oFWMsExcelEx:GetXMLFile(cArquivo)
                else
                    cArquivo := AllTrim(mv_par08) + " - "+cTable+" "+AllTrim(cHora)+".xml
                    oFWMsExcelEx:GetXMLFile(cArquivo)
                ENDIF
                
                MsgInfo("Arquivo pronto para conferência! Acesse: "+cArquivo, "Aviso")
                
                oDlgMark:end()
            ENDIF    
            
            
            QRY_PRO->(DbCloseArea())
        ENDIF
    ENDIF
   

    
RETURN

Static Function lerAquivo(oSay)

    // Variaveis.
    Local     cFile       := ""
    Local     nHandle     := 0
    Local     qtd         := 0
    Local     cQueryB     := ""
    Local     cLinha      := ""
    Local     cTit        := ""
    Local     lPrim       := .T.
    Local     aCampos     := {}
    Local     aDados      := {}
    

    //Chama a função para buscar arquivos
    cFile := cArqSel := tFileDialog(;
        "CSV (separado por virgula) (*.csv)",;  // Filtragem de tipos de arquivos que serão selecionados
        "Seleção de Arquivos para Processamento",;  // Título da Janela para seleção dos arquivos
        ,;         // Compatibilidade
        "C:\spool",;  // Diretório inicial da busca de arquivos
        .F.,;  // Se for .T., será uma Save Dialog, senão será Open Dialog
        ;          // Se não passar parâmetro, irá pegar apenas 1 arquivo; Se for informado GETF_MULTISELECT será possível pegar mais de 1 arquivo; Se for informado GETF_RETDIRECTORY será possível selecionar o diretório
    )

    // Trava file para uso.
    nHandle := FT_FUSE(cFile)
    // Se houver erro de abertura abandona processamento
    IF nHandle = -1
        alert("Erro de processamento")
        nErro := 1
        RETURN 
    ENDIF
    // Posiciona na primeria linha
    FT_FGoTop()
    dbSelectArea(AliasTmp)
    // Enquanto não for final do arquivo continua lendo o mesmo.
    WHILE !FT_FEOF()
        qtd++;
        // Le conteudo da linha posicionada.
        cLinha := FT_FREADLN()
        IF (qtd == 1)
            cTit := Separa(cLinha, ";", .T.)
        elseIF (qtd == 2)
            aCampos := Separa(cLinha, ";", .T.)
			lPrim   := .F.
		Else
            pos++
            AADD(aDados, Separa(cLinha, ";", .T.))
            // realizando a busca se encontra o R_E_C_N_O_ do arquivo no SE1
            cQueryB := "SELECT * FROM "+ RetSqlName("SE1") + " WHERE R_E_C_N_O_ = " + aDados[pos][2] + " AND E1_PRODUTO = '"+ aDados[pos][1] +"' AND E1_FILIAL = '"+aDados[pos][5]+"'"
            TCQuery cQueryB New Alias "QRY_B"
            //se não possuir nenhum dado da busca então o R_E_C_N_O_ está invalido ou não existe
            IF QRY_B->(EoF())
                Alert("Existe uma inconsistencia dos dados originais, gere uma nova carga e coloque o arquivo novamente")
                QRY_B->(DbCloseArea())
                (AliasTmp)->(DbCloseArea())
                nErro := 1
                RETURN 
            ELSE 
                RecLock((AliasTmp),.T.)
                (AliasTmp)->PRODUTO    := aDados[pos][1]
                (AliasTmp)->ID_UNICO      := aDados[pos][2]
                (AliasTmp)->ZZSECUR    := aDados[pos][3]
                (AliasTmp)->ZZHTSEC    := QRY_B->E1_ZZSECUR
                (AliasTmp)->FILIAL     := aDados[pos][5]
                (AliasTmp)->PREFIXO    := aDados[pos][6]
                (AliasTmp)->NUM        := aDados[pos][7]
                (AliasTmp)->PARCELA    := aDados[pos][8]
                (AliasTmp)->RAZAO      := aDados[pos][12]
                (AliasTmp)->VALOR      := aDados[pos][13]
                (AliasTmp)->ACRESC     := aDados[pos][14]
                (AliasTmp)->PORTADO    := aDados[pos][17]
                (AliasTmp)->NUMBCO     := aDados[pos][18]
                (AliasTmp)->NUMBCO     := aDados[pos][18]
                MsUnLock()
            ENDIF
            QRY_B->(DbCloseArea())
		ENDIF

        
        // Proxima linha.
        FT_FSKIP()

    ENDDO


    // Libera arquivo.
    FT_FUSE()
    
    MsgInfo("Realizado a leitura de "+cValToChar(pos)+" linhas do documento" , "Aviso")

    
RETURN 

Static Function SelAll()
    Local cQuery := ''

    cQuery := "SELECT COUNT(OK) AS NumOk FROM " +  cTitulo + " WHERE OK <> '' "
    TCQuery cQuery New Alias "QRY_QTAD"

    IF QRY_QTAD->NumOk == pos
        oMarkBrowse:AllMark()
    ElseIF QRY_QTAD->NumOk == 0
        oMarkBrowse:AllMark()
    Else
        cQuery := "SELECT ID_UNICO FROM " +  cTitulo + " WHERE OK = '' "
        TCQuery cQuery New Alias "QRY_SEL"
        (AliasTmp)->(DBGoTop())
        WHILE !(AliasTmp)->(EOF())
            IF (AliasTMP)->ID_UNICO == QRY_SEL->ID_UNICO
                oMarkBrowse:MarkRec()
            ENDIF
            (AliasTmp)->(DbSkip())
            QRY_SEL->(DbSkip())
        ENDDO
        QRY_SEL->(dbCloseArea())
    ENDIF
    QRY_QTAD->(dbCloseArea())
    
RETURN 

//função para alterar a securitizadoras do SE1
STATIC Function AltDados()
    Local cMarca := oMarkBrowse:Mark()
    local n:=0
    (AliasTmp)->(DBGoTop())
    WHILE !(AliasTmp)->(EOF())
        if oMarkBrowse:IsMark(cMarca)
            n++
        endif
        (AliasTmp)->(DbSkip())
    Enddo

    IF MsgYesNo("Você selecionou:" + cValtoChar(n) + " linhas" + Chr(13) + Chr(10) + "Você confirma essa seleção?" , "Confirma?")
        FWMsgRun(, {|oSay| ProcessaDados(oSay) }, "Processando", "Processando os dados")
    
        

    ENDIF
    
    
    
RETURN .T.

Static Function ProcessaDados(oSay)
    Local cMarca := oMarkBrowse:Mark()
    Local ctn := 0

    oSay:SetText("Iniciando processamento...")
    dbSelectArea("SE1")
    SE1->(DBGoTop())
    (AliasTmp)->(DBGoTop())
    WHILE !(AliasTmp)->(EOF())
        IF oMarkBrowse:IsMark(cMarca)
            SE1->(MSSEEK((AliasTmp)->FILIAL + (AliasTmp)->PREFIXO + (AliasTmp)->NUM + (AliasTmp)->PARCELA) )
            IF SE1->(RecNo()) = Val((AliasTmp)->ID_UNICO)
                RecLock("SE1",.F.)
                    SE1->E1_ZZSECUR := (AliasTmp)->ZZSECUR
                    SE1->E1_ZZHTSEC := (AliasTmp)->ZZHTSEC
                SE1->(MsUnLock())
                
                ctn ++
            ENDIF
            IF(mv_par02 == 1)
                dbSelectArea("SB1")
                SB1->(DBGoTop())
                SB1->(MSSEEK(SE1->E1_FILIAL + SE1->E1_PRODUTO))
                IF SB1->B1_COD == SE1->E1_PRODUTO
                    RecLock("SB1",.F.)
                        SB1->B1_ZZSECUR := ALLTRIM(SE1->E1_ZZSECUR)
                    SB1->(MsUnLock())
                ENDIF
                SB1->(dbCloseArea())
            ENDIF
        ENDIF
        (AliasTmp)->(DbSkip())
    ENDDO

    
        
       
    oSay:SetText("Concluindo o processamento...")

    MsgInfo("Foram atualizados " +cValToChar(ctn)+ " valores")
    SE1->(dbCloseArea())
    oDlgMk:end()
RETURN
