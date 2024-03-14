Public Ordem As String
Public Material As String
Public Carcaca As String
Public Distanciadoras As String
Public Eixo As String

Public SapGuiAuto As Variant
Public SAPApp As Variant
Public SAPCon As Variant
Public session As Variant
Public Connection As Variant
Public WScript As Variant

' Numero do componente na tabela co03
Public ComponenteBarraNumero As String

' Texto do componente na tabela co03
Public ComponenteBarraDescricao As String

' Texto do componente na tabela co03
Public ComponenteBarraTipo As String
Public ComponenteBarraComprimento As String
Public ComponenteBarraQuantidade As Integer

Public MaterialCobreIdentificador As String
Public MaterialCobreAlmoxarifado As String

Sub Main()
    Call DeclararVariaveis
    Call ConectarSAP
    Call ArrumarPlanilha
    
    MsgBox "Programa encerrado"
    MACRO_RODADA
End Sub


Sub MACRO_RODADA()
    caminhoArquivo = "Q:\GROUPS\BR_SC_JGS_WM_DEPARTAMENTO_CALDEIRARIA\DEPARTAMENTO DE CALDEIRARIA\02 - DOCUMENTOS\16 - RELATORIOS BI\Banco de dados de macros\banco de dados - macros.txt"
    fileNumber = FreeFile
    On Error Resume Next
        Open caminhoArquivo For Append As fileNumber
        Print #fileNumber, "Macro | " & ActiveWorkbook.Name & " | " & Now
        Close fileNumber
    On Error GoTo 0
End Sub

Sub ArrumarPlanilhaSAP(PEP As String)
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncoois"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_DISPO-LOW").Text = "410"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_DISPO-HIGH").Text = "412"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PROJN-LOW").Text = PEP
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    If Not session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell", False) Is Nothing Then
        TabelaSAP = "wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell"
        Ordem = session.findById(TabelaSAP).GetCellValue(0, "AUFNR")
        Material = session.findById(TabelaSAP).GetCellValue(0, "MATNR")
        Carcaca = session.findById(TabelaSAP).GetCellValue(0, "MATXT")
        Carcaca = Replace(Carcaca, "ROTOR COMPLETO MIT ", "")
        Carcaca = Replace(Carcaca, " A/H", "")
        
    Else
        session.findById("wnd[0]").sendVKey 0
        MsgBox "Projeto " & PEP & " não achado"
    End If
    
End Sub

Sub ArrumarPlanilha()
    quantidadeLinhas = ContarLinhas(PlanilhaComando)
    Projeto = PlanilhaComando.Range("C1").Value
    VaiArrumarPlanilha = True
    
    For Linha = 8 To quantidadeLinhas
        Projeto = PlanilhaComando.Cells(Linha, 3).Value
    
        If PlanilhaComando.Cells(Linha, 3) <> "" Then
            Arrumou = False
            'se faltar ordem
            If PlanilhaComando.Cells(Linha, 1) = "" Then
                ArrumarPlanilhaSAP (Projeto)
                PlanilhaComando.Cells(Linha, 1).Value = Ordem
                Arrumou = True
            End If
            'se faltar material
            If PlanilhaComando.Cells(Linha, 2) = "" Then
                
                ArrumarPlanilhaSAP (Projeto)
                PlanilhaComando.Cells(Linha, 2).Value = Material
            End If
            'se faltar carcaça
            If PlanilhaComando.Cells(Linha, 10) = "" Then
                ArrumarPlanilhaSAP (Projeto)
                PlanilhaComando.Cells(Linha, 10).Value = Carcaca
            End If
            'se faltar distanciadoras
            If PlanilhaComando.Cells(Linha, 6) = "" Then
                PuxarDistanciadoras
                PlanilhaComando.Cells(Linha, 6).Value = Distanciadoras
            End If
            
            
            
            If Arrumou Then
                Call PegarDadosBarra
            End If
            
        End If
        
    Next Linha
    
End Sub

Sub PuxarDistanciadoras()
    Call DeclararVariaveis
    TotalLinhasEstatores = ContarLinhas(ws_principal)
    
    For Linha = 3 To TotalLinhasEstatores
        If ws_principal.Cells(Linha, 6).Value = "" And ws_principal.Cells(Linha, 1).Value <> "" Then
            
            If ws_principal.Cells(Linha, 1).Value <> "ORDEM" Then
                texto = ConsultarSAP(ws_principal.Cells(Linha, 1).Value)
                ws_principal.Cells(Linha, 6) = texto
            End If
            
        End If
    Next Linha
End Sub

Function ConsultarSAP(Ordem)
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nco03"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = Ordem
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    
    On Error GoTo semdist
    
    For Linha = 0 To 25
        On Error GoTo errodosap
        descricaoMaterial = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2," & Linha & "]").Text
        On Error GoTo semdist
        If InStr(1, descricaoMaterial, "DIST") <> 0 Then
            QtdNec = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3," & Linha & "]").Text
            QtdConf = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DVMENG[11," & Linha & "]").Text
            QtdRet = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DENMNG[12," & Linha & "]").Text
            
            If QtdNec = QtdConf Or QtdNec = QtdRet Then
                If ConsultarSAP = "" Then
                    ConsultarSAP = "OK"
                End If
            Else
                If ConsultarSAP = "" Then
                    ConsultarSAP = "Falta dist " & session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & Linha & "]").Text
                ElseIf ConsultarSAP = "OK" Then
                    ConsultarSAP = "Falta dist " & session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & Linha & "]").Text
                Else
                    ConsultarSAP = ConsultarSAP & " e " & session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & Linha & "]").Text
                End If
            End If
        'Exit Function
        End If
        
        If InStr(1, descricaoMaterial, "EIXO") <> 0 Then
            QtdNec = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3," & Linha & "]").Text
            QtdConf = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DVMENG[11," & Linha & "]").Text
            QtdRet = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DENMNG[12," & Linha & "]").Text
            If QtdNec = QtdConf Or QtdNec = QtdRet Then
                If ConsultarSAP = "" Then
                    Eixo = "OK"
                End If
            Else
                If ConsultarSAP = "" Then
                    Eixo = "Falta " & session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & Linha & "]").Text
                ElseIf ConsultarSAP = "OK" Then
                    Eixo = "Falta " & session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & Linha & "]").Text
                Else
                    Eixo = Eixo & " e " & session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & Linha & "]").Text
                End If
            End If
        End If
    Next Linha
    
semdist:
    If ConsultarSAP = "" Then
        ConsultarSAP = "Não tem dist"
    End If
    Exit Function
errodosap:
    MsgBox ("FECHA E ABRE O SAP!")
    End
End Function

Sub LimparAuxiliares()
    Call LimparTabela(ws_aux, 1)
    Call LimparTabela(ws_aux2, 1)
End Sub

Function CalcularPesoSolido()
    ComprimentoBarrona = 6000
    QuantidadeBarrinhaPorBarrona = ComprimentoBarrona / ComponenteBarraComprimento
    QuantidadeBarrinhaPorBarrona = Application.WorksheetFunction.RoundDown(QuantidadeBarrinhaPorBarrona, 0)
    
    QuantidadeBarrona = ComponenteBarraQuantidade / QuantidadeBarrinhaPorBarrona
    QuantidadeBarrona = Application.WorksheetFunction.RoundUp(QuantidadeBarrinhaPorBarrona, 0)
    
    PesoTotal = CalcularQuilosBarrinha()
    Debug.Print "Buscou peso total a ser solicitado"
    
    CalcularPesoSolido = ComponenteBarraQuantidade * PesoTotal
    
    Debug.Print "Comprimento barrinha:", ComponenteBarraComprimento
    Debug.Print "Comprimento barrona:", ComprimentoBarrona
    Debug.Print "Quantidade barrinha por barrona:", QuantidadeBarrinhaPorBarrona
    Debug.Print "Quantidade barrinha por barrona:", QuantidadeBarrinhaPorBarrona
End Function

Sub PegarDadosBarra()

    Call ConectarSAP
    Call EntrarTelaCO03
    Call EntrarSinteseComponentes
    Call PegarDadosBarraTabelaSAP
    
    QuiloBarrinha = CalcularQuilosBarrinha()
    TotalQuilos = QuiloBarrinha * ComponenteBarraQuantidade

    Call EntrarTelaMMBE
        MaterialCobreAlmoxarifado = PegarMaterialCobreAlmoxarifado()
    
    If PegarMaterialCobreAlmoxarifado = False Then
        Debug.Print ("Almoxarifado não encontrado!")
        Exit Sub
    End If
    
    Debug.Print "Ordem: ", Ordem
    Debug.Print "Almoxarifado: ", MaterialCobreAlmoxarifado
    Debug.Print "Total kg a ser solicitado: ", TotalQuilos
End Sub

Function PegarBarraComprimento(ComponenteBarraDescricao)
    X = InStrRev(ComponenteBarraDescricao, "X")
    If X > 0 Then
        mm = InStr(ComponenteBarraDescricao, "mm")
        If mm > 0 Then
            numero = Mid(ComponenteBarraDescricao, X + 1, mm - X - 1)
            numero = Trim(numero)
            PegarBarraComprimento = numero + 0
            Exit Function
        End If
    End If
    PegarNumeroAntesDeMM = False
End Function

Function CalcularQuilosBarrinha()
    
    Call EntrarTelaCS12
    MaterialCobreIdentificador = PegarMaterialCobreIdentificador()
    ComponenteBarraDescricao = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").GetCellValue(0, "OJTXP")
    DescricaoOriginal = ComponenteBarraDescricao
    X = InStrRev(ComponenteBarraDescricao, "X")
    
    mm = InStr(ComponenteBarraDescricao, "mm")
    If mm > 0 Then
        ComponenteBarraDescricao = Left(ComponenteBarraDescricao, mm - 1)
        ComponenteBarraDescricao = Trim(ComponenteBarraDescricao)
    End If
        
     For i = 1 To Len(ComponenteBarraDescricao)
        If IsNumeric(Mid(ComponenteBarraDescricao, i, 1)) Then
            NumerosExtraidos = Mid(ComponenteBarraDescricao, i)
            Exit For
        End If
    Next i
    
    ' Verificar se é trapezoidal:
    
    If InStr(DescricaoOriginal, "TRAP") Then
        ' Se na string da descrição tem algo como "C11000", vai resultar em "11000 10"
        
            posicaoX1 = InStr(NumerosExtraidos, "X")
            LarguraCima = Left(NumerosExtraidos, posicaoX1 - 1)
            LarguraCima = CDec(Right(LarguraCima, 2)) ' Pra resolver isso, extraímos os dois ultimos numeros da string
    
            NumerosExtraidos = Mid(NumerosExtraidos, posicaoX1 + 1)
        
            posicaoX2 = InStr(NumerosExtraidos, "X")
            LarguraBaixo = CDec(Left(NumerosExtraidos, posicaoX2 - 1))
            NumerosExtraidos = Mid(NumerosExtraidos, posicaoX2 + 1)
            
            posicaoX3 = InStr(NumerosExtraidos, "X")
            Altura = CDec(Left(NumerosExtraidos, posicaoX3 - 1))
            NumerosExtraidos = Mid(NumerosExtraidos, posicaoX3 + 1)
            
            Comprimento = CDec(NumerosExtraidos)
            
            ' O comprimento as vezes falta caractéres, arrumar quando acontece
            If Comprimento < 100 Then
                Message = ""
                Comprimento = ""
                While Comprimento = ""
                    Comprimento = InputBox("Comprimento não detectado, veja o comprimento no SAP e informe ele" & vbNewLine & "EX: se é 1.041mm informe 1041", "Atenção")
                Wend
                Comprimento = CDec(Comprimento)
            End If
            
            ' Calculo do peso do trapézio
            AreaTrapezio = (LarguraBaixo + LarguraCima) * Altura / 2
            VolumeTrapezio = AreaTrapezio * Comprimento
            
            If InStr(DescricaoOriginal, "COBRE") Then
                Densidade = 8.96
            End If
            
            If InStr(DescricaoOriginal, "LATAO") Then
                Densidade = 8.73
            End If
            
            PesoTrapezio = VolumeTrapezio * Densidade / 1000000
            PesoTrapezio = Format(PesoTrapezio, "0.00")
            CalcularQuilosBarrinha = CDec(PesoTrapezio)
    End If
    
    ' Verificar se é retangular:
    If InStr(DescricaoOriginal, "RET") Then
        ' Se na string da descrição tem algo como "C11000", vai resultar em "11000 10"
        
            posicaoX1 = InStr(NumerosExtraidos, "X")
            Largura = Left(NumerosExtraidos, posicaoX1 - 1)
            Largura = Right(Largura, 2) ' Pra resolver isso, extraímos os dois ultimos numeros da string
            
            NumerosExtraidos = Mid(NumerosExtraidos, posicaoX1 + 1)
            
            posicaoX3 = InStr(NumerosExtraidos, "X")
            Altura = Left(NumerosExtraidos, posicaoX3 - 1)
            NumerosExtraidos = Mid(NumerosExtraidos, posicaoX3 + 1)
            
            Comprimento = NumerosExtraidos + 0
            
             If InStr(DescricaoOriginal, "COBRE") Then
                Densidade = 8.96
            End If
            
            If InStr(DescricaoOriginal, "LATAO") Then
                Densidade = 8.73
            End If
            
            ' Calculo do peso do retangulo
            VolumeRetangulo = Largura * Altura * Comprimento
            PesoBarra = VolumeRetangulo * Densidade / 1000000
            CalcularQuilosBarrinha = PesoBarra
    End If
End Function



Function PegarDadosBarraTabelaSAP()
    For Linha = 0 To 25:
        ComponenteBarraDescricao = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2," & Linha & "]").Text
        
        If InStr(1, ComponenteBarraDescricao, "BARRA COBRE") <> 0 Or InStr(1, ComponenteBarraDescricao, "BARRA LATAO") <> 0 Then
            If InStr(1, ComponenteBarraDescricao, "RET") = 0 Then
                ComponenteBarraNumero = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & Linha & "]").Text
                ComponenteBarraQuantidade = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3," & Linha & "]").Text
                ComponenteBarraComprimento = PegarBarraComprimento(ComponenteBarraDescricao)
                Debug.Print "Buscou comprimento"
                Exit For
            End If
        End If
    Next Linha
End Function


Function PegarBarraQuantidade()
    session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3,9]").SetFocus
End Function

Function PegarMaterialCobreIdentificador()
    Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
    grid.currentCellColumn = "DOBJT"
    PegarMaterialCobreIdentificador = grid.GetCellValue(grid.currentCellRow, "DOBJT")
End Function

Function PegarMaterialCobreAlmoxarifado()

    On Error GoTo NotExist
        session.findById("wnd[0]/usr/ctxtIO_MENGEINH").SetFocus
        Set tree = session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]")
    On Error GoTo RS02
        session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").selectedNode = "          4"
        PegarMaterialCobreAlmoxarifado = "RS01"
        Exit Function
        
RS02:
    PegarMaterialCobreAlmoxarifado = "RS02"
    Exit Function
NotExist:
    PegarMaterialCobreAlmoxarifado = False
    Exit Function
    
End Function

Sub EntrarTelaCO03()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nco03"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = Ordem
    session.findById("wnd[0]/tbar[0]/btn[0]").press
End Sub

Sub EntrarSinteseComponentes()
    session.findById("wnd[0]/tbar[1]/btn[6]").press
End Sub

Sub EntrarTelaCS12()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncs12"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRC29L-MATNR").Text = ComponenteBarraNumero
    session.findById("wnd[0]/usr/ctxtRC29L-CAPID").Text = "pp01"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
End Sub

Sub EntrarTelaMMBE()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmmbe"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").Text = MaterialCobreIdentificador
    session.findById("wnd[0]/usr/ctxtMS_WERKS-LOW").Text = "1200"
    session.findById("wnd[0]/usr/ctxtMS_LGORT-LOW").Text = "RS01"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
End Sub

Sub EntrarTelaZTMM292()
    CodigoLocalEntrega = "000000071"
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nztmm292"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/chkWA_TRANSFERENCIA-SOL_TRANSF").Selected = True
    session.findById("wnd[0]/usr/ctxtWA_ZTBMM_248-CD_DELIVERY_LOC").Text = CodigoLocalEntrega
    session.findById("wnd[0]/usr/chkWA_TRANSFERENCIA-SOL_TRANSF").SetFocus
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtWA_TRANSFERENCIA-WERKS").Text = "1200"
    session.findById("wnd[0]/usr/ctxtWA_TRANSFERENCIA-LGORT").Text = "IA04"
End Sub



'Função para contar n° de linhas ativas na planilha
Function ContarLinhas(ws_temp As Worksheet)
    ContarLinhas = ws_temp.UsedRange.Rows.Count
End Function

'Função para contar n° de colunas ativas na planilha
Function ContarColunas(ws_temp As Worksheet)
    ContarLinhas = ws_temp.UsedRange.Columns.Count
End Function

' Função para limpar uma planilha a partir da linha "n"
Sub LimparTabela(ws As Worksheet, n As Long)
    Dim r As String
    r = n & ":1048576"
    ws.Activate
    Rows(r).Select
    Selection.Delete Shift:=xlUp
End Sub

' Função para incrementar a performance do código, desativando atualização de toda a planilha durante o código
Sub IncrementarPerformance(op As Boolean)
    Select Case op
    Case True
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    Case False
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End Select
End Sub

' Faz um autofit em todas as colunas
Sub AjustarColunas(ws As Worksheet)
    ws.Activate
    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub

Sub ConectarSAP()
    'SAP
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
     Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)
    If Not IsObject(Application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject Application, "on"
    End If
End Sub
