Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Threading
Imports Alianca.Seguranca.BancoDados
Imports Alianca.Sistema.Funcoes.TrataParametros
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports SEGP1619.SEGP1619.Utils

Public Class frmPrincipal
    Private oclsPrincipal As New clsPrincipal
    
#Region "Variáveis"
    Public iConexao
    Public motivoRecusa As String
    Public ArquivoSS As String
    Public dsDetalheErros As New DataSet
    Public tbDetalheErros As New System.Data.DataTable

    Private dadosArray() As String = {"11111111111", "22222222222", "33333333333", "44444444444", "55555555555", _
                                      "66666666666", "77777777777", "88888888888", "99999999999", "00000000000"}
#End Region

#Region "Funções"

    Private Function ValidaCPF(ByVal CPF As String) As Boolean
        Dim CPF_valida As String
        Dim i, x, n1, n2 As Integer

        CPF_valida = CPF.Replace(".", "").Replace("-", ".")
        CPF_valida = "00000000000" + CPF_valida.Trim
        CPF_valida = CPF_valida.Substring(CPF_valida.Length - 11, 11)

        For i = 0 To dadosArray.Length - 1
            If dadosArray(i).Equals(CPF_valida) Then
                Return False
            End If
        Next
        For x = 0 To 1
            n1 = 0
            For i = 0 To 8 + x
                n1 = n1 + Val(CPF_valida.Substring(i, 1)) * (10 + x - i)
            Next
            n2 = 11 - (n1 - (Int(n1 / 11) * 11))
            If n2 = 10 Or n2 = 11 Then n2 = 0

            If n2 <> Val(CPF_valida.Substring(9 + x, 1)) Then
                Return False
            End If
        Next

        Return True
    End Function
#End Region

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub frmPrincipal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Parametro para Control-M 
        'C:\SEGP1619.0001.log CTM 3 SVCCONTROLM SURA001 N

        Try
            If oclsPrincipal.VerificaControlM(Command) Then
                iConexao = cCon.ConfiguraConexao(System.Reflection.Assembly.GetExecutingAssembly, oclsPrincipal.AmbienteId)
                oclsPrincipal.TituloAmbiente(Me, oclsPrincipal.AmbienteId)
            Else
                Me.Dispose()
                End
            End If

            cCon.ConnectionTimeout = cCon.TimeTimeout.Infinito

            TrataParametros(oclsPrincipal.codigoTarefa)

            'oclsPrincipal.GravaLogCTM("OK", "SEGP1619", "Processo executado com sucesso: " & oclsPrincipal.codigoTarefa)

        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
        End Try
    End Sub

    Private Sub TrataParametros(ByVal codRelatorio As String)
        codRelatorio = codRelatorio.Replace("""", "")

        Select Case codRelatorio
            Case "SURA001"
                ImportacaoArquivosSura()
            Case "SURA002"
                ProcessarBaixasRecebidas()
        End Select
    End Sub


    Private Sub ImportacaoArquivosSura()
        Dim sSQL As String

        Dim dsConsulta As DataSet
        Dim caminhoEntradaSURA As String
        Dim caminhoSaidaSURA As String
        Dim nomeArquivo As String
        Dim nomeArquivoErros As String
        Dim tp_ramo_id As Integer

        Dim campo As String

        'captar os parametros
        Try
            sSQL = "SELECT "
            sSQL = sSQL & "campo, valor "
            sSQL = sSQL & "FROM "
            sSQL = sSQL & " controle_sistema_db.dbo.parametro_tb with (nolock) "
            sSQL = sSQL & "WHERE "
            sSQL = sSQL & " sigla_sistema = 'SEGP1619' "
            sSQL = sSQL & " and ambiente_id = " & oclsPrincipal.AmbienteId.ToString

            dsConsulta = cCon.ExecuteDataset(CommandType.Text, sSQL)

            caminhoEntradaSURA = ""
            caminhoSaidaSURA = ""

            For Each registro As DataRow In dsConsulta.Tables(0).Rows
                campo = registro("campo").ToString()
                Select Case campo
                    Case "caminho_entrada_sura"
                        caminhoEntradaSURA = registro("valor").ToString()
                    Case "caminho_saida_sura"
                        caminhoSaidaSURA = registro("valor").ToString()
                End Select
            Next

            dsConsulta.Dispose()

            'Iniciar o dataset para armazenar os erros
            dsDetalheErros.Tables.Add(tbDetalheErros)
            dsDetalheErros.Tables(0).Columns.Add("Data", GetType(String))
            dsDetalheErros.Tables(0).Columns.Add("Arquivo", GetType(String))
            dsDetalheErros.Tables(0).Columns.Add("MotivoRecusa", GetType(String))

            'PARA TESTAR DESCOMENTAR ISSO
            'caminhoEntradaSURA = "\\10.0.35.107\users\temporario\Nova\SURA\ENTRADA\"
            'caminhoSaidaSURA = "\\10.0.35.107\users\temporario\Nova\SURA\SAIDA\"
            'FIM TESTE

            'Para testes

            'caminhoEntradaSURA = "C:\Temp\Sura\Entrada\"
            'caminhoSaidaSURA = "C:\Temp\Sura\Saida\"
            'fim teste

            'LENDO OS ARQUIVOS TXT DA PASTA INDICADA NA TABELA PARAMETROS

            Dim di As New IO.DirectoryInfo(caminhoEntradaSURA)
            Dim diar1() As IO.FileInfo
            Dim dra As IO.FileInfo
            Dim i As Integer

            oclsPrincipal.GravaLogCTM("", "", "Iniciando leitura na pasta " & caminhoEntradaSURA)
            
            diar1 = di.GetFiles("*.xlsx")

            For i = 0 To diar1.Length() - 1
                dra = diar1.GetValue(i)

                ArquivoSS = dra.Name
                nomeArquivo = caminhoEntradaSURA & dra.Name

                'NTendencia - Retirando pontos no nome do arquivo
                Dim novoCaminho As String = RenomearArquivo(nomeArquivo)
                nomeArquivo = novoCaminho
                ArquivoSS = Path.GetFileName(nomeArquivo)
                'NTendencia

                motivoRecusa = ""

                oclsPrincipal.GravaLogCTM("", "", "Abrindo arquivo " & ArquivoSS)

                ''O processo de ler linha a linha da planilha e gravar no banco estava muito demorado
                ''Foi substituido pelo prooesso abaixo de gerar um CSV e importat via BULK INSERT
                'ProcessarArquivosSura(nomeArquivo, caminhoSaidaSURA)

                'Novo procwesso de importação da planilha do faturamento
                'Gera um arquivo CSV e importa via BULK INSERT
                TransformarArquivoSuraCsv(caminhoEntradaSURA, ArquivoSS)

                If oclsPrincipal.ignoraFaturamento <> "S" Then

                    oclsPrincipal.GravaLogCTM("", "", "Verificando o ramo do arquivo para o processamento!")

                    sSQL = "SELECT ramo_tb.tp_ramo_id "
                    sSQL = sSQL & "FROM seguros_db.dbo.sura_arquivo_importacao_tb sura_arquivo_importacao_tb WITH (NOLOCK) "
                    sSQL = sSQL & "INNER JOIN seguros_db.dbo.sura_apolice_capital_tb sura_apolice_capital_tb WITH (NOLOCK) "
                    sSQL = sSQL & "   ON sura_arquivo_importacao_tb.apolice_arq = sura_apolice_capital_tb.apolice_sura "
                    sSQL = sSQL & "  AND sura_arquivo_importacao_tb.cod_ramo_arq = sura_apolice_capital_tb.ramo_sura "
                    sSQL = sSQL & "  AND sura_arquivo_importacao_tb.premio_arq = sura_apolice_capital_tb.val_premio "
                    sSQL = sSQL & "INNER JOIN seguros_db.dbo.apolice_tb apolice_tb WITH (NOLOCK) "
                    sSQL = sSQL & "   ON sura_apolice_capital_tb.apolice_id = apolice_tb.apolice_id "
                    sSQL = sSQL & "  AND sura_apolice_capital_tb.ramo_id = apolice_tb.ramo_id "
                    sSQL = sSQL & "  AND sura_apolice_capital_tb.seguradora_cod_susep = apolice_tb.seguradora_cod_susep "
                    sSQL = sSQL & "  AND sura_apolice_capital_tb.sucursal_seguradora_id = apolice_tb.sucursal_seguradora_id "
                    sSQL = sSQL & "INNER JOIN seguros_db.dbo.ramo_tb ramo_tb WITH (NOLOCK) "
                    sSQL = sSQL & "   ON apolice_tb.ramo_id = ramo_tb.ramo_id "
                    sSQL = sSQL & "WHERE sura_arquivo_importacao_tb.nome_arquivo_importado = '" & ArquivoSS & "' "
                    sSQL = sSQL & "  AND sura_arquivo_importacao_tb.processado IS NULL "
                    sSQL = sSQL & "GROUP BY ramo_tb.tp_ramo_id "

                    dsConsulta = cCon.ExecuteDataset(CommandType.Text, sSQL)

                    If dsConsulta.Tables(0).Rows.Count = 0 Then
                        If (motivoRecusa.Length = 0) Then
                            motivoRecusa = "Arquivo sem registros de pagamentos ou Apólice não identificada no SEGBR - Importação cancelada!"
                            RejeitaArquivo()
                            oclsPrincipal.GravaLogCTM("", "", "Arquivo sem registros de pagamentos ou Apólice não identificada no SEGBR - Importação cancelada!")
                        End If
                    Else
                        tp_ramo_id = dsConsulta.Tables(0).Rows(0)("tp_ramo_id")
                    End If
                    dsConsulta.Dispose()

                    If motivoRecusa.Length > 0 Then
                        oclsPrincipal.GravaLogCTM("", "", "Limpando a base com as infornmações do arquivo com erro!")
                        LimparTabelaRegistrosArquivo(ArquivoSS, "1")
                        MoverArquivoLido(nomeArquivo, "Erros\")
                    Else
                        If tp_ramo_id = 1 Then
                            oclsPrincipal.GravaLogCTM("", "", "Gerando a planilha para importação do movimento da apólice de vida!")
                            GerarPlanilhasFaturamentoVida(caminhoSaidaSURA)

                            If motivoRecusa.Length > 0 Then
                                oclsPrincipal.GravaLogCTM("", "", "Falha no processo de geração da planilha de faturamento do arquivo " + ArquivoSS + "!")
                                LimparTabelaRegistrosArquivo(ArquivoSS, "1")
                                MoverArquivoLido(nomeArquivo, "Erros\")
                            Else
                                oclsPrincipal.GravaLogCTM("", "", "Arquivo " + ArquivoSS + " processado e gerada a planilha de faturamento!")
                                MoverArquivoLido(nomeArquivo, "Processados\")
                            End If
                        Else
                            LimparTabelaRegistrosArquivo(ArquivoSS, "2")

                            ProcessarAdesoesResidencial(ArquivoSS)

                            If motivoRecusa.Length = 0 Then
                                ProcessarFaturamentoResidencial(ArquivoSS)
                                'ProcessarProLaboreResidencial(ArquivoSS)
                                If motivoRecusa.Length > 0 Then
                                    LimparTabelaRegistrosArquivo(ArquivoSS, "3")
                                    MoverArquivoLido(nomeArquivo, "Erros\")
                                Else
                                    MoverArquivoLido(nomeArquivo, "Processados\")
                                End If
                            Else
                                LimparTabelaRegistrosArquivo(ArquivoSS, "3")
                                MoverArquivoLido(nomeArquivo, "Erros\")
                            End If
                        End If
                    End If
                Else
                    oclsPrincipal.GravaLogCTM("", "", "Arquivo foi lido apenas para composição de base, sem processamento!")
                    MoverArquivoLido(nomeArquivo, "Processados\")
                End If
                oclsPrincipal.GravaLogCTM("", "", "Finalizou tratamento do arquivo " + ArquivoSS + "!")
            Next

            oclsPrincipal.GravaLogCTM("OK", "SEGP1619", "Processo executado com sucesso: " & oclsPrincipal.codigoTarefa)

            'GERAR PLANILHA EXCEL COM OS ERROS PROCESSADOS
            If dsDetalheErros.Tables(0).Rows.Count > 0 Then
                nomeArquivoErros = "Registros_Importados_Erro_SURA_" & DateTime.Now.ToString("ddMMyyyyHHmmss")
                oclsPrincipal.GravaLogCTM("", "", "Gerando planilha com as inconsistencias apuradas nos arquivos!")
                Dim excel As New ExcelUtil
                excel.ExportarExcelErros(caminhoSaidaSURA + "Erros\", dsDetalheErros, nomeArquivoErros)
                oclsPrincipal.GravaLogCTM("", "", "Gerada a planilha " + nomeArquivoErros + " Na pasta " + caminhoSaidaSURA + "!")
            End If
        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            oclsPrincipal.GravaLogCTM("NOT OK", "SEGP1619", "Erro ocorrido no processamento da tarefa!")
        End Try
    End Sub

    Private Sub TransformarArquivoSuraCsv(ByVal caminhoEntradaSURA As String, ByVal ArquivoSS As String)
        Dim _excel As Application = New Application()
        Dim _excelworkBook As Workbook

        Dim nomeArquivo As String = caminhoEntradaSURA & ArquivoSS
        Dim nomeArquivoTempCsvPagos As String
        Dim nomeArquivoTempCsvEmitidos As String
        Dim bTemEmitidos As Boolean = False
        Dim bTemPagos As Boolean = False

        _excelworkBook = _excel.Workbooks.Open(nomeArquivo)

        Try
            nomeArquivoTempCsvPagos = caminhoEntradaSURA & "temp_Pagos_" & ArquivoSS.Replace("xlsx", "csv").Replace("xls", "csv")
            nomeArquivoTempCsvEmitidos = caminhoEntradaSURA & "temp_Emitidos_" & ArquivoSS.Replace("xlsx", "csv").Replace("xls", "csv")

            If File.Exists(nomeArquivoTempCsvPagos) Then
                File.Delete(nomeArquivoTempCsvPagos)
            End If

            If File.Exists(nomeArquivoTempCsvEmitidos) Then
                File.Delete(nomeArquivoTempCsvEmitidos)
            End If

            If Not _excelworkBook.Sheets("Emitidos") Is Nothing Then
                ' definir qual a planilha de trabalho
                _excelworkBook.Sheets("Emitidos").Select()

                oclsPrincipal.GravaLogCTM("", "", "Gerando arquivo csv temporário da aba Emitidos.")

                _excelworkBook.SaveAs(nomeArquivoTempCsvEmitidos, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV)

                bTemEmitidos = True

            End If

            If oclsPrincipal.ignoraFaturamento <> "S" Then
                If Not _excelworkBook.Sheets("Pagos") Is Nothing Then
                    ' definir qual a planilha de trabalho
                    _excelworkBook.Sheets("Pagos").Select()

                    oclsPrincipal.GravaLogCTM("", "", "Gerando arquivo csv da aba Pagos.")

                    _excelworkBook.SaveAs(nomeArquivoTempCsvPagos, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV)

                    bTemPagos = True
                Else
                    oclsPrincipal.GravaLogCTM("", "", "Aba Pagos não localizada no arquivo " & ArquivoSS)
                    motivoRecusa = "Aba Pagos não localizada no arquivo!"
                    RejeitaArquivo()
                End If
            End If

                ' Fechar a planilha sem salvar alterações (para salvar mude False para True)
                _excelworkBook.Close(False)
                _excel.Application.Quit()

                'Liberar arquivo
                FreeFile(nomeArquivo)

            ' Liberar a memória
                ReleaseObject(_excelworkBook)
                ReleaseObject(_excel)

            If bTemEmitidos Then
                'Grava na base os registros
                oclsPrincipal.GravaLogCTM("", "", "Gravando na base a aba Emitidos!")
                GravaRegistroSuraLoteEmitidos(ArquivoSS, nomeArquivoTempCsvEmitidos)
                'Remove o temporário
                File.Delete(nomeArquivoTempCsvEmitidos)
            End If

            If bTemPagos Then
                'Grava na base os registros
                oclsPrincipal.GravaLogCTM("", "", "Gravando na base a aba Pagos para o processamento!")
                GravaRegistroSuraLote(ArquivoSS, nomeArquivoTempCsvPagos)
                'Remove o temporário
                File.Delete(nomeArquivoTempCsvPagos)
            End If

        Catch ex As Exception
            'motivoRecusa = ex.Message
            'NTendencia - Otimizando retorno de erro
            oclsPrincipal.GravaLogCTM("", "", "Exceção ocorrida na leitura do arquivo!" + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message))
            motivoRecusa = "Exceção ocorrida na leitura do arquivo!" + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
        End Try

    End Sub

    Private Sub ProcessarArquivosSura(ByVal nomeArquivo As String, ByVal caminhoSaidaSURA As String)
        Dim sSQL As String
        Dim dsGrupoPorApolice As New System.Data.DataSet()
        Dim ds As DataSet
        Dim PropostaId As String

        Dim _excel As Application = New Application()
        Dim _excelworkBook As Workbook

        Dim dataHoje As DateTime = DateTime.Now
        Dim varLin, qtdeLin As Integer

        Try
            _excelworkBook = _excel.Workbooks.Open(nomeArquivo)

            If Not _excelworkBook.Sheets("Pagos") Is Nothing Then
                ' definir qual a planilha de trabalho
                _excelworkBook.Sheets("Pagos").Select()

                qtdeLin = _excelworkBook.Application.Cells(_excelworkBook.Application.Rows.Count, "A").End(-4162).row

                For varLin = 2 To qtdeLin
                    If Trim(_excelworkBook.Application.Cells(varLin, 4).Value) = "Pagamento" Then

                        If Trim(_excelworkBook.Application.Cells(varLin, 13).Value).Length = 0 Then 'Nome segurado
                            motivoRecusa = "Nome em branco - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If Not ValidaCPF(_excelworkBook.Application.Cells(varLin, 12).Value) Then 'CPF
                            motivoRecusa = "CPF inválido - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If IsDate(_excelworkBook.Application.Cells(varLin, 14).Value) = False Then 'Data de nascimento
                            motivoRecusa = "Data de nascimento inválida - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If IsDate(_excelworkBook.Application.Cells(varLin, 18).Value) = False Then 'Data de emissão
                            motivoRecusa = "Data de emissão inválida - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If IsDate(_excelworkBook.Application.Cells(varLin, 3).Value) = False Then 'Data de inclusao
                            motivoRecusa = "Data de inclusão inválida - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        'Rafael Soares - Nova Tendência - STRY0118573
                        If IsNumeric(_excelworkBook.Application.Cells(varLin, 17).Value) = False Then 'Prêmio
                            motivoRecusa = "Valor do prêmio inválido - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If Trim(_excelworkBook.Application.Cells(varLin, 30).Value).Length = 0 Then 'Bairro
                            motivoRecusa = "Nome do bairro em branco - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If IsNumeric(Trim(Replace(_excelworkBook.Application.Cells(varLin, 31).Value, "-", ""))) = False Or _excelworkBook.Application.Cells(varLin, 31).Value = "00000000" Then 'CEP
                            motivoRecusa = "Valor do CEP inválido - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If



                        sSQL = " SELECT ISNULL(sura_apolice_capital_tb.val_premio, 0) val_premio, "
                        sSQL = sSQL & " ISNULL(sura_apolice_capital_tb.val_capital_segurado, 0) val_capital_segurado, "
                        sSQL = sSQL & " ISNULL(sura_apolice_capital_tb.apolice_id, 0) apolice_id, "
                        sSQL = sSQL & " ISNULL(sura_apolice_capital_tb.ramo_id, 0) ramo_id "
                        sSQL = sSQL & " FROM seguros_db.dbo.sura_apolice_capital_tb sura_apolice_capital_tb WITH (NOLOCK)"
                        sSQL = sSQL & " WHERE sura_apolice_capital_tb.apolice_sura = " & Trim(_excelworkBook.Application.Cells(varLin, 8).Value) & " "
                        sSQL = sSQL & "   AND sura_apolice_capital_tb.ramo_sura = " & Trim(_excelworkBook.Application.Cells(varLin, 7).Value) & " "

                        ds = cCon.ExecuteDataset(CommandType.Text, sSQL)

                        If ds.Tables(0).Rows.Count = 0 Then
                            ds.Dispose()
                            motivoRecusa = "Valor do prêmio ou capital segurado não encontrado para a apolice da SURA - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If Not (ds.Tables(0).Rows(0)("val_capital_segurado") > 0) And (ds.Tables(0).Rows(0)("val_premio") > 0) Then
                            ds.Dispose()
                            motivoRecusa = "Valor do prêmio ou capital segurado não encontrado para a apolice da SURA - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        sSQL = " SELECT proposta_tb.situacao, "
                        sSQL = sSQL & " proposta_tb.proposta_id "
                        sSQL = sSQL & " FROM seguros_db.dbo.apolice_tb apolice_tb WITH (NOLOCK)"
                        sSQL = sSQL & "INNER JOIN seguros_db.dbo.proposta_tb proposta_tb WITH (NOLOCK)"
                        sSQL = sSQL & "   ON apolice_tb.proposta_id = proposta_tb.proposta_id"
                        sSQL = sSQL & " WHERE apolice_tb.apolice_id = " & ds.Tables(0).Rows(0)("apolice_id") & " "
                        sSQL = sSQL & "   AND apolice_tb.ramo_id = " & ds.Tables(0).Rows(0)("ramo_id") & " "

                        ds.Dispose()

                        ds = cCon.ExecuteDataset(CommandType.Text, sSQL)

                        If ds.Tables(0).Rows.Count = 0 Then
                            ds.Dispose()
                            motivoRecusa = "Apólice não vinculada e/ou inexistente - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If ds.Tables(0).Rows(0)("situacao").ToString.ToUpper = "C" Then
                            ds.Dispose()
                            motivoRecusa = "Apólice vinculada cancelada! - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        If ds.Tables(0).Rows(0)("situacao").ToString.ToUpper <> "I" And _
                           ds.Tables(0).Rows(0)("situacao").ToString.ToUpper <> "A" Then
                            ds.Dispose()
                            motivoRecusa = "Apólice vinculada pendente! - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        PropostaId = ds.Tables(0).Rows(0)("proposta_id").ToString

                        ds.Dispose()

                        sSQL = " SELECT proposta_suspensa_tb.proposta_id "
                        sSQL = sSQL & " FROM seguros_db.dbo.proposta_suspensa_tb proposta_suspensa_tb WITH (NOLOCK)"
                        sSQL = sSQL & " WHERE proposta_suspensa_tb.proposta_id = " & PropostaId & " "
                        sSQL = sSQL & "   AND proposta_suspensa_tb.dt_inicio_suspensao <= '" & DateTime.Parse(Trim(_excelworkBook.Application.Cells(varLin, 20).Value)).ToString("yyyyMMdd") & "' "
                        sSQL = sSQL & "   AND (proposta_suspensa_tb.dt_fim_suspensao IS NULL "
                        sSQL = sSQL & "    OR (proposta_suspensa_tb.dt_fim_suspensao IS NOT NULL AND proposta_suspensa_tb.cod_suspensao = 1 )) "

                        ds = cCon.ExecuteDataset(CommandType.Text, sSQL)

                        If ds.Tables(0).Rows.Count > 0 Then
                            ds.Dispose()
                            motivoRecusa = "Apolice vinculada suspensa - linha " + varLin.ToString + " - Item " + Trim(_excelworkBook.Application.Cells(varLin, 1).Value)
                            RejeitaArquivo()
                            Exit Sub
                        End If

                        ds.Dispose()

                        'Validar informação SEXO

                        GravaRegistroSURA(Trim(_excelworkBook.Application.Cells(varLin, 1).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 2).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 3).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 4).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 5).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 6).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 7).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 8).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 9).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 10).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 11).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 12).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 13).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 14).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 15).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 16).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 17).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 18).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 19).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 20).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 21).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 22).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 23).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 24).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 25).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 26).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 27).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 28).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 29).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 30).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 31).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 32).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 33).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 34).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 35).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 36).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 37).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 38).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 39).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 40).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 41).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 42).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 43).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 44).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 45).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 46).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 47).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 48).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 49).Value), _
                        Trim(_excelworkBook.Application.Cells(varLin, 50).Value), _
                        ArquivoSS)
                    End If

                Next varLin
            Else
                motivoRecusa = "Aba Pagos não localizada no arquivo!"
                RejeitaArquivo()
                Exit Sub
            End If

        Catch ex As Exception
            'motivoRecusa = ex.Message
            'NTendencia - Otimizando retorno de erro
            motivoRecusa = "Exceção ocorrida na leitura do arquivo!" + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
        Finally
            ' Fechar a planilha sem salvar alterações (para salvar mude False para True)
            _excelworkBook.Close(False)
            _excel.Application.Quit()

            'Liberar arquivo
            FreeFile(nomeArquivo)

            ' Liberar a memória
            ReleaseObject(_excelworkBook)
            ReleaseObject(_excel)
        End Try

    End Sub
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub FreeFile(ByVal filePath As String)
        ' Aqui, você deve saber qual processo está usando o arquivo
        Dim processName As String = "EXCEL" ' Exemplo: "EXCEL"

        Try
            ' Obtém todos os processos pelo nome
            Dim processes() As Process = Process.GetProcessesByName(processName)

            ' Itera pelos processos e força o término
            For Each proc As Process In processes
                If Not proc.HasExited Then
                    proc.Kill() ' Força o término do processo
                    proc.WaitForExit() ' Aguarda o processo terminar
                End If
            Next

        Catch ex As Exception
            'NTendencia - Otimizando retorno de erro
            Throw New Exception("Erro ao tentar terminar o processo: " & IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message))
        End Try

    End Sub
    Public Sub GerarPlanilhasFaturamentoVida(ByVal caminhoSaidaSURA As String)
        Dim sSQL As String
        Dim tbDetalhe As New System.Data.DataTable()
        Dim dsDetalhePlanilha As New DataSet()
        Dim dsGrupoPorApolice As New DataSet()

        Dim dataHoje As DateTime = DateTime.Now

        oclsPrincipal.GravaLogCTM("", "", "Gerando a planilha para o faturamento!")
        'INÍCIO - SURA - Criar planilha dos pagantes para gerar o faturamento das apolices do vida.
        Try
            dsDetalhePlanilha.Tables.Add(tbDetalhe)
            dsDetalhePlanilha.Tables(0).Columns.Add("Sequencial", GetType(Integer))
            dsDetalhePlanilha.Tables(0).Columns.Add("tipo_registro", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("inicio_cobertura", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("entrada_subgrupo", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("nome_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("cpf_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("nascimento_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("sexo_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("estado_civil_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("capital_segurado", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("salario", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("endereco_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("bairro_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("municipio_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("cep_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("uf_principal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("email", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("ddd_telefone", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("telefone", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("nome_conjuge", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("cpf_conjuge", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("nascimento_conjuge", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("sexo_conjuge", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("inicio_cobertura_conjuge", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("beneficiario1", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("beneficiario2", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("beneficiario3", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("beneficiario4", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("beneficiario5", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("beneficiario6", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("beneficiario7", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("campo1", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("campo2", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("campo3", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("campo4", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("campo5", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("subgrupo", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("vlr_premio_mensal", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("extrato", GetType(String))
            dsDetalhePlanilha.Tables(0).Columns.Add("contrato", GetType(String))

            sSQL = "SELECT DISTINCT "
            sSQL = sSQL & " '' AS sequencial, "
            sSQL = sSQL & " 10 AS tipo_registro, "
            sSQL = sSQL & " CONVERT(VARCHAR(10), sura_arquivo_importacao_tb.dt_inicio_vigencia_arq, 103) AS inicio_cobertura, "
            sSQL = sSQL & " CONVERT(VARCHAR(10), sura_arquivo_importacao_tb.dt_inicio_vigencia_arq, 103) AS entrada_subgrupo, "
            sSQL = sSQL & " sura_arquivo_importacao_tb.nome_segurado_arq AS nome_principal, "
            sSQL = sSQL & " sura_arquivo_importacao_tb.CPF_CNPJ_segurado_arq AS cpf_principal, "
            sSQL = sSQL & " CONVERT(VARCHAR(10), sura_arquivo_importacao_tb.dt_nascimento_segurado_arq, 103) AS nascimento_principal, "
            sSQL = sSQL & " sura_arquivo_importacao_tb.sexo_arq AS sexo_principal, "
            sSQL = sSQL & " '' AS estado_civil_principal, "
            sSQL = sSQL & " REPLACE(STR(CONVERT(FLOAT, sura_apolice_capital_tb.val_capital_segurado), 10, 2), '.', ',') AS capital_segurado, "
            sSQL = sSQL & " '' AS salario, "
            sSQL = sSQL & " ISNULL(sura_arquivo_importacao_tb.logradouro_arq, '') + ',' + ISNULL(sura_arquivo_importacao_tb.numero_arq, '') + ' ' + ISNULL(sura_arquivo_importacao_tb.complemento_arq, '') AS endereco_principal, "
            sSQL = sSQL & " ISNULL(sura_arquivo_importacao_tb.bairro_arq, '') AS bairro_principal, "
            sSQL = sSQL & " sura_arquivo_importacao_tb.cidade_arq AS municipio_principal, "
            sSQL = sSQL & " sura_arquivo_importacao_tb.cep_arq AS cep_principal, "
            sSQL = sSQL & " sura_arquivo_importacao_tb.uf_arq AS uf_principal, "
            sSQL = sSQL & " '' AS email, "
            sSQL = sSQL & " '' AS ddd_telefone, "
            sSQL = sSQL & " '' AS telefone, "
            sSQL = sSQL & " '' AS nome_conjuge, "
            sSQL = sSQL & " '' AS cpf_conjuge, "
            sSQL = sSQL & " '' AS nascimento_conjuge, "
            sSQL = sSQL & " '' AS sexo_conjuge, "
            sSQL = sSQL & " '' AS inicio_cobertura_conjuge, "
            sSQL = sSQL & " '' AS beneficiario1, "
            sSQL = sSQL & " '' AS beneficiario2, "
            sSQL = sSQL & " '' AS beneficiario3, "
            sSQL = sSQL & " '' AS beneficiario4, "
            sSQL = sSQL & " '' AS beneficiario5, "
            sSQL = sSQL & " '' AS beneficiario6, "
            sSQL = sSQL & " '' AS beneficiario7, "
            sSQL = sSQL & " '' AS campo1, "
            sSQL = sSQL & " '' AS campo2, "
            sSQL = sSQL & " '' AS campo3, "
            sSQL = sSQL & " '' AS campo4, "
            sSQL = sSQL & " '' AS campo5, "
            sSQL = sSQL & " 1 AS subgrupo, "
            sSQL = sSQL & " REPLACE(STR(CONVERT(FLOAT, sura_apolice_capital_tb.val_premio), 10, 2), '.', ',') AS vlr_premio_mensal, "
            sSQL = sSQL & " sura_apolice_capital_tb.apolice_id, "
            sSQL = sSQL & " sura_apolice_capital_tb.ramo_id "
            sSQL = sSQL & "FROM seguros_db.dbo.sura_arquivo_importacao_tb sura_arquivo_importacao_tb WITH (NOLOCK) "
            sSQL = sSQL & "INNER JOIN seguros_db.dbo.sura_apolice_capital_tb sura_apolice_capital_tb WITH (NOLOCK) "
            sSQL = sSQL & "   ON sura_arquivo_importacao_tb.apolice_arq = sura_apolice_capital_tb.apolice_sura "
            sSQL = sSQL & "  AND sura_arquivo_importacao_tb.cod_ramo_arq = sura_apolice_capital_tb.ramo_sura "
            sSQL = sSQL & "  AND sura_arquivo_importacao_tb.premio_arq = sura_apolice_capital_tb.val_premio "
            sSQL = sSQL & "INNER JOIN seguros_db.dbo.apolice_tb apolice_tb WITH (NOLOCK) "
            sSQL = sSQL & "   ON sura_apolice_capital_tb.apolice_id = apolice_tb.apolice_id "
            sSQL = sSQL & "  AND sura_apolice_capital_tb.ramo_id = apolice_tb.ramo_id "
            sSQL = sSQL & "  AND sura_apolice_capital_tb.seguradora_cod_susep = apolice_tb.seguradora_cod_susep "
            sSQL = sSQL & "  AND sura_apolice_capital_tb.sucursal_seguradora_id = apolice_tb.sucursal_seguradora_id "
            sSQL = sSQL & "INNER JOIN seguros_db.dbo.proposta_tb proposta_tb WITH (NOLOCK) "
            sSQL = sSQL & "   ON apolice_tb.proposta_id = proposta_tb.proposta_id "
            sSQL = sSQL & "INNER JOIN seguros_db.dbo.ramo_tb ramo_tb WITH (NOLOCK) "
            sSQL = sSQL & "   ON proposta_tb.ramo_id = ramo_tb.ramo_id "
            sSQL = sSQL & "WHERE sura_arquivo_importacao_tb.processado IS NULL "
            sSQL = sSQL & "  AND sura_arquivo_importacao_tb.nome_arquivo_importado = '" & ArquivoSS & "' "

            oclsPrincipal.GravaLogCTM("", "", "Buscando as informações na base processadas!")

            dsDetalhePlanilha = cCon.ExecuteDataset(CommandType.Text, sSQL)

            'Valida se retornou algum registro para gerar a planilha de vida
            If dsDetalhePlanilha.Tables(0).Rows.Count > 0 Then
                sSQL = "SELECT sura_apolice_capital_tb.apolice_id, "
                sSQL = sSQL & "sura_apolice_capital_tb.ramo_id, "
                sSQL = sSQL & "cliente_tb.nome, "
                sSQL = sSQL & "CONVERT(VARCHAR(10), MAX(sura_arquivo_importacao_tb.dt_inicio_vigencia_arq), 103) AS inicio_cobertura "
                sSQL = sSQL & "FROM seguros_db.dbo.sura_arquivo_importacao_tb sura_arquivo_importacao_tb WITH (NOLOCK) "
                sSQL = sSQL & "INNER JOIN seguros_db.dbo.sura_apolice_capital_tb sura_apolice_capital_tb WITH (NOLOCK) "
                sSQL = sSQL & "   ON sura_arquivo_importacao_tb.apolice_arq = sura_apolice_capital_tb.apolice_sura "
                sSQL = sSQL & "  AND sura_arquivo_importacao_tb.cod_ramo_arq = sura_apolice_capital_tb.ramo_sura "
                sSQL = sSQL & "  AND sura_arquivo_importacao_tb.premio_arq = sura_apolice_capital_tb.val_premio "
                sSQL = sSQL & "INNER JOIN seguros_db.dbo.apolice_tb apolice_tb WITH (NOLOCK) "
                sSQL = sSQL & "   ON sura_apolice_capital_tb.apolice_id = apolice_tb.apolice_id "
                sSQL = sSQL & "  AND sura_apolice_capital_tb.ramo_id = apolice_tb.ramo_id "
                sSQL = sSQL & "  AND sura_apolice_capital_tb.seguradora_cod_susep = apolice_tb.seguradora_cod_susep "
                sSQL = sSQL & "  AND sura_apolice_capital_tb.sucursal_seguradora_id = apolice_tb.sucursal_seguradora_id "
                sSQL = sSQL & "INNER JOIN seguros_db.dbo.proposta_tb proposta_tb WITH (NOLOCK) "
                sSQL = sSQL & "   ON apolice_tb.proposta_id = proposta_tb.proposta_id "
                sSQL = sSQL & "INNER JOIN seguros_db.dbo.cliente_tb cliente_tb WITH (NOLOCK) "
                sSQL = sSQL & "   ON proposta_tb.prop_cliente_id = cliente_tb.cliente_id "
                sSQL = sSQL & "INNER JOIN seguros_db.dbo.ramo_tb ramo_tb WITH (NOLOCK) "
                sSQL = sSQL & "   ON proposta_tb.ramo_id = ramo_tb.ramo_id "
                sSQL = sSQL & "WHERE sura_arquivo_importacao_tb.processado IS NULL "
                sSQL = sSQL & "  AND sura_arquivo_importacao_tb.nome_arquivo_importado = '" & ArquivoSS & "' "
                sSQL = sSQL & "GROUP BY sura_apolice_capital_tb.apolice_id, "
                sSQL = sSQL & "         sura_apolice_capital_tb.ramo_id, "
                sSQL = sSQL & "         cliente_tb.nome "

                dsGrupoPorApolice = cCon.ExecuteDataset(CommandType.Text, sSQL)

                'GERAR PLANILHA EXCEL DO FATURAMENTO
                If dsDetalhePlanilha.Tables(0).Rows.Count > 0 Then
                    oclsPrincipal.GravaLogCTM("", "", "Trabalhando na geração da planilha de faturamento!")
                    Dim nomeArquivo As String
                    Dim excel As New ExcelUtil
                    nomeArquivo = Path.GetFileNameWithoutExtension(ArquivoSS)
                    excel.ExportarExcelFaturamento(caminhoSaidaSURA, dsDetalhePlanilha, nomeArquivo, dsGrupoPorApolice)
                End If
                'FIM - SURA - Criar planilha dos pagantes para gerar o faturamento.

                AtualizarRegistrosProcessados()

                dsGrupoPorApolice.Dispose()

            End If

            dsDetalhePlanilha.Dispose()

        Catch ex As Exception
            'NTendencia - Otimizando retorno de erro
            'motivoRecusa = "Falha ao gerar a planilha para o faturamento - " + ex.Message
            oclsPrincipal.GravaLogCTM("", "", "Falha ao gerar a planilha para o faturamento - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message))
            motivoRecusa = "Falha ao gerar a planilha para o faturamento - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
        End Try
    End Sub
    Private Sub GravaRegistroSuraLote(ByVal nome_arquivo_original As String, _
                                      ByVal nome_arquivo As String)
        Dim dsResultado As DataSet
        Dim sSQL As String
        Dim iQtde As Integer

        Try
            sSQL = "EXEC seguros_db.dbo.SEGS16088_SPI "

            sSQL = sSQL & "@nome_arquivo_original  = '" & nome_arquivo_original & "', "
            sSQL = sSQL & "@caminho_arquivo  = '" & nome_arquivo & "', "
            sSQL = sSQL & "@usuario = '" & oclsPrincipal.nomeUsuario & "' "

            dsResultado = cCon.ExecuteDataset(CommandType.Text, sSQL)

            For Each registro As DataRow In dsResultado.Tables(0).Rows
                iQtde = Convert.ToInt32(registro("Qtde").ToString())
                If iQtde = 0 Then
                    motivoRecusa = registro("mensagem_recusa").ToString() + " - Item " + registro("item_recusa").ToString()
                    RejeitaArquivo()
                End If
            Next

            dsResultado.Dispose()

        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            motivoRecusa = "Exceção ocorrida ao inserir registros na base!" + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            oclsPrincipal.GravaLogCTM("", "", "Exceção ocorrida ao inserir registros na base!" + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message))
            RejeitaArquivo()
        End Try

    End Sub
    Private Sub GravaRegistroSuraLoteEmitidos(ByVal nome_arquivo_original As String, _
                                 ByVal nome_arquivo As String)

        Dim sSQL As String

        Try
            sSQL = "EXEC seguros_db.dbo.SEGS16335_SPI "
            sSQL = sSQL & "@nome_arquivo_original = '" & nome_arquivo_original & "', "
            sSQL = sSQL & "@caminho_arquivo  = '" & nome_arquivo & "' ,"
            sSQL = sSQL & "@usuario  = '" & oclsPrincipal.nomeUsuario & "' "

            cCon.ExecuteNonQuery(CommandType.Text, sSQL)

        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            oclsPrincipal.GravaLogCTM("", "", "Erro ao compor base de registros no arquivo!" + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message))
        End Try

    End Sub

    Private Sub GravaRegistroSURA(ByVal cod_item_arq As String, _
                                  ByVal cod_controle_arq As String, _
                                  ByVal dt_inclusao_arq As String, _
                                  ByVal operacao_arq As String, _
                                  ByVal status_arq As String, _
                                  ByVal op_arq As String, _
                                  ByVal cod_ramo_arq As String, _
                                  ByVal apolice_arq As String, _
                                  ByVal movimento_arq As String, _
                                  ByVal certificado_parceiro_arq As String, _
                                  ByVal CNPJ_estipulante_arq As String, _
                                  ByVal CPF_CNPJ_segurado_arq As String, _
                                  ByVal nome_segurado_arq As String, _
                                  ByVal dt_nascimento_segurado_arq As String, _
                                  ByVal cod_produto_arq As String, _
                                  ByVal cod_plano_arq As String, _
                                  ByVal premio_arq As String, _
                                  ByVal dt_emissao_arq As String, _
                                  ByVal motivo_canc_arq As String, _
                                  ByVal dt_inicio_vigencia_arq As String, _
                                  ByVal dt_fim_vigencia_arq As String, _
                                  ByVal dt_cancelamento_arq As String, _
                                  ByVal dt_inicio_vigencia_cancelamento_arq As String, _
                                  ByVal valor_cancelamento_arq As String, _
                                  ByVal logradouro_arq As String, _
                                  ByVal numero_arq As String, _
                                  ByVal complemento_arq As String, _
                                  ByVal uf_arq As String, _
                                  ByVal cidade_arq As String, _
                                  ByVal bairro_arq As String, _
                                  ByVal cep_arq As String, _
                                  ByVal dt_nota_fiscal_arq As String, _
                                  ByVal numero_nota_fiscal_arq As String, _
                                  ByVal pz_garantia_fabricante_arq As String, _
                                  ByVal vl_equipamento_arq As String, _
                                  ByVal tp_equipamento_arq As String, _
                                  ByVal cd_equipamento_arq As String, _
                                  ByVal ds_equipamento_arq As String, _
                                  ByVal cd_vendedor_arq As String, _
                                  ByVal cd_loja_arq As String, _
                                  ByVal nu_cnpj_loja_arq As String, _
                                  ByVal vl_faixa_arq As String, _
                                  ByVal ds_marca_arq As String, _
                                  ByVal ds_modelo_arq As String, _
                                  ByVal nu_serie_equipamento_arq As String, _
                                  ByVal placa_veículo_arq As String, _
                                  ByVal desc_mensagem_arq As String, _
                                  ByVal nome_arquivo_arq As String, _
                                  ByVal linha_arq As String, _
                                  ByVal sexo_arq As String, _
                                  ByVal nome_arquivo_importado As String)

        Dim sSQL As String
        Dim dataHoje As DateTime = DateTime.Now
        Dim data As String

        Try
            sSQL = "EXEC seguros_db.dbo.SEGS16088_SPI "
            If cod_item_arq.Length > 0 Then
                sSQL = sSQL & "@cod_item_arq = '" & TratarCamposRecebidos(cod_item_arq) & "', "
            Else
                sSQL = sSQL & "@cod_item_arq = NULL, "
            End If

            If cod_controle_arq.Length > 0 Then
                sSQL = sSQL & "@cod_controle_arq = '" & TratarCamposRecebidos(cod_controle_arq) & "', "
            Else
                sSQL = sSQL & "@cod_controle_arq = NULL, "
            End If

            If dt_inclusao_arq.Length > 0 Then
                data = TratarCamposRecebidos(dt_inclusao_arq)
                sSQL = sSQL & "@dt_inclusao_arq = '" & DateTime.Parse(data).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            Else
                sSQL = sSQL & "@dt_inclusao_arq = NULL, "
            End If

            If operacao_arq.Length > 0 Then
                sSQL = sSQL & "@operacao_arq = '" & TratarCamposRecebidos(operacao_arq) & "', "
            Else
                sSQL = sSQL & "@operacao_arq = NULL, "
            End If

            If status_arq.Length > 0 Then
                sSQL = sSQL & "@status_arq = '" & TratarCamposRecebidos(status_arq) & "', "
            Else
                sSQL = sSQL & "@status_arq = NULL, "
            End If

            If op_arq.Length > 0 Then
                sSQL = sSQL & "@op_arq = '" & TratarCamposRecebidos(op_arq) & "', "
            Else
                sSQL = sSQL & "@op_arq = NULL, "
            End If

            If cod_ramo_arq.Length > 0 Then
                sSQL = sSQL & "@cod_ramo_arq = '" & TratarCamposRecebidos(cod_ramo_arq) & "', "
            Else
                sSQL = sSQL & "@cod_ramo_arq = NULL, "
            End If

            If apolice_arq.Length > 0 Then
                sSQL = sSQL & "@apolice_arq = '" & TratarCamposRecebidos(apolice_arq) & "', "
            Else
                sSQL = sSQL & "@apolice_arq = NULL, "
            End If

            If movimento_arq.Length > 0 Then
                sSQL = sSQL & "@movimento_arq = '" & TratarCamposRecebidos(movimento_arq) & "', "
            Else
                sSQL = sSQL & "@movimento_arq = NULL, "
            End If

            If certificado_parceiro_arq.Length > 0 Then
                sSQL = sSQL & "@certificado_parceiro_arq = '" & TratarCamposRecebidos(certificado_parceiro_arq) & "', "
            Else
                sSQL = sSQL & "@certificado_parceiro_arq = NULL, "
            End If

            If CNPJ_estipulante_arq.Length > 0 Then
                sSQL = sSQL & "@CNPJ_estipulante_arq = '" & TratarCamposRecebidos(CNPJ_estipulante_arq) & "', "
            Else
                sSQL = sSQL & "@CNPJ_estipulante_arq = NULL, "
            End If

            If CPF_CNPJ_segurado_arq.Length > 0 Then
                sSQL = sSQL & "@CPF_CNPJ_segurado_arq = '" & TratarCamposRecebidos(CPF_CNPJ_segurado_arq.PadLeft(11, "0"c)) & "', "
            Else
                sSQL = sSQL & "@CPF_CNPJ_segurado_arq = NULL, "
            End If

            If nome_segurado_arq.Length > 0 Then
                sSQL = sSQL & "@nome_segurado_arq = '" & TratarCamposRecebidos(nome_segurado_arq) & "', "
            Else
                sSQL = sSQL & "@nome_segurado_arq = NULL, "
            End If

            If dt_nascimento_segurado_arq.Length > 0 Then
                data = TratarCamposRecebidos(dt_nascimento_segurado_arq)
                sSQL = sSQL & "@dt_nascimento_segurado_arq = '" & DateTime.Parse(data).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            Else
                sSQL = sSQL & "@dt_nascimento_segurado_arq = NULL, "
            End If

            If cod_produto_arq.Length > 0 Then
                sSQL = sSQL & "@cod_produto_arq = '" & TratarCamposRecebidos(cod_produto_arq) & "', "
            Else
                sSQL = sSQL & "@cod_produto_arq = NULL, "
            End If

            If cod_plano_arq.Length > 0 Then
                sSQL = sSQL & "@cod_plano_arq = '" & TratarCamposRecebidos(cod_plano_arq) & "', "
            Else
                sSQL = sSQL & "@cod_plano_arq = NULL, "
            End If

            If premio_arq.Length > 0 Then
                sSQL = sSQL & "@premio_arq = '" & TratarCamposRecebidos(premio_arq.Replace(",", ".")) & "', "
            Else
                sSQL = sSQL & "@premio_arq = NULL, "
            End If

            If dt_emissao_arq.Length > 0 Then
                data = TratarCamposRecebidos(dt_emissao_arq)
                sSQL = sSQL & "@dt_emissao_arq = '" & DateTime.Parse(data).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            Else
                sSQL = sSQL & "@dt_emissao_arq = NULL, "
            End If

            If motivo_canc_arq.Length > 0 Then
                sSQL = sSQL & "@motivo_canc_arq = '" & TratarCamposRecebidos(motivo_canc_arq) & "', "
            Else
                sSQL = sSQL & "@motivo_canc_arq = NULL, "
            End If

            If dt_inicio_vigencia_arq.Length > 0 Then
                data = TratarCamposRecebidos(dt_inicio_vigencia_arq)
                sSQL = sSQL & "@dt_inicio_vigencia_arq = '" & DateTime.Parse(data).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            Else
                sSQL = sSQL & "@dt_inicio_vigencia_arq = NULL, "
            End If

            If dt_fim_vigencia_arq.Length > 0 Then
                data = TratarCamposRecebidos(dt_fim_vigencia_arq)
                sSQL = sSQL & "@dt_fim_vigencia_arq = '" & DateTime.Parse(data).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            Else
                sSQL = sSQL & "@dt_fim_vigencia_arq = NULL, "
            End If

            If dt_cancelamento_arq.Length > 0 Then
                data = TratarCamposRecebidos(dt_cancelamento_arq)
                sSQL = sSQL & "@dt_cancelamento_arq = '" & DateTime.Parse(data).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            Else
                sSQL = sSQL & "@dt_cancelamento_arq = NULL, "
            End If

            If dt_inicio_vigencia_cancelamento_arq.Length > 0 Then
                data = TratarCamposRecebidos(dt_inicio_vigencia_cancelamento_arq)
                sSQL = sSQL & "@dt_inicio_vigencia_cancelamento_arq = '" & DateTime.Parse(data).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            Else
                sSQL = sSQL & "@dt_inicio_vigencia_cancelamento_arq = NULL, "
            End If

            If valor_cancelamento_arq.Length > 0 Then
                sSQL = sSQL & "@valor_cancelamento_arq = '" & TratarCamposRecebidos(valor_cancelamento_arq) & "', "
            Else
                sSQL = sSQL & "@valor_cancelamento_arq = NULL, "
            End If

            If logradouro_arq.Length > 0 Then
                sSQL = sSQL & "@logradouro_arq = '" & TratarCamposRecebidos(logradouro_arq) & "', "
            Else
                sSQL = sSQL & "@logradouro_arq = NULL, "
            End If

            If numero_arq.Length > 0 Then
                sSQL = sSQL & "@numero_arq = '" & TratarCamposRecebidos(numero_arq) & "', "
            Else
                sSQL = sSQL & "@numero_arq = NULL, "
            End If

            If complemento_arq.Length > 0 Then
                sSQL = sSQL & "@complemento_arq = '" & TratarCamposRecebidos(complemento_arq) & "', "
            Else
                sSQL = sSQL & "@complemento_arq = NULL, "
            End If

            If uf_arq.Length > 0 Then
                sSQL = sSQL & "@uf_arq = '" & TratarCamposRecebidos(uf_arq) & "', "
            Else
                sSQL = sSQL & "@uf_arq = NULL, "
            End If

            If cidade_arq.Length > 0 Then
                sSQL = sSQL & "@cidade_arq = '" & TratarCamposRecebidos(cidade_arq) & "', "
            Else
                sSQL = sSQL & "@cidade_arq = NULL, "
            End If

            If bairro_arq.Length > 0 Then
                sSQL = sSQL & "@bairro_arq = '" & TratarCamposRecebidos(bairro_arq) & "', "
            Else
                sSQL = sSQL & "@bairro_arq = NULL, "
            End If

            If cep_arq.Length > 0 Then
                sSQL = sSQL & "@cep_arq = '" & TratarCamposRecebidos(cep_arq.PadLeft(8, "0"c)) & "', "
            Else
                sSQL = sSQL & "@cep_arq = NULL, "
            End If

            If dt_nota_fiscal_arq.Length > 0 Then
                data = TratarCamposRecebidos(dt_nota_fiscal_arq)
                sSQL = sSQL & "@dt_nota_fiscal_arq = '" & DateTime.Parse(data).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            Else
                sSQL = sSQL & "@dt_nota_fiscal_arq = NULL, "
            End If

            If numero_nota_fiscal_arq.Length > 0 Then
                sSQL = sSQL & "@numero_nota_fiscal_arq = '" & TratarCamposRecebidos(numero_nota_fiscal_arq) & "', "
            Else
                sSQL = sSQL & "@numero_nota_fiscal_arq = NULL, "
            End If

            If pz_garantia_fabricante_arq.Length > 0 Then
                sSQL = sSQL & "@pz_garantia_fabricante_arq = '" & TratarCamposRecebidos(pz_garantia_fabricante_arq) & "', "
            Else
                sSQL = sSQL & "@pz_garantia_fabricante_arq = NULL, "
            End If

            If vl_equipamento_arq.Length > 0 Then
                sSQL = sSQL & "@vl_equipamento_arq = '" & TratarCamposRecebidos(vl_equipamento_arq.Replace(",", ".")) & "', "
            Else
                sSQL = sSQL & "@vl_equipamento_arq = NULL, "
            End If

            If tp_equipamento_arq.Length > 0 Then
                sSQL = sSQL & "@tp_equipamento_arq = '" & TratarCamposRecebidos(tp_equipamento_arq) & "', "
            Else
                sSQL = sSQL & "@tp_equipamento_arq = NULL, "
            End If

            If cd_equipamento_arq.Length > 0 Then
                sSQL = sSQL & "@cd_equipamento_arq = '" & TratarCamposRecebidos(cd_equipamento_arq) & "', "
            Else
                sSQL = sSQL & "@cd_equipamento_arq = NULL, "
            End If

            If ds_equipamento_arq.Length > 0 Then
                sSQL = sSQL & "@ds_equipamento_arq = '" & TratarCamposRecebidos(ds_equipamento_arq) & "', "
            Else
                sSQL = sSQL & "@ds_equipamento_arq = NULL, "
            End If

            If cd_vendedor_arq.Length > 0 Then
                sSQL = sSQL & "@cd_vendedor_arq = '" & TratarCamposRecebidos(cd_vendedor_arq) & "', "
            Else
                sSQL = sSQL & "@cd_vendedor_arq = NULL, "
            End If

            If cd_loja_arq.Length > 0 Then
                sSQL = sSQL & "@cd_loja_arq = '" & TratarCamposRecebidos(cd_loja_arq) & "', "
            Else
                sSQL = sSQL & "@cd_loja_arq = NULL, "
            End If

            If nu_cnpj_loja_arq.Length > 0 Then
                sSQL = sSQL & "@nu_cnpj_loja_arq = '" & TratarCamposRecebidos(nu_cnpj_loja_arq) & "', "
            Else
                sSQL = sSQL & "@nu_cnpj_loja_arq = NULL, "
            End If

            If vl_faixa_arq.Length > 0 Then
                sSQL = sSQL & "@vl_faixa_arq = '" & TratarCamposRecebidos(vl_faixa_arq.Replace(",", ".")) & "', "
            Else
                sSQL = sSQL & "@vl_faixa_arq = NULL, "
            End If

            If ds_marca_arq.Length > 0 Then
                sSQL = sSQL & "@ds_marca_arq = '" & TratarCamposRecebidos(ds_marca_arq) & "', "
            Else
                sSQL = sSQL & "@ds_marca_arq = NULL, "
            End If

            If ds_modelo_arq.Length > 0 Then
                sSQL = sSQL & "@ds_modelo_arq = '" & TratarCamposRecebidos(ds_modelo_arq) & "', "
            Else
                sSQL = sSQL & "@ds_modelo_arq = NULL, "
            End If

            If nu_serie_equipamento_arq.Length > 0 Then
                sSQL = sSQL & "@nu_serie_equipamento_arq = '" & TratarCamposRecebidos(nu_serie_equipamento_arq) & "', "
            Else
                sSQL = sSQL & "@nu_serie_equipamento_arq = NULL, "
            End If

            If placa_veículo_arq.Length > 0 Then
                sSQL = sSQL & "@placa_veículo_arq = '" & TratarCamposRecebidos(placa_veículo_arq) & "', "
            Else
                sSQL = sSQL & "@placa_veículo_arq = NULL, "
            End If

            If desc_mensagem_arq.Length > 0 Then
                sSQL = sSQL & "@desc_mensagem_arq = '" & TratarCamposRecebidos(desc_mensagem_arq) & "', "
            Else
                sSQL = sSQL & "@desc_mensagem_arq = NULL, "
            End If

            If nome_arquivo_arq.Length > 0 Then
                sSQL = sSQL & "@nome_arquivo_arq = '" & TratarCamposRecebidos(nome_arquivo_arq) & "', "
            Else
                sSQL = sSQL & "@nome_arquivo_arq = NULL, "
            End If

            If linha_arq.Length > 0 Then
                sSQL = sSQL & "@linha_arq = '" & TratarCamposRecebidos(linha_arq) & "', "
            Else
                sSQL = sSQL & "@linha_arq = NULL, "
            End If

            If sexo_arq.Length > 0 Then
                sSQL = sSQL & "@sexo_arq = '" & TratarCamposRecebidos(sexo_arq) & "', "
            Else
                sSQL = sSQL & "@sexo_arq = 'M', "
            End If

            sSQL = sSQL & "@nome_arquivo_importado  = '" & nome_arquivo_importado & "', "
            sSQL = sSQL & "@dt_inclusao  = '" & DateTime.Parse(dataHoje).ToString("yyyy-MM-dd HH:mm:ss") & "', "
            sSQL = sSQL & "@usuario = '" & oclsPrincipal.nomeUsuario & "' "

            cCon.ExecuteNonQuery(CommandType.Text, sSQL)

        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            'NTendencia - Otimizando retorno de erro
            'motivoRecusa = "Exceção ocorrida ao inserir registros na base!"
            motivoRecusa = "Exceção ocorrida ao inserir registros na base!" + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
        End Try

    End Sub

    Public Sub RejeitaArquivo()
        Dim rowLinha As DataRow
        Dim dataHoje As DateTime = DateTime.Now

        Try
            'Adicionando a linha atual no DataSet
            rowLinha = dsDetalheErros.Tables(0).NewRow
            rowLinha("Data") = dataHoje.ToString("dd/MM/yyyy")
            rowLinha("Arquivo") = ArquivoSS
            rowLinha("MotivoRecusa") = motivoRecusa

            dsDetalheErros.Tables(0).Rows.Add(rowLinha)
            dsDetalheErros.AcceptChanges()
        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
        End Try

    End Sub

    Public Sub AtualizarRegistrosProcessados()
        Dim sSQL As String

        Try
            'Atualizando os registros processados na tabela
            oclsPrincipal.GravaLogCTM("", "", "Executando a atualização dos registros processados!")
            sSQL = "EXEC seguros_db.dbo.SEGS16072_SPU "
            sSQL = sSQL & "   @nome_arquivo_importado = '" & ArquivoSS & "' "
            cCon.ExecuteNonQuery(CommandType.Text, sSQL)
        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            'NTendencia - Otimizando retorno de erro
            'motivoRecusa = "Falha ao atualizar a base processada - " + ex.Message
            motivoRecusa = "Falha ao atualizar a base processada - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
            oclsPrincipal.GravaLogCTM("", "", "Erro ao atualizar os registros ´processados na base!")
        End Try

    End Sub

    Public Sub LimparTabelaRegistrosArquivo(ByVal nomeArquivo As String, ByVal tabela As String)
        Dim sSQL As String

        Try
            'Apagando registros do arquivo recusado
            sSQL = "EXEC seguros_db.dbo.SEGS16084_SPD "
            sSQL = sSQL & "   @nome_arquivo_importado = '" & nomeArquivo & "', "
            sSQL = sSQL & "   @tabela = " & tabela
            cCon.ExecuteNonQuery(CommandType.Text, sSQL)
        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            'NTendencia - Otimizando retorno de erro
            'motivoRecusa = "Falha ao apagar a base do arquivo recusado - " + ex.Message
            motivoRecusa = "Falha ao apagar a base do arquivo recusado - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
        End Try

    End Sub

    Public Sub MoverArquivoLido(ByVal nomeArquivo As String, ByVal Destino As String)
        Dim caminhoMovimentacao As String

        Try
            caminhoMovimentacao = nomeArquivo.Substring(0, nomeArquivo.Length - ArquivoSS.Length) + Destino
            'Mover o arquivo lido
            My.Computer.FileSystem.MoveFile(nomeArquivo, caminhoMovimentacao & ArquivoSS, True)
        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            'NTendencia - Otimizando retorno de erro
            'motivoRecusa = "Falha mover o arquivo recusado - " + ex.Message
            motivoRecusa = "Falha mover o arquivo recusado - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
        End Try

    End Sub

    Public Sub ProcessarAdesoesResidencial(ByVal nomeArquivo As String)
        Dim sSQL As String
        Dim ds As DataSet
        Dim transacao As Integer

        Try
            oclsPrincipal.GravaLogCTM("", "", "Gerando as novas adesões e vinculando as adesões já efetuadas!")
            'Registrando as novas adesões
            sSQL = "EXEC seguros_db.dbo.SEGS16135_SPI "
            sSQL = sSQL & "   @nome_arquivo_importado = '" & nomeArquivo & "', "
            sSQL = sSQL & "   @usuario = '" & oclsPrincipal.nomeUsuario & "' "

            transacao = cCon.BeginTransaction()
            ds = cCon.ExecuteDataset(CommandType.Text, sSQL)

            If ds.Tables(0).Rows(0)("Total") = 0 Then
                motivoRecusa = "Falha ao registrar novas adesoes"
                RejeitaArquivo()
                cCon.RollBackTransaction(transacao)
                oclsPrincipal.GravaLogCTM("", "", "Falha ao processas adesões do arquivo " + nomeArquivo + "!")
                Exit Sub
            End If

            cCon.CommitTransaction(transacao)

        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            cCon.RollBackTransaction(transacao)
            'NTendencia - Otimizando retorno de erro
            'motivoRecusa = "Falha ao registrar as novas adesões do arquivo processado - " + ex.Message
            motivoRecusa = "Falha ao registrar as novas adesões do arquivo processado - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
            oclsPrincipal.GravaLogCTM("", "", "Falha ao registrar as novas adesões do arquivo processado - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message))
        End Try
    End Sub
    Public Sub ProcessarFaturamentoResidencial(ByVal nomeArquivo As String)
        Dim sSQL As String
        Dim ds As DataSet

        Try
            'Registrando o faturamento dos pagamentos recebidos
            oclsPrincipal.GravaLogCTM("", "", "Processando o faturamento da planilha importada!")
            sSQL = "EXEC seguros_db.dbo.SEGS16137_SPI " 'NTendencia - Nova procedure baseada na SEGS16075_SPI 'Vilela - Para testes
            sSQL = sSQL & "   @nome_arquivo_importado = '" & nomeArquivo & "', "
            sSQL = sSQL & "   @usuario = '" & oclsPrincipal.nomeUsuario & "' "

            ds = cCon.ExecuteDataset(CommandType.Text, sSQL)

            If ds.Tables(0).Rows(0)("qtde") = 0 Then
                motivoRecusa = "Falha ao gerar o faturmemto dos pagamentos recebidos!"
                RejeitaArquivo()
                oclsPrincipal.GravaLogCTM("", "", "Falha ao gerar o faturamemto dos pagamentos recebidos da planilha " + nomeArquivo + "!")
                Exit Sub
            End If

        Catch ex As Exception
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            'NTendencia - Otimizando retorno de erro
            'motivoRecusa = "Falha ao registrar o faturamento do arquivo processado - " + ex.Message
            motivoRecusa = "Falha ao registrar o faturamento do arquivo processado - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
            oclsPrincipal.GravaLogCTM("", "", "Falha ao registrar o faturamento do arquivo processado - " + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message))
        End Try

    End Sub

    Private Sub ProcessarProLaboreResidencial(ByVal nomeArquivo As String)
        Dim dsPercentual As DataSet
        Dim dsPropostas As DataSet
        Dim ds As DataSet
        Dim sSQL As String
        Dim transacao As Integer
        Dim proposta_id As String
        Dim num_cobranca As String
        Dim perc_pro_labore As Double

        Try

            'Rafael Soares - Nova - 11/04/2025
            sSQL = " SELECT TOP 1 ap.perc_pro_labore "
            sSQL = sSQL & "  FROM seguros_db.dbo.sura_arquivo_importacao_tb sai WITH(NOLOCK)"
            sSQL = sSQL & " INNER JOIN seguros_db.dbo.sura_apolice_capital_tb sac WITH(NOLOCK)"
            sSQL = sSQL & "    ON sai.apolice_arq = sac.apolice_sura "
            sSQL = sSQL & "   AND sai.cod_ramo_arq = sac.ramo_sura "
            sSQL = sSQL & " INNER JOIN seguros_db.dbo.apolice_perc_tb ap WITH(NOLOCK)"
            sSQL = sSQL & "    ON ap.apolice_id = sac.apolice_id "
            sSQL = sSQL & "   AND ap.ramo_id = sac.ramo_id "
            sSQL = sSQL & "   AND ap.sucursal_seguradora_id = sac.sucursal_seguradora_id "
            sSQL = sSQL & "   AND ap.seguradora_cod_susep = sac.seguradora_cod_susep "
            sSQL = sSQL & " WHERE sai.nome_arquivo_importado = '" & nomeArquivo & "'"
            dsPercentual = cCon.ExecuteDataset(CommandType.Text, sSQL)

            If dsPercentual.Tables(0).Rows.Count = 0 Then
                motivoRecusa = "Falha ao apurar o percentual de pró labora da apólice!"
                RejeitaArquivo()
                Exit Sub
            Else
                perc_pro_labore = dsPercentual.Tables(0).Rows(0)("perc_pro_labore").ToString()
            End If
            'Fim


            sSQL = "SELECT DISTINCT"
            sSQL = sSQL & " agendamento_cobranca_atual_tb.proposta_id, agendamento_cobranca_atual_tb.num_cobranca "
            sSQL = sSQL & " FROM seguros_db.dbo.fatura_tb fatura_tb WITH (NOLOCK) "
            sSQL = sSQL & " INNER JOIN seguros_db.dbo.agendamento_cobranca_atual_tb agendamento_cobranca_atual_tb WITH (NOLOCK) "
            sSQL = sSQL & " ON fatura_tb.proposta_id = agendamento_cobranca_atual_tb.proposta_id  "
            sSQL = sSQL & " AND fatura_tb.endosso_id = agendamento_cobranca_atual_tb.num_endosso "
            sSQL = sSQL & " INNER JOIN seguros_db.dbo.sura_apolice_capital_tb sura_apolice_capital_tb WITH (NOLOCK) "
            sSQL = sSQL & " ON fatura_tb.apolice_id = sura_apolice_capital_tb.apolice_id "
            sSQL = sSQL & " AND fatura_tb.ramo_id = sura_apolice_capital_tb.ramo_id "
            sSQL = sSQL & " INNER JOIN seguros_db.dbo.sura_arquivo_importacao_tb sura_arquivo_importacao_tb WITH (NOLOCK)"
            sSQL = sSQL & " ON sura_apolice_capital_tb.apolice_sura = sura_arquivo_importacao_tb.apolice_arq"
            'NTendencia - STRY0124515 - [SEGBR][Equatorial] Faturamento Residencial (124) - Adequação Competência e Vencimento da Fatura
            sSQL = sSQL & " AND fatura_tb.fatura_id = sura_arquivo_importacao_tb.fatura_id"
            'sSQL = sSQL & " AND fatura_tb.dt_inicio_vigencia = sura_arquivo_importacao_tb.dt_inicio_vigencia_arq"
            'STRY0124515 - [SEGBR][Equatorial] Faturamento Residencial (124) - Adequação Competência e Vencimento da Fatura
            sSQL = sSQL & " WHERE nome_arquivo_importado = '" & nomeArquivo & "'"
            sSQL = sSQL & "   AND fatura_tb.fatura_id = (SELECT MAX(f.fatura_id) "
            sSQL = sSQL & "                                FROM seguros_db.dbo.fatura_tb f WITH (NOLOCK) "
            sSQL = sSQL & "                               WHERE f.apolice_id = fatura_tb.apolice_id "
            sSQL = sSQL & "                    			    AND f.ramo_id = fatura_tb.ramo_id) "

            dsPropostas = cCon.ExecuteDataset(CommandType.Text, sSQL)

            If dsPropostas.Tables(0).Rows.Count = 0 Then
                motivoRecusa = "Falha ao apurar fatura emitida para calcular o valor do pró labore!"
                RejeitaArquivo()
                Exit Sub
            Else
                transacao = cCon.BeginTransaction()

                For Each registro As DataRow In dsPropostas.Tables(0).Rows
                    proposta_id = registro("proposta_id").ToString()
                    num_cobranca = registro("num_cobranca").ToString()


                    sSQL = ""
                    'Atualizando o pro labore
                    sSQL = "EXEC seguros_db.dbo.SEGS16138_SPI " 'NTendencia - Nova procedure para gerar pro labore - ASR 09/01/25
                    sSQL = sSQL & "   @proposta_id = '" & proposta_id & "', "
                    sSQL = sSQL & "   @num_cobranca = '" & num_cobranca & "', "
                    sSQL = sSQL & "   @perc_pro_labore = '" & perc_pro_labore & "', "
                    sSQL = sSQL & "   @usuario = '" & oclsPrincipal.nomeUsuario & "' "


                    ds = cCon.ExecuteDataset(CommandType.Text, sSQL)

                    ds.Dispose()
                Next
                cCon.CommitTransaction(transacao)
            End If

            dsPercentual.Dispose()
            dsPropostas.Dispose()
        Catch ex As Exception
            cCon.RollBackTransaction(transacao)
            oclsPrincipal.GravaLogCTM(ex.Message.ToString, _
                System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            'NTendencia - Otimizando retorno de erro
            'motivoRecusa = "Falha ao calcular o pró labore do arquivo processado!"
            motivoRecusa = "Falha ao calcular o pró labore do arquivo processado!" + IIf(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
            RejeitaArquivo()
        End Try
    End Sub

    Private Sub ProcessarBaixasRecebidas()
        Dim sSQL As String
        Dim ds As DataSet
        Dim transacao As Integer

        oclsPrincipal.GravaLogCTM("", "", "Iniciando processamento das baixas recebidas!")

        sSQL = "select cod_convenio_cob, txt_nom_arquivo, usuario, dt_alteracao,COD_MOTIVO_COMP "
        sSQL = sSQL & " FROM interface_dados_db.dbo.retorno_baixa_agendamento_erp_atual_tb retorno WITH(NOLOCK)  "
        sSQL = sSQL & " where(cod_produto = 124) "
        sSQL = sSQL & " AND retorno.cod_convenio_cob = 'CBR643BAB' "
        sSQL = sSQL & " AND mca_situacao_vcr = 0 "
        sSQL = sSQL & " AND COD_MOTIVO_COMP NOT IN (57,58) "
        sSQL = sSQL & " and txt_nom_arquivo = ''"

        ds = cCon.ExecuteDataset(CommandType.Text, sSQL)

        If ds.Tables(0).Rows.Count = 0 Then
            oclsPrincipal.GravaLogCTM("", "", "não foi identificada nenhuma baixa pendente!")
            ds.Dispose()
            Exit Sub
        Else
            ds.Dispose()
        End If

        Try
            oclsPrincipal.GravaLogCTM("", "", "Executando a procedure das baixas!")
            sSQL = "EXEC seguros_db.dbo.SEGS16154_SPI '" & oclsPrincipal.nomeUsuario & "' "

            transacao = cCon.BeginTransaction()

            cCon.ExecuteNonQuery(CommandType.Text, sSQL)

            cCon.CommitTransaction(transacao)

            oclsPrincipal.GravaLogCTM("OK", "SEGP1619", "Processo das baixas executados com sucesso! Tarefa " + oclsPrincipal.codigoTarefa + " finalizada!")

            Exit Sub
        Catch ex As Exception
            cCon.RollBackTransaction(transacao)
            oclsPrincipal.GravaLogCTM("Falha ao processar as baixas recebidas!", _
                    System.Reflection.MethodBase.GetCurrentMethod.Name().ToString, ex.StackTrace)
            oclsPrincipal.GravaLogCTM("NOT OK", "SEGP1619", "Falha ao processar as baixas recebidas!")
        End Try
    End Sub

    Public Function RenomearArquivo(ByVal caminhoArquivo As String) As String
        oclsPrincipal.GravaLogCTM("", "", "Retirando pontos do nome do arquivo")
        ' Verifica se o arquivo existe
        If Not File.Exists(caminhoArquivo) Then
            Throw New FileNotFoundException("Arquivo não encontrado: " & caminhoArquivo)
            oclsPrincipal.GravaLogCTM("", "", "Arquivo não encontrado: " & caminhoArquivo & "!")
        End If

        ' Obtém o diretório do arquivo
        Dim diretorio As String = Path.GetDirectoryName(caminhoArquivo)

        ' Obtém apenas o nome do arquivo com extensão
        Dim nomeCompleto As String = Path.GetFileName(caminhoArquivo)

        ' Separa nome e extensão
        Dim extensao As String = Path.GetExtension(nomeCompleto)
        Dim nomeSemExtensao As String = Path.GetFileNameWithoutExtension(nomeCompleto)

        ' Remove todos os pontos do nome (menos o da extensão)
        nomeSemExtensao = nomeSemExtensao.Replace(".", "")

        ' Monta o novo nome e caminho completo
        Dim novoNome As String = nomeSemExtensao & extensao
        Dim novoCaminho As String = Path.Combine(diretorio, novoNome)

        ' Renomeia o arquivo (move para o novo nome)
        File.Move(caminhoArquivo, novoCaminho)

        ' Retorna o novo caminho completo
        Return novoCaminho
    End Function

    Public Function TratarCamposRecebidos(ByVal campo As String) As String
        Dim campoRetorno As String

        'campoRetorno = campo.Replace("'", "''").Replace(";", "").Replace("--", "").Replace("/*", "").Replace("*/", "").Replace("xp_", "")
        campoRetorno = campo.Replace("'", "").Replace(";", "").Replace("--", "").Replace("/*", "").Replace("*/", "").Replace("xp_", "")

        Return campoRetorno

    End Function
End Class
