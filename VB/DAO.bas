Attribute VB_Name = "DAO"

Option Explicit

Private mConexao As ADODB.Connection


Public Sub AbrirConexao()

    If mConexao Is Nothing Then
        Set mConexao = New ADODB.Connection
    End If

    If mConexao.State = adStateClosed Then

        mConexao.ConnectionString = _
            "Provider=SQLOLEDB;" & _
            "Data Source=localhost\SQLEXPRESS;" & _
            "Initial Catalog=XYZCartoes;" & _
            "Integrated Security=SSPI;"

        mConexao.Open

    End If

End Sub

Public Function ObterConexao() As ADODB.Connection

    AbrirConexao

    Set ObterConexao = mConexao

End Function

Public Function BuscaUltimasTransacoes() As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT TOP 20 " & _
          "Id_Transacao, " & _
          "Numero_Cartao, " & _
          "Valor_Transacao, " & _
          "Data_Transacao, " & _
          "Descricao, " & _
          "Status_Transacao " & _
          "FROM dbo.Transacao " & _
          "ORDER BY Data_Transacao DESC, Id_Transacao DESC"

    Set rs = New ADODB.Recordset

    rs.Open sql, ObterConexao(), adOpenForwardOnly, adLockReadOnly

    Set BuscaUltimasTransacoes = rs

End Function

Public Sub InserirTransacao(ByVal transacao As clsTransacao)

    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = ObterConexao()
        .CommandType = adCmdText
        .CommandText = _
            "INSERT INTO dbo.Transacao " & _
            "(Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao) " & _
            "VALUES (?, ?, ?, ?, ?)"

        .Parameters.Append .CreateParameter("Numero_Cartao", adVarChar, adParamInput, 16, transacao.numeroCartao)
        .Parameters.Append .CreateParameter("Valor_Transacao", adDecimal, adParamInput)
        .Parameters("Valor_Transacao").Precision = 18
        .Parameters("Valor_Transacao").NumericScale = 2
        .Parameters("Valor_Transacao").Value = transacao.ValorTransacao

        .Parameters.Append .CreateParameter("Data_Transacao", adDBTimeStamp, adParamInput, , transacao.DataTransacao)

        .Parameters.Append .CreateParameter("Descricao", adVarChar, adParamInput, 255, transacao.Descricao)

        .Parameters.Append .CreateParameter("Status_Transacao", adVarChar, adParamInput, 10, transacao.StatusTransacao)

        .Execute
    End With

    Set cmd = Nothing

End Sub
Public Function ConsultarTransacoes( _
    ByVal numeroCartao As String, _
    ByVal dataInicial As Variant, _
    ByVal dataFinal As Variant, _
    ByVal valorMinimo As Variant, _
    ByVal valorMaximo As Variant, _
    ByVal status As String, _
    ByVal pagina As Long, _
    ByVal registrosPorPagina As Long) As ADODB.Recordset

    
    Dim registrosIgnorar As Long
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    registrosIgnorar = (pagina - 1) * registrosPorPagina
    
    sql = "SELECT " & _
      "Id_Transacao, " & _
      "Numero_Cartao, " & _
      "Valor_Transacao, " & _
      "Data_Transacao, " & _
      "Descricao, " & _
      "Status_Transacao " & _
      "FROM dbo.Transacao " & _
      "WHERE 1 = 1 "

    If Trim(numeroCartao) <> "" Then
        sql = sql & "AND Numero_Cartao = '" & Replace(numeroCartao, "'", "''") & "' "
    End If

    If Not IsNull(dataInicial) Then
        sql = sql & "AND Data_Transacao >= '" & Format(dataInicial, "yyyy-mm-dd hh:nn:ss") & "' "
    End If

    If Not IsNull(dataFinal) Then
        sql = sql & "AND Data_Transacao < '" & Format(dataFinal, "yyyy-mm-dd hh:nn:ss") & "' "
    End If

    If Not IsNull(valorMinimo) Then
        sql = sql & "AND Valor_Transacao >= " & Replace(CStr(valorMinimo), ",", ".") & " "
    End If

    If Not IsNull(valorMaximo) Then
        sql = sql & "AND Valor_Transacao <= " & Replace(CStr(valorMaximo), ",", ".") & " "
    End If

    If Trim(status) <> "" And Trim(status) <> "SELECIONE" Then
        sql = sql & "AND Status_Transacao = '" & Replace(status, "'", "''") & "' "
    End If

    sql = sql & _
      "ORDER BY Data_Transacao DESC, Id_Transacao DESC " & _
      "OFFSET " & CStr(registrosIgnorar) & " ROWS " & _
      "FETCH NEXT " & CStr(registrosPorPagina + 1) & " ROWS ONLY" 'Aqui é a paginação do grid

    Set rs = New ADODB.Recordset

    rs.Open sql, ObterConexao(), adOpenForwardOnly, adLockReadOnly

    Set ConsultarTransacoes = rs

End Function

Public Function ObterTransacaoPorId(ByVal idTransacao As Long) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT " & _
          "Id_Transacao, " & _
          "Numero_Cartao, " & _
          "Valor_Transacao, " & _
          "Data_Transacao, " & _
          "Descricao, " & _
          "Status_Transacao " & _
          "FROM dbo.Transacao " & _
          "WHERE Id_Transacao = " & CStr(idTransacao)

    Set rs = New ADODB.Recordset

    rs.Open sql, ObterConexao(), adOpenForwardOnly, adLockReadOnly

    Set ObterTransacaoPorId = rs

End Function

Public Function AtualizarTransacao(ByVal transacao As clsTransacao) As Boolean

    Dim sql As String
    Dim registrosAfetados As Long

    sql = "UPDATE dbo.Transacao SET " & _
          "Numero_Cartao = '" & _
          Replace(transacao.numeroCartao, "'", "''") & "', " & _
          "Valor_Transacao = " & _
          Replace(CStr(transacao.ValorTransacao), ",", ".") & ", " & _
          "Data_Transacao = '" & _
          Format(transacao.DataTransacao, "yyyy-mm-dd hh:nn:ss") & "', " & _
          "Descricao = '" & _
          Replace(transacao.Descricao, "'", "''") & "', " & _
          "Status_Transacao = '" & _
          Replace(transacao.StatusTransacao, "'", "''") & "' " & _
          "WHERE Id_Transacao = " & _
          CStr(transacao.idTransacao) & " " & _
          "AND Status_Transacao <> 'Aprovada'"

    ObterConexao().Execute sql, registrosAfetados

    AtualizarTransacao = (registrosAfetados > 0)

End Function

Public Function ExcluirTransacao(ByVal idTransacao As Long) As Boolean

    Dim sql As String
    Dim registrosAfetados As Long

    sql = "DELETE FROM dbo.Transacao " & _
          "WHERE Id_Transacao = " & CStr(idTransacao)

    ObterConexao().Execute sql, registrosAfetados

    ExcluirTransacao = (registrosAfetados > 0)

End Function


Public Sub RegistrarLog(ByVal mensagem As String)

    Dim caminhoLog As String
    Dim numeroArquivo As Integer

    If Right(App.Path, 1) = "\" Then
        caminhoLog = App.Path & "logs"
    Else
        caminhoLog = App.Path & "\logs"
    End If

    If Dir(caminhoLog, vbDirectory) = "" Then
        MkDir caminhoLog
    End If

    caminhoLog = caminhoLog & "\erros.log"

    numeroArquivo = FreeFile

    Open caminhoLog For Append As #numeroArquivo

    Print #numeroArquivo, _
        Format(Now, "yyyy-mm-dd hh:nn:ss") & _
        " - " & mensagem

    Close #numeroArquivo

End Sub

Public Function BuscarTransacoesUltimoMes() As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT " & _
          "Id_Transacao, " & _
          "Numero_Cartao, " & _
          "Valor_Transacao, " & _
          "Data_Transacao, " & _
          "Descricao, " & _
          "Status_Transacao, " & _
          "Categoria " & _
          "FROM dbo.vw_TransacoesFinanceiras " & _
          "WHERE Data_Transacao >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0) " & _
          "AND Data_Transacao < DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0) " & _
          "ORDER BY Data_Transacao ASC, Id_Transacao ASC"

    Set rs = New ADODB.Recordset

    rs.Open sql, ObterConexao(), adOpenForwardOnly, adLockReadOnly

    Set BuscarTransacoesUltimoMes = rs

End Function

Public Sub GerarRelatorioMesAnterior()

    Dim rs As ADODB.Recordset
    Dim excelApp As Object
    Dim pastaTrabalho As Object
    Dim planilha As Object

    Dim linha As Long
    Dim caminhoArquivo As String
    
    Dim nomeArquivo As String

    On Error GoTo TrataErro

    'Busca as transações do mês anterior
    Set rs = BuscarTransacoesUltimoMes()

    If rs.EOF Then

        MsgBox "Não existem transações no mês anterior para gerar o relatório.", _
               vbInformation, _
               "Relatório"

        rs.Close
        Set rs = Nothing

        Exit Sub

    End If

    'Cria o Excel
    Set excelApp = CreateObject("Excel.Application")

    Set pastaTrabalho = excelApp.Workbooks.Add

    Set planilha = pastaTrabalho.Worksheets(1)

    planilha.Name = "Transações"

    'Cabeçalho
    planilha.Cells(1, 1).Value = "ID"
    planilha.Cells(1, 2).Value = "Número do Cartão"
    planilha.Cells(1, 3).Value = "Valor"
    planilha.Cells(1, 4).Value = "Data"
    planilha.Cells(1, 5).Value = "Descrição"
    planilha.Cells(1, 6).Value = "Status"
    planilha.Cells(1, 7).Value = "Categoria"

    'Dados
    linha = 2

    Do While Not rs.EOF

        planilha.Cells(linha, 1).Value = rs!Id_Transacao
        planilha.Cells(linha, 2).Value = rs!Numero_Cartao
        planilha.Cells(linha, 3).Value = rs!Valor_Transacao
        planilha.Cells(linha, 4).Value = rs!Data_Transacao

        If IsNull(rs!Descricao) Then
            planilha.Cells(linha, 5).Value = ""
        Else
            planilha.Cells(linha, 5).Value = rs!Descricao
        End If

        planilha.Cells(linha, 6).Value = rs!Status_Transacao
        planilha.Cells(linha, 7).Value = rs!Categoria

        linha = linha + 1

        rs.MoveNext

    Loop

    'Formatação
    planilha.Rows(1).Font.Bold = True

    planilha.Columns(2).NumberFormat = "@"
    planilha.Columns(3).NumberFormat = "#,##0.00"
    planilha.Columns(4).NumberFormat = "dd/mm/yyyy hh:mm"

    planilha.Columns("A:G").AutoFit

    'Define a pasta dos relatórios
    Dim pastaRelatorios As String
    
    If Right(App.Path, 1) = "\" Then
        pastaRelatorios = App.Path & "Relatorios"
    Else
        pastaRelatorios = App.Path & "\Relatorios"
    End If
    
    'Cria a pasta caso ela não exista
    If Dir(pastaRelatorios, vbDirectory) = "" Then
        MkDir pastaRelatorios
    End If
    
    'Nome do arquivo sai com a data
    nomeArquivo = "Relatorio_Transacoes_" & _
                  Format(DateAdd("m", -1, Date), "mm-yyyy") & _
                  ".xlsx"
    
    caminhoArquivo = pastaRelatorios & "\" & nomeArquivo

    'Salva
    pastaTrabalho.SaveAs caminhoArquivo, 51

    'Fecha
    pastaTrabalho.Close False
    excelApp.Quit

    Set planilha = Nothing
    Set pastaTrabalho = Nothing
    Set excelApp = Nothing

    rs.Close
    Set rs = Nothing

    MsgBox "Relatório gerado com sucesso!" & vbCrLf & vbCrLf & _
           "Arquivo disponível em:" & vbCrLf & _
           caminhoArquivo, _
           vbInformation, _
           "Relatório"

    Exit Sub

TrataErro:

    On Error Resume Next

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    If Not pastaTrabalho Is Nothing Then
        pastaTrabalho.Close False
    End If

    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If

    Set planilha = Nothing
    Set pastaTrabalho = Nothing
    Set excelApp = Nothing
    Set rs = Nothing

    RegistrarLog _
        "Erro ao gerar relatório do mês anterior: " & _
        CStr(Err.Number) & " - " & Err.Description

    Err.Raise Err.Number, , Err.Description

End Sub

