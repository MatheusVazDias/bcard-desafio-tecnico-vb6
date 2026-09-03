VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consultar Transações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   11415
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   615
         Left            =   9720
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdProxima 
         Caption         =   "Próxima"
         Height          =   495
         Left            =   5880
         TabIndex        =   10
         Top             =   8760
         Width           =   1575
      End
      Begin VB.TextBox txtValorMaximo 
         Height          =   375
         Left            =   9720
         TabIndex        =   5
         Text            =   "99999.99"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtValorMinimo 
         Height          =   375
         Left            =   6960
         TabIndex        =   4
         Text            =   "0"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtDataFinal 
         Height          =   375
         Left            =   9720
         TabIndex        =   2
         Text            =   "01/01/2070"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtCartao 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtDataInicial 
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Text            =   "01/01/1900"
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "Limpar"
         Height          =   615
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1320
         Width           =   3855
      End
      Begin MSFlexGridLib.MSFlexGrid grdTransacoes 
         Height          =   6255
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   11033
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.Label lblPagina 
         Caption         =   "Página ***"
         Height          =   375
         Left            =   9360
         TabIndex        =   19
         Top             =   8760
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Max.:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   8880
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Min.:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   16
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Data Final:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cartão:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Data Inicial:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Status:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Paginacao do Grid
Private paginaAtual As Long
Private Const registrosPorPag As Long = 10

Private Sub cmdAnterior_Click()
    If paginaAtual <= 1 Then
        Exit Sub
    End If

    paginaAtual = paginaAtual - 1

    ConsultarPagina

End Sub

Private Sub cmdExcluir_Click()
    
    Dim idTransacao As Long
    Dim excluiu As Boolean
    Dim resposta As VbMsgBoxResult

    On Error GoTo TrataErro

    'Ignora o cabeçalho
    If grdTransacoes.Row <= 0 Then

        MsgBox "Selecione uma transação para excluir.", _
               vbExclamation, _
               "Atenção"

        Exit Sub

    End If

    idTransacao = CLng(grdTransacoes.TextMatrix(grdTransacoes.Row, 0))

    resposta = MsgBox( _
        "Deseja realmente excluir a transação " & CStr(idTransacao) & "?", _
        vbQuestion + vbYesNo, _
        "Confirmar exclusão")

    If resposta <> vbYes Then
        Exit Sub
    End If

    'Forçando erro pra cair no log
    'Err.Raise 9999, , "Erro"
    
    excluiu = ExcluirTransacao(idTransacao)
    

    If Not excluiu Then

        MsgBox "A transação não foi encontrada ou já foi excluída.", _
               vbExclamation, _
               "Exclusão não realizada"

        Exit Sub

    End If

    MsgBox "Transação excluída com sucesso!", _
           vbInformation, _
           "Sucesso"

    ConsultarPagina

    Exit Sub

TrataErro:

    RegistrarLog _
        "[frmConsulta.cmdExcluir_Click] Erro ao excluir transação ID=" & _
        CStr(idTransacao) & _
        ". Erro " & CStr(Err.Number) & _
        " - " & Err.Description

    MsgBox "Não foi possível excluir a transação." & vbCrLf & vbCrLf & _
           "O erro foi registrado no log do sistema.", _
           vbCritical, _
           "Erro"

End Sub

Private Sub cmdLimpar_Click(Index As Integer)

    txtCartao.Text = ""
    txtDataInicial.Text = "01/01/1900"
    txtDataFinal.Text = "01/01/2070"
    txtValorMinimo.Text = "0"
    txtValorMaximo.Text = "99999.99"
    cboStatus.ListIndex = 0
End Sub

Private Sub cmdProxima_Click()

    If Not cmdProxima.Enabled Then
        Exit Sub
    End If

    paginaAtual = paginaAtual + 1

    ConsultarPagina

End Sub

Private Sub Form_Load()

    With grdTransacoes

        .Cols = 6
        .Rows = 2
        .FixedRows = 1

        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Cartão"
        .TextMatrix(0, 2) = "Valor"
        .TextMatrix(0, 3) = "Data"
        .TextMatrix(0, 4) = "Descrição"
        .TextMatrix(0, 5) = "Status"
        
        .ColWidth(0) = 1000    ' ID
        .ColWidth(1) = 2000    ' Cartão
        .ColWidth(2) = 1500    ' Valor
        .ColWidth(3) = 2500    ' Data
        .ColWidth(4) = 2500    ' Descrição
        .ColWidth(5) = 1000    ' Status
    End With
    
    cboStatus.AddItem "SELECIONE"
    cboStatus.AddItem "Aprovada"
    cboStatus.AddItem "Pendente"
    cboStatus.AddItem "Cancelada"

    cboStatus.ListIndex = 0
    
    paginaAtual = 1

    cmdAnterior.Enabled = False
    cmdProxima.Enabled = True

End Sub
Private Sub cmdConsultar_Click()

    If Not ValidarFiltros() Then
        Exit Sub
    End If

    paginaAtual = 1

    ConsultarPagina

End Sub

Private Sub ConsultarPagina()

    Dim rs As ADODB.Recordset
    Dim temProxima As Boolean

    Set rs = ConsultarTransacoes( _
        Trim(txtCartao.Text), _
        ObterDataInicial(), _
        ObterDataFinal(), _
        ObterValorMinimo(), _
        ObterValorMaximo(), _
        Trim(cboStatus.Text), _
        paginaAtual, _
        registrosPorPag)

    temProxima = CarregarGrid(rs)

    rs.Close
    Set rs = Nothing

    lblPagina.Caption = "Página " & CStr(paginaAtual)

    cmdAnterior.Enabled = (paginaAtual > 1)
    cmdProxima.Enabled = temProxima

End Sub
Private Function ObterDataInicial() As Variant

    Dim valor As String
    Dim data As Date

    valor = Trim(txtDataInicial.Text)

    If valor = "" Then
        ObterDataInicial = Null
        Exit Function
    End If

    If Not IsDate(valor) Then

        MsgBox "Informe uma data inicial válida.", _
               vbExclamation, _
               "Atenção"

        txtDataInicial.SetFocus
        ObterDataInicial = Null
        Exit Function

    End If

    data = CDate(valor)

    'Se informou somente a data, considera o dia inteiro
    If InStr(valor, ":") = 0 Then
        ObterDataInicial = DateAdd("d", 1, DateValue(data))
    Else
        ObterDataInicial = data
    End If

End Function
Private Function ObterDataFinal() As Variant

    Dim valor As String
    Dim data As Date

    valor = Trim(txtDataFinal.Text)

    If valor = "" Then
        ObterDataFinal = Null
        Exit Function
    End If

    If Not IsDate(valor) Then

        MsgBox "Informe uma data final válida.", _
               vbExclamation, _
               "Atenção"

        txtDataFinal.SetFocus
        ObterDataFinal = Null
        Exit Function

    End If

    data = CDate(valor)

    'Se informou somente a data, considera o dia inteiro
    If InStr(valor, ":") = 0 Then
        ObterDataFinal = DateAdd("d", 1, DateValue(data))
    Else
        ObterDataFinal = data
    End If

End Function

Private Function ObterValorMinimo() As Variant

    Dim valor As String

    valor = Trim(txtValorMinimo.Text)

    If valor = "" Then
        ObterValorMinimo = Null
        Exit Function
    End If

    If Not IsNumeric(valor) Then

        MsgBox "Informe um valor mínimo válido.", _
               vbExclamation, _
               "Atenção"

        txtValorMinimo.SetFocus
        ObterValorMinimo = Null
        Exit Function

    End If

    If CDbl(valor) < 0 Then

        MsgBox "O valor mínimo não pode ser negativo.", _
               vbExclamation, _
               "Atenção"

        txtValorMinimo.SetFocus
        ObterValorMinimo = Null
        Exit Function

    End If

    ObterValorMinimo = CDbl(valor)

End Function
Private Function ObterValorMaximo() As Variant

    Dim valor As String

    valor = Trim(txtValorMaximo.Text)

    If valor = "" Then
        ObterValorMaximo = Null
        Exit Function
    End If

    If Not IsNumeric(valor) Then
        MsgBox "Informe um valor máximo válido.", vbExclamation, "Atenção"
        ObterValorMaximo = Null
        Exit Function
    End If

    If CDbl(valor) < 0 Then
        MsgBox "O valor máximo não pode ser negativo.", vbExclamation, "Atenção"
        ObterValorMaximo = Null
        Exit Function
    End If

    ObterValorMaximo = CDbl(valor)

End Function

Private Function CarregarGrid(ByVal rs As ADODB.Recordset) As Boolean

    Dim linha As Integer
    Dim quantidade As Long

    With grdTransacoes

        .Rows = 1

        linha = 1

        quantidade = 0

        Do While Not rs.EOF And quantidade < registrosPorPag
        
            .Rows = .Rows + 1
        
            .TextMatrix(linha, 0) = CStr(rs!Id_Transacao)
            .TextMatrix(linha, 1) = CStr(rs!Numero_Cartao)
            .TextMatrix(linha, 2) = FormatCurrency(rs!Valor_Transacao)
            .TextMatrix(linha, 3) = Format(rs!Data_Transacao, "dd/mm/yyyy hh:nn")
        
            If IsNull(rs!Descricao) Then
                .TextMatrix(linha, 4) = ""
            Else
                .TextMatrix(linha, 4) = CStr(rs!Descricao)
            End If
        
            .TextMatrix(linha, 5) = CStr(rs!Status_Transacao)
        
            linha = linha + 1
            quantidade = quantidade + 1
        
            rs.MoveNext
        
        Loop

    End With
    CarregarGrid = Not rs.EOF
End Function

Private Function ValidarCartao() As Boolean

    Dim valor As String

    valor = Trim(txtCartao.Text)

    'Campo vazio é permitido
    If valor = "" Then
        ValidarCartao = True
        Exit Function
    End If

    If Len(valor) <> 15 Then
        MsgBox "O número do cartão deve conter 16 dígitos.", _
               vbExclamation, _
               "Atenção"

        txtCartao.SetFocus
        ValidarCartao = False
        Exit Function
    End If

    If Not IsNumeric(valor) Then
        MsgBox "O número do cartão deve conter apenas dígitos numéricos.", _
               vbExclamation, _
               "Atenção"

        txtCartao.SetFocus
        ValidarCartao = False
        Exit Function
    End If

    ValidarCartao = True

End Function

Private Function ValidarFiltros() As Boolean

    Dim dataInicial As Variant
    Dim dataFinal As Variant
    Dim valorMinimo As Variant
    Dim valorMaximo As Variant

    'Valida cartão
    If Not ValidarCartao() Then
        Exit Function
    End If

    'Valida data inicial
    dataInicial = ObterDataInicial()

    If Trim(txtDataInicial.Text) <> "" And IsNull(dataInicial) Then
        Exit Function
    End If

    'Valida data final
    dataFinal = ObterDataFinal()

    If Trim(txtDataFinal.Text) <> "" And IsNull(dataFinal) Then
        Exit Function
    End If

    'Valida valor mínimo
    valorMinimo = ObterValorMinimo()

    If Trim(txtValorMinimo.Text) <> "" And IsNull(valorMinimo) Then
        Exit Function
    End If

    'Valida valor máximo
    valorMaximo = ObterValorMaximo()

    If Trim(txtValorMaximo.Text) <> "" And IsNull(valorMaximo) Then
        Exit Function
    End If

    'Valida relação entre as datas
    If Not IsNull(dataInicial) And Not IsNull(dataFinal) Then

        If dataInicial >= dataFinal Then

            MsgBox "A data inicial deve ser menor que a data final.", _
                   vbExclamation, _
                   "Atenção"

            txtDataInicial.SetFocus
            Exit Function

        End If

    End If

    'Valida relação entre os valores
    If Not IsNull(valorMinimo) And Not IsNull(valorMaximo) Then

        If valorMinimo > valorMaximo Then

            MsgBox "O valor mínimo não pode ser maior que o valor máximo.", _
                   vbExclamation, _
                   "Atenção"

            txtValorMinimo.SetFocus
            Exit Function

        End If

    End If

    ValidarFiltros = True

End Function

Private Sub grdTransacoes_DblClick()

    Dim idTransacao As Long

    'Ignora o cabeçalho
    If grdTransacoes.Row <= 0 Then
        Exit Sub
    End If

    idTransacao = CLng(grdTransacoes.TextMatrix(grdTransacoes.Row, 0))

    frmTransacao.CarregarTransacao idTransacao
    frmTransacao.Show vbModal

End Sub
