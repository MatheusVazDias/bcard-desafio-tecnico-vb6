VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "XYZ Administradora de Cartões de Crédito"
   ClientHeight    =   6375
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4939.098
   ScaleMode       =   0  'User
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "20 Últimas Transações"
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
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid grdUltimasTransacoes 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   9340
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.Label Label1 
         Caption         =   "O GRID ACIMA NÃO É MANIPULÁVEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   5760
         Width           =   3255
      End
   End
   Begin VB.Menu mnuNovaTransacao 
      Caption         =   "Nova Transação"
      Index           =   1
   End
   Begin VB.Menu mnuConsultaTransacoes 
      Caption         =   "Consultar Transações"
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "Relatórios"
      Index           =   3
      Begin VB.Menu mnuRelMesAnterior 
         Caption         =   "Relatório de Transações do Mês Anterior"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'Conexão com o Banco de dados Utilizando a DAO
    On Error GoTo TrataErro

    Call AbrirConexao

'Configurando o tamanho das colunas e linhas
    Call ConfigurarGrid
    
'Popula FlexGrid com as últimas 20 transações
    Call CarregarUltimasTransacoes

    Exit Sub

TrataErro:

    RegistrarLog _
        "[frmPrincipal.Form_Load] Erro " & _
        CStr(Err.Number) & " - " & Err.Description

    MsgBox "Não foi possível conectar ao banco de dados." & vbCrLf & vbCrLf & _
           "O erro foi registrado no log do sistema.", _
           vbCritical, _
           "Erro de conexão"

    End



End Sub

Private Sub CarregarUltimasTransacoes()

    Dim rs As ADODB.Recordset
    Dim linha As Integer

    On Error GoTo TrataErro

    Set rs = BuscaUltimasTransacoes()

    With grdUltimasTransacoes

        .Clear

        .Rows = 1
        .Cols = 6

        'Cabeçalho
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Cartão"
        .TextMatrix(0, 2) = "Valor"
        .TextMatrix(0, 3) = "Data"
        .TextMatrix(0, 4) = "Descrição"
        .TextMatrix(0, 5) = "Status"

        linha = 1

        Do While Not rs.EOF

            .Rows = .Rows + 1

            .TextMatrix(linha, 0) = CStr(rs!Id_Transacao)

            .TextMatrix(linha, 1) = CStr(rs!Numero_Cartao)

            .TextMatrix(linha, 2) = FormatCurrency(rs!Valor_Transacao)

            .TextMatrix(linha, 3) = _
                Format(rs!Data_Transacao, "dd/mm/yyyy hh:nn")

            If IsNull(rs!Descricao) Then
                .TextMatrix(linha, 4) = ""
            Else
                .TextMatrix(linha, 4) = CStr(rs!Descricao)
            End If

            .TextMatrix(linha, 5) = CStr(rs!Status_Transacao)

            linha = linha + 1
            rs.MoveNext

        Loop

    End With

    rs.Close
    Set rs = Nothing

    Exit Sub

TrataErro:

    RegistrarLog _
        "[frmPrincipal.CarregarUltimasTransacoes] Erro " & _
        CStr(Err.Number) & " - " & Err.Description

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing

    MsgBox "Não foi possível carregar as transações." & vbCrLf & vbCrLf & _
           "O erro foi registrado no log do sistema.", _
           vbCritical, _
           "Erro"

End Sub

Private Sub ConfigurarGrid()

    With grdUltimasTransacoes

        .Rows = 2
        .Cols = 6
        .FixedRows = 1

        .ColWidth(0) = 1000    ' ID
        .ColWidth(1) = 2000    ' Cartão
        .ColWidth(2) = 1500    ' Valor
        .ColWidth(3) = 2500    ' Data
        .ColWidth(4) = 2500    ' Descrição
        .ColWidth(5) = 1000    ' Status

    End With

End Sub

Private Sub mnuTestaConexao_Click()

End Sub

Private Sub mnuConsultaTransacoes_Click()
    frmConsulta.Show vbModal
End Sub

Private Sub mnuNovaTransacao_Click(Index As Integer)

    frmTransacao.Show vbModal

End Sub

Private Sub mnuRelMesAnterior_Click()

    On Error GoTo TrataErro

    GerarRelatorioMesAnterior

    Exit Sub

TrataErro:

    RegistrarLog _
    "[frmPrincipal.mnuRelMesAnterior_Click] Erro " & _
    CStr(Err.Number) & " - " & Err.Description

    MsgBox "Não foi possível gerar o relatório." & vbCrLf & vbCrLf & _
           "O erro foi registrado no log do sistema.", _
           vbCritical, _
           "Erro"

End Sub


