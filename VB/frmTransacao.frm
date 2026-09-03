VERSION 5.00
Begin VB.Form frmTransacao 
   Caption         =   "Nova Transação"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescricao 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox txtValor 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Cadastrais"
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
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   3240
         Width           =   3855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   3240
         TabIndex        =   11
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   615
         Left            =   600
         TabIndex        =   10
         Top             =   7080
         Width           =   1575
      End
      Begin VB.TextBox txtData 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox txtCartao 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Status:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Descrição:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Data:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Valor:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cartão:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mModoEdicao As Boolean
Private mIdTransacao As Long

Private mTransacaoAprovada As Boolean




Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdSalvar_Click()

    Dim transacao As clsTransacao
    Dim valorData As String

    On Error GoTo TrataErro

    ' Validação do cartão
    If Trim(txtCartao.Text) = "" Then
        MsgBox "Informe o número do cartão.", vbExclamation, "Validação"
        txtCartao.SetFocus
        Exit Sub
    End If

    If Len(Trim(txtCartao.Text)) <> 16 Then
        MsgBox "O número do cartão deve possuir 16 dígitos.", vbExclamation, "Validação"
        txtCartao.SetFocus
        Exit Sub
    End If

    ' Validação do valor
    If Trim(txtValor.Text) = "" Then
        MsgBox "Informe o valor da transação.", vbExclamation, "Validação"
        txtValor.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtValor.Text) Then
        MsgBox "Informe um valor válido.", vbExclamation, "Validação"
        txtValor.SetFocus
        Exit Sub
    End If

    If CDbl(txtValor.Text) <= 0 Then
        MsgBox "O valor da transação deve ser maior que zero.", vbExclamation, "Validação"
        txtValor.SetFocus
        Exit Sub
    End If

    ' Validação da data
    valorData = Trim(txtData.Text)

    If InStr(valorData, ".") > 0 Then
        valorData = Left(valorData, InStr(valorData, ".") - 1)
    End If
    
    If Not IsDate(valorData) Then
        MsgBox "Informe uma data válida.", vbExclamation, "Validação"
        txtData.SetFocus
        Exit Sub
    End If

    ' Validação da descrição
    If Trim(txtDescricao.Text) = "" Then
        MsgBox "Informe uma Descrição.", vbExclamation, "Validação"
        txtDescricao.SetFocus
        Exit Sub
    End If

    If Len(txtDescricao.Text) > 255 Then
        MsgBox "A descrição pode conter até 255 caracteres.", vbExclamation, "Validação"
        txtDescricao.SetFocus
        Exit Sub
    End If

    ' Validação do status
    If Trim(cboStatus.Text) = "" Or Trim(cboStatus.Text) = "SELECIONE" Then
        MsgBox "Selecione o status da transação.", vbExclamation, "Validação"
        cboStatus.SetFocus
        Exit Sub
    End If

    ' Cria o objeto da transação
    Set transacao = New clsTransacao

    ' Preenche os dados da transação
    transacao.numeroCartao = Trim(txtCartao.Text)
    transacao.ValorTransacao = CDbl(txtValor.Text)
    transacao.DataTransacao = CDate(valorData)
    transacao.Descricao = Trim(txtDescricao.Text)
    transacao.StatusTransacao = Trim(cboStatus.Text)

    If mModoEdicao Then

        Dim atualizou As Boolean
    
        transacao.idTransacao = mIdTransacao
    
        atualizou = AtualizarTransacao(transacao)
    
        If Not atualizou Then
    
            MsgBox "A transação não pôde ser alterada." & vbCrLf & vbCrLf & _
                   "Ela pode ter sido aprovada por outro processo enquanto estava sendo editada.", _
                   vbExclamation, _
                   "Alteração não realizada"
    
            Exit Sub
    
        End If
    
        MsgBox "Transação alterada com sucesso!", _
               vbInformation, _
               "Sucesso"

    Else
        Call InserirTransacao(transacao)
    
        MsgBox "Transação salva com sucesso!", _
               vbInformation, _
               "Sucesso"
    End If

    Unload Me

    Exit Sub

TrataErro:

    RegistrarLog _
        "[frmTransacao.cmdSalvar_Click] Erro " & _
        CStr(Err.Number) & " - " & Err.Description

    MsgBox "Não foi possível salvar a transação." & vbCrLf & vbCrLf & _
           "O erro foi registrado no log do sistema.", _
           vbCritical, _
           "Erro"

End Sub

Private Sub Form_Load()

    cboStatus.AddItem "SELECIONE"
    cboStatus.AddItem "Aprovada"
    cboStatus.AddItem "Pendente"
    cboStatus.AddItem "Cancelada"

    If Not mModoEdicao Then

        txtData.Text = Format(Now, "dd/mm/yyyy hh:nn")
        cboStatus.ListIndex = 0

    End If

End Sub

Public Sub CarregarTransacao(ByVal idTransacao As Long)

    Dim rs As ADODB.Recordset

    On Error GoTo TrataErro

    mModoEdicao = True
    mIdTransacao = idTransacao

    Set rs = ObterTransacaoPorId(idTransacao)

    If rs.EOF Then

        MsgBox "A transação não foi encontrada.", _
               vbExclamation, _
               "Atenção"

        rs.Close
        Set rs = Nothing

        Exit Sub

    End If

    txtCartao.Text = CStr(rs!Numero_Cartao)
    txtValor.Text = CStr(rs!Valor_Transacao)
    txtData.Text = Format(rs!Data_Transacao, "dd/mm/yyyy hh:nn")

    If IsNull(rs!Descricao) Then
        txtDescricao.Text = ""
    Else
        txtDescricao.Text = CStr(rs!Descricao)
    End If

    cboStatus.Text = CStr(rs!Status_Transacao)
    
    mTransacaoAprovada = (CStr(rs!Status_Transacao) = "Aprovada")
    
    If mTransacaoAprovada Then
        txtCartao.Enabled = False
        txtValor.Enabled = False
        txtData.Enabled = False
        txtDescricao.Enabled = False
        cboStatus.Enabled = False
    
        cmdSalvar.Enabled = False
    End If

    rs.Close
    Set rs = Nothing

    Exit Sub

TrataErro:

    RegistrarLog _
        "[frmTransacao.CarregarTransacao] Erro " & _
        CStr(Err.Number) & " - " & Err.Description

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing

    MsgBox "Não foi possível carregar a transação." & vbCrLf & vbCrLf & _
           "O erro foi registrado no log do sistema.", _
           vbCritical, _
           "Erro"

End Sub

