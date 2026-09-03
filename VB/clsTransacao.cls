VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mNumeroCartao As String
Private mValorTransacao As Double
Private mDataTransacao As Date
Private mDescricao As String
Private mStatusTransacao As String

'Para editar as transações
Private mIdTransacao As Long

Public Property Get idTransacao() As Long
    idTransacao = mIdTransacao
End Property

Public Property Let idTransacao(ByVal valor As Long)
    mIdTransacao = valor
End Property


Public Property Get numeroCartao() As String
    numeroCartao = mNumeroCartao
End Property

Public Property Let numeroCartao(ByVal valor As String)
    mNumeroCartao = valor
End Property


Public Property Get ValorTransacao() As Double
    ValorTransacao = mValorTransacao
End Property

Public Property Let ValorTransacao(ByVal valor As Double)
    mValorTransacao = valor
End Property


Public Property Get DataTransacao() As Date
    DataTransacao = mDataTransacao
End Property

Public Property Let DataTransacao(ByVal valor As Date)
    mDataTransacao = valor
End Property


Public Property Get Descricao() As String
    Descricao = mDescricao
End Property

Public Property Let Descricao(ByVal valor As String)
    mDescricao = valor
End Property


Public Property Get StatusTransacao() As String
    StatusTransacao = mStatusTransacao
End Property

Public Property Let StatusTransacao(ByVal valor As String)
    mStatusTransacao = valor
End Property


