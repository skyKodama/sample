#Region "宣言"
Option Explicit On
Option Strict On
Imports System.Data.SqlClient
Imports skysystem.common.SystemUtil
#End Region

'*************************************************************************
'*　機能　　：DatabaseConnectionServer（メインサーバーへの接続）
'*            <DB接続>
'*　作成日　：
'*
'*　＜変更内容＞
'*
'*************************************************************************
''' <summary>
''' データベースコネクションクラス
''' </summary>
''' <remarks>コネクションの接続・切断・戻しﾒｿｯﾄﾞ</remarks>
Public Class DBconnectionServer

#Region "Private変数"
    Private OraCon As New SqlConnection             'オラクルコネクション
#End Region

#Region "コンストラクタ"
    Public Sub New()

    End Sub
#End Region

#Region "ChkConnection：OracleOpen接続：チェック"
    ''' <summary>
    ''' 接続オープン処理
    ''' </summary>
    ''' <remarks>コネクションが閉じられている場合は、再接続を行う。</remarks>
    Public Sub Open()

        'コネクションが閉じられている場合のみ
        If OraCon Is Nothing OrElse (OraCon.State = ConnectionState.Closed) Then
            Dim Str As String = XMLReadConnectionS()
            OraCon.ConnectionString = Str
            OraCon.Open()
        End If


    End Sub
#End Region

#Region "RtnCon：接続を返す"
    ''' <summary>
    ''' 接続情報を戻す
    ''' </summary>
    ''' <returns>Sqlコネクション</returns>
    ''' <remarks></remarks>
    Public Function rtncon() As SqlConnection

        'コネクションが閉じられている場合のみ
        Open()

        Return OraCon
    End Function
#End Region

#Region "Close：接続を閉じる"
    ''' <summary>
    ''' 接続を閉じる
    ''' </summary>
    Public Sub Close()
        OraCon.Close()
    End Sub
#End Region

#Region "XMLReadConnection：XML接続文字列取得(CONNECTION)"
    ''' <summary>
    ''' XMLより接続文字列を取得する
    ''' </summary>
    ''' <returns>接続文字列</returns>
    ''' <remarks></remarks>
    Private Function XMLReadConnectionS() As String
        Return PB_ReadXML("/SKY/SKY_DB/CONNECTIONS", "", SystemConst.C_SYSTEMPRM)
    End Function
#End Region



End Class
