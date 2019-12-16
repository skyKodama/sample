
Option Explicit On
Option Strict On
Imports Npgsql
Imports skysystem.common.SystemUtil


'*************************************************************************
'*　機能　　：DatabaseConnection
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
Public Class DBconnectionN

#Region "Private変数"
    Private Con As New NpgsqlConnection             'オラクルコネクション
#End Region

#Region "コンストラクタ"
    Public Sub New()

    End Sub
#End Region



#Region "RtnCon：接続を返す"
    ''' <summary>
    ''' 接続情報を戻す
    ''' </summary>
    ''' <returns>Sqlコネクション</returns>
    ''' <remarks></remarks>
    Public Function rtncon() As NpgsqlConnection

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
        If Not OraCon Is Nothing Then
            If OraCon.State = ConnectionState.Open Then
                OraCon.ClearPool()
                OraCon.Close()
            End If
        End If
    End Sub
#End Region



End Class
