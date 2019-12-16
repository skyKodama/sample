#Region "宣言"
Option Explicit On
Option Strict On
Imports System.Data.SqlClient
Imports skysystem.common.SystemUtil
#End Region

'*************************************************************************
'*　機能　　：DBconnection
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
Public Class DBconnectionSS

    Private Const DecryptPass As String = "#2030" '20101208_1

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

            Dim strDataSorce As String = PB_ReadXML("/SKY/SKY_DB/DataSorce", "", SystemConst.C_SYSTEMPRM)
            Dim strCatalog As String = PB_ReadXML("/SKY/SKY_DB/Catalog", "", SystemConst.C_SYSTEMPRM)
            Dim strUserId As String = PB_ReadXML("/SKY/SKY_DB/UserId", "", SystemConst.C_SYSTEMPRM)
            Dim strPassword As String = PB_ReadXML("/SKY/SKY_DB/Password", "", SystemConst.C_SYSTEMPRM)
            Dim strTimeOut As String = PB_ReadXML("/SKY/SKY_DB/TimeOut", "", SystemConst.C_SYSTEMPRM)

            Dim Str As String = CreateConnectionString(strDataSorce, strCatalog, strUserId, strPassword, strTimeOut)
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

#Region "CreateConnectionString:ConectionStringの生成"
    ''20101213_1
    Private Function CreateConnectionString(ByVal strDataSorce As String, ByVal strCatalog As String, ByVal strUserId As String, ByVal strPassword As String, ByVal strTimeOut As String) As String
        ''暗号化されている部分は複合化する。
        Dim rtn As String = ""
        rtn += "Data Source=" + strDataSorce & ";"
        rtn += "Initial Catalog=" + SystemUtil.doDecrypt(strCatalog, DecryptPass) & ";"
        rtn += "User ID=" + SystemUtil.doDecrypt(strUserId, DecryptPass) & ";"
        rtn += "Password=" + SystemUtil.doDecrypt(strPassword, DecryptPass) & ";"
        rtn += "Connection Lifetime=" + strTimeOut & ";"
        Return rtn
    End Function
#End Region


End Class
