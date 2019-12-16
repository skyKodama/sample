
Option Explicit On
Option Strict On

Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports skysystem.common.SystemUtil
Imports skysystem.common

'*************************************************************************
'*　機能　　：共通DatabasePBクラス
'*            <DB接続>
'*            <DB結果戻り>
'*　作成日　：2006.05.22    黄
'*
'*　＜変更内容＞
'*　20060619_1　駒形　SQL作成AtoZ(PBFSTR_CreatSqlAtoZ)
'*  20060721_1  駒形　レポート用データセット取得   
'*  20060728_1  駒形  データテーブルに最大行に満たない空レコードを追加する(AddRow)
'*  20061010_1  駒形  PBFSTR_SQLMltSgl(SQL構築(全角半角を区別しない))
'*  20061023_1  黄    ExecuteDB、更新行数戻り値追加
'*  20070523_1  駒形  DataVeiw取得
'*
'*************************************************************************
''' <summary>
''' データベース用ユーティリティ集
''' </summary>
''' <remarks></remarks>
Public Module DBUtilSS

    Private Const PrmTimeOut As Integer = 60 'ComandTimeOut値

#Region "SQL実行・データ確認・取得"

#Region "OracleOpenチェック"
    ''' <summary>
    ''' コネクションが確立しているかどうかを確認する
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <returns>sqlコネクション</returns>
    ''' <remarks></remarks>
    Public Function PB_ChkConnection(ByVal ocon As SqlConnection) As SqlConnection

        'コネクションが閉じられている場合のみ
        If (ocon.State = ConnectionState.Closed) Then
            Dim Str As String = XMLReadConnection()
            ocon.ConnectionString = Str
            ocon.Open()
        End If

        Return ocon
    End Function
#End Region


#Region "SQL文実行"
    '---------------------------------------------------------
    '　機能：SQL文(INSERT, UPDATE, DELETE)実行
    '
    '　引数　：Connection, 実行SQL文, Optional(Transaction)
    '　戻り値：Boolean(成功可否)
    '---------------------------------------------------------
    'Public Function ExecuteDB(ByVal ocon As SqlConnection, ByVal SQL As String, _
    '                                 Optional ByVal tran As SqlTransaction = Nothing) As Boolean
    ''' <summary>
    ''' SQL文の実行
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文</param>
    ''' <param name="tran">トランザクション</param>
    ''' <param name="intUpdLine"></param>
    ''' <returns>True：sql実行成功　False：sql実行失敗</returns>
    ''' <remarks></remarks>
    Public Function ExecuteDB(ByVal ocon As SqlConnection, ByVal SQL As String, _
                                     Optional ByVal tran As SqlTransaction = Nothing, _
                                     Optional ByVal intUpdLine As Integer = 0) As Boolean
        Dim ocd As New SqlCommand
        ocd.CommandTimeout = PrmTimeOut

        Try
            If IsNothing(tran) Then
                ocon = PB_ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL
            If Not IsNothing(tran) Then
                ocd.Transaction = tran
            End If

            'Modify 20061023_1
            'If ocd.ExecuteNonQuery() < 1 Then
            '    Return False
            'End If
            intUpdLine = ocd.ExecuteNonQuery

            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "データ確認(Boolean)"
    ''' <summary>
    ''' データの存在確認
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <returns>True：存在する　False：存在しない</returns>
    ''' <remarks></remarks>
    Public Function ChkDB(ByVal ocon As SqlConnection, ByVal SQL As String, _
                                 Optional ByVal tran As SqlTransaction = Nothing) As Boolean
        Dim ocd As New SqlCommand
        Dim odr As SqlDataReader = Nothing
        ocd.CommandTimeout = PrmTimeOut



        Try
            If IsNothing(tran) Then
                ocon = PB_ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL

            If Not tran Is Nothing Then
                ocd.Transaction = tran
            End If

            ''DEL 20090807_1
            ''If CInt(ocd.ExecuteScalar()) < 1 Then
            ''    Return False
            ''End If

            odr = ocd.ExecuteReader
            If odr.HasRows Then
                Return True
            Else
                Return False
            End If


        Catch ex As Exception
            Throw ex
        Finally
            If Not odr Is Nothing Then
                odr.Close()
            End If
        End Try
    End Function
#End Region

#Region "データ取得(ONE)"
    ''' <summary>
    ''' １項目のみデータを取得する
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <returns>取得した１項目</returns>
    ''' <remarks></remarks>
    Public Function getOneDataDB(ByVal ocon As SqlConnection, ByVal SQL As String, _
                                          Optional ByVal tran As SqlTransaction = Nothing) As String

        Dim ocd As New SqlCommand
        Try
            If IsNothing(tran) Then
                ocon = PB_ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL

            If Not tran Is Nothing Then
                ocd.Transaction = tran
            End If

            Return CStr(IIf(ocd.ExecuteScalar() Is DBNull.Value, Nothing, ocd.ExecuteScalar))

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "データ取得(ONE RECORD)"
    ''' <summary>
    ''' 配列で１レコードデータを取得する
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <returns>1レコード情報</returns>
    ''' <remarks></remarks>
    Public Function PB_GetARLDataDB(ByVal ocon As SqlConnection, ByVal SQL As String, _
                                        Optional ByVal tran As SqlTransaction = Nothing) As ArrayList
        Dim ocd As New SqlCommand
        Dim odr As SqlDataReader
        Dim arlData As New ArrayList

        Try

            If IsNothing(tran) Then
                ocon = PB_ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL

            If Not tran Is Nothing Then
                ocd.Transaction = tran
            End If

            odr = ocd.ExecuteReader

            While (odr.Read)
                For i As Integer = 0 To odr.FieldCount - 1
                    With arlData
                        .Add(PBCStr(odr.Item(i)))
                    End With
                Next
            End While

            odr.Close()
            Return arlData
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region


#Region "データ取得(DataTable)"
    ''' <summary>
    ''' DataRowオブジェクトでデータを取得する
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <returns>DataRowオブジェクト</returns>
    ''' <remarks></remarks>
    Public Function GetDataRow(ByVal ocon As SqlConnection, ByVal SQL As String, _
                                        Optional ByVal tran As SqlTransaction = Nothing) As DataRow
        Dim ocd As New SqlCommand
        Dim dts As DataSet = New DataSet
        Dim oda As New SqlDataAdapter
        Dim dtt As DataTable
        ocd.CommandTimeout = PrmTimeOut

        Try
            If IsNothing(tran) Then
                ocon = PB_ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL

            If Not tran Is Nothing Then
                ocd.Transaction = tran
            End If

            oda.SelectCommand = ocd

            dts.Tables.Clear()
            oda.Fill(dts)
            dtt = dts.Tables(0)

            If dtt.Rows.Count > 0 Then
                '1件目のみ
                Return dtt.Rows(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            'SkyLog.Debug(SQL)
            Throw ex
        End Try
    End Function
#End Region

#Region "データ取得(DataTable)"
    ''' <summary>
    ''' データテーブルオブジェクトでデータを取得する
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <returns>データテーブルオブジェクト</returns>
    ''' <remarks></remarks>
    Public Function GetDtDataDB(ByVal ocon As SqlConnection, ByVal SQL As String, _
                                        Optional ByVal tran As SqlTransaction = Nothing) As DataTable
        Dim ocd As New SqlCommand
        Dim dts As DataSet = New DataSet
        Dim oda As New SqlDataAdapter
        Dim dtt As DataTable



        Try
            If IsNothing(tran) Then
                ocon = PB_ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL

            ocd.CommandTimeout = PrmTimeOut

            If Not tran Is Nothing Then
                ocd.Transaction = tran
            End If

            oda.SelectCommand = ocd

            dts.Tables.Clear()
            oda.Fill(dts)
            dtt = dts.Tables(0)

            Return dtt

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "データ取得(DataSet)レポート用"
    ''' <summary>
    ''' DataSetオブジェクトを利用した、データの取得メソッド
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <param name="dts">データセットオブジェクト</param>
    ''' <param name="tblName">データテーブル名</param>
    ''' <param name="inMaxRow">最大行数</param>
    ''' <remarks></remarks>
    Public Sub PB_GetDTTSetDB(ByVal ocon As SqlConnection, ByVal SQL As String, _
                                        ByRef dts As DataSet, ByVal tblName As String, _
                                        Optional ByVal tran As SqlTransaction = Nothing, Optional ByVal inMaxRow As Integer = 0)
        Dim ocd As New SqlCommand
        Dim oda As New SqlDataAdapter
        Dim dtt As DataTable
        Dim dtRow As DataRow
        Dim inN As Integer
        Dim inCnt As Integer

        Try
            If IsNothing(tran) Then
                ocon = PB_ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL

            If Not tran Is Nothing Then
                ocd.Transaction = tran
            End If

            oda.SelectCommand = ocd

            'dts.Tables.Clear()
            oda.Fill(dts, tblName)
            dtt = dts.Tables(0)

            '最大行数に満たないレコード数を取得
            inCnt = inMaxRow - dtt.Rows.Count - 1
            For inN = 0 To inCnt
                dtRow = dtt.NewRow
                dtt.Rows.Add(dtRow)
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "データ取得(DateView)"
    ''' <summary>
    ''' 指定したデータテーブルよりレコードを抽出する
    ''' </summary>
    ''' <param name="dt">指定データテーブル</param>
    ''' <param name="szWhere">問合せ文</param>
    ''' <param name="szSort">並び替え条件</param>
    ''' <returns>データビューオブジェクト(抽出されたレコード)</returns>
    ''' <remarks></remarks>
    Public Function GetDtView(ByVal dt As DataTable, _
                                        Optional ByVal szWhere As String = "", _
                                                Optional ByVal szSort As String = "") As DataView
        Try

            Dim dtView As DataView
            dtView = New DataView(dt, szWhere, szSort, DataViewRowState.CurrentRows)

            Return dtView

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "レコード件数の取得"
    ''' <summary>
    ''' レコード件数の取得
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <returns>レコード件数</returns>
    ''' <remarks></remarks>
    Public Function GetRecCount(ByVal ocon As SqlConnection, ByVal SQL As String, _
                                        Optional ByVal tran As SqlTransaction = Nothing) As Integer
        Dim ocd As New SqlCommand
        Dim dts As DataSet = New DataSet
        Dim oda As New SqlDataAdapter
        Dim dtt As DataTable


        ocd.CommandTimeout = PrmTimeOut

        Try
            If IsNothing(tran) Then
                ocon = PB_ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL

            ocd.CommandTimeout = PrmTimeOut

            If Not tran Is Nothing Then
                ocd.Transaction = tran
            End If

            oda.SelectCommand = ocd

            dts.Tables.Clear()
            oda.Fill(dts)
            dtt = dts.Tables(0)

            Return dtt.Rows.Count


        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#End Region





End Module

