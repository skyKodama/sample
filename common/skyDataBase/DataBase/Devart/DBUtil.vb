
Option Explicit On
Option Strict On

Imports System.Data
Imports Devart.Data.Universal
Imports System.IO
Imports skysystem.common.SystemUtil
Imports skysystem.common

'****************************************************************************************
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
'*  20190716_1  児玉  アプリケーションフォルダへのログの出力を追加
'*                    独立性を保たせるため、外部参照ではなく同モジュール内に処理を複製
'*
'****************************************************************************************
''' <summary>
''' データベース用ユーティリティ集
''' </summary>
''' <remarks></remarks>
Public Module DBUtil

    Private Const PrmTimeOut As Integer = 120 'ComandTimeOut値
    ''' <summary>
    ''' コネクションタイプ
    ''' </summary>
    ''' <remarks></remarks>
    Enum DBTYPE
        ORACLE = 1
        SQLSERVER = 1
        POSTGRESQL = 0
    End Enum


#Region "SQL実行・データ確認・取得"


    ''' <summary>
    ''' XMLより接続文字列を取得する
    ''' </summary>
    ''' <returns>接続文字列</returns>
    ''' <remarks></remarks>
    Public Function XMLReadConnection() As String
        Return PB_ReadXML("/SKY/SKY_DB/CONNECTION", "", SystemConst.C_SYSTEMPRM)
    End Function

    ''' <summary>
    ''' コネクション種別を取得
    ''' </summary>
    ''' <returns>接続文字列</returns>
    ''' <remarks></remarks>
    Public Function XMLReadConnectionType() As DBTYPE
        Return CType(PB_ReadXML("/SKY/SKY_DB/DBTYPE", "", SystemConst.C_SYSTEMPRM), DBTYPE)
    End Function



#Region "OracleOpenチェック"
    ''' <summary>
    ''' コネクションが確立しているかどうかを確認する
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <returns>sqlコネクション</returns>
    ''' <remarks></remarks>
    Public Function PB_ChkConnection(ByVal ocon As UniConnection) As UniConnection

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
    'Public Function ExecuteDB(ByVal ocon As uniConnection, ByVal SQL As String, _
    '                                 Optional ByVal tran As UniTransaction = Nothing) As Boolean
    ''' <summary>
    ''' SQL文の実行
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文</param>
    ''' <param name="tran">トランザクション</param>
    ''' <param name="intUpdLine"></param>
    ''' <returns>True：sql実行成功　False：sql実行失敗</returns>
    ''' <remarks></remarks>
    Public Function ExecuteDB(ByVal ocon As UniConnection, ByVal SQL As String, _
                                     Optional ByVal tran As UniTransaction = Nothing, _
                                     Optional ByVal intUpdLine As Integer = 0) As Boolean
        Dim ocd As New UniCommand
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

      
        Catch ex As UniException
            LogOutPut_Error(SQL, "DBUtil.ExecuteDB", ex.Message)
            If ocon.Provider = "SQL Server" Then
                ''SQLServerはBiginTransactionがないとRollBackできない
                If CType(ex.InnerException, SqlClient.SqlException).Number = 3903 Then
                    Return True
                End If
            End If

            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try

            Throw ex
        End Try
    End Function
#End Region

#Region "データ確認(Boolean)"
    ''' <summary>
    ''' データの存在確認
    ''' True：存在する　False：存在しない
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <returns>True：存在する　False：存在しない</returns>
    ''' <remarks>True：DBに存在　False：DBに存在しない</remarks>
    Public Function ChkDB(ByVal ocon As UniConnection, ByVal SQL As String, _
                                 Optional ByVal tran As UniTransaction = Nothing) As Boolean
        Dim ocd As New UniCommand
        Dim odr As UniDataReader = Nothing
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
            LogOutPut_Error(SQL, "DBUtil.ChkDB", ex.Message)

            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try

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
    Public Function getOneDataDB(ByVal ocon As UniConnection, ByVal SQL As String, _
                                          Optional ByVal tran As UniTransaction = Nothing) As String

        Dim ocd As New UniCommand
        Dim reader As UniDataReader = Nothing
        Dim rtnValue As String = ""


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


            reader = ocd.ExecuteReader()

            If reader.HasRows Then
                Do While reader.Read()
                    rtnValue = reader(0).ToString
                    Exit Do
                Loop
            End If

            Return rtnValue

        Catch ex As UniException
            LogOutPut_Error(SQL, "DBUtil.getOneDataDB", ex.Message)
            'If ex.Number = 54 Then
            '    ''ロック情報のLoggin
            '    Dim lockInfo As String
            '    lockInfo = doWriteLockInfo(ocon, tran)
            '    MessageUtil.ShowErrorMsg("現在、別のユーザーによってデータが使用中です。" & vbCrLf & lockInfo)
            'End If
            Throw ex : Return Nothing
        Catch ex As Exception
            LogOutPut(SQL, "DBUtil")

            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try

            Throw ex : Return Nothing
        Finally
            If Not reader Is Nothing Then
                reader.Close()
                reader = Nothing
            End If
        End Try

        'Dim ocd As New uniCommand
        'Try
        '    If IsNothing(tran) Then
        '        ocon = PB_ChkConnection(ocon)
        '    End If

        '    ocd.Connection = ocon
        '    ocd.CommandText = SQL

        '    If Not tran Is Nothing Then
        '        ocd.Transaction = tran
        '    End If

        '    Return CStr(IIf(ocd.ExecuteScalar() Is DBNull.Value, Nothing, ocd.ExecuteScalar))

        'Catch ex As Exception
        '    Throw ex
        'End Try
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
    Public Function getAryDataDB(ByVal ocon As UniConnection, ByVal SQL As String, _
                                        Optional ByVal tran As UniTransaction = Nothing) As ArrayList
        Dim ocd As New UniCommand
        Dim odr As UniDataReader
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
            LogOutPut_Error(SQL, "DBUtil.getAryDataDB", ex.Message)
            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try
            Throw ex
        End Try
    End Function
#End Region
#Region "データ取得(DataReader)"
    '---------------------------------------------------------
    '　機能：データゲット(DataTable)
    '
    '　引数　：Connection, SQL文, Optional(Transaction)
    '　戻り値：DataTable(ゲットしたもの)
    '---------------------------------------------------------
    Public Function getDataReader(ByVal ocon As UniConnection, ByVal SQL As String, _
                                        Optional ByVal tran As UniTransaction = Nothing) As UniDataReader
        Dim ocd As New UniCommand
        Dim reader As UniDataReader = Nothing



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


            reader = ocd.ExecuteReader()

            Return reader

        Catch ex As UniException
            LogOutPut_Error(SQL, "DBUtil.getDataReader", ex.Message)
            'If ex.Number = 54 Then
            '    ''ロック情報のLoggin
            '    Dim lockInfo As String
            '    lockInfo = doWriteLockInfo(ocon, tran)
            '    MessageUtil.ShowErrorMsg("現在、別のユーザーによってデータが使用中です。" & vbCrLf & lockInfo)
            'End If
            Throw ex : Return Nothing
        Catch ex As Exception
            LogOutPut_Error(SQL, "DBUtil.getDataReader", ex.Message)

            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try

            Throw ex : Return Nothing
        Finally
            'If tran Is Nothing Then
            '    ocon.Close()
            'End If
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
    Public Function GetDataRow(ByVal ocon As UniConnection, ByVal SQL As String, _
                                        Optional ByVal tran As UniTransaction = Nothing) As DataRow
        Dim ocd As New UniCommand
        Dim dts As DataSet = New DataSet
        Dim oda As New UniDataAdapter
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
            LogOutPut_Error(SQL, "DBUtil.GetDataRow", ex.Message)
            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try
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
    Public Function GetDtDataDB(ByVal ocon As UniConnection, ByVal SQL As String, _
                                        Optional ByVal tran As UniTransaction = Nothing, Optional prmTableName As String = "") As DataTable
        Dim ocd As New UniCommand
        'Dim dts As DataSet = New DataSet
        Dim oda As New UniDataAdapter
        Dim dtt As New DataTable(prmTableName)



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

            'dts.Tables.Clear()
            oda.Fill(dtt)
            'dtt = dts.Tables(0)

            Return dtt

        Catch ex As Exception
            LogOutPut_Error(SQL, "DBUtil.GetDtDataDB", ex.Message)

            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try

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
    Public Sub PB_GetDTTSetDB(ByVal ocon As UniConnection, ByVal SQL As String, _
                                        ByRef dts As DataSet, ByVal tblName As String, _
                                        Optional ByVal tran As UniTransaction = Nothing, Optional ByVal inMaxRow As Integer = 0)
        Dim ocd As New UniCommand
        Dim oda As New UniDataAdapter
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
            LogOutPut_Error(SQL, "DBUtil.PB_GetDTTSetDB", ex.Message)
            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try
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
            LogOutPut_Error(szWhere, "DBUtil.GetDtView", ex.Message)
            Throw ex
        End Try
    End Function
#End Region

#Region "Seq取得"
    ''' <summary>
    ''' Seqを取得する
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="seqObject">Sequenceオブジェクト(</param>
    ''' <param name="tran">sqlトランザクション</param>
    ''' <returns>1レコード情報</returns>
    ''' <remarks></remarks>
    Public Function getSequence(ByVal ocon As UniConnection, ByVal seqObject As String, _
                                      Optional ByVal tran As UniTransaction = Nothing) As Integer
        Try


            Dim iseq As Integer
            iseq = PBCint(DBUtil.getOneDataDB(ocon, String.Format("Select NextVal('{0}')", seqObject), tran))

            Return iseq

        Catch ex As Exception
            LogOutPut_Error(String.Format("Select NextVal('{0}')", seqObject), "DBUtil.getSequence", ex.Message)
            Try
                'エラー時にコネクションを閉じる
                ocon.Close()
            Catch ex2 As Exception

            End Try
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "ログ出力"
    '*****************************************************
    '* テスト用簡易ログ出力
    '*****************************************************

    Private LogPath As String = ".\Logs\"
    Private FileNm As String = "Log_999999.txt"

    Private Sub LogOutPut(prmMsg As String, prmPGMID As String)
        Try
            MakeLogFileName()

            Dim sw As New System.IO.StreamWriter(LogPath & FileNm, True)

            sw.WriteLine(setCommonInfo() & prmPGMID & vbTab & prmMsg)

            sw.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LogOutPut_Error(prmMsg As String, prmPGMID As String, ex_message As String)
        Try
            MakeLogFileName()

            Dim sw As New System.IO.StreamWriter(LogPath & FileNm, True)

            sw.WriteLine(setCommonInfo() & prmPGMID & vbTab & ex_message & vbCrLf & prmMsg)

            sw.Close()
        Catch ex As Exception

        End Try
    End Sub


#Region "ファイル名作成"
    Private Sub MakeLogFileName()
        Try
            '日付が付加されていない場合、パスを生成
            If LogPath = ".\Logs\" Then
                LogPath = LogPath + Now.ToString("yyyy") + "\" + Now.ToString("MM") + "\"
            End If

            If Not System.IO.Directory.Exists(LogPath) Then
                System.IO.Directory.CreateDirectory(LogPath)
            End If


            FileNm = String.Format("Log_{0}.txt", Now.ToString("yyyyMMdd"))

        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "共通情報"
    Private Function setCommonInfo() As String
        Dim rtnInfo As String = ""

        rtnInfo += Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab

        Return rtnInfo
    End Function

#End Region

#End Region

End Module

