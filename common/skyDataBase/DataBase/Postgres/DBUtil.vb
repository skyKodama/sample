
Option Explicit On
Option Strict On

Imports System.Data
Imports Npgsql
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
Public Module DBUtilNpg

    Private Const PrmTimeOut As Integer = 60 'ComandTimeOut値
    'Private Const PrmTimeOut As Integer = 20 'ComandTimeOut値

#Region "SQL実行・データ確認・取得"

#Region "OracleOpenチェック"
    ''' <summary>
    ''' コネクションが確立しているかどうかを確認する
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <returns>sqlコネクション</returns>
    ''' <remarks></remarks>
    Public Function PB_ChkConnection(ByVal ocon As NpgsqlConnection) As NpgsqlConnection

        'コネクションが閉じられている場合のみ
        If (ocon.State = ConnectionState.Closed) Then
            Dim Str As String = XMLReadConnection()
            ocon.ConnectionString = Str
            ocon.Open()
        End If

        Return ocon
    End Function
#End Region

#Region "XMLReadConnection：XML接続文字列取得(CONNECTION)"
    ''' <summary>
    ''' XMLより接続文字列を取得する
    ''' </summary>
    ''' <returns>接続文字列</returns>
    ''' <remarks></remarks>
    Public Function XMLReadConnection() As String
        Return PB_ReadXML("/SKY/SKY_DB/CONNECTION", "", SystemConst.C_SYSTEMPRM)
    End Function

#End Region

#Region "SQL文実行"
    '---------------------------------------------------------
    '　機能：SQL文(INSERT, UPDATE, DELETE)実行
    '
    '　引数　：Connection, 実行SQL文, Optional(Transaction)
    '　戻り値：Boolean(成功可否)
    '---------------------------------------------------------
    'Public Function ExecuteDB(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
    '                                 Optional ByVal tran As NpgsqlTransaction = Nothing) As Boolean
    ''' <summary>
    ''' SQL文の実行
    ''' </summary>
    ''' <param name="ocon">sqlコネクション</param>
    ''' <param name="SQL">sql文</param>
    ''' <param name="tran">トランザクション</param>
    ''' <param name="intUpdLine"></param>
    ''' <returns>True：sql実行成功　False：sql実行失敗</returns>
    ''' <remarks></remarks>
    Public Function ExecuteDB(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
                                     Optional ByVal tran As NpgsqlTransaction = Nothing, _
                                     Optional ByVal intUpdLine As Integer = 0) As Boolean
        Dim ocd As New NpgsqlCommand
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
    Public Function ChkDB(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
                                 Optional ByVal tran As NpgsqlTransaction = Nothing) As Boolean
        Dim ocd As New NpgsqlCommand
        Dim odr As NpgsqlDataReader = Nothing
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
    Public Function getOneDataDB(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
                                          Optional ByVal tran As NpgsqlTransaction = Nothing) As String

        Dim ocd As New NpgsqlCommand
        Dim reader As NpgsqlDataReader = Nothing
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

        Catch ex As NpgsqlException
            'If ex.Number = 54 Then
            '    ''ロック情報のLoggin
            '    Dim lockInfo As String
            '    lockInfo = doWriteLockInfo(ocon, tran)
            '    MessageUtil.ShowErrorMsg("現在、別のユーザーによってデータが使用中です。" & vbCrLf & lockInfo)
            'End If
            Throw ex : Return Nothing
        Catch ex As Exception
            Throw ex : Return Nothing
        Finally
            reader.Close()
            reader = Nothing
            'If tran Is Nothing Then
            '    ocon.Close()
            'End If
        End Try

        'Dim ocd As New NpgsqlCommand
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
    Public Function getAryDataDB(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
                                        Optional ByVal tran As NpgsqlTransaction = Nothing) As ArrayList
        Dim ocd As New NpgsqlCommand
        Dim odr As NpgsqlDataReader
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
#Region "データ取得(DataReader)"
    '---------------------------------------------------------
    '　機能：データゲット(DataTable)
    '
    '　引数　：Connection, SQL文, Optional(Transaction)
    '　戻り値：DataTable(ゲットしたもの)
    '---------------------------------------------------------
    Public Function getDataReader(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
                                        Optional ByVal tran As NpgsqlTransaction = Nothing) As NpgsqlDataReader
        Dim ocd As New NpgsqlCommand
        Dim reader As NpgsqlDataReader = Nothing



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

        Catch ex As NpgsqlException
            'If ex.Number = 54 Then
            '    ''ロック情報のLoggin
            '    Dim lockInfo As String
            '    lockInfo = doWriteLockInfo(ocon, tran)
            '    MessageUtil.ShowErrorMsg("現在、別のユーザーによってデータが使用中です。" & vbCrLf & lockInfo)
            'End If
            Throw ex : Return Nothing
        Catch ex As Exception
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
    Public Function GetDataRow(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
                                        Optional ByVal tran As NpgsqlTransaction = Nothing) As DataRow
        Dim ocd As New NpgsqlCommand
        Dim dts As DataSet = New DataSet
        Dim oda As New NpgsqlDataAdapter
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
    Public Function GetDtDataDB(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
                                        Optional ByVal tran As NpgsqlTransaction = Nothing) As DataTable
        Dim ocd As New NpgsqlCommand
        'Dim dts As DataSet = New DataSet
        Dim oda As New NpgsqlDataAdapter
        Dim dtt As New DataTable



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
    Public Sub PB_GetDTTSetDB(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
                                        ByRef dts As DataSet, ByVal tblName As String, _
                                        Optional ByVal tran As NpgsqlTransaction = Nothing, Optional ByVal inMaxRow As Integer = 0)
        Dim ocd As New NpgsqlCommand
        Dim oda As New NpgsqlDataAdapter
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


#End Region


#Region "未使用のためPrivateに変更"
#Region "シングルコーテーション追加"
    '------------------------------------------------------------------
    ' 機能         : シングルコーテーション追加
    '
    ' 返り値       : 正常終了 = 変換後の文字列
    '                異常終了 = ""
    '
    ' 引き数       : (IN) strVal  入力文字列
    '
    ' 機能説明     : 文字列中のシングルコーテーションを検索しシングルコーテーションを二重にする
    '
    ' 備考         : 2005.11.07 駒形
    '
    '------------------------------------------------------------------
    Private Function PBFSTR_ChangeQuotation(ByVal strVal As String) As String

        Dim intLocation As Integer        'シングルコーテーションの位置
        Dim strOutputVal As String
        Dim strInputVal As String

        ''出力値変数を初期化
        strOutputVal = ""

        ''入力値を入力値変数に移送
        strInputVal = strVal

        ''入力値変数のシングルコーテーションの位置を検査
        intLocation = InStr(strInputVal, "'")

        ''シングルコーテーションの位置が０より大きい間ループする。
        While intLocation > 0
            ''出力値変数に入力値変数のシングルコーテーションの位置までとシングルコーテーションを出力。
            strOutputVal = strOutputVal & Left$(strInputVal, intLocation) & "'"
            ''入力値変数から出力値変数に出力した文字列を削除する。
            strInputVal = Mid$(strInputVal, intLocation + 1, Len(strInputVal) - intLocation)
            ''入力値変数のシングルコーテーションの位置を検査
            intLocation = InStr(strInputVal, "'")
            ''ループ終了
        End While

        ''戻り値を設定
        PBFSTR_ChangeQuotation = strOutputVal & strInputVal
    End Function
#End Region

#Region "SQL構築(AtoZ)"
    ' ------------------------------------------------------------------ 
    ' @(e) 
    ' 
    ' 機能        : PBFSTR_CreatSqlAtoZ
    ' 
    ' 返り値      : String()
    ' 
    ' 引き数      : strCDST：開始ｺｰﾄﾞ
    '               strCDED：終了ｺｰﾄﾞ
    '               strFLD：
    '               strWhere：
    '
    ' 機能説明    : 開始～終了のSQLを構築
    ' 備考        : 
    '               
    ''------------------------------------------------------------------
    Private Function PBFSTR_CreatSqlAtoZ(ByVal strCDST As String, ByVal strCDED As String, _
                                        ByVal strFLD As String, Optional ByVal strWhere As String = "") As String

        Dim strResult As String

        If strCDST = "" And strCDED = "" Then '開始終了ブランク
            strResult = ""

        ElseIf strCDST <> "" And strCDED = "" Then  ''開始のみ
            strResult = strFLD & ">=" & strCDST

        ElseIf strCDST = "" And strCDED <> "" Then ''終了のみ
            strResult = strFLD & "<=" & strCDED

        Else
            strResult = strFLD & ">=" & strCDST & " AND " & strFLD & "<=" & strCDED

        End If

        Return strResult
    End Function
#End Region

#Region "SQL構築(全角半角を区別しない)"
    '------------------------------------------------------------------------
    '(引数)　
    'strFLD         フィールド名
    'strText        検索値
    'inKBN          0:中間一致　1:前方一致  2:後方一致  (OPT=0)
    '------------------------------------------------------------------------
    Private Function PBFSTR_SQLMltSgl(ByVal strFLD As String, ByVal strText As String, _
                                    Optional ByVal inKBN As Integer = 0) As String
        Dim strSQL As String = ""
        Select Case inKBN

            Case 0 '中間一致検索
                ''全角
                strSQL = strSQL & " ( " & strFLD & " LIKE  ('%" & StrConv(strText, VbStrConv.Wide) & "%')"
                ''半角
                strSQL = strSQL & " OR " & strFLD & " LIKE  ('%" & StrConv(strText, VbStrConv.Narrow) & "%'))"

            Case 1 '前方一致検索
                ''全角
                strSQL = strSQL & " ( " & strFLD & " LIKE  ('" & StrConv(strText, VbStrConv.Wide) & "%')"
                ''半角
                strSQL = strSQL & " OR " & strFLD & " LIKE  ('" & StrConv(strText, VbStrConv.Narrow) & "%'))"

            Case 2 '後方一致検索
                ''全角
                strSQL = strSQL & " ( " & strFLD & " LIKE  ('%" & StrConv(strText, VbStrConv.Wide) & "')"
                ''半角
                strSQL = strSQL & " OR " & strFLD & " LIKE  ('%" & StrConv(strText, VbStrConv.Narrow) & "'))"
        End Select
        Return strSQL
    End Function
#End Region
#End Region


End Module

