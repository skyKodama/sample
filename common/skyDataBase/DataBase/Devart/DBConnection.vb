
Option Explicit On
Option Strict On
Imports Devart.Data.Universal
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
Public Class DBconnection

#Region "Private変数"
    Private Con As New UniConnection             'オラクルコネクション
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
        If Con Is Nothing OrElse (Con.State = ConnectionState.Closed) Then
            Dim Str As String = DBUtil.XMLReadConnection()
            ''ConnetionString複合化
            'Str = skysystem.common.SystemUtil.doDecrypt(Str, "skysystem")

            '''プロバイダ設定
            Str = SetProvider(DBUtil.XMLReadConnectionType(), Str)


            Con.ConnectionString = Str
            Con.Open()


            Select Case DBUtil.XMLReadConnectionType()
                Case DBTYPE.POSTGRESQL
                    ''SearthPath取得
                    Dim path As String
                    path = PB_ReadXML("/SKY/SKY_DB/PATH", "public", SystemConst.C_SYSTEMPRM)

                    ''サーチパス設定
                    DBUtil.ExecuteDB(Con, "SET search_path TO " & path)
            End Select

        Else
            Try
                DBUtil.getOneDataDB(Con, "Select 1 ")
            Catch ex As Exception
                ''States=Brokenが検知できないため、Catch後CloseしOpen
                Con.Close()
                Me.Open()
            End Try
        End If


    End Sub
#End Region

#Region "RtnCon：接続を返す"
    ''' <summary>
    ''' 接続情報を戻す
    ''' </summary>
    ''' <returns>Sqlコネクション</returns>
    ''' <remarks></remarks>
    Public Function rtncon() As UniConnection

        'コネクションが閉じられている場合のみ
        Open()

        Return Con
    End Function
#End Region

#Region "Close：接続を閉じる"
    ''' <summary>
    ''' 接続を閉じる
    ''' </summary>
    Public Sub Close()
        If Not Con Is Nothing Then
            If Con.State = ConnectionState.Open Then
                UniConnection.ClearPool(Con)
                Con.Close()
            End If
        End If
    End Sub
#End Region

    ''' <summary>
    ''' コネクションにプロバイダ設定
    ''' </summary>
    ''' <param name="prmConStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetProvider(prmDBType As DBTYPE, prmConStr As String) As String

        'connStrings.Add("MySql", "Provider=MySql;Host=server;User Id=root;Password=;Database=test;Port=3306")
        'connStrings.Add("PostgreSql", "Provider=PostgreSQL;Host=server;User Id=postgres;Password=;Database=test;Port=5432")
        'connStrings.Add("Oracle", "Provider=Oracle;Data Source=ora;User Id=scott;Password=tiger;Direct=true;SID=;Port=1521")
        'connStrings.Add("OracleClient", "Provider=OracleClient;Data Source=ora;User Id=scott;Password=tiger")
        'connStrings.Add("ODP", "Provider=Odp;Data Source=ora;User Id=scott;Password=tiger")
        'connStrings.Add("SQLite", "Provider=SQLite;Data Source=test.db")
        'connStrings.Add("SQL Server", "Provider=Sql Server;Data Source=server;Initial Catalog=pubs;User Id=sa")
        'connStrings.Add("ODBC", "Provider=ODBC;Driver={Sql Server};UID=sa;Server=server;Database=pubs")
        'connStrings.Add("OLE DB", "Provider=Ole Db;User Id=sa;Data Source=server;Initial Catalog=pubs;Ole Db Provider=SQLOLEDB.1")


        Dim ht As New Hashtable
        ht.Add("0", "PostgreSql")
        ht.Add("1", "SQL Server")
        ht.Add("2", "Oracle")

        Return String.Format("Provider={0}", ht.Item(PBCint(prmDBType).ToString).ToString) & ";" & prmConStr

    End Function


End Class
