
Option Explicit On
Option Strict On
Imports Devart.Data.Universal
Imports skysystem.common.SystemUtil


'*************************************************************************
'*�@�@�\�@�@�FDatabaseConnection
'*            <DB�ڑ�>
'*�@�쐬���@�F
'*
'*�@���ύX���e��
'*
'*************************************************************************
''' <summary>
''' �f�[�^�x�[�X�R�l�N�V�����N���X
''' </summary>
''' <remarks>�R�l�N�V�����̐ڑ��E�ؒf�E�߂�ҿ���</remarks>
Public Class DBconnection

#Region "Private�ϐ�"
    Private Con As New UniConnection             '�I���N���R�l�N�V����
#End Region

#Region "�R���X�g���N�^"
    Public Sub New()

    End Sub
#End Region

#Region "ChkConnection�FOracleOpen�ڑ��F�`�F�b�N"
    ''' <summary>
    ''' �ڑ��I�[�v������
    ''' </summary>
    ''' <remarks>�R�l�N�V�����������Ă���ꍇ�́A�Đڑ����s���B</remarks>
    Public Sub Open()

        '�R�l�N�V�����������Ă���ꍇ�̂�
        If Con Is Nothing OrElse (Con.State = ConnectionState.Closed) Then
            Dim Str As String = DBUtil.XMLReadConnection()
            ''ConnetionString������
            'Str = skysystem.common.SystemUtil.doDecrypt(Str, "skysystem")

            '''�v���o�C�_�ݒ�
            Str = SetProvider(DBUtil.XMLReadConnectionType(), Str)


            Con.ConnectionString = Str
            Con.Open()


            Select Case DBUtil.XMLReadConnectionType()
                Case DBTYPE.POSTGRESQL
                    ''SearthPath�擾
                    Dim path As String
                    path = PB_ReadXML("/SKY/SKY_DB/PATH", "public", SystemConst.C_SYSTEMPRM)

                    ''�T�[�`�p�X�ݒ�
                    DBUtil.ExecuteDB(Con, "SET search_path TO " & path)
            End Select

        Else
            Try
                DBUtil.getOneDataDB(Con, "Select 1 ")
            Catch ex As Exception
                ''States=Broken�����m�ł��Ȃ����߁ACatch��Close��Open
                Con.Close()
                Me.Open()
            End Try
        End If


    End Sub
#End Region

#Region "RtnCon�F�ڑ���Ԃ�"
    ''' <summary>
    ''' �ڑ�����߂�
    ''' </summary>
    ''' <returns>Sql�R�l�N�V����</returns>
    ''' <remarks></remarks>
    Public Function rtncon() As UniConnection

        '�R�l�N�V�����������Ă���ꍇ�̂�
        Open()

        Return Con
    End Function
#End Region

#Region "Close�F�ڑ������"
    ''' <summary>
    ''' �ڑ������
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
    ''' �R�l�N�V�����Ƀv���o�C�_�ݒ�
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
