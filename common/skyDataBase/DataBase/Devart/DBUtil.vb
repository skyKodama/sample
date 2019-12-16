
Option Explicit On
Option Strict On

Imports System.Data
Imports Devart.Data.Universal
Imports System.IO
Imports skysystem.common.SystemUtil
Imports skysystem.common

'****************************************************************************************
'*�@�@�\�@�@�F����DatabasePB�N���X
'*            <DB�ڑ�>
'*            <DB���ʖ߂�>
'*�@�쐬���@�F2006.05.22    ��
'*
'*�@���ύX���e��
'*�@20060619_1�@��`�@SQL�쐬AtoZ(PBFSTR_CreatSqlAtoZ)
'*  20060721_1  ��`�@���|�[�g�p�f�[�^�Z�b�g�擾   
'*  20060728_1  ��`  �f�[�^�e�[�u���ɍő�s�ɖ����Ȃ��󃌃R�[�h��ǉ�����(AddRow)
'*  20061010_1  ��`  PBFSTR_SQLMltSgl(SQL�\�z(�S�p���p����ʂ��Ȃ�))
'*  20061023_1  ��    ExecuteDB�A�X�V�s���߂�l�ǉ�
'*  20070523_1  ��`  DataVeiw�擾
'*  20190716_1  ����  �A�v���P�[�V�����t�H���_�ւ̃��O�̏o�͂�ǉ�
'*                    �Ɨ�����ۂ����邽�߁A�O���Q�Ƃł͂Ȃ������W���[�����ɏ����𕡐�
'*
'****************************************************************************************
''' <summary>
''' �f�[�^�x�[�X�p���[�e�B���e�B�W
''' </summary>
''' <remarks></remarks>
Public Module DBUtil

    Private Const PrmTimeOut As Integer = 120 'ComandTimeOut�l
    ''' <summary>
    ''' �R�l�N�V�����^�C�v
    ''' </summary>
    ''' <remarks></remarks>
    Enum DBTYPE
        ORACLE = 1
        SQLSERVER = 1
        POSTGRESQL = 0
    End Enum


#Region "SQL���s�E�f�[�^�m�F�E�擾"


    ''' <summary>
    ''' XML���ڑ���������擾����
    ''' </summary>
    ''' <returns>�ڑ�������</returns>
    ''' <remarks></remarks>
    Public Function XMLReadConnection() As String
        Return PB_ReadXML("/SKY/SKY_DB/CONNECTION", "", SystemConst.C_SYSTEMPRM)
    End Function

    ''' <summary>
    ''' �R�l�N�V������ʂ��擾
    ''' </summary>
    ''' <returns>�ڑ�������</returns>
    ''' <remarks></remarks>
    Public Function XMLReadConnectionType() As DBTYPE
        Return CType(PB_ReadXML("/SKY/SKY_DB/DBTYPE", "", SystemConst.C_SYSTEMPRM), DBTYPE)
    End Function



#Region "OracleOpen�`�F�b�N"
    ''' <summary>
    ''' �R�l�N�V�������m�����Ă��邩�ǂ������m�F����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <returns>sql�R�l�N�V����</returns>
    ''' <remarks></remarks>
    Public Function PB_ChkConnection(ByVal ocon As UniConnection) As UniConnection

        '�R�l�N�V�����������Ă���ꍇ�̂�
        If (ocon.State = ConnectionState.Closed) Then
            Dim Str As String = XMLReadConnection()
            ocon.ConnectionString = Str

            ocon.Open()
        End If

        Return ocon
    End Function
#End Region



#Region "SQL�����s"
    '---------------------------------------------------------
    '�@�@�\�FSQL��(INSERT, UPDATE, DELETE)���s
    '
    '�@�����@�FConnection, ���sSQL��, Optional(Transaction)
    '�@�߂�l�FBoolean(������)
    '---------------------------------------------------------
    'Public Function ExecuteDB(ByVal ocon As uniConnection, ByVal SQL As String, _
    '                                 Optional ByVal tran As UniTransaction = Nothing) As Boolean
    ''' <summary>
    ''' SQL���̎��s
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��</param>
    ''' <param name="tran">�g�����U�N�V����</param>
    ''' <param name="intUpdLine"></param>
    ''' <returns>True�Fsql���s�����@False�Fsql���s���s</returns>
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
                ''SQLServer��BiginTransaction���Ȃ���RollBack�ł��Ȃ�
                If CType(ex.InnerException, SqlClient.SqlException).Number = 3903 Then
                    Return True
                End If
            End If

            Try
                '�G���[���ɃR�l�N�V���������
                ocon.Close()
            Catch ex2 As Exception

            End Try

            Throw ex
        End Try
    End Function
#End Region

#Region "�f�[�^�m�F(Boolean)"
    ''' <summary>
    ''' �f�[�^�̑��݊m�F
    ''' True�F���݂���@False�F���݂��Ȃ�
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>True�F���݂���@False�F���݂��Ȃ�</returns>
    ''' <remarks>True�FDB�ɑ��݁@False�FDB�ɑ��݂��Ȃ�</remarks>
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
                '�G���[���ɃR�l�N�V���������
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

#Region "�f�[�^�擾(ONE)"
    ''' <summary>
    ''' �P���ڂ̂݃f�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>�擾�����P����</returns>
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
            '    ''���b�N����Loggin
            '    Dim lockInfo As String
            '    lockInfo = doWriteLockInfo(ocon, tran)
            '    MessageUtil.ShowErrorMsg("���݁A�ʂ̃��[�U�[�ɂ���ăf�[�^���g�p���ł��B" & vbCrLf & lockInfo)
            'End If
            Throw ex : Return Nothing
        Catch ex As Exception
            LogOutPut(SQL, "DBUtil")

            Try
                '�G���[���ɃR�l�N�V���������
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

#Region "�f�[�^�擾(ONE RECORD)"
    ''' <summary>
    ''' �z��łP���R�[�h�f�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>1���R�[�h���</returns>
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
                '�G���[���ɃR�l�N�V���������
                ocon.Close()
            Catch ex2 As Exception

            End Try
            Throw ex
        End Try
    End Function
#End Region
#Region "�f�[�^�擾(DataReader)"
    '---------------------------------------------------------
    '�@�@�\�F�f�[�^�Q�b�g(DataTable)
    '
    '�@�����@�FConnection, SQL��, Optional(Transaction)
    '�@�߂�l�FDataTable(�Q�b�g��������)
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
            '    ''���b�N����Loggin
            '    Dim lockInfo As String
            '    lockInfo = doWriteLockInfo(ocon, tran)
            '    MessageUtil.ShowErrorMsg("���݁A�ʂ̃��[�U�[�ɂ���ăf�[�^���g�p���ł��B" & vbCrLf & lockInfo)
            'End If
            Throw ex : Return Nothing
        Catch ex As Exception
            LogOutPut_Error(SQL, "DBUtil.getDataReader", ex.Message)

            Try
                '�G���[���ɃR�l�N�V���������
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
#Region "�f�[�^�擾(DataTable)"
    ''' <summary>
    ''' DataRow�I�u�W�F�N�g�Ńf�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>DataRow�I�u�W�F�N�g</returns>
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
                '1���ڂ̂�
                Return dtt.Rows(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            LogOutPut_Error(SQL, "DBUtil.GetDataRow", ex.Message)
            Try
                '�G���[���ɃR�l�N�V���������
                ocon.Close()
            Catch ex2 As Exception

            End Try
            'SkyLog.Debug(SQL)
            Throw ex
        End Try
    End Function
#End Region

#Region "�f�[�^�擾(DataTable)"
    ''' <summary>
    ''' �f�[�^�e�[�u���I�u�W�F�N�g�Ńf�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>�f�[�^�e�[�u���I�u�W�F�N�g</returns>
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
                '�G���[���ɃR�l�N�V���������
                ocon.Close()
            Catch ex2 As Exception

            End Try

            Throw ex
        End Try
    End Function
#End Region

#Region "�f�[�^�擾(DataSet)���|�[�g�p"
    ''' <summary>
    ''' DataSet�I�u�W�F�N�g�𗘗p�����A�f�[�^�̎擾���\�b�h
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <param name="dts">�f�[�^�Z�b�g�I�u�W�F�N�g</param>
    ''' <param name="tblName">�f�[�^�e�[�u����</param>
    ''' <param name="inMaxRow">�ő�s��</param>
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

            '�ő�s���ɖ����Ȃ����R�[�h�����擾
            inCnt = inMaxRow - dtt.Rows.Count - 1
            For inN = 0 To inCnt
                dtRow = dtt.NewRow
                dtt.Rows.Add(dtRow)
            Next

        Catch ex As Exception
            LogOutPut_Error(SQL, "DBUtil.PB_GetDTTSetDB", ex.Message)
            Try
                '�G���[���ɃR�l�N�V���������
                ocon.Close()
            Catch ex2 As Exception

            End Try
            Throw ex
        End Try
    End Sub
#End Region

#Region "�f�[�^�擾(DateView)"
    ''' <summary>
    ''' �w�肵���f�[�^�e�[�u����背�R�[�h�𒊏o����
    ''' </summary>
    ''' <param name="dt">�w��f�[�^�e�[�u��</param>
    ''' <param name="szWhere">�⍇����</param>
    ''' <param name="szSort">���ёւ�����</param>
    ''' <returns>�f�[�^�r���[�I�u�W�F�N�g(���o���ꂽ���R�[�h)</returns>
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

#Region "Seq�擾"
    ''' <summary>
    ''' Seq���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="seqObject">Sequence�I�u�W�F�N�g(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>1���R�[�h���</returns>
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
                '�G���[���ɃR�l�N�V���������
                ocon.Close()
            Catch ex2 As Exception

            End Try
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "���O�o��"
    '*****************************************************
    '* �e�X�g�p�ȈՃ��O�o��
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


#Region "�t�@�C�����쐬"
    Private Sub MakeLogFileName()
        Try
            '���t���t������Ă��Ȃ��ꍇ�A�p�X�𐶐�
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

#Region "���ʏ��"
    Private Function setCommonInfo() As String
        Dim rtnInfo As String = ""

        rtnInfo += Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab

        Return rtnInfo
    End Function

#End Region

#End Region

End Module

