
Option Explicit On
Option Strict On

Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports skysystem.common.SystemUtil
Imports skysystem.common

'*************************************************************************
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
'*
'*************************************************************************
''' <summary>
''' �f�[�^�x�[�X�p���[�e�B���e�B�W
''' </summary>
''' <remarks></remarks>
Public Module DBUtilSS

    Private Const PrmTimeOut As Integer = 60 'ComandTimeOut�l

#Region "SQL���s�E�f�[�^�m�F�E�擾"

#Region "OracleOpen�`�F�b�N"
    ''' <summary>
    ''' �R�l�N�V�������m�����Ă��邩�ǂ������m�F����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <returns>sql�R�l�N�V����</returns>
    ''' <remarks></remarks>
    Public Function PB_ChkConnection(ByVal ocon As SqlConnection) As SqlConnection

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
    'Public Function ExecuteDB(ByVal ocon As SqlConnection, ByVal SQL As String, _
    '                                 Optional ByVal tran As SqlTransaction = Nothing) As Boolean
    ''' <summary>
    ''' SQL���̎��s
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��</param>
    ''' <param name="tran">�g�����U�N�V����</param>
    ''' <param name="intUpdLine"></param>
    ''' <returns>True�Fsql���s�����@False�Fsql���s���s</returns>
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

#Region "�f�[�^�m�F(Boolean)"
    ''' <summary>
    ''' �f�[�^�̑��݊m�F
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>True�F���݂���@False�F���݂��Ȃ�</returns>
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

#Region "�f�[�^�擾(ONE)"
    ''' <summary>
    ''' �P���ڂ̂݃f�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>�擾�����P����</returns>
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

#Region "�f�[�^�擾(ONE RECORD)"
    ''' <summary>
    ''' �z��łP���R�[�h�f�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>1���R�[�h���</returns>
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


#Region "�f�[�^�擾(DataTable)"
    ''' <summary>
    ''' DataRow�I�u�W�F�N�g�Ńf�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>DataRow�I�u�W�F�N�g</returns>
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
                '1���ڂ̂�
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

#Region "�f�[�^�擾(DataTable)"
    ''' <summary>
    ''' �f�[�^�e�[�u���I�u�W�F�N�g�Ńf�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>�f�[�^�e�[�u���I�u�W�F�N�g</returns>
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

            '�ő�s���ɖ����Ȃ����R�[�h�����擾
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
            Throw ex
        End Try
    End Function
#End Region

#Region "���R�[�h�����̎擾"
    ''' <summary>
    ''' ���R�[�h�����̎擾
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>���R�[�h����</returns>
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

