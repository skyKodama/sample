
Option Explicit On
Option Strict On

Imports System.Data
Imports Npgsql
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
Public Module DBUtilNpg

    Private Const PrmTimeOut As Integer = 60 'ComandTimeOut�l
    'Private Const PrmTimeOut As Integer = 20 'ComandTimeOut�l

#Region "SQL���s�E�f�[�^�m�F�E�擾"

#Region "OracleOpen�`�F�b�N"
    ''' <summary>
    ''' �R�l�N�V�������m�����Ă��邩�ǂ������m�F����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <returns>sql�R�l�N�V����</returns>
    ''' <remarks></remarks>
    Public Function PB_ChkConnection(ByVal ocon As NpgsqlConnection) As NpgsqlConnection

        '�R�l�N�V�����������Ă���ꍇ�̂�
        If (ocon.State = ConnectionState.Closed) Then
            Dim Str As String = XMLReadConnection()
            ocon.ConnectionString = Str
            ocon.Open()
        End If

        Return ocon
    End Function
#End Region

#Region "XMLReadConnection�FXML�ڑ�������擾(CONNECTION)"
    ''' <summary>
    ''' XML���ڑ���������擾����
    ''' </summary>
    ''' <returns>�ڑ�������</returns>
    ''' <remarks></remarks>
    Public Function XMLReadConnection() As String
        Return PB_ReadXML("/SKY/SKY_DB/CONNECTION", "", SystemConst.C_SYSTEMPRM)
    End Function

#End Region

#Region "SQL�����s"
    '---------------------------------------------------------
    '�@�@�\�FSQL��(INSERT, UPDATE, DELETE)���s
    '
    '�@�����@�FConnection, ���sSQL��, Optional(Transaction)
    '�@�߂�l�FBoolean(������)
    '---------------------------------------------------------
    'Public Function ExecuteDB(ByVal ocon As NpgsqlConnection, ByVal SQL As String, _
    '                                 Optional ByVal tran As NpgsqlTransaction = Nothing) As Boolean
    ''' <summary>
    ''' SQL���̎��s
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��</param>
    ''' <param name="tran">�g�����U�N�V����</param>
    ''' <param name="intUpdLine"></param>
    ''' <returns>True�Fsql���s�����@False�Fsql���s���s</returns>
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

#Region "�f�[�^�m�F(Boolean)"
    ''' <summary>
    ''' �f�[�^�̑��݊m�F
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>True�F���݂���@False�F���݂��Ȃ�</returns>
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

#Region "�f�[�^�擾(ONE)"
    ''' <summary>
    ''' �P���ڂ̂݃f�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>�擾�����P����</returns>
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
            '    ''���b�N����Loggin
            '    Dim lockInfo As String
            '    lockInfo = doWriteLockInfo(ocon, tran)
            '    MessageUtil.ShowErrorMsg("���݁A�ʂ̃��[�U�[�ɂ���ăf�[�^���g�p���ł��B" & vbCrLf & lockInfo)
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

#Region "�f�[�^�擾(ONE RECORD)"
    ''' <summary>
    ''' �z��łP���R�[�h�f�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>1���R�[�h���</returns>
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
#Region "�f�[�^�擾(DataReader)"
    '---------------------------------------------------------
    '�@�@�\�F�f�[�^�Q�b�g(DataTable)
    '
    '�@�����@�FConnection, SQL��, Optional(Transaction)
    '�@�߂�l�FDataTable(�Q�b�g��������)
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
            '    ''���b�N����Loggin
            '    Dim lockInfo As String
            '    lockInfo = doWriteLockInfo(ocon, tran)
            '    MessageUtil.ShowErrorMsg("���݁A�ʂ̃��[�U�[�ɂ���ăf�[�^���g�p���ł��B" & vbCrLf & lockInfo)
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
#Region "�f�[�^�擾(DataTable)"
    ''' <summary>
    ''' DataRow�I�u�W�F�N�g�Ńf�[�^���擾����
    ''' </summary>
    ''' <param name="ocon">sql�R�l�N�V����</param>
    ''' <param name="SQL">sql��(</param>
    ''' <param name="tran">sql�g�����U�N�V����</param>
    ''' <returns>DataRow�I�u�W�F�N�g</returns>
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


#End Region


#Region "���g�p�̂���Private�ɕύX"
#Region "�V���O���R�[�e�[�V�����ǉ�"
    '------------------------------------------------------------------
    ' �@�\         : �V���O���R�[�e�[�V�����ǉ�
    '
    ' �Ԃ�l       : ����I�� = �ϊ���̕�����
    '                �ُ�I�� = ""
    '
    ' ������       : (IN) strVal  ���͕�����
    '
    ' �@�\����     : �����񒆂̃V���O���R�[�e�[�V�������������V���O���R�[�e�[�V�������d�ɂ���
    '
    ' ���l         : 2005.11.07 ��`
    '
    '------------------------------------------------------------------
    Private Function PBFSTR_ChangeQuotation(ByVal strVal As String) As String

        Dim intLocation As Integer        '�V���O���R�[�e�[�V�����̈ʒu
        Dim strOutputVal As String
        Dim strInputVal As String

        ''�o�͒l�ϐ���������
        strOutputVal = ""

        ''���͒l����͒l�ϐ��Ɉڑ�
        strInputVal = strVal

        ''���͒l�ϐ��̃V���O���R�[�e�[�V�����̈ʒu������
        intLocation = InStr(strInputVal, "'")

        ''�V���O���R�[�e�[�V�����̈ʒu���O���傫���ԃ��[�v����B
        While intLocation > 0
            ''�o�͒l�ϐ��ɓ��͒l�ϐ��̃V���O���R�[�e�[�V�����̈ʒu�܂łƃV���O���R�[�e�[�V�������o�́B
            strOutputVal = strOutputVal & Left$(strInputVal, intLocation) & "'"
            ''���͒l�ϐ�����o�͒l�ϐ��ɏo�͂�����������폜����B
            strInputVal = Mid$(strInputVal, intLocation + 1, Len(strInputVal) - intLocation)
            ''���͒l�ϐ��̃V���O���R�[�e�[�V�����̈ʒu������
            intLocation = InStr(strInputVal, "'")
            ''���[�v�I��
        End While

        ''�߂�l��ݒ�
        PBFSTR_ChangeQuotation = strOutputVal & strInputVal
    End Function
#End Region

#Region "SQL�\�z(AtoZ)"
    ' ------------------------------------------------------------------ 
    ' @(e) 
    ' 
    ' �@�\        : PBFSTR_CreatSqlAtoZ
    ' 
    ' �Ԃ�l      : String()
    ' 
    ' ������      : strCDST�F�J�n����
    '               strCDED�F�I������
    '               strFLD�F
    '               strWhere�F
    '
    ' �@�\����    : �J�n�`�I����SQL���\�z
    ' ���l        : 
    '               
    ''------------------------------------------------------------------
    Private Function PBFSTR_CreatSqlAtoZ(ByVal strCDST As String, ByVal strCDED As String, _
                                        ByVal strFLD As String, Optional ByVal strWhere As String = "") As String

        Dim strResult As String

        If strCDST = "" And strCDED = "" Then '�J�n�I���u�����N
            strResult = ""

        ElseIf strCDST <> "" And strCDED = "" Then  ''�J�n�̂�
            strResult = strFLD & ">=" & strCDST

        ElseIf strCDST = "" And strCDED <> "" Then ''�I���̂�
            strResult = strFLD & "<=" & strCDED

        Else
            strResult = strFLD & ">=" & strCDST & " AND " & strFLD & "<=" & strCDED

        End If

        Return strResult
    End Function
#End Region

#Region "SQL�\�z(�S�p���p����ʂ��Ȃ�)"
    '------------------------------------------------------------------------
    '(����)�@
    'strFLD         �t�B�[���h��
    'strText        �����l
    'inKBN          0:���Ԉ�v�@1:�O����v  2:�����v  (OPT=0)
    '------------------------------------------------------------------------
    Private Function PBFSTR_SQLMltSgl(ByVal strFLD As String, ByVal strText As String, _
                                    Optional ByVal inKBN As Integer = 0) As String
        Dim strSQL As String = ""
        Select Case inKBN

            Case 0 '���Ԉ�v����
                ''�S�p
                strSQL = strSQL & " ( " & strFLD & " LIKE  ('%" & StrConv(strText, VbStrConv.Wide) & "%')"
                ''���p
                strSQL = strSQL & " OR " & strFLD & " LIKE  ('%" & StrConv(strText, VbStrConv.Narrow) & "%'))"

            Case 1 '�O����v����
                ''�S�p
                strSQL = strSQL & " ( " & strFLD & " LIKE  ('" & StrConv(strText, VbStrConv.Wide) & "%')"
                ''���p
                strSQL = strSQL & " OR " & strFLD & " LIKE  ('" & StrConv(strText, VbStrConv.Narrow) & "%'))"

            Case 2 '�����v����
                ''�S�p
                strSQL = strSQL & " ( " & strFLD & " LIKE  ('%" & StrConv(strText, VbStrConv.Wide) & "')"
                ''���p
                strSQL = strSQL & " OR " & strFLD & " LIKE  ('%" & StrConv(strText, VbStrConv.Narrow) & "'))"
        End Select
        Return strSQL
    End Function
#End Region
#End Region


End Module

