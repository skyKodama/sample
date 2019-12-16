
Option Explicit On
Option Strict On
Imports Npgsql
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
Public Class DBconnectionN

#Region "Private�ϐ�"
    Private Con As New NpgsqlConnection             '�I���N���R�l�N�V����
#End Region

#Region "�R���X�g���N�^"
    Public Sub New()

    End Sub
#End Region



#Region "RtnCon�F�ڑ���Ԃ�"
    ''' <summary>
    ''' �ڑ�����߂�
    ''' </summary>
    ''' <returns>Sql�R�l�N�V����</returns>
    ''' <remarks></remarks>
    Public Function rtncon() As NpgsqlConnection

        '�R�l�N�V�����������Ă���ꍇ�̂�
        Open()

        Return OraCon
    End Function
#End Region

#Region "Close�F�ڑ������"
    ''' <summary>
    ''' �ڑ������
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
