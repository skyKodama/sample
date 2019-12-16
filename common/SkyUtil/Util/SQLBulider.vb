#Region "�錾"
Option Explicit On 
Option Strict On

Imports System
Imports System.Text
Imports skysystem.common.SystemUtil
#End Region

'*************************************************************************
'*�@�@�\�@�@�FSQL��Builder�N���X
'*
'*�@�쐬���@�F2006.06.05    ���
'*
'*�@���ύX���e��
'*
'*************************************************************************
''' <summary>
''' SQL��Builder�N���X
''' </summary>
''' <remarks></remarks>
Public Class SQLBulider


#Region "Pirvate�ϐ�"
    Private _commandText As String
    Private _table As String

    Private fields As New ArrayList
    Private parameters As New ArrayList
    Private PrszWhere As String
#End Region



#Region "��Poroperty��"
    Public ReadOnly Property IsEmpty() As Boolean
        Get
            If fields.Count = 0 Or parameters.Count = 0 Then
                Return True
            End If
            Return False
        End Get
    End Property

    ''' <summary>
    ''' �������ꂽSQL�����擾
    ''' </summary>
    ''' <value>SQL��</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CommandText() As String
        Get
            Return _commandText
        End Get
    End Property

    ''' <summary>
    ''' �e�[�u�������`
    ''' </summary>
    ''' <value>SQL�F�e�[�u����</value>
    ''' <remarks></remarks>
    Public WriteOnly Property Table() As String
        Set(ByVal Value As String)
            _table = Value
        End Set
    End Property
#End Region


#Region "Private���b�\�h"
    Private Function createStatementIns() As String
        Dim stringBuffer As New StringBuilder
        stringBuffer.Append("INSERT INTO ")
        stringBuffer.Append(_table)
        stringBuffer.Append(" (")
        stringBuffer.Append(creatorIns(fields))
        stringBuffer.Append(") VALUES (")
        stringBuffer.Append(creatorIns(parameters))
        stringBuffer.Append(")")
        Return stringBuffer.ToString
    End Function
    Private Function createStatementUpd() As String
        Dim stringBuffer As New StringBuilder
        stringBuffer.Append("UPDATE  ")
        stringBuffer.Append(_table)
        stringBuffer.Append(" SET ")
        stringBuffer.Append(creatorUpd(fields, parameters))
        stringBuffer.Append(" ")
        stringBuffer.Append(PrszWhere) 'WHERE��
        Return stringBuffer.ToString
    End Function

    Private Sub arrayClear()
        fields.Clear()
        parameters.Clear()
        PrszWhere = ""
    End Sub

    Private Function decorateString(ByVal decorateBase As String) As String
        Return "'" & decorateBase & "'"
    End Function
    Private Function creatorUpd(ByVal iterateF As ArrayList, ByVal iterateP As ArrayList) As String
        Dim stringBuffer As New StringBuilder
        Dim i As Integer

        For i = 0 To iterateF.Count - 1
            stringBuffer.Append(iterateF.Item(i))
            stringBuffer.Append("=")
            stringBuffer.Append(iterateP.Item(i))
            stringBuffer.Append(", ")
        Next

        stringBuffer.Remove(stringBuffer.Length - 2, 2)
        Return stringBuffer.ToString
    End Function
    Private Function creatorIns(ByVal iterate As ArrayList) As String
        Dim stringBuffer As New StringBuilder
        Dim iterator As IEnumerator = iterate.GetEnumerator
        While iterator.MoveNext
            stringBuffer.Append(iterator.Current)
            stringBuffer.Append(", ")
        End While
        stringBuffer.Remove(stringBuffer.Length - 2, 2)
        Return stringBuffer.ToString
    End Function
#End Region


#Region "Public���b�\�h"
    ''' <summary>
    ''' Add���\�b�h
    ''' </summary>
    ''' <param name="fieldElement">�t�B�[���h��</param>
    ''' <param name="parameterElement">�l(String�^)</param>
    ''' <param name="bytDeco">�N�H�[�e�[�V�����t�^���邩�ǂ���(�f�t�H���g�͕t�^����)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As String, _
                             Optional ByVal bytDeco As Byte = 9)

        '*** Field�����
        fields.Add(fieldElement)

        '*** Parameter�l���
        If PB_ChkNUll(parameterElement) Then
            parameters.Add("NULL")                  '�k���̏ꍇ'NULL'���̂܂�
        Else
            If bytDeco <> 9 Then
                parameters.Add(parameterElement)    ''SYSDATE'�̏ꍇ��Oracle�ɂ��̂܂�
            Else
                parameters.Add(decorateString(parameterElement.Replace("'", "''")))
            End If
        End If
    End Sub
    ''' <summary>
    ''' Add���\�b�h
    ''' </summary>
    ''' <param name="fieldElement">�t�B�[���h��</param>
    ''' <param name="parameterElement">�l(Boolean�^)</param>
    ''' <remarks>20131009_1 </remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Boolean)
        fields.Add(fieldElement)

        If parameterElement Then
            parameters.Add(PBCint(CHK.TRUE))
        Else
            parameters.Add(PBCint(CHK.FALSE))
        End If

    End Sub
    ''' <summary>
    ''' Add���\�b�h
    ''' </summary>
    ''' <param name="fieldElement">�t�B�[���h��</param>
    ''' <param name="parameterElement">�l(Integer�^)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Integer)
        fields.Add(fieldElement)
        parameters.Add(parameterElement)
    End Sub
    ''' <summary>
    ''' Add���\�b�h
    ''' </summary>
    ''' <param name="fieldElement">�t�B�[���h��</param>
    ''' <param name="parameterElement">�l(Byte�^)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Byte)
        fields.Add(fieldElement)
        parameters.Add(parameterElement)
    End Sub
    ''' <summary>
    ''' Add���\�b�h
    ''' </summary>
    ''' <param name="fieldElement">�t�B�[���h��</param>
    ''' <param name="parameterElement">�l(Long�^)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Long)
        fields.Add(fieldElement)
        parameters.Add(parameterElement)
    End Sub
    ''' <summary>
    ''' Add���\�b�h
    ''' </summary>
    ''' <param name="fieldElement">�t�B�[���h��</param>
    ''' <param name="parameterElement">�l(Decimal�^)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Decimal)
        fields.Add(fieldElement)
        parameters.Add(parameterElement)
    End Sub
    ''' <summary>
    ''' Add���\�b�h
    ''' </summary>
    ''' <param name="fieldElement">�t�B�[���h��</param>
    ''' <param name="parameterElement">�l(Date�^)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Date)
        fields.Add(fieldElement)
        parameters.Add(decorateString(parameterElement.ToString))
    End Sub
    ''' <summary>
    ''' Add���\�b�h
    ''' </summary>
    ''' <param name="fieldElement">�t�B�[���h��</param>
    ''' <param name="parameterElement">�l(Double�^)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Double)
        fields.Add(fieldElement)
        parameters.Add(decorateString(parameterElement.ToString))
    End Sub
    ''' <summary>
    ''' Where��̒ǉ�
    ''' </summary>
    ''' <param name="szWhere">SQL�FWhere��</param>
    ''' <remarks></remarks>
    Public Overloads Sub addWhere(ByVal szWhere As String)
        PrszWhere = szWhere
    End Sub
    Public Sub clear()
        arrayClear()
        _commandText = Nothing
    End Sub

    ''' <summary>
    ''' SQL�C���T�[�g���𐶐�����
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub buildIns()
        _commandText = createStatementIns()
        arrayClear()
    End Sub
    ''' <summary>
    ''' SQL�A�b�v�f�[�g���𐶐�����
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub buildUpd()
        _commandText = createStatementUpd()
        arrayClear()
    End Sub
#End Region


End Class

