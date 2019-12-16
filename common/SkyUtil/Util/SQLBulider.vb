#Region "宣言"
Option Explicit On 
Option Strict On

Imports System
Imports System.Text
Imports skysystem.common.SystemUtil
#End Region

'*************************************************************************
'*　機能　　：SQL文Builderクラス
'*
'*　作成日　：2006.06.05    樋口
'*
'*　＜変更内容＞
'*
'*************************************************************************
''' <summary>
''' SQL文Builderクラス
''' </summary>
''' <remarks></remarks>
Public Class SQLBulider


#Region "Pirvate変数"
    Private _commandText As String
    Private _table As String

    Private fields As New ArrayList
    Private parameters As New ArrayList
    Private PrszWhere As String
#End Region



#Region "■Poroperty■"
    Public ReadOnly Property IsEmpty() As Boolean
        Get
            If fields.Count = 0 Or parameters.Count = 0 Then
                Return True
            End If
            Return False
        End Get
    End Property

    ''' <summary>
    ''' 生成されたSQL文を取得
    ''' </summary>
    ''' <value>SQL文</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CommandText() As String
        Get
            Return _commandText
        End Get
    End Property

    ''' <summary>
    ''' テーブル名を定義
    ''' </summary>
    ''' <value>SQL：テーブル名</value>
    ''' <remarks></remarks>
    Public WriteOnly Property Table() As String
        Set(ByVal Value As String)
            _table = Value
        End Set
    End Property
#End Region


#Region "Privateメッソド"
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
        stringBuffer.Append(PrszWhere) 'WHERE句
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


#Region "Publicメッソド"
    ''' <summary>
    ''' Addメソッド
    ''' </summary>
    ''' <param name="fieldElement">フィールド名</param>
    ''' <param name="parameterElement">値(String型)</param>
    ''' <param name="bytDeco">クォーテーション付与するかどうか(デフォルトは付与する)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As String, _
                             Optional ByVal bytDeco As Byte = 9)

        '*** Field名ｾｯﾄ
        fields.Add(fieldElement)

        '*** Parameter値ｾｯﾄ
        If PB_ChkNUll(parameterElement) Then
            parameters.Add("NULL")                  'ヌルの場合'NULL'そのまま
        Else
            If bytDeco <> 9 Then
                parameters.Add(parameterElement)    ''SYSDATE'の場合はOracleにそのまま
            Else
                parameters.Add(decorateString(parameterElement.Replace("'", "''")))
            End If
        End If
    End Sub
    ''' <summary>
    ''' Addメソッド
    ''' </summary>
    ''' <param name="fieldElement">フィールド名</param>
    ''' <param name="parameterElement">値(Boolean型)</param>
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
    ''' Addメソッド
    ''' </summary>
    ''' <param name="fieldElement">フィールド名</param>
    ''' <param name="parameterElement">値(Integer型)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Integer)
        fields.Add(fieldElement)
        parameters.Add(parameterElement)
    End Sub
    ''' <summary>
    ''' Addメソッド
    ''' </summary>
    ''' <param name="fieldElement">フィールド名</param>
    ''' <param name="parameterElement">値(Byte型)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Byte)
        fields.Add(fieldElement)
        parameters.Add(parameterElement)
    End Sub
    ''' <summary>
    ''' Addメソッド
    ''' </summary>
    ''' <param name="fieldElement">フィールド名</param>
    ''' <param name="parameterElement">値(Long型)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Long)
        fields.Add(fieldElement)
        parameters.Add(parameterElement)
    End Sub
    ''' <summary>
    ''' Addメソッド
    ''' </summary>
    ''' <param name="fieldElement">フィールド名</param>
    ''' <param name="parameterElement">値(Decimal型)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Decimal)
        fields.Add(fieldElement)
        parameters.Add(parameterElement)
    End Sub
    ''' <summary>
    ''' Addメソッド
    ''' </summary>
    ''' <param name="fieldElement">フィールド名</param>
    ''' <param name="parameterElement">値(Date型)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Date)
        fields.Add(fieldElement)
        parameters.Add(decorateString(parameterElement.ToString))
    End Sub
    ''' <summary>
    ''' Addメソッド
    ''' </summary>
    ''' <param name="fieldElement">フィールド名</param>
    ''' <param name="parameterElement">値(Double型)</param>
    ''' <remarks></remarks>
    Public Overloads Sub add(ByVal fieldElement As Object, ByVal parameterElement As Double)
        fields.Add(fieldElement)
        parameters.Add(decorateString(parameterElement.ToString))
    End Sub
    ''' <summary>
    ''' Where句の追加
    ''' </summary>
    ''' <param name="szWhere">SQL：Where句</param>
    ''' <remarks></remarks>
    Public Overloads Sub addWhere(ByVal szWhere As String)
        PrszWhere = szWhere
    End Sub
    Public Sub clear()
        arrayClear()
        _commandText = Nothing
    End Sub

    ''' <summary>
    ''' SQLインサート文を生成する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub buildIns()
        _commandText = createStatementIns()
        arrayClear()
    End Sub
    ''' <summary>
    ''' SQLアップデート文を生成する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub buildUpd()
        _commandText = createStatementUpd()
        arrayClear()
    End Sub
#End Region


End Class

