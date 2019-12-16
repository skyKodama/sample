Option Explicit On
Option Strict On

Imports System.Text
'Imports IG = Infragistics.Win.UltraWinEditors
'Imports IGM = Infragistics.Win.Misc
'Imports IGMASK = Infragistics.Win.UltraWinMaskedEdit
Imports skysystem.common.SystemUtil

Public Class dailogUtil
    '*************************************************************************
    '*　機能　　：共通初期化・dailogUtil
    '*　作成日　：2013/12/01   駒方
    '*
    '*  変更日  ：
    '*
    '*************************************************************************

    ''取得したファイルパス
    Private rtnFile As String = ""
    Private filtertext As String = ""
    Private filtertp As FILEKIND = FILEKIND.XLS
    Private title As String = "ファイルを選択してください"
    Public Property _title As String
        Get
            Return title
        End Get
        Set(value As String)
            title = value
        End Set
    End Property
    Public Property _filtertp As FILEKIND
        Get
            Return filtertp
        End Get
        Set(value As FILEKIND)
            filtertp = value
        End Set
    End Property
    Public Property _filtertext As String
        Get
            Return filtertext
        End Get
        Set(value As String)
            filtertext = value
        End Set
    End Property
    Public ReadOnly Property _filePath As String
        Get
            Return rtnFile
        End Get
    End Property

    ''' <summary>
    ''' ファイルOpenダイアログボックスの表示
    ''' </summary>
    ''' <param name="szDefaultPath">初期ディレクトリ(省略時はC:\)</param>
    ''' <param name="title">フォームタイトル</param>
    ''' <returns>ファイル名</returns>
    ''' <remarks>キャンセル時は空値を戻す</remarks>
    Public Function OpenFileDialog(Optional ByVal szDefaultPath As String = "", _
                                         Optional ByVal title As String = "") As String

        Dim sfd As New OpenFileDialog
        'はじめのファイル名を指定する
        'sfd.FileName = szDefaltFileName
        'はじめに表示されるフォルダを指定する
        '20080201_1 Add
        If Not szDefaultPath.Equals("") Then
            sfd.InitialDirectory = szDefaultPath
        End If

        ''If szDefaultPath.Equals("") Then
        ''    sfd.InitialDirectory = "C:\"
        ''Else
        ''    sfd.InitialDirectory = szDefaultPath
        ''End If

        '[ファイルの種類]に表示される選択肢を指定する
        Select Case filtertp
            Case FILEKIND.XLS
                sfd.Filter = "Excelファイル (*.xls)|*.xls"
            Case FILEKIND.TXT
                sfd.Filter = "テキストファイル (*.txt)|*.txt"
            Case FILEKIND.CSV
                sfd.Filter = "CSVファイル (*.csv)|*.csv"
            Case FILEKIND.ETC
                sfd.Filter = filtertext
        End Select

        '[ファイルの種類]ではじめに
        '「すべてのファイル」が選択されているようにする
        sfd.FilterIndex = 2
        'タイトルを設定する
        If title.Equals("") Then
            sfd.Title = "保存先のファイルを選択してください"
        Else
            sfd.Title = title
        End If

        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        sfd.RestoreDirectory = True
        '存在しないパスが指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        sfd.CheckPathExists = True

        'ダイアログを表示する
        If sfd.ShowDialog() = DialogResult.OK Then
            rtnFile = sfd.FileName
            Return sfd.FileName
        Else
            Return ""
        End If
    End Function
End Class
