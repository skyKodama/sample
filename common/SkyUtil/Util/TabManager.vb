''' <summary>
'''  タブコントロール制御クラス
''' </summary>
''' <remarks></remarks>
Public Class TabManager
    '********************************************************************
    '* ソースファイル名 : TabController.vb
    '* クラス名　　	    : TabController
    '* クラス説明　	    : タブページ・表示非表示
    '* 備考　           :
    '* 作成  　         : 2008/5/1
    '* 更新履歴         :
    '********************************************************************
    Private Class TabPageInfo
        Public TabPage As TabPage
        Public Visible As Boolean

        Public Sub New(ByVal page As TabPage, ByVal v As Boolean)
            TabPage = page
            Visible = v
        End Sub
    End Class

    Private _tabPageInfos As TabPageInfo() = Nothing
    Private _tabControl As TabControl = Nothing

    ''' <summary>
    ''' TabPageManagerクラスのインスタンスを作成する
    ''' </summary>
    ''' <param name="crl">基になるTabControlオブジェクト</param>
    Public Sub New(ByVal crl As TabControl)
        _tabControl = crl
        _tabPageInfos = _
            New TabPageInfo(_tabControl.TabPages.Count - 1) {}
        Dim i As Integer
        For i = 0 To _tabControl.TabPages.Count - 1
            '配色を設定
            _tabControl.TabPages(i).BackColor = System.Drawing.SystemColors.Control
            _tabPageInfos(i) = New TabPageInfo(_tabControl.TabPages(i), True)
        Next i

        'DrawItemイベントハンドラを追加
        crl.DrawMode = TabDrawMode.OwnerDrawFixed
        AddHandler crl.DrawItem, AddressOf TabControl_DrawItem

    End Sub

    ''' <summary>
    ''' TabPageの表示・非表示を変更する
    ''' </summary>
    ''' <param name="index">変更するTabPageのIndex番号</param>
    ''' <param name="v">表示するときはTrue。
    ''' 非表示にするときはFalse。</param>
    Public Sub ChangeTabPageVisible( _
        ByVal index As Integer, ByVal v As Boolean)
        If _tabPageInfos(index).Visible = v Then
            Return
        End If
        _tabPageInfos(index).Visible = v
        _tabControl.SuspendLayout()
        _tabControl.TabPages.Clear()
        Dim i As Integer
        For i = 0 To _tabPageInfos.Length - 1
            If _tabPageInfos(i).Visible Then
                _tabControl.TabPages.Add(_tabPageInfos(i).TabPage)
            End If
        Next i
        _tabControl.ResumeLayout()
    End Sub


#Region "イベント"
    Private Sub TabControl_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        '対象のTabControlを取得
        Dim tab As TabControl = CType(sender, TabControl)
        'タブページのテキストを取得
        Dim txt As String = tab.TabPages(e.Index).Text

        'タブのテキストと背景を描画するためのブラシを決定する
        Dim foreBrush, backBrush As Brush
        If e.State = DrawItemState.Selected Then
            '選択されているタブのテキスト・背景を設定
            foreBrush = Brushes.DarkBlue
            'backBrush = Brushes.Lime
            backBrush = Brushes.LightGreen

        Else
            '選択されていないタブのテキストは灰色、背景を白とする
            foreBrush = Brushes.Gray
            backBrush = Brushes.White

        End If


        'StringFormatを作成
        Dim sf As New StringFormat
        '中央に表示する
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center

        '背景の描画
        e.Graphics.FillRectangle(backBrush, e.Bounds)
        'Textの描画
        Dim rectf As New RectangleF( _
            e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height + 1)
        e.Graphics.DrawString(txt, e.Font, foreBrush, rectf, sf)

    End Sub

#End Region
End Class

