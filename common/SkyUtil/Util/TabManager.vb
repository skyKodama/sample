''' <summary>
'''  �^�u�R���g���[������N���X
''' </summary>
''' <remarks></remarks>
Public Class TabManager
    '********************************************************************
    '* �\�[�X�t�@�C���� : TabController.vb
    '* �N���X���@�@	    : TabController
    '* �N���X�����@	    : �^�u�y�[�W�E�\����\��
    '* ���l�@           :
    '* �쐬  �@         : 2008/5/1
    '* �X�V����         :
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
    ''' TabPageManager�N���X�̃C���X�^���X���쐬����
    ''' </summary>
    ''' <param name="crl">��ɂȂ�TabControl�I�u�W�F�N�g</param>
    Public Sub New(ByVal crl As TabControl)
        _tabControl = crl
        _tabPageInfos = _
            New TabPageInfo(_tabControl.TabPages.Count - 1) {}
        Dim i As Integer
        For i = 0 To _tabControl.TabPages.Count - 1
            '�z�F��ݒ�
            _tabControl.TabPages(i).BackColor = System.Drawing.SystemColors.Control
            _tabPageInfos(i) = New TabPageInfo(_tabControl.TabPages(i), True)
        Next i

        'DrawItem�C�x���g�n���h����ǉ�
        crl.DrawMode = TabDrawMode.OwnerDrawFixed
        AddHandler crl.DrawItem, AddressOf TabControl_DrawItem

    End Sub

    ''' <summary>
    ''' TabPage�̕\���E��\����ύX����
    ''' </summary>
    ''' <param name="index">�ύX����TabPage��Index�ԍ�</param>
    ''' <param name="v">�\������Ƃ���True�B
    ''' ��\���ɂ���Ƃ���False�B</param>
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


#Region "�C�x���g"
    Private Sub TabControl_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        '�Ώۂ�TabControl���擾
        Dim tab As TabControl = CType(sender, TabControl)
        '�^�u�y�[�W�̃e�L�X�g���擾
        Dim txt As String = tab.TabPages(e.Index).Text

        '�^�u�̃e�L�X�g�Ɣw�i��`�悷�邽�߂̃u���V�����肷��
        Dim foreBrush, backBrush As Brush
        If e.State = DrawItemState.Selected Then
            '�I������Ă���^�u�̃e�L�X�g�E�w�i��ݒ�
            foreBrush = Brushes.DarkBlue
            'backBrush = Brushes.Lime
            backBrush = Brushes.LightGreen

        Else
            '�I������Ă��Ȃ��^�u�̃e�L�X�g�͊D�F�A�w�i�𔒂Ƃ���
            foreBrush = Brushes.Gray
            backBrush = Brushes.White

        End If


        'StringFormat���쐬
        Dim sf As New StringFormat
        '�����ɕ\������
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center

        '�w�i�̕`��
        e.Graphics.FillRectangle(backBrush, e.Bounds)
        'Text�̕`��
        Dim rectf As New RectangleF( _
            e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height + 1)
        e.Graphics.DrawString(txt, e.Font, foreBrush, rectf, sf)

    End Sub

#End Region
End Class

