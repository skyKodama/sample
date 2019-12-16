Option Explicit On
Option Strict On
Imports skysystem.common
Imports skysystem.common.MessageUtil
Imports skysystem.common.SystemUtil
Imports Infragistics.Win
Imports Devart.Data.Universal

Public Class comboUtil




   

    ''' <summary>
    '''  コンボボックスバインド処理
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="strTable"></param>
    ''' <param name="valueMember"></param>
    ''' <param name="displayMember"></param>
    ''' <param name="strWhere"></param>
    ''' <param name="szOrder"></param>
    ''' <param name="blnSpace"></param>
    ''' <param name="tran"></param>
    ''' <remarks>SkyBaseComboを使用しているため、現在未使用</remarks>
    Public Shared Sub BindCmbData(targetCombo As Infragistics.Win.UltraWinEditors.UltraComboEditor, ByVal con As uniConnection, ByVal strTable As String, _
                         ByVal valueMember As String, ByVal displayMember As String, _
                         Optional ByVal strWhere As String = "", Optional ByVal szOrder As String = Nothing, _
                         Optional ByVal blnSpace As Boolean = True, Optional ByVal tran As uniTransaction = Nothing)

        Dim strSql As String
        Dim oda As UniDataAdapter
        Dim ocd As UniCommand
        Dim dts As DataSet

        Try


            'blnFlgInit = False   'セットコンボ初期化フラグ(False)
            Dim strAS As String = "NM" 'As句


            'SQL文作成
            strSql = ""
            strSql = strSql & " SELECT " & valueMember & "," & displayMember & " AS " & strAS
            strSql = strSql & " FROM " & strTable & " "
            If strWhere <> "" Then
                strSql = strSql & " WHERE " & strWhere
            End If

            If szOrder = "" Then
                strSql = strSql & " ORDER BY " & valueMember
            Else
                strSql = strSql & " ORDER BY " & szOrder
            End If

            'コマンド作成
            oda = New UniDataAdapter
            dts = New DataSet

            ocd = New UniCommand(strSql, con)

            oda.SelectCommand = ocd
            oda.Fill(dts, strTable)

            '** ADD 2006.07.05
            targetCombo.Items.Clear()  'ｾｯﾄする前初期化


            If dts.Tables(0).Rows.Count > 0 Then

                '初期化
                targetCombo.DataSource = Nothing
                targetCombo.Items.Clear()


                Dim dtTbl_Clone As DataTable = dts.Tables(0).Copy()

                'ValuMenbの最大文字数を取得
                Dim iMaxLength As Integer = 0
                For i As Integer = 0 To dtTbl_Clone.Rows.Count - 1
                    If PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)).Length > iMaxLength Then
                        iMaxLength = PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)).Length
                    End If
                Next

                '配列クラスに格納
                Dim cmbAry As New ArrayList

                '空項目を追加するか？
                If blnSpace Then
                    cmbAry.Add(New cConf("", "", 0))
                End If

                For i As Integer = 0 To dtTbl_Clone.Rows.Count - 1
                    cmbAry.Add(New cConf(PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)), _
                                                            PBCStr(dtTbl_Clone.Rows.Item(i)(strAS)), iMaxLength))
                Next

                'データソース設定
                targetCombo.DataSource = cmbAry
                targetCombo.ValueMember = "ValueData"
                targetCombo.DisplayMember = "DisplayDataEdtingVertical"
                'targetCombo.ValueList.DropDownResizeHandleStyle = Infragistics.Win.DropDownResizeHandleStyle.VerticalResize
                '自動調整 20071025_1
                'targetCombo.AutoSize = True
                '候補取得
                'targetCombo.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Append

            End If

        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Sub


    Class cConf
        '*********************************************************************************************
        '*　機能　　：配列格納クラス
        '*            
        '*　
        '*********************************************************************************************
        Private DisMeb As String
        Private DescMeb As String
        Private ValMeb As String
        Private pMaxLength As Integer

#Region "コンストラクタ"
        Sub New(ByVal prmDisMeb As String, ByVal prmDescMeb As String, ByVal prmValMeb As String)
            MyBase.New()
            Me.DisMeb = prmDisMeb
            Me.DescMeb = prmDescMeb
            Me.ValMeb = prmValMeb
        End Sub
        Sub New(ByVal prmValMeb As String, ByVal prmDisMeb As String, ByVal prmMaxLengeth As Integer)
            MyBase.New()
            Me.DisMeb = prmDisMeb
            Me.ValMeb = prmValMeb
            Me.pMaxLength = prmMaxLengeth
        End Sub
#End Region

#Region "プロパティ"
#Region "DisplayMemberEditing "
        ReadOnly Property DisplayDataEdting() As String
            Get
                If Me.ValMeb.Equals("") Then
                    Return ""
                Else
                    Return Me.DisMeb
                End If
            End Get
        End Property
#End Region
#Region "DisplayDataEdtingVertical "
        ReadOnly Property DisplayDataEdtingVertical() As String
            Get
                If Me.ValMeb.Equals("") Then
                    Return ""
                Else
                    Return Me.ValMeb.PadRight(pMaxLength) & "｜" & DisMeb
                End If
            End Get
        End Property
#End Region
#Region "DisplayMember "
        ReadOnly Property DisplayData() As String
            Get
                Return DisMeb
            End Get
        End Property
#End Region
#Region "DescriptionMember "
        ReadOnly Property DescriptionData() As String
            Get
                Return DescMeb
            End Get
        End Property
#End Region
#Region "ValueMember"
        ReadOnly Property ValueData() As String
            Get
                Return ValMeb
            End Get
        End Property
#End Region
#Region "ImageMenber"
        'Public ReadOnly Property ImageData() As Object
        '    Get
        '        Return aImageData
        '    End Get
        'End Property
#End Region
#End Region

    End Class
End Class

