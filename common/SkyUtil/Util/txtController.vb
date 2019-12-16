''' <summary>
''' テキスト生成クラス
''' </summary>
''' <remarks>使用不可</remarks>
Public Class txtController

    Private PrdtTbl As New DataTable
    Private outFile As String '出力先パス＋ファイル
    Private Const PGMID As String = "M_MA9000"
    Private Const PrszFileName = "LogFile"
    Private boExist As Boolean = True


#Region "コンストラクト"
    Sub New(ByVal dttbl As DataTable)
        Me.PrdtTbl = dttbl

    End Sub
#End Region

#Region "出力処理開始"
    Private Function WriteTxtFile() As Boolean
        Try
            Dim inRowCnt As Integer
            Dim inColCnt As Integer = PrdtTbl.Columns.Count
            Dim Sw As New System.IO.StreamWriter(outFile, False, System.Text.Encoding.Default)

            '書込処理
            For inRowCnt = 0 To PrdtTbl.Rows.Count - 1
                Dim szValue As String = ""
                Dim i As Integer

                For i = 0 To inColCnt - 1
                    If i <> inColCnt - 1 Then
                        szValue += PrdtTbl.Rows.Item(inRowCnt)(i) & ","
                    Else
                        szValue += PrdtTbl.Rows.Item(inRowCnt)(i)
                    End If
                Next

                Sw.WriteLine(szValue)
            Next


            Sw.Close()

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "テキスト出力"
    Friend Function txtOutAction() As Boolean

        'Dim outPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) & "\" 'マイドキュメント
        'Dim FileName As String = PrszFileName & Date.Now.ToString("yyMMdd")

        'Try

        '    '*-----------------
        '    'データ存在チェック
        '    '*-----------------
        '    If PrdtTbl Is Nothing OrElse PrdtTbl.Rows.Count = 0 Then
        '        PBS_ShowErrorMsg("出力するデータが存在しません。")
        '        Exit Function
        '    End If

        '    '*-----------------
        '    'ﾀﾞｲｱﾛｸﾞボックス
        '    '*-----------------
        '    If Not FileDialog(FileName, outPath, ReadFilePath(PBFSTR_GetRstDir, "/SKY/" & PGMID, outPath), LibraryPB.FILEKIND.TXT) Then
        '        Return False
        '    Else
        '        '保存先をセーブ 20070315_1
        '        PBS_SavePath(PGMID, outPath)
        '    End If

        '    '*-----------------
        '    '出力先ファイル名をセット
        '    '*-----------------
        '    outFile = outPath & FileName

        '    '*-----------------
        '    ' 出力処理開始
        '    '*-----------------
        '    WriteTxtFile()

        '    ShowInfoMsg("ログファイルを出力しました。")

        '    Return True


        'Catch ex As Exception
        '    PBS_ShowErrorMsg("出力処理に失敗しました。")
        '    Exit Function
        'End Try

    End Function
#End Region

#Region "取込処理(CSV)"
    Public Shared Function ReadCsvFile(ByVal prmPath As String) As ArrayList

        Dim csvRecords As New System.Collections.ArrayList()

        'CSVファイル名
        Dim csvFileName As String = prmPath

        'Shift JISで読み込む
        Dim tfp As New FileIO.TextFieldParser(csvFileName, _
            System.Text.Encoding.GetEncoding(932))
        'フィールドが文字で区切られているとする
        'デフォルトでDelimitedなので、必要なし

        Try
            tfp.TextFieldType = FileIO.FieldType.Delimited
            '区切り文字を,とする
            tfp.Delimiters = New String() {","}
            'フィールドを"で囲み、改行文字、区切り文字を含めることができるか
            'デフォルトでtrueなので、必要なし
            tfp.HasFieldsEnclosedInQuotes = True
            'フィールドの前後からスペースを削除する
            'デフォルトでtrueなので、必要なし
            tfp.TrimWhiteSpace = True

            While Not tfp.EndOfData
                'フィールドを読み込む
                Dim fields As String() = tfp.ReadFields()
                '保存
                csvRecords.Add(fields)
            End While


            Return csvRecords

        Catch ex As Exception
            Throw ex
        Finally
            '後始末
            tfp.Close()

        End Try

    End Function

#End Region '20100323_1

End Class
