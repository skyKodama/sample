''' <summary>
''' �e�L�X�g�����N���X
''' </summary>
''' <remarks>�g�p�s��</remarks>
Public Class txtController

    Private PrdtTbl As New DataTable
    Private outFile As String '�o�͐�p�X�{�t�@�C��
    Private Const PGMID As String = "M_MA9000"
    Private Const PrszFileName = "LogFile"
    Private boExist As Boolean = True


#Region "�R���X�g���N�g"
    Sub New(ByVal dttbl As DataTable)
        Me.PrdtTbl = dttbl

    End Sub
#End Region

#Region "�o�͏����J�n"
    Private Function WriteTxtFile() As Boolean
        Try
            Dim inRowCnt As Integer
            Dim inColCnt As Integer = PrdtTbl.Columns.Count
            Dim Sw As New System.IO.StreamWriter(outFile, False, System.Text.Encoding.Default)

            '��������
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

#Region "�e�L�X�g�o��"
    Friend Function txtOutAction() As Boolean

        'Dim outPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) & "\" '�}�C�h�L�������g
        'Dim FileName As String = PrszFileName & Date.Now.ToString("yyMMdd")

        'Try

        '    '*-----------------
        '    '�f�[�^���݃`�F�b�N
        '    '*-----------------
        '    If PrdtTbl Is Nothing OrElse PrdtTbl.Rows.Count = 0 Then
        '        PBS_ShowErrorMsg("�o�͂���f�[�^�����݂��܂���B")
        '        Exit Function
        '    End If

        '    '*-----------------
        '    '�޲�۸ރ{�b�N�X
        '    '*-----------------
        '    If Not FileDialog(FileName, outPath, ReadFilePath(PBFSTR_GetRstDir, "/SKY/" & PGMID, outPath), LibraryPB.FILEKIND.TXT) Then
        '        Return False
        '    Else
        '        '�ۑ�����Z�[�u 20070315_1
        '        PBS_SavePath(PGMID, outPath)
        '    End If

        '    '*-----------------
        '    '�o�͐�t�@�C�������Z�b�g
        '    '*-----------------
        '    outFile = outPath & FileName

        '    '*-----------------
        '    ' �o�͏����J�n
        '    '*-----------------
        '    WriteTxtFile()

        '    ShowInfoMsg("���O�t�@�C�����o�͂��܂����B")

        '    Return True


        'Catch ex As Exception
        '    PBS_ShowErrorMsg("�o�͏����Ɏ��s���܂����B")
        '    Exit Function
        'End Try

    End Function
#End Region

#Region "�捞����(CSV)"
    Public Shared Function ReadCsvFile(ByVal prmPath As String) As ArrayList

        Dim csvRecords As New System.Collections.ArrayList()

        'CSV�t�@�C����
        Dim csvFileName As String = prmPath

        'Shift JIS�œǂݍ���
        Dim tfp As New FileIO.TextFieldParser(csvFileName, _
            System.Text.Encoding.GetEncoding(932))
        '�t�B�[���h�������ŋ�؂��Ă���Ƃ���
        '�f�t�H���g��Delimited�Ȃ̂ŁA�K�v�Ȃ�

        Try
            tfp.TextFieldType = FileIO.FieldType.Delimited
            '��؂蕶����,�Ƃ���
            tfp.Delimiters = New String() {","}
            '�t�B�[���h��"�ň͂݁A���s�����A��؂蕶�����܂߂邱�Ƃ��ł��邩
            '�f�t�H���g��true�Ȃ̂ŁA�K�v�Ȃ�
            tfp.HasFieldsEnclosedInQuotes = True
            '�t�B�[���h�̑O�ォ��X�y�[�X���폜����
            '�f�t�H���g��true�Ȃ̂ŁA�K�v�Ȃ�
            tfp.TrimWhiteSpace = True

            While Not tfp.EndOfData
                '�t�B�[���h��ǂݍ���
                Dim fields As String() = tfp.ReadFields()
                '�ۑ�
                csvRecords.Add(fields)
            End While


            Return csvRecords

        Catch ex As Exception
            Throw ex
        Finally
            '��n��
            tfp.Close()

        End Try

    End Function

#End Region '20100323_1

End Class
