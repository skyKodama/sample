
Option Explicit On
Option Strict On
Imports System.IO
Imports System.Xml
Imports System.Text
Imports skysystem.common
Imports skysystem.common.SystemConst
Imports skysystem.common.MessageUtil
Imports System.Windows.Forms
Imports Devart.Data.Universal
Imports System.Globalization



'********************************************************************
'* �\�[�X�t�@�C���� : SystemUtil.vb
'* �N���X���@�@	    : SystemUtil
'* �N���X�����@	    : �V�X�e�����ʃ��[�e�B���e�B�[
'* ���l�@           :
'* �쐬  �@         : 2007/07/08 ���
'* �X�V����         :
' 20090201_1 Komagta OpenFileDialog�̉��P(�����l�̃t�@�C���p�X�\��)
'********************************************************************
''' <summary>
''' �V�X�e�����ʃ��[�e�B���e�B�[
''' </summary>
''' <remarks></remarks>
Public Class SystemUtil

#Region "�񋓑�"
#Region "�S�p�E���p"
    ''' <summary>
    ''' �S�p�E���p�E���݂̗񋓑�
    ''' </summary>
    ''' <remarks>�S�p�E���p�E���݂̗񋓑�</remarks>
    Public Enum CHAR_SIZE

        FULLHALF = 0 '����'
        FULL = 1 '�S�p
        HALF = 2 '���p
    End Enum
#End Region
#Region "�޲�۸�̨���"
    ''' <summary>
    ''' �t�@�C�A���O�{�b�N�X�̃t�@�C���t�B���^�["
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FILEKIND
        XLS = 0
        TXT = 1
        CSV = 2
        PDF = 3
        XLSX = 5
        ETC = 9
    End Enum
#End Region
#Region "DLL�EEXE���s�敪"
    Public Enum START_PG
        DLL     'DLL, �Q�ƋN��
        EXE     'EXE, �P�ƋN��
    End Enum
#End Region
#Region "�X�V�m�F"
    Public Enum ACTION
        INS = 0 '�V�K���[�h
        UPD = 1 '�C�����[�h
        RO = 2 '�\�����[�h
        DEL = 3 '�\�����[�h
        ERR = 9 '�G���[��
    End Enum
#End Region
#Region "�{�x�X"
    Public Enum KBHNS
        HNSYA = 0 '�{��
        SISYA = 1 '�V��
    End Enum
#End Region
#Region "�`�F�b�N"
    ''' <summary>
    ''' �`�F�b�N�̗L��
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum CHK
        [FALSE] = 0 '�Ȃ�
        [TRUE] = 1 '����
    End Enum
#End Region
#Region "�E������"
    Public Enum LorR
        [LEFT] = 0 '�Ȃ�
        [RIGHT] = 1 '����
    End Enum
#End Region

#End Region

#Region "���yCombo�֘A�z"
    '#Region "�R���{�f�[�^�̐���"
    '    ''' <summary>
    '    ''' �R���{�f�[�^�̐���
    '    ''' </summary>
    '    ''' <param name="kbn"></param>
    '    ''' <returns></returns>
    '    ''' <remarks>NFU�Ŏg�p�E���ݔ񋤒ʂ̂���Private�֕ύX</remarks>
    '    Private Shared Function CreateComboData(Optional ByVal kbn As Integer = Nothing) As DataTable
    '        Dim dtset As DataSet = New DataSet
    '        Dim dtt As DataTable
    '        Dim dtRow As DataRow
    '        Dim FCode As String = "CODE"
    '        Dim FName As String = "NAME"


    '        dtt = dtset.Tables.Add("TEMP")
    '        dtt.Columns.Add(FCode, Type.GetType("System.String"))
    '        dtt.Columns.Add(FName, Type.GetType("System.String"))

    '        '���R�[�h�ǉ�
    '        dtRow = dtt.NewRow
    '        dtRow(FCode) = "0" : dtRow(FName) = "�����s" : dtt.Rows.Add(dtRow)

    '        dtRow = dtt.NewRow
    '        dtRow(FCode) = "1" : dtRow(FName) = "���s" : dtt.Rows.Add(dtRow)

    '        Return dtt

    '    End Function
    '#End Region '20071019_1
    '#Region "�l����"
    '    ''' <summary>
    '    ''' �R���{�`�F�b�N�`�F�b�N����
    '    ''' </summary>
    '    ''' <param name="ctrlCombo">�R���{�R���g���[��</param>
    '    ''' <param name="blnPermitNULL">NULL���L��(Optional, Default�F�����Ȃ�) </param>
    '    ''' <returns>Boolean(True�F����@False�F�ُ�)</returns>
    '    ''' <remarks></remarks>
    '    Public Overloads Shared Function ChkCombo(ByVal ctrlCombo As IM.Combo, _
    '                                              Optional ByVal blnPermitNULL As Boolean = False) As Boolean
    '        'Dim blnResult As Boolean
    '        If Trim(ctrlCombo.Value) = "" Then
    '            'blnResult = blnPermitNULL
    '            Return blnPermitNULL
    '        Else
    '            If Not IsNothing(ctrlCombo.SelectedItem) Then
    '                'blnResult = True
    '                Return True
    '            Else
    '                If ctrlCombo.Value <> "" Then
    '                    Dim i As Integer
    '                    For i = 0 To ctrlCombo.Items.Count - 1
    '                        If GetCmbContent(ctrlCombo) = CStr(ctrlCombo.Items.Item(i).Value) Then
    '                            'blnResult = True
    '                            Return True
    '                        End If
    '                    Next
    '                End If
    '            End If
    '        End If
    '        'Return blnResult
    '        Return False
    '    End Function
    '    '-------------------------------------------------------------------------------------------------
    '    '�����FcmbCtrl           ( ComboControl�FIM.Combo )
    '    '      blFlg_PermitNull  ( NULL���L���FBoolean )
    '    '      strTitle          ( ���b�Z�[�WTitle�FString )
    '    '      inDegit           ( �����FInteger(Default�F1��) )
    '    '      ocon              ( SqlConnection�FMSG�\���p )
    '    '
    '    '���l�F               �y�V�K�z�FNULL����(True) �^�y�ύX�z�FNULL������(False)
    '    '�쐬��  �F2006.06.06  ��
    '    '--------------------------------------------------------------------------------------------------
    '    Public Overloads Shared Function ChkCombo(ByVal cmbCtrl As IM.Combo, _
    '                                    ByVal blFlg_PermitNull As Boolean, _
    '                                    ByVal strTitle As String) As Boolean

    '        If Not ChkCombo(cmbCtrl, blFlg_PermitNull) Then
    '            If Trim(cmbCtrl.Value) = "" And blFlg_PermitNull = False Then

    '                ''DEL 20061018_1
    '                ''20061113_1�@���A
    '                '< MSG(0,11)�F���̍��ڂ͕K�{���͂ł��B�������l����͂��Ă��������B>
    '                'PBS_ShowMsg(ocon, PBCSTR_MSGCTG_MUST_INPUT, PBCSTR_MSGID_MUST_INPUT, strTitle & vbCrLf)
    '                ShowErrorMsg(PBCSTR_MSG_ERROR_1, strTitle & vbCrLf)
    '                Return False
    '            Else
    '                '< MSG(0,2)�F���̍��ڂ̓��X�g���ɂ��鍀�ڂ���I�����Ă��������B>
    '                'PBS_ShowMsg(ocon, PBCSTR_MSGCTG_NO_LIST, PBCSTR_MSGID_NO_LIST, strTitle & vbCrLf)
    '                ShowErrorMsg(PBCSTR_MSG_ERROR_2, strTitle & vbCrLf)
    '                cmbCtrl.Text = ""
    '                Return False
    '            End If
    '        End If
    '        Return True
    '    End Function
    '    '-------------------------------------------------------------------------------------------------
    '    '�����FcmbCtrl           ( ComboControl�FIM.Combo )
    '    '      strContent        ( �w��Content )
    '    '�߂�l�FContent����=True
    '    '���l�F   
    '    '�쐬���F2006.07.14     ��
    '    '�C�����F2006.07.21     �� (�I�����ꂽIndex�ԍ��Ԃ�)
    '    '--------------------------------------------------------------------------------------------------
    '    Public Shared Function ChkComboContent(ByVal cmbCtrl As IM.Combo, _
    '                                           ByVal strContent As String, _
    '                                           Optional ByRef index As Integer = 0) As Boolean
    '        Dim blnRtn As Boolean
    '        If strContent <> "" Then
    '            For i As Integer = 0 To cmbCtrl.Items.Count - 1
    '                If strContent = CStr(cmbCtrl.Items.Item(i).Value) Then
    '                    index = i       '�� ADD 2006.07.21
    '                    blnRtn = True
    '                    Exit For
    '                End If
    '            Next
    '        End If
    '        Return blnRtn
    '    End Function
    '#End Region

    '#Region "�l�擾(Content�EDescription)"
    '    ' ------------------------------------------------------------------ 
    '    '�@�@�\�@�F�R���{�e�L�X�g��Content��Ԃ�
    '    '
    '    '�@�����@�F�R���{�R���g���[��(cmbCtrl)
    '    '�@�߂�l�F�R���{�e�L�X�g����'�b'�O�̕�����Ԃ�
    '    '
    '    '  �쐬���F2006.05.10�@��
    '    ' ------------------------------------------------------------------ 
    '    Public Overloads Shared Function GetCmbContent(ByVal cmbCtrl As IM.Combo) As String
    '        Dim intLength As Integer
    '        Dim strCmbVal, strContent As String
    '        Try
    '            If cmbCtrl.Value = "" Then Return ""
    '            strCmbVal = Trim(cmbCtrl.Value)

    '            intLength = InStr(strCmbVal, PBCSTR_VERTICAL)
    '            If intLength = 0 Then
    '                strContent = strCmbVal
    '            Else
    '                strContent = PBFSTR_MidB(strCmbVal, 1, intLength - 1)
    '            End If
    '            Return strContent
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    '    ' ------------------------------------------------------------------ 
    '    '�@�@�\�@�F�R���{�e�L�X�g��Content��Ԃ�
    '    '
    '    '�@�����@�F�R���{�R���g���[��(cmbCtrl)
    '    '          strValue(÷��)
    '    '�@�߂�l�F�R���{�e�L�X�g����'�b'�O�̕�����Ԃ�
    '    '
    '    '  �쐬���F2006.07.20�@��
    '    ' ------------------------------------------------------------------ 
    '    Public Overloads Shared Function GetCmbContent(ByVal cmbCtrl As IM.Combo, _
    '                                                   ByVal strValue As String) As String
    '        Dim intLength As Integer
    '        Dim strContent As String
    '        Try
    '            strContent = Trim(strValue)
    '            intLength = InStr(strContent, PBCSTR_VERTICAL)
    '            If intLength = 0 Then
    '                'Modify 2006.08.04
    '                'strContent = strContent
    '                strContent = Trim(strContent)
    '            Else
    '                'Modify 2006.08.04
    '                'strContent = PBFSTR_MidB(strContent, 1, intLength - 1)
    '                strContent = Trim(PBFSTR_MidB(strContent, 1, intLength - 1))
    '            End If
    '            If ChkComboContent(cmbCtrl, strContent) Then
    '                Return strContent
    '            Else
    '                Return ""
    '            End If
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    ''' <summary>
    ''' �R���{�̃e�L�X�g��"�b"�E��������擾
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function GetCmbText(ByVal strValue As String) As String
        Dim intLength As Integer
        Dim strCmbVal As String = ""
        Dim strContent As String = ""
        Try

            intLength = InStr(strValue, PBCSTR_VERTICAL)
            If intLength = 0 Then
                strContent = strValue
            Else
                strContent = strValue.Substring(intLength)
            End If
            Return strContent
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' �R���{�̃e�L�X�g��"�b"����������擾
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function GetCmbCode(ByVal strValue As String) As String
        Dim intLength As Integer
        Dim strCmbVal As String = ""
        Dim strContent As String = ""
        Try

            intLength = InStr(strValue, PBCSTR_VERTICAL)
            'intLength = strValue.Length - intLength

            If intLength = 0 Then
                strContent = strValue
            Else
                strContent = strValue.Substring(0, intLength - 1).Trim
            End If
            Return strContent
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    '    ' ------------------------------------------------------------------ 
    '    '�@�@�\�@�F�R���{�e�L�X�g��Description(����)��Ԃ�
    '    '
    '    '�@�����@�F�R���{�R���g���[��(cmbCtrl)
    '    '�@�߂�l�F�R���{�e�L�X�g����Description(����)
    '    '
    '    '  �쐬���F2006.06.08�@��
    '    '  �C�����F2006.08.28  ��
    '    ' ------------------------------------------------------------------ 
    '    Public Shared Function GetCmbDescription(ByVal cmbCtrl As IM.Combo, _
    '                                             Optional ByVal strContent As String = "") As String
    '        Dim strDescription As String = ""
    '        If cmbCtrl.Value <> "" OrElse strContent <> "" Then

    '            'ADD 2006.08.28
    '            If strContent = "" Then strContent = GetCmbContent(cmbCtrl)
    '            For i As Integer = 0 To cmbCtrl.Items.Count - 1
    '                'Modify 2006.08.28
    '                'If GetCmbContent(cmbCtrl) = CStr(cmbCtrl.Items.Item(i).Value) Then
    '                If strContent = CStr(cmbCtrl.Items.Item(i).Content) Then
    '                    strDescription = PBCStr(cmbCtrl.Items.Item(i).Description)
    '                End If
    '            Next
    '        End If
    '        Return strDescription
    '    End Function
    '#End Region
    '#Region "�����ޯ��(DropDownList)�F�l�����E�Ԃ�"
    '    ''' <summary>
    '    ''' �R���{�{�b�N�X�̃��X�g���̂�߂�
    '    ''' </summary>
    '    ''' <param name="ctrlCombo">�Ώۂ̃R���{�{�b�N�X�R���g���[��</param>
    '    ''' <param name="strValue">�^�[�Q�b�g�R�[�h(�l)</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    Public Shared Function GetCmbListDescription(ByVal ctrlCombo As IM.Combo, _
    '                                                           Optional ByVal strValue As String = "") As String
    '        With ctrlCombo
    '            Dim strcontents As String = ""
    '            If strValue = "" Then strValue = .Text
    '            If strValue = "" Then Return ""

    '            For i As Integer = 0 To .Items.Count - 1
    '                If strValue = CStr(.Items.Item(i).Value) Then
    '                    strcontents = CStr(.Items.Item(i).Content)
    '                    Exit For
    '                End If
    '            Next
    '            Return strcontents
    '        End With
    '    End Function
    '#End Region
#End Region

#Region "���yXML�֘A�z"
#Region "XML�t�@�C����������"
    Public Shared Sub PB_WriteXML(ByVal xmlPath As String, ByVal prmElement As String, ByVal prmValue As String)
        Try


            Dim domDoc As New XmlDocument
            Dim domNode As XmlNode

            'XML �`���̕�����f�[�^��ݒ肷�� 
            domDoc.Load(xmlPath)

            '����̗v�f�ɃA�N�Z�X���� 

            domNode = domDoc.SelectSingleNode(prmElement)
            domNode.InnerText = prmValue

            ''����̑����ɃA�N�Z�X���� 
            'domNode = domDoc.SelectSingleNode("//Item/@att")
            'Console.WriteLine("{0} => {1}", domNode.LocalName, domNode.Value)

            '�t�@�C���Ƃ��ĕۑ����� 

            domDoc.Save(xmlPath)



        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''--------------------------------------------------------
    ''  �@�\    �FXML��������������
    ''  ����    �F�P�D�p�X�w��
    ''  �@�@�@  �F�Q�D�Ȃ������ꍇ�̃f�t�H���g�l
    ''  �߂�l  �F��������
    ''  �쐬��  �F
    ''--------------------------------------------------------
    'Public Shared Sub PBS_AppendXML(ByVal xmlPath As String, ByVal prmElement As String, ByVal prmValue As String)

    '    Try
    '        Dim xmlDoc As New System.Xml.XmlDocument
    '        xmlDoc.Load(xmlPath)
    '        '�v�f��ǉ�����
    '        Dim xmlRoot As System.Xml.XmlElement = xmlDoc.DocumentElement
    '        Dim xmlEle As System.Xml.XmlElement = xmlRoot.Item(prmElement)
    '        Dim xmlValue As System.Xml.XmlText
    '        '���݃`�F�b�N
    '        Dim xmlList As System.Xml.XmlNodeList = xmlDoc.GetElementsByTagName(prmElement)
    '        If xmlList.Count > 0 Then
    '            '�폜
    '            xmlRoot.RemoveChild(xmlEle)
    '        End If
    '        '�ǉ�
    '        xmlEle = xmlDoc.CreateElement(prmElement)
    '        xmlValue = xmlDoc.CreateTextNode(prmValue)
    '        xmlRoot.AppendChild(xmlEle)
    '        xmlEle.AppendChild(xmlValue)


    '        xmlDoc.Save(xmlPath)

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
#End Region
#Region "XML�t�@�C���ǂݍ���"
    '--------------------------------------------------------
    '  �@�\    �FXML������ǂݍ���
    '  ����    �F�P�D�p�X�w��
    '  �@�@�@  �F�Q�D�Ȃ������ꍇ�̃f�t�H���g�l
    '  �߂�l  �F��������
    '  �쐬��  �F
    '--------------------------------------------------------
    Public Shared Function PB_ReadXML(ByVal Path As String, ByVal DefaultVal As String, ByVal xmlPath As String) As String

        Try
            Dim xmlDoc As XmlDocument = New XmlDocument
            xmlDoc.Load(xmlPath)
            Dim list As XmlNodeList = xmlDoc.SelectNodes(Path)
            Dim node As XmlNode
            If list.Count <= 0 Then
                Return DefaultVal
            End If

            node = list.Item(0)

            Return node.InnerText

        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "XML�t�@�C���ǂݍ���"
    '--------------------------------------------------------
    '  �@�\    �FXML������ǂݍ���
    '  ����    �F�P�D�p�X�w��
    '  �@�@�@  �F�Q�D�Ȃ������ꍇ�̃f�t�H���g�l
    '  �߂�l  �F��������
    '  �쐬��  �F
    '--------------------------------------------------------
    Public Shared Function PB_ReadXmlNodeList(ByVal Path As String, ByVal xmlPath As String) As XmlNodeList

        Dim ary As New ArrayList

        Try
            Dim xmlDoc As XmlDocument = New XmlDocument
            xmlDoc.Load(xmlPath)
            Dim list As XmlNodeList = xmlDoc.SelectNodes(Path)
            Dim node As XmlNode = Nothing
            If list.Count <= 0 Then
                Return Nothing
            End If

            node = list.Item(0)


            Return node.ChildNodes

        Catch ex As Exception
            Throw ex
        End Try
    End Function


#End Region
#Region "�ʌč���"

#End Region

#End Region

#Region "���y�f�[�^�^�ϊ����b�\�h�z"

    ''---------------------------------------------------------------------
    '' �@�\    �FINPUTMAN Date���t�Z�b�g(�����̎��F�N���A)
    '' ����    �F1. Value(���t�FYYYY/MM/DD(OracleDefaultValue), YY/MM/DD)
    ''           2. DateController(�Z�b�g����R���g���[���[)
    '' �߂�l  �F����
    '' �쐬��  �F2005.12.20 ��
    ''---------------------------------------------------------------------
    'Public Shared Sub PBS_SetIMDate(ByVal objVal As Object, ByVal editDate As IM.Date)

    '    If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
    '        editDate.Clear()
    '    Else
    '        editDate.Value = ToDateTimeEx(CStr(objVal), _
    '            System.Globalization.CultureInfo.CurrentCulture)
    '    End If
    'End Sub
    ''' <summary>
    ''' �l�ϊ�(String)
    ''' </summary>
    ''' <param name="objVal">�ϊ�����l</param>
    ''' <param name="rtnValue">��l�̏ꍇ�̖߂�l</param>
    ''' <returns>String�ߒl</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCStr(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As String = "") As String

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        Else
            Return CStr(objVal)
        End If

    End Function
    ''' <summary>
    ''' �l�ϊ�(Integer)
    ''' </summary>
    ''' <param name="objVal">�ϊ�����l</param>
    ''' <param name="rtnValue">��l�̏ꍇ�̖߂�l</param>
    ''' <returns>Integer�߂�l</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCint(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Integer = 0) As Integer
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CInt(objVal)
        End If
    End Function

    ''' <summary>
    ''' �l�ϊ�(Boolean)
    ''' </summary>
    ''' <param name="objVal">�ϊ�����l</param>
    ''' <param name="rtnValue">��l�̏ꍇ�̖߂�l</param>
    ''' <returns>Boolean�߂�l</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCBool(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Boolean = False) As Boolean
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        Else
            Return CBool(objVal)
        End If
    End Function
    ''' <summary>
    ''' �l�ϊ�(Byte)
    ''' </summary>
    ''' <param name="objVal">�ϊ�����l</param>
    ''' <param name="rtnValue">��l�̏ꍇ�̖߂�l</param>
    ''' <returns>Byte�߂�l</returns>
    ''' <remarks></remarks>
    Public Function PBCbyt(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Byte = 0) As Byte
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CByte(objVal)
        End If
    End Function
    ''' <summary>
    ''' �l�ϊ�(Long)
    ''' </summary>
    ''' <param name="objVal">�ϊ�����l</param>
    ''' <param name="rtnValue">��l�̏ꍇ�̖߂�l</param>
    ''' <returns>Long�߂�l</returns>
    ''' <remarks></remarks>
    Public Shared Function PBClng(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Long = 0) As Long
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CLng(objVal)
        End If
    End Function

    ''' <summary>
    ''' �l�ϊ�(Decimal)
    ''' </summary>
    ''' <param name="objVal">�ϊ�����l</param>
    ''' <param name="rtnValue">��l�̏ꍇ�̖߂�l</param>
    ''' <returns>Decimal�߂�l</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCdec(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Decimal = 0D) As Decimal
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CDec(objVal)
        End If
    End Function
    ''' <summary>
    ''' �l�ϊ�(Double)
    ''' </summary>
    ''' <param name="objVal">�ϊ�����l</param>
    ''' <param name="rtnValue">��l�̏ꍇ�̖߂�l</param>
    ''' <returns>Double�߂�l</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCdbl(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Double = 0) As Double
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CDbl(objVal)
        End If
    End Function
    ''' <summary>
    '''  NULL�`�F�b�N
    ''' </summary>
    ''' <param name="objVal">�`�F�b�N����l</param>
    ''' <param name="bln">��l�ȊO�̏ꍇ�̖߂�l</param>
    ''' <returns>True�FNull�@False:Nukk�ȊO</returns>
    ''' <remarks></remarks>
    Public Shared Function PB_ChkNUll(ByVal objVal As Object, _
                                   Optional ByVal bln As Boolean = False) As Boolean

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return True
        Else
            Return bln
        End If
    End Function

    '---------------------------------------------------------------------
    '  �@�\    �F�k��(Nothing,DBNull, "")�`�F�b�N
    '  ����    �F1�DObject, 2.Byte(�����`�E�����`�敪) �� 0:���� 9:���ݒ�(����)
    '  �߂�l  �FString(�k���̏ꍇ"NULL"�Ԃ��A�Ȃ��ꍇ��''������)
    '            Integer(���̂܂ܕԂ�)
    '  �쐬��  �F2006.01.17  ��
    '---------------------------------------------------------------------
    '�����`�ł�������������ꂽ���͐����`�ɔF�����Ă�̂ŁB�B�B�B
    '������\�������܂��Ă�̂ŁB
    Public Shared Function PBFSTR_SetQTT(ByVal objVal As Object, _
                                    Optional ByVal byKBN As Byte = 9) As String

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
            Return "NULL"
        Else
            Select Case byKBN
                Case 9    '�����`�̏ꍇ
                    If IsDate(objVal) Then '20070308_1 �ǂ�����x��
                        'If PB_IsDate(CStr(objVal)) Then
                        Return "'" & CStr(objVal) & "'"
                    ElseIf IsNumeric(objVal) Then
                        Return "'" & addQuot(CStr(objVal)) & "'"
                    Else
                        Return "'" & addQuot(CStr(objVal)) & "'"
                    End If

                Case 0      '�����`�̏ꍇ
                    If IsNumeric(objVal) Then
                        Return CStr(objVal)
                    Else
                        Return CStr(objVal)
                    End If
                Case Else
                    Return CStr(objVal)

            End Select
        End If
    End Function
    '---------------------------------------------------------------------
    '  �@�\    �F�k��(Nothing,DBNull, "")��Decimal�ɕϊ�
    '  ����    �F�P�DObject, (�Q�DDecimal )
    '  �߂�l  �FDecimal
    '  �쐬��  �F2006.07.25  F.Nishida
    '---------------------------------------------------------------------
    Public Shared Function PBCsng(ByVal objVal As Object, _
                                   Optional ByVal rtnValue As Single = 0) As Single
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CSng(objVal)
        End If
    End Function
    '---------------------------------------------------------------------
    '  �@�\    �F�k��(Nothing,DBNull, "")��Date�ɕϊ�
    '  ����    �F�P�DObject, (�Q�DData )
    '  �߂�l  �FDate
    '  �쐬��  �F2007.08.01 ����
    '---------------------------------------------------------------------
    Public Shared Function PBCDate(ByVal objVal As Object, _
                                   Optional ByVal rtnValue As Date = Nothing) As Date

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
            Return rtnValue
        ElseIf IsDate(objVal) Then
            Return CDate(objVal)
        Else
            Return Nothing
        End If
    End Function

    '---------------------------------------------------------------------
    '  �@�\    �F�k��(Nothing,DBNull, "")��Date�ɕϊ�
    '  ����    �F�P�DObject, (�Q�DData )
    '  �߂�l  �FDate
    '  �쐬��  �F2007.08.01 ����
    '---------------------------------------------------------------------
    Public Shared Function PBCDateTime(ByVal objVal As Object, _
                                   Optional ByVal rtnValue As Date = Nothing) As Date

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
            Return Nothing
        ElseIf IsDate(objVal) Then
            Return DateTime.Parse(PBCStr(objVal))
        End If
    End Function
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
    Public Shared Function addQuot(ByVal strVal As String) As String

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
        Return strOutputVal & strInputVal

    End Function
#End Region


#End Region

#Region "���y�f�[�^�擾�֘A�z"
#Region "�f�[�^�Z�b�g���Y���f�[�^���擾����"
    Public Shared Function GetOneItemData(ByVal dtTbl As DataTable, ByVal FldName As String, ByVal szWhere As String) As String
        Dim dtView As DataView

        Try
            dtView = New DataView(dtTbl, szWhere, "", DataViewRowState.CurrentRows)
            If dtView.Count <= 0 Then
                Return ""
            Else
                Return PBCStr(dtView.Item(0)(FldName))
            End If

        Catch ex As Exception
            'SkyLog.Error(ex.Message)
            Return ""
        End Try

    End Function
#End Region
#Region "�f�[�^�Z�b�g���₢���킹����(View)���擾����"
    Public Shared Function GetResultDataView(ByVal dtTbl As DataTable, ByVal szWhere As String) As DataView
        Dim dtView As DataView

        Try
            dtView = New DataView(dtTbl, szWhere, "", DataViewRowState.CurrentRows)
            If dtView.Count <= 0 Then
                Return Nothing
            Else
                Return dtView
            End If

        Catch ex As Exception
            'SkyLog.Error(ex.Message)
            Return Nothing
        End Try

    End Function
#End Region '20090729_1
#Region "�f�[�^�Z�b�g�����f�[�^�̑��݊m�F"
    Public Shared Function ExistValue(ByVal dtTbl As DataTable, ByVal szWhere As String) As Boolean
        Dim dtView As DataView

        Try
            dtView = New DataView(dtTbl, szWhere, "", DataViewRowState.CurrentRows)
            If dtView.Count <= 0 Then
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            'SkyLog.Error(ex.Message)
            Throw ex
        End Try

    End Function
#End Region
#Region "ArrayList���f�[�^�e�[�u���֕ϊ�����"
    Public Shared Function GetdtFromArrayList(ByVal prmAryList As ArrayList, ByVal boHeader As Boolean) As DataTable
        Dim dt As New DataTable
        Dim dtRow As DataRow
        Dim arydt As New ArrayList

        arydt = prmAryList

        For i As Integer = 0 To prmAryList.Count - 1

            If boHeader Then
                '*---------------------
                'CSV��1�s�ڂ��w�b�_�[��
                '*---------------------
                If i.Equals(0) Then
                    '
                    Dim fields As String()
                    fields = CType(arydt.Item(i), String())

                    For j As Integer = 0 To fields.Length - 1
                        Dim headName As String = fields(j).ToString
                        dt.Columns.Add(headName)
                    Next
                Else
                    ''dataTable �� Row ��1�s���ǉ�.
                    dtRow = dt.NewRow()
                    dtRow.ItemArray = CType(prmAryList.Item(i), [Object]())
                    dt.Rows.Add(dtRow)
                End If

            Else
                '*---------------------
                '���ׂĂ𖾍׈���
                '*---------------------
                If i.Equals(0) Then
                    '
                    Dim fields As String()
                    fields = CType(arydt.Item(i), String())

                    For j As Integer = 0 To fields.Length - 1
                        dt.Columns.Add((j + 1).ToString)
                    Next
                End If


                ''dataTable �� Row ��1�s���ǉ�.
                dtRow = dt.NewRow()
                dtRow.ItemArray = CType(prmAryList.Item(i), [Object]())
                dt.Rows.Add(dtRow)
            End If

        Next

        Return dt

    End Function
#End Region '20100326_1
#End Region

#Region "���y����ŁE�[�������z"
    '#Region "�ŗ��擾"
    '    ''' <summary>
    '    ''' �ŗ��擾
    '    ''' </summary>
    '    ''' <param name="con"></param>
    '    ''' <param name="prmZeitp"></param>
    '    ''' <param name="prmBaseDate">����iyyyy/MM/dd�`���j</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    Public Shared Function getZeiRt(con As uniConnection, prmZeitp As Integer, prmBaseDate As String,
    '                                                                        Optional ByVal tran As uniTransaction = Nothing) As Decimal
    '        Try

    '            Dim szSQL As String = ""

    '            szSQL = ""
    '            szSQL += " select ZEI_RT from M_ZEI  "
    '            szSQL += " where 1=1 "
    '            szSQL += " and  zei_tp = " & prmZeitp
    '            szSQL += " and aply_dt = (select max(aply_dt) from m_zei zei_max where zei_max.zei_tp=" & prmZeitp & "and TO_CHAR(zei_max.aply_dt,'YYYY/MM/DD') <='" & prmBaseDate & "') "

    '            Dim ary As New ArrayList
    '            ary = getAryDataDB(con, szSQL, tran)
    '            If ary.Count.Equals(0) Then
    '                Return 0
    '            Else
    '                If PBCint(ary.Item(0)).Equals(0) Then
    '                    Return 0
    '                Else
    '                    Return PBCdec(PBCint(ary.Item(0)) / 100)
    '                End If
    '            End If
    '        Catch ex As Exception
    '            'Throw ex
    '            Return 0
    '        End Try
    '    End Function
    '#End Region
    '#Region "�Ŋz�擾"
    '    ''' <summary>
    '    ''' ����ŋ��z�̎擾
    '    ''' </summary>
    '    ''' <param name="prmItemKn">���i���z</param>
    '    ''' <param name="prmZeiRt">�ŗ�</param>
    '    ''' <param name="prmZeikKbn">�ō��敪</param>
    '    ''' <param name="prmHasu">�[�������敪</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    Public Shared Function GetZeiKn(ByVal prmItemKn As Integer, prmZeiRt As Decimal, prmZeikKbn As SystemConst.ZEIK, _
    '                                                    Optional ByVal prmHasu As SystemConst.Round = Round.Down, Optional prmKazeiKbn As SystemConst.KAZEI = KAZEI.KAZEI
    '                                                    ) As Decimal

    '        Try

    '            Dim rtnZeiKn As Decimal = 0

    '            If prmKazeiKbn = KAZEI.HIKAZEI Then
    '                rtnZeiKn = 0
    '            Else
    '                Select Case prmZeikKbn

    '                    Case SystemConst.ZEIK.ZEIOUT
    '                        ''�Ŕ��� (���i���z���ŗ�)�[������
    '                        rtnZeiKn = SystemUtil.doCalHASU(PBCdec(prmItemKn * prmZeiRt), prmHasu)

    '                    Case SystemConst.ZEIK.ZEIIN
    '                        ''����� ���z*(5/105) '20090907_1
    '                        ''�ō��� :���i���z�[(���i���z/1+�ŗ�)�[������
    '                        ''rtnZeiKn = prmItemKn - SystemUtil.PBF_CalHASU(PBCdec(prmItemKn / (ZeiRt + 1)), prmHasu)
    '                        'rtnZeiKn = SystemUtil.PBF_CalHASU(prmItemKn - PBCdec(prmItemKn / (ZeiRt + 1)), prmHasu)
    '                        rtnZeiKn = SystemUtil.doCalHASU(PBCdec(prmItemKn * prmZeiRt / (prmZeiRt + 1)), prmHasu)

    '                    Case Else
    '                        rtnZeiKn = 0
    '                End Select
    '            End If


    '            Return rtnZeiKn

    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    '#End Region
#Region "�[�������敪"
    Public Shared Function doCalHASU(ByVal value As Decimal, _
                                Optional ByVal bytHASU As Round = Round.Half) As Decimal
        Dim tempValue As Decimal = value
        If value < 0 Then
            tempValue = 0 - value
        End If
        Dim result As Decimal
        Select Case bytHASU
            Case Round.UP       '�؏グ
                result = Decimal.Truncate(tempValue)
                If result <> tempValue Then
                    result += 1D
                End If
            Case Round.Down      '�؎̂�
                result = Decimal.Truncate(tempValue)

            Case Round.Half     '�l�̌ܓ�
                result = Decimal.Truncate(tempValue + 0.5D)
            Case Else
                Return value
        End Select

        If value < 0 Then
            result = 0 - result
        End If

        Return result
    End Function
#End Region
#Region "�[�������敪(Double)"
    ''' <summary>
    ''' �[������(Double)
    ''' </summary>
    ''' <param name="dblValue">�ۂߑΏےl</param>
    ''' <param name="intDigits">�߂�l�̗L�������̐��x</param>
    ''' <param name="bytHASU">�[�������敪</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function doCalHASU2(ByVal dblValue As Double, _
                                   Optional ByVal intDigits As Integer = 0, _
                                   Optional ByVal bytHASU As Round = Round.Half) As Double
        Dim dblSquare As Double = Math.Pow(10, intDigits)
        Select Case bytHASU
            Case Round.UP  '�؏グ
                If dblValue > 0 Then
                    Return Math.Ceiling(dblValue * dblSquare) / dblSquare
                Else
                    Return Math.Floor(dblValue * dblSquare) / dblSquare
                End If

            Case Round.Down  '�؎̂�
                If dblValue > 0 Then
                    Return Math.Floor(dblValue * dblSquare) / dblSquare
                Else
                    Return Math.Ceiling(dblValue * dblSquare) / dblSquare
                End If

            Case Round.Half  '�l�̌ܓ�
                If dblValue > 0 Then
                    Return Math.Floor((dblValue * dblSquare) + 0.5) / dblSquare
                Else
                    Return Math.Ceiling((dblValue * dblSquare) - 0.5) / dblSquare
                End If
            Case Else
                Return dblValue
        End Select
    End Function
#End Region

#End Region

#Region "���y���t�֘A�z"
#Region "���t���ǂ����𒲂ׂ�"
    ''' <summary>
    ''' "���t���ǂ����𒲂ׂ�
    ''' </summary>
    ''' <param name="szObj">���؂���l</param>
    ''' <returns>True�F���t�ł��@False�F���t�łȂ�</returns>
    ''' <remarks></remarks>
    Public Shared Function PB_IsDate(ByVal szObj As String) As Boolean


        'DateTime�ɕϊ��ł��邩�m���߂�
        Try
            DateTime.Parse(szObj)
            Return True
        Catch
            Return False
        End Try
    End Function
#End Region
#Region "���t�̐���(�N��)"
    ''' <summary>
    ''' ���t�̐���(�N��)
    ''' </summary>
    ''' <param name="inY">�N</param>
    ''' <param name="inM">��</param>
    ''' <param name="inD">��</param>
    ''' <param name="inAddDate">���Z�l</param>
    ''' <param name="interval">�C���^�[�o��</param>
    ''' <param name="szFormat">�߂�l�̃t�H�[�}�b�g</param>
    ''' <returns>���t(����)</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function PBGetDate(ByVal inY As Integer, ByVal inM As Integer, _
                            Optional ByVal inD As Integer = 1, Optional ByVal inAddDate As Integer = 0, _
                            Optional ByVal interval As DateInterval = DateInterval.Day, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date
        Dim szDate As String

        Try


            szDate = inY & "/" & inM & "/" & inD
            dtDate = CDate(DateValue(szDate).ToString("yyyy/MM/dd"))
            ''InterVal�Z�b�g
            dtDate = DateAdd(interval, inAddDate, dtDate)
            ''�t�H�[�}�b�g
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "���t�̐���(�N����)"
    ''' <summary>
    '''  ���t�̐���(�N����)
    ''' </summary>
    ''' <param name="szYMD">���t(yyyyMMdd�`��)</param>
    ''' <param name="inAddDate">���Z�l</param>
    ''' <param name="interval">�C���^�[�o��</param>
    ''' <param name="szFormat">�߂�l�̃t�H�[�}�b�g</param>
    ''' <returns>���t(����)</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function getDate(ByVal szYMD As String, Optional ByVal inAddDate As Integer = 0, _
                            Optional ByVal interval As DateInterval = DateInterval.Day, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date
        Dim szDate As String

        Try

            If szYMD.Equals("00010101") Then
                Return ""
            End If


            szDate = GetWantedByte(szYMD, 0, 4) & "/" & GetWantedByte(szYMD, 4, 2) & "/" & GetWantedByte(szYMD, 6, 2)
            dtDate = CDate(DateValue(szDate).ToString("yyyy/MM/dd"))
            ''InterVal�Z�b�g
            dtDate = DateAdd(interval, inAddDate, dtDate)
            ''�t�H�[�}�b�g
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    '''  ���t�̐���(�N����)
    ''' </summary>
    ''' <param name="prmDate">���t</param>
    ''' <param name="inAddDate">���Z�l</param>
    ''' <param name="interval">�C���^�[�o��</param>
    ''' <param name="szFormat">�߂�l�̃t�H�[�}�b�g</param>
    ''' <returns>���t(����)</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function getDate(ByVal prmDate As Date, Optional ByVal inAddDate As Integer = 0, _
                            Optional ByVal interval As DateInterval = DateInterval.Day, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date

        Try

            If prmDate = Nothing Then
                Return ""
            End If

            ''InterVal�Z�b�g
            dtDate = DateAdd(interval, inAddDate, prmDate)
            ''�t�H�[�}�b�g
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Return ""
        End Try
    End Function
#End Region
#Region "���t�̐���(yyyy/MM/dd��yyyyMMdd)"
    ''' <summary>
    ''' ���t�̏����ϊ�
    ''' </summary>
    ''' <param name="szYMD">�ϊ����镶����</param>
    ''' <returns>�ϊ���̓��t������</returns>
    ''' <remarks>yyyy/MM/dd��yyyyMMdd</remarks>
    Public Overloads Shared Function PBGetCngDate(ByVal szYMD As String) As String

        Dim szDate As String = ""
        Dim dtDate As Date
        Try

            '���t���ۂ�����
            dtDate = CDate(DateValue(szYMD))


            szDate += GetWantedByte(szYMD, 0, 4)
            szDate += GetWantedByte(szYMD, 5, 2)
            szDate += GetWantedByte(szYMD, 8, 2)

            Return szDate

        Catch ex As Exception
            Return ""
        End Try
    End Function
#End Region
#Region "���t�ϊ�(yyyyMMdd��yyyy/MM/dd)"
    ''' <summary>
    ''' ���t�ϊ�(yyyyMMdd��yyyy/MM/dd)"
    ''' </summary>
    ''' <param name="szYMD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function getCngDate(ByVal szYMD As String, Optional formatString As String = "yyyy/MM/dd") As String

        Try
            Dim result As DateTime

            If DateTime.TryParseExact(szYMD, "yyyyMMdd", Nothing, DateTimeStyles.None, result) Then
                Return result.ToString(formatString)
            Else
                Return ""
            End If

        Catch ex As Exception
            Return ""
        End Try

    End Function
#End Region
#Region "���������߂�B"
    ''' <summary>
    ''' ���������߂�
    ''' </summary>
    ''' <param name="prmDate">yyyy/MM/dd�`��</param>
    ''' <param name="szFormat">�߂�l�ƂȂ���t�̏���</param>
    ''' <returns>���t�̕�����</returns>
    ''' <remarks></remarks>
    Public Shared Function getLastDate(ByVal prmDate As String, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date

        Try



            dtDate = CDate(DateValue(prmDate).ToString("yyyy/MM/dd"))
            ''InterVal�Z�b�g'1������
            dtDate = DateAdd(DateInterval.Month, 1, dtDate)
            '-1���Ō������Z�b�g
            dtDate = DateAdd(DateInterval.Day, -1, dtDate)

            ''�t�H�[�}�b�g
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' ���������߂�
    ''' </summary>
    ''' <param name="inY">�N</param>
    ''' <param name="inM">��</param>
    ''' <param name="szFormat">�߂�l�ƂȂ���t�̏���</param>
    ''' <returns>���t�̕�����</returns>
    ''' <remarks></remarks>
    Public Shared Function PB_GetLastDate(ByVal inY As Integer, ByVal inM As Integer, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date
        Dim szDate As String

        Try


            szDate = inY & "/" & inM & "/" & "01"
            dtDate = CDate(DateValue(szDate).ToString("yyyy/MM/dd"))
            ''InterVal�Z�b�g'1������
            dtDate = DateAdd(DateInterval.Month, 1, dtDate)
            '-1���Ō������Z�b�g
            dtDate = DateAdd(DateInterval.Day, -1, dtDate)

            ''�t�H�[�}�b�g
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "�c�莞�Ԃ����߂�"
    ''' <summary>
    ''' �c�莞�Ԃ����߂�
    ''' </summary>
    ''' <param name="dtNow">���݂̎���</param>
    ''' <param name="dtTarget">�^�[�Q�b�g�ƂȂ鎞��</param>
    ''' <param name="boGuide">�߂�l�̐ݒ�(True:�K�C�h False:����)</param>
    ''' <returns>String(����or�K�C�hor���ݸ)</returns>
    ''' <remarks>
    ''' <para>boGuide = Ture �̏ꍇ(�K�C�h)</para>
    ''' <para>1���ȉ��̏ꍇXX�b�Ŗ߂��i0�`59�b�j</para>
    ''' <para>1���ȏ��1���Ԉȉ��̏ꍇ�AXX���Ŗ߂��i1�`59���j</para>
    ''' <para>1���Ԉȏ�ň���ȉ��̏ꍇ�AXX���ԂŖ߂��i1�`24���ԁj</para>
    ''' <para>����ȏ�̏ꍇ�AXX���Ŗ߂��i1���`�j</para>
    ''' </remarks>
    Public Shared Function PBGetTimeRemit(ByVal dtNow As Date, ByVal dtTarget As Date, Optional ByVal boGuide As Boolean = True) As String

        Dim intTime As Integer
        Dim intHour As Integer = 0
        Dim intMinute As Integer = 0
        Dim intSecond As Integer = 0


        '���݂̎�������^�[�Q�b�g�ƂȂ鎞��������
        intTime = PBCint(dtTarget.Subtract(dtNow).TotalSeconds)

        Select Case boGuide
            Case True
                '-----------------
                '�K�C�h�L��̏ꍇ
                '-----------------
                If 0 < intTime And intTime < 60 Then
                    '1���ȉ��̏ꍇ�AXX�b�Ŗ߂�
                    Return CStr(intTime) & "�b"

                ElseIf 60 <= intTime And intTime < 3600 Then
                    '1���ȏ��1���Ԉȉ��AXX���Ŗ߂�
                    Return CStr(CInt(intTime / 60)) & "��"

                ElseIf 3600 <= intTime And intTime < 86400 Then
                    '1���Ԉȏ�ň���ȉ��̏ꍇ�AXX���ԂŖ߂�
                    Return CStr(CInt(intTime / 3600)) & "����"

                ElseIf 86400 <= intTime Then
                    '����ȏ�̏ꍇ�AXX���Ŗ߂�
                    Return CStr(CInt(intTime / 86400)) & "��"

                Else

                    Return ""

                End If

            Case Else
                ''-----------------
                ''�K�C�h�����̏ꍇ
                ''-----------------
                If intTime > 0 Then

                    Return CStr(intTime)

                Else

                    Return ""

                End If

        End Select

    End Function

#End Region
#Region "�a������߂�"
    ''' <summary>
    ''' �a������߂�
    ''' </summary>
    ''' <param name="prmDate">yyyy/MM/dd�`��</param>
    ''' <param name="szFormat"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getWareki(ByVal prmDate As Date, Optional ByVal szFormat As String = "ggyy/MM/dd") As String

        Dim szDate As String

        Try

            Dim culture As Globalization.CultureInfo = New Globalization.CultureInfo("ja-JP")
            culture.DateTimeFormat.Calendar = New System.Globalization.JapaneseCalendar

            szDate = prmDate.ToString(szFormat, culture)

            Return szDate


        Catch ex As Exception
            Return ""
        End Try

    End Function
    ''' <summary>
    ''' �a������߂�
    ''' </summary>
    ''' <param name="prmDate">yyyy/MM/dd�`��</param>
    ''' <param name="szFormat"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getWareki(ByVal prmDate As String, Optional ByVal szFormat As String = "ggyy/MM/dd") As String

        Dim dtDate As Date
        Dim szDate As String

        dtDate = PBCDate(prmDate)

        Try

            Dim culture As Globalization.CultureInfo = New Globalization.CultureInfo("ja-JP")
            culture.DateTimeFormat.Calendar = New System.Globalization.JapaneseCalendar

            szDate = dtDate.ToString(szFormat, culture)

            Return szDate


        Catch ex As Exception
            Return ""
        End Try

    End Function
#End Region
#End Region

#Region "���t�@�C���֘A"
#Region "�t�@�C���֘A"

#Region "�ۑ��_�C�A���O"
    ''' <summary>
    ''' �t�@�C���ۑ��_�C�A���O�{�b�N�X�̕\��
    ''' </summary>
    ''' <param name="szDefaltFileName">�t�@�C����</param>
    ''' <param name="szFileDir">�t�@�C���f�B���N�g��</param>
    ''' <param name="szDefaultPath">�����f�B���N�g��</param>
    ''' <param name="Filter">�t�@�C�����(�����l�FXLS)</param>
    ''' <returns>Boolean(�����E���s)</returns>
    ''' <remarks></remarks>
    Public Shared Function SaveFileDialog(ByRef szDefaltFileName As String, ByRef szFileDir As String, _
                               Optional ByVal szDefaultPath As String = "", Optional ByVal Filter As FILEKIND = FILEKIND.XLS) As Boolean

        Dim sfd As New SaveFileDialog
        '�͂��߂̃t�@�C�������w�肷��
        sfd.FileName = szDefaltFileName
        '�͂��߂ɕ\�������t�H���_���w�肷��
        If Not szDefaultPath.Equals("") Then
            sfd.InitialDirectory = szDefaultPath
        End If

        '[�t�@�C���̎��]�ɕ\�������I�������w�肷��
        Select Case Filter
            Case FILEKIND.XLS
                sfd.Filter = "Excel�t�@�C�� (*.xls)|*.xls"
            Case FILEKIND.XLSX
                sfd.Filter = "Excel�t�@�C�� (*.xlsx)|*.xlsx"
            Case FILEKIND.TXT
                sfd.Filter = "�e�L�X�g�t�@�C�� (*.txt)|*.txt"
            Case FILEKIND.CSV
                sfd.Filter = "CSV�t�@�C�� (*.csv)|*.csv"
            Case FILEKIND.PDF
                sfd.Filter = "PDF�t�@�C�� (*.pdf)|*.pdf"
            Case FILEKIND.ETC
                sfd.Filter = "���ׂẴt�@�C�� (*.*)|*.*"
        End Select

        '[�t�@�C���̎��]�ł͂��߂�
        '�u���ׂẴt�@�C���v���I������Ă���悤�ɂ���
        sfd.FilterIndex = 2
        '�^�C�g����ݒ肷��
        sfd.Title = "�ۑ���̃t�@�C����I�����Ă�������"
        '�_�C�A���O�{�b�N�X�����O�Ɍ��݂̃f�B���N�g���𕜌�����悤�ɂ���
        sfd.RestoreDirectory = True
        '���ɑ��݂���t�@�C�������w�肵���Ƃ��x������
        '�f�t�H���g��True�Ȃ̂Ŏw�肷��K�v�͂Ȃ�
        sfd.OverwritePrompt = True
        '���݂��Ȃ��p�X���w�肳�ꂽ�Ƃ��x����\������
        '�f�t�H���g��True�Ȃ̂Ŏw�肷��K�v�͂Ȃ�
        sfd.CheckPathExists = True
        '�g���q���w�肳��Ȃ��ꍇ�Ɋg���q��ݒ肷��悤�ɂ���
        '�f�t�H���g��True�Ȃ̂Ŏw�肷��K�v�͂Ȃ�
        sfd.AddExtension = True

        '�_�C�A���O��\������
        If sfd.ShowDialog() = DialogResult.OK Then
            szFileDir = System.IO.Path.GetDirectoryName(sfd.FileName) & "\"
            szDefaltFileName = System.IO.Path.GetFileName(sfd.FileName)
            Return True
        Else
            Return False
        End If
    End Function
#End Region
#Region "�t�@�C��Open�_�C�A���O"
    ''' <summary>
    ''' �t�@�C��Open�_�C�A���O�{�b�N�X�̕\��
    ''' </summary>
    ''' <param name="szDefaultPath">�����f�B���N�g��(�ȗ�����C:\)</param>
    ''' <param name="Filter">�t�@�C�����(�����l�FXLS)</param>
    ''' <param name="title">�t�H�[���^�C�g��</param>
    ''' <returns>�t�@�C����</returns>
    ''' <remarks>�L�����Z�����͋�l��߂�</remarks>
    Public Shared Function OpenFileDialog(Optional ByVal szDefaultPath As String = "", _
                                         Optional ByVal Filter As FILEKIND = FILEKIND.XLS, _
                                          Optional ByVal FilterTxt As String = "", _
                                         Optional ByVal title As String = "") As String

        Using sfd As New OpenFileDialog
            '�͂��߂̃t�@�C�������w�肷��
            'sfd.FileName = szDefaltFileName
            '�͂��߂ɕ\�������t�H���_���w�肷��
            '20080201_1 Add
            If Not szDefaultPath.Equals("") Then
                sfd.InitialDirectory = szDefaultPath
            End If

            ''If szDefaultPath.Equals("") Then
            ''    sfd.InitialDirectory = "C:\"
            ''Else
            ''    sfd.InitialDirectory = szDefaultPath
            ''End If

            '[�t�@�C���̎��]�ɕ\�������I�������w�肷��
            Select Case Filter
                Case FILEKIND.XLS
                    sfd.Filter = "Excel�t�@�C�� (*.xls)|*.xls"
                Case FILEKIND.XLSX
                    sfd.Filter = "Excel�t�@�C�� (*.xlsx)|*.xlsx"
                Case FILEKIND.TXT
                    sfd.Filter = "�e�L�X�g�t�@�C�� (*.txt)|*.txt"
                Case FILEKIND.CSV
                    sfd.Filter = "CSV�t�@�C�� (*.csv)|*.csv"
                Case FILEKIND.ETC
                    sfd.Filter = FilterTxt
            End Select

            '[�t�@�C���̎��]�ł͂��߂�
            '�u���ׂẴt�@�C���v���I������Ă���悤�ɂ���
            sfd.FilterIndex = 2
            '�^�C�g����ݒ肷��
            If title.Equals("") Then
                sfd.Title = "�ۑ���̃t�@�C����I�����Ă�������"
            Else
                sfd.Title = title
            End If

            '�_�C�A���O�{�b�N�X�����O�Ɍ��݂̃f�B���N�g���𕜌�����悤�ɂ���
            sfd.RestoreDirectory = True
            '���݂��Ȃ��p�X���w�肳�ꂽ�Ƃ��x����\������
            '�f�t�H���g��True�Ȃ̂Ŏw�肷��K�v�͂Ȃ�
            sfd.CheckPathExists = True

            '�_�C�A���O��\������
            If sfd.ShowDialog() = DialogResult.OK Then
                Return sfd.FileName
            Else
                Return ""
            End If
        End Using

    End Function
#End Region
#Region "�t�H���_�_�C�A���O"
    ''' <summary>
    ''' �t�H���_�̎Q�ƃ_�C�A���O�\��
    ''' </summary>
    ''' <param name="szDefaultPath">�����ݒ�p�X</param>
    ''' <param name="szTitle">�t�@�C�A���O�^�C�g��</param>
    ''' <returns>�I�������t�H���_�Q�ƃp�X</returns>
    ''' <remarks>�L�����Z���̏ꍇ�͋�l</remarks>
    Public Shared Function FolderDialog(Optional ByVal szDefaultPath As String = "C:\", _
                                            Optional ByVal szTitle As String = "�t�H���_���Q�Ƃ��Ă�������") As String

        Dim szPath As String = ""
        Dim Dialog As New FolderBrowserDialog
        '�����Q��Path
        Dialog.SelectedPath = szDefaultPath

        '�_�C�A���O�{�b�N�X��[�V�����t�H���_�̍쐬]�{�^����\�����Ȃ��ꍇ�� False 
        Dialog.ShowNewFolderButton = False
        '�_�C�A���O�^�C�g��
        Dialog.Description = szTitle

        If Dialog.ShowDialog() = DialogResult.OK Then
            '�t�@�C���̎擾
            szPath = Dialog.SelectedPath
        End If

        Return szPath

    End Function
#End Region
#Region "�t�@�C���擾"
    ''' <summary>
    ''' �Ώۃf�B���N�g���̃t�@�C���ꗗ��߂�
    ''' </summary>
    ''' <param name="szPath">�Ώۃf�B���N�g���p�X</param>
    ''' <param name="arySerachPattarn">�����Ώۃt�@�C��</param>
    ''' <returns>�t�@�C���̈ꗗ</returns>
    ''' <remarks></remarks>
    Public Shared Function GetFileList(ByVal szPath As String, Optional ByVal arySerachPattarn As ArrayList = Nothing) As ArrayList
        Dim aryextension As New ArrayList
        Dim szFile As String
        Dim aryFiles As New ArrayList


        If arySerachPattarn Is Nothing Then
            aryextension.Add("*.*")
        Else
            aryextension = arySerachPattarn
        End If

        '�t�@�C�������݂���΂��̂܂ܕԂ�
        If IO.File.Exists(szPath) Then
            aryFiles.Add(Path.GetFileName(szPath))
            Return aryFiles
        End If


        For i As Integer = 0 To aryextension.Count - 1
            '�t�@�C���擾(�g�b�v�f�B���N�g���̂�)
            For Each szFile In Directory.GetFiles(szPath, PBCStr(aryextension.Item(i)), SearchOption.TopDirectoryOnly)
                aryFiles.Add(Path.GetFileName(szFile))
            Next
        Next


        Return aryFiles
    End Function
#End Region
#End Region
#End Region

#Region "���f�[�^�ҏW�֘A"
#Region "�w�蕶����̃o�C�g����߂�"
    Public Shared Function GetLengthASByte(ByVal prmVal As String) As Integer
        Return System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(prmVal)

    End Function
#End Region
#Region "Mid�֐��̃o�C�g��"
    '-----------------------------------------------------------------------------------
    '�@�@�\�@�F�������ƈʒu���o�C�g���Ŏw�肵�ĕ������؂蔲��
    '
    '�@�����@�FstrVal(�Ώۂ̕�����), 
    '          intStart(�؂蔲���J�n�ʒu�B
    '                   �S�p�����𕪊�����悤�ʒu���w�肳�ꂽ�ꍇ�A�߂�l�̕�����̐擪�͈Ӗ��s���̔��p�����ƂȂ�), 
    '          intLength(�؂蔲��������̃o�C�g��)
    '
    '�@�߂�l�FString(�؂蔲���ꂽ������)
    '
    '�@���l�@�F�Ō�̂P�o�C�g���S�p�����̔����ɂȂ�ꍇ�A���̂P�o�C�g�͖��������B
    '-----------------------------------------------------------------------------------
    Public Shared Function PBFSTR_MidB(ByVal strVal As String, _
                            ByVal intStart As Integer, _
                            ByVal intLength As Integer) As String

        Try
            '*** �󕶎��ɑ΂��Ă͏�ɋ󕶎���Ԃ�
            If strVal = "" Then Return ""

            '*** intLength�̃`�F�b�N
            'intLength��0���AintStart�ȍ~�̃o�C�g�����I�[�o�[����ꍇ��intStart�ȍ~�̑S�o�C�g���w�肳�ꂽ���̂Ƃ݂Ȃ��B
            Dim intResetLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(strVal) - intStart + 1
            If intLength = 0 OrElse intLength > intResetLength Then
                intLength = intResetLength
            End If

            '*** �؂蔲��
            Dim SJIS As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift-JIS")
            Dim bytBIG() As Byte = CType(Array.CreateInstance(GetType(Byte), intLength), Byte())
            Array.Copy(SJIS.GetBytes(strVal), intStart - 1, bytBIG, 0, intLength)


            Dim strNewVal As String = SJIS.GetString(bytBIG)

            '*** �؂蔲�������ʁA�Ō�̂P�o�C�g���S�p�����̔����������ꍇ�A���̔����͐؂�̂Ă�B
            Dim intResultLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(strNewVal) - intStart + 1

            If Asc(Strings.Right(strNewVal, 1)) = 0 Then
                'VB.NET2002,2003�̏ꍇ�A�Ō�̂P�o�C�g���S�p�̔����̎�
                Return strNewVal.Substring(0, strNewVal.Length - 1)

            ElseIf intLength = intResultLength - 1 Then
                'VB2005�̏ꍇ�ōŌ�̂P�o�C�g���S�p�̔����̎�
                Return strNewVal.Substring(0, strNewVal.Length - 1)

            Else
                '���̑��̏ꍇ
                Return strNewVal
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region
#Region "�w��o�C�g����������o��"
    ''' <summary>
    ''' �w��o�C�g������������o��
    ''' </summary>
    ''' <param name="strText">�^�[�Q�b�g�̕�����</param>
    ''' <param name="intStart">�J�n�ʒu</param>
    ''' <param name="intEnd">�I���ʒu</param>
    ''' <param name="intMultiLine">���s�r�� 0:���@1:�r��</param>
    ''' <returns>�ҏW��̕�����</returns>
    ''' <remarks>
    ''' �QByte�̕����̏ꍇ�A�w��Byte�����PByte�؂��ĕԂ�
    '''  Ex) test = "������"
    ''' "������" = PB_GetWantedString(test, 0, 5)
    ''' "����" = PB_GetWantedString(test, 0, 4)
    '''  "����" = PB_GetWantedString(test, 1, 4)
    ''' </remarks>  
    Public Shared Function GetWantedByte(ByVal strText As String, _
                                         ByVal intStart As Integer, _
                                         ByVal intEnd As Integer, _
                                        Optional ByVal intMultiLine As Integer = 0) As String

        '�w��o�C�g�ʒu����w��o�C�g�����̕���������o���֐�
        Dim strJIS As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        ''20061109_1 ���s�u��
        If intMultiLine <> 0 Then
            'strText = strText.Replace(ControlChars.CrLf, "")
            strText = strText.Replace(vbCrLf, "")
        End If

        If strText <> "" Then
            '�w�蕶������o�C�g�z��
            Dim bytText() As Byte = strJIS.GetBytes(strText)
            Dim intSumText As Integer = strJIS.GetByteCount(strText)

            '�����̃o�C�g������
            If intStart < 0 Or intEnd <= 0 Or intStart > intSumText Then Return ""

            '�X�^�[�g�������Q�b�g�A�QByte�̏ꍇ�̓X�^�[�g�o�C�g�� +1�ɂ���
            Dim strTemp As String = strJIS.GetString(bytText, 0, intStart)

            If intStart > 0 And strTemp.EndsWith(ControlChars.NullChar) Then
                intStart += 1   '�J�n�ʒu�������̒��Ȃ玟(�O)�̕�������J�n
            End If

            If intStart + intEnd > intSumText Then    '��������葽���擾���悤�Ƃ����ꍇ
                intEnd = intSumText - intStart        '������̍Ō�܂ł̕��Ƃ���
            End If

            '�w��o�C�g�̌��؂���薳���̏ꍇ�A���o����������Ԃ�
            '2005��2003�o���Ή� 20080806_1
            If strJIS.GetString(bytText, intStart, intEnd).EndsWith(ControlChars.NullChar) Or _
                                                strJIS.GetString(bytText, intStart, intEnd).EndsWith("�E") Then
                Return strJIS.GetString(bytText, intStart, intEnd - 1)
            End If

            Return strJIS.GetString(bytText, intStart, intEnd)

            ''Return strJIS.GetString(bytText, intStart, intEnd).TrimEnd(ControlChars.NullChar)
        Else
            Return strText
        End If
    End Function
#End Region
#Region "�w�蕶���O��̕�����擾"
    ' ------------------------------------------------------------------ 
    ' @(e) 
    ' �@�\        : PBFSTR_GetWantedText 
    ' �Ԃ�l      : String(���o����������)
    ' 
    ' ������      : strText:���o������������
    '               szLeft�F�ǂ��瑤�̕�������擾���邩�@True:�����@False:�E��
    '               szTarget:�w�蕶��
    '
    ' �@�\����    : �w�蕶���O��̕���������o��
    ' 
    ''------------------------------------------------------------------
    Public Shared Function PBFSTR_GetWantedText(ByVal szText As String, Optional ByVal LorR As LorR = LorR.LEFT, _
                                                  Optional ByVal szTarget As String = " ") As String
        Try

            Dim intBarCnt As Integer

            intBarCnt = InStr(szText, szTarget)

            If intBarCnt > 1 Then
                If LorR = LorR.LEFT Then
                    Return Left(szText, intBarCnt - 1)
                Else
                    Return Right(szText, Len(szText) - intBarCnt)
                End If
            Else
                If szTarget = " " Then
                    intBarCnt = InStr(szText, "�@")
                    If intBarCnt > 1 Then
                        If LorR = LorR.LEFT Then
                            Return Left(szText, intBarCnt - 1)
                        Else
                            Return Right(szText, Len(szText) - intBarCnt)
                        End If
                    End If
                    Return szText
                End If
                Return szText
            End If

        Catch ex As Exception
            'SkyLog.Error(ex.Message, ex)
            Return szText
        End Try
    End Function
#End Region
#Region "�S�p���p�̔��f"
    Public Shared Function ChkFullHalf(ByVal szText As String) As CHAR_SIZE

        Dim SJISEnc As Encoding = Encoding.GetEncoding("Shift_Jis")
        Dim inCnt As Integer = SJISEnc.GetByteCount(szText)

        '�uGetByteCount(str)�Ŏ擾�����o�C�g���v�Ɓustr.Length * 2�v����v����΁A������͂��ׂđS�p
        '�uGetByteCount(str)�Ŏ擾�����o�C�g���v�Ɓustr.Length�v����v����΁A������͂��ׂĔ��p
        '�uGetByteCount(str)�Ŏ擾�����o�C�g�� / 2�v�ŗ]�肪�o���ꍇ�́A�S�p�E���p���݂ł��B

        If szText.Length = inCnt Then
            Return CHAR_SIZE.HALF

        ElseIf szText.Length * 2 = inCnt Then
            Return CHAR_SIZE.FULL

        Else
            Return CHAR_SIZE.FULLHALF
        End If

    End Function
#End Region
#Region "�w�蕶���񂩂琔�������������݂̂��擾���A�A�����Ė߂�"
    Public Shared Function GetCharFromValue(ByVal prmValue As String) As String

        Dim CharVal As String = "" '�擾����������
        Dim MaxLength As Integer = prmValue.Length '�ő啶����

        For i As Integer = 0 To MaxLength - 1
            If Not IsNumeric(prmValue.Substring(i, 1)) Then
                CharVal += PBCStr(prmValue.Substring(i, 1))
            End If
        Next

        Return CharVal
    End Function
#End Region
#Region "�w�蕶����������Ēl��߂�"
    Public Shared Function RemoveFromValue(ByVal prmValue As String, ByVal prmtCar As String) As String

        Dim CharVal As String = "" '�擾����������
        Dim MaxLength As Integer = prmValue.Length '�ő啶����

        For i As Integer = 0 To MaxLength - 1
            If prmValue.Substring(i, 1) <> prmtCar Then
                CharVal += PBCStr(prmValue.Substring(i, 1))
            End If
        Next

        Return CharVal
    End Function
#End Region
#Region "���s�R�[�h��u������"
    '*************************************************************
    '* �@�\     : ���s�R�[�h��u������
    '* �Ԃ�l   : prmbaseText : ��{�ƂȂ镶����
    '* ������   : prmReplacement : �u����̕�����
    '* �@�\���� : �@�@�@�@�@�@
    '* �쐬     : 
    '* �X�V���� : 20090407_1 ����   Lf�̃P�[�X�������s�̕ϊ����ł��Ă��Ȃ��������ߏC��
    '*************************************************************
    Public Shared Function doReplaceLine(ByVal prmbaseText As String, Optional ByVal prmReplacement As String = " ") As String


        If IsNothing(prmbaseText) OrElse IsDBNull(prmbaseText) OrElse CStr(prmbaseText).Equals("") Then
            Return ""
        End If

        Dim rtn As String = ""

        rtn = Replace(prmbaseText, ControlChars.CrLf, prmReplacement) '�L�����b�W���^�[�������ƃ��C���t�B�[�h����
        rtn = Replace(rtn, ControlChars.Cr, prmReplacement) '�L�����b�W���^�[������
        rtn = Replace(rtn, ControlChars.Lf, prmReplacement) '���C���t�B�[�h����

        Return rtn

    End Function
#End Region '20091207_1
#Region "ArrayList����IN��𐶐�"
    '*************************************************************
    '* �@�\     : PBF_SQLIN
    '* �Ԃ�l   : CharVal : ������
    '* ������   : prmAry : ArrayList
    '*            prmText : XXXX.XXXXX
    '* �@�\���� : ��ʂ̃`�F�b�N
    '* ���l     : Ary����SQL��IN����쐬(ORACLE10g��IN���1000�܂łȂ̂�IN��𕪊�����)
    '*�@�@�@�@�@�@prmKey IN ('XXX','XXX','XXX', �c) or prmKey IN ('XXX','XXX','XXX', �c)
    '* �쐬     : 
    '* �X�V���� :
    '*************************************************************
    Public Shared Function PbfCreateSqlIN(ByVal prmAry As ArrayList, ByVal prmText As String) As String
        Try

            Dim intAry As Integer 'IN��̕K�v��
            Dim iCount As Integer '�s��(prmAry)
            Dim CharVal As String = "" '�擾����������
            Dim arytemp As New ArrayList '�ꎞAry
            Dim Max_Count As Integer = 999  ' AryTblMax�l

            '--------------
            'IN��̕K�v�������߂�
            '--------------
            If PBCint(prmAry.Count) > Max_Count Then
                '999���ȏ�̏ꍇ�A999�Ŋ��邱�Ƃ�IN�傪���K�v�Ȃ̂����v�Z
                intAry = PBCint(doCalHASU2(PBCdbl(prmAry.Count / Max_Count), 0, Round.Down))
                '��]��0�ȊO�̏ꍇ�͏�]�̕���IN�傪�K�v�Ȃ̂�+1������
                If PBCint(prmAry.Count) Mod Max_Count <> 0 Then intAry = intAry + 1
            Else
                '999�������̏ꍇ��1
                intAry = 1
            End If

            For j As Integer = 1 To intAry
                '--------------
                'IN��̐���
                '--------------
                '999�����ɂȂ�悤prmAry�𕪊�
                For i As Integer = 1 To Max_Count
                    If iCount < PBCint(prmAry.Count) Then
                        arytemp.Add(prmAry(iCount)) '�ꎞAry�ɃZ�b�g
                        iCount += 1
                    Else
                        i = Max_Count
                    End If
                Next

                CharVal += prmText & " IN ("

                '��ς�����  'XXX','XXX','XXX', �c
                For iCnt As Integer = 0 To arytemp.Count - 1
                    If iCnt = 0 Then
                        CharVal += "'" & PBCStr(arytemp.Item(iCnt)) & "'"
                    Else
                        CharVal += ",'" & PBCStr(arytemp.Item(iCnt)) & "'"
                    End If
                Next

                If j <> intAry Then
                    CharVal += ") OR " '�ŏI�łȂ��ꍇOR��t����
                Else
                    CharVal += ")"
                End If

                arytemp.Clear() '�ꎞAry�N���A

            Next

            Return CharVal

        Catch ex As Exception
            Throw ex
            Return ""
        End Try
    End Function
#End Region
#Region "�_�u���N�H�[�e�[�V�����Ŋ���"
    Public Shared Function setDoubleQuotes(field As String) As String
        If field.IndexOf(""""c) > -1 Then
            '"��""�Ƃ���
            field = field.Replace("""", """""")
        End If
        Return """" & field & """"
    End Function
#End Region
#End Region

#Region "�Í���"
    ''' <summary>
    ''' ��������Í�������
    ''' </summary>
    ''' <param name="sourceString">�Í������镶����</param>
    ''' <param name="password">�Í����Ɏg�p����p�X���[�h</param>
    ''' <returns>�Í������ꂽ������</returns>
    Public Shared Function doEncrypt(ByVal sourceString As String, _
                                         ByVal password As String) As String

        Try

            'RijndaelManaged�I�u�W�F�N�g���쐬
            Dim rijndael As New System.Security.Cryptography.RijndaelManaged()

            '�p�X���[�h���狤�L�L�[�Ə������x�N�^���쐬
            Dim key As Byte() = Nothing
            Dim iv As Byte() = Nothing
            GenerateKeyFromPassword(password, rijndael.KeySize, key, rijndael.BlockSize, iv)
            rijndael.Key = key
            rijndael.IV = iv

            '��������o�C�g�^�z��ɕϊ�����
            Dim strBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(sourceString)

            '�Ώ̈Í����I�u�W�F�N�g�̍쐬
            Dim encryptor As System.Security.Cryptography.ICryptoTransform = _
                rijndael.CreateEncryptor()
            '�o�C�g�^�z����Í�������
            Dim encBytes As Byte() = encryptor.TransformFinalBlock(strBytes, 0, strBytes.Length)
            '����
            encryptor.Dispose()

            '�o�C�g�^�z��𕶎���ɕϊ����ĕԂ�
            Return System.Convert.ToBase64String(encBytes)


        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' �Í������ꂽ������𕜍�������
    ''' </summary>
    ''' <param name="sourceString">�Í������ꂽ������</param>
    ''' <param name="password">�Í����Ɏg�p�����p�X���[�h</param>
    ''' <returns>���������ꂽ������</returns>
    Public Shared Function doDecrypt(ByVal sourceString As String, _
                                         ByVal password As String) As String

        Try

            'RijndaelManaged�I�u�W�F�N�g���쐬
            Dim rijndael As New System.Security.Cryptography.RijndaelManaged()

            '�p�X���[�h���狤�L�L�[�Ə������x�N�^���쐬
            Dim key As Byte() = Nothing
            Dim iv As Byte() = Nothing
            GenerateKeyFromPassword(password, rijndael.KeySize, key, rijndael.BlockSize, iv)
            rijndael.Key = key
            rijndael.IV = iv

            '��������o�C�g�^�z��ɖ߂�
            Dim strBytes As Byte() = System.Convert.FromBase64String(sourceString)

            '�Ώ̈Í����I�u�W�F�N�g�̍쐬
            Dim decryptor As System.Security.Cryptography.ICryptoTransform = _
                rijndael.CreateDecryptor()
            '�o�C�g�^�z��𕜍�������
            '�������Ɏ��s����Ɨ�OCryptographicException������
            Dim decBytes As Byte() = decryptor.TransformFinalBlock(strBytes, 0, strBytes.Length)
            '����
            decryptor.Dispose()

            '�o�C�g�^�z��𕶎���ɖ߂��ĕԂ�
            Return System.Text.Encoding.UTF8.GetString(decBytes)

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' �p�X���[�h���狤�L�L�[�Ə������x�N�^�𐶐�����
    ''' </summary>
    ''' <param name="password">��ɂȂ�p�X���[�h</param>
    ''' <param name="keySize">���L�L�[�̃T�C�Y�i�r�b�g�j</param>
    ''' <param name="key">�쐬���ꂽ���L�L�[</param>
    ''' <param name="blockSize">�������x�N�^�̃T�C�Y�i�r�b�g�j</param>
    ''' <param name="iv">�쐬���ꂽ�������x�N�^</param>
    Private Shared Sub GenerateKeyFromPassword(ByVal password As String, _
                                               ByVal keySize As Integer, _
                                               ByRef key As Byte(), _
                                               ByVal blockSize As Integer, _
                                               ByRef iv As Byte())
        '�p�X���[�h���狤�L�L�[�Ə������x�N�^���쐬����
        'salt�����߂�
        Dim salt As Byte() = System.Text.Encoding.UTF8.GetBytes("salt�͕K��8�o�C�g�ȏ�")
        'Rfc2898DeriveBytes�I�u�W�F�N�g���쐬����
        Dim deriveBytes As New System.Security.Cryptography.Rfc2898DeriveBytes( _
            password, salt)
        '.NET Framework 1.1�ȉ��̎��́APasswordDeriveBytes���g�p����
        'Dim deriveBytes As New System.Security.Cryptography.PasswordDeriveBytes( _
        '    password, salt)

        '���������񐔂��w�肷�� �f�t�H���g��1000��
        deriveBytes.IterationCount = 1000

        '���L�L�[�Ə������x�N�^�𐶐�����
        key = deriveBytes.GetBytes(keySize \ 8)
        iv = deriveBytes.GetBytes(blockSize \ 8)
    End Sub

#End Region

#Region "���K�C�h�p "
    Public Shared Function GUID_RowsCount(ByVal iCnt As Integer) As String
        Return "�������ʁF" & iCnt & "������܂����B"
    End Function
    Public Shared Function GUID_RegUserInfo(ByVal InUser As String, ByVal InDate As String, ByVal UpUser As String, ByVal Update As String) As String
        Return "[�쐬��]" & InUser & "(" & InDate & ")   [�X�V��]" & UpUser & "(" & Update & ") "
    End Function
#End Region

    '#Region "�f�[�^�擾(arrlyList)"
    '    '---------------------------------------------------------
    '    '�@�@�\�F�f�[�^�Q�b�g(Return ArrayList)
    '    '
    '    '�@�����@�FConnection, SQL��, Optional(Transaction)
    '    '�@�߂�l�FArrayList(�Q�b�g��������)
    '    '---------------------------------------------------------
    '    Private Shared Function getAryDataDB(ByVal ocon As uniConnection, ByVal SQL As String, _
    '                                        Optional ByVal tran As uniTransaction = Nothing) As ArrayList
    '        Dim ocd As New uniCommand
    '        Dim odr As UniDataReader
    '        Dim arlData As New ArrayList

    '        Try

    '            ocd.Connection = ocon
    '            ocd.CommandText = SQL

    '            If Not tran Is Nothing Then
    '                ocd.Transaction = tran
    '            End If

    '            odr = ocd.ExecuteReader

    '            While (odr.Read)
    '                For i As Integer = 0 To odr.FieldCount - 1
    '                    With arlData
    '                        .Add(PBCStr(odr.Item(i)))
    '                    End With
    '                Next
    '            End While

    '            odr.Close()
    '            Return arlData

    '        Catch ex As Exception
    '            Throw ex
    '        Finally

    '        End Try
    '    End Function
    '#End Region

#Region "�ۊ�"
#Region "���t�̐���(�N����)"
    ''Public Overloads Shared Function PBDataAdd(ByVal szYMD As String, Optional ByVal inAddDate As Integer = 0, _
    ''                        Optional ByVal interval As DateInterval = DateInterval.Day, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

    ''    Dim dtDate As Date

    ''    Try

    ''        dtDate = CDate(DateValue(szYMD).ToString("yyyy/MM/dd"))
    ''        ''InterVal�Z�b�g
    ''        dtDate = DateAdd(interval, inAddDate, dtDate)
    ''        ''�t�H�[�}�b�g
    ''        Return dtDate.ToString(szFormat)

    ''    Catch ex As Exception
    ''        Return ""
    ''    End Try
    ''End Function
#End Region
#Region "���t�̑Ó����`�F�b�N"
    'Public Function PBFBL_CheckDay(ByVal intY As Integer, ByVal intM As Integer, ByVal intD As Integer) As Boolean
    '    If (DateTime.MinValue.Year > intY) OrElse (intY > DateTime.MaxValue.Year) Then
    '        Return False
    '    End If

    '    If (DateTime.MinValue.Month > intM) OrElse (intM > DateTime.MaxValue.Month) Then
    '        Return False
    '    End If

    '    Dim iLastDay As Integer = DateTime.DaysInMonth(intY, intM)
    '    If (DateTime.MinValue.Day > intD) OrElse (intD > iLastDay) Then
    '        Return False
    '    End If

    '    Return True
    'End Function
#End Region
#Region "����ŗ��擾"
    ''�ŗ��F�œK�p�J�n�� �` �œK�p�I���� �Ԃ̐ŗ��擾
    'Public Function PBF_GetRTZEI(ByVal strDTZEI As String, ByVal con As SqlConnection) As Decimal
    '    Dim strSQL As String
    '    strSQL = ""
    '    strSQL = strSQL & " SELECT TO_CHAR(ZEI_RTZEI, '0.00') "
    '    strSQL = strSQL & " FROM M_ZEI "
    '    'strSQL = strSQL & " WHERE TO_CHAR(SYSDATE, 'YYYY/MM/DD') "
    '    strSQL = strSQL & " WHERE '" & strDTZEI & "'"
    '    strSQL = strSQL & "       BETWEEN TO_CHAR(ZEI_DTST, 'YYYY/MM/DD') "
    '    strSQL = strSQL & "       AND TO_CHAR(ZEI_DTED, 'YYYY/MM/DD') "

    '    Return PBFDEC_RtnDec(PBFSTR_GetOneDataDB(con, strSQL))
    'End Function
#End Region
#Region "����ŋ��z�v�Z (�d����)"
    ' -------------------------------------------------------------------------------------------
    ' �@�\        : PBF_CalSIRZEI 
    ' 
    ' �Ԃ�l      : Decimal(�Z�o���z)
    ' 
    '����         �FstrCDSIR�F�d����M.�d����CD
    '               decKNSIR�F�d�����z(�����z)
    '               intSURYO�F����
    '               decRTZEI�F����ŗ� 
    '               con : SqlConnection
    '               tran : SqlTransaction
    ' �@�\����    : 
    ' ���l        : 
    ''-------------------------------------------------------------------------------------------
    'Public Function PBF_CalSIRZEI(ByVal strCDSIR As String, ByVal lngKNSIR As Long, _
    '                              ByVal lngSURYO As Long, ByVal decRTZEI As Decimal, _
    '                              ByVal con As SqlConnection, _
    '                              Optional ByVal tran As SqlTransaction = Nothing) As Decimal


    '    Dim bytKBHASU As Byte   '�d����M.�[�������敪
    '    Dim strSQL As String

    '    Try
    '        strSQL = ""
    '        strSQL = strSQL & " SELECT SIR_KBHASU "
    '        strSQL = strSQL & " FROM M_SIR "
    '        strSQL = strSQL & " WHERE SIR_CDSIR =  " & strCDSIR

    '        bytKBHASU = PBCbyt(PBFSTR_GetOneDataDB(con, strSQL, tran))

    '        '�d�����z * ����ŗ� 
    '        Return PBF_CalHASU(lngKNSIR * decRTZEI, bytKBHASU)

    '    Catch ex As Exception
    '        SkyLog.Error(ex.Message, ex)
    '    End Try
    'End Function
#End Region
#Region "����ŋ��z�v�Z (���Ӑ�)"
    '' -------------------------------------------------------------------------------------------
    '' �@�\        : PBF_CalTOKZEI 
    '' 
    '' �Ԃ�l      : Decimal(�Z�o���z)
    '' 
    ''����         �FstrCDTOK�F���Ӑ�M.���Ӑ�CD
    ''               decKNTOK�F������z
    ''               intSURYO�F����
    ''               decRTZEI�F����ŗ� 
    ''               con : SqlConnection
    ''               tran : SqlTransaction
    '' �@�\����    : 
    '' ���l        : 
    ' ''-------------------------------------------------------------------------------------------
    'Public Function PBF_CalTOKZEI(ByVal strCDTOK As String, ByVal lngKNTOK As Long, _
    '                              ByVal lngSURYO As Long, ByVal decRTZEI As Decimal, _
    '                              ByVal con As SqlConnection, _
    '                              Optional ByVal tran As SqlTransaction = Nothing) As Decimal


    '    Dim bytKBHASU As Byte   '���Ӑ�M.�[�������敪
    '    Dim strSQL As String

    '    Try
    '        strSQL = ""
    '        strSQL = strSQL & " SELECT TOK_KBHASU "
    '        strSQL = strSQL & " FROM M_TOK "
    '        strSQL = strSQL & " WHERE TOK_CDTOK =  " & strCDTOK

    '        '20060818_1 �[�������͎l�̌ܓ�
    '        '' bytKBHASU = PBCbyt(PBFSTR_GetOneDataDB(con, strSQL, tran))
    '        bytKBHASU = 2

    '        '������z * ����ŗ� 
    '        Return PBF_CalHASU(lngKNTOK * decRTZEI, bytKBHASU)

    '    Catch ex As Exception
    '        SkyLog.Error(ex.Message, ex)
    '    End Try
    'End Function
#End Region
    '---------------------------------------------------------------------
    '  �@�\    �FDate��DateFormat�ɕϊ�
    '  ����    �F�P�DObject, (�Q�DString )
    '  �߂�l  �FString
    '  �쐬��  �F2006.07.25  F.Nishida
    '---------------------------------------------------------------------
    'Public Function PBFSTR_RtnDTE(ByVal objVal As Object, _
    ''                               Optional ByVal strFormat As String = "") As String
    '    '    Dim arrTemp() As String
    '    '    If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
    '    '        Return ""
    '    '    ElseIf IsDate(objVal) Then
    '    '        If strFormat = "" Or strFormat = "1" Then
    '    '            arrTemp = CStr(objVal).Split(CChar("/"))
    '    '            Return arrTemp(0) & "�N" & arrTemp(1) & "��" & arrTemp(2) & "��"
    '    '        End If
    '    '    Else
    '    '        Return ""
    '    '    End If
    'End Function
    '---------------------------------------------------------------------
    '  �@�\    �FDate��DateFormat�ɕϊ�
    '  ����    �F�P�DObject, (�Q�DString )
    '  �߂�l  �FString
    '  �쐬��  �F2006.07.25  F.Nishida
    '---------------------------------------------------------------------
    ''Public Function PBFSTR_IsDATE(ByVal objVal As Object, _
    ''                               Optional ByVal strFormat As String = "yyyyMMdd") As String
    ''    Dim arrTemp() As String
    ''    If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
    ''        Return ""
    ''    ElseIf IsDate(objVal) Then
    ''        Return CDate(objVal).ToString(strFormat)
    ''    Else
    ''        Return ""
    ''    End If
    ''End Function
#End Region
End Class
