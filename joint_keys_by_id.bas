Attribute VB_Name = "joint_keys_by_id"
    
Sub ConsolidateRowsUniqueValuesAndSaveAsCSV()
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim id_num As Variant
    Dim dict As Object, info As Object
    Dim outputPath As String
    Dim baseFileName As String
    Dim csvFileName As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set srcSheet = ThisWorkbook.Sheets("original")
    Set destSheet = ThisWorkbook.Sheets.Add
    destSheet.Name = "converted"
    
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row
    
    
    
    ' �f�[�^���H���W�b�N...
    For i = 2 To lastRow
        ' A��̒l���疖����5�������擾
        id_num = Right(srcSheet.Cells(i, "A").Value, 5)
        
        If Not dict.Exists(id_num) Then
            Set info = CreateObject("Scripting.Dictionary")
            ' �����l�Ƃ��Ċe�񂩂�̃f�[�^��ݒ�
            info("����/�o�T") = srcSheet.Cells(i, "AA").Value
            info("����/����") = srcSheet.Cells(i, "AB").Value
            info("�W������") = srcSheet.Cells(i, "Z").Value
            info("���M��") = srcSheet.Cells(i, "W").Value
            info("�o�Ŏ�") = srcSheet.Cells(i, "AE").Value
            info("�o�ŔN") = srcSheet.Cells(i, "AF").Value
            info("unidic") = "" ' ��̒l
            info("����") = srcSheet.Cells(i, "E").Value
            dict.Add id_num, info
        Else
            ' "����"�̗�̂ݒl������
            dict(id_num)("����") = dict(id_num)("����") & srcSheet.Cells(i, "E").Value
        End If
    Next i
    
    ' �w�b�_�[�o��
    With destSheet
        .Cells(1, 1).Value = "id_num"
        .Cells(1, 2).Value = "����/�o�T"
        .Cells(1, 3).Value = "����/����"
        .Cells(1, 4).Value = "�W������"
        .Cells(1, 5).Value = "���M��"
        .Cells(1, 6).Value = "�o�Ŏ�"
        .Cells(1, 7).Value = "�o�ŔN"
        .Cells(1, 8).Value = "unidic"
        .Cells(1, 9).Value = "����"
    End With
    
    ' �f�[�^�o��
    i = 2
    For Each id_num In dict.Keys
        With destSheet
            .Cells(i, 1).Value = id_num
            .Cells(i, 2).Value = dict(id_num)("����/�o�T") 
            .Cells(i, 3).Value = dict(id_num)("����/����") 
            .Cells(i, 4).Value = dict(id_num)("�W������") 
            .Cells(i, 5).Value = dict(id_num)("���M��") 
            .Cells(i, 6).Value = dict(id_num)("�o�Ŏ�") 
            .Cells(i, 7).Value = dict(id_num)("�o�ŔN") 
            .Cells(i, 8).Value = dict(id_num)("unidic") 
            .Cells(i, 9).Value = dict(id_num)("����") 
        End With
        i = i + 1
    Next id_num
    
    ' ����Excel�t�@�C�����i�g���q�Ȃ��j���擾
    baseFileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' �o�̓p�X��ݒ�iExcel�t�@�C���Ɠ����f�B���N�g���j
    outputPath = ThisWorkbook.Path & "\outputs"
    
    ' "outputs" �t�H���_�����݂��Ȃ��ꍇ�͍쐬
    If Dir(outputPath, vbDirectory) = "" Then
        MkDir outputPath
    End If
    
    csvFileName = baseFileName & "_j.csv"
    
    ' ���S�ȏo�̓t�@�C���p�X�̐���
    outputPath = outputPath & "\" & csvFileName

    
    ' �ꎞ�I�ɍ쐬�����V�[�g��CSV�t�@�C���Ƃ��ĕۑ�
    destSheet.SaveAs Filename:=outputPath, FileFormat:=xlCSV, Local:=True
    
    ' �ꎞ�V�[�g���폜�i���[�U�[�Ɋm�F�Ȃ��Łj
    Application.DisplayAlerts = False
    destSheet.Delete
    Application.DisplayAlerts = True
    
    MsgBox "CSV�t�@�C�����ۑ�����܂���: " & outputPath
End Sub
