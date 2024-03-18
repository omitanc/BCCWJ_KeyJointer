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
    
    
    
    ' データ加工ロジック...
    For i = 2 To lastRow
        ' A列の値から末尾の5文字を取得
        id_num = Right(srcSheet.Cells(i, "A").Value, 5)
        
        If Not dict.Exists(id_num) Then
            Set info = CreateObject("Scripting.Dictionary")
            ' 初期値として各列からのデータを設定
            info("書名/出典") = srcSheet.Cells(i, "AA").Value
            info("副題/分類") = srcSheet.Cells(i, "AB").Value
            info("ジャンル") = srcSheet.Cells(i, "Z").Value
            info("執筆者") = srcSheet.Cells(i, "W").Value
            info("出版者") = srcSheet.Cells(i, "AE").Value
            info("出版年") = srcSheet.Cells(i, "AF").Value
            info("unidic") = "" ' 空の値
            info("原文") = srcSheet.Cells(i, "E").Value
            dict.Add id_num, info
        Else
            ' "原文"の列のみ値を結合
            dict(id_num)("原文") = dict(id_num)("原文") & srcSheet.Cells(i, "E").Value
        End If
    Next i
    
    ' ヘッダー出力
    With destSheet
        .Cells(1, 1).Value = "id_num"
        .Cells(1, 2).Value = "書名/出典"
        .Cells(1, 3).Value = "副題/分類"
        .Cells(1, 4).Value = "ジャンル"
        .Cells(1, 5).Value = "執筆者"
        .Cells(1, 6).Value = "出版者"
        .Cells(1, 7).Value = "出版年"
        .Cells(1, 8).Value = "unidic"
        .Cells(1, 9).Value = "原文"
    End With
    
    ' データ出力
    i = 2
    For Each id_num In dict.Keys
        With destSheet
            .Cells(i, 1).Value = id_num
            .Cells(i, 2).Value = dict(id_num)("書名/出典") 
            .Cells(i, 3).Value = dict(id_num)("副題/分類") 
            .Cells(i, 4).Value = dict(id_num)("ジャンル") 
            .Cells(i, 5).Value = dict(id_num)("執筆者") 
            .Cells(i, 6).Value = dict(id_num)("出版者") 
            .Cells(i, 7).Value = dict(id_num)("出版年") 
            .Cells(i, 8).Value = dict(id_num)("unidic") 
            .Cells(i, 9).Value = dict(id_num)("原文") 
        End With
        i = i + 1
    Next id_num
    
    ' 元のExcelファイル名（拡張子なし）を取得
    baseFileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' 出力パスを設定（Excelファイルと同じディレクトリ）
    outputPath = ThisWorkbook.Path & "\outputs"
    
    ' "outputs" フォルダが存在しない場合は作成
    If Dir(outputPath, vbDirectory) = "" Then
        MkDir outputPath
    End If
    
    csvFileName = baseFileName & "_j.csv"
    
    ' 完全な出力ファイルパスの生成
    outputPath = outputPath & "\" & csvFileName

    
    ' 一時的に作成したシートをCSVファイルとして保存
    destSheet.SaveAs Filename:=outputPath, FileFormat:=xlCSV, Local:=True
    
    ' 一時シートを削除（ユーザーに確認なしで）
    Application.DisplayAlerts = False
    destSheet.Delete
    Application.DisplayAlerts = True
    
    MsgBox "CSVファイルが保存されました: " & outputPath
End Sub
