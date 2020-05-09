Option Explicit

Dim fsObj: Set fsObj = CreateObject("Scripting.FileSystemObject")

Dim nowStr: nowStr = Now()
Dim nowTimeStampStr: nowTimeStampStr = Year(nowStr) & _
                                       Right("0" & Month(nowStr) , 2) & _
                                       Right("0" & Day(nowStr) , 2) & _
                                       Right("0" & Hour(nowStr) , 2) & _
                                       Right("0" & Minute(nowStr) , 2) & _
                                       Right("0" & Second(nowStr) , 2)  ' yyyymmddhhmmssフォーマットの現在日時
Dim searchResultFilePath: searchResultFilePath = ".\SearchResult_" & nowTimeStampStr & ".tsv"  ' 検索結果出力先TSVファイルパス
Dim searchResultFileWriterStream
Dim hasSearched: hasSearched = False

Dim regExReplaceOrgPatternObj
Dim replaceDestStr
Dim replaceTgtFilePathList: Set replaceTgtFilePathList = CreateObject("System.Collections.ArrayList")  ' 置換対象ファイルパスリスト

Call Main()

Set fsObj = Nothing
Set replaceTgtFilePathList = Nothing


Sub Main()
    MsgBox "PowerPointが起動中の場合は、終了してから操作を行なって下さい。"
    
    Dim startDirPath: startDirPath = InputBox("検索/置換したい文字列を含むファイルが存在するトップフォルダのパスを入力して下さい。")
    If fsObj.FolderExists(startDirPath) = False Then
        MsgBox "不正なフォルダです"
        WScript.Quit
    End If
    Dim startDirObj: Set startDirObj = fsObj.GetFolder(startDirPath)
    
    Dim replaceOrgStrPattern: replaceOrgStrPattern = InputBox("検索する/置換元とする文字列の正規表現を入力して下さい。" & vbCrLf & _
                                                              "* 正規表現として扱われますので注意してください。" & vbCrLf & _
                                                              "* 大文字小文字の区別をしません。")
    If replaceOrgStrPattern = "" Then
        WScript.Quit  ' 置換元が指定されなければ終了
    End If
    Set regExReplaceOrgPatternObj = New RegExp
    regExReplaceOrgPatternObj.Pattern    = replaceOrgStrPattern
    regExReplaceOrgPatternObj.IgnoreCase = True  ' 大文字小文字区別しない
    regExReplaceOrgPatternObj.Global     = True  ' 文字列内のすべてを対象に検索
    
    replaceDestStr = InputBox("置換後の文字列を入力して下さい。")
    
    If MsgBox("トップフォルダと下位フォルダから置換元文字正規表現を含むPower Pointファイル・箇所を置換前にファイル出力して先に確認しますか？", vbYesNo + vbQuestion) = vbYes Then
        Call SearchEntry(startDirObj)
    End If
    
    If MsgBox("トップフォルダと下位フォルダから置換元文字正規表現を含むPower Pointファイルの文字列を置換しますか？", vbYesNo + vbQuestion) = vbYes Then
        Call ReplaceEntry(startDirObj)
    End If
    
    Set startDirObj = Nothing
    Set regExReplaceOrgPatternObj = Nothing
    
    MsgBox "処理が終了しました。"
End Sub

Sub SearchEntry(dirObj)
    Call OpenSearchResultFile()
    Call WriteAppendToSearchResultFile("Folder Path" & vbTab & "File Name" & vbTab & "Replace Target Part Text")
    
    Call SearchRecursively(dirObj)
    hasSearched = True
    
    Call CloseSearchResultFile()
End Sub

Sub SearchRecursively(dirObj)
    Dim fileObj
    For Each fileObj in dirObj.Files
        Dim fileExt: fileExt = fsObj.GetExtensionName(fileObj.Name)
        If LCase(fileExt) = "ppt" Or LCase(fileExt) = "pptx" Then
            Call SearchWithinFile(fileObj)
        End If
    Next
    
    Dim subDirObj
    For Each subDirObj in dirObj.SubFolders
        Call SearchRecursively(subDirObj)
    Next
End Sub

Sub SearchWithinFile(fileObj)
    Dim pptAppObj: Set pptAppObj = CreateObject("PowerPoint.Application")
    pptAppObj.Visible = True
    Call pptAppObj.Presentations.Open(fileObj.Path)
    
    Dim containsReplaceOrgStrPattern: containsReplaceOrgStrPattern = False
    
    Dim slide
    For Each slide In pptAppObj.ActiveWindow.Parent.Slides
        Dim shape
        For Each shape In slide.Shapes
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    Dim shapeTxt: shapeTxt = shape.TextFrame.TextRange.Text
                    If regExReplaceOrgPatternObj.Test(shapeTxt) = True Then
                        containsReplaceOrgStrPattern = True
                        Call WriteAppendToSearchResultFile(fileObj.ParentFolder.Path & vbTab & fileObj.Name & vbTab & shapeTxt)
                    End If
                End If
            End If
        Next
    Next
    
    If containsReplaceOrgStrPattern Then
        Call replaceTgtFilePathList.Add(fileObj.Path)
    End If
    
    pptAppObj.Quit
    Set pptAppObj = Nothing
End Sub

Sub OpenSearchResultFile()
    ' 追加書き込みモード, ファイルがなければ作成, 文字コード:UTF-16
    Set searchResultFileWriterStream = fsObj.OpenTextFile(searchResultFilePath, 8, True, -1)
End Sub

Sub CloseSearchResultFile()
    searchResultFileWriterStream.Close
    Set searchResultFileWriterStream = Nothing
End Sub

Sub WriteAppendToSearchResultFile(outputStr)
    Call searchResultFileWriterStream.WriteLine(outputStr)
End Sub


Sub ReplaceEntry(dirObj)
    If hasSearched = False Then
        Call SearchEntry(dirObj)
    End If
    
    Dim pptAppObj: Set pptAppObj = CreateObject("PowerPoint.Application")
    pptAppObj.Visible = True
    
    Dim replaceTgtFilePath
    For Each replaceTgtFilePath In replaceTgtFilePathList
        Dim replaceTgtPresenFile: Set replaceTgtPresenFile = pptAppObj.Presentations.Open(replaceTgtFilePath)
        
        Dim slide
        For Each slide In pptAppObj.ActiveWindow.Parent.Slides
            Dim shape
            For Each shape In slide.Shapes
                If shape.HasTextFrame Then
                    If shape.TextFrame.HasText Then
                        shape.TextFrame.TextRange.Text = regExReplaceOrgPatternObj.Replace(shape.TextFrame.TextRange.Text, replaceDestStr)
                    End If
                End If
            Next
        Next
        
        replaceTgtPresenFile.Save
        replaceTgtPresenFile.Close
    Next
    
    pptAppObj.Quit
    Set pptAppObj = Nothing
End Sub
