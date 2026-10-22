Attribute VB_Name = "修改涂色"
Option Explicit

'=========================================================
' Engineering Change Marker V3.1
' 工程文件修改标记插件
'
' 功能：
' 1. Rev.A ~ Rev.G 七种版本颜色
' 2. 当前版本保存在 Word 文档自身
' 3. Word 关闭后重新打开仍可读取
' 4. 使用 Word 修订识别新增 / 删除 / 替换
' 5. 完成修改后自动接受修订
' 6. 新增或替换后的文字永久使用高亮颜色
' 7. 删除内容直接正式删除
' 8. 支持清除选中高亮 / 全文高亮
'=========================================================


Private Const PROP_REVISION As String = "ECM_CurrentRevision"


'=========================================================
' 打开修改面板
'=========================================================
Sub 打开修改面板()

    frmChangeMarker.Show vbModeless

End Sub


'=========================================================
' 获取当前版本
'=========================================================
Function 获取当前版本() As String

    On Error GoTo NoProperty

    获取当前版本 = _
        ActiveDocument.CustomDocumentProperties(PROP_REVISION).Value

    Exit Function


NoProperty:

    获取当前版本 = ""

End Function


'=========================================================
' 保存当前版本到当前 Word 文件
'=========================================================
Sub 保存当前版本(ByVal RevName As String)

    On Error GoTo CreateProperty

    ActiveDocument.CustomDocumentProperties(PROP_REVISION).Value = RevName

    Exit Sub


CreateProperty:

    Err.Clear

    ActiveDocument.CustomDocumentProperties.Add _
        Name:=PROP_REVISION, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=RevName

End Sub


'=========================================================
' 设置当前版本
'=========================================================
Sub 设置版本(ByVal RevName As String)

    保存当前版本 RevName

End Sub


'=========================================================
' 根据版本号返回对应高亮颜色
'
' Rev.A = 绿色
' Rev.B = 黄色
' Rev.C = 青色
' Rev.D = 粉红
' Rev.E = 蓝色
' Rev.F = 红色
' Rev.G = 灰色
'=========================================================
Function 获取版本颜色(ByVal RevName As String) As WdColorIndex

    Select Case UCase(RevName)

        Case "REV.A"
            获取版本颜色 = wdBrightGreen

        Case "REV.B"
            获取版本颜色 = wdYellow

        Case "REV.C"
            获取版本颜色 = wdTurquoise

        Case "REV.D"
            获取版本颜色 = wdPink

        Case "REV.E"
            获取版本颜色 = wdBlue

        Case "REV.F"
            获取版本颜色 = wdRed

        Case "REV.G"
            获取版本颜色 = wdGray25

        Case Else
            获取版本颜色 = wdBrightGreen

    End Select

End Function


'=========================================================
' 开始修改
'=========================================================
Sub 开始工程修改()

    Dim RevName As String

    RevName = 获取当前版本()


    If RevName = "" Then

        MsgBox _
            "请先选择本次修改对应的版本。", _
            vbExclamation, _
            "工程修改标记"

        frmChangeMarker.Show vbModeless

        Exit Sub

    End If


    '开启 Word 修订
    ActiveDocument.TrackRevisions = True

End Sub


'=========================================================
' 完成修改
'=========================================================
Sub 完成工程修改()

    Dim RevName As String
    Dim MarkColor As WdColorIndex

    Dim Story As Range
    Dim CurrentStory As Range

    Dim rev As Revision
    Dim MarkRange As Range

    Dim i As Long

    Dim CountInsert As Long
    Dim CountDelete As Long
    Dim TotalRevisions As Long


    '=====================================================
    ' 获取当前版本
    '=====================================================
    RevName = 获取当前版本()


    If RevName = "" Then

        MsgBox _
            "尚未选择当前修改版本。", _
            vbExclamation, _
            "工程修改标记"

        Exit Sub

    End If


    MarkColor = 获取版本颜色(RevName)


    '=====================================================
    ' 首先统计整个 Word 文件中的修订
    '
    ' 包括：
    ' 正文
    ' 页眉
    ' 页脚
    ' 其它 StoryRange
    '=====================================================
    TotalRevisions = 0


    For Each Story In ActiveDocument.StoryRanges

        Set CurrentStory = Story

        Do

            TotalRevisions = _
                TotalRevisions + CurrentStory.Revisions.Count

            Set CurrentStory = CurrentStory.NextStoryRange

        Loop Until CurrentStory Is Nothing

    Next Story


    '=====================================================
    ' 没有发现修订
    '=====================================================
    If TotalRevisions = 0 Then

        ActiveDocument.TrackRevisions = False

        MsgBox _
            "没有发现需要处理的修订内容。" & vbCrLf & vbCrLf & _
            "已检查：" & vbCrLf & _
            "· 正文" & vbCrLf & _
            "· 页眉" & vbCrLf & _
            "· 页脚", _
            vbInformation, _
            "工程修改标记"

        Exit Sub

    End If


    Application.ScreenUpdating = False


    '=====================================================
    ' 遍历 Word 中所有 StoryRange
    '=====================================================
    For Each Story In ActiveDocument.StoryRanges

        Set CurrentStory = Story


        Do

            '=================================================
            ' 每一个 StoryRange 的修订必须倒序处理
            '=================================================
            For i = CurrentStory.Revisions.Count To 1 Step -1

                Set rev = CurrentStory.Revisions(i)


                Select Case rev.Type


                    '=========================================
                    ' 新增文字
                    '=========================================
                    Case wdRevisionInsert


                        '先复制新增文字的 Range
                        Set MarkRange = rev.Range.Duplicate


                        '接受修订
                        rev.Accept


                        '-------------------------------------
                        ' 排除段落标记、单元格结束符
                        '-------------------------------------
                        Do While MarkRange.End > MarkRange.Start

                            If Right$(MarkRange.Text, 1) = vbCr _
                               Or Right$(MarkRange.Text, 1) = Chr(7) Then

                                MarkRange.MoveEnd _
                                    Unit:=wdCharacter, _
                                    Count:=-1

                            Else

                                Exit Do

                            End If

                        Loop


                        '-------------------------------------
                        ' 添加本版本高亮
                        '-------------------------------------
                        If MarkRange.End > MarkRange.Start Then

                            MarkRange.HighlightColorIndex = _
                                MarkColor

                        End If


                        CountInsert = CountInsert + 1


                    '=========================================
                    ' 删除文字
                    '=========================================
                    Case wdRevisionDelete

                        rev.Accept

                        CountDelete = CountDelete + 1


                    '=========================================
                    ' 其它类型修订
                    '=========================================
                    Case Else

                        rev.Accept

                End Select

            Next i


            '进入下一个相同类型的 StoryRange
            '例如下一节的页眉
            Set CurrentStory = CurrentStory.NextStoryRange


        Loop Until CurrentStory Is Nothing

    Next Story


    '=====================================================
    ' 完成
    '=====================================================
    ActiveDocument.TrackRevisions = False

    Application.ScreenUpdating = True


    MsgBox _
        "本轮修改已经完成。" & vbCrLf & vbCrLf & _
        "版本：" & RevName & vbCrLf & _
        "新增 / 替换：" & CountInsert & " 处" & vbCrLf & _
        "删除：" & CountDelete & " 处" & vbCrLf & vbCrLf & _
        "正文、页眉及页脚中的修改已处理。", _
        vbInformation, _
        "工程修改完成"

End Sub


'=========================================================
' 显示当前版本
'=========================================================
Sub 显示当前版本()

    Dim RevName As String

    RevName = 获取当前版本()


    If RevName = "" Then

        MsgBox _
            "当前文档尚未指定修改版本。", _
            vbInformation, _
            "当前版本"

    Else

        MsgBox _
            "当前修改版本：" & RevName, _
            vbInformation, _
            "当前版本"

    End If

End Sub


'=========================================================
' 清除选中文字的高亮
'=========================================================
Sub 清除选中内容颜色()

    If Selection.Range.Start = Selection.Range.End Then

        MsgBox _
            "请先选择需要清除颜色的文字。", _
            vbExclamation, _
            "清除版本颜色"

        Exit Sub

    End If


    If MsgBox( _
        "确定清除选中文字的高亮颜色吗？", _
        vbYesNo + vbQuestion, _
        "清除版本颜色") <> vbYes Then

        Exit Sub

    End If


    Selection.Range.HighlightColorIndex = wdNoHighlight


    MsgBox _
        "选中文字的高亮颜色已经清除。", _
        vbInformation, _
        "处理完成"

End Sub


'=========================================================
' 清除全文高亮
' 谨慎使用
'=========================================================
Sub 清除全文颜色()

    Dim Answer As VbMsgBoxResult


    Answer = MsgBox( _
        "确定清除当前文档中的所有文字高亮吗？" & vbCrLf & vbCrLf & _
        "注意：" & vbCrLf & _
        "如果文档原来就存在人工设置的高亮，" & _
        "也会被一起清除。", _
        vbYesNo + vbExclamation, _
        "清除全文高亮")


    If Answer <> vbYes Then Exit Sub


    Application.ScreenUpdating = False


    ActiveDocument.Content.HighlightColorIndex = wdNoHighlight


    Application.ScreenUpdating = True


    MsgBox _
        "正文中的所有高亮已经清除。", _
        vbInformation, _
        "处理完成"

End Sub

