' ===================================================================
' PowerPoint 自动化格式宏
' 功能:
' 1. AddProgressBar: 在幻灯片【最下方】添加一个动态进度条。
' 2. AddSectionNamesToHeader: 在顶部创建完整的 Beamer 风格导航。
' 3. UpdatePageFormat: 统一右下角的页码格式为 "当前页 / 总页数"。
' 4. RunAllFunctions: 一键运行以上所有功能。
'
' 作者: Gemini（乙方）& ZilongWang (甲方)
' 日期: 2025-10-14
' 兼容性: 已针对 macOS 和 Windows 调整。
' ===================================================================

' Subroutine 1: 添加底部进度条 (已修改)
Sub AddProgressBar()
    Dim X As Long
    Dim S As shape
    Dim slideHeight As Single
    
    On Error Resume Next
    With ActivePresentation
        slideHeight = .PageSetup.slideHeight
        
        For X = 3 To .Slides.Count - 1
            On Error Resume Next
            Do
                .Slides(X).Shapes("PB").Delete
            Loop Until .Slides(X).Shapes("PB") Is Nothing
            Do
                .Slides(X).Shapes("PC").Delete
            Loop Until .Slides(X).Shapes("PC") Is Nothing
            On Error GoTo 0
            
            ' 绘制背景进度条 (灰色)，Y 坐标已修改为页面底部
            Set S = .Slides(X).Shapes.AddLine(-1, slideHeight - 2, .PageSetup.SlideWidth + 1, slideHeight - 2)
            S.Line.Weight = 3
            S.Line.ForeColor.RGB = RGB(205, 205, 205)
            S.Name = "PB"
            
            ' 绘制当前进度条 (蓝色)，Y 坐标已修改为页面底部
            Set S = .Slides(X).Shapes.AddLine(-1, slideHeight - 2, (X - 2) * .PageSetup.SlideWidth / (.Slides.Count - 3) + 1, slideHeight - 2)
            S.Line.Weight = 3
            S.Line.ForeColor.RGB = RGB(50, 100, 200)
            S.Name = "PC"
        Next X
    End With
End Sub

' Subroutine 2: 创建完整的顶部导航系统 (标题 + 圆圈)
Sub AddSectionNamesToHeader()
    Dim sld As slide
    Dim headerShape As shape, circleShape As shape, sepShape As shape
    Dim i As Long, j As Long
    
    ' --- 数据准备 ---
    Dim sectionNames As New Collection
    Dim sectionStartSlides As New Collection
    Dim sectionSlideCounts As New Collection
    
    ' 收集所有章节的名称、起始幻灯片索引和幻灯片数量
    For i = 1 To ActivePresentation.SectionProperties.Count
        sectionNames.Add ActivePresentation.SectionProperties.Name(i)
        sectionStartSlides.Add ActivePresentation.SectionProperties.FirstSlide(i), ActivePresentation.SectionProperties.Name(i)
        sectionSlideCounts.Add ActivePresentation.SectionProperties.SlidesCount(i), ActivePresentation.SectionProperties.Name(i)
    Next i
    
    ' --- 筛选章节 (移除目录、首页和尾页) ---
    For i = sectionNames.Count To 1 Step -1
        If StrComp(sectionNames(i), "目录", vbTextCompare) = 0 Then
            sectionNames.Remove i
        End If
    Next i
    If sectionNames.Count > 0 Then sectionNames.Remove 1
    If sectionNames.Count > 0 Then sectionNames.Remove sectionNames.Count

    ' --- 遍历每张幻灯片，绘制导航元素 ---
    For Each sld In ActivePresentation.Slides
        ' 清理旧的导航元素
        For i = sld.Shapes.Count To 1 Step -1
            If Left(sld.Shapes(i).Name, 17) = "HeaderSectionName" Or _
               Left(sld.Shapes(i).Name, 15) = "HeaderSeparator" Or _
               Left(sld.Shapes(i).Name, 17) = "BeamerSlideCircle" Then
                sld.Shapes(i).Delete
            End If
        Next i

        ' 获取当前幻灯片所在的章节名称
        Dim currentSectionName As String
        currentSectionName = ActivePresentation.SectionProperties.Name(sld.sectionIndex)
        
        If sectionNames.Count > 0 Then
            ' 计算每个章节标题所占的宽度
            Dim portion As Single
            portion = ActivePresentation.PageSetup.SlideWidth / sectionNames.Count

            ' 循环绘制每个章节的标题和其下方的圆圈
            For i = 1 To sectionNames.Count
                Dim sectionTitle As String
                sectionTitle = sectionNames(i)
                
                ' 1. 绘制章节标题
                Set headerShape = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, (i - 1) * portion, 0, portion, 12)
                headerShape.Name = "HeaderSectionName" & i
                With headerShape.TextFrame.TextRange
                    .Text = sectionTitle
                    .Font.Size = 9
                    .Font.NameFarEast = "黑体"
                    .Font.Name = "Times New Roman"
                    .ParagraphFormat.Alignment = ppAlignCenter
                    .Font.Color.RGB = RGB(205, 205, 205)
                End With

                ' 高亮当前章节的标题
                If sectionTitle = currentSectionName Then
                    headerShape.TextFrame.TextRange.Font.Bold = msoTrue
                    headerShape.TextFrame.TextRange.Font.Color.RGB = RGB(240, 180, 50)
                End If
                
                ' 为标题添加超链接
                Dim sectionStartIndex As Long
                sectionStartIndex = sectionStartSlides(sectionTitle)
                With headerShape.ActionSettings(ppMouseClick)
                    .Action = ppActionHyperlink
                    .Hyperlink.SubAddress = ActivePresentation.Slides(sectionStartIndex).SlideID & "," & sectionStartIndex & "," & ActivePresentation.Slides(sectionStartIndex).Name
                End With

                ' 2. 在标题下方绘制本章节的导航圆圈
                Dim slidesInSection As Long
                slidesInSection = sectionSlideCounts(sectionTitle)
                
                Dim circleDiameter As Single, circleSpacing As Single, verticalPos As Single
                circleDiameter = 5
                circleSpacing = 4
                verticalPos = 16 ' 调整垂直位置，使其在标题下方
                
                ' 计算这组圆圈的起始位置，使其在标题 portion 内居中
                Dim circlesTotalWidth As Single, circlesStartLeft As Single
                circlesTotalWidth = (slidesInSection * circleDiameter) + ((slidesInSection - 1) * circleSpacing)
                circlesStartLeft = ((i - 1) * portion) + (portion - circlesTotalWidth) / 2
                
                For j = 1 To slidesInSection
                    Dim targetSlideIndex As Long
                    targetSlideIndex = sectionStartIndex + j - 1
                    
                    Set circleShape = sld.Shapes.AddShape(msoShapeOval, circlesStartLeft + ((j - 1) * (circleDiameter + circleSpacing)), verticalPos, circleDiameter, circleDiameter)
                    circleShape.Name = "BeamerSlideCircle" & i & "_" & j

                    ' 为圆圈添加超链接
                    With circleShape.ActionSettings(ppMouseClick)
                        .Action = ppActionHyperlink
                        .Hyperlink.SubAddress = ActivePresentation.Slides(targetSlideIndex).SlideID & "," & targetSlideIndex & "," & ActivePresentation.Slides(targetSlideIndex).Name
                    End With
                    
                    ' 高亮当前幻灯片对应的圆圈
                    If sld.SlideIndex = targetSlideIndex Then
                        circleShape.Fill.Visible = msoTrue
                        circleShape.Fill.ForeColor.RGB = RGB(180, 180, 180)
                        circleShape.Line.Visible = msoFalse
                    Else
                        circleShape.Fill.Visible = msoFalse
                        circleShape.Line.Visible = msoTrue
                        circleShape.Line.ForeColor.RGB = RGB(205, 205, 205)
                        circleShape.Line.Weight = 1
                    End If
                Next j

                ' 3. 绘制分隔符
                If i < sectionNames.Count Then
                    Set sepShape = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, i * portion - 10, 6, 20, 12)
                    sepShape.Name = "HeaderSeparator" & i
                    With sepShape.TextFrame.TextRange
                        .Text = "|"
                        .Font.Size = 10
                        .Font.NameFarEast = "黑体"
                        .Font.Name = "Times New Roman"
                        .ParagraphFormat.Alignment = ppAlignCenter
                        .Font.Color.RGB = RGB(205, 205, 205)
                    End With
                End If
            Next i
        End If
    Next sld
End Sub


' Subroutine 3: 更新页码格式
Sub UpdatePageFormat()
    Dim sld As slide
    Dim shp As shape
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                    With shp
                        .TextFrame.TextRange.Text = sld.SlideIndex & " / " & ActivePresentation.Slides.Count
                        .Width = 60
                        .Left = ActivePresentation.PageSetup.SlideWidth - .Width
                        With .TextFrame.TextRange.Font
                            .NameFarEast = "黑体"
                            .Name = "Times New Roman"
                            .Size = 14
                            .Color.RGB = RGB(25, 25, 25)
                        End With
                    End With
                End If
            End If
        Next shp
    Next sld
End Sub

' 主控宏: 运行所有格式化功能
Sub RunAllFunctions()
    AddProgressBar
    AddSectionNamesToHeader
    UpdatePageFormat
End Sub
