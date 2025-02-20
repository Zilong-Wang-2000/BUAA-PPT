Sub AddProgressBar()
    On Error Resume Next
    With ActivePresentation
        For X = 3 To .Slides.Count - 1 ' Skip the first and last slide
            On Error Resume Next
            Do
                .Slides(X).Shapes("PB").Delete
            Loop Until .Slides(X).Shapes("PB") Is Nothing
            Do
                .Slides(X).Shapes("PC").Delete
            Loop Until .Slides(X).Shapes("PC") Is Nothing
            On Error GoTo 0
            Set S = .Slides(X).Shapes.AddLine( _
            0, 0, .PageSetup.SlideWidth, 0) ' Specify the start and end points of the line, with the top-left corner as the origin
            S.Line.Weight = 6 ' Set the stroke width, the stroke defaults to center outwards
            S.Line.ForeColor.RGB = RGB(205, 205, 205) ' Set the stroke color
            S.Name = "PB"
            Set S = .Slides(X).Shapes.AddLine( _
            0, 0, (X - 2) * .PageSetup.SlideWidth / (.Slides.Count - 3), 0) ' Specify the start and end points of the line, with the top-left corner as the origin
            S.Line.Weight = 6 ' Set the stroke width, the stroke defaults to center outwards
            S.Line.ForeColor.RGB = RGB(255, 255, 0) ' Set the stroke color
            S.Name = "PC"
        Next X
    End With
End Sub

Sub AddSectionNamesToHeader()
    Dim slide As slide
    Dim sectionIndex As Integer
    Dim sectionName As String
    Dim headerShape As Shape
    Dim sectionNames As Collection
    Dim i As Integer
    Dim currentSectionName As String
    Dim sectionSlideIndex As Integer

    ' Collect all section names
    Set sectionNames = New Collection
    For i = 1 To ActivePresentation.SectionProperties.Count
        sectionNames.Add ActivePresentation.SectionProperties.Name(i)
    Next i

    ' Remove any section named "目录"
    For i = sectionNames.Count To 1 Step -1
        If StrComp(sectionNames(i), "目录", vbTextCompare) = 0 Then
            sectionNames.Remove i
        End If
    Next i

    ' Remove the first and last items, if present
    If sectionNames.Count > 0 Then
        sectionNames.Remove 1
    End If

    If sectionNames.Count > 0 Then
        sectionNames.Remove sectionNames.Count
    End If

    ' Create a dictionary to store the starting slide index for each section
    Dim sectionStartSlides As Object
    Set sectionStartSlides = CreateObject("Scripting.Dictionary")

    ' Populate the dictionary with section names and their starting slide index
    For i = 1 To ActivePresentation.SectionProperties.Count
        sectionStartSlides.Add ActivePresentation.SectionProperties.Name(i), ActivePresentation.SectionProperties.FirstSlide(i)
        ' print section name and starting slide index
        Debug.Print ActivePresentation.SectionProperties.Name(i) & " " & ActivePresentation.SectionProperties.FirstSlide(i)
    Next i

    ' Add section names to each slide
    For Each slide In ActivePresentation.Slides 
        sectionIndex = slide.sectionIndex
        currentSectionName = ActivePresentation.SectionProperties.Name(sectionIndex)

        ' Clear existing header shapes with header section names
        For Each headerShape In slide.Shapes
            If Left(headerShape.Name, 17) = "HeaderSectionName" Or Left(headerShape.Name, 15) = "HeaderSeparator" Then
                Debug.Print "Delete " & headerShape.Name
                headerShape.Delete
            End If
        Next headerShape

        ' Add section names to the header horizontally
        Dim portion As Single
        If sectionNames.Count > 0 Then
            portion = ActivePresentation.PageSetup.SlideWidth / sectionNames.Count
            For i = 1 To sectionNames.Count
                Set headerShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, (i - 1) * portion, 6, portion, 10)
                headerShape.Name = "HeaderSectionName" & i
                headerShape.TextFrame.TextRange.Font.NameFarEast = "黑体"
                headerShape.TextFrame.TextRange.Font.Name = "Times New Roman"
                headerShape.TextFrame.TextRange.Font.Size = 16
                headerShape.TextFrame.TextRange.Text = sectionNames(i)
                headerShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                headerShape.TextFrame.TextRange.Font.Color.RGB = RGB(205, 205, 205)

                If sectionNames(i) = currentSectionName Then
                    headerShape.TextFrame.TextRange.Font.Bold = msoTrue
                    headerShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 0)
                End If

                ' Add hyperlink to section name
                sectionSlideIndex = sectionStartSlides(sectionNames(i))
                headerShape.ActionSettings(ppMouseClick).Hyperlink.Address = ""
                headerShape.ActionSettings(ppMouseClick).Action = ppActionHyperlink
                headerShape.ActionSettings(ppMouseClick).Hyperlink.SubAddress = ActivePresentation.Slides(sectionSlideIndex).SlideID & "," & sectionSlideIndex & "," & ActivePresentation.Slides(sectionSlideIndex).Name
                headerShape.TextFrame.TextRange.Font.Underline = msoFalse

                ' Add a separator "|" between titles
                If i < sectionNames.Count Then
                    Dim sepShape As Shape
                    Set sepShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, i * portion - 10, 6, 20, 10)
                    sepShape.Name = "HeaderSeparator" & i
                    sepShape.TextFrame.TextRange.Font.NameFarEast = "黑体"
                    sepShape.TextFrame.TextRange.Font.Name = "Times New Roman"
                    sepShape.TextFrame.TextRange.Font.Size = 16
                    sepShape.TextFrame.TextRange.Text = "|"
                    sepShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                    sepShape.TextFrame.TextRange.Font.Color.RGB = RGB(205, 205, 205)
                End If
            Next i
        End If
    Next slide
End Sub


' print all shape names in the slide
Sub UpdatePageFormat()
    Dim slide As slide
    Dim shape As Shape

    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            Debug.Print shape.Name
            If shape.Type = msoPlaceholder Then
                If shape.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                    shape.TextFrame.TextRange.Text = "" & slide.SlideIndex & " / " & ActivePresentation.Slides.Count & ""
                    Debug.Print "Slide Number Placeholder" & shape.Name & " " & shape.TextFrame.TextRange.Text
                    shape.Width = 60
                    shape.Left = ActivePresentation.PageSetup.SlideWidth - 60
                    shape.TextFrame.TextRange.Font.NameFarEast = "黑体"
                    shape.TextFrame.TextRange.Font.Name = "Times New Roman"
                    shape.TextFrame.TextRange.Font.Size = 14
                    shape.TextFrame.TextRange.Font.Color.RGB = RGB(25, 25, 25)
                End If
            End If
        Next shape
        
        Debug.Print "----------------------"
    Next slide
End Sub

' 在页面的左下角加入上一页，下一页，首页，目录，尾页的链接
Sub AddNavigationLinks()
    Dim slide As slide
    Dim navShape As Shape
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim linkNames As Variant
    Dim linkActions As Variant
    Dim linkIcons As Variant
    Dim i As Integer

    linkNames = Array("上一页", "下一页", "首页", "目录", "尾页")
    linkActions = Array(ppActionPreviousSlide, ppActionNextSlide, ppActionFirstSlide, ppActionLastSlide, ppActionLastSlide)
    linkIcons = Array(ChrW(9194), ChrW(9193), ChrW(9198), ChrW(9208), ChrW(9197))

    For Each slide In ActivePresentation.Slides
        Debug.Print slide.Name

        slideWidth = ActivePresentation.PageSetup.SlideWidth
        slideHeight = ActivePresentation.PageSetup.SlideHeight

        ' Clear existing navigation shapes
        For Each navShape In slide.Shapes
            If Left(navShape.Name, 14) = "NavigationLink" Then
                navShape.Delete
            End If
        Next navShape

        For i = LBound(linkNames) To UBound(linkNames)
            Set navShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 10 + (i * 30), slideHeight - 30, 30, 16)
            navShape.Name = "NavigationLink" & i
            navShape.TextFrame.TextRange.Text = linkIcons(i) 
            navShape.TextFrame.TextRange.Font.Name = "Segoe UI Symbol"
            navShape.TextFrame.TextRange.Font.Size = 12

            navShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
            navShape.TextFrame.TextRange.Font.Color.RGB = RGB(205, 205, 205)
            navShape.ActionSettings(ppMouseClick).Action = linkActions(i)

            If linkNames(i) = "目录" Then
                navShape.ActionSettings(ppMouseClick).Hyperlink.Address = ""
                navShape.ActionSettings(ppMouseClick).Hyperlink.SubAddress = ActivePresentation.Slides(2).SlideID & ",2," & ActivePresentation.Slides(2).Name
                navShape.TextFrame.TextRange.Font.Underline = msoFalse
            End If
        Next i
    Next slide
End Sub

Sub RunAllFunctions()
    AddProgressBar
    AddSectionNamesToHeader
    UpdatePageFormat
    AddNavigationLinks
End Sub