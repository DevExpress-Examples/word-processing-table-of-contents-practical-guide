Imports DevExpress.Portable.Input
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Drawing

Namespace RichEditTOCGeneration
    Module Program
        Sub Main()
            Console.WriteLine("Select an approach to generate a TOC:" & vbCrLf &
                              "Based on Styles - enter 1" & vbCrLf &
                              "Based on outline levels - enter 2" & vbCrLf &
                              "Based on TC fields - enter 3" & vbCrLf)
            Dim answer As String = Console.ReadLine()
            Dim documentName As String = ""

            Select Case answer
                Case "1"
                    documentName = ApplyStyles()
                Case "2"
                    documentName = AssignOutlineLevels()
                Case "3"
                    documentName = AddTCFields()
            End Select

            Dim p As New Process()
            p.StartInfo = New ProcessStartInfo(documentName) With {
                .UseShellExecute = True
            }
            p.Start()
        End Sub

        Private Function ApplyStyles() As String
            Using wordProcessor As New RichEditDocumentServer()
                wordProcessor.Options.Hyperlinks.ModifierKeys = PortableKeys.None
                wordProcessor.LoadDocument("Employees.rtf")
                Dim document As Document = wordProcessor.Document
                document.BeginUpdate()

                SearchForTOCEntries(document, Sub(location, level)
                                                  document.Paragraphs.[Get](location).Style = GetStyleForLevel(document, level)
                                              End Sub)

                InsertTOC(document, "\h", True)
                document.EndUpdate()
                Dim documentName As String = "Employees_with_Styles_TOC.docx"
                wordProcessor.SaveDocument(documentName, DocumentFormat.OpenXml)
                Console.WriteLine(documentName & " is created")
                Return documentName
            End Using
        End Function

        Private Function AssignOutlineLevels() As String
            Using wordProcessor As New RichEditDocumentServer()
                wordProcessor.Options.Hyperlinks.ModifierKeys = PortableKeys.None
                wordProcessor.LoadDocument("Employees.rtf")
                Dim document As Document = wordProcessor.Document
                document.BeginUpdate()

                SearchForTOCEntries(document, Sub(location, level)
                                                  document.Paragraphs.[Get](location).OutlineLevel = level
                                              End Sub)

                InsertTOC(document, "\h \u", True)
                document.EndUpdate()
                Dim documentName As String = "Employees_with_Outlines_TOC.docx"
                wordProcessor.SaveDocument(documentName, DocumentFormat.OpenXml)
                Console.WriteLine(documentName & " is created")
                Return documentName
            End Using
        End Function

        Private Function AddTCFields() As String
            Using wordProcessor As New RichEditDocumentServer()
                wordProcessor.Options.Hyperlinks.ModifierKeys = PortableKeys.None
                wordProcessor.LoadDocument("Employees.rtf")
                Dim document As Document = wordProcessor.Document
                document.BeginUpdate()

                SearchForTOCEntries(document, Function(ByVal location As DocumentPosition, ByVal level As Integer)
                                                  document.Fields.Create(location, String.Format("TC ""{0}"" \f {1} \l {2}",
                                                  document.GetText(document.Paragraphs.[Get](location).Range), "defaultGroup", level))
                                              End Function)

                InsertTOC(document, "\h \f defaultGroup", True)
                document.Fields.Update()
                document.EndUpdate()
                Dim documentName As String = "Employees_with_TCFields_TOC.docx"
                wordProcessor.SaveDocument(documentName, DocumentFormat.OpenXml)
                Console.WriteLine(documentName & " is created")
                Return documentName
            End Using
        End Function

        Private Sub SearchForTOCEntries(document As Document, callback As Action(Of DocumentPosition, Integer))
            For i As Integer = 0 To document.Paragraphs.Count - 1
                Dim range As DocumentRange = document.CreateRange(document.Paragraphs(i).Range.Start, 1)
                Dim cp As CharacterProperties = document.BeginUpdateCharacters(range)
                Dim level As Integer = 0

                If cp.FontSize.Equals(14.0F) Then
                    level = 1
                ElseIf cp.FontSize.Equals(13.0F) Then
                    level = 2
                ElseIf cp.FontSize.Equals(11.0F) Then
                    level = 3
                End If

                document.EndUpdateCharacters(cp)

                If level <> 0 Then
                    callback(range.Start, level)
                End If
            Next
        End Sub

        Private Sub InsertTOC(document As Document, switches As String, insertHeading As Boolean)
            If insertHeading Then
                InsertContentHeading(document)
            End If

            Dim field As Field = document.Fields.Create(document.Paragraphs(If(insertHeading, 1, 0)).Range.Start, "TOC " & switches)
            Dim cp As CharacterProperties = document.BeginUpdateCharacters(field.Range)
            cp.Bold = False
            cp.FontSize = 12
            cp.ForeColor = Color.Blue
            document.EndUpdateCharacters(cp)
            document.InsertSection(field.Range.End)
            field.Update()
        End Sub

        Private Sub InsertContentHeading(document As Document)
            Dim range As DocumentRange = document.InsertText(document.Range.Start, "Contents" & vbCrLf)
            Dim cp As CharacterProperties = document.BeginUpdateCharacters(range)
            cp.FontSize = 18
            cp.ForeColor = Color.DarkBlue
            document.EndUpdateCharacters(cp)
            Dim paragraph As Paragraph = document.Paragraphs(0)
            paragraph.Alignment = ParagraphAlignment.Center
            paragraph.Style = document.ParagraphStyles("Normal")
            paragraph.OutlineLevel = 0
        End Sub

        Private Function GetStyleForLevel(document As Document, level As Integer) As ParagraphStyle
            Dim styleName As String = "Paragraph Level " & level.ToString()
            Dim paragraphStyle As ParagraphStyle = document.ParagraphStyles(styleName)

            If paragraphStyle Is Nothing Then
                paragraphStyle = document.ParagraphStyles.CreateNew()
                paragraphStyle.Name = styleName
                paragraphStyle.Parent = document.ParagraphStyles("Normal")
                paragraphStyle.OutlineLevel = level
                document.ParagraphStyles.Add(paragraphStyle)
            End If

            Return paragraphStyle
        End Function
    End Module
End Namespace