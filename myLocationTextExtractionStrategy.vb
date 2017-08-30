Imports System
Imports System.Collections.Generic
Imports System.Text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser


Namespace MyNSpace

    Public Class myLocationTextExtractionStrategy
        Implements ITextExtractionStrategy

        '* set to true for debugging 

        Private _UndercontentCharacterSpacing = 0
        Private _UndercontentHorizontalScaling = 0
        Private ThisPdfDocFonts As SortedList(Of String, DocumentFont)

        Public Shared DUMP_STATE As Boolean = False

        '* a summary of all found text 

        Private locationalResult As New List(Of TextChunk)()

        '*
        '         * Creates a new text extraction renderer.
        '         

        Public Sub New()
            ThisPdfDocFonts = New SortedList(Of String, DocumentFont)
        End Sub

        '*
        '         * @see com.itextpdf.text.pdf.parser.RenderListener#beginTextBlock()
        '         

        Public Overridable Sub BeginTextBlock() Implements ITextExtractionStrategy.BeginTextBlock
        End Sub

        '*
        '         * @see com.itextpdf.text.pdf.parser.RenderListener#endTextBlock()
        '         

        Public Overridable Sub EndTextBlock() Implements ITextExtractionStrategy.EndTextBlock
        End Sub

        '*
        '         * @param str
        '         * @return true if the string starts with a space character, false if the string is empty or starts with a non-space character
        '         

        Private Function StartsWithSpace(ByVal str As [String]) As Boolean
            If str.Length = 0 Then
                Return False
            End If
            Return str(0) = " "c
        End Function

        '*
        '         * @param str
        '         * @return true if the string ends with a space character, false if the string is empty or ends with a non-space character
        '         

        Private Function EndsWithSpace(ByVal str As [String]) As Boolean
            If str.Length = 0 Then
                Return False
            End If
            Return str(str.Length - 1) = " "c
        End Function

        Public Property UndercontentCharacterSpacing
            Get
                Return _UndercontentCharacterSpacing
            End Get
            Set(ByVal value)
                _UndercontentCharacterSpacing = value
            End Set
        End Property

        Public Property UndercontentHorizontalScaling
            Get
                Return _UndercontentHorizontalScaling
            End Get
            Set(ByVal value)
                _UndercontentHorizontalScaling = value
            End Set
        End Property

        Public Overridable Function GetResultantText() As [String] Implements ITextExtractionStrategy.GetResultantText

            If DUMP_STATE Then
                DumpState()
            End If

            locationalResult.Sort()

            Dim sb As New StringBuilder()
            Dim lastChunk As TextChunk = Nothing

            For Each chunk As TextChunk In locationalResult

                If lastChunk Is Nothing Then
                    sb.Append(chunk.text)
                Else
                    If chunk.SameLine(lastChunk) Then
                        Dim dist As Single = chunk.DistanceFromEndOf(lastChunk)
                        If dist < -chunk.charSpaceWidth Then
                            sb.Append(" "c)
                            ' we only insert a blank space if the trailing character of the previous string wasn't a space, and the leading character of the current string isn't a space
                        ElseIf dist > chunk.charSpaceWidth / 2.0F AndAlso Not StartsWithSpace(chunk.text) AndAlso Not EndsWithSpace(lastChunk.text) Then
                            sb.Append(" "c)
                        End If

                        sb.Append(chunk.text)
                    Else
                        sb.Append(ControlChars.Lf)
                        sb.Append(chunk.text)
                    End If
                End If
                lastChunk = chunk
            Next

            Return sb.ToString()

        End Function

        Public Function GetTextLocations(ByVal pSearchString As String, ByVal pStrComp As System.StringComparison) As List(Of iTextSharp.text.Rectangle)
            Dim FoundMatches As New List(Of iTextSharp.text.Rectangle)
            Dim sb As New StringBuilder()
            Dim ThisLineChunks As List(Of TextChunk) = New List(Of TextChunk)
            Dim bStart As Boolean, bEnd As Boolean
            Dim FirstChunk As TextChunk = Nothing, LastChunk As TextChunk = Nothing
            Dim sTextInUsedChunks As String = vbNullString

            For Each chunk As TextChunk In locationalResult

                If ThisLineChunks.Count > 0 AndAlso Not chunk.SameLine(ThisLineChunks.Last) Then
                    If sb.ToString.IndexOf(pSearchString, pStrComp) > -1 Then
                        Dim sLine As String = sb.ToString

                        'Check how many times the Search String is present in this line:
                        Dim iCount As Integer = 0
                        Dim lPos As Integer
                        lPos = sLine.IndexOf(pSearchString, 0, pStrComp)
                        Do While lPos > -1
                            iCount += 1
                            If lPos + pSearchString.Length > sLine.Length Then Exit Do Else lPos = lPos + pSearchString.Length
                            lPos = sLine.IndexOf(pSearchString, lPos, pStrComp)
                        Loop

                        'Process each match found in this Text line:
                        Dim curPos As Integer = 0
                        For i As Integer = 1 To iCount
                            Dim sCurrentText As String, iFromChar As Integer, iToChar As Integer

                            iFromChar = sLine.IndexOf(pSearchString, curPos, pStrComp)
                            curPos = iFromChar
                            iToChar = iFromChar + pSearchString.Length - 1
                            sCurrentText = vbNullString
                            sTextInUsedChunks = vbNullString
                            FirstChunk = Nothing
                            LastChunk = Nothing

                            'Get first and last Chunks corresponding to this match found, from all Chunks in this line
                            For Each chk As TextChunk In ThisLineChunks
                                sCurrentText = sCurrentText & chk.text

                                'Check if we entered the part where we had found a matching String then get this Chunk (First Chunk)
                                If Not bStart AndAlso sCurrentText.Length - 1 >= iFromChar Then
                                    FirstChunk = chk
                                    bStart = True
                                End If

                                'Keep getting Text from Chunks while we are in the part where the matching String had been found
                                If bStart And Not bEnd Then
                                    sTextInUsedChunks = sTextInUsedChunks & chk.text
                                End If

                                'If we get out the matching String part then get this Chunk (last Chunk)
                                If Not bEnd AndAlso sCurrentText.Length - 1 >= iToChar Then
                                    LastChunk = chk
                                    bEnd = True
                                End If

                                'If we already have first and last Chunks enclosing the Text where our String pSearchString has been found 
                                'then it's time to get the rectangle, GetRectangleFromText Function below this Function, there we extract the pSearchString locations
                                If bStart And bEnd Then
                                    FoundMatches.Add(GetRectangleFromText(FirstChunk, LastChunk, pSearchString, sTextInUsedChunks, iFromChar, iToChar, pStrComp))
                                    curPos = curPos + pSearchString.Length
                                    bStart = False : bEnd = False
                                    Exit For
                                End If
                            Next
                        Next
                    End If
                    sb.Clear()
                    ThisLineChunks.Clear()
                End If
                ThisLineChunks.Add(chunk)
                sb.Append(chunk.text)
            Next

            Return FoundMatches
        End Function

        Private Function GetRectangleFromText(ByVal FirstChunk As TextChunk, ByVal LastChunk As TextChunk, ByVal pSearchString As String, _
                                   ByVal sTextinChunks As String, ByVal iFromChar As Integer, ByVal iToChar As Integer, ByVal pStrComp As System.StringComparison) As iTextSharp.text.Rectangle

            'There are cases where Chunk contains extra text at begining and end, we don't want this text locations, we need to extract the pSearchString location inside
            'for these cases we need to crop this String (left and Right), and measure this excedent at left and right, at this point we don't have any direct way to make a
            'Transformation from text space points to User Space units, the matrix for making this transformation is not accesible from here, so for these special cases when
            'the String needs to be cropped (Left/Right) We'll interpolate between the width from Text in Chunk (we have this value in User Space units), then i'll measure Text corresponding
            'to the same String but in Text Space units, finally from the relation betweeenthese 2 values I get the TransformationValue I need to use for all cases

            'Text Width in User Space Units
            Dim LineRealWidth As Single = LastChunk.PosRight - FirstChunk.PosLeft

            'Text Width in Text Units
            Dim LineTextWidth As Single = GetStringWidth(sTextinChunks, LastChunk.curFontSize, _
                                                         LastChunk.charSpaceWidth, _
                                                         ThisPdfDocFonts.Values.ElementAt(LastChunk.FontIndex))
            'TransformationValue value for Interpolation
            Dim TransformationValue As Single = LineRealWidth / LineTextWidth

            'In the worst case, we'll need to crop left and right:
            Dim iStart As Integer = sTextinChunks.IndexOf(pSearchString, pStrComp)

            Dim iEnd As Integer = iStart + pSearchString.Length - 1

            Dim sLeft As String
            If iStart = 0 Then sLeft = vbNullString Else sLeft = sTextinChunks.Substring(0, iStart)

            Dim sRight As String
            If iEnd = sTextinChunks.Length - 1 Then sRight = vbNullString Else sRight = sTextinChunks.Substring(iEnd + 1, sTextinChunks.Length - iEnd - 1)

            'Measure cropped Text at left:
            Dim LeftWidth As Single = 0
            If iStart > 0 Then
                LeftWidth = GetStringWidth(sLeft, LastChunk.curFontSize, _
                                                  LastChunk.charSpaceWidth, _
                                                  ThisPdfDocFonts.Values.ElementAt(LastChunk.FontIndex))
                LeftWidth = LeftWidth * TransformationValue
            End If

            'Measure cropped Text at right:
            Dim RightWidth As Single = 0
            If iEnd < sTextinChunks.Length - 1 Then
                RightWidth = GetStringWidth(sRight, LastChunk.curFontSize, _
                                                    LastChunk.charSpaceWidth, _
                                                    ThisPdfDocFonts.Values.ElementAt(LastChunk.FontIndex))
                RightWidth = RightWidth * TransformationValue
            End If

            'LeftWidth is the text width at left we need to exclude, FirstChunk.distParallelStart is the distance to left margin, both together will give us this LeftOffset
            Dim LeftOffset As Single = FirstChunk.distParallelStart + LeftWidth
            'RightWidth is the text width at right we need to exclude, FirstChunk.distParallelEnd is the distance to right margin, we substract RightWidth from distParallelEnd to get RightOffset
            Dim RightOffset As Single = LastChunk.distParallelEnd - RightWidth
            'Return this Rectangle
            Return New iTextSharp.text.Rectangle(LeftOffset, FirstChunk.PosBottom, RightOffset, FirstChunk.PosTop)

        End Function

        Private Function GetStringWidth(ByVal str As String, ByVal curFontSize As Single, ByVal pSingleSpaceWidth As Single, ByVal pFont As DocumentFont) As Single
            Dim chars() As Char = str.ToCharArray()
            Dim totalWidth As Single = 0
            Dim w As Single = 0

            For Each c As Char In chars
                w = pFont.GetWidth(c) / 1000
                totalWidth += (w * curFontSize + Me.UndercontentCharacterSpacing) * Me.UndercontentHorizontalScaling / 100
            Next

            Return totalWidth
        End Function

        Private Sub DumpState()
            For Each location As TextChunk In locationalResult
                location.PrintDiagnostics()
                Console.WriteLine()
            Next
        End Sub

        Public Overridable Sub RenderText(ByVal renderInfo As TextRenderInfo) Implements ITextExtractionStrategy.RenderText
            Dim segment As LineSegment = renderInfo.GetBaseline()
            Dim location As New TextChunk(renderInfo.GetText(), segment.GetStartPoint(), segment.GetEndPoint(), renderInfo.GetSingleSpaceWidth())

            With location

                'Chunk Location:
                Debug.Print(renderInfo.GetText)
                .PosLeft = renderInfo.GetDescentLine.GetStartPoint(Vector.I1)
                .PosRight = renderInfo.GetAscentLine.GetEndPoint(Vector.I1)
                .PosBottom = renderInfo.GetDescentLine.GetStartPoint(Vector.I2)
                .PosTop = renderInfo.GetAscentLine.GetEndPoint(Vector.I2)
                'Chunk Font Size: (Height)
                .curFontSize = .PosTop - segment.GetStartPoint()(Vector.I2)
                'Use Font name  and Size as Key in the SortedList
                Dim StrKey As String = renderInfo.GetFont.PostscriptFontName & .curFontSize.ToString
                'Add this font to ThisPdfDocFonts SortedList if it's not already present
                If Not ThisPdfDocFonts.ContainsKey(StrKey) Then ThisPdfDocFonts.Add(StrKey, renderInfo.GetFont)
                'Store the SortedList index in this Chunk, so we can get it later
                .FontIndex = ThisPdfDocFonts.IndexOfKey(StrKey)
            End With
            locationalResult.Add(location)
        End Sub

        '*
        '         * Represents a chunk of text, it's orientation, and location relative to the orientation vector
        '         

        Public Class TextChunk
            Implements IComparable(Of TextChunk)
            '* the text of the chunk 

            Friend text As [String]
            '* the starting location of the chunk 

            Friend startLocation As Vector
            '* the ending location of the chunk 

            Friend endLocation As Vector
            '* unit vector in the orientation of the chunk 

            Friend orientationVector As Vector
            '* the orientation as a scalar for quick sorting 

            Friend orientationMagnitude As Integer
            '* perpendicular distance to the orientation unit vector (i.e. the Y position in an unrotated coordinate system)
            '             * we round to the nearest integer to handle the fuzziness of comparing floats 

            Friend distPerpendicular As Integer
            '* distance of the start of the chunk parallel to the orientation unit vector (i.e. the X position in an unrotated coordinate system) 

            Friend distParallelStart As Single
            '* distance of the end of the chunk parallel to the orientation unit vector (i.e. the X position in an unrotated coordinate system) 

            Friend distParallelEnd As Single
            '* the width of a single space character in the font of the chunk 

            Friend charSpaceWidth As Single

            Private _PosLeft As Single

            Private _PosRight As Single

            Private _PosTop As Single

            Private _PosBottom As Single

            Private _curFontSize As Single

            Private _FontIndex As Integer

            Public Property FontIndex As Integer
                Get
                    Return _FontIndex
                End Get
                Set(ByVal value As Integer)
                    _FontIndex = value
                End Set
            End Property

            Public Property PosLeft As Single
                Get
                    Return _PosLeft
                End Get
                Set(ByVal value As Single)
                    _PosLeft = value
                End Set
            End Property

            Public Property PosRight As Single
                Get
                    Return _PosRight
                End Get
                Set(ByVal value As Single)
                    _PosRight = value
                End Set
            End Property

            Public Property PosTop As Single
                Get
                    Return _PosTop
                End Get
                Set(ByVal value As Single)
                    _PosTop = value
                End Set
            End Property

            Public Property PosBottom As Single
                Get
                    Return _PosBottom
                End Get
                Set(ByVal value As Single)
                    _PosBottom = value
                End Set
            End Property

            Public Property curFontSize As Single
                Get
                    Return _curFontSize
                End Get
                Set(ByVal value As Single)
                    _curFontSize = value
                End Set
            End Property

            Public Sub New(ByVal str As [String], ByVal startLocation As Vector, ByVal endLocation As Vector, ByVal charSpaceWidth As Single)
                Me.text = str
                Me.startLocation = startLocation
                Me.endLocation = endLocation
                Me.charSpaceWidth = charSpaceWidth

                Dim oVector As Vector = endLocation.Subtract(startLocation)
                If oVector.Length = 0 Then
                    oVector = New Vector(1, 0, 0)
                End If
                orientationVector = oVector.Normalize()
                orientationMagnitude = CInt(Math.Truncate(Math.Atan2(orientationVector(Vector.I2), orientationVector(Vector.I1)) * 1000))

                Dim origin As New Vector(0, 0, 1)
                distPerpendicular = CInt((startLocation.Subtract(origin)).Cross(orientationVector)(Vector.I3))

                distParallelStart = orientationVector.Dot(startLocation)
                distParallelEnd = orientationVector.Dot(endLocation)
            End Sub

            Public Sub PrintDiagnostics()
                Console.WriteLine("Text (@" & Convert.ToString(startLocation) & " -> " & Convert.ToString(endLocation) & "): " & text)
                Console.WriteLine("orientationMagnitude: " & orientationMagnitude)
                Console.WriteLine("distPerpendicular: " & distPerpendicular)
                Console.WriteLine("distParallel: " & distParallelStart)
            End Sub

            '*
            '             * @param as the location to compare to
            '             * @return true is this location is on the the same line as the other
            '             

            Public Function SameLine(ByVal a As TextChunk) As Boolean
                If orientationMagnitude <> a.orientationMagnitude Then
                    Return False
                End If
                If distPerpendicular <> a.distPerpendicular Then
                    Return False
                End If
                Return True
            End Function

            '*
            '             * Computes the distance between the end of 'other' and the beginning of this chunk
            '             * in the direction of this chunk's orientation vector.  Note that it's a bad idea
            '             * to call this for chunks that aren't on the same line and orientation, but we don't
            '             * explicitly check for that condition for performance reasons.
            '             * @param other
            '             * @return the number of spaces between the end of 'other' and the beginning of this chunk
            '             

            Public Function DistanceFromEndOf(ByVal other As TextChunk) As Single
                Dim distance As Single = distParallelStart - other.distParallelEnd
                Return distance
            End Function

            '*
            '             * Compares based on orientation, perpendicular distance, then parallel distance
            '             * @see java.lang.Comparable#compareTo(java.lang.Object)
            '             

            Public Function CompareTo(ByVal rhs As TextChunk) As Integer Implements System.IComparable(Of TextChunk).CompareTo
                If Me Is rhs Then
                    Return 0
                End If
                ' not really needed, but just in case
                Dim rslt As Integer
                rslt = CompareInts(orientationMagnitude, rhs.orientationMagnitude)
                If rslt <> 0 Then
                    Return rslt
                End If

                rslt = CompareInts(distPerpendicular, rhs.distPerpendicular)
                If rslt <> 0 Then
                    Return rslt
                End If

                ' note: it's never safe to check floating point numbers for equality, and if two chunks
                ' are truly right on top of each other, which one comes first or second just doesn't matter
                ' so we arbitrarily choose this way.
                rslt = If(distParallelStart < rhs.distParallelStart, -1, 1)

                Return rslt
            End Function

            '*
            '             *
            '             * @param int1
            '             * @param int2
            '             * @return comparison of the two integers
            '             

            Private Shared Function CompareInts(ByVal int1 As Integer, ByVal int2 As Integer) As Integer
                Return If(int1 = int2, 0, If(int1 < int2, -1, 1))
            End Function


        End Class

        '*
        '         * no-op method - this renderer isn't interested in image events
        '         * @see com.itextpdf.text.pdf.parser.RenderListener#renderImage(com.itextpdf.text.pdf.parser.ImageRenderInfo)
        '         * @since 5.0.1
        '         

        Public Sub RenderImage(ByVal renderInfo As ImageRenderInfo) Implements IRenderListener.RenderImage
            ' do nothing
        End Sub
    End Class
End Namespace