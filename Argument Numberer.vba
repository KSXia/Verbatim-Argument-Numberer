' ---Argument Numberer v1.0.2---
' Updated on 2024-09-03.
' https://github.com/KSXia/Verbatim-Argument-Numberer/
' Based on Verbatim 6.0.0's "AutoNumberTags" function.
Sub NumberArguments()
	Dim NumberPlaceholder as String
	Dim TemplateToNumber as String
	Dim ResetArgumentNumberAtPocket As Boolean
	Dim ResetArgumentNumberAtHat As Boolean
	Dim ResetArgumentNumberAtBlock As Boolean
	
	' ---USER CUSTOMIZATION---
	' Set the NumberPlaceholder to the character that you want the number to replace.
	NumberPlaceholder = "x"
	
	' Set the TemplateToNumber your numbering template, with the character you set as the NumberPlaceholder in place of where the number should go.
	' In your document, you must put the TemplateToNumber at the beginning of any tag or analytic you want this macro to number.
	' WARNING: The NumberPlaceholder MUST not be repeated in the template. The character used as the NumberPlaceholder MUST only show up once in the TemplateToNumber.
	TemplateToNumber = "[x]"
	
	' Set the header types that the argument number should reset at.
	' If you want the argument number to reset at a certain header type, set the corresponding variable to True.
	' If you do NOT want the argument number to reset at a certain header type, set the corresponding variable to False.
	ResetArgumentNumberAtPocket = True
	ResetArgumentNumberAtHat = True
	ResetArgumentNumberAtBlock = False
	
	' ---INITIAL VARIABLE SETUP---
	Dim TemplateLength As Integer
	TemplateLength = Len(TemplateToNumber)
	
	' The following code for numbering arguments is based on Verbatim 6.0.0's "Auto Number Tags" function.
	Dim p As Paragraph
	Dim CurrentArgumentNumber As Long
	
	' ---PROCESS TO NUMBER ARGUMENTS---
	' Loop through each paragraph and insert the number if the numbering template is present at the start of the paragraph.
	' Reset the numbering on any specified larger heading.
	For Each p In ActiveDocument.Paragraphs
		Select Case p.OutlineLevel
			Case Is = 1
				If ResetArgumentNumberAtPocket = True Then
					CurrentArgumentNumber = 0
				End If
			Case Is = 2
				If ResetArgumentNumberAtHat = True Then
					CurrentArgumentNumber = 0
				End If
			Case Is = 3
				If ResetArgumentNumberAtBlock = True Then
					CurrentArgumentNumber = 0
				End If
			Case Is = 4
				If Len(p.Range.Text) >= TemplateLength Then
					Dim IsTheNumberingTemplatePresent As Boolean
					IsTheNumberingTemplatePresent = True
					Dim i As Integer
					For i = 1 to TemplateLength Step 1
						' Going character-by-character, compare the characters at the start of the paragraph with the characters in the TemplateToNumber to see if they are the same.
						If p.Range.Characters(i) <> Mid(TemplateToNumber, i, 1) Then
							IsTheNumberingTemplatePresent = False
						End If
					Next i
					If IsTheNumberingTemplatePresent = True Then
						CurrentArgumentNumber = CurrentArgumentNumber + 1
						Dim j As Integer
						For j = 1 to TemplateLength Step 1
							If p.Range.Characters(j) = NumberPlaceholder Then
								p.Range.Characters(j) = CurrentArgumentNumber
							End If
						Next j
					End If
				End If
		End Select
	Next p
	' End of code based on Verbatim 6.0.0's functions.
End Sub