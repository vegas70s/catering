<% 

' Script: kvh-WrapText.asp'
' Source: KVH Information Systems'
' Changes:'
'   2014-10-08 jcathcart - created'

' Procedure:'
'   WrapText( string, integer )'

'   @param string strTxtIn - String of text without line breaks.'
'   @param integer intLineLength - Desired maximum line length.'
'   @return string '

' Description:'
'   Inserts CRLF (line breaks) into text by replacing the'
'   last space before intLineLength with a CRLF.'

' Example: '
' <!--#include file="kvh-WrapText.asp"-->'
' < %'
' margin = 20'
' generatedText = "Double, double toil and trouble, Fire burn and cauldron bubble."'
' correctedText = WrapText( generatedText, margin )'
' response.write correctedText'
' % >'

Function WrapText( strTxtIn, intLineLength )

	Dim arrTxt()
	intTxtLength = Len(strTxtIn) - 1
	Redim arrTxt(intTxtLength)

	For i = 0 to intTxtLength

		arrTxt(i) = Mid(strTxtIn, i + 1, 1)
		If Asc(arrTxt(i)) = 32 Then
			intLastSpace = i
		End If

		intLimitCounter = intLimitCounter + 1
		If intLimitCounter = intLineLength Then
			arrTxt(intLastSpace) = vbCrLF
			intLimitCounter = 0
		End If

	Next

	WrapText = Join(arrTxt,"")
End Function
%>
