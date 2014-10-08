<pre>
<!--#include file="kvh-WrapText.asp"-->
<%
margin = 30'
generatedText = "Double, double toil and trouble, Fire burn and cauldron bubble."'
correctedText = WrapText( generatedText, margin )'
response.write correctedText'

%>
