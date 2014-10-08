<!--#include file="kvh-WrapText.asp"-->
<%
Dim I, EmailFrom, EmailTo, Subject, mail
Dim printer, name, email, rdate, deptchargecode, grandtotal
Dim rtime, attendees, location, groupname, setuptype, notes
Dim amt1, amt2, amt3, amt4, amt5, amt6, amt7, amt8, amt9, amt10
Dim amt11, amt12, amt13, amt14, amt15, amt16, amt17, amt18, amt19, amt20
Dim amt21, amt22, amt23
Dim desc1, desc2, desc3, desc4, desc5, desc6, desc7, desc8, desc9, desc10
Dim desc11, desc12, desc13, desc14, desc15, desc16, desc17, desc18, desc19, desc20
Dim desc21, desc22, desc23
Dim price1, price2, price3, price4, price5, price6, price7, price8, price9, price10
Dim price11, price12, price13, price14, price15, price16, price17, price18, price19, price20
Dim price21, price22, price23
Dim total1, total2, total3, total4, total5, total6, total7, total8, total9, total10
Dim total11, total12, total13, total14, total15, total16, total17, total18, total19, total20
Dim total21, total22, total23, pipe
Dim Body, PrnBody, br, tr, td, ctr, ctd, nbsp
Dim strUsername, strPassword, objFS, objWSHNet, objPrinter

'printer = "\\KVCHPRINT\FNS-PR1"
printer = "\\KVCHPRINT\MIS-PR1"
strUsername = "kvch\ttester"
strPassword = "ttester"

pipe=" | "
nbsp = "&nbsp;"
br = "<br>"
td = "<td>"
ctd = "&nbsp;</td>"
tr = "<tr>"
ctr ="</tr>"

'get data from form
name = request.form(rtrim("name"))
email = request.form(trim("email"))
emailFrom = "form.catering-request@kvhealthcare.org"
emailTo = "jcathcart@kvhealthcare.org"
'EmailTo = email & "; mhanson@kvhealthcare.org"
Subject = "Catering Request"
rtime = request.form(trim("time"))
rdate = request.form(trim("date"))
attendees = request.form(trim("attendees"))
groupname = request.form(trim("groupname"))
location = request.form(trim("location"))
setuptype = request.form(trim("setuptype"))
deptchargecode = request.form(trim("deptchargecode"))
'notes = request.form(trim("notes"))
' 2014-10-09 jcaathcart - added WrapText function to inject line breaks
notes = WrapText( request.form(trim("notes")), 60 )

price1 = .08
price2 = .05
price3 = .08
price4 = 2.65
price5 = 1.45
price6 = 5.90
price7 = 8.60
price8 = 2.10
price9 = 1.25
price10 = .55
price11 = 1.40
price12 = 1.40
price13 = 1.15
price14 = .85
price15 = 1.05
price16 = .85
price17 = 8.30
price18 = 63
price19 = 31.50
price20 = 46.20
price21 = 23.10
price22 = .50
price23 = .40

desc1 = "12 oz. dixie coffee cups, each"
desc2 = "Napkin, each"
desc3 = "Plasticware, each"
desc4 = "Tablecloth, disposable, each"
desc5 = "20 oz. Bottled Beverages (Tea, Juice, Soda, Water)"
desc6 = "Pitcher Fruit Punch, serves 10"
desc7 = "Pump Pot Coffee (Reg or Decaf + coffee condiments)"
desc8 = "Pump Pot Hot Water (includes tea bags, cocoa packets)"
desc9 = "Cheese slice, packaged, each"
desc10 = "Yogurt, 6 oz."
desc11 = "Bagels, each, sliced (includes cream cheese portion pkt)"
desc12 = "Danish"
desc13 = "Cinnamon Roll, each"
desc14 = "Cookies, brownies, bars, each"
desc15 = "Cookies, specialty, holiday (specify)"
desc16 = "Fresh muffin, each"
desc17 = "Box Lunch, each"
desc18 = "Meat/Cheese Tray,Large ( 35- 40 servings)"
desc19 = "Meat/Cheese Tray,Small (15- 20 servings)"
desc20 = "Veg or Fruit Tray, Large ( 35- 40 servings)"
desc21 = "Veg or Fruit Tray, small (15-25 servings)"
desc22 = "Ice Cream, 4 oz."
desc23 = "Sherbet, 4 oz."

For I = 1 to 23
	execute("amt" & cstr(I) & " = cint(request.form(" & chr(34) & "amt" & cstr(I) & chr(34) & "))")
	execute("total"& I &" = amt" & I & " * price" & I)
	execute("grandtotal = grandtotal + total" & I)
next


if (name = "") or (rdate = "") or (email = "") or (deptchargecode = "") or (rtime = "") or (attendees = "") or (groupname = "") or (location = "") or (setuptype = "") then
	Response.Write "<p>We are sorry but there seems to be an error in the form. Please click back on your browser and complete the following field(s) : </p>" 
	if (name<>"") Then
	Else
	Response.Write "<font color='red'>Name is blank.</font><br>"
	end if
	if (rtime<>"") Then
	Else
	Response.Write "<font color='red'>Time is blank.</font><br>"
	end if
	if (attendees<>"") Then
	Else
	Response.Write "<font color='red'>Attendees is blank.</font><br>"
	end if

	if (groupname<>"") Then
	Else
	Response.Write "<font color='red'>Group name is blank.</font><br>"
	end if

	if (location<>"") Then
	Else
	Response.Write "<font color='red'>Location is blank.</font><br>"
	end if
	
	if (setuptype<>"") Then
	Else
	Response.Write "<font color='red'>Setup type is blank.</font><br>"
	end if

	if (rdate<>"") Then
	ELSE
	Response.Write "<font color='red'>Date has been left blank.</font><br>"
	end if
	if (email<>"") Then
	ELSE
	Response.Write "<font color='red'>Email address is blank.</font><br>"
	end if
	if (deptchargecode<>"") Then
	ELSE
	Response.Write "<font color='red'>Department Charge code is blank.</font><br>"
	end if
end if

' prepare email body text

Body = Body & "Requestor name: " & name & br 
Body = Body & "Requestor email: " & email & br 
Body = Body & "Event Date: " & rdate & br 
Body = Body & "Event Time: " & rtime & br 
Body = Body & "Attendees: " & attendees & br
Body = Body & "Group name: " & groupname & br
Body = Body & "Location: " & location & br
Body = Body & "Setup type: " & setuptype & br & br
Body = Body & "Department Charge Code: " & deptchargecode & br & br
Body = Body & "Notes: " & replace(notes,chr(13),"<br>") & br & br
Body = Body & "<table border='0' cellpadding='0' cellspacing='0'>" & br
Body = Body & tr & td & "Item &nbsp;" & ctd & td & "Cost &nbsp;" & ctd & td & "Amount &nbsp;" & ctd & td & "Subtotal &nbsp;" & ctd & ctr
For I = 1 to 23
	if eval("amt" & cstr(i) &" > 0") then
		execute("Body = Body & tr & td & desc" & cstr(i) & " & nbsp & ctd")
		execute("Body = Body & td & price" & cstr(i) & " & ctd")
		execute("Body = Body & td & amt" & cstr(i) & " & ctd")
		execute("Body = Body & td & formatnumber(total" & cstr(i) & ",2) & ctd & ctr")
	end if 
Next
Body = Body & "</table>" & br
Body = Body & "Grand total = " & formatnumber(grandtotal,2) & br

'prepare printed body test

PrnBody = PrnBody & "Catering Request " & Vbcrlf & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Requestor name: " & name & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Requestor email: " & email & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Event Date: " & rdate & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Event Time: " & rtime & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Attendees: " & attendees & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Group Name: " & groupname & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Location: " & location & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Setup type: " & setuptype & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Department Charge Code: " & deptchargecode & Vbcrlf & Vbcrlf 
PrnBody = PrnBody & "Notes: " & notes & Vbcrlf & Vbcrlf 

For I = 1 to 23
	if eval("amt" &cstr(i)&" > 0") then
		if eval("amt" &cstr(i)&" < 10") then
			Execute("PrnBody = PrnBody & " & chr(34) & "Quantity = 0" & chr(34) & " & amt" & cstr(i) & " & pipe")
		else
			Execute("PrnBody = PrnBody & " & chr(34) & "Quantity = "  & chr(34) & " & amt" & cstr(i) & " & pipe")		
		end if
		Execute("PrnBody = PrnBody & desc" & cstr(i) & " & VbCrLf & VbCrLf")
		' PrnBody = PrnBody & "Subtotal = " & formatnumber(total1,2)  & VbCrLf & VbCrLf
	end if
next

PrnBody = PrnBody & VbCrLf & "Total = $" & formatnumber(grandtotal,2) & VbCrLf

' Create FileSystem Object and Windows Script Host Network Object
'On Error Resume Next

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objWSHNet = CreateObject("WScript.Network")

on error resume next
objWSHNet.RemovePrinterConnection "LPT1:"
objWSHNet.AddPrinterConnection "LPT1:", printer, False, strUsername, strPassword
Set objPrinter = objFS.CreateTextFile("LPT1:", True)
objPrinter.Write(PrnBody)
'objPrinter.Write(Body)
Response.Write "<br><font size='5'>Your catering order has been submitted and sent to the printer.</font><br>"


' send email 

if (name<>"") AND (rdate<>"") AND (email<>"") AND (deptchargecode<>"") THEN
	Set mail = Server.CreateObject("CDO.Message") 
	mail.To = EmailTo
	mail.From = EmailFrom
	mail.Subject = Subject
	mail.HTMLBody = Body
	mail.Send()
	' redirect to success page 
	
	' Close the print device object and trap for errors
	'On Error Resume Next
	objPrinter.Close
	objWSHNet.RemovePrinterConnection "LPT1:"
	Set objWSHNet  = Nothing
	Set objFS      = Nothing
	Set objPrinter = Nothing

	Response.Write "<br><font size='5'>An email containing your order has also been sent.</font><br>"
ELSE
	Response.Write "<br><font color='red' size='5'> Your catering order has not been submitted.</font><br>"
end if
Response.Write br + Body

%>
