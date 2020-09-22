<div align="center">

## DevinsHandyASPstuff


</div>

### Description

This is a compilation of functions I use when working on ASP projects. There are functions to build HTML form elements (and whole forms), HTML tables, 'smart' date drop down boxes, capitalization functions, date functions, a sql quote handler, a bunch of stuff. I'm providing it hoping someone out there might get use from one of these at some point. I'd also like opinions, suggestions, and contributions too.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Devin Garlit](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/devin-garlit.md)
**Level**          |Beginner
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__4-35.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/devin-garlit-devinshandyaspstuff__4-6631/archive/master.zip)





### Source Code

```
<%
	''''''DevinsHandyASPstuff'''''''''''''''
	'
	'purpose: This is just a compilation of numerous ASP function I have built and use. Each one should be
	'     commented.
	'
	'programmer: Devin Garlit dgarlit@hotmail.com
	'
	'write(strString)
	'buildTextBox(strValue, strFieldName, intSize, intMaxSize, blnLabel, strLabel)
	'buildPasswordBox(strValue, strFieldName, intSize, intMaxSize, blnLabel, strLabel)
	'buildHidden(strValue, strFieldName, intSize, intMaxSize, blnDisplayValue, strDisplayValue)
	'buildCheckBox(strValue, strFieldName, blnChecked, blnDisplayValue, strDisplayValue)
	'buildRadioButton(strValue, strFieldName, blnDisplayValue, strDisplayValue)
	'buildTextArea(strValue, strFieldName, intCols, intRows, strWrap)
	'buildDropDownFromDB( objConnection, strSQL, strName)
	'buildDropDownFromDBwithTitle( objConnection, strSQL, strName, strTitle)
	'createAForm(RS, strFormName, strFormMethod, strFormAction)
	'requestAndIncludeAsHidden()
	'CheckQuotes (strValue)
	'a cut and paste cache-control script
	'write(strString) 'instead of response.write
	'RemoveHTMLTags (strString)
	'isOdd (strNum)
	'Caps(strString) - capitalize the first letter of a string
	'capAllWords (strString)
	'GetYear (strDate)
	'GetMonthNum (strDate)
	'GetDayNum (strDate)
	'GetDateWithDay (strDate) 'return day and date like this: Saturday, September 24, 1977
	'GetLongDate (strDate)
	'GetDateFromParts(strMonth, strDay, strYear) 'returns a date from the month, day and year, allows an empty string for day( but will pull the first of the month
	'writeTable(intCols, intRows, arrValues, strTableAttributes, strRowAttributes, strCellAttributes )
	'writeTable2(arrValues, strTableAttributes, strRowAttributes, strCellAttributes )
	'createAForm2WHidden(RS, intColumnSplit, strFormName, strFormMethod, strFormAction, strButton)
	'createAForm2(RS, intColumnSplit, strFormName, strFormMethod, strFormAction, strButton, strEditFlag)
	'getDaysInMonth(strMonth,strYear)
	'
	'writeDropDowns()
	' writeDropDowns is a way I used MonthDropDown, DayDropDown, and YearDropDown together
	' basically, the point was that I didn't want someone to select 30 for the month of february
	' so it resubmits to the page(that could be costly depending on what else is goin on) with the selected
	' day,month,year and it sets/resets the days according to the month and year so the user cannot select
	' day 30 for month 2
	'MonthDropDown(strName, blnNum, strSelected, strSelfLink)
	'YearDropDown(strName, intStartYear, intEndYear, strSelected, strSelfLink)
	'DayDropDown(strName, intStartDay, intEndDay, strSelected )
	'beginDoc (strTitle)
	'endDoc()
  '''instead of writing out response.write all the time
  sub write(strString)
		Response.Write strString
  end sub
	'**************************************************************
	'Function: buildTextBox(strValue, strFieldName, intSize, intMaxSize, blnLabel, strLabel)
	'
	'Returns: an string of an HTML input field
	'
	'Inputs:
	'			strValue = a string of the value for the input field
	'   strFieldName = a string of the name of the input field
	'   intSize = an integer of the size of the input field
	'   intMaxsize = an integer of the maxlength of the input field
	'   blnLabel = a true/false to determine if a label will be placed in front of the input field
	'   strLabel = the label to be used if blnLabel is true
	'
	'Notes:
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
	function buildTextBox(strValue, strFieldName, intSize, intMaxSize, blnLabel, strLabel)
		if cbool(blnLabel) then
			buildTextBox = strLabel & " " & "<input type='text' name='" & strFieldName & "' value='" & strValue & "' size='" & intSize & "' maxlength='"& intMaxSize & "'>"
		else
			buildTextBox = "<input type='text' name='" & strFieldName & "' value='" & strValue & "' size='" & intSize & "' maxlength='"& intMaxSize & "'>"
		end if
	end function
	function buildPasswordBox(strValue, strFieldName, intSize, intMaxSize, blnLabel, strLabel)
		if cbool(blnLabel) then
			buildPasswordBox = strLabel & " " & "<input type='Password' name='" & strFieldName & "' value='" & strValue & "' size='" & intSize & "' maxlength='"& intMaxSize & "'>"
		else
			buildPasswordBox = "<input type='Password' name='" & strFieldName & "' value='" & strValue & "' size='" & intSize & "' maxlength='"& intMaxSize & "'>"
		end if
	end function
	'**************************************************************
	'Function: buildHidden(strValue, strFieldName, intSize, intMaxSize, blnDisplayValue, strDisplayValue)
	'
	'Returns: an string of an HTML hidden field
	'
	'Inputs:
	'			strValue = a string of the value for the input field
	'   strFieldName = a string of the name of the input field
	'   blnDisplayValue = a true/false to determine if a value will be displayed
	'   strDisplayValue = the value to be displayed if blnDisplayValue is true
	'
	'Notes:
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
	function buildHidden(strValue, strFieldName, blnDisplayValue, strDisplayValue)
		if cbool(blnDisplayValue) then
			buildHidden = strDisplayValue & " " & "<input type='hidden' name='" & strFieldName & "' value='" & strValue & "'>"
		else
			buildHidden = "<input type='hidden' name='" & strFieldName & "' value='" & strValue & "'>"
		end if
	end function
	'**************************************************************
	'Function: buildCheckBox(strValue, strFieldName, blnChecked, blnDisplayValue, strDisplayValue)
	'
	'Returns: an string of an HTML checkbox
	'
	'Inputs:
	'			strValue = a string of the value for the checkbox
	'   strFieldName = a string of the name of the checkbox
	'   blnChecked = a true/false whether the box is checked(true) or uncheck(false)
	'   blnDisplayValue = a true/false to determine if a value will be displayed
	'   strDisplayValue = the value to be displayed if blnDisplayValue is true
	'
	'Notes: if true the display value is displayed after the checkbox
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
	function buildCheckBox(strValue, strFieldName, blnChecked, blnDisplayValue, strDisplayValue)
		dim strChecked
		if cbool(blnChecked) then
			strChecked = "CHECKED"
		else
			strChecked = ""
		end if
		if cbool(blnDisplayValue) then
				buildCheckBox = "<input type='checkbox' name='" & strFieldName & "' value='" & strValue &"' " & strChecked & ">" & " " & strDisplayValue
		else
				buildCheckBox = "<input type='checkbox' name='" & strFieldName & "' value='" & strValue &"'" & strChecked & ">"
		end if
	end function
	'**************************************************************
	'Function: buildRadioButton(strValue, strFieldName, blnDisplayValue, strDisplayValue)
	'
	'Returns: an string of an HTML radio button
	'
	'Inputs:
	'			strValue = a string of the value for the radio button
	'   strFieldName = a string of the name of the radio button
	'   blnDisplayValue = a true/false to determine if a value will be displayed
	'   strDisplayValue = the value to be displayed if blnDisplayValue is true
	'
	'Notes: if true the display value is displayed after the radio button
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
	function buildRadioButton(strValue, strFieldName, blnDisplayValue, strDisplayValue)
		if cbool(blnDisplayValue) then
				buildRadioButton = "<input type='radio' name='" & strFieldName & "' value='" & strValue &"'>" & " " & strDisplayValue
		else
				buildRadioButton = "<input type='radio' name='" & strFieldName & "' value='" & strValue &"'>"
		end if
	end function
	'**************************************************************
	'Function: buildTextArea(strValue, strFieldName, intCols, intRows, strWrap)
	'
	'Returns: an string of an HTML textarea
	'
	'Inputs:
	'			strValue = a string of the value for the textarea
	'   strFieldName = a string of the name of the textarea
	'   intCols = an integer for the cols attribute
	'   intRows = an integer for the rows attribute
	'   strWrap = a string for the wrap attribute i.e. "virtual"
	'
	'Notes:
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
	function buildTextArea(strValue, strFieldName, intCols, intRows, strWrap)
		dim strTemp
		strTemp = "<textarea cols=" & intCols & " rows=" & intRows & " name='" & strFieldName & "' wrap=" & strWrap & ">"
		strTemp = strTemp & buildTextArea & vbcrlf & strValue & vbcrlf & "</textarea>"
		buildTextArea = strTemp
	end function
	'**************************************************************
	'Function: buildDropDownFromDB( objConnection, strSQL, strName)
	'
	'Returns: an string of an HTML checkbox
	'
	'Inputs:
	'			objConnection = a connection object
	'   strSQL = a string of a SQL statement
	'   strName = a string of the name attribute of the select box
	'
	'Notes:
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
	function buildDropDownFromDB( objConnection, strSQL, strName)
		dim RS 'recordset
		dim strTemp
		set RS = objConnection.execute(strSQL)
		strTemp = "<select name='" & strName& "'>" & vbcrlf
		do while not RS.EOF
			strTemp = strTemp & "<option value='" & RS.fields(0) & "'>" & RS.fields(0) & "</option>" & vbcrlf
			RS.MoveNext
		Loop
		set RS = nothing
		strTemp = strTemp & "</select>"
		buildDropDownFromDB = strTemp
	end function
	'**************************************************************
	'Function: buildDropDownFromDBwithTitle( objConnection, strSQL, strName, strTitle)
	'
	'Returns: an string of an HTML checkbox
	'
	'Inputs:
	'			objConnection = a connection object
	'   strSQL = a string of a SQL statement
	'   strName = a string of the name attribute of the select box
	'   strTitle = a string for the value of the first option of the select box i.e. "Select"
	'
	'Notes:
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
	function buildDropDownFromDBwithTitle( objConnection, strSQL, strName, strTitle)
		dim RS 'recordset
		dim strTemp
		set RS = objConnection.execute(strSQL)
		strTemp = "<select name='" & strName& "'>" & vbcrlf
		strTemp = strTemp & "<option value='" & strTitle & "'>" & strTitle & "</option>" & vbcrlf
		do while not RS.EOF
			strTemp = strTemp & "<option value='" & RS.fields(0) & "'>" & RS.fields(0) & "</option>" & vbcrlf
			RS.MoveNext
		Loop
		set RS = nothing
		strTemp = strTemp & "</select>"
		buildDropDownFromDBwithTitle = strTemp
	end function
	'**************************************************************
	'Function: createAForm(RS, strFormName, strFormMethod, strFormAction)
	'
	'Returns: creates a simple html form of text boxes using buildTextBox from a recordset
	'
	'Inputs:
	'			RS = a recordset object
	'   strFormName = a string of the name of the form
	'   strFormMethod = a string of the forms method i.e. "post"
	'   strFormAction = a string of the forms action
	'
	'Notes: real simple, just lines them up in a simple table and gives a simple submit button
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
function createAForm(RS, strFormName, strFormMethod, strFormAction)
	dim x
	Response.Write "<Form method='" & strFormMethod & "' name='" & strFormName & "' action='" & strFormAction & "'>" & vbcrlf
	Response.Write "<table>" & vbcrlf
	for x = 0 to RS.Fields.Count-1
		Response.Write "<tr><td>"
		Response.write RS.Fields(x).Name & "</td><td>"
		Response.Write buildTextBox("", RS.Fields(x).Name, 25, RS.Fields(x).DefinedSize, false, "") & "<br>"
		Response.Write "</td></tr>" & vbcrlf
	next
	Response.Write "<tr><td>&nbsp;</td><td><input type=submit name=Submit value=Submit></td></tr>" & vbcrlf
	Response.Write "</table>" & vbcrlf
	Response.Write "</Form>"
end function
function requestAndIncludeAsHidden()
	dim field
	for each Field in Request.Form
		buildHidden request(field), field.name, true, request(field)
	next
end function
'a classic to take care of those pesky quotes when working with SQL
function CheckQuotes(strValue)
		if not isnull(strValue) and strValue <> "" then
			CheckQuotes = replace(strValue,"'","''")
		else
			CheckQuotes = strValue
		end if
end function
	''''cachecontrol
	'''included right after option explicit
  'Response.Buffer=TRUE
  'Response.Expires = 0
  'Response.AddHeader "Pragma","no-cache"
  'Response.AddHeader "cache-control","no-store"
 'capitilize first letter
	function Caps(strString)
		Caps = ucase(left(strString,1)) & lcase(mid(strString,2))
	end function
	'capitializ all words in a string
	'write capAllWords("we actually do listen to our users once in a while")
	function capAllWords(strString)
		dim arrTemp, strTemp, i
		arrTemp = split(strString, " ")
		for i = 0 to Ubound(arrTemp)
			strTemp = strTemp & " " & ucase(left(arrTemp(i),1)) & lcase(mid(arrTemp(i),2))
		next
		capAllWords = strTemp
	end function
	'write GetYear("09/24/1977")
	'return a simple year # from a string in format of yyyy
	function GetYear(strDate)
		GetYear = datepart("yyyy",strDate)
	end function
	'return a month #
	function GetMonthNum(strDate)
		GetMonthNum = datepart("m",strDate)
	end function
	'return a day #
	function GetDayNum(strDate)
		GetDayNum = datepart("d",strDate)
	end function
	'return day and date like this: Saturday, September 24, 1977
	function GetDateWithDay(strDate)
		GetDateWithDay = formatdatetime(strDate,1)
	end function
	'return long date like 9/24/1977
	function GetLongDate(strDate)
		GetLongDate = formatdatetime(strDate,2)
	end function
	'returns a date from the month, day and year, allows an empty string for day( but will pull the first of the month
	'write GetDateFromParts("9", "", "77")
	'write GetDateFromParts("9", "24", "77")
	function GetDateFromParts(strMonth, strDay, strYear)
		if strDay <> "" then
			GetDateFromParts = formatdatetime(strMonth & "/" & strDay & "/" & strYear)
		else
			GetDateFromParts = formatdatetime(strMonth & "/" & strYear)
		end if
	end function
	'''''''''''
	''''vbs function FormatDateTime formats'''
	'd Short Date
	'D Long Date
	'f Full (long date + short time)
	'F Full (long date + long time)
	'g General (short date + short time)
	'G General (short date + long time)
	'm, M Month/Day Date
	'r, R RFC Standard
	's Sortable without TimeZone info
	't Short Time
	'T Long Time
	'u Universal with sort able format
	'U Universal with Full (long date + long time) format
	'y, Y Year/Month Date
'returns a true if the number (an int or string) is odd, a false otherwise
function isOdd(strNum)
	if cint(strNum) mod 2 = 0 then
		isOdd = false
	else
		isOdd = true
	end if
end function
'remove HTML tags from a string, note, this won't handle html encoding.
'write RemoveHTMLTags("<B>BOB</B> rules")
Function RemoveHTMLTags(strString)
	Dim nCharPos, sOut, bInTag, sChar
	sOut = ""
	bInTag = False
	For nCharPos = 1 To Len(strString)
		sChar = Mid(strString, nCharPos, 1)
		If sChar = "<" Then
			bInTag = True
		End If
		If Not bInTag Then sOut = sOut & sChar
		If sChar = ">" Then
			bInTag = False
		End If
	Next
	RemoveHTMLTags = sOut
End Function
'''''''''''''''''''''''''''''''''''sortable table
'dim objConn
	'Set objConn = server.CreateObject("ADODB.Connection")
	'objConn.Open "passwordlist"
	'strSQL = "Select * From passwords"
	'createSortableList objConn,strSQL, "id", request("sort"),request("page"),"sort.asp",5, "border=1 bgcolor=steelblue"
	'creates a sortable html table
	sub createSortableList(objConn,strSQL, strDefaultSort, strSort, intCurrentPage, strPageName, intPageSize, strLinkedColumnName,strLink,strTableAttributes)
		dim RS 'recordset
		dim strTemp, field, strMoveFirst, strMoveNext, strMovePrevious, strMoveLast
		dim i, intTotalPages, intCurrentRecord, intTotalRecords
		i = 0
		if strSort = "" then
			strSort = strDefaultSort
		end if
		if intCurrentPage = "" then
			intCurrentPage = 1
		end if
		set RS = server.CreateObject("adodb.recordset")
		with RS
			.CursorLocation=3
			.Open strSQL & " order by " & replace(strSort,"desc"," desc"), objConn,adOpenStatic
			if not rs.EOF then
				.PageSize = cint(intPageSize)
				intTotalPages = .PageCount
				intCurrentRecord = .AbsolutePosition
				.AbsolutePage = intCurrentPage
				intTotalRecords = .RecordCount
			end if
		end with
		Response.Write "<table " & strTableAttributes & " >" & vbcrlf
		Response.Write "<tr>" & vbcrlf
		'if not rs.EOF then
			for each field in RS.Fields
				Response.Write "<td>" & vbcrlf
				if instr(strSort, "desc") then
					Response.Write "<a href=" & strPageName & "?sort="& field.name &"&page="&intCurrentPage&">" & field.name & "</a>" & vbcrlf
				else
					Response.Write "<a href=" & strPageName & "?sort="& field.name &"desc&page="&intCurrentPage&">" & field.name & "</a>"	& vbcrlf
				end if
				Response.Write "<td>"	& vbcrlf
			next
		'end if
		Response.Write "<tr>"
		for i = intCurrentRecord to RS.PageSize
			if not RS.eof then
			Response.Write "<tr>" & vbcrlf
			for each field in RS.Fields
				Response.Write "<td>" & vbcrlf
				if lcase(strLinkedColumnName) = lcase(field.name) then
					Response.Write "<a href=" & strLink & "?sort="& strSort &"&page="&intCurrentPage&">" & field.value & "</a>" & vbcrlf
				else
					Response.Write field.value
				end if
				Response.Write "<td>" & vbcrlf
			next
			Response.Write "<tr>" & vbcrlf
			RS.MoveNext
			end if
		next
		Response.Write "<table>" & vbcrlf
	'Response.Write intTotalPages		& "  " & intCurrentPage
		select case cint(intCurrentPage)
			case cint(intTotalPages) 'last page
				strMoveFirst = "<a href=" & strPageName & "?sort="& strSort &"&page=1 >"& "first" &"</a>"
				strMoveNext = ""
				strMovePrevious = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage - 1 & " >"& "Prev" &"</a>"
				strMoveLast = "" '"<a href=" & strPageName & "?sort="& strSort &"&page=" & intTotalPages & " >"
			case 1 'first page
				strMoveFirst = "" '"<a href=" & strPageName & "?sort="& strSort &"&page=1 >"
				strMoveNext = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage + 1 & " >"& "next" &"</a>"
				strMovePrevious = "" '"<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage - 1 & " >"
				strMoveLast = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intTotalPages & " >"& "last" &"</a>"
			case else
				strMoveFirst = "<a href=" & strPageName & "?sort="& strSort &"&page=1 >"& "first" &"</a>"
				strMoveNext = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage + 1 & " >"& "next" &"</a>"
				strMovePrevious = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage - 1 & " >"& "Prev" &"</a>"
				strMoveLast = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intTotalPages & " >"& "last" &"</a>"
		end select
		with Response
			.Write strMoveFirst & " "
			.Write strMovePrevious
			.Write " " & intCurrentPage & " of " & intTotalPages & " "
			.Write strMoveNext & " "
			.Write strMoveLast
		end with
		if RS.State = &H00000001 then 'its open
			RS.Close
		end if
		set RS = nothing
	end sub
	'**************************************************************
	'Function: writeTable(intCols, intRows, strTableAttributes, strRowAttributes, arrValues)
	'
	'Returns: writes a html table
	'
	'Inputs:
	'			intCols = # of column
	'   intRows = # of rows
	'   strTableAttributes = string of table attributes seperated by a space i.e. "border=1 bgcolor=steelblue"
	'   strRowAttriutes = string of row attributes seperated by a space i.e. "valign=top"
	'   arrValues = a multidimensional array in format of arr(rows,cols)
	'
	'Notes:
	'
	'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
	'**************************************************************
	function writeTable(intCols, intRows, arrValues, strTableAttributes, strRowAttributes, strCellAttributes )
		dim i, j
		write "<table " & strTableAttributes & " >" & vbcrlf
		for i = 0 to intRows - 1
			write "<tr " & strRowAttributes & " >" & vbcrlf
			for j = 0 to intCols - 1
				write "<td " & strCellAttributes & " >" & vbcrlf
				write arrValues(i,j)
				write "</td>" & vbcrlf
			next
			write "</tr>" & vbcrlf
		next
		write "</table>" & vbcrlf
	end function
function writeTable2(arrValues, strTableAttributes, strRowAttributes, strCellAttributes )
		dim i, j
		'write ubound(arrValues,1)
		'write ubound(arrValues,1)
		'Response.end
		write "<table " & strTableAttributes & " >" & vbcrlf
		for i = 0 to ubound(arrValues)-1
			write "<tr " & strRowAttributes & " >" & vbcrlf
			for j = 0 to ubound(arrValues,1)-1
				write "<td " & strCellAttributes & " >" & vbcrlf
				write arrValues(i,j)
				write "</td>" & vbcrlf
			next
			write "</tr>" & vbcrlf
		next
		write "</table>" & vbcrlf
	end function
'**************************************************************
'Function: createAForm2WHidden(RS, strFormName, strFormMethod, strFormAction, strButton)
'
'Returns: creates a simple html form of hidden fields from a recordset
'
'Inputs:
'			RS = a recordset object
'   intColumnSplit = the number at which to stop the first column, the rest of the fields will go to the next
'   strFormName = a string of the name of the form
'   strFormMethod = a string of the forms method i.e. "post"
'   strFormAction = a string of the forms action
'			strButton = a string of html for the submit and other action type buttons
'
'Notes: real simple, just lines them up in a simple table and gives a simple submit button
'
'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
'**************************************************************
function createAForm2WHidden(RS, intColumnSplit, strFormName, strFormMethod, strFormAction, strButton)
	dim x
	write "<Form method='" & strFormMethod & "' name='" & strFormName & "' action='" & strFormAction & "'>" & vbcrlf
	write "<table>" & vbcrlf
		write "<tr>" & vbcrlf
		 write "<td valign=top >" & vbcrlf
				write "<table border=1>" & vbcrlf
				for x = 0 to intColumnSplit
					write "<tr><td>" & vbcrlf
					write RS.Fields(x).Name & "</td><td>"
					write buildHidden(request(cstr(RS.Fields(x).Name)), RS.Fields(x).Name,true, request(cstr(RS.Fields(x).Name)) )
					write "</td></tr>" & vbcrlf
				next
				write "</table>" & vbcrlf
			write "</td>"
			write "<td valign=top >"
				write "<table border=1>" & vbcrlf
				for x = intColumnSplit + 1 to RS.Fields.Count-1
					write "<tr><td>" & vbcrlf
					write RS.Fields(x).Name & "</td><td>"
					write buildHidden(request(cstr(RS.Fields(x).Name)), RS.Fields(x).Name,true, request(cstr(RS.Fields(x).Name)) )
					write "</td></tr>" & vbcrlf
				next
				write "</table>" & vbcrlf
			write "</td>"	& vbcrlf
		write "</tr>"	& vbcrlf
	write "</table>" & vbcrlf
	write strButton & vbcrlf
	write "</Form>"
end function
'**************************************************************
'Function: createAForm2(RS, intColumnSplit, strFormName, strFormMethod, strFormAction, strButton, strEditFlag)
'
'Returns: creates a simple html form of hidden fields from a recordset
'
'Inputs:
'			RS = a recordset object
'   intColumnSplit = the number at which to stop the first column, the rest of the fields will go to the next
'   strFormName = a string of the name of the form
'   strFormMethod = a string of the forms method i.e. "post"
'   strFormAction = a string of the forms action
'			strButton = a string of html for the submit and other action type buttons
'   strEditFlag = a string of whether to fill the txtboxes with requested false, true or false
'
'Notes: real simple, just lines them up in a simple table and gives a simple submit button
'
'Programmer: Devin Garlit dgarlit@hotmail.com. 4/01/01
'**************************************************************
function createAForm2(RS, intColumnSplit, strFormName, strFormMethod, strFormAction, strButton, strEditFlag)
	dim x
	write "<Form method='" & strFormMethod & "' name='" & strFormName & "' action='" & strFormAction & "'>" & vbcrlf
	write "<table>" & vbcrlf
		write "<tr>" & vbcrlf
		 write "<td valign=top >" & vbcrlf
				write "<table border=1>" & vbcrlf
				for x = 0 to intColumnSplit
					write "<tr><td>" & vbcrlf
					write RS.Fields(x).Name & "</td><td>"
					if cbool(strEditFlag) then
						write buildTextBox(request(cstr(RS.Fields(x).Name)), RS.Fields(x).Name, 25, RS.Fields(x).DefinedSize, false, "") & "<br>"
					else
						write buildTextBox("", RS.Fields(x).Name, 25, RS.Fields(x).DefinedSize, false, "") & "<br>"
					end if
					write "</td></tr>" & vbcrlf
				next
				write "</table>" & vbcrlf
			write "</td>"
			write "<td valign=top >"
				write "<table border=1>" & vbcrlf
				for x = intColumnSplit + 1 to RS.Fields.Count-1
					write "<tr><td>" & vbcrlf
					write RS.Fields(x).Name & "</td><td>"
					if cbool(strEditFlag) then
						write buildTextBox(request(cstr(RS.Fields(x).Name)), RS.Fields(x).Name, 25, RS.Fields(x).DefinedSize, false, "") & "<br>"
					else
						write buildTextBox("", RS.Fields(x).Name, 25, RS.Fields(x).DefinedSize, false, "") & "<br>"
					end if
					write "</td></tr>" & vbcrlf
				next
				write "</table>" & vbcrlf
			write "</td>"	& vbcrlf
		write "</tr>"	& vbcrlf
	write "</table>" & vbcrlf
	write strButton & vbcrlf
	write "</Form>"
end function
function getDaysInMonth(strMonth,strYear)
		dim strDays
  Select Case cint(strMonth)
    Case 1,3,5,7,8,10,12:
						strDays = 31
    Case 4,6,9,11:
						strDays = 30
    Case 2:
						if ( (cint(strYear) mod 4 = 0 and cint(strYear) mod 100 <> 0) or ( cint(strYear) mod 400 = 0) ) then
							strDays = 29
						else
							strDays = 28
						end if
    'Case Else:
  End Select
  getDaysInMonth = strDays
end function
'''writeDropDowns is a way I used MonthDropDown, DayDropDown, and YearDropDown together
'basically, the point was that I didn't want someone to select 30 for the month of february
'so it resubmits to the page(that could be costly depending on what else is goin on) with the selected
'day,month,year and it sets/resets the days according to the month and year so the user cannot select
'day 30 for month 2
	sub writeDropDowns()
		dim strSelfLink
		strSelfLink = "InvoiceList.asp?sort=" & request("sort") & "&page=" & request("page")
		write "<form name=dates method=post>" & vbcrlf
		write MonthDropDown("month1",true,request("month1"),strSelfLink) & " " & DayDropDown("day1", "",getDaysInMonth(request("month1"),request("year1")),request("day1")) & " " & YearDropDown("year1","","", request("year1"),strSelfLink) & _
		" To " & MonthDropDown("month2",true, request("month2"),strSelfLink) & " " & DayDropDown("day2", "",getDaysInMonth(request("month2"),request("year2")),request("day2")) & " " & YearDropDown("year2","","", request("year2"),strSelfLink) & vbcrlf
	 write "<a href='javascript: fnSubmit(" & chr(34) & strSelfLink& "&datechange=true" & chr(34) & ",1)'>Submit</a>"
		write "</form>"	& vbcrlf
	end sub
	'write MonthDropDown("Month1",true)
	function MonthDropDown(strName, blnNum, strSelected, strSelfLink) 'if blnNUM is true, then show as numbers
		dim strTemp, i, strSelectedString
		strTemp = "<select name='" & strName& "' onchange='javascript: fnSubmit(" & chr(34) & strSelfLink & chr(34) & ",0)'>" & vbcrlf
		strTemp = strTemp & "<option value='" & 0 & "'>" & "Month" & "</option>" & vbcrlf
		for i = 1 to 12
			if strSelected = cstr(i) then
				strSelectedString = "Selected"
			else
				strSelectedString = ""
			end if
			if blnNum then
				strTemp = strTemp & "<option value='" & i & "' " & strSelectedString & " >" & i & "</option>" & vbcrlf
			else
				strTemp = strTemp & "<option value='" & i & "' " & strSelectedString & " >" & MonthName(i) & "</option>" & vbcrlf
			end if
		next
		strTemp = strTemp & "</select>" & vbcrlf
		MonthDropDown = strTemp
	end function
	'write YearDropDown("Year1", 2001, 2010)
	function YearDropDown(strName, intStartYear, intEndYear, strSelected, strSelfLink)
	 dim strTemp, i, strSelectedString
	 if intStartYear = "" then
			intStartYear = Year(now())
		end if
		if intEndYear = "" then
			intEndYear = Year(now()) + 9
		end if
		strTemp = "<select name='" & strName& "' onchange='javascript: fnSubmit(" & chr(34) & strSelfLink & chr(34) & ",0)'>" & vbcrlf
		strTemp = strTemp & "<option value='" & 0 & "'>" & "Year" & "</option>" & vbcrlf
		for i = intStartYear to intEndYear
			if strSelected = cstr(i) then
				strSelectedString = "Selected"
			else
				strSelectedString = ""
			end if
			strTemp = strTemp & "<option value='" & i & "' " & strSelectedString & " >" & i & "</option>" & vbcrlf
		next
		strTemp = strTemp & "</select>" & vbcrlf
		YearDropDown = strTemp
	end function
	'write DayDropDown("Day1",1,getDaysInMonth(2,2001) )
	function DayDropDown(strName, intStartDay, intEndDay, strSelected )
		dim strTemp, i, strSelectedString
		if intStartDay = "" then
			intStartDay = 1
		end if
		if intEndDay = "" then
			intEndDay = getDaysInMonth(Month(now()),Year(now()))
		end if
		strTemp = "<select name='" & strName& "'>" & vbcrlf
		strTemp = strTemp & "<option value='" & 0 & "'>" & "Day" & "</option>" & vbcrlf
		for i = intStartDay to intEndDay
			if strSelected = cstr(i) then
				strSelectedString = "Selected"
			else
				strSelectedString = ""
			end if
			strTemp = strTemp & "<option value='" & i & "' " & strSelectedString & " >" & i & "</option>" & vbcrlf
		next
		strTemp = strTemp & "</select>" & vbcrlf
		DayDropDown = strTemp
	end function
 sub beginDoc(strTitle)
		write "<html>" & vbcrlf
		write "<head>" & vbcrlf
		write "<title>" & strTitle & "</title>" & vbcrlf
		write "</head>" & vbcrlf
		write "<body>" & vbcrlf
	end sub
	sub endDoc()
		write "</body>" & vbcrlf
		write "</html>" & vbcrlf
	end sub
	Const KERMITTHEFROGGREEN = "#beff43"
%>
```

