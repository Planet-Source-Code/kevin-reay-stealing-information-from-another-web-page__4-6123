<div align="center">

## Stealing Information from Another Web Page \(


</div>

### Description

This code steals info/output from other pages! This can be used create meta-searches by grabing the output of other pages!!! For example, you could pass a search string into 2 existing pages and return the results to a single page!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kevin Reay](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kevin-reay.md)
**Level**          |Intermediate
**User Rating**    |3.5 (21 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kevin-reay-stealing-information-from-another-web-page__4-6123/archive/master.zip)





### Source Code

```
<% Option Explicit %>
<%
Dim url    'The URL to download
Dim sInfo   'string to hold the collected info
Dim sHTML   'String to hold HTML from download
Dim rReg   'var to hold regular expression
Dim objCols  'Object to hold collections from Regular expression
Dim objMatch 'Object for matches
Dim inet   'Object for Inet Control
url = "WhatEverURL"
'Create instance of Inet Control
Set inet = Server.CreateObject("InetCtls.Inet")
'Set the timeout
inet.RequestTimeOut = 20
'Set the URL property of the control
inet.Url = url
'Actually download the file
sHTML = inet.OpenURL()
'Regular expression to find the string stored between
'the tags. This is where information is.
Set rReg = New regexp
'the TagGoesHere and EndTagGoesHere tags below represent
'the tags surrounding the information we want
'these tags can be more complex if required
rReg.Pattern = "TagGoesHere(.*)EndTagGoesHere"
rReg.Global = False
rReg.IgnoreCase = True
'Execute the regular expression on the raw HTML
Set objCols = rReg.Execute( sHTML )
'Step through our matches
For Each objMatch in objCols
	sInfo = sInfo & objMatch.Value
Next
'Clean up
Set rWorldPop = Nothing
Set objCols = Nothing
'Strip the TagGoesHere tags off of the info
sInfo = Replace(Replace(sInfo, "TagGoesHere", ""), "EndTagGoesHere", "")
%>
<HTML>
<HEAD>
<TITLE>Web output Stealer</TITLE>
</HEAD>
<BODY>
<P>The output of the page is: <%=sInfo %></P>
</BODY>
</HTML>
```

