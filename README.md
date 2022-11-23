# Classic-ASP-Tools

Classic ASP includes I wrote for work.

Also, checkout [ASPJSON](https://github.com/gerritvankuipers/aspjson)


## logger.inc

This reconstructs the path of the file the logger class is initiated in inside `/logs/`.

```
File Path: /dashboard/reports/reportA.asp
Log Path: /logs/dashboard/reports/reportA.asp.log
```

Example:

```vbs
<!--#include virtual="/includes/logger.inc.asp"-->

<%
doLogs = True
doClearOnStart = False
Set l = (New Logger)(doLogs, doClearOnStart)
l.write "This is log one"
l.write "This is log two"
%>
```

## console-log.inc

Example:

```vbs
<!--#include virtual="/includes/console-log.inc.asp"-->

<%
clog("Message")
cwarn("Warning Message")
cerr("Error Mesage")
%>
```

## random.inc

Example:

```vbs
<!--#include virtual="/includes/random.inc.asp"-->

<%
Response.Write(RandomInt(1, 100))
Response.Write(GUID())
%>
```
