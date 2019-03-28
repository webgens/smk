<%
Option Explicit
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<%

Dim strDType, strFile, strDir
strDType = Request("dtype")
strFile	 = Request("file")



SELECT CASE strDType
	CASE "Notice" : strDir = Server.MapPath(D_Notice) 
END SELECT

Response.ContentType = "application/unknown"
Response.AddHeader "Content-Disposition","attachment; filename=" & strFile


Dim objStream, dwnFile, strFileURL
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type = 1


strFileURL = strDir &"\"& strFile
'response.write strFileURL & "<BR>"
'response.end
objStream.LoadFromFile strFileURL

dwnFile = objStream.Read

Response.BinaryWrite dwnFile

Set objStream = Nothing 
%> 
