<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ReviewImageTempUpload.asp - 상품후기 이미지 임지 저장
'Date		: 2018.12.20
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM FileName

DIM SaveFolder
DIM FSO
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

FileName		= Request("FileName")


SaveFolder		= D_REVIEW & "Temp/"


SET FSO = Server.CreateObject("Scripting.FileSystemObject")

IF FSO.FileExists(Server.MapPath(SaveFolder & FileName)) THEN
		FSO.DeleteFile(Server.MapPath(SaveFolder & FileName))
END IF

SET FSO = Nothing

Response.Write "OK|||||"
%>