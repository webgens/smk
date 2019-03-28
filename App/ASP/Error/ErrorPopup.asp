<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************/
'ErrorPopup.asp - 에러알림팝업
'Date		: 2018.12.28
'Update		: 
'/****************************************************************************************/

'//페이지 응답헤더 설정------------------------------------------------------
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//-------------------------------------------------------------------------------

'//페이지 코드-----------------------------------------------------------------
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "ER"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'//-------------------------------------------------------------------------------
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include virtual = "/Common/ProgID1.asp" -->

<%
'/****************************************************************************************/
'변수 선언 START
'-----------------------------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

Dim x
DIM i
DIM j
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM Title
DIM Msg
DIM Script
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


Title					 = sqlFilter(Request("Title"))
Msg						 = sqlFilter(Request("Msg"))
Script					 = sqlFilter(Request("Script"))

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
%>


<!--#INCLUDE VIRTUAL = "/INC/Header.asp"-->
<!--#INCLUDE VIRTUAL = "/INC/TopSub.asp"-->

	<main id="container" class="container">
	</main>

<!--#INCLUDE VIRTUAL = "/INC/Footer.asp"-->
<!--#INCLUDE VIRTUAL = "/INC/Bottom.asp"-->

<script type="text/javascript">
		common_msgPopOpen("<%=Title%>", "<%=Msg%>", "<%=Script%>", "msgPopup", "N");
</script>


<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>