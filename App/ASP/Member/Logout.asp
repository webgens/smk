<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Logout.asp - 로그아웃 처리 페이지
'Date		: 2018.10.30
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드----------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "01"
PageCode2 = "01"
PageCode3 = "99"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'





'*****************************************************************************************'
'쿠키 처리 START
'-----------------------------------------------------------------------------------------'
Response.Cookies("UIP").Expires			 = Now - 1000
Response.Cookies("UMFLAG").Expires		 = Now - 1000
Response.Cookies("UNUM").Expires		 = Now - 1000
Response.Cookies("UID").Expires			 = Now - 1000
Response.Cookies("UNAME").Expires		 = Now - 1000
Response.Cookies("UETYPE").Expires		 = Now - 1000
Response.Cookies("UETYPE").Expires		 = Now - 1000
Response.Cookies("UGROUP").Expires		 = Now - 1000
'-----------------------------------------------------------------------------------------'
'쿠키 처리 END
'-----------------------------------------------------------------------------------------'

Call AlertMessage2("로그아웃 되었습니다.", "APP_goMain();")
%>
