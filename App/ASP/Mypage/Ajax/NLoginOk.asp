<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'NLoginOk.asp - 비회원 로그인 처리 - 주문/배송 조회용
'Date		: 2019.01.17
'Update	: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드----------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "00"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절
	
DIM Name
DIM HP
DIM Email
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	

Name			 = sqlFilter(Request("Name"))
HP				 = sqlFilter(Request("HP1")) & "-" & sqlFilter(Request("HP2")) & "-" & sqlFilter(Request("HP3"))
Email			 = TRIM(sqlFilter(Request("Email")))


IF Name = "" THEN
		Response.Write "FAIL|||||이름을 입력하여 주십시오."
		Response.End
END IF
IF HP = "--" THEN
		Response.Write "FAIL|||||휴대폰번호를 입력하여 주십시오."
		Response.End
END IF
IF Email = "" THEN
		Response.Write "FAIL|||||이메일을 입력하여 주십시오."
		Response.End
END IF


Response.Write "OK|||||"



Response.Cookies("N_NAME")	 = Name
Response.Cookies("N_HP")	 = HP
Response.Cookies("N_EMAIL")	 = Email
%>