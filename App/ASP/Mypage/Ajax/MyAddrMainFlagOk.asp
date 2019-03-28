<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyAddrOk.asp - 마이페이지 > 배송지 주소록 입력/수정
'Date		: 2018.12.17
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
PageCode1 = "05"
PageCode2 = "05"
PageCode3 = "02"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->
<!-- #include virtual="/Common/md5.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절


DIM Idx
DIM MainFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
Idx				 = sqlFilter(Request("Idx"))


SET oConn		 = ConnectionOpen()							'# 커넥션 생성


SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_EShop_MyAddress_MainFlag_Update"

			.Parameters.Append .CreateParameter("@Idx",				 adInteger,	 adParamInput,		, Idx)
			.Parameters.Append .CreateParameter("@MemberNum",		 adVarChar,	 adParamInput,	  20, U_NUM)
			.Parameters.Append .CreateParameter("@UpdateID",		 adVarChar,	 adParamInput,	  20, U_NUM)
			.Parameters.Append .CreateParameter("@UpdateIP",		 adVarChar,	 adParamInput,	  15, U_IP)

			.Execute, , adExecuteNoRecords
	END WITH

SET oCmd = Nothing

IF Err.number <> 0 THEN
	oConn.Close : SET oConn = Nothing
	Response.Write "FAIL|||||처리 도중 오류가 발생하였습니다."
	Response.End
END IF


Response.Write "OK|||||기본 배송지로 설정 되었습니다."

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>