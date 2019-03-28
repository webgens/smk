<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyReentryList.asp - 재입고 알림 해제
'Date		: 2019.01.06
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

DIM RecCnt
DIM i

DIM Idx
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Idx	= sqlFilter(request("Idx"))


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Reentry_Update_For_DelFlag"

		.Parameters.Append .CreateParameter("@IDX"		,	 adInteger, adParaminput,	 , Idx)
		.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar, adParamInput,  20, U_ID)
		.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

Response.Write "OK|||||알림신청이 해제되었습니다."

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>