<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'RefundAccountDelOk.asp - 마이페이지 > 회원정보 수정 > 환불계좌 삭제 처리
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

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

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


Dim Idx
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Idx					= sqlFilter(Request("Idx"))


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


'-----------------------------------------------------------------------------------------------------------'	
'# 환불계좌 삭제 Start
'-----------------------------------------------------------------------------------------------------------'	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_RefundAccount_Delete"

		.Parameters.Append .CreateParameter("@Idx",						adInteger,	adParamInput,     ,	 Idx)
		.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 U_NUM)

		.Execute, , adExecuteNoRecords

END WITH
SET oCmd = Nothing


IF Err.Number <> 0 THEN

		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||환불계좌 삭제 중 오류가 발생하였습니다."
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'	
'# 상품후기 삭제 End
'-----------------------------------------------------------------------------------------------------------'	

Response.Write "OK|||||삭제되었습니다."


SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>