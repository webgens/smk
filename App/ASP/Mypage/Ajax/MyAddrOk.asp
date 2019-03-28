<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyAddrOk.asp - 배송지 주소록 입력/수정
'Date		: 2019.01.16
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
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

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
DIM RecCnt
DIM AddrType
DIM AddressName
DIM ReceiveName
DIM ReceiveTel
DIM ReceiveTel1
DIM ReceiveTel23
DIM ReceiveHP
DIM ReceiveHP1
DIM ReceiveHP23
DIM ReceiveEmail
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2
DIM MainFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
Idx				 = sqlFilter(Request("Idx"))
RecCnt			 = sqlFilter(Request("RecCnt"))
AddrType		 = sqlFilter(Request("AddrType"))
AddressName		 = sqlFilter(Request("AddressName"))
ReceiveName		 = sqlFilter(Request("ReceiveName"))
ReceiveTel1		 = sqlFilter(Request("ReceiveTel1"))
ReceiveTel23	 = sqlFilter(Request("ReceiveTel23"))
ReceiveHP1		 = sqlFilter(Request("ReceiveHP1"))
ReceiveHP23		 = sqlFilter(Request("ReceiveHP23"))
ReceiveEmail	 = ""
ReceiveZipCode	 = sqlFilter(Request("ReceiveZipCode"))
ReceiveAddr1	 = sqlFilter(Request("ReceiveAddr1"))
ReceiveAddr2	 = sqlFilter(Request("ReceiveAddr2"))
MainFlag		 = sqlFilter(Request("MainFlag"))
ReceiveTel		 = ChgTel(ReceiveTel1 & ReceiveTel23)
ReceiveHP		 = ChgTel(ReceiveHP1 & ReceiveHP23)
IF MainFlag<>"Y" THEN
	MainFlag = "N"
END IF
'# 초기등록 또는 초기등록분 수정 시 기본배송지로 설정
IF (RecCnt=1 AND AddrType="modify") OR (RecCnt=0 AND AddrType="insert") THEN
	MainFlag = "Y"
END IF


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


SET oCmd = Server.CreateObject("ADODB.Command")
IF Idx<>"" AND Idx<>"0" THEN
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_EShop_MyAddress_Update"

			.Parameters.Append .CreateParameter("@Idx",				 adInteger,	 adParamInput,		, Idx)
			.Parameters.Append .CreateParameter("@MemberNum",		 adVarChar,	 adParamInput,	  20, U_NUM)
			.Parameters.Append .CreateParameter("@AddressName",		 adVarChar,	 adParamInput,	  20, AddressName)
			.Parameters.Append .CreateParameter("@ReceiveName",		 adVarChar,	 adParamInput,	  50, ReceiveName)
			.Parameters.Append .CreateParameter("@ReceiveTel",		 adVarChar,	 adParamInput,	  20, ReceiveTel)
			.Parameters.Append .CreateParameter("@ReceiveHp",		 adVarChar,	 adParamInput,    20, ReceiveHp)
			.Parameters.Append .CreateParameter("@ReceiveEmail",	 adVarChar,	 adParamInput,    50, ReceiveEmail)
			.Parameters.Append .CreateParameter("@ReceiveZipCode",	 adVarChar,	 adParamInput,	   7, ReceiveZipCode)
			.Parameters.Append .CreateParameter("@ReceiveAddr1",	 adVarChar,	 adParamInput,   200, ReceiveAddr1)
			.Parameters.Append .CreateParameter("@ReceiveAddr2",	 adVarChar,	 adParamInput,   200, ReceiveAddr2)
			.Parameters.Append .CreateParameter("@MainFlag",		 adChar,	 adParamInput,     1, MainFlag)
			.Parameters.Append .CreateParameter("@UpdateID",		 adVarChar,	 adParamInput,	  20, U_NUM)
			.Parameters.Append .CreateParameter("@UpdateIP",		 adVarChar,	 adParamInput,	  15, U_IP)

			.Execute, , adExecuteNoRecords
	END WITH
ELSE
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_EShop_MyAddress_Insert"

			.Parameters.Append .CreateParameter("@MemberNum",		 adVarChar,	 adParamInput,	  20, U_NUM)
			.Parameters.Append .CreateParameter("@AddressName",		 adVarChar,	 adParamInput,	  20, AddressName)
			.Parameters.Append .CreateParameter("@ReceiveName",		 adVarChar,	 adParamInput,	  50, ReceiveName)
			.Parameters.Append .CreateParameter("@ReceiveTel",		 adVarChar,	 adParamInput,	  20, ReceiveTel)
			.Parameters.Append .CreateParameter("@ReceiveHp",		 adVarChar,	 adParamInput,    20, ReceiveHp)
			.Parameters.Append .CreateParameter("@ReceiveEmail",	 adVarChar,	 adParamInput,    50, ReceiveEmail)
			.Parameters.Append .CreateParameter("@ReceiveZipCode",	 adVarChar,	 adParamInput,	   7, ReceiveZipCode)
			.Parameters.Append .CreateParameter("@ReceiveAddr1",	 adVarChar,	 adParamInput,   200, ReceiveAddr1)
			.Parameters.Append .CreateParameter("@ReceiveAddr2",	 adVarChar,	 adParamInput,   200, ReceiveAddr2)
			.Parameters.Append .CreateParameter("@MainFlag",		 adChar,	 adParamInput,     1, MainFlag)
			.Parameters.Append .CreateParameter("@CreateID",		 adVarChar,	 adParamInput,	  20, U_NUM)
			.Parameters.Append .CreateParameter("@CreateIP",		 adVarChar,	 adParamInput,	  15, U_IP)

			.Execute, , adExecuteNoRecords
	END WITH
END IF
SET oCmd = Nothing

IF Err.number <> 0 THEN
	oConn.Close : SET oConn = Nothing
	Response.Write "FAIL|||||처리 도중 오류가 발생하였습니다."
	Response.End
END IF


IF Idx<>"" AND Idx<>"0" THEN
	Response.Write "OK|||||수정이 완료되었습니다."
ELSE
	Response.Write "OK|||||입력이 완료되었습니다."
END IF

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>