﻿<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'Brand_Pick_Delete.asp - 브랜드 찜하기 삭제
'Date		: 2019.01.06
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
<!-- #include virtual = "/Common/CheckID_Ajax.asp" -->

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

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM BrandCode
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

	
BrandCode			= sqlFilter(Request("BrandCode"))


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


IF BrandCode = "ALL" THEN

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Product_Brand_Pick_Delete_For_ALL_By_MemberNum"

				.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 U_NUM)
				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing

ELSEIF BrandCode <> "" THEN

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Product_Brand_Pick_Delete"

				.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 U_NUM)
				.Parameters.Append .CreateParameter("@BrandCode",			adVarChar,	adParamInput,   10,	 BrandCode)
				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
END IF


Response.Write "OK|||||"

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>