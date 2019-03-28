<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Get_Category1.asp - 대 카테고리 가져오기
'Date		: 2019.01.04
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

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

If U_Num = "" Then
	wQuery = " WHERE D.GuestInfo = '" & U_GuestInfo & "' "
Else
	wQuery = " WHERE D.MemberNum = " & U_Num & " OR D.GuestInfo = '" & U_GuestInfo & "' "
End If
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Latest_Select_By_wQuery_For_Top30"
		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

Do While Not oRs.EOF
%>
                                    <li>
                                        <a href="javascript:close_ProductLatest();APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>');"><img src="<%=oRs("ImageUrl_180")%>" alt="<%=oRs("ProductName")%>"></a>
                                    </li>
<%
	oRs.MoveNext
Loop
oRs.Close
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>