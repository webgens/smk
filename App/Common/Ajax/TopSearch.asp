<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'TopSearch.asp - 탑 검색
'Date		: 2019.01.12
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정 ------------------------------------------------------------------'
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

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절
DIM sqlQuery

DIM i
DIM j
DIM x
DIM y

Dim vType
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

vType = sqlFilter(Request("vType"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_SearchWord_Manage_Select_By_sType"
		.Parameters.Append .CreateParameter("@sType", adChar, adParamInput, 1, vType)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

Do While Not oRs.EOF
%>

                                            <li>
												<% If oRs("LinkUrl") = "" Then %>
                                                <a href="javascript:APP_GoUrl('/ASP/Product/SearchProductList.asp?SearchWord=<%=oRs("Contents")%>');">
												<% Else %>
												<a href="javascript:APP_GoUrl('<%=Server.UrlEncode(oRs("LinkUrl"))%>');">
												<% End If %>
													<%=oRs("Contents")%>
                                                </a>
                                            </li>
<%
	oRs.MoveNext
Loop
oRs.Close

SET oRs = Nothing
oConn.Close
SET oConn = Nothing	
%>