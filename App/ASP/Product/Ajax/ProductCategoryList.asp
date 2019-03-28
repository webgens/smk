<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ProductList.asp - 상품리스트
'Date		: 2019.01.05
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
PageCode1 = "00"
PageCode2 = "00"
PageCode3 = "00"
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
DIM oRs1						'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절

Dim SCode1
Dim SCode2
Dim SCode3
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SCode1 = sqlFilter(Request("SCode1"))
SCode2 = sqlFilter(Request("SCode2"))
SCode3 = sqlFilter(Request("SCode3"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

If SCode2 = "" Then

	SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_EShop_Category2_Select"
			.Parameters.Append .CreateParameter("@CategoryCode1", adChar, adParamInput, 2, SCode1)
	END WITH
	oRs.CursorLocation = adUseClient
	oRs.Open oCmd, , adOpenStatic, adLockReadOnly
	SET oCmd = Nothing

	If oRs.EOF Then
%>	
		<button type="button" class="on" onclick="location.href='/ASP/Product/ProductList.asp?SCode1=<%=SCode1%>';"><span>전체</span></button>
<%
	Else
%>
		<button type="button" class="on" onclick="location.href='/ASP/Product/ProductList.asp?SCode1=<%=SCode1%>';"><span>전체</span></button>

<%
		Do While Not oRs.EOF
%>
			<button type="button" onclick="location.href='/ASP/Product/ProductList.asp?SCode1=<%=SCode1%>&SCode2=<%=oRs("CategoryCode2")%>';"><span><%=oRs("CategoryName2")%></span></button>
<%
			oRs.MoveNext
		Loop
	End If
	oRs.Close

Else

	SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_EShop_Category3_Select"
			.Parameters.Append .CreateParameter("@CategoryCode1", adChar, adParamInput, 2, SCode1)
			.Parameters.Append .CreateParameter("@CategoryCode2", adChar, adParamInput, 2, SCode2)
	END WITH
	oRs.CursorLocation = adUseClient
	oRs.Open oCmd, , adOpenStatic, adLockReadOnly
	SET oCmd = Nothing

	If oRs.EOF Then
%>	
		<button type="button" class="on" onclick="location.href='/ASP/Product/ProductList.asp?SCode1=<%=SCode1%>&SCode2=<%=SCode2%>';"><span>전체</span></button>
<%
	Else
%>
			<button type="button" <% If SCode3 = "" Then %>class="on"<% End If %> onclick="location.href='/ASP/Product/ProductList.asp?SCode1=<%=SCode1%>&SCode2=<%=SCode2%>';"><span>전체</span></button>

<%
		Do While Not oRs.EOF
%>
			<button type="button" <% If SCode3 = oRs("CategoryCode3") Then %>class="on"<% End If %> onclick="location.href='/ASP/Product/ProductList.asp?SCode1=<%=SCode1%>&SCode2=<%=SCode2%>&SCode3=<%=oRs("CategoryCode3")%>';"><span><%=oRs("CategoryName3")%></span></button>
<%
			oRs.MoveNext
		Loop
	End If
	oRs.Close

End If
	
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
