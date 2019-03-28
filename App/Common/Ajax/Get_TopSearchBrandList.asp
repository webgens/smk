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

Dim TopBrandSearchWord
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

TopBrandSearchWord = sqlFilter(Request("TopBrandSearchWord"))

If TopBrandSearchWord = "" Then
	Response.Write "FAIL|||||"&TopBrandSearchWord&"찾으시는 브랜드를 입력하여 주세요."
	Response.End
End If
	
SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>

<%
Response.Write "OK|||||"
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Brand_Select_For_SearchWord"

		.Parameters.Append .CreateParameter("@SearchWord", adVarChar, adParamInput, 100, TopBrandSearchWord)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF oRs.EOF Then
%>
                                <p class="tit">총 <em>0</em>개의 브랜드</p>
                                <div class="result">
                                    <p class="empty">찾으시는 브랜드가 없습니다.</p>
                                </div>
<%
Else
%>
                                <p class="tit">총 <em><%=oRs.RecordCount %></em>개의 브랜드</p>
                                <div class="result" style="line-height:30px;">
<%
	Do While Not oRs.EOF
%>
                                    <a href="javascript:APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=oRs("BrandCode")%>');GetCategory1Close();" class="brandName"><%=oRs("BrandName")%></a> <br />
<%
		oRs.MoveNext
	Loop
%>
                                </div>
<%
End If
oRs.Close

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>