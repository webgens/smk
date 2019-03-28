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
DIM oRs1						'# ADODB Recordset 개체
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
SET oRs1		 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>
                    <div class="contents bg-ty1">
                        <div class="wrap-depth-sel">
                            <!-- *** 수정 *** 190123 : 카테고리(햄버거 메뉴) 선택 수정 -->
                            <p class="h-level6">카테고리</p>
                            <div class="area-accord" id="CategoryView">

<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Category1_Select"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

i = 1
Do While Not oRs.EOF
%>

                                <div id="category_<%=i%>" class="categoryList">
                                    <div class="ly-title">
                                        <button type="button" class="btn-list clickAct" data-target="category_<%=i%>"><%=oRs("CategoryName1")%></button>
                                    </div>
                                    <div class="ly-content">
                                        <ul class="sub-menu">
                                            <li><a href="/ASP/Product/ProductList.asp?SCode1=<%=oRs("CategoryCode1")%>"><%=oRs("CategoryName1")%> 전체</a></li>

<%
	SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Mobile_EShop_Category2_Select"
			.Parameters.Append .CreateParameter("@CategoryCode1", adChar, adParamInput, 2, oRs("CategoryCode1"))
	END WITH
	oRs1.CursorLocation = adUseClient
	oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
	SET oCmd = Nothing

	Do While Not oRs1.EOF
%>
                                            <li><a href="/ASP/Product/ProductList.asp?SCode1=<%=oRs("CategoryCode1")%>&SCode2=<%=oRs1("CategoryCode2")%>"><%=oRs1("CategoryName2")%></a></li>
<%
		oRs1.MoveNext
	Loop
	oRs1.Close	
%>
                                        </ul>
                                    </div>
                                </div>
<%
	i = i + 1
	oRs.MoveNext
Loop
oRs.Close
%>
                            </div>

                            <div class="banner-menu">
                                <a href="/ASP/Street306/" class="street"><img src="/images/img/logo-street.png" alt="Street306"></a>
                                <a href="/ASP/ShoemarkerOnly/" class="only"><img src="/images/img/logo-only.png" alt="SHOEMARKER ONLY"></a>
                                <a href="/ASP/Product/Brand.asp?SBrandCode=TE" class="teva"><img src="/images/img/logo-teva.png" alt="TeVa"></a>
                            </div>

							<form name="TopBrandSearchForm" id="TopBrandSearchForm" method="get" onsubmit="return false;">
                            <div class="fieldset ty-col3">
                                <label class="fieldset-label">브랜드 검색</label>
                                <div class="fieldset-row">
                                    <button type="button" class="button ty-black" onclick="TopBrandSearch();">검색</button>
                                    <span class="input">
                                        <input type="text" title="브랜드명을 입력해 주세요." placeholder="브랜드명을 입력해 주세요." name="TopBrandSearchWord" id="TopBrandSearchWord">
                                    </span>
                                </div>
                            </div>
							</form>

                            <div class="area-result" id="BrandSearchView" style="display:none;">
                            </div>


                            <p class="h-level6">전체 브랜드</p>
<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Brand_Select_For_Prefix"

		.Parameters.Append .CreateParameter("@Prefix", adChar, adParamInput, 1, "")
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
%>
                            <ul class="txtList">
<%
	Do While Not oRs.EOF	
%>

                                <li>
                                    <a href="javascript:APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=oRs("BrandCode")%>');GetCategory1Close();"><%=oRs("BrandName")%></a>
                                </li>
<%
		oRs.MoveNext
	Loop
oRs.Close
%>

                            </ul>
                            <!-- *** 수정 *** 190123 : 카테고리 선택 수정 -->
                        </div>
                    </div>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>