<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Brands.asp - 브랜드 리스트
'Date		: 2019.01.15
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
PageCode1 = "BR"
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

%>

<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		.brand-all { position: relative; padding: 15px 10px; border-bottom: 0; }
		.brand-all li { height:auto; }
		.brand-all li>a { position: relative; display: block; width: 100%; height: 100%; margin: 0 auto; border: 0; border-radius: 50%; }
		.brand-all li>a>img { position: relative; width: 100%; max-width: initial; top: initial; left: initial; -webkit-transform: initial; transform: initial; }
	</style>
<!-- #include virtual="/INC/TopMain.asp" -->

    <main id="container" class="container">
        <div class="content">
            <div class="slider-for">
                <div class="wrap-item-list">
<%
'# 메인 브랜드 TOP5
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Brand_Select_For_TopBrand_Top5"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If Not oRs.EOF Then	
%>
                    <section class="brand-main">
                        <div class="swiper-container">
                            <ul class="swiper-wrapper">
							<%
							Do While Not oRs.EOF	
							%>

                                <li class="swiper-slide">
                                    <div class="img">
                                        <img src="<%=oRs("MobileMainLogoImg")%>" alt="<%=oRs("BrandName")%>" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=oRs("BrandCode")%>');">
                                    </div>
                                    <div class="ly-tit">
                                        <div class="brand-name-1"><%=oRs("BrandName")%></div>
                                        <div class="brand-name-2"><%=oRs("BrandNameKor")%></div>
                                    </div>
                                    <!--<a href="/ASP/Product/Brand.asp?SBrandCode=<%=oRs("BrandCode")%>" class="a-more" style="z-index:201">BRAND SHOP</a>-->
                                </li>
							<%
								oRs.MoveNext
							Loop	
							%>
                            </ul>

                            <div class="swiper-pagination ty-red"></div>
                        </div>
                    </section>
<%
End If
oRs.Close	
%>
                    <section class="tit-ty2">
                        <p>브랜드관 바로가기</p>
                    </section>
<%
'# 브랜드 리스트
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

If Not oRs.EOF Then
%>
                    <section class="brand-all">
                        <ul>
						<%
						Do While Not oRs.EOF	
						%>
                            <li>
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=oRs("BrandCode")%>')">
									<img src="<%=oRs("MobileLogoImg")%>" alt="<%=oRs("BrandName")%>">
								</a>
                            </li>
						<%
							oRs.MoveNext
						Loop	
						%>
                        </ul>
                    </section>
<%
End If
oRs.Close	
%>
                </div>
            </div>
        </div>
    </main>


<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
