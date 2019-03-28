<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'LookBookView.asp - Street306
'Date		: 2019.01.15
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

Dim IDX

Dim BrandCode
Dim BrandName
Dim BrandStory
Dim Mobile_TopImage
Dim Mobile_BottomImage
Dim Title1
Dim Title2
Dim Contents
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
IDX = sqlFilter(Request("IDX"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

wQuery = "WHERE A.DelFlag = 'N' AND A.IDX = " & IDX & " "
sQuery = "ORDER BY IDX DESC "
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Street306_LookBook_Select_By_wQuery"
		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
		.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If oRs.EOF Then
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing
	
	Call AlertMessage("잘못 된 경로 입니다.", "history.back();")
	Response.End
Else
	BrandCode = oRs("BrandCode")
	BrandName = oRs("BrandName")
	BrandStory = oRs("BrandStory")
	Mobile_TopImage = oRs("Mobile_TopImage")
	Mobile_BottomImage = oRs("Mobile_BottomImage")
	Title1 = oRs("Title1")
	Title2 = oRs("Title2")
	Contents = oRs("Contents")
End If
oRs.Close
%>						<style>
							.add-img img {width:100% !important;}
  						</style>
						<div class="lookbook-view">
                            <section class="inform">
                                <div class="typical-img">
                                    <img src="<%=Mobile_TopImage%>" alt="<%=BrandName%>">
                                </div>
                                <p class="tit"><%=BrandName%></p>
                                <div class="sub-tit"><%=Title1%></div>
                                <p class="cont">
                                    <%=Title2%>
                                </p>
                                <span class="view-all">
									<a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=BrandCode%>')">브랜드샵 바로가기</a>
								</span>
                            </section>

                            <div class="add-img">
                                <%=Contents%>
                            </div>

                            <section class="a-brand-shop">
                                <img src="<%=Mobile_BottomImage%>" alt="">
                                <div class="tit">
                                    <p><%=BrandName%></p>
                                    <div class="view-all">
                                        <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=BrandCode%>')">BRAND SHOP</a>
                                    </div>
                                </div>
                            </section>
							<!--
                            <div class="a-back-2">
                                <a href="javascript:LookBookClose();">목록으로</a>
                            </div>
							-->
                        </div>
<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>