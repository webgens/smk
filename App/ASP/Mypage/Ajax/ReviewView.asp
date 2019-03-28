<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ReviewView.asp - 구매후기 조회 폼 페이지
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

<%
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 다시 로그인하여 주십시오."
		Response.End
END IF

'/****************************************************************************************'
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


DIM Idx

DIM Order_Product_IDX
DIM ReviewType
DIM SizeGrade
DIM WearGrade
DIM DesignGrade
DIM QualityGrade
DIM AvgGrade
DIM Contents

DIM ProductCode
DIM ProductName
DIM SizeCD
DIM BrandName
DIM ProductImage
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


Idx					= sqlFilter(Request("Idx"))




IF Idx = "" THEN
		Response.Write "FAIL|||||선택한 상품후기 정보가 없습니다."
		Response.End
END IF



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'-----------------------------------------------------------------------------------------------------------'
'# 상품후기 정보 Start
'-----------------------------------------------------------------------------------------------------------'
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Product_Review_Select_By_IDX"

		.Parameters.Append .CreateParameter("@IDX",		 adInteger, adParaminput, 	,	Idx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		Order_Product_IDX	= oRs("Order_Product_IDX")
		ReviewType			= oRs("ReviewType")
		SizeGrade			= oRs("SizeGrade")
		WearGrade			= oRs("WearGrade")
		DesignGrade			= oRs("DesignGrade")
		QualityGrade		= oRs("QualityGrade")
		AvgGrade			= oRs("AvgGrade")
		Contents			= oRs("Contents")
ELSE
		Response.Write "FAIL|||||상품후기 정보가 없습니다."
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'# 상품후기 정보 End
'-----------------------------------------------------------------------------------------------------------'

'-----------------------------------------------------------------------------------------------------------'
'# 주문상품 정보 Start
'-----------------------------------------------------------------------------------------------------------'
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Order_Product_Select_By_Idx"

		.Parameters.Append .CreateParameter("@Idx",		 adInteger, adParaminput, 	,	Order_Product_IDX)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		BrandName		= oRs("BrandName")
		SizeCD			= oRs("SizeCD")

		IF oRs("ProductImage_180") = "" THEN
				ProductImage	= "/Images/180_noimage.png"
		ELSE
				ProductImage	= oRs("ProductImage_180")
		END IF
ELSE
		Response.Write "FAIL|||||주문정보가 없습니다."
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'# 주문상품 정보 End
'-----------------------------------------------------------------------------------------------------------'

Response.Write "OK|||||"
%>					

		<div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">상품 후기</div>
                    <div class="btn-hide-pop" onclick="closePop('DimDepth1');">닫기</div>
                </div>

                <div class="container-pop mypage-ty2">
                    <div class="contents">
                        <div class="wrap-review">
                            <div class="informView">
								<%IF ReviewType="P" THEN%>
								<div class="exclamation-mark">
									<p>포토후기 <em><%=MALL_REVIEW_POINT_P%> 포인트</em>가 지급되었습니다.</p>
								</div>
								<%ELSE%>
								<div class="exclamation-mark">
									<p>일반후기 <em><%=MALL_REVIEW_POINT_B%> 포인트</em>가 지급되었습니다.</p>
								</div>
								<%END IF%>
                                <div class="informItem">
                                    <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=ProductCode%>">
										<span class="cont">
											<span class="thumbNail">
												<span class="img">
													<img src="<%=ProductImage%>" alt="상품 이미지">
												</span>
											</span>
											
											<span class="detail">
												<span class="brand">
													<span class="name"><%=BrandName%></span>
												</span>
												<span class="product-name"><em><%=ProductName%></em></span>
												
												<span class="inform">
													<span class="list">
														<span class="tit">옵션</span>
														<span class="opt"><%=SizeCD%></span>
													</span>
												</span>
											</span>
										</span>
									</a>
                                </div>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">나의 평가</h3>
                            </div>

                            <!-- psd_181212수정 -->
                            <ul class="assessment">
                                <li class="star-score">
                                    <span class="tit">사이즈</span>
									<span class="point val<%=SizeGrade * 2%>0"></span>
                                    <span class="score"><%=SizeGrade%></span>
                                </li>
                                <li class="star-score">
                                    <span class="tit">착화감</span>
									<span class="point val<%=WearGrade * 2%>0"></span>
                                    <span class="score"><%=WearGrade%></span>
                                </li>
                                <li class="star-score">
                                    <span class="tit">디자인</span>
									<span class="point val<%=DesignGrade * 2%>0"></span>
                                    <span class="score"><%=DesignGrade%></span>
                                </li>
                                <li class="star-score">
                                    <span class="tit">품질</span>
									<span class="point val<%=QualityGrade * 2%>0"></span>
                                    <span class="score"><%=QualityGrade%></span>
                                </li>
                            </ul>
                            <!-- psd_181212수정 -->

                            <div class="h-line">
                                <h3 class="h-level4">후기 내용</h3>
                            </div>

                            <div class="review-write">
                                <div class="input" style="height:auto; padding:10px;">
                                    <%=ReplaceDetails(Contents)%>
                                </div>
                            </div>

<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Review_Image_Select_By_ReviewIdx"

		.Parameters.Append .CreateParameter("@ReviewIdx",		 adInteger, adParaminput, 	,	Idx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
%>
                            <div class="added-photo" style="margin:0; padding:10px;">
<%
		Do Until oRs.EOF
%>
								<img src="<%=D_REVIEW%><%=oRs("FileName")%>" style="width:100%;" alt="후기 이미지">
<%
				oRs.MoveNext
		Loop
%>
                            </div>
<%
END IF
oRs.Close
%>
                            <div class="inf-type1">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">포토후기의 경우 직접 촬영한 사진이 아닐 경우 당첨과 쿠폰이 취소됩니다.</li>
                                    <li class="bullet-ty1">상품후기와 관련없는 내용일 경우 관리자에 의해 통보 없이 미등록, 삭제 될 수 있습니다.</li>
                                </ul>
                            </div>
                        </div>
                    </div>

                    <div class="btns">
                        <button type="button" class="button ty-black" onclick="closePop('DimDepth1');">창닫기</button>
                    </div>
                </div>
            </div>
	    </div>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>