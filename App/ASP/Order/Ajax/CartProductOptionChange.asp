<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'CartProductOptionChange.asp - 장바구니 옵션 변경 폼 페이지
'Date		: 2018.12.27
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
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM CartIdx

DIM ProductCode
DIM ProductName
DIM SizeCD
DIM BrandName
DIM ProductImage
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


CartIdx			= sqlFilter(Request("CartIdx"))




IF CartIdx = "" THEN
		Response.Write "FAIL|||||선택한 상품이 없습니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_Select_By_Idx"

		.Parameters.Append .CreateParameter("@Idx", adInteger, adParamInput, , CartIdx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		BrandName		= oRs("BrandName")
		SizeCD			= oRs("SizeCD")
		ProductImage	= oRs("ProductImage_180")

		IF ProductImage = "" THEN
				ProductImage	= "/Images/180_noimage.png"
		END IF
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||선택한 상품이 없습니다."
		Response.End
END IF
oRs.Close

Response.Write "OK|||||"
%>					
        <div class="area-dim"></div>

        <div class="area-pop" id="CartOptionChangeLayer">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">옵션 변경</p>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents ofh">
                        <!-- ofh 더블 클래스 사용, 팝업에서 콘텐츠가 height 100%보다 클 때, 전체 스크롤 생기는 것을 방지 -->
                        <div class="wrap-change-option h100p">
                            <div class="ly-items">
                                <div class="thumbNail">
                                    <img src="<%=ProductImage%>" alt="<%=ProductName%>">
                                </div>
                                <div class="cont">
                                    <span class="brand"><%=BrandName%></span>
                                    <span class="line"><%=ProductName%></span>
                                </div>
                            </div>

                            <div class="select-option">
                                <div class="footSize-table">
                                    <div id="footSize_all" class="accordion">
                                        <div class="selector is-focus">
                                            <button type="button" class="btn-select" data-target="footSize_all">
											<span>사이즈 선택</span>
										</button>
                                        </div>
                                        <div class="option" style="display:block">
                                            <div class="pop-size">
											<%
											SET oCmd = Server.CreateObject("ADODB.Command")
											WITH oCmd
													.ActiveConnection	 = oConn
													.CommandType		 = adCmdStoredProc
													.CommandText		 = "USP_Front_EShop_Product_SizeCD_Select_With_EShop_Stock"

													.Parameters.Append .CreateParameter("@ProductCode", adInteger,	adParamInput,  , ProductCode)
											END WITH
											oRs.CursorLocation = adUseClient
											oRs.Open oCmd, , adOpenStatic, adLockReadOnly
											SET oCmd = Nothing

											IF NOT oRs.EOF THEN
													Do Until oRs.EOF
															IF oRs("SizeCD") = SizeCD THEN
											%>
                                                <span class="check-style"><input type="radio" name="cSizeCD" id="cSizeCD_<%=oRs("SizeCD")%>" value="<%=oRs("SizeCD")%>" disabled /><label for="cSizeCD_<%=oRs("SizeCD")%>"><span style="color:#ff201b"><%=oRs("SizeCD")%></span></label></span>
											<%
															ELSEIF oRs("StockCnt") > 0 THEN
											%>
                                                <span class="check-style"><input type="radio" name="cSizeCD" id="cSizeCD_<%=oRs("SizeCD")%>" value="<%=oRs("SizeCD")%>" /><label for="cSizeCD_<%=oRs("SizeCD")%>"><span><%=oRs("SizeCD")%></span></label></span>
											<%
															ELSE
											%>
                                                <!--<span class="check-style"><input type="radio" name="cSizeCD" id="cSizeCD_<%=oRs("SizeCD")%>" value="<%=oRs("SizeCD")%>" disabled /><label for="cSizeCD_<%=oRs("SizeCD")%>"><span><%=oRs("SizeCD")%></span></label></span>-->
											<%
															END IF

															oRs.MoveNext
													Loop
											END IF
											oRs.Close
											%>
                                            </div>
                                        </div>
                                    </div>
                                </div>
								<!--
                                <div class="check-amount">
                                    <div id="checkAmount" class="accordion">
                                        <div class="selector">
                                            <button type="button" class="btn-select clickEvt" data-target="checkAmount">
											<span>수량 선택</span>
										</button>
                                        </div>
                                        <div class="option">
                                            <div class="selected-cont">
                                                <div class="cont">
                                                    <div class="amount">
                                                        <button type="button" class="btn-minus">-</button>
                                                        <span class="product-length">1</span>
                                                        <button type="button" class="btn-plus">+</button>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
								-->
                            </div>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="changeOption('<%=CartIdx%>')" class="button ty-red">확인</button>
                    </div>
                </div>
            </div>
        </div>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>