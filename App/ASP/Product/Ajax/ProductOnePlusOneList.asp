<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductOnePlusOneList.asp - 1+1상품목록 페이지
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

DIM ProductCode
DIM SizeCD

DIM ProductImage
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


ProductCode		= sqlFilter(Request("ProductCode"))
SizeCD			= sqlFilter(Request("SizeCD"))




IF ProductCode = "" THEN
		Response.Write "FAIL|||||선택한 상품이 없습니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs1 = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



Response.Write "OK|||||"
%>					
        <!-- 1+1 상품 선택 POP -->
        <div class="area-dim"></div>

        <div class="area-pop" id="ProductOnePlusOneLayer">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">1+1 상품 선택</p>
                    <button onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="onePlusOne-list">
                            <div class="onePlus-item-list">
								<%
								SET oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection	 = oConn
										.CommandType		 = adCmdStoredProc
										.CommandText		 = "USP_Front_EShop_Product_OnePlusOne_Select_By_ProductCode"

										.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode)
								END WITH
								oRs.CursorLocation = adUseClient
								oRs.Open oCmd, , adOpenStatic, adLockReadOnly
								SET oCmd = Nothing

								IF NOT oRs.EOF THEN
										i = 0
										Do Until oRs.EOF

												IF oRs("ProductImage") = "" THEN
														ProductImage	= "/Images/180_noimage.png"
												ELSE
														ProductImage	= oRs("ProductImage")
												END IF
								%>
                                <span class="inform">
									<input type="radio" name="OProductCode" id="OProductCode<%=i%>" onclick="openOSizeList('<%=i%>')" data-num="<%=i%>" data-name="<%=oRs("ProductName")%>" value="<%=oRs("ProductCode")%>" />
									<label for="OProductCode<%=i%>">
										<span class="img">
											<img src="<%=ProductImage%>" alt="">
										</span>
										<span class="cont">
											<span class="brand"><%=oRs("BrandName")%></span>
											<span class="line"><%=oRs("ProductName")%></span>
										</span>
									</label>
                                </span>

								<div class="footSize-table">
									<div class="accordion">
										<div class="option" id="OSizeList<%=i%>">
											<div class="pop-size">
												<%
												SET oCmd = Server.CreateObject("ADODB.Command")
												WITH oCmd
														.ActiveConnection	 = oConn
														.CommandType		 = adCmdStoredProc
														.CommandText		 = "USP_Front_EShop_Product_SizeCD_Select_With_EShop_Stock"

														.Parameters.Append .CreateParameter("@ProductCode", adInteger,	adParamInput,  , oRs("ProductCode"))
												END WITH
												oRs1.CursorLocation = adUseClient
												oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
												SET oCmd = Nothing

												IF NOT oRs1.EOF THEN
														j = 1
														Do Until oRs1.EOF
																IF oRs1("StockCnt") > 0 THEN
												%>
												<span class="check-style"><input type="radio" name="OSizeCD<%=i%>" id="OSizeCD<%=i%>_<%=j%>" value="<%=oRs1("SizeCD")%>"><label for="OSizeCD<%=i%>_<%=j%>"><span><%=oRs1("SizeCD")%></span></label></span>
												<%
																ELSE
												%>
												<span class="check-style"><input type="radio" name="OSizeCD<%=i%>" id="OSizeCD<%=i%>_<%=j%>" value="<%=oRs1("SizeCD")%>" disabled><label for="OSizeCD<%=i%>_<%=j%>"><span><%=oRs1("SizeCD")%></span></label></span>
												<%
																END IF

																oRs1.MoveNext
																j = j + 1
														Loop
												END IF
												oRs1.Close
												%>
											</div>
										</div>
									</div>
								</div>
								<%
												oRs.MoveNext
												i = i + 1
										Loop 
								END IF
								oRs.Close
								%>
                            </div>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="selectOnePlusOne('<%=ProductCode%>','<%=SizeCD%>')" class="button ty-red">확인</button>
                    </div>
                </div>
            </div>
        </div>

		<script type="text/javascript">
			function openOSizeList(num) {
				$("#ProductOnePlusOneLayer .footSize-table .option").slideUp('fast');

				if ($("#OSizeList" + num).css("display") == "none") {
					$("#OSizeList" + num).slideDown('fast');
				} else {
					$("#OSizeList" + num).slideUp('fast');
				}
			}
		</script>

<%
Set oRs = Nothing
Set oRs1 = Nothing
oConn.Close
Set oConn = Nothing
%>