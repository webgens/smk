<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'AfterServiceWriteOk.asp - A/S 등록 처리
'Date		: 2019.01.23
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderCode
DIM Order_Product_IDX
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

OrderCode = sqlFilter(Request("OrderCode"))
Order_Product_IDX = sqlFilter(Request("Order_Product_IDX"))


wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'P' AND A.OrderState IN ('1', '3', '4', '5', '6', '7') "
wQuery = wQuery & "AND A.Idx = " & Order_Product_IDX & " "

sQuery = ""

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Order_Detail"

		.Parameters.Append .CreateParameter("@WQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@SQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF oRs.EOF THEN
	oRs.Close
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing

	Response.Write "FAIL|||||주문정보가 없습니다."
	Response.End
END IF
Response.Write "OK|||||"
'/****************************************************************************************/
'회원 배송지정보 SELECT END
'-----------------------------------------------------------------------------------------------------------'
%>

	    <!-- PopUp -->
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">A/S 신청하기</div>
                    <button type="button" class="btn-hide-pop" onclick="closeAfterService();">닫기</button>
                </div>

                <div class="container-pop mypage-ty2">
                    <!-- 팝업 스타일 변경으로 'mypage-ty2'클래스 명 추가 -->
                    <div class="contents">
						<form name="ASForm" id="ASForm" method="post">
						<input type="hidden" name="OrderCode"			id="OrderCode"			value="<%=OrderCode%>" />
						<input type="hidden" name="OrderProductIDX"		id="OrderProductIDX"	value="<%=Order_Product_IDX%>" />
						<input type="hidden" name="ProductCode"			id="ProductCode"		value="<%=oRs("ProductCode")%>" />
						<input type="hidden" name="ShopCD"				id="ShopCD"				value="<%=oRs("ShopCD")%>" />
						<input type="hidden" name="UploadFiles"			id="UploadFiles"		value="" />
						<input type="hidden" name="UploadFilesCount"	id="UploadFilesCount"	value="0" />
                        <div class="wrap-review">
                            <div class="informView">
                                <div class="informItem">
                                    <a href="#">
										<span class="cont">
											<span class="thumbNail">
												<span class="img">
													<img src="<%=oRs("ProductImage")%>" alt="<%=oRs("ProductName")%>">
												</span>
											</span>
											
											<span class="detail">
												<span class="brand">
													<span class="name"><%=oRs("BrandName")%></span>
													<span class="item-code"><%=oRs("ProdCD")%>=<%=oRs("ColorCD")%></span>
												</span>
												<span class="product-name"><em><%=oRs("ProductName")%></em></span>
												
												<span class="inform">
													<span class="list">
														<span class="tit">옵션</span>
														<span class="opt"><%=oRs("SizeCD")%></span>
													</span>
												</span>
											</span>
										</span>
									</a>
                                </div>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">고객 정보 확인</h3>
                            </div>

                            <ul class="detailView apply-as">
                                <!-- 콘텐츠 스타일 변경으로 'apply-as'클래스 명 추가 -->
                                <li class="detailList">
                                    <div class="tit">받는분</div>
                                    <div class="cont"><span class="general"><%=oRs("OrderName")%></span></div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">휴대전화</div>
                                    <div class="cont"><span class="general"><%=oRs("OrderHp")%></span></div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">주소</div>
                                    <div class="cont"><span class="general">[<%=oRs("OrderZipCode")%>] <%=oRs("OrderAddr1")%>&nbsp;<%=oRs("OrderAddr2")%></span></div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">픽업주소</div>
                                    <div class="cont"><span class="general">[<%=oRs("ReceiveZipCode")%>] <%=oRs("ReceiveAddr1")%>&nbsp;<%=oRs("ReceiveAddr2")%></span></div>
                                    <!--<a href="#" class="all-view is-right">변경</a>-->
                                </li>
                            </ul>

                            <div class="h-line">
                                <h3 class="h-level4">신청 내용</h3>
                            </div>

                            <div class="area-radio">
                                <span class="rad-ty1">
									<input type="radio" id="as_1" name="RequestCode" value="A" checked="">
									<label for="as_1">수선</label>
								</span>
                                <span class="rad-ty1">
									<input type="radio" id="as_2" name="RequestCode" value="C">
									<label for="as_2">반품</label>
								</span>
                                <span class="rad-ty1">
									<input type="radio" id="as_3" name="RequestCode" value="R">
									<label for="as_3">교환</label>
								</span>
                            </div>

                            <div class="review-write">
                                <div class="input">
                                    <textarea name="Contents" id="Contents" cols="30" rows="10"></textarea>
                                </div>
                            </div>

                            <div class="buttongroup">
                                <input type="file" name="FileName" id="file_btn" style="display:none;" />
                                <button type="button" class="button-ty2 is-expand ty-bd-gray" onclick="openAsImageSearch();"><span class="icon ico-add-photo">사진 첨부</span></button>
								<input type="text" value="선택된 파일 없음" disabled style="display:none;" />
                            </div>

                            <div class="added-photo">
                                <ul class="as-photo" style="margin-bottom:10px">
                                </ul>
                            </div>

                            <div class="inf-type1">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">A/S 처리는 슈마커에서 구매한 상품만 대상으로 합니다. 처음 상태로의 수선과 기장 및 디자인 변경은 불가합니다.</li>
                                    <li class="bullet-ty1">브랜드에서 통보한 판정 결과에 따라 신청내용과 보상내용이 상이 할 수 있습니다.</li>
                                    <li class="bullet-ty1">아래 고객 정보는 슈마커 개인정보보호 정책에 따라 자동으로 제공 동의 되었습니다.</li>
                                    <li class="bullet-ty1">일부 경우에 따라 유상 수선으로 진행될 수 있으며, 진행 전 유선이나 문자메시지를 통해 견적 사항을 안내 받으실 수 있습니다.</li>
                                </ul>
                            </div>
                        </div>
						</form>
                    </div>

                    <div class="btns">
                        <button type="button" class="button ty-red" onclick="asWrite();">A/S 신청하기</button>
                    </div>
                </div>
            </div>
        </div>
	    <!-- // PopUp -->

		<script>


			var formInit = function(){
				$('.checkbox').each(function (i, el) {
					FormCheckbox.build(el);
					$(el).find('input').on('change', function () {
						FormCheckbox.change(this);
						if ($(this).data('allchk') != undefined) {
							FormCheckbox.allchk(this);
						} else if ($(this).data('allparts') != undefined) {
							FormCheckbox.allparts(this);
						}
					});
					$(el).find('input').on('focus blur click', function () {
						FormCheckbox.focusin(this);
					});
				});
				$('.select').each(function(i, el){
					FormSelect.build(el);
					$(el).find('select').on('change', function(){
						FormSelect.change(this);
					});
					$(el).find('select').on('focus blur click', function(){
						FormSelect.focusin(this);
					});
				});
			};

			$(document).ready(function(){
				formInit();
			});

			$(function () {
				//.upload-file 사진 업로드
				var fileTarget = $('.buttongroup input:file');

				fileTarget.on('change', function () {
					if (window.FileReader) {
						var fileName = $(this)[0].files[0].name;
					} else {
						var fileName = $(this).val().split('/').pop().split('\\').pop();
					}

					$(this).siblings('.file-name').val(fileName);

					asImageAdd();
				});
			})

		</script>
<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>