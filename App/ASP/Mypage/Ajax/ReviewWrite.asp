<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ReviewWrite.asp - 구매후기 작성 폼 페이지
'Date		: 2018.12.24
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


DIM OrderCode
DIM Idx

DIM OPIdx_Org
DIM ProductCode
DIM ProductName
DIM SizeCD
DIM BrandName
DIM ProductImage
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
Idx					= sqlFilter(Request("Idx"))



IF OrderCode = "" OR Idx = "" THEN
		Response.Write "FAIL|||||선택한 주문정보가 없습니다."
		Response.End
END IF



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'-----------------------------------------------------------------------------------------------------------'
'# 주문정보 Start
'-----------------------------------------------------------------------------------------------------------'
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.Idx = " & Idx & " "
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND B.UserID = '" & U_NUM & "' "
ELSE
		wQuery = wQuery & "AND B.OrderName = '" & N_NAME & "' AND B.OrderHp = '" & N_HP & "' AND B.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = ""

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

IF NOT oRs.EOF THEN
		OPIdx_Org		= oRs("OPIdx_Org")
		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		SizeCD			= oRs("SizeCD")
		BrandName		= oRs("BrandName")

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
'# 주문정보 End
'-----------------------------------------------------------------------------------------------------------'
Response.Write "OK|||||"
%>					

		<div class="area-dim"></div>

		<form name="ReviewWriteForm" method="post" enctype="multipart/form-data">
		<input type="hidden" name="OrderCode"			value="<%=OrderCode%>" />
		<input type="hidden" name="Idx"					value="<%=Idx%>" />
		<input type="hidden" name="OPIdx_Org"			value="<%=OPIdx_Org%>" />
		<input type="hidden" name="ProductCode"			value="<%=ProductCode%>" />
		<input type="hidden" name="UploadFiles"			value="" />
		<input type="hidden" name="UploadFilesCount"	value="0" />

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">상품 후기 작성하기</div>
                    <div class="btn-hide-pop" onclick="closePop('DimDepth1');">닫기</div>
                </div>

                <div class="container-pop mypage-ty2">
                    <!-- 팝업 스타일 변경으로 'mypage-ty2'클래스 명 추가 -->
                    <div class="contents">
                        <div class="wrap-review">
                            <div class="informView">
                                <div class="informItem">
                                    <a href="#">
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
                                <h3 class="h-level4">평가 하기</h3>
                            </div>

                            <!-- psd_181212수정 -->
                            <ul class="assessment">
                                <li class="star-score">
                                    <span class="tit">사이즈</span>
                                    <div class="post-body">
                                        <div class="star-grade">
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                        </div>
                                    </div>
                                    <div class="score">0</div>
									<input type="hidden" name="SizeGrade" value="0" />
                                </li>
                                <li class="star-score">
                                    <span class="tit">착화감</span>
                                    <div class="post-body">
                                        <div class="star-grade">
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                        </div>
                                    </div>
                                    <div class="score">0</div>
									<input type="hidden" name="WearGrade" value="0" />
                                </li>
                                <li class="star-score">
                                    <span class="tit">디자인</span>
                                    <div class="post-body">
                                        <div class="star-grade">
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                        </div>
                                    </div>
                                    <div class="score">0</div>
									<input type="hidden" name="DesignGrade" value="0" />
                                </li>
                                <li class="star-score">
                                    <span class="tit">품질</span>
                                    <div class="post-body">
                                        <div class="star-grade">
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                            <span class="star"></span>
                                        </div>
                                    </div>
                                    <div class="score">0</div>
									<input type="hidden" name="QualityGrade" value="0" />
                                </li>
                            </ul>
                            <!-- psd_181212수정 -->

                            <div class="h-line">
                                <h3 class="h-level4">후기 작성</h3>
                            </div>

                            <div class="review-write">
                                <div class="input">
                                    <textarea name="Contents" id="Contents" cols="30" rows="10"></textarea>
                                </div>
                            </div>

                            <div class="buttongroup">
                                <input type="file" name="FileName" id="file_btn" style="display:none;" />
                                <button type="button" class="button-ty2 is-expand ty-bd-gray" onclick="openReviewImageSearch();"><span class="icon ico-add-photo">사진 첨부</span></button>
								<input type="text" value="선택된 파일 없음" disabled style="display:none;" />
                            </div>

                            <div class="added-photo">
                                <ul class="review-photo" style="margin-bottom:10px">
                                </ul>
                            </div>

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
                        <button type="button" class="button ty-red" onclick="reviewWrite();">등록하기</button>
                    </div>
                </div>
            </div>
	    </div>
		</form>



		<script type="text/javascript">
			$(function () {
				$('.star-grade').each(function() {
					var _this = $(this);
					var _thisSpan = _this.children('span');
					_thisSpan.click(function () {
						var _spanIndex = $(this).index();
						var _starNum = _spanIndex + 1;
						$(this).closest('.star-score').children('.score').text(_starNum);
						$(this).closest('.star-score').children('input').val(_starNum);

						$(this).parent().children('span').removeClass('on');
						$(this).addClass('on').prevAll('span').addClass('on');
						return false;
					});
				});
				//.upload-file 사진 업로드
				var fileTarget = $('.buttongroup input:file');

				fileTarget.on('change', function () {
					if (window.FileReader) {
						var fileName = $(this)[0].files[0].name;
					} else {
						var fileName = $(this).val().split('/').pop().split('\\').pop();
					}

					$(this).siblings('.file-name').val(fileName);

					reviewImageAdd();
				});
			})
		</script>


<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>