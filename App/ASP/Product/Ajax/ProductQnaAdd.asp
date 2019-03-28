<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductQnaAdd.asp - 상품 문의 등록
'Date		: 2019.01.10
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

'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

ProductCode				= sqlFilter(Request("ProductCode"))

IF U_Num = "" THEN
		Response.Write "FAIL|||||로그인 정보가 없습니다."
		Response.End
END IF

IF ProductCode = "" THEN
		Response.Write "FAIL|||||상품 정보가 없습니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_By_ProductCode"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 상품 정보 입니다."
		Response.End
END IF
oRs.Close

Response.Write "OK|||||"
%>

        <div class="area-dim"></div>

		<form name="ProductQnaForm" id="ProductQnaForm" method="post">
		<input type="hidden" name="ProductCode" value="<%=ProductCode%>" />
        <div class="area-pop">

            <!-- 팝업 문의하기 -->
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">문의하기</p>
                    <button type="button" class="btn-hide-pop" onclick="closePop('DimDepth1');">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="wrap-question">
                            <div class="inf-type1">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">상품과 관련된 문의만 답변을 받으실 수 있습니다.</li>
                                    <li class="bullet-ty1">해당 게시판에 맞지 않는 질문이나 상업/홍보성 글은 통보없이 삭제 될 수 있습니다.</li>
                                    <li class="bullet-ty1">결제/배송/교환/반품에 대한 문의는 고객센터 1:1문의를 이용해 주시기 바랍니다.</li>
                                </ul>
                            </div>

                            <div class="a-question">
                                <a href="javascript:void(0)" onclick="closePop('DimDepth1');APP_GoUrl('/ASP/Mypage/Qna.asp?QnaType=2')">1:1문의 바로가기</a>
                            </div>

                            <div class="fieldset">
                                <label class="fieldset-label">문의내용 <em class="essential">*</em></label>
                                <div class="fieldset-row">
                                   <span class="select is-expand" style="margin-bottom:13px;">
										<select name="ClassName" id="ClassName">
											<option value="사이즈문의">사이즈문의</option>
											<option value="색상문의">색상문의</option>
											<option value="기타문의">기타문의</option>
										</select>
									   <span class="value">사이즈문의</span>
									</span>
									
                                   <span class="input is-expand">
										<input type="text" name="Title" id="Title" maxlength="100">
									</span>
                                    <span class="input is-expand">
										<textarea name="Contents" id="Contents" rows="7"></textarea>
									</span>
                                </div>
                                <p class="message icon ico-essential">*필수입력 항목</p>
                            </div>

                            <div class="fieldset ty-row">
                                <label class="fieldset-label" for="yes_ans_sns" style="min-width: 35%;">SMS로 답변알림 받기</label>
                                <div class="fieldset-row">
                                    <div class="radiogroup">
                                        <div class="inner" style="width:30%;">
                                            <span class="radio">
												<input type="radio" id="yes_ans_sns" name="SMSReturnFlag" value="1" checked>
											</span>
                                            <label for="yes_ans_sns">예</label>
										</div>
										<div class="inner">
                                            <span class="radio">
												<input type="radio" id="no_ans_sns" name="SMSReturnFlag" value="0">
											</span>
                                            <label for="no_ans_sns">아니요</label>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div class="fieldset ty-col2">
                                <div class="fieldset-row">
                                    <span class="select">
										<select name="Mobile1">
												<option value="010" selected>010</option>
												<option value="011">011</option>
												<option value="016">016</option>
												<option value="017">017</option>
												<option value="018">018</option>
												<option value="019">019</option>
										</select>
										<span class="value">010</span>
                                    </span>									
                                    <span class="input">
										<input type="tel" name="Mobile2" id="Mobile2" maxlength="8" placeholder="휴대폰 번호를 입력해주세요.">
									</span>
                                </div>
                            </div>

                            <div class="fieldset ty-row">
                                <label class="fieldset-label" for="yes_ans_email" style="min-width: 35%;">이메일로 답변알림 받기</label>
                                <div class="fieldset-row">
                                    <div class="radiogroup">
                                        <div class="inner" style="width:30%;">
                                            <span class="radio">
												<input type="radio" id="yes_ans_email" name="EMailReturnFlag" value="1" checked>
											</span>
                                            <label for="yes_ans_email">예</label>
                                        </div>
                                        <div class="inner">
                                            <span class="radio">
												<input type="radio" id="no_ans_email" name="EMailReturnFlag" value="0">
											</span>
                                            <label for="no_ans_email">아니요</label>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div class="fieldset">
                                <div class="fieldset-row">
                                    <span class="input is-expand">
										<input type="email" name="Email" id="Email" maxlength="50" placeholder="이메일 계정을 입력해주세요.">
									</span>
                                </div>
                            </div>

                            <div class="fieldset ty-row">
                                <label class="fieldset-label" style="min-width: 35%;">비밀글 설정</label>
                                <div class="fieldset-row">
                                    <div class="radiogroup">
                                        <div class="inner" style="width:30%;">
                                            <span class="radio">
												<input type="radio" id="yes_secret_txt" name="SecretFlag" value="1" checked>
											</span>
                                            <label for="yes_secret_txt">예</label>
                                        </div>
                                        <div class="inner">
                                            <span class="radio">
												<input type="radio" id="no_secret_txt" name="SecretFlag" value="0">
											</span>
                                            <label for="no_secret_txt">아니요</label>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" class="button ty-red" onclick="ins_ProductQna();">등록하기</button>
                    </div>
                </div>
            </div>
            <!-- // 팝업 문의하기 -->
        </div>
		</form>

		<script>
			var FormSelect = {
				build : function(el){
					$('.value', el).text($('option:selected', el).text());
					if($('select', el).is(':disabled')){
						$(el).addClass('is-disabled');
					}
					if($('select', el).prop('readonly')){
						$(el).addClass('is-readonly');
					}
				},
				change : function(el){
					$(el).parent().find('.value').text($('option:selected', el).text());
				},
				focusin : function(el){
					if($(el).is(':focus')){
						$(el).parent().addClass('is-focus');
					} else {
						$(el).parent().removeClass('is-focus');
					}
				}
			};

			$(document).ready(function(){
				formInit();
			});

		</script>
