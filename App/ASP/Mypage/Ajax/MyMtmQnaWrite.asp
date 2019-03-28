<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyQnaList.asp - 마이페이지 > 상품문의 > 1:1문의 작성폼
'Date		: 2018.12.26
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
PageCode1 = "05"
PageCode2 = "03"
PageCode3 = "05"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

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

Response.Write "OK|||||"
%>
    <!-- PopUp -->
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">1:1 문의하기</div>
                    <button class="btn-hide-pop" onclick="common_PopClose('DimDepth1');">닫기</button>
                </div>

                <div class="container-pop mypage-ty2">
                    <!-- 팝업 스타일 변경으로 'mypage-ty2'클래스 명 추가 -->
                    <div class="contents">
                        <div class="wrap-mtom">
							<form name="MtmQnaWriteForm" id="MtmQnaWriteForm" method="post" autocomplete="off">
							<input type="hidden" name="UploadFiles"			value="" />
							<input type="hidden" name="UploadFilesCount"	value="0" />
                            <div class="reason">
                                <span class="select">
									<select id="selectBank1" name="category1" onchange="selCate2();" title="문의 유형 선택">
										<option value="">문의 유형을 선택하세요.</option>
										<option value="온라인">온라인</option>
										<option value="오프라인">오프라인</option>
										<option value="기타">기타</option>
									</select>
                                    <span class="value" id="selval1">문의 유형을 선택하세요.</span>
                                </span>
                                <span class="select" style="display:none;" id="span_cate2">
									<select id="selectBank2" name="category2" onchange="selCate3();">
										<option value=''>온라인 상담유형을 선택하세요.</option>
									</select>
                                    <span class="value" id="selval2">온라인 상담유형을 선택하세요.</span>
                                </span>
                                <span class="select" style="display:none;" id="span_cate3">
									<select id="selectBank3" name="category3" onchange="selCate4();">
										<option value=''>상세 유형을 선택하세요.</option>
									</select>
                                    <span class="value"  id="selval3">상세 유형을 선택하세요.</span>
                                </span>
                                <span class="select" style="display:none;" id="span_cate4">
									<select id="selectBank4" name="category4">
										<option value=''>구분 유형을 선택하세요.</option>
									</select>
                                    <span class="value"  id="selval4">구분 유형을 선택하세요.</span>
                                </span>

                                <span class="input">
                                    <input type="text" id="Title" name="Title" placeholder="제목을 입력해 주세요.">
                                </span>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">문의 내용 작성</h3>
                            </div>

                            <div class="review-write">
                                <div class="input">
                                    <textarea name="Contents" id="Contents" cols="30" rows="10" placeholder="내용을 입력 해 주세요."></textarea>
                                </div>
                            </div>
                            <!-- psd_181212수정 -->
                            <div class="answer-agree">
                                <div class="answer-sms">
                                    <div class="fieldset ty-row">
                                        <label class="fieldset-label" for="yes_ans_sns">답변 문자 받기</label>
                                        <div class="fieldset-row">
                                            <div class="radiogroup">
                                                <div class="inner">
                                                    <span class="radio is-checked">
                                                        <input type="radio" id="yes_ans_sns" name="SMSReturnFlag" value="1" checked="">
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
													<option value="010">010</option>
													<option value="011">011</option>
													<option value="016">016</option>
													<option value="017">017</option>
													<option value="019">019</option>
												</select>
                                                <span class="value">010</span>
                                            </span>
                                            <span class="input">
                                                <input type="tel" name="Mobile23" maxlength="8" placeholder="휴대폰 번호를 입력해주세요.">
                                            </span>
                                        </div>
                                    </div>
                                </div>
                                <div class="answer-mail">
                                    <div class="fieldset ty-row">
                                        <label class="fieldset-label" for="yes_ans_email">답변 이메일 받기</label>
                                        <div class="fieldset-row">
                                            <div class="radiogroup">
                                                <div class="inner">
                                                    <span class="radio">
                                                        <input type="radio" id="yes_ans_email" name="EMailReturnFlag" value="1" checked="">
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
                                                <input type="email" name="EMail" id="EMail" maxlength="50" placeholder="이메일 주소를 입력해주세요.">
                                            </span>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div class="buttongroup">
                                <input type="file" name="FileName" id="file_btn" style="display:none;" />
                                <button type="button" class="button-ty2 is-expand ty-bd-gray" onclick="openMtmQnaImageSearch();"><span class="icon ico-add-photo">사진 첨부</span></button>
								<input type="text" value="선택된 파일 없음" disabled style="display:none;" />
                            </div>

                            <div class="added-photo">
                                <ul class="review-photo" style="margin-bottom:10px">
                                </ul>
                            </div>

						</form>
                            <!-- //psd_181212수정 -->
                        </div>
                    </div>

                    <div class="btns">
                        <button type="button" onclick="chk_MtmWrite();" class="button ty-red">1:1 문의하기</button>
                    </div>
                </div>
            </div>
        </div>
    <!-- // PopUp -->

		<script>
			var FormRadio = {
				build : function(el){
					if($(el).find('input').is(':disabled')){
						$(el).addClass('is-disabled');
					}
					if($(el).find('input').prop('readonly')){
						$(el).addClass('is-readonly');
					}
					if($(el).find('input').is(':checked')){
						$(el).addClass('is-checked');
					}
				},
				change : function(el){
					var groupName = $(el).attr('name');
					$('[name=' + groupName + ']').parent().removeClass('is-checked');
					$('[name=' + groupName + ']:checked').parent().addClass('is-checked');
				},
				focusin : function(el){
					if($(el).is(':focus')){
						$(el).parent().addClass('is-focus');
					} else {
						$(el).parent().removeClass('is-focus');
					}
				}
			};
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

			var formInit = function(){
				$('.radio, .radio2').each(function(i, el){
					FormRadio.build(el);
					$(el).find('input').on('change', function(){
						FormRadio.change(this);
					});
					$(el).find('input').on('focus blur click', function(){
						FormRadio.focusin(this);
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

				if ($('#remaintime').get(0) != undefined) {
					setInterval(timedeal, 1);
				}
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

					mtmQnaImageAdd();
				});
			});
		</script>
<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>