<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyInfoModify.asp - 마이페이지 > 회원정보 수정
'Date		: 2018.12.18
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
PageCode2 = "05"
PageCode3 = "03"
PageCode4 = "02"
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
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절


DIM Name
DIM Birth
DIM Sex
DIM HP
DIM HP1
DIM HP2
DIM HP3
DIM Email
DIM SmsFlag
DIM EmailFlag
DIM ZipCode
DIM Address1
DIM Address2
DIM FTFlag

DIM ParentSDupInfo
DIM ParentName
DIM ParentBirth
DIM ParentHP
DIM PHP1
DIM PHP2
DIM PHP3
DIM ParentEmail



DIM arrHP1
arrHP1	= ARRAY("010", "011", "016", "017", "018", "019")

'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


'/****************************************************************************************/
'회원 기본정보 SELECT START
'-----------------------------------------------------------------------------------------------------------'
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Member_Select_By_UserID"
		.Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 30, U_ID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF oRs.EOF THEN
	oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	Response.Write "FAIL|||||"
	Response.End
ELSE
	Name		= sqlFilter(oRs("Name"))
	Birth		= sqlFilter(oRs("Birth"))
	Sex			= sqlFilter(oRs("Sex"))
	HP			= sqlFilter(oRs("HP"))
	IF HP <> "" THEN
			IF UBound(SPLIT(HP,"-")) = 2 THEN
					HP1		= SPLIT(HP, "-")(0)
					HP2		= SPLIT(HP, "-")(1)
					HP3		= SPLIT(HP, "-")(2)
			ELSEIF UBound(SPLIT(HP,"-")) = 1 THEN
					HP1		= SPLIT(HP, "-")(0)
					HP2		= SPLIT(HP, "-")(1)
					HP3		= ""
			ELSEIF UBound(SPLIT(HP,"-")) = 0 THEN
					HP1		= SPLIT(HP, "-")(0)
					HP2		= ""
					HP3		= ""
			ELSE
					HP1		= HP
					HP2		= ""
					HP3		= ""
			END IF
	END IF

	Email		= sqlFilter(oRs("Email"))
	SmsFlag		= sqlFilter(oRs("SmsFlag"))
	EmailFlag	= sqlFilter(oRs("EmailFlag"))
	ZipCode		= sqlFilter(oRs("ZipCode"))
	Address1	= sqlFilter(oRs("Address1"))
	Address2	= sqlFilter(oRs("Address2"))
	FTFlag		= sqlFilter(oRs("FTFlag"))

	ParentName	= sqlFilter(oRs("ParentName"))
	ParentBirth = sqlFilter(oRs("ParentBirth"))
	ParentHP	= sqlFilter(oRs("ParentHP"))
	IF ParentHP <> "" THEN
			IF UBound(SPLIT(ParentHP,"-")) = 2 THEN
					PHP1	= SPLIT(ParentHP, "-")(0)
					PHP2	= SPLIT(ParentHP, "-")(1)
					PHP3	= SPLIT(ParentHP, "-")(2)
			ELSEIF UBound(SPLIT(ParentHP,"-")) = 1 THEN
					PHP1	= SPLIT(ParentHP, "-")(0)
					PHP2	= SPLIT(ParentHP, "-")(1)
					PHP3	= ""
			ELSEIF UBound(SPLIT(ParentHP,"-")) = 0 THEN
					PHP1	= SPLIT(ParentHP, "-")(0)
					PHP2	= ""
					PHP3	= ""
			ELSE
					PHP1	= HP
					PHP2	= ""
					PHP3	= ""
			END IF
	END IF
	ParentEmail = sqlFilter(oRs("ParentEmail"))


END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'회원 기본정보 SELECT END
'-----------------------------------------------------------------------------------------------------------'
Response.Write "OK|||||"
%>

							<form name="MyInfoModify" id="MyInfoModify">
							<input type="hidden" name="FTFlag" id="FTFlag" value="<%=FTFlag%>" />
                            <!-- 나의 정보 수정 -->
                            <div class="h-line">
                                <h2 class="h-level4">나의 정보 수정</h2>
                            </div>
                            <div class="edit-info">
                                <p class="tit no-border">가입자 정보</p>
                                <fieldset>
                                    <legend class="hidden">기본 정보 입력</legend>
                                    <div class="fieldset">
                                        <label for="join-id" class="fieldset-label">아이디</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="text" name="UserID" id="UserID" maxlength="25" value="<%=U_ID %>" readonly="readonly" />
											</span>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <label for="join-pw" class="fieldset-label">비밀번호</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="password" id="Pwd" name="Pwd" placeholder="비밀번호 (8~16자리 영문 대소문자/숫자 포함)">
											</span>
                                        </div>
                                        <p class="message icon ico-caution">8~16자리 영문 대소문자/숫자 포함하여 입력하세요.</p>
                                    </div>
                                    <div class="fieldset">
                                        <label for="join-pw-confirm" class="fieldset-label">비밀번호 확인</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="password" id="Pwd1" name="Pwd1" placeholder="비밀번호를 한번 더 입력해주세요.">
											</span>
                                        </div>
                                    </div>
                                </fieldset>
                                <fieldset class="">
                                    <legend class="hidden">부가 정보 입력</legend>
                                    <div class="fieldset">
                                        <label for="join-name" class="fieldset-label">이름</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="text" name="Name" id="login-name" maxlength="25" placeholder="본인의 실명을 입력해 주세요." value="<%=Name%>" <%IF FTFlag<>"Y" THEN Response.Write "readonly='readonly'" %> />
											</span>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <label for="join-birth" class="fieldset-label">생년월일</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="text" name="Birth" id="login-birth2" maxlength="8" placeholder="EX 19800809" value="<%=Birth%>" <%IF FTFlag<>"Y" THEN Response.Write "readonly='readonly'" %> />
											</span>
                                        </div>
                                    </div>
                                    <div class="fieldset ty-row">
                                        <label class="fieldset-label">성별</label>
                                        <div class="fieldset-row">
                                            <div class="radiogroup">
												<div class="inner">
													<span class="radio">
														<input type="radio" id="join-male" name="Sex" value="M" <%IF SEX="M" THEN Response.write "checked=''" %>>
													</span>
													<label for="join-male">남</label>
												</div>
												<div class="inner">
													<span class="radio">
														<input type="radio" id="join-female" name="Sex" value="F" <%IF SEX="F" THEN Response.write "checked=''" %>>
													</span>
													<label for="join-female">여</label>
												</div>
                                            </div>
                                        </div>
                                    </div>
                                </fieldset>
                                <fieldset class="no-border">
                                    <div class="fieldset ty-col2">
                                        <label for="join-phone" class="fieldset-label">휴대폰 번호</label>
                                        <div class="fieldset-row mb8">
                                            <span class="select">
												<select name="HP1" title="휴대폰 국번 선택">
													<%FOR i = 0 TO UBOUND(arrHP1)%>
													<option value="<%=arrHP1(i)%>"<%IF arrHP1(i) = HP1 THEN%> selected="selected"<%END IF%>><%=arrHP1(i)%></option>
													<%NEXT%>
												</select>
												<span class="value"></span>
                                            </span>
                                            <span class="input">
												<input type="text" id="HP23" name="HP23" placeholder="휴대폰 뒷자리를 입력하세요." value="<%=HP2 & HP3 %>" maxlength="8" />
											</span>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <label for="join-email" class="fieldset-label">이메일주소</label>
                                        <div class="fieldset-row mb8">
                                            <span class="input is-expand">
												<input type="text" id="Email" name="Email" placeholder="이메일계정" value="<%=Email %>" maxlength="50">
											</span>
                                        </div>
                                    </div>
                                </fieldset>
								<%IF FTFlag="Y" THEN%>
                                <div class="fieldset parent">
                                    <p class="tit">보호자 정보</p>
                                    <p class="message">*만 14세 미만 가입 시 필수 기재사항입니다.</p>
                                </div>
                                <fieldset class="">
                                    <legend class="hidden">기본 정보 입력</legend>
                                    <div class="fieldset">
                                        <label for="join-id" class="fieldset-label">이름</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="text" id="ParentName" name="ParentName" placeholder="보호자의 실명을 입력해 주세요." value="<%=ParentName %>" readonly='readonly'>
											</span>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <label for="join-pw" class="fieldset-label">생년월일</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="text" name="ParentBirth" id="ParentBirth" maxlength="8" placeholder="EX 19800809" value="<%=ParentBirth%>" readonly='readonly' %> />
											</span>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <legend class="hidden">인증 정보 입력</legend>
                                        <div class="fieldset ty-col2 pt0">
                                            <label for="join-phone" class="fieldset-label">휴대폰 번호</label>
                                            <div class="fieldset-row">
                                                <span class="select">
													<select name="PHP1" title="휴대폰 국번 선택">
														<%FOR i = 0 TO UBOUND(arrHP1)%>
														<option value="<%=arrHP1(i)%>"<%IF arrHP1(i) = PHP1 THEN%> selected="selected"<%END IF%>><%=arrHP1(i)%></option>
														<%NEXT%>
													</select>
													<span class="value"></span>
                                                </span>
                                                <span class="input">
													<input type="text" name="PHP23" id="PHP23" title="연락처의 앞 번호와 뒷 번호 입력" style="width: 205px" value="<%=PHP2 & PHP3 %>" maxlength="8">
												</span>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <label for="join-email" class="fieldset-label">이메일주소</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="text" name="ParentEmail" id="ParentEmail" placeholder="보호자 이메일계정" value="<%=ParentEmail %>" maxlength="50">
											</span>
                                        </div>
                                    </div>
                                </fieldset>
								<%END IF%>
                            </div>
                            
							<!-- 고객서비스 수신동의 -->
                            <div class="h-line">
                                <h2 class="h-level4">고객서비스 수신동의</h2>
                                <span class="h-date  color-dg">이벤트,세일,쿠폰지급 정보등 수신</span>
                            </div>
                            <div class="agree-receive">
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">이메일 수신동의</label>
                                    <div class="fieldset-row">
                                        <div class="radiogroup">
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" id="mail-agree" name="EmailFlag" value="Y" <%IF EmailFlag="Y" THEN Response.write "checked=''"%> />
												</span>
                                                <label for="mail-agree">수신함</label>
                                            </div>
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" id="mail-disagree" name="EmailFlag" value="N" <%IF EmailFlag<>"Y" THEN Response.write "checked=''"%> />
												</span>
                                                <label for="mail-disagree">수신안함</label>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="fieldset ty-row">
                                    <label class="fieldset-label">문자 수신동의</label>
                                    <div class="fieldset-row">
                                        <div class="radiogroup">
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" id="sms-agree"  name="SmsFlag" value="Y" <%IF SmsFlag="Y" THEN Response.write "checked=''"%>>
												</span>
                                                <label for="sms-agree">수신함</label>
                                            </div>
                                            <div class="inner">
                                                <span class="radio">
													<input type="radio" id="sms-disagree"  name="SmsFlag" value="N" <%IF SmsFlag<>"Y" THEN Response.write "checked=''"%>>
												</span>
                                                <label for="sms-disagree">수신안함</label>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            
                            <!-- 수정완료/취소 -->
                            <div class="edit-complete">
                                <div class="buttongroup is-space">
                                    <button type="button" onclick="chk_MyInfoModify();" class="button-ty2 is-expand ty-red">수정 완료</button>
                                    <button type="button" onclick="reset();" class="button-ty2 is-expand ty-black">취소</button>
                                </div>
                                <button type="button" onclick="common_getPage('getMyInfoModify','Withdraw');" class="button-ty2 is-expand ty-bd-gray">회원 탈퇴</button>
                            </div>
							</form>

							<script>
								//ajax 이용시 라디오버튼 disabled 처리되는 문제로 추가 (2018.12.18 DJ)							
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
							</script>
    <!-- PopUp -->
	<div id="DimDepth1" class="area-dim" style="display:none"></div>
    <!-- // PopUp -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>