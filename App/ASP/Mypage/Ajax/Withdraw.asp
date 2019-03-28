<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Withdraw.asp - 마이페이지 > 회원탈퇴
'Date		: 2018.12.19
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
PageCode4 = "03"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM i
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

'/****************************************************************************************/
'탈퇴사유 SELECT START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
	.ActiveConnection = oConn
	.CommandType = adCmdStoredProc
	.CommandText = "USP_Front_EShop_Member_Withdraw_Reason_Select_For_SelectBox"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF oRs.EOF THEN
	oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	Response.Write "|||||탈퇴 사유정보가 없습니다."
	Response.End
ELSE
	Response.Write "OK|||||"
END IF
%>

							<form name="withDrawForm" id="withDrawForm" method="post" autocomplete="off">
                            <!-- 회원탈퇴 -->
                            <div class="h-line">
                                <h2 class="h-level4">회원탈퇴 신청</h2>
                            </div>
                            <div class="draw-info">
                                <p class="tit no-border"></p>
                                <fieldset>
                                    <legend class="hidden">회원탈퇴 신청</legend>
                                    <div class="fieldset">
                                        <label for="join-id" class="fieldset-label"> 하시기 전 아래 내용을 확인 부탁드립니다.</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand" style="padding:5px;">
												• 진행 중인 주문건이 있는 경우 배송확인이 끝난 후에 다시 탈퇴 신청을 부탁드립니다.<br />
												• 탈퇴 즉시 그동안 보유하셨던 멤버십 등급, 거래정보, 쿠폰등의 정보가 삭제됩니다.<br />
												• 회원 탈퇴 후에는 그 어떤 경우에도 철회/취소가 불가능합니다.<br />
												• 동일한 아이디로 재 가입이 불가능합니다.<br />
												• 회원 탈퇴 후 1개월동안(셀러의 경우 2개월) 회원의 성명, 아이디, 이메일(E-mail), 연락처 정보를 보관 하며, 로그기록, 접속아이피(IP)정보는 12개월간 보관합니다.<br />
												• 거래 정보가 있는 경우, 판매 거래 정보관리를 위하여 구매와 관련된 상품정보, 아이디, 거래 내역 등에 대한 기본 정보는 탈퇴 후 5년간 보관합니다.<br />
												• 포인트의 경우 잔액을 모두 인출한 이후 회원 탈퇴가 가능합니다.<br />
												  (포인트 인출관련 문의는 고객센터 상담시간을 이용하여 유선 확인 후 진행됩니다.)
											</span>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <label for="join-id" class="fieldset-label">탈퇴사유 선택</label>
										<div class="fieldset-row">
											<div class="radiogroup">
												<%
													i=1
													DO UNTIL oRs.EOF 
												%>
												<div class="inner">
													<span class="radio">
														<input type="radio" id="draw<%=i%>" name="wdReason" value="<%=oRs("WReason")%>">
													</span>
													<label for="draw<%=i%>"><%=oRs("WReasonDesc")%></label>
												</div>
												<%
														i = i+1
														oRs.MoveNext
													LOOP
												%>
											</div>
										</div>
									</div>
                                    <div class="fieldset">
                                        <label for="join-pw" class="fieldset-label">비밀번호</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
												<input type="password" id="Pwd" name="Pwd" placeholder="비밀번호를 입력">
											</span>
                                        </div>
                                        <p class="message icon ico-caution">비밀번호를 입력하여 주십시오.</p>
                                    </div>
                                </fieldset>

                            </div>


                          

                            
                            <!-- 수정완료/취소 -->
                            <div class="edit-complete">
                                <div class="buttongroup is-space">
                                    <button type="button" onclick="chk_MyDraw();" class="button-ty2 is-expand ty-red">탈퇴완료</button>
                                    <button type="button" onclick="location.replace('/ASP/Mypage/MyInfoModify.asp');" class="button-ty2 is-expand ty-black">취소</button>
                                </div>
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
								};


								$(document).ready(function(){
									formInit();

									if ($('#remaintime').get(0) != undefined) {
										setInterval(timedeal, 1);
									}
								});
							</script>

<%
oRs.Close
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>