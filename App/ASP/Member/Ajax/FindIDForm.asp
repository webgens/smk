                <div id="tabs" class="tab ty-theme1">
                    <ul class="tab-selector">
                        <li class="part-3"><a href="javascript:;" data-target="tabs-col1">등록정보</a></li>
                        <li class="part-3"><a href="javascript:;" data-target="tabs-col2">휴대폰 인증</a></li>
                        <li class="part-3"><a href="javascript:;" data-target="tabs-col3">아이핀 인증</a></li>
                    </ul>
                    <div id="tabs-col1" class="tab-panel">
                        <ul class="id-find-info">
                            <li class="bullet-ty1">슈마커 회원정보에 등록되어 있는 정보 중 1가지를 택하여 입력해 주세요. </li>
                            <li class="bullet-ty1">등록된 정보로 아이디의 일부를 찾을 수 있습니다.</li>
                            <li class="bullet-ty1 red">아이디 전체보기를 원하시면 휴대폰 인증이나 아이핀 인증을 선택해주세요.</li>
                        </ul>
                        <div class="radiogroup">
                            <div class="inner">
                                <span class="radio">
                                    <input type="radio" id="find-phone" name="FindIDType" value="mobile" checked onclick="chg_FindID_Normal_Type()">
                                </span>
                                <label for="find-phone">휴대폰 번호로 찾기</label>
                            </div>
                            <div class="inner">
                                <span class="radio">
									<input type="radio" id="find-mail" name="FindIDType" value="email" onclick="chg_FindID_Normal_Type()">
                                </span>
                                <label for="find-mail">이메일로 찾기</label>
                            </div>
                        </div>


						<form name="formFindID" id="formFindID" method="post" autocomplete="off">
						<input type="hidden" name="FindIDType" value="mobile">
						<div id="FI_N_Rst" class="t-type4-inform-hp" style="display:none">
							<span id="FI_N_Rst_Msg" class="nt-type-bl-r"></span>
						</div>
                        <fieldset id="FI_N_Form">
                            <legend class="hidden">휴대폰/이메일 정보 입력</legend>
                            <div class="fieldset">
                                <label for="user-name" class="fieldset-label">이름</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="Name" id="Name" maxlength="25" placeholder="본인의 실명을 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset ty-col2" id="FI_N_Mobile">
                                <label for="user-phone2" class="fieldset-label">휴대폰</label>
                                <div class="fieldset-row">
                                    <span class="select2">
                                        <select name="HP1" id="HP1" title="휴대폰 국번 선택">
											<option value="010">010</option>
											<option value="011">011</option>
											<option value="016">016</option>
											<option value="017">017</option>
											<option value="018">018</option>
											<option value="019">019</option>
                                        </select>
                                        <span class="value"></span>
                                    </span>
									<span class="dash1">-</span>
									<span class="input2">
										<input type="tel" id="HP2" name="HP2" maxlength="4">
									</span>
									<span class="dash2">-</span>
									<span class="input3">
										<input type="tel" id="HP3" name="HP3" maxlength="4">
									</span>
                                </div>
                            </div>
                            <div class="fieldset" id="FI_N_Email" style="display:none">
                                <label for="user-phone2" class="fieldset-label">이메일</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="email" name="Email" id="Email" maxlength="50" placeholder="이메일을 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                        </fieldset>
						<div id="FI_N_Btn" class="area-btn-c">
							<button type="button" onclick="chk_FindID_Normal()" class="button is-expand ty-red">확 인</button>
                        </div>
                        <div id="FI_N_Fail_Btn" class="area-btn-c" style="display:none">
                            <button type="button" onclick="chg_FindID_Normal_Type()" class="button is-expand ty-red">다시찾기</button>
                        </div>
                        <div id="FI_N_Succ_Btn" class="area-btn-c" style="display:none">
                            <button type="button" onclick="APP_PopupHistoryBack()" class="button ty-red login">로그인하기</button>
                            <button type="button" onclick="chg_FindIDPwForm(1)" class="button ty-black pw">비밀번호찾기</button>
                        </div>
						</form>
                    </div>
                    <div id="tabs-col2" class="tab-panel">
						<form name="form" id="form">
							<input type="hidden" name="SMode" value="SearchID" />
							<ul class="id-find-info  confirm-info no-border">
								<li class="bullet-ty1">본인 명의 휴대폰 번호로 가입여부 및 본인여부를 확인합니다. </li>
								<li class="bullet-ty1">타인명의/법인휴대폰은 본인인증이 불가능합니다.</li>
							</ul>
							<div id="FI_HP_Rst" class="t-type4-inform-hp" style="display:none">
								<span id="FI_HP_Rst_Msg" class="nt-type-bl-r"></span>
							</div>
							<div id="FI_HP_Btn" class="area-btn-c">
								<button type="button" onclick="auth_HP('form')" class="button is-expand ty-red">휴대폰 인증하기</button>
							</div>
							<div id="FI_HP_Fail_Btn" class="area-btn-c" style="display:none">
								<button type="button" onclick="re_FindID_AuthHP()" class="button is-expand ty-red">다시찾기</button>
							</div>
							<div id="FI_HP_Succ_Btn" class="area-btn-c" style="display:none">
								<button type="button" onclick="APP_PopupHistoryBack()" class="button ty-red login">로그인하기</button>
								<button type="button" onclick="chg_FindIDPwForm(1)" class="button ty-black pw">비밀번호찾기</button>
							</div>
						</form>
                    </div>
                    <div id="tabs-col3" class="tab-panel">
						<form name="form1" id="form1">
							<input type="hidden" name="SMode" value="SearchID" />
							<ul class="id-find-info  confirm-info no-border">
								<li class="bullet-ty1">아이핀으로 가입하신 경우,<br> 아이핀 인증을 통해 아이디를 찾을 수 있습니다. </li>
							</ul>
							<div id="FI_Ipin_Rst" class="t-type4-inform-hp" style="display:none">
								<span id="FI_Ipin_Rst_Msg" class="nt-type-bl-r"></span>
							</div>
							<div id="FI_Ipin_Btn" class="area-btn-c">
								<button type="button" onclick="auth_Ipin('form1')" class="button is-expand ty-red">아이핀 인증하기</button>
							</div>
							<div id="FI_Ipin_Fail_Btn" class="area-btn-c" style="display:none">
								<button type="button" onclick="re_FindID_AuthIpin()" class="button is-expand ty-red">다시찾기</button>
							</div>
							<div id="FI_Ipin_Succ_Btn" class="area-btn-c" style="display:none">
								<button type="button" onclick="APP_PopupHistoryBack()" class="button ty-red login">로그인하기</button>
								<button type="button" onclick="chg_FindIDPwForm(1)" class="button ty-black pw">비밀번호찾기</button>
							</div>
						</form>
                    </div>
                </div>



				<script type="text/javascript">
					$(function() {
						$(".select2 .value").text($("select[name='HP1']").val());

						$(".select2 select").on("focus", function() {
							$(".select2").addClass("is-focus");
						});
						$(".select2 select").on("blur", function() {
							$(".select2").removeClass("is-focus");
						});
						$(".select2 select").on("change", function() {
							$(".select2 .value").text($("select[name='HP1']").val());
						});
		
						$(".input2 input").on("focus", function() {
							$(".input2").addClass("is-focus");
						});
						$(".input2 input").on("blur", function() {
							$(".input2").removeClass("is-focus");
						});
		
						$(".input3 input").on("focus", function() {
							$(".input3").addClass("is-focus");
						});
						$(".input3 input").on("blur", function() {
							$(".input3").removeClass("is-focus");
						});
		
						$(".input1 input").on("focus", function() {
							$(".input1").addClass("is-focus");
						});
						$(".input1 input").on("blur", function() {
							$(".input1").removeClass("is-focus");
						});

						tabBuild();
						formInit();
					});
				</script>