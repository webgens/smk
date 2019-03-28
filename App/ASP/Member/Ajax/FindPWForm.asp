                <div id="tabs" class="tab ty-theme1">
                    <ul class="tab-selector">
                        <li class="part-3"><a href="javascript:intTabPanel1();" data-target="tabs-col1">등록정보 인증</a></li>
                        <li class="part-3"><a href="javascript:intTabPanel1();" data-target="tabs-col2">휴대폰 인증</a></li>
                        <li class="part-3"><a href="javascript:intTabPanel1();" data-target="tabs-col3">아이핀 인증</a></li>
                    </ul>
                    <div id="tabs-col1" class="tab-panel">
                        <ul class="id-find-info pw-find-info">
                            <li class="bullet-ty1">슈마커 회원정보에 등록되어있는 정보 중 1가지를 택하여 입력해 주세요.</li>
                            <li class="bullet-ty1">등록정보로 비밀번호를 재설정 할 수 있습니다.</li>
                        </ul>
                        <div class="radiogroup">
                            <div class="inner">
                                <span class="radio">
                                    <input type="radio" id="find-phone" name="FindPWType" value="mobile" checked onclick="chg_FindPW_Normal_Type()">
                                </span>
                                <label for="find-phone">휴대폰 번호로 찾기</label>
                            </div>
                            <div class="inner">
                                <span class="radio">
									<input type="radio" id="find-mail" name="FindPWType" value="email" onclick="chg_FindPW_Normal_Type()">
                                </span>
                                <label for="find-mail">이메일로 찾기</label>
                            </div>
                        </div>

						<form name="formFindPW" id="formFindPW" method="post" autocomplete="off">
						<input type="hidden" name="FindPWType" value="mobile">
						<div id="FW_N_Rst" class="t-type4-inform-hp" style="display:none">
							<span id="FW_N_Rst_Msg" class="nt-type-bl-r"></span>
						</div>
                        <fieldset id="FW_N_Form">
                            <legend class="hidden">이메일 정보 입력</legend>
                            <div class="fieldset">
                                <label for="user-name" class="fieldset-label">이름</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="Name" id="Name1" maxlength="25" placeholder="본인의 실명을 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="user-id" class="fieldset-label">아이디</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="UserID" id="UserID1" maxlength="30" placeholder="아이디를 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                            <div id="FW_N_Mobile" class="fieldset ty-col2">
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
										<input type="tel" name="HP2" id="HP2" maxlength="4">
									</span>
									<span class="dash2">-</span>
									<span class="input3">
										<input type="tel" name="HP3" id="HP3" maxlength="4">
									</span>
                                </div>
                            </div>
                            <div id="FW_N_Email" style="display:none" class="fieldset">
                                <label for="user-email" class="fieldset-label">이메일</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="email" name="Email" id="Email1" maxlength="50" placeholder="가입 시 등록하신 이메일주소를 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                        </fieldset>
                        <div id="FW_N_Btn" class="area-btn-c">
                            <button type="button" onclick="chk_FindPW_Normal()" class="button is-expand ty-red">확 인</button>
                        </div>
                        <div id="FW_N_Fail_Btn" class="area-btn-c" style="display:none">
                            <button type="button" onclick="chg_FindPW_Normal_Type()" class="button is-expand ty-red">다시찾기</button>
                        </div>
						</form>
                    </div>
                    <div id="tabs-col2" class="tab-panel">
                        <ul class="id-find-info pw-find-info">
                            <li class="bullet-ty1">본인 명의 휴대폰 번호로 가입여부 및 본인여부를 확인합니다. </li>
                            <li class="bullet-ty1">타인명의/법인휴대폰은 본인인증이 불가능합니다.</li>
                        </ul>
						<form name="form" id="form" autocomplete="off">
						<input type="hidden" name="SMode" value="SearchPwd" />
						<div id="FW_HP_Rst" class="t-type4-inform-hp" style="display:none">
							<span id="FW_HP_Rst_Msg" class="nt-type-bl-r"></span>
						</div>
                        <fieldset id="FW_HP_Form">
                            <legend class="hidden">휴대폰 정보 입력</legend>
                            <div class="fieldset">
                                <label for="user-name" class="fieldset-label">이름</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="Name" id="Name2" maxlength="25" placeholder="본인의 실명을 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="user-id" class="fieldset-label">아이디</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="UserID" id="UserID2" maxlength="30" placeholder="아이디를 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                        </fieldset>
                        <div id="FW_HP_Btn" class="area-btn-c">
                            <button type="button" onclick="chk_FindPW_AuthHP('form')" class="button is-expand ty-red">휴대폰 인증</button>
                        </div>
                        <div id="FW_HP_Fail_Btn" class="area-btn-c" style="display:none">
                            <button type="button" onclick="re_FindPW_AuthHP()" class="button is-expand ty-red">다시찾기</button>
                        </div>
						</form>
                    </div>
                    <div id="tabs-col3" class="tab-panel">
                        <ul class="id-find-info pw-find-info">
                            <li class="bullet-ty1">아이핀으로 회원가입을 하신 경우, 아이핀 인증을 통해 비밀번호를 찾을 수 있습니다. </li>
                        </ul>
						<form name="form1" id="form1" autocomplete="off">
						<input type="hidden" name="SMode" value="SearchPwd" />
						<div id="FW_Ipin_Rst" class="t-type4-inform-hp" style="display:none">
							<span id="FW_Ipin_Rst_Msg" class="nt-type-bl-r"></span>
						</div>
                        <fieldset id="FW_Ipin_Form">
                            <legend class="hidden">이메일 정보 입력</legend>
                            <div class="fieldset">
                                <label for="user-name" class="fieldset-label">이름</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="Name" id="Name3" maxlength="25" placeholder="본인의 실명을 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="user-id" class="fieldset-label">아이디</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="UserID" id="UserID3" maxlength="30" placeholder="아이디를 입력해주세요.">
                                    </span>
                                </div>
                            </div>
                        </fieldset>
                        <div id="FW_Ipin_Btn" class="area-btn-c">
                            <button type="button" onclick="chk_FindPW_AuthIpin('form1')" class="button is-expand ty-red">아이핀 인증</button>
                        </div>
                        <div id="FW_Ipin_Fail_Btn" class="area-btn-c" style="display:none">
                            <button type="button" onclick="re_FindPW_AuthIpin()" class="button is-expand ty-red">다시찾기</button>
                        </div>
						</form>
                    </div>

                    <div id="tabs-col4" class="tab-panel1" style="display:none;">
                        <ul class="id-find-info pw-find-info">
                            <li class="bullet-ty1">본인인증에 성공하였습니다.새로운 비밀번호를 입력해 주세요.</li>
                            <li class="bullet-ty1">회원님의 비밀번호는 암호화되어 저장되기 때문에 재설정으로 진행됩니다.</li>
                        </ul>
						<form name="formChgPwd" id="formChgPwd" method="post" autocomplete="off">
						<input type="hidden" id="getUserID" name="UserID">
                        <fieldset>
                            <legend class="hidden">비밀번호 초기화</legend>
                            <div class="fieldset">
                                <label for="user-name" class="fieldset-label">비밀번호</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="password" name="Pwd" id="newPw" maxlength="12" placeholder="비밀번호 (영문 숫자포함 6~12)">
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="user-id" class="fieldset-label">비밀번호 확인</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="password" name="Pwd1" id="newPwCheck" maxlength="12" placeholder="비밀번호를 한번 더 입력해 주세요.">
                                    </span>
                                </div>
                            </div>
                        </fieldset>
                        <div class="area-btn-c">
                            <button type="button" onclick="chk_ChgPwd()" class="button is-expand ty-red">확 인</button>
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