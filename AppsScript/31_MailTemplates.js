// 메일 본문 HTML(처음 메일·리마인드·[변경사항 없음] 버튼)

/** 메일 하단에 넣는 [변경사항 없음] 버튼 블록 */
function buildResponseButtons_(token) {
  // [안전장치①] 토큰이 비면 버튼 링크가 '?token=' 로 깨져 직원이 "접근 권한 필요"를 보게 됨.
  // 그런 메일은 아예 만들지 않고 발송을 실패시켜(→ 발송실패 기록 + 요약메일 경고) 눈에 띄게 함.
  if (!token) throw new Error('응답토큰이 비어 있어 버튼 메일을 만들 수 없습니다. (발송전_점검 실행 필요)');
  const noneUrl = `${CONFIG.WEB_APP_URL}?token=${encodeURIComponent(token)}`;
  return `
<div style="margin:22px 0 4px;">
  <div style="font-size:13px; color:#333; margin-bottom:10px;">
    <b>변경사항이 없으신 경우</b> 아래 <b>[변경사항 없음]</b> 버튼을 눌러주세요.<br>
    변경사항이 있으신 경우에는 위 안내대로 <b>2주 이내</b>에 인사총무팀으로 <b>이메일 또는 서면</b>으로 통보해 주시기 바랍니다.
  </div>
  <table cellpadding="0" cellspacing="0" border="0"><tr>
    <td>
      <a href="${noneUrl}" style="display:inline-block; background:#15803d; color:#fff; text-decoration:none; font-size:14px; font-weight:bold; padding:12px 22px; border-radius:8px;">✅ 변경사항 없음</a>
    </td>
  </tr></table>
  <div style="font-size:12px; color:#888; margin-top:9px;">
    * 버튼을 누르면 확인 화면이 한 번 뜨고, 거기서 [제출]을 눌러야 최종 접수됩니다.
  </div>
</div>`;
}

function buildRequestHtml_(d, token) {
  return `
<div style="font-size:14px; line-height:1.6; color:#000;">
안녕하십니까 ${d.recipientName} ${d.recipientTitle}님<br>
${CONFIG.발신표시명} ${CONFIG.담당자성함} ${CONFIG.담당자직책}입니다.<br><br>

당 그룹사에서 운용중인 직원태양광 법인 정보가 변경될 경우,<br>
경영지원본부 인사총무팀에 변경사항을 알려주시기를 요청드리려 메일드렸습니다.<br><br>

요청사항은 아래와 같습니다.<br>
-------------------------------------------------------------------<br>

<table cellpadding="0" cellspacing="0" border="0" style="font-size:14px; line-height:1.6; color:#000;">
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
등기부등본 내 정보의 변동사항 발생 시(<u>대표이사 주소지 변경 등</u>),<br>
변경일로부터 <b><span style="color:#0000FF;">2주 이내</span></b> 인사총무팀에 이메일 또는 서면으로 변경된 정보 통보
</td></tr>
<tr><td colspan="2" style="height:8px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
통보 시, <b><span style="color:#0000FF;">변경 내용과 변경된 날짜</span></b>를 명확히 기재
</td></tr>
<tr><td colspan="2" style="height:8px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
<b>첨부드린 등기부등본</b>에서 아래 항목이 현재와 같은지 확인<br>
${buildCheckItemsHtml_()}
</td></tr>
<tr><td colspan="2" style="height:8px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
법인 정보<br>
- 법인명: ${d.corpName}<br>
- 임원: ${d.exec1Name} ${d.exec1Title}${d.exec2Name ? ` / ${d.exec2Name} ${d.exec2Title}` : ''}<br>
- 태양광 대상자: ${d.solarTarget}
</td></tr>
<tr><td colspan="2" style="height:8px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
<tr><td style="width:10px; vertical-align:top;">•</td><td style="vertical-align:top;">
인사총무팀 ${CONFIG.담당자성함} ${CONFIG.담당자직책} E-mail: ${CONFIG.담당자이메일}<br>
인사총무팀 ${CONFIG.사수성함} ${CONFIG.사수직책} E-mail: ${CONFIG.사수이메일}<br>
본사 사무실 주소: ${CONFIG.본사주소}
</td></tr>
</table>

-------------------------------------------------------------------<br>

<b>변경 사항을 2주 내로 보고하지 않을 경우 과태료가 부과되오니, 신속하게 변경 사항을 통보해주시기 바랍니다.<br>
미통보로 인한 과태료 발생시 당사에서는 어떠한 책임도 지지 않음을 알려드립니다.</b><br><br>

내용 확인하여 주시고, 관련하여 궁금하신 사항은 언제든 편하게 문의주시면 감사하겠습니다.<br>

감사합니다.<br>
${buildResponseButtons_(token)}
<br>
<span style="text-decoration:underline; font-size:12px;">*본 메일은 분기마다 정기적으로 발송되는 메일입니다.</span>
</div>`;
}

/** 처음 메일·리마인드 공통: 첨부한 등기부등본에서 확인할 항목 */
function buildCheckItemsHtml_() {
  return `- 대표이사 주소<br>
- 임원 구성(대표이사·사내이사)<br>
- 본점 주소`;
}

function buildReminderHtml_(d, token) {
  return `
<div style="font-size:14px; line-height:1.6; color:#000;">
안녕하십니까 ${d.recipientName} ${d.recipientTitle}님<br>
${CONFIG.발신표시명} ${CONFIG.담당자성함} ${CONFIG.담당자직책}입니다.<br><br>

지난 ${formatDate_(d.lastSendDate)} ${d.corpName} 법인 등기부등본 변경사항 관련하여 안내드린 메일에 대해<br>
아직 회신을 받지 못하여 다시 한 번 안내드립니다.<br><br>

확인이 쉽도록 <b>${d.corpName} 등기부등본을 첨부</b>하였습니다.<br>
첨부 파일에서 아래 항목이 현재와 같은지 확인 부탁드립니다.<br>
${buildCheckItemsHtml_()}<br><br>

바쁘시겠지만 확인 부탁드립니다.<br>
감사합니다.
${buildResponseButtons_(token)}
</div>`;
}
