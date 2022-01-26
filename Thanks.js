const WEBINAR = {
  morning : {
    url : 'https://zoom.us/s/99663862892',
    id : '996 6386 2892'
  },
  afternoon : {
    url : 'https://zoom.us/s/92479112437',
    id : '924 7911 2437'
  }
};



// shift と @ 同時押しで ``
// 自動返信メール本文
function createBody(name, attendDate, webinar) {
  return `
${name} 様

いつも大変お世話になっております。
株式会社LillyHoldingsの加藤でございます。

この度はLillyMEOサービス説明へのお申し込みをいただきましてありがとうございます！
ウェビナーの参加リンクをお送りさせていただきますのでご確認ください。

【参加日時】
${attendDate}

【参加URL】
${webinar.url}
ウェビナーID：${webinar.id}

パソコンで計測ツールにログインした状態でご参加ください。
※計測ツールは下記のアドレスよりお送りさせていただいているのでご確認ください。

【計測ツール送信アドレス】
lillyholdings.customers@gmail.com


当日は、何卒よろしくお願い致します。

株式会社LillyHoldings 加藤`;
}


function isAfternoon(attendDate) {
  const day = attendDate.match(/(?<=\().*?(?=\))/)[0];
  const afternoonDayArray = ['月', '水', '金'];
  console.log(afternoonDayArray.indexOf(day), afternoonDayArray.indexOf(day) >= 0);
  return afternoonDayArray.indexOf(day) >= 0;
}


function sendWebinerUrl(e) {
  // フォームの回答を取得
  const name = e.namedValues['担当者様名'][0];
  const email = e.namedValues['メールアドレス'][0];
  const attendDate = e.namedValues['参加希望日'][0];
  
  // 自動返信メール件名
  const subject = '【ウェビナー】LillyMEOサービス説明';

  // 自動返信メール本文
  let body;
  if ( isAfternoon(attendDate) ) {
    body = createBody(name, attendDate, WEBINAR['afternoon']);
  } else {
    body = createBody(name, attendDate, WEBINAR['morning']);
  }
  console.log(email, subject, body);
  
  // メール送信
  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body
  });
}