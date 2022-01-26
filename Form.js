const FORM_ID = "1F1MBTnGLyDdwSD6_Hil0wsuGbGx8hT6G_FUZIZQ0BeI";

/**
 * main関数
 */
function updateWebinarDates() {
  const countDate = new Date();
  countDate.setDate(countDate.getDate() + 2);

  const maximumDate = new Date();
  maximumDate.setMonth(maximumDate.getMonth() + 1);

  const attendableDates = getAttendableDates(countDate, maximumDate);
  replacePulldownList(attendableDates);
}


function getWeekNum(endDate){
  const startDate = new Date('2021/11/01');
  const differentDays = (endDate - startDate) / 1000 / 60 / 60 / 24;
  const weekNum = Math.floor(differentDays / 7) + 1;
  return weekNum;
}


function getWebinarTimeRange(day) {
  if (day === 1 || day === 3 || day === 5) {
    return '15:00~16:00';
  } else if (day === 2 || day === 4) {
    return '11:00~12:00';
  }
}


function replaceDay(day) {
  const dayArray = ['日', '月', '火', '水', '木', '金', '土'];
  return dayArray[day];
}


/**
 * リリー　休業日カレンダーに予定があるかチェックする関数
 */
function isHoliday(date) {
  const startDate = new Date(date.setHours(0, 0, 0, 0));
  const endDate = new Date(date.setHours(23, 59, 59));
  const cal = CalendarApp.getCalendarById("c_aguhmn5sn8v3sjjk6am3t8gjms@group.calendar.google.com");  // [リリー　休業日]を取得
  const holidays =  cal.getEvents(startDate, endDate);

  return holidays.length !== 0;
}


/**
 * 開催不可日程をチェックする関数
 */
function isNonHeldDay(date, nonHeldDays) {
  let isNonHeldDay = false;
  nonHeldDays.forEach(nonHeldDay => {
    const nonHeldDate = new Date(nonHeldDay[0]);
    date.setHours(0, 0, 0, 0);
    nonHeldDate.setHours(0, 0, 0, 0);
    if ( date.getTime() === nonHeldDate.getTime() ) {
      isNonHeldDay = true;
    };
  });
  return isNonHeldDay;
}


function replacePulldownList(attendableDates){
  const form = FormApp.openById(FORM_ID);
  const items = form.getItems();
  const attendDateQuestion = items[2];
  console.log(attendDateQuestion.getTitle());

  attendDateQuestion.asMultipleChoiceItem().setChoiceValues(attendableDates);
}


function getAttendableDate(countDate) {
  const day = countDate.getDay();

  const dateValue = Utilities.formatDate(countDate, 'JST', 'M月dd日');
  const dayValue = ` (${ replaceDay(day) })`;
  const timeRange = getWebinarTimeRange(day);

  return dateValue + dayValue + timeRange;
}


function addAttendableDate(attendableDates, countDate, validation) {
  if (validation) {
    const attendableDate = getAttendableDate(countDate);
    attendableDates.push(attendableDate);
    return countDate.setDate( countDate.getDate() + 1 );
  } else {
    return countDate.setDate( countDate.getDate() + 1 );
  }
}


function getAttendableDates(countDate, maximumDate) {
  const attendableDates = [];
  const nonHeldDays = SpreadsheetApp.openById('1i1clofuGoWpB-lgZ1g04d6TlgcAScYibvdBfD0O_rFg').getSheetByName('開催不可日程').getDataRange().getValues();
  nonHeldDays.shift();

  while (countDate.getTime() <= maximumDate.getTime()) {
    const day = countDate.getDay();
    const weekNum = getWeekNum(countDate);
    if ( isHoliday(countDate) || isNonHeldDay(countDate, nonHeldDays) ) {
      countDate.setDate( countDate.getDate() + 1 );
    } else if (weekNum % 2 === 0) {
      const validation = (day === 2 || day === 4);
      addAttendableDate(attendableDates, countDate, validation);
    } else {
      const validation = (day === 1 || day === 3 || day === 5);
      addAttendableDate(attendableDates, countDate, validation);
    }
  }
  console.log(attendableDates);
  return attendableDates;
}