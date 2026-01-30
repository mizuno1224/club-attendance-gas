/**
 * 設定・定数
 */
const DEFAULT_ROSTER = ["ゆうり", "みゆ", "のぞみ", "えみり", "まな", "まほ", "まい", "しん"];
const TZ = "Asia/Tokyo";
const PROP_LOG_SHEET_ID = "LOG_SHEET_ID";
// キャッシュキーを "V5" に変更して、過去の全キャッシュ（祝日データ含む）を無効化
const CACHE_PREFIX = "CLUB_DATA_V5_"; 
const CACHE_EXPIRATION = 21600;
// キャッシュ有効時間（秒） = 6時間

const SCHEDULE_SHEET_CANDIDATES = ["活動予定", "スケジュール", "Schedule", "schedule"];
const ATTENDANCE_SHEET_CANDIDATES = ["出欠", "出席", "Attendance", "attendance"];
const HEADER_ALIASES = {
  // schedule
  date: ["日付", "日", "Date", "date", "DATE", "日にち"],
  off: ["休み", "オフ", "OFF", "off", "Off", "中止"],
  morning: ["午前", "AM", "am", "朝", "morning", "Morning"],
  afternoon: ["午後", "PM", "pm", "afternoon", "Afternoon"],
  after: ["業後", "after", "After", "afterWork", "夜", "夜間", "ナイター"],
  note: ["メモ", "備考", "内容", "note", "Note"],
  place: ["場所", "会場", "place", "Place", "グラウンド"],
  time: ["時間", "活動時間", "集合", "集合時間", "gather", "gatherAt", "time", "Time"],
  // attendance
  present: ["出席", "参加", "present", "Present"],
  absent: ["欠席", "不参加", "absent", "Absent"],
  tardy: ["遅刻", "tardy", "Tardy", "Late", "遅参"],
  early: ["早退", "early", "Early", "早上がり"]
};