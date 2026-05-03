var SPREADSHEET_ID = "1DQKketDHjco_qsyc4jTFfALkXLomlij_srDC5gf0JyA";
var SHEET_NAME = "Assessment Results";

function doPost(e) {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet =
    spreadsheet.getSheetByName(SHEET_NAME) ||
    spreadsheet.insertSheet(SHEET_NAME);

  var payload = JSON.parse((e.postData && e.postData.contents) || "{}");
  var headers = [
    "submitted_at",
    "attempt_id",
    "student_name",
    "student_age",
    "school_grade",
    "parent_email",
    "self_rating",
    "final_level_code",
    "final_level_name",
    "suggested_class_level",
    "final_level_summary",
    "final_level_detail",
    "final_note",
    "total_correct",
    "total_questions",
    "grammar_correct",
    "grammar_questions",
    "reading_correct",
    "reading_questions",
    "rounds_completed",
    "round_scores_json",
    "reading_results_json",
    "writing_prompt_text",
    "writing_prompt_id",
    "writing_prompt_image",
    "writing_prompt_alt",
    "writing_prompt_responses_json",
    "writing_score",
    "writing_word_count",
    "writing_sentence_count",
    "writing_level_judgement",
    "writing_task_response",
    "writing_main_weakness",
    "writing_improvement_tip",
    "writing_original_text",
    "writing_strengths_json",
    "writing_issues_json",
    "writing_areas_json",
    "writing_corrections_json",
    "answers_json"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var writing = payload.writing || {};

  sheet.appendRow([
    payload.submittedAt || "",
    payload.attemptId || "",
    payload.studentName || "",
    payload.studentAge || "",
    payload.schoolGrade || "",
    payload.parentEmail || "",
    payload.selfRating || "",
    payload.finalLevelCode || "",
    payload.finalLevelName || "",
    payload.suggestedClassLevel || "",
    payload.finalLevelSummary || "",
    payload.finalLevelDetail || "",
    payload.finalNote || "",
    payload.totalCorrect || "",
    payload.totalQuestions || "",
    payload.grammarCorrect || "",
    payload.grammarQuestions || "",
    payload.readingCorrect || "",
    payload.readingQuestions || "",
    payload.roundsCompleted || "",
    JSON.stringify(payload.roundScores || []),
    JSON.stringify(payload.readingResults || []),
    writing.promptText || "",
    writing.promptId || "",
    writing.promptImageSrc || "",
    writing.promptImageAlt || "",
    JSON.stringify(writing.promptResponses || []),
    writing.score || "",
    writing.wordCount || "",
    writing.sentenceCount || "",
    writing.levelJudgement || "",
    writing.taskResponse || "",
    writing.mainWeakness || "",
    writing.improvementTip || "",
    writing.originalText || "",
    JSON.stringify(writing.strengths || []),
    JSON.stringify(writing.issues || []),
    JSON.stringify(writing.areasToImprove || []),
    JSON.stringify(writing.correctionExamples || []),
    JSON.stringify(payload.answers || [])
  ]);

  return ContentService.createTextOutput(
    JSON.stringify({ ok: true })
  ).setMimeType(ContentService.MimeType.JSON);
}
