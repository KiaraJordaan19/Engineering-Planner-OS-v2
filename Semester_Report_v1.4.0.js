/**
 * Semester_Report_v1.4.0.gs — Engineering Planner OS v1.4.0, Semester
 * Lifecycle & Analytics.
 *
 * Builds the deterministic Semester Report (semReportBuild_), its
 * accompanying deterministic charts (semChartsBuild_), and semester
 * completion detection (semDetectCompletion_). Nothing in this file writes
 * to the workbook except the one-time, idempotent "assign a Semester ID if
 * one isn't set yet" step inside semGetOrCreateSemesterId_ — every other
 * function here is a pure, read-only aggregation over data every other
 * system already owns (Marks Intelligence, Study Planner, Revision Tracker,
 * Attendance, Resources, Academic Intelligence, Automation Log).
 *
 * INTEGRATION, NOT DUPLICATION: this file never re-implements a calculation
 * another system already owns. Per-module Official Final Mark, winning FM
 * route, required A3, and risk level all come from calling
 * computeAcademicIntelligence_() (API_v1.3.1.gs) — the exact same pure
 * function api_getPlannerData() already calls on every page load — with a
 * bundle built from api_getPlannerData()'s own already-mapped arrays. No
 * Marks Intelligence, Study Planner, Revision, Attendance, Resources, or
 * Academic Intelligence file is modified by v1.4.0.
 *
 * Depends on (already present, unmodified): api_getPlannerData, ok_/fail_,
 * getSetting_, readSheetRows_, num_, isoDate_, round1_, daysBetween_,
 * isoWeekKey_, pctToNumber_, computeAcademicIntelligence_, nextId_,
 * HEADER_ROW, SETTINGS_SHEET.
 */

// ------------------------------------------------------------------
// SEMESTER ID
// ------------------------------------------------------------------
/**
 * Every archive/backup/report needs a stable Semester ID to key off. Most
 * workbooks won't have one until the New Semester Wizard is completed once
 * (semWizardApply_ in Semester_Wizard_v1.4.0.gs sets it explicitly) — this
 * lazily assigns one (format "SEM-001", next after every ID already used in
 * "23 Semester Archive" or "26 Semester History") ONLY if "02 Settings" >
 * "Semester ID" is still blank, so Generate Report / Generate Backup never
 * fail on a pre-Wizard workbook. Idempotent: once set, always returns the
 * same value until a new semester is started.
 */
function semGetOrCreateSemesterId_() {
  var existing = getSetting_('Semester ID');
  if (existing) return existing;
  var archiveIds = readSheetRows_('23 Semester Archive', HEADER_ROW, 'Semester ID').map(function (r) { return r['Semester ID']; });
  var historyIds = readSheetRows_('26 Semester History', HEADER_ROW, 'Semester ID').map(function (r) { return r['Semester ID']; });
  var newId = nextId_('SEM', archiveIds.concat(historyIds));
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  var cell = settingsSheet.createTextFinder('Semester ID').matchEntireCell(false).findNext();
  if (cell) settingsSheet.getRange(cell.getRow(), 3).setValue(newId);
  return newId;
}

// ------------------------------------------------------------------
// COMPLETION DETECTION (spec §1)
// ------------------------------------------------------------------
/**
 * A semester is "ready to finish" when every scheduled assessment has
 * finished (date has passed AND status is not "Scheduled" — i.e. every
 * assessment is Completed or Cancelled, never inferred from date alone if
 * the status still says "Scheduled") OR the user has explicitly pressed
 * "Finish Semester" (Settings > "Semester manually finished (TRUE/FALSE)").
 * NEVER triggers the actual delete/reset workflow by itself — this only
 * decides whether the Semester screen shows the "Finish Semester" prompt.
 * Every call to the real workflow still requires its own separate explicit
 * confirmation (see Semester_Delete_v1.4.0.gs).
 */
function semDetectCompletion_(pdData) {
  var todayIso = isoDate_(new Date());
  var assessments = pdData.assessments || [];
  var total = assessments.length;
  var finished = assessments.filter(function (a) {
    if (a.status === 'Completed' || a.status === 'Cancelled') return true;
    if (a.status === 'Scheduled') return false; // still scheduled -- never assume finished from date alone
    // No status recorded: treat as finished only once its date has clearly passed.
    return a.date && daysBetween_(todayIso, a.date) !== null && daysBetween_(todayIso, a.date) < 0;
  });
  var manuallyFinished = getSetting_('Semester manually finished (TRUE/FALSE)') === true;
  var allAssessmentsFinished = total > 0 && finished.length === total;
  return {
    totalAssessments: total,
    finishedAssessments: finished.length,
    allAssessmentsFinished: allAssessmentsFinished,
    manuallyFinished: manuallyFinished,
    readyToFinish: allAssessmentsFinished || manuallyFinished,
    reason: manuallyFinished
      ? 'Manually marked finished via "Finish Semester".'
      : (allAssessmentsFinished ? 'All ' + total + ' scheduled assessment(s) have finished.' : (total - finished.length) + ' assessment(s) still scheduled.')
  };
}

// ------------------------------------------------------------------
// REPORT BUILDER (spec §2)
// ------------------------------------------------------------------
/**
 * reflectionInput: optional { whatWorkedWell, whatShouldImprove,
 * goalsNextSemester } strings typed by the user on the Semester screen —
 * reflection has no sheet of its own (nothing to guess or infer), so it is
 * only ever whatever was typed immediately before this call. Pass {} or
 * omit to generate a report with blank reflection fields (e.g. a
 * "preview" report, not tied to archiving).
 */
function semReportBuild_(reflectionInput) {
  var pd = api_getPlannerData();
  if (!pd.ok) throw new Error('Could not read workbook data: ' + pd.error);
  var data = pd.data;
  var todayIso = isoDate_(new Date());
  var reflection = reflectionInput || {};

  var modules = data.modules.filter(function (m) { return m.active; });
  var byModuleName = {};
  modules.forEach(function (m) { byModuleName[m.name] = m; });

  // ---- Reuse the exact same Academic Intelligence bundle api_getPlannerData
  // already builds internally -- see file header note. Nothing recomputed a
  // second way; this is one more caller of the same pure function. ----
  var academicIntelligence = computeAcademicIntelligence_({
    modules: data.modules, marksTracker: data.marksTracker, assessments: data.assessments,
    studySessions: data.studySessions, studyTasks: data.studyTasks, revisionTracker: data.revisionTracker,
    resources: data.resources, afComponents: data.afComponents, settings: data.settings,
    logRows: data.syncLog, todayIso: todayIso
  });
  var miByName = {};
  academicIntelligence.marksIntelligence.forEach(function (mi) { miByName[mi.name] = mi; });

  // =========================== OVERVIEW ===========================
  var studyStats = semBuildStudyStatistics_(data.studySessions, modules);
  var revisionHoursTotal = round1_((data.revisionTracker || []).reduce(function (s, r) { return s + (num_(r.totalActualMinutes) || 0); }, 0) / 60);
  var assignmentsCompleted = (data.assignments || []).filter(function (a) { return a.status === 'Submitted' || a.status === 'Graded'; }).length;
  var assessmentsWritten = (data.assessments || []).filter(function (a) { return a.status === 'Completed'; }).length;
  var attendanceSummary = semBuildAttendance_(data.attendance, byModuleName, data.studySessions);

  var overview = {
    semester: data.settings.semesterName || '',
    academicYear: getSetting_('Academic year') || '',
    semesterStart: data.settings.semesterStart || '',
    semesterEnd: data.settings.semesterEnd || '',
    modules: modules.map(function (m) { return { code: m.code, name: m.name }; }),
    studyHours: studyStats.totalHours,
    revisionHours: revisionHoursTotal,
    attendanceSummary: attendanceSummary,
    assignmentsCompleted: assignmentsCompleted,
    assessmentsWritten: assessmentsWritten,
    // Overall workload: a single deterministic combined figure -- study
    // hours + revision hours + estimated assignment hours logged as
    // completed. Disclosed judgment call (same spirit as pctToNumber_'s doc
    // comment elsewhere in this codebase): "workload" here means TIME spent
    // across the three trackers that record it, not a difficulty score.
    overallWorkload: round1_(studyStats.totalHours + revisionHoursTotal +
      (data.assignments || []).reduce(function (s, a) { return s + (num_(a.hoursCompleted) || 0); }, 0))
  };

  // =========================== MARKS ===========================
  var marksPerModule = modules.map(function (m) {
    var row = (data.marksTracker || []).filter(function (r) { return r.module === m.name; })[0];
    var mi = miByName[m.name];
    var requiredA3 = null, requiredA3Status = 'insufficient_data';
    if (mi && mi.status === 'ok' && mi.a3Recovery) {
      var targetEntry = mi.a3Recovery.targets.filter(function (t) { return t.label === 'Pass'; })[0] || mi.a3Recovery.targets[0];
      var bestRoute = targetEntry ? (targetEntry.routes.filter(function (r) { return r.isBest; })[0] || targetEntry.routes[0]) : null;
      if (bestRoute) { requiredA3 = bestRoute.regulationValidRequiredA3; requiredA3Status = bestRoute.status; }
    }
    return {
      code: m.code, module: m.name,
      af: row ? num_(row.af) : null, a1: row ? num_(row.a1) : null, a2: row ? num_(row.a2) : null, a3: row ? num_(row.a3) : null,
      officialFinalMark: (mi && typeof mi.officialFinalMark === 'number') ? mi.officialFinalMark : null,
      highestFmRoute: (mi && mi.winningRoute) ? mi.winningRoute : null,
      requiredA3: requiredA3, requiredA3Status: requiredA3Status,
      currentMark: (mi && typeof mi.currentMark === 'number') ? mi.currentMark : null,
      status: mi ? mi.status : 'insufficient_data'
    };
  });
  var numericMarks = marksPerModule.filter(function (r) { return typeof r.currentMark === 'number'; });
  var passMarkDefault = data.settings.passMark || 50, distinctionMarkDefault = data.settings.distinctionMark || 75;
  var distinctionCount = numericMarks.filter(function (r) {
    var m = byModuleName[r.module]; var d = (m && typeof m.distinctionMark === 'number') ? m.distinctionMark : distinctionMarkDefault;
    return r.currentMark >= d;
  }).length;
  var passCount = numericMarks.filter(function (r) {
    var m = byModuleName[r.module]; var p = (m && typeof m.passMark === 'number') ? m.passMark : passMarkDefault;
    return r.currentMark >= p; // NOTE: pass count includes distinctions (a distinction is still a pass) -- disclosed convention.
  }).length;
  var marks = {
    perModule: marksPerModule,
    overallAverage: numericMarks.length ? round1_(numericMarks.reduce(function (s, r) { return s + r.currentMark; }, 0) / numericMarks.length) : null,
    highestMark: numericMarks.length ? round1_(Math.max.apply(null, numericMarks.map(function (r) { return r.currentMark; }))) : null,
    lowestMark: numericMarks.length ? round1_(Math.min.apply(null, numericMarks.map(function (r) { return r.currentMark; }))) : null,
    distinctionCount: distinctionCount, passCount: passCount
  };

  // ===================== REVISION STATISTICS =====================
  var revRows = data.revisionTracker || [];
  var withConfidence = revRows.filter(function (r) { return typeof r.confidence === 'number'; });
  var sortedByConf = withConfidence.slice().sort(function (a, b) { return a.confidence - b.confidence; });
  var revisionStatistics = {
    totalTopics: revRows.length,
    topicsCompleted: revRows.filter(function (r) { return r.completed; }).length,
    // "Mastered" = disclosed judgment call: completed AND confidence >= 4 (of a 1-5 confidence scale, this workbook's existing convention).
    topicsMastered: revRows.filter(function (r) { return r.completed && typeof r.confidence === 'number' && r.confidence >= 4; }).length,
    topicsNeedingReview: revRows.filter(function (r) { return r.weakTopic || (typeof r.confidence === 'number' && r.confidence <= 2); }).length,
    averageConfidence: withConfidence.length ? round1_(withConfidence.reduce(function (s, r) { return s + r.confidence; }, 0) / withConfidence.length) : null,
    weakestTopics: sortedByConf.slice(0, 5).map(function (r) { return { module: r.module, topic: r.topic, confidence: r.confidence }; }),
    strongestTopics: sortedByConf.slice(-5).reverse().map(function (r) { return { module: r.module, topic: r.topic, confidence: r.confidence }; })
  };

  // ======================== RESOURCES =========================
  var resRows = data.resources || [];
  var pastPapers = resRows.filter(function (r) { return r.type === 'Past paper'; });
  var formulaBible = resRows.filter(function (r) { return r.type === 'Formula bible'; });
  var resources = {
    resourcesAdded: resRows.length,
    resourcesCompleted: resRows.filter(function (r) { return r.status === 'Completed'; }).length,
    resourcesReviewed: resRows.filter(function (r) { return r.status === 'Reviewed' || r.reviewed; }).length,
    pastPapersAttempted: pastPapers.filter(function (r) { return r.flagA; }).length, // flag A = "Attempted" for Past paper (see RESOURCE_PROGRESS_LABELS in Index.html)
    formulaBibleUsage: {
      count: formulaBible.length,
      reviewed: formulaBible.filter(function (r) { return r.flagA; }).length,
      memorised: formulaBible.filter(function (r) { return r.flagB; }).length,
      recallTested: formulaBible.filter(function (r) { return r.flagC; }).length,
      appliedInQuestions: formulaBible.filter(function (r) { return r.flagD; }).length
    }
  };

  // =================== ACADEMIC INTELLIGENCE ===================
  var riskCounts = { Low: 0, Medium: 0, High: 0, Impossible: 0, 'Insufficient data': 0 };
  academicIntelligence.marksIntelligence.forEach(function (mi) {
    var level = (mi.riskAnalysis && mi.riskAnalysis.level) ? mi.riskAnalysis.level : 'Insufficient data';
    riskCounts[level] = (riskCounts[level] || 0) + 1;
  });
  var academicIntel = {
    priorityHistory: academicIntelligence.priorities.map(function (p) { return { code: p.code, name: p.name, category: p.category, score: p.score }; }),
    examFocusActivations: {
      count: (academicIntelligence.examFocus.active || []).length,
      modules: (academicIntelligence.examFocus.active || []).map(function (a) { return a.module; })
    },
    // No dedicated log of "recovery plans generated" exists anywhere in the
    // workbook (the A3 Recovery Calculator computes on demand, live, and is
    // never itself logged) -- disclosed proxy: modules with a currently
    // active recovery scenario (recoveryDifficulty is neither "Not
    // applicable" nor "Insufficient data") at report time.
    recoveryPlansGenerated: academicIntelligence.marksIntelligence.filter(function (mi) {
      return mi.recoveryDifficulty && mi.recoveryDifficulty !== 'Not applicable' && mi.recoveryDifficulty !== 'Insufficient data';
    }).length,
    riskDistribution: riskCounts,
    improvementTrackerSummary: academicIntelligence.improvementTracker
  };

  // ========================== TIMELINE ==========================
  var timeline = semBuildTimeline_(data, overview.semesterStart, overview.semesterEnd, todayIso);

  var report = {
    meta: {
      semesterId: semGetOrCreateSemesterId_(),
      semesterName: overview.semester,
      academicYear: overview.academicYear,
      generatedAt: new Date().toISOString(),
      plannerVersion: getSetting_('Planner version') || '1.4.0'
    },
    overview: overview,
    marks: marks,
    studyStatistics: studyStats,
    revisionStatistics: revisionStatistics,
    attendance: attendanceSummary,
    resources: resources,
    academicIntelligence: academicIntel,
    timeline: timeline,
    reflection: {
      whatWorkedWell: reflection.whatWorkedWell || '',
      whatShouldImprove: reflection.whatShouldImprove || '',
      goalsNextSemester: reflection.goalsNextSemester || ''
    }
  };
  return report;
}

/** Study Statistics section (spec §2) -- also reused for the Overview's
 *  "Study hours" figure so the two never disagree. */
function semBuildStudyStatistics_(studySessions, activeModules) {
  var rows = studySessions || [];
  var totalMinutes = 0;
  var byModule = {}, byWeek = {}, byType = {};
  var CANONICAL_TYPES = ['Preview', 'Learn', 'Practice', 'Mistake', 'Recall', 'Formula', 'Programming', 'Exam', 'Weekly Review'];
  CANONICAL_TYPES.forEach(function (t) { byType[t] = 0; });
  byType['Other'] = 0;
  var completed = 0;
  rows.forEach(function (s) {
    var minutes = num_(s.actualDuration) || num_(s.duration) || 0;
    totalMinutes += minutes;
    byModule[s.module] = (byModule[s.module] || 0) + minutes;
    if (s.date) {
      var wk = isoWeekKey_(s.date);
      if (wk) byWeek[wk] = (byWeek[wk] || 0) + minutes;
    }
    var typeKey = CANONICAL_TYPES.indexOf(s.type) !== -1 ? s.type : 'Other';
    byType[typeKey] += minutes;
    if (s.status === 'Completed' || s.status === 'Partly completed' || s.completed) completed++;
  });
  var hoursPerModule = Object.keys(byModule).map(function (m) { return { module: m, hours: round1_(byModule[m] / 60) }; })
    .sort(function (a, b) { return b.hours - a.hours; });
  var hoursPerWeek = Object.keys(byWeek).sort().map(function (wk) { return { weekOf: wk, hours: round1_(byWeek[wk] / 60) }; });
  var typeDistribution = {};
  Object.keys(byType).forEach(function (t) { typeDistribution[t] = round1_(byType[t] / 60); });
  return {
    totalHours: round1_(totalMinutes / 60),
    hoursPerModule: hoursPerModule,
    hoursPerWeek: hoursPerWeek,
    typeDistribution: typeDistribution,
    completionRate: rows.length ? round1_((completed / rows.length) * 100) : null
  };
}

/** Attendance section (spec §2) -- also used at Overview level. */
function semBuildAttendance_(attendanceRows, byModuleName, studySessions) {
  var rows = attendanceRows || [];
  var recorded = rows.filter(function (r) { return r.status !== 'Not recorded'; });
  var attended = recorded.filter(function (r) { return r.status === 'Attended' || r.attended; });
  var missed = recorded.filter(function (r) { return r.status === 'Missed'; });
  var tutorials = recorded.filter(function (r) { return (r.type || '').toLowerCase().indexOf('tutorial') !== -1; });
  var tutorialsAttended = tutorials.filter(function (r) { return r.status === 'Attended' || r.attended; });
  var compulsory = recorded.filter(function (r) { var m = byModuleName[r.module]; return m && m.compulsoryPractical; });
  var compulsoryAttended = compulsory.filter(function (r) { return r.status === 'Attended' || r.attended; });
  var catchUps = (studySessions || []).filter(function (s) {
    return (s.origin || '').toLowerCase() === 'catch-up' || /^catch-up:/i.test(s.topic || '');
  }).length;
  return {
    attendanceRate: recorded.length ? round1_((attended.length / recorded.length) * 100) : null,
    tutorialAttendance: tutorials.length ? round1_((tutorialsAttended.length / tutorials.length) * 100) : null,
    compulsorySessionAttendance: compulsory.length ? round1_((compulsoryAttended.length / compulsory.length) * 100) : null,
    missedSessions: missed.length,
    catchUpSessionsCreated: catchUps
  };
}

/** Timeline section (spec §2). Deliberately derives A1/A2/Exam periods from
 *  REAL assessment records (09 Assessments, filtered by "Assessment type"
 *  containing "A1"/"A2"/"A3"/"Exam") rather than the hardcoded phase-boundary
 *  table in currentSemesterPhaseIndex_ (API_v1.3.1.gs) -- that table is
 *  specific to one semester's actual calendar and already carries its own
 *  "update this by hand every semester" caveat (see updateSemesterDates()'s
 *  own warning message). Deriving from real assessment dates instead is the
 *  only way this section stays correct across many future semesters without
 *  code changes -- directly serving spec §"supporting multiple years without
 *  becoming cluttered". */
function semBuildTimeline_(data, semesterStart, semesterEnd, todayIso) {
  function periodFor(matchFn) {
    var dates = (data.assessments || []).filter(matchFn).map(function (a) { return a.date; }).filter(Boolean).sort();
    if (!dates.length) return { start: null, end: null, count: 0 };
    return { start: dates[0], end: dates[dates.length - 1], count: dates.length };
  }
  var a1Period = periodFor(function (a) { return (a.type || '').toUpperCase().indexOf('A1') !== -1; });
  var a2Period = periodFor(function (a) { return (a.type || '').toUpperCase().indexOf('A2') !== -1; });
  var examPeriod = periodFor(function (a) { return /A3|EXAM/i.test(a.type || ''); });

  var milestones = [];
  if (semesterStart) milestones.push({ date: semesterStart, label: 'Semester start' });
  (data.assessments || []).forEach(function (a) {
    if (a.date) milestones.push({ date: a.date, label: (a.type || 'Assessment') + ' -- ' + a.module + (a.title ? ' (' + a.title + ')' : '') });
  });
  (data.assignments || []).forEach(function (a) {
    if (a.dueDate && (a.status === 'Submitted' || a.status === 'Graded')) {
      milestones.push({ date: a.dueDate, label: 'Assignment submitted -- ' + a.module + ': ' + a.title });
    }
  });
  if (semesterEnd) milestones.push({ date: semesterEnd, label: 'Semester end' });
  milestones.sort(function (x, y) { return (x.date || '').localeCompare(y.date || ''); });

  // Achievements: deterministic, rule-based, templated strings only -- no
  // generative/AI text of any kind, every line traceable to a real record.
  // (See semBuildAchievements_ below.)
  return { a1Period: a1Period, a2Period: a2Period, examPeriod: examPeriod, majorMilestones: milestones, achievements: semBuildAchievements_(data) };
}

function semBuildAchievements_(data) {
  var achievements = [];
  var byModuleName = {};
  data.modules.forEach(function (m) { byModuleName[m.name] = m; });
  (data.marksTracker || []).forEach(function (r) {
    var m = byModuleName[r.module];
    if (!m) return;
    var finalMark = pctToNumber_(r.finalOrProvisional) || pctToNumber_(r.provisionalAfterA2);
    if (typeof finalMark === 'number' && typeof m.distinctionMark === 'number' && finalMark >= m.distinctionMark) {
      achievements.push(finalMark + '% (Distinction) achieved in ' + r.module + '.');
    }
  });
  var attendanceByModule = {};
  (data.attendance || []).forEach(function (a) {
    if (a.status === 'Not recorded') return;
    attendanceByModule[a.module] = attendanceByModule[a.module] || { attended: 0, total: 0 };
    attendanceByModule[a.module].total++;
    if (a.status === 'Attended' || a.attended) attendanceByModule[a.module].attended++;
  });
  Object.keys(attendanceByModule).forEach(function (mod) {
    var d = attendanceByModule[mod];
    if (d.total >= 5 && (d.attended / d.total) >= 0.95) achievements.push('95%+ attendance maintained in ' + mod + '.');
  });
  (data.revisionTracker || []).filter(function (r) { return r.completed; }).length &&
    achievements.push((data.revisionTracker || []).filter(function (r) { return r.completed; }).length + ' revision topic(s) completed this semester.');
  return achievements;
}

// ------------------------------------------------------------------
// CHARTS (spec §3) -- deterministic, no external libraries. Every chart is
// plain data ({label, value[, color]} arrays); the Semester screen renders
// them as simple CSS width-percentage bars, the same "no framework, no CDN
// chart library" approach the rest of this app already uses for every other
// visual element.
// ------------------------------------------------------------------
function semChartsBuild_(report) {
  return {
    studyHoursByWeek: report.studyStatistics.hoursPerWeek.map(function (w) { return { label: w.weekOf, value: w.hours }; }),
    studyHoursByModule: report.studyStatistics.hoursPerModule.map(function (m) { return { label: m.module, value: m.hours }; }),
    marksByModule: report.marks.perModule.map(function (m) { return { label: m.module, value: (typeof m.currentMark === 'number') ? m.currentMark : 0 }; }),
    attendance: [
      { label: 'Attended', value: report.attendance.attendanceRate === null ? 0 : report.attendance.attendanceRate },
      { label: 'Missed sessions', value: report.attendance.missedSessions }
    ],
    revisionCompletion: [
      { label: 'Completed', value: report.revisionStatistics.topicsCompleted },
      { label: 'Needing review', value: report.revisionStatistics.topicsNeedingReview }
    ],
    resourceCompletion: [
      { label: 'Added', value: report.resources.resourcesAdded },
      { label: 'Completed', value: report.resources.resourcesCompleted },
      { label: 'Reviewed', value: report.resources.resourcesReviewed }
    ],
    // Proxy for "priority history" -- see academicIntelligence.priorityHistory
    // doc comment in semReportBuild_ for why there is no literal week-by-week
    // priority score log anywhere in the workbook. This chart instead shows
    // how many pieces of Improvement Tracker evidence exist per category,
    // which IS a real, dated, deterministic record.
    priorityHistory: (function () {
      var byCategory = {};
      (report.academicIntelligence.improvementTrackerSummary || []).forEach(function (m) {
        byCategory[m.category] = (byCategory[m.category] || 0) + (m.evidenceCount || 1);
      });
      return Object.keys(byCategory).map(function (c) { return { label: c, value: byCategory[c] }; });
    })()
  };
}
