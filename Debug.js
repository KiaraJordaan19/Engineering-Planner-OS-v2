function debugCalendarSync() {
  console.log('DEBUG: starting api_syncCalendar');

  const started = Date.now();
  const result = api_syncCalendar();

  console.log(
    'DEBUG: finished after ' +
    (Date.now() - started) +
    ' ms'
  );

  console.log(JSON.stringify(result, null, 2));
}