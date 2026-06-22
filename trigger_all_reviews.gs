function setGitHubTokenAllReviews() {
  PropertiesService.getScriptProperties()
    .setProperty('GITHUB_TOKEN_ALL_REVIEWS', 'YOUR_NEW_GITHUB_PAT_HERE');
  Logger.log('All-Reviews Token saved!');
}

function triggerGitHubAllReviews() {
  const GITHUB_TOKEN = PropertiesService.getScriptProperties()
                       .getProperty('GITHUB_TOKEN_ALL_REVIEWS');
  const GITHUB_REPO  = 'Likithayathvik/boat-review-all';   // update if repo name differs

  const response = UrlFetchApp.fetch(
    `https://api.github.com/repos/${GITHUB_REPO}/dispatches`,
    {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + GITHUB_TOKEN,
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ event_type: 'weekly-review-trigger' }),
      muteHttpExceptions: true
    }
  );
  Logger.log('All-Reviews GitHub triggered! Status: ' + response.getResponseCode());
}

function createMondayTriggerAllReviews() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'triggerGitHubAllReviews') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('triggerGitHubAllReviews')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .nearMinute(0)
    .create();
  Logger.log('All-Reviews Monday 11 AM trigger created!');
}
