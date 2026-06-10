/**
 * Demo Request Form — Google Form generator
 * ------------------------------------------
 * Paste this whole file into https://script.google.com (new project),
 * then run buildDemoRequestForm() once. It will:
 *   1. Create a new Google Form with all the fields below
 *   2. Create a Google Sheet and wire responses into it
 *   3. Print the Form edit link, the shareable (send-to-reps) link,
 *      and the responses Sheet link in the Execution log
 *
 * Re-running it creates a brand-new form each time, so you normally
 * only run it once. Tweak the values here, re-run, and you get a fresh copy.
 */

function buildDemoRequestForm() {
  // ---- Create the form -----------------------------------------------------
  var form = FormApp.create('Demo Request Form');
  form.setDescription(
    'Request a change to a demo app. Fill in the basics, set a priority, ' +
    'and describe what needs to change. Fields marked * are required.'
  );

  // Capture each submitter's email so you know who sent the request.
  // Set to false if your reps are outside the org / you don't want sign-in.
  form.setCollectEmail(true);
  form.setProgressBar(true);

  // ========================================================================
  // SECTION 1 — THE BASICS
  // ========================================================================
  form.addSectionHeaderItem()
    .setTitle('The Basics');

  form.addTextItem()
    .setTitle('Request Name')
    .setHelpText('App name + what needs to change. Keep it skimmable. ' +
                 'e.g. "Agent — Screen Pop Customization"')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Your Name')
    .setHelpText('First Last')
    .setRequired(true);

  form.addDateItem()
    .setTitle('Date Requested')
    .setRequired(false);

  // ========================================================================
  // SECTION 2 — PRIORITY & TRACKING
  // ========================================================================
  form.addSectionHeaderItem()
    .setTitle('Priority & Tracking');

  form.addTextItem()
    .setTitle('Application(s)')
    .setHelpText('Which demo app(s) does this change apply to? e.g. "Agent, Supervisor"')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Jira #')
    .setHelpText('Leave blank — Eternals will link after approval. e.g. "DEMO-482"')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Sprint #')
    .setHelpText('e.g. "Sprint 44"')
    .setRequired(false);

  form.addMultipleChoiceItem()
    .setTitle('Priority')
    .setHelpText('How urgent is this request?')
    .setChoiceValues([
      'Critical — Imminent, no flexibility',
      'High — Timely, deal at stake',
      'Medium — Standard timeline',
      'Low — No rush, early stage'
    ])
    .setRequired(true);

  // ========================================================================
  // SECTION 3 — TELL US MORE
  // ========================================================================
  form.addSectionHeaderItem()
    .setTitle('Tell Us More');

  form.addParagraphTextItem()
    .setTitle('Full Description')
    .setHelpText('What needs to change and why? Describe current behavior, ' +
                 'expected behavior, and any specific scenarios to cover. ' +
                 'The more detail, the better we can scope it.')
    .setRequired(true);

  // ---- Wire responses into a Google Sheet ---------------------------------
  var ss = SpreadsheetApp.create('Demo Request Form — Responses');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // ---- Report the links ---------------------------------------------------
  Logger.log('✅ Form created!');
  Logger.log('Edit the form here:        ' + form.getEditUrl());
  Logger.log('Send THIS link to reps:    ' + form.getPublishedUrl());
  Logger.log('Responses Sheet:           ' + ss.getUrl());
}
