/**
 * @OnlyCurrentDoc
 */

/**
 * The script property key for the OpenAI API key.
 * @type {string}
 */
const OPENAI_API_KEY_PROPERTY = 'OPENAI_API_KEY';

/**
 * The base URL for the OpenAI API.
 * @type {string}
 */
const OPENAI_API_BASE_URL = 'https://api.openai.com/v1/';

/**
 * Retrieves the OpenAI API key from the script properties.
 * @returns {string} The OpenAI API key.
 * @throws {Error} If the OPENAI_API_KEY script property is not set.
 */
function getOpenAiApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(OPENAI_API_KEY_PROPERTY);
  if (!apiKey) {
    throw new Error('The OPENAI_API_KEY script property is not set.');
  }
  return apiKey;
}

/**
 * Calls the OpenAI API with the specified parameters.
 * @param {string} model The OpenAI model to use.
 * @param {Array<Object>} messages The messages to send to the model.
 * @returns {Object} The response from the OpenAI API.
 */
function logError(error, context) {
  Logger.log(`Error in ${context}: ${error.toString()}`);
}

function callOpenAI(model, messages) {
  try {
    const apiKey = getOpenAiApiKey();
    const url = `${OPENAI_API_BASE_URL}chat/completions`;
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify({
        model: model,
        messages: messages,
      }),
    };

    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    logError(e, 'callOpenAI');
    return null; // Return null or a more specific error object
  }
}

/**
 * Creates course content using OpenAI.
 * @param {string} topic The topic for the content.
 * @returns {string} The generated content.
 */
function createCourseContent(topic) {
  const messages = [
    {
      role: 'system',
      content: 'You are a helpful assistant that generates educational content.',
    },
    {
      role: 'user',
      content: `Create a short lesson on the topic: ${topic}`,
    },
  ];

  const response = callOpenAI('gpt-3.5-turbo', messages);
  if (!response || !response.choices || response.choices.length === 0) {
    logError(new Error('Invalid response from OpenAI'), 'createCourseContent');
    return 'Error generating content.';
  }
  return response.choices[0].message.content;
}

/**
 * Creates a quiz using OpenAI.
 * @param {string} topic The topic for the quiz.
 * @param {number} numQuestions The number of questions to generate.
 * @returns {Array<Object>} The generated quiz questions.
 */
function createQuiz(topic, numQuestions) {
  const messages = [
    {
      role: 'system',
      content: 'You are a helpful assistant that generates quizzes in JSON format.',
    },
    {
      role: 'user',
      content: `Create a ${numQuestions}-question multiple-choice quiz on the topic: ${topic}. Provide the output in JSON format with an array of objects, where each object has "question", "options" (an array of strings), and "answer" keys.`,
    },
  ];

  const response = callOpenAI('gpt-3.5-turbo', messages);
  if (!response || !response.choices || response.choices.length === 0) {
    logError(new Error('Invalid response from OpenAI'), 'createQuiz');
    return [];
  }
  const content = response.choices[0].message.content;
  try {
    // Attempt to parse the JSON from the response.
    const quiz = JSON.parse(content);
    return quiz;
  } catch (e) {
    logError(e, 'createQuiz - JSON parsing');
    // Handle cases where the response is not valid JSON
    return [];
  }
}

/**
 * Grades an assignment using OpenAI.
 * @param {string} submission The student's submission.
 * @param {string} rubric The grading rubric.
 * @returns {number} The calculated grade.
 */
function gradeAssignment(submission, rubric) {
  const messages = [
    {
      role: 'system',
      content: 'You are a helpful assistant that grades assignments based on a rubric and provides a numerical score.',
    },
    {
      role: 'user',
      content: `Please grade the following submission based on the provided rubric.\n\nSubmission:\n${submission}\n\nRubric:\n${rubric}\n\nPlease provide a single numerical score from 0 to 100 and nothing else.`,
    },
  ];

  const response = callOpenAI('gpt-3.5-turbo', messages);
  if (!response || !response.choices || response.choices.length === 0) {
    logError(new Error('Invalid response from OpenAI'), 'gradeAssignment');
    return 0; // or handle the error as appropriate
  }
  const grade = parseInt(response.choices[0].message.content.trim(), 10);
  if (isNaN(grade)) {
    logError(new Error(`Could not parse grade from response: ${response.choices[0].message.content}`), 'gradeAssignment');
    return 0; // or handle the error as appropriate
  }
  return grade;
}

/**
 * Provides feedback on a student's submission using OpenAI.
 * @param {string} submission The student's submission.
 * @returns {string} The generated feedback.
 */
function provideFeedback(submission) {
  const messages = [
    {
      role: 'system',
      content: 'You are a helpful assistant that provides constructive feedback on student work.',
    },
    {
      role: 'user',
      content: `Please provide feedback on the following submission:\n\n${submission}`,
    },
  ];

  const response = callOpenAI('gpt-3.5-turbo', messages);
  if (!response || !response.choices || response.choices.length === 0) {
    logError(new Error('Invalid response from OpenAI'), 'provideFeedback');
    return 'Could not generate feedback.';
  }
  return response.choices[0].message.content;
}

const PROGRESS_TRACKING_SHEET_NAME = 'Progress Tracking';

function setupProgressTrackingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(PROGRESS_TRACKING_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(PROGRESS_TRACKING_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Course ID', 'Student ID', 'Grade']);
  }
  return sheet;
}

/**
 * Tracks student progress by logging their grade for a course.
 * @param {string} studentId The ID of the student.
 * @param {string} courseId The ID of the course.
 * @param {number} grade The grade received by the student.
 */
function trackProgress(studentId, courseId, grade) {
  const sheet = setupProgressTrackingSheet();
  const timestamp = new Date();
  sheet.appendRow([timestamp, courseId, studentId, grade]);
}

/**
 * Retrieves course data from the Google Classroom API.
 * @param {string} courseId The ID of the course.
 * @returns {Object} The course data.
 */
function getCourseData(courseId) {
  try {
    const course = Classroom.Courses.get(courseId);
    return course;
  } catch (e) {
    logError(e, 'getCourseData');
    return null;
  }
}

/**
 * Submits a grade to the Google Classroom API.
 * @param {string} courseId The ID of the course.
 * @param {string} courseworkId The ID of the coursework.
 * @param {string} studentId The ID of the student.
 * @param {number} grade The grade to submit.
 */
function submitGrade(courseId, courseworkId, studentId, grade) {
  try {
    const studentSubmissions = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, courseworkId, {
      userId: studentId
    }).studentSubmissions;

    if (studentSubmissions && studentSubmissions.length > 0) {
      const submissionId = studentSubmissions[0].id;
      const submission = {
        assignedGrade: grade,
        draftGrade: grade
      };
      const updateMask = 'assignedGrade,draftGrade';
      Classroom.Courses.CourseWork.StudentSubmissions.patch(submission, courseId, courseworkId, submissionId, {
        updateMask: updateMask
      });
      Logger.log(`Successfully submitted grade for student ${studentId} in course ${courseId}`);
      trackProgress(studentId, courseId, grade);
    } else {
      Logger.log(`No submission found for student ${studentId} in course ${courseId}`);
    }
  } catch (e) {
    logError(e, 'submitGrade');
  }
}

/**
 * Renders the homepage card for the add-on.
 * @param {Object} e The event object.
 * @returns {Card} The homepage card.
 */
function onHomepage(e) {
  return createHomepageCard();
}

/**
 * Creates the homepage card for the add-on.
 * @returns {Card} The homepage card.
 */
function createHomepageCard() {
  const builder = CardService.newCardBuilder();
  builder.setHeader(CardService.newCardHeader().setTitle('OpenAI Classroom Assistant'));
  const section = CardService.newCardSection();

  const createContentButton = CardService.newTextButton()
      .setText('Create Course Content')
      .setOnClickAction(CardService.newAction().setFunctionName('showCreateContentUi'));

  const createQuizButton = CardService.newTextButton()
      .setText('Create a Quiz')
      .setOnClickAction(CardService.newAction().setFunctionName('showCreateQuizUi'));

  const gradeAssignmentsButton = CardService.newTextButton()
      .setText('Grade Assignments')
      .setOnClickAction(CardService.newAction().setFunctionName('showGradeAssignmentsUi'));

  const provideFeedbackButton = CardService.newTextButton()
      .setText('Provide Feedback')
      .setOnClickAction(CardService.newAction().setFunctionName('showProvideFeedbackUi'));

  const viewProgressButton = CardService.newTextButton()
      .setText('View Progress')
      .setOpenLink(CardService.newOpenLink()
          .setUrl(SpreadsheetApp.getActiveSpreadsheet().getUrl()));

  section.addWidget(createContentButton)
      .addWidget(createQuizButton)
      .addWidget(gradeAssignmentsButton)
      .addWidget(provideFeedbackButton)
      .addWidget(viewProgressButton);
  builder.addSection(section);
  return builder.build();
}

/**
 * Shows the UI for creating course content.
 * @returns {ActionResponse} The action response.
 */
function showCreateContentUi() {
  const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Create Course Content'))
      .addSection(
          CardService.newCardSection()
              .addWidget(CardService.newTextInput().setFieldName('topic').setTitle('Topic'))
              .addWidget(
                  CardService.newTextButton()
                      .setText('Generate')
                      .setOnClickAction(
                          CardService.newAction().setFunctionName('handleCreateContent'))
              )
      )
      .build();

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card))
      .build();
}

function showGradeAssignmentsUi() {
  const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Grade Assignments'))
      .addSection(
          CardService.newCardSection()
              .addWidget(CardService.newTextInput().setFieldName('studentId').setTitle('Student ID'))
              .addWidget(CardService.newTextInput().setFieldName('courseworkId').setTitle('Coursework ID'))
              .addWidget(CardService.newTextInput().setFieldName('submissionText').setTitle('Submission Text'))
              .addWidget(CardService.newTextInput().setFieldName('rubric').setTitle('Rubric'))
              .addWidget(
                  CardService.newTextButton()
                      .setText('Grade')
                      .setOnClickAction(CardService.newAction().setFunctionName('handleGradeAssignment'))
              )
      )
      .build();

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card))
      .build();
}

function showProvideFeedbackUi() {
  const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Provide Feedback'))
      .addSection(
          CardService.newCardSection()
              .addWidget(CardService.newTextInput().setFieldName('studentId').setTitle('Student ID'))
              .addWidget(CardService.newTextInput().setFieldName('courseworkId').setTitle('Coursework ID'))
              .addWidget(CardService.newTextInput().setFieldName('submissionText').setTitle('Submission Text'))
              .addWidget(
                  CardService.newTextButton()
                      .setText('Provide Feedback')
                      .setOnClickAction(CardService.newAction().setFunctionName('handleProvideFeedback'))
              )
      )
      .build();

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card))
      .build();
}

function handleGradeAssignment(e) {
  const studentId = e.formInput.studentId;
  const courseworkId = e.formInput.courseworkId;
  const submissionText = e.formInput.submissionText;
  const rubric = e.formInput.rubric;
  const courseId = e.classroom.courseId;

  const grade = gradeAssignment(submissionText, rubric);
  submitGrade(courseId, courseworkId, studentId, grade);

  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Assignment graded successfully.'))
      .setNavigation(CardService.newNavigation().popToRoot())
      .build();
}

function handleProvideFeedback(e) {
  const studentId = e.formInput.studentId;
  const courseworkId = e.formInput.courseworkId;
  const submissionText = e.formInput.submissionText;
  const courseId = e.classroom.courseId;

  const feedback = provideFeedback(submissionText);

  try {
    const studentSubmissions = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, courseworkId, {
      userId: studentId
    }).studentSubmissions;

    if (!studentSubmissions || studentSubmissions.length === 0) {
      return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification().setText('Submission not found.'))
          .build();
    }
    const submissionId = studentSubmissions[0].id;

    // Create a Google Doc with the feedback
    const feedbackDoc = DocumentApp.create(`Feedback for submission`);
    feedbackDoc.getBody().setText(feedback);
    const feedbackDocUrl = feedbackDoc.getUrl();

    const addAttachmentsRequest = {
      addAttachments: [{
        link: { url: feedbackDocUrl }
      }]
    };

    Classroom.Courses.CourseWork.StudentSubmissions.modifyAttachments(addAttachmentsRequest, courseId, courseworkId, submissionId);

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Feedback provided successfully as an attached document.'))
        .setNavigation(CardService.newNavigation().popToRoot())
        .build();
  } catch (err) {
    logError(err, 'handleProvideFeedback');
    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Failed to provide feedback.'))
        .build();
  }
}

/**
 * Handles the creation of course content.
 * @param {Object} e The event object from the UI.
 * @returns {ActionResponse} The action response.
 */
function handleCreateContent(e) {
  const topic = e.formInput.topic;
  const content = createCourseContent(topic);
  const courseId = e.classroom.courseId;

  const announcement = {
    text: content,
    state: 'PUBLISHED'
  };
  Classroom.Courses.Announcements.create(announcement, courseId);

  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Content created successfully.'))
      .setNavigation(CardService.newNavigation().popToRoot())
      .build();
}

/**
 * Handles the creation of a quiz.
 * @param {Object} e The event object from the UI.
 * @returns {ActionResponse} The action response.
 */
function handleCreateQuiz(e) {
  const topic = e.formInput.topic;
  const numQuestions = parseInt(e.formInput.numQuestions, 10);
  const quizQuestions = createQuiz(topic, numQuestions);
  const courseId = e.classroom.courseId;

  if (quizQuestions.length === 0) {
    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Failed to create quiz.'))
        .build();
  }

  // Create a new Google Form
  const form = FormApp.create(`Quiz: ${topic}`);
  form.setDescription('Please complete the following quiz.');
  form.setQuiz(true);

  quizQuestions.forEach(q => {
    const item = form.addMultipleChoiceItem();
    item.setTitle(q.question);
    const choices = q.options.map(optionText => {
      return item.createChoice(optionText, optionText === q.answer);
    });
    item.setChoices(choices);
    item.setRequired(true);
  });

  const formUrl = form.getPublishedUrl();

  const quizAssignment = {
    title: `Quiz: ${topic}`,
    description: 'Please complete the following quiz.',
    workType: 'ASSIGNMENT',
    state: 'PUBLISHED',
    materials: [
      {
        link: {
          url: formUrl
        }
      }
    ]
  };
  Classroom.Courses.CourseWork.create(quizAssignment, courseId);

  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Quiz created successfully.'))
      .setNavigation(CardService.newNavigation().popToRoot())
      .build();
}

/**
 * Shows the UI for creating a quiz.
 * @returns {ActionResponse} The action response.
 */
function showCreateQuizUi() {
  const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Create a Quiz'))
      .addSection(
          CardService.newCardSection()
              .addWidget(CardService.newTextInput().setFieldName('topic').setTitle('Topic'))
              .addWidget(CardService.newTextInput().setFieldName('numQuestions').setTitle('Number of Questions'))
              .addWidget(
                  CardService.newTextButton()
                      .setText('Generate')
                      .setOnClickAction(CardService.newAction().setFunctionName('handleCreateQuiz'))
              )
      )
      .build();

  return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card))
      .build();
}
