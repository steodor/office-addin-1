export const insertText = async (text: string) => {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      //body.insertBreak(Word.BreakType.line, Word.InsertLocation.end);
      text.split("\n").forEach((line) => {
        // body.insertParagraph(line, Word.InsertLocation.end);
        body.insertText(line, Word.InsertLocation.end);
        body.insertBreak(Word.BreakType.line, Word.InsertLocation.end);
      });
      // body.insertHtml(text.replace(/\n/g, "<br />"), Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export const insertFootnote = async (text: string) => {
  try {
    await Word.run(async (context) => {
      context.document.getSelection().insertFootnote(text);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export const insertEndnote = async (text: string) => {
  try {
    await Word.run(async (context) => {
      context.document.getSelection().insertEndnote(text);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export const getSelectedText = (callback: (string) => void) => {
  try {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      (asyncResult: Office.AsyncResult<string>) => {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          console.error("Action failed. Error: " + asyncResult.error.message);
          return "";
        } else {
          callback(asyncResult.value);
          return asyncResult.value;
        }
      }
    );
  } catch (error) {
    console.log("Error: " + error);
  }
};

export const addSelectionHandler = async (handler: (event: Office.DocumentSelectionChangedEventArgs) => void) => {
  try {
    return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
  } catch (error) {
    console.log("Error: " + error);
  }
};

export const removeSelectionHandler = async (handler: (eventArgs?: any) => any) => {
  try {
    return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export default insertText;
