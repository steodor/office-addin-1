import React, { useEffect } from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import {
  addSelectionHandler,
  getSelectedText,
  insertEndnote,
  insertFootnote,
  removeSelectionHandler,
} from "../office-document";
import { AppTabProps } from "../types";
import useEvent from "react-use-event-hook";

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
  button: {
    marginBottom: "10px",
  },
});

const ReferenceInsertion: React.FC<AppTabProps> = () => {
  const [text, setText] = useState<string>("Some reference text");

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const handleFootnoteInsertion = useEvent(async () => {
    await insertFootnote(text);
  });

  const handleEndnoteInsertion = useEvent(async () => {
    await insertEndnote(text);
  });

  const [selection, setSelection] = useState<string>("");

  useEffect(() => {
    getSelectedText((result: string) => {
      setSelection(result);
    });

    const handler = (event: Office.DocumentSelectionChangedEventArgs) => {
      getSelectedText((result: string) => {
        setSelection(result);
      });
      console.log(event);
    };

    addSelectionHandler(handler);

    return () => {
      removeSelectionHandler(handler);
    };
  }, []);

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Current selection:">
        <Textarea size="large" value={selection} disabled />
      </Field>
      <Field className={styles.textAreaField} size="large" label="Enter reference text for the current selection">
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      <Field className={styles.instructions}>Click any button to insert the reference.</Field>
      <Field>
        <Button
          appearance="primary"
          disabled={false}
          size="large"
          onClick={handleFootnoteInsertion}
          className={styles.button}
        >
          Insert Footnote
        </Button>
        <Button
          appearance="primary"
          disabled={false}
          size="large"
          onClick={handleEndnoteInsertion}
          className={styles.button}
        >
          Insert Endnote
        </Button>
      </Field>
    </div>
  );
};

export default ReferenceInsertion;
