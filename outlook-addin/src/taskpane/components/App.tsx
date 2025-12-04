/* eslint-disable no-undef */
import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
//import { ChatCompletionRequestMessage, Configuration, OpenAIApi } from "openai";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  generatedText: string;
  startText: string;
  finalMailText: string;
  isLoading: boolean;
  isGenerateBusinessMailActive: boolean;
  isSummarizeMailActive: boolean;
  summary: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props) {
    super(props);

    let isGenerateBusinessMailActive;
    let isSummarizeMailActive;

    //get the current URL
    const url = window.location.href;
    console.log("URL: " + url);
    //check if the URL contains the parameter "generate"
    if (url.indexOf("compose") > -1) {
      console.log("Action: generate business mail");
      isGenerateBusinessMailActive = true;
      isSummarizeMailActive = false;
    }
    //check if the URL contains the parameter "summarize"
    if (url.indexOf("summary") > -1) {
      console.log("Action: summarize mail");
      isGenerateBusinessMailActive = false;
      isSummarizeMailActive = true;
    }

    this.state = {
      generatedText: "",
      startText: "",
      finalMailText: "",
      isLoading: false,
      isGenerateBusinessMailActive: isGenerateBusinessMailActive,
      isSummarizeMailActive: isSummarizeMailActive,
      summary: "",
    };
  }

  showGenerateBusinessMail = () => {
    this.setState({ isGenerateBusinessMailActive: true, isSummarizeMailActive: false });
  };

  showSummarizeMail = () => {
    this.setState({ isGenerateBusinessMailActive: false, isSummarizeMailActive: true });
  };

generateText = async () => {
  this.setState({ isLoading: true });

  try {
    const response = await fetch("http://localhost:11434/api/generate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "llama3",
        prompt: `
          Sei un assistente che scrive email professionali.
          Riscrivimi questa richiesta come una mail professionale:

          ${this.state.startText}
        `,
        stream: false
      })
    });

    const data = await response.json();

    this.setState({ 
      generatedText: data.response,
      isLoading: false
    });

  } catch (error) {
    console.error(error);
    this.setState({ generatedText: "Errore nella generazione del testo", isLoading: false });
  }
};


  insertIntoMail = () => {
    const finalText = this.state.finalMailText.length === 0 ? this.state.generatedText : this.state.finalMailText;
    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, {
      coercionType: Office.CoercionType.Text,
    });
  };

  onSummarize = async () => {
    try {
      this.setState({ isLoading: true });
      const summary = await this.summarizeMail();
      this.setState({ summary: summary, isLoading: false });
    } catch (error) {
      this.setState({ summary: error, isLoading: false });
    }
  };

  summarizeMail(): Promise<any> {
  return new Office.Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Text,
        async (asyncResult) => {

          const mailText = asyncResult.value.split(" ").slice(0, 800).join(" ");

          try {
            const response = await fetch("http://localhost:11434/api/generate", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                model: "llama3",
                prompt: `
                  Riassumi questa email in una lista puntata:

                  ${mailText}
                `,
                stream: false
              })
            });

            const data = await response.json();
            resolve(data.response);

          } catch (err) {
            reject(err);
          }
        }
      );
    } catch (error) {
      reject(error);
    }
  });
}

  ProgressSection = () => {
    if (this.state.isLoading) {
      return <Progress title="Loading..." message="The AI is working..." />;
    } else {
      return <> </>;
    }
  };

  BusinessMailSection = () => {
    if (this.state.isGenerateBusinessMailActive) {
      return (
        <>
          <p>Briefly describe what you want to communicate in the mail:</p>
          <textarea
            className="ms-welcome"
            onChange={(e) => this.setState({ startText: e.target.value })}
            rows={5}
            cols={40}
          />
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.generateText}
            >
              Generate text
            </DefaultButton>
          </p>
          <this.ProgressSection />
          <textarea
            className="ms-welcome"
            defaultValue={this.state.generatedText}
            onChange={(e) => this.setState({ finalMailText: e.target.value })}
            rows={15}
            cols={40}
          />
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.insertIntoMail}
            >
              Insert into mail
            </DefaultButton>
          </p>
        </>
      );
    } else {
      return <div> </div>;
    }
  };

  SummarizeMailSection = () => {
    if (this.state.isSummarizeMailActive) {
      return (
        <>
          <p>Summarize mail</p>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.onSummarize}
          >
            Summarize mail
          </DefaultButton>
          <this.ProgressSection />
          <textarea className="ms-welcome" defaultValue={this.state.summary} rows={15} cols={40} />
        </>
      );
    } else {
      return <div> </div>;
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <main className="ms-welcome__main">
          <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            Outlook AI Assistant
          </h2>

          <p className="ms-font-l ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            Choose your service:
          </p>
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.showGenerateBusinessMail}
            >
              Generate business mail
            </DefaultButton>
          </p>
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.showSummarizeMail}
            >
              Summarize mail
            </DefaultButton>
          </p>
          <div>
            <this.BusinessMailSection />
          </div>
          <div>
            <this.SummarizeMailSection />
          </div>
        </main>
      </div>
    );
  }
}
