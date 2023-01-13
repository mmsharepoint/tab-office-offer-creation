import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult, ActionTypes, MessagingExtensionAttachment } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import GraphSearchService from "../api/GraphSearchService";
import { IOfferDocument } from "../../model/IOfferDcoument";
import jwtDecode from "jwt-decode";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/offerReviewYoteamsMessageExtension/config.html")
export default class OfferReviewYoteamsMessageExtension implements IMessagingExtensionMiddlewareProcessor {
  private connectionName = process.env.ConnectionName;
  private documents: IOfferDocument[];
  public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
    const attachments: MessagingExtensionAttachment[] = [];
    const adapter: any = context.adapter;
    const magicCode = (query.state && Number.isInteger(Number(query.state))) ? query.state : '';        
    const tokenResponse = await adapter.getUserToken(context, this.connectionName, magicCode);
    if (!tokenResponse || !tokenResponse.token) {
      // There is no token, so the user has not signed in yet.

      // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
      const signInLink = await adapter.getSignInLink(context, this.connectionName);
      let composeExtension: MessagingExtensionResult = {
        type: 'auth',
        suggestedActions: {
          actions: [{
            title: 'Sign in as user',
            value: signInLink,
            type: ActionTypes.OpenUrl
          }]
        }
      };
      return Promise.resolve(composeExtension);
    }
    let documents: IOfferDocument[] = [];
    if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
      const controller = new GraphSearchService();
      documents = await controller.getFiles(tokenResponse.token);
      this.documents = documents;
    }
    documents.forEach((doc) => {
      const card = CardFactory.adaptiveCard(
        {
          type: "AdaptiveCard",
          body: [
                  {
                    type: "ColumnSet",
                    columns: [
                          {
                            type: "Column",
                            width: 25,
                            items: [
                              {
                                type: "Image",
                                url: `https://${process.env.HOSTNAME}/assets/icon.png`,
                                style: "Person"
                              }
                            ]
                          },
                          {
                            type: "Column",
                            width: 75,
                            items: [
                              {
                                type: "TextBlock",
                                text: doc.name,
                                size: "Large",
                                weight: "Bolder"
                              },
                              {
                                type: "TextBlock",
                                text: doc.description,
                                size: "Medium"
                              },
                              {
                                type: "TextBlock",
                                text: `Author: ${doc.author}`
                              },
                              {
                                type: "TextBlock",
                                text: `Modified: ${doc.modified.toLocaleDateString()}`
                              }
                            ]
                          }
                      ]
                  }                     
              ],
              actions: [
                  {
                      type: "Action.OpenUrl",
                      title: "View",
                      url: doc.url
                  },
                  {
                    type: "Action.ShowCard",
                    title: "Review",
                    card: {
                      type: "AdaptiveCard",
                      body: [
                        {
                          type: "Input.Text",
                          isVisible: false,
                          value: doc.id,
                          id: "id"
                        },
                        {
                          type: "Input.Text",
                          isVisible: false,
                          value: "reviewed",
                          id: "action"
                        }
                      ],
                      actions: [
                        {
                          type: "Action.Submit",
                          title: "Reviewed"
                            
                        }
                      ],
                      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                      version: "1.0"
                    }
                  }                        
              ],
              $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
              version: "1.0"
          });            
      const preview = {
        contentType: "application/vnd.microsoft.card.thumbnail",
        content: {
          title: doc.name,
          text: doc.description,
          images: [
              {
                  url: `https://${process.env.HOSTNAME}/assets/icon.png`
              }
          ]             
        }
      };
      attachments.push({ contentType: card.contentType, content: card.content, preview: preview });
    });
    
    return Promise.resolve({
        type: "result",
        attachmentLayout: "list",
        attachments: attachments
      } as MessagingExtensionResult);
    }

    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
      // Handle the Action.Submit action on the adaptive card
      if (value.action === "reviewed") {
        log(`I got this ${value.id}`);
        const adapter: any = context.adapter;
        const magicCode = (value.state && Number.isInteger(Number(value.state))) ? value.state : '';        
        const tokenResponse = await adapter.getUserToken(context, this.connectionName, magicCode);
        if (!tokenResponse || !tokenResponse.token) {
          // There is no token, so the user has not signed in yet.            
          return Promise.reject();
        }
        // Get user's Email from the token (as the context.activity only offers display name)
        const decoded: { [key: string]: any; } = jwtDecode(tokenResponse.token) as { [key: string]: any; };
        const controller = new GraphSearchService();
        await controller.reviewItem(tokenResponse.token, value.id, decoded.upn!);
      }
      return Promise.resolve();
    }

    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
      return Promise.resolve({
        title: "Offer Review (yoteams) Configuration",
        value: `https://${process.env.PUBLIC_HOSTNAME}/offerReviewYoteamsMessageExtension/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`
      });
    }

    public async onSettings(context: TurnContext): Promise<void> {
      // take care of the setting returned from the dialog, with the value stored in state
      const setting = context.activity.value.state;
      log(`New setting: ${setting}`);
      return Promise.resolve();
    }

}
