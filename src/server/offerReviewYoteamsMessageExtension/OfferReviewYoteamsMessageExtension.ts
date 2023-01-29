import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult, ActionTypes, MessagingExtensionAttachment, TeamsInfo, StatusCodes } from "botbuilder";
import { AdaptiveCardRequestValue, AdaptiveCardResponseBody, IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import GraphSearchService from "../api/GraphSearchService";
import { IOfferDocument } from "../../model/IOfferDocument";
import jwtDecode from "jwt-decode";
import CardService from "../api/CardService";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/offerReviewYoteamsMessageExtension/config.html")
export default class OfferReviewYoteamsMessageExtension implements IMessagingExtensionMiddlewareProcessor {
  private connectionName = process.env.ConnectionName;
  // private documents: IOfferDocument[];
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
    let memberIDs: string[] = [];
    // Members only available in Teams
    if (context.activity.channelId === 'msteams') {
      const memberResponse = await TeamsInfo.getPagedMembers(context, 60, '');      
      memberResponse.members.forEach((m) => {
        memberIDs.push(m.id!);
      });
    }
    log(context.activity.value);
    if (query.commandId === 'offerReviewYoteamsMessageExtension') {
      let documents: IOfferDocument[] = [];
      if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
        const controller = new GraphSearchService();
        documents = await controller.getFiles(tokenResponse.token);        
      }
      documents.forEach((doc) => {
        const card = CardFactory.adaptiveCard(CardService.reviewCardUA(doc, memberIDs));
        const preview = {
          contentType: "application/vnd.microsoft.card.thumbnail",
          content: {
            title: doc.name,
            text: doc.description,
            images: [
              {
                url: `https://${process.env.PUBLIC_HOSTNAME}/assets/icon.png`
              }
            ]             
          }
        };
        attachments.push({ contentType: card.contentType, content: card.content, preview: preview });
      });
    }
    
    if (query.commandId === 'offerPublishYoteamsMessageExtension') {
      let documents: IOfferDocument[] = [];
      if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
        const controller = new GraphSearchService();
        documents = await controller.getFiles(tokenResponse.token);        
      }
      documents.forEach((doc) => {
        const card = CardFactory.adaptiveCard(CardService.publishCardUA(doc, memberIDs));
        const preview = {
          contentType: "application/vnd.microsoft.card.thumbnail",
          content: {
            title: doc.name,
            text: doc.description,
            images: [
              {
                url: `https://${process.env.PUBLIC_HOSTNAME}/assets/icon.png`
              }
            ]             
          }
        };
        attachments.push({ contentType: card.contentType, content: card.content, preview: preview });
      });
    }

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

  public async onActionExecute(context: TurnContext): Promise<AdaptiveCardResponseBody> {
    console.log(context);
    console.log(context.activity.value);
    const doc: IOfferDocument = context.activity.value.action.data.doc as IOfferDocument;
    if (typeof doc.modified === 'string') {
      doc.modified = new Date(doc.modified);
    }
    const adapter: any = context.adapter;
    const magicCode = (context.activity.value.state && Number.isInteger(Number(context.activity.value.state))) ? context.activity.value.state : '';        
    const tokenResponse = await adapter.getUserToken(context, this.connectionName, magicCode);
    if (!tokenResponse || !tokenResponse.token) {
      // There is no token, so the user has not signed in yet.            
      return Promise.reject();
    }
    // Get user's Email from the token (as the context.activity only offers display name)
    const decoded: { [key: string]: any; } = jwtDecode(tokenResponse.token) as { [key: string]: any; };
    const controller = new GraphSearchService();
    switch (context.activity.value.action.verb) {
      case 'review':
        await controller.reviewItem(tokenResponse.token, doc.id, decoded.upn!);
        break;
      case 'alreadyreviewed':
        let currentDoc: IOfferDocument;
        currentDoc = await controller.getItem(tokenResponse.token, doc.id)
          .catch(e => { 
            console.log(e);
            return doc; // Use card's doc instead
        });
        if (typeof currentDoc.reviewer !== 'undefined') {
          return Promise.resolve({
            statusCode: StatusCodes.OK,
            type: 'application/vnd.microsoft.card.adaptive',
            value: CardService.reviewedCardUA(currentDoc)
          });
        }
        else {
          // let memberIDs: string[] = [];
          // const memberResponse = await TeamsInfo.getPagedMembers(context, 60, '');      
          // memberResponse.members.forEach((m) => {
          //   memberIDs.push(m.id!);
          // });
          return Promise.resolve({
            statusCode: StatusCodes.OK,
            type: 'application/vnd.microsoft.card.adaptive',
            value: CardService.reviewCardUA(currentDoc, context.activity.value.action.data.userIds)
          });
        }
        break;
      case 'publish':
        const publishedFileUrl = await controller.publishItem(tokenResponse.token, doc.name, doc.id, doc.fileId!, decoded.upn!);
        let finalDoc = await controller.getItem(tokenResponse.token, doc.id)
          .catch(e => { 
            console.log(e);
            return doc; // Use card's doc instead
        });
        finalDoc.url = publishedFileUrl;
        const card = CardService.publishedCardUA(finalDoc);
        return Promise.resolve({
          statusCode: StatusCodes.OK,
          type: 'application/vnd.microsoft.card.adaptive',
          value: card
        });
        break;
      case 'alreadypublished':
        let publishDoc: IOfferDocument;
        publishDoc = await controller.getItem(tokenResponse.token, doc.id)
          .catch(e => { 
            console.log(e);
            return doc; // Use card's doc instead
        });
        if (typeof publishDoc.publisher !== 'undefined') {
          return Promise.resolve({
            statusCode: StatusCodes.OK,
            type: 'application/vnd.microsoft.card.adaptive',
            value: CardService.publishedCardUA(publishDoc)
          });
        }
        else {
          return Promise.resolve({
            statusCode: StatusCodes.OK,
            type: 'application/vnd.microsoft.card.adaptive',
            value: CardService.publishCardUA(publishDoc, context.activity.value.action.data.userIds)
          });
        }
        break;
    }
    const card = CardService.reviewedCardUA(doc);
    return Promise.resolve({
      statusCode: StatusCodes.OK,
      type: 'application/vnd.microsoft.card.adaptive',
      value: card
    });
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
