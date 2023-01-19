import { IOfferDocument } from "../../model/IOfferDcoument";

export default class CardService {
  public static reviewCard = (doc: IOfferDocument) => {
    return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.4",
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
                          url: `https://${process.env.PUBLIC_HOSTNAME}/assets/icon.png`,
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
          type: "Action.Execute",
          title: "Reviewed",
          verb: "review",
          data: {
            doc: doc
          }
        }
      ]
    }
  }

  public static reviewedCard = (doc: IOfferDocument) => {
    return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.4",
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
                          url: `https://${process.env.PUBLIC_HOSTNAME}/assets/icon.png`,
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
      }
    ]
    }
  }

  public static reviewCardUA = (doc: IOfferDocument, userIds: string[]) => {
    return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.4",
    refresh: {
        action: {
            type: "Action.Execute",
            title: "Refresh",
            verb: "alreadyrevied",
            data: {
              doc: doc
            }
        },
        userIds: userIds
    },

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
                          url: `https://${process.env.PUBLIC_HOSTNAME}/assets/icon.png`,
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
          ]
        }
      }
    ]
    }
  }
};