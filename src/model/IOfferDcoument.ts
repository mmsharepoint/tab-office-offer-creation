export interface IOfferDocument {
  name: string;
  description: string;
  modified: Date;
  author: string;
  id: string;
  url: string;
  reviewer?: string;
  reviewedOn?: Date;
}