import{ IOffer } from '../../model/IOffer';

export interface IOfferCreationFormProps {
  createOffer: (offer: IOffer) => void;
}