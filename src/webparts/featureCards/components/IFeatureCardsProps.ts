import { IFeatureCardItem } from '../FeatureCardsWebPart';

export interface IFeatureCardsProps {
  tituloSeccion: string;
  tarjetas: IFeatureCardItem[];
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
}