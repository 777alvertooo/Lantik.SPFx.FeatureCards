import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IconPicker } from '@pnp/spfx-controls-react/lib/IconPicker';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';

import FeatureCards from './components/FeatureCards';
import { IFeatureCardsProps } from './components/IFeatureCardsProps';

export interface IFeatureCardItem {
  texto: string;
  subtexto: string;
  icono: string;
  enlace?: string;
}

export interface IFeatureCardsWebPartProps {
  tituloSeccion: string;
  tarjetas: IFeatureCardItem[];
}

export default class FeatureCardsWebPart extends BaseClientSideWebPart<IFeatureCardsWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IFeatureCardsProps> = React.createElement(
      FeatureCards,
      {
        tituloSeccion: this.properties.tituloSeccion !== undefined ? this.properties.tituloSeccion : 'Valores',
        tarjetas: this.properties.tarjetas || [],
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configuración del bloque de tarjetas"
          },
          groups: [
            {
              groupName: "Contenido",
              groupFields: [
                PropertyPaneTextField('tituloSeccion', {
                  label: "Título superior",
                  value: "Valores"
                }),
                PropertyFieldCollectionData("tarjetas", {
                  key: "tarjetas",
                  label: "Tarjetas",
                  panelHeader: "Configurar tarjetas de valores",
                  manageBtnLabel: "Gestionar tarjetas",
                  value: this.properties.tarjetas,
                  enableSorting: true,
                  fields: [
                    {
                      id: "texto",
                      title: "Texto",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "subtexto",
                      title: "Subtexto",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "icono",
                      title: "Icono",
                      type: CustomCollectionFieldType.custom,
                      required: true,
                      onCustomRender: (field, value, onUpdate, item, itemId) => {
                        return React.createElement(IconPicker, {
                          buttonLabel: value || "Elegir icono",
                          currentIcon: value,
                          renderOption: "dialog",
                          onSave: (iconName: string) => {
                            onUpdate(field.id, iconName);
                          },
                          key: itemId
                        });
                      }
                    },
                    {
                      id: "enlace",
                      title: "Enlace de redirección",
                      type: CustomCollectionFieldType.custom,
                      required: false,
                      onCustomRender: (field, value, onUpdate, item, itemId) => {
                        return React.createElement(FilePicker, {
                          buttonLabel: value ? "Cambiar destino" : "Elegir destino",
                          context: this.context as any,
                          hideLocalUploadTab: true,
                          hideStockImages: true,
                          onSave: (filePickerResult: IFilePickerResult[]) => {
                            if (filePickerResult && filePickerResult.length > 0) {
                              onUpdate(field.id, filePickerResult[0].fileAbsoluteUrl);
                            }
                          },
                          onChange: (filePickerResult: IFilePickerResult[]) => {
                            if (filePickerResult && filePickerResult.length > 0) {
                              onUpdate(field.id, filePickerResult[0].fileAbsoluteUrl);
                            }
                          },
                          key: itemId
                        });
                      }
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}