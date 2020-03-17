import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneTextFieldProps,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NasaApolloMissionViewerWebPart.module.scss';
import * as strings from 'NasaApolloMissionViewerWebPartStrings';

import { IMission } from '../../models';
import { MissionService } from '../../services';

export interface INasaApolloMissionViewerWebPartProps {
  description: string;
  selectedMission: string;
}

export default class NasaApolloMissionViewerWebPart extends BaseClientSideWebPart<INasaApolloMissionViewerWebPartProps> {
  
  private selectedMission: IMission;
  private missionDetailElement: HTMLElement;

  protected onInit(): Promise<void> {
    return new Promise<void>(
      (
        resolve: () => void,
        reject: (error: any) => void
      ): void => {
        this.selectedMission = this._getSelectedMission();
        resolve();
      }
    );
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.nasaApolloMissionViewer}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <div class="apolloMissionDetails"></div>
            </div>
          </div>
        </div>
      </div>`;
    
    this.missionDetailElement = this.domElement.getElementsByClassName("apolloMissionDetails")[0] as HTMLElement;

    if (this.selectedMission) {
      this._renderMissionDetails(this.missionDetailElement, this.selectedMission);
    } else {
      this.missionDetailElement.innerHTML = "";
    }
  }

  private _getSelectedMission(): IMission {
    // Determine the mission ID, default to Apollo 11
    const selectedMissionId: string = (this.properties.selectedMission)
      ? this.properties.selectedMission
      : "AS-506";

    // get the specified mission
    return MissionService.getMission(selectedMissionId);
  }

  private _renderMissionDetails(element: HTMLElement, mission: IMission): void {
    element.innerHTML = `
      <p class="ms-font-m">
        <span class="ms-fontWeight-semibold">Mission: </span>
        ${escape(mission.name)}
      </p>
      <p class="ms-font-m">
        <span class="ms-fontWeight-semibold">Duration: </span>
        ${escape(this._getMissionTimeLine(mission))}
      </p>
      <a href="${mission.wiki_href}" target="_blank" class="${styles.button}">
        <span class="${styles.label}">Learn more about ${escape(mission.name)} on Wikipedia &raquo;</span>
      </a>`;
  }

  private _getMissionTimeLine(mission: IMission): string {
    let missionDate = mission.end_date !== ''
      ? `${mission.launch_date.toString()} - ${mission.end_date.toString()}`
      : `${mission.launch_date.toString()}`;
    return missionDate;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        // <Page 1>
        {
          header: {
            description: "NASA Apollo Mission Viewer Web Part (Voitanos SPFx Course)"
          },
          displayGroupsAsAccordion: true,
          groups: [
            // <General>
            {
              groupName: 'General',
              isCollapsed: false,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('selectedMission', <IPropertyPaneTextFieldProps> {
                  label: 'Apollo Mission to Show',
                  onGetErrorMessage: this._validateMissionCode.bind(this)
                })
              ]
            },
            // </General>
            // <Custom>
            {
              groupName: 'Custom',
              isCollapsed: true,
              groupFields: [
                PropertyPaneLabel('', {
                  text: 'Custom Fields'
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('selectedMission', <IPropertyPaneTextFieldProps> {
                  label: 'Apollo Mission to Show',
                  onGetErrorMessage: this._validateMissionCode.bind(this)
                })
              ]
            }
            // </Custom>
          ]
        },
        // </Page 1>
        // <Page 2>
        {
          header: {
            description: 'About this web part'
          },
          groups: [
            {
              groupFields: [
                PropertyPaneLabel('', {
                  text: 'This is a killer first web part! This \'About\' page is defined in the "getPropertyPaneConfiguration" method of the "NasaApolloMissionViewer.ts" file.'
                })
              ]
            }
          ]
        }
        // </Page 2>

      ]
    };
  }

  private _validateMissionCode(value: string): string {
    const validMissionCodeRegEx = /AS-[2,5][0,1][0-9]/g;
    return value.match(validMissionCodeRegEx)
      ? ''
      : 'Invalid mission code; should be \'AS-###\'.';
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    // update selected mission
    this.selectedMission = this._getSelectedMission();

    // update rendering
    if (this.selectedMission) {
      this._renderMissionDetails(this.missionDetailElement, this.selectedMission);
    } else {
      this.missionDetailElement.innerHTML = "";
    }
  }
}
