/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BuddyAppWebPartStrings';
import BuddyApp from './components/BuddyApp';
import { IBuddyAppProps } from './components/IBuddyAppProps';
import { sp } from "@pnp/sp";

export interface IBuddyAppWebPartProps {
  description: string;
}

export default class BuddyAppWebPart extends BaseClientSideWebPart<IBuddyAppWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
    private _userLoginDetails: {
    name: string;
    department: string;
    email: string;
    workPhone: string;
    mobilePhone: string;
    office: string;
    photoUrl: string;
  } | null = null;


 public async render(): Promise<void> {
    if (!this._userLoginDetails) {
      await this._fetchUserDetails();
    }
        const element: React.ReactElement<IBuddyAppProps> = React.createElement(
      BuddyApp,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userloginDetails: this._userLoginDetails,
        currentUserId: this.context.pageContext.legacyPageContext.userId,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        sp: sp,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context as any
    });
    
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    }).then(() => {
      return this._fetchUserDetails();
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  
  private async _fetchUserDetails(): Promise<void> {
    try {
      const currentUserId = this.context.pageContext.legacyPageContext.userId;
      
      // Ensure sp.web is properly initialized
      if (!(sp as any).web) {
        throw new Error('sp.web is not initialized');
      }
      
      const userDetails = await (sp as any).web.siteUserInfoList.items
        .getById(currentUserId)
        .get();

      this._userLoginDetails = {
        name: userDetails.Title,
        department: userDetails.Department,
        // email: "gautam.raj@credentinfotech.com",
        email: userDetails.EMail,
        workPhone: userDetails.WorkPhone,
        mobilePhone: userDetails.MobilePhone,
        office: userDetails.Office,
        photoUrl:
          (typeof userDetails.Picture === 'string' ? userDetails.Picture : '') ||
          userDetails.Picture?.Url ||
          userDetails.Picture?.URL ||
          userDetails.PictureUrl ||
          userDetails.PictureURL ||
          '',
      };

      console.log("User Details:", this._userLoginDetails);
    } catch (error) {
      console.error("Error fetching user details:", error);
      // Fallback to basic context info
      const user = this.context.pageContext.user;
      this._userLoginDetails = {
        name: user.displayName || 'Unknown',
        department: 'N/A',
        email: user.email || 'N/A',
        workPhone: 'N/A',
        mobilePhone: 'N/A',
        office: 'N/A',
        photoUrl: '',
      };
    }
  }
}
