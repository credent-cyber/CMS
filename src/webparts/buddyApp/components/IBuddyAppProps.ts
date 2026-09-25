/*eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBuddyAppProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userloginDetails: any;
  currentUserId: any;
  context: WebPartContext;
  siteUrl?: string;
  sp: any;
}
