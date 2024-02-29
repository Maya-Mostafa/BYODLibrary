import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IByodLibraryProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;

  context: WebPartContext;
  targetAudience: any;
  siteUrl: string;
  listName: string;
  displayState: string;
  color: string;
  openInNewTab: boolean;
  showDivider: boolean;
  sectionTitle: string;
  isCollapsible: boolean;
  iconAlignment: string;
  iconPicker: any;
  thumbnail: any;
  customImgPicker: any;

  groupBy: boolean;
  groupByField: string;
  sectionDescription: string;
  enableSearch: boolean;
  searchPlaceholder: string;
  enableTargetAudience: boolean;
  showBasedOnTargetAudience: boolean;
}
