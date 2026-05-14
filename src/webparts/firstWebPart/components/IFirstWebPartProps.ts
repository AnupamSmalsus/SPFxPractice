import SearchService from "../../../services/SearchService";

export interface IFirstWebPartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  searchService: SearchService;
  siteUrl: any;
  rootUrl: string;
}
