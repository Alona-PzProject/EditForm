declare interface IEditFormWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'EditFormWebPartStrings' {
  const strings: IEditFormWebPartStrings;
  export = strings;
}
