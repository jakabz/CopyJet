declare interface ICopyJetSetupWebPartStrings {
  Title: string;
  StepSelect: string;
  StepOptions: string;
  StepSummary: string;
  StepExport: string;
  Next: string;
  Back: string;

  Loading: string;
  LoadError: string;
  DiscoverPartialError: string;
  WholeSite: string;
  CategoryList: string;
  CategoryGroup: string;
  CategorySiteField: string;
  CategoryContentType: string;
  SelectedOf: string;
  Columns: string;
  Views: string;
  ItemCount: string;
  UnsupportedLIST_TEMPLATE_UNSUPPORTED: string;
  UnsupportedVIEW_TYPE_UNSUPPORTED: string;
  UnsupportedOther: string;
  SelectionSummary: string;
  NothingSelected: string;

  NameLabel: string;
  NameRequired: string;
  DefaultTemplateName: string;
  DescriptionLabel: string;
  ContentLaterInfo: string;

  SummaryIntro: string;
  KindGroup: string;
  KindSiteField: string;
  KindContentType: string;
  KindList: string;
  KindListField: string;
  KindView: string;
  CreateTemplate: string;

  ExportRunning: string;
  ExportDone: string;
  ExportFailed: string;
  ExportStopped: string;
  Stop: string;
  Download: string;
  TemplateInvalid: string;
  MissingTitle: string;
  AddMissingAndRebuild: string;
  NewTemplate: string;

  LogTitle: string;
  LogAll: string;
  LogInfo: string;
  LogWarn: string;
  LogError: string;
  LogEmpty: string;
  LogExportCsv: string;
  LogExportJson: string;
}

declare module 'CopyJetSetupWebPartStrings' {
  const strings: ICopyJetSetupWebPartStrings;
  export = strings;
}
