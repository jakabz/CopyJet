declare interface ICopyJetInstallWebPartStrings {
  Title: string;
  StepLoad: string;
  StepPreview: string;
  StepInstall: string;
  StepResult: string;
  Next: string;
  Back: string;

  DropHint: string;
  ChooseFile: string;
  Reading: string;
  ZipLater: string;
  InvalidTemplate: string;
  ParseError: string;
  VersionError: string;
  TemplateLabel: string;
  SourceLabel: string;
  CreatedLabel: string;
  ContentsLabel: string;
  PermissionsChecking: string;
  PermissionsMissing: string;
  PermissionsOk: string;

  PreviewLoading: string;
  PreviewFailed: string;
  ColumnInclude: string;
  ColumnKind: string;
  ColumnName: string;
  ColumnStatus: string;
  StatusNew: string;
  StatusSame: string;
  StatusDifferent: string;
  StatusUnsupported: string;
  StatusError: string;
  ExcludedDisabled: string;
  ExcludedDependency: string;
  ExcludedCycle: string;
  ConflictModeLabel: string;
  ModeSkip: string;
  ModeUpdate: string;
  ModeRename: string;
  PreviewCounts: string;
  StartInstall: string;

  Installing: string;
  Stop: string;
  LeaveWarning: string;
  InstallFailed: string;

  ResultTitle: string;
  ResultStopped: string;
  CountCreated: string;
  CountUpdated: string;
  CountSkipped: string;
  CountFailed: string;
  CountBlocked: string;
  CountCancelled: string;
  CreatedLists: string;
  NewInstall: string;

  KindGroup: string;
  KindSiteField: string;
  KindContentType: string;
  KindList: string;
  KindListField: string;
  KindView: string;

  LogTitle: string;
  LogAll: string;
  LogInfo: string;
  LogWarn: string;
  LogError: string;
  LogEmpty: string;
  LogExportCsv: string;
  LogExportJson: string;
}

declare module 'CopyJetInstallWebPartStrings' {
  const strings: ICopyJetInstallWebPartStrings;
  export = strings;
}
