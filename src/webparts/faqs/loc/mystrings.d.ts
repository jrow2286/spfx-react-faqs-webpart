declare interface IFaqsWebPartStrings {
  TitleFieldLabel: string;
  SubTitleFieldLabel: string;
  CollapseCategoriesFieldLabel: string;
  CollapseAnswersFieldLabel: string;
  WebpartSettingsGroupName: string;
  ListSettingsGroupName: string;
  ListNameFieldLabel: string;
  QuestionFieldLabel: string;
  AnswerFieldLabel: string;
  CategoryFieldLabel: string;
}

declare module 'FaqsWebPartStrings' {
  const strings: IFaqsWebPartStrings;
  export = strings;
}
