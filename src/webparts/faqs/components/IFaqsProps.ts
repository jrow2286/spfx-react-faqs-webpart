export interface IFaqsProps {
  title: string;
  subTitle: string;
  error: string;
  categories: IFaqCategory[];
  collapseCategories: boolean;
  collapseAnswers: boolean;
}

export interface IFaqCategory {
  title: string;
  faqs: IFaq[];
}

export interface IFaq {
  question: string;
  answer: string;
}