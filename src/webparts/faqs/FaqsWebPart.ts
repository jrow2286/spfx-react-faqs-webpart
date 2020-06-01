import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import * as strings from 'FaqsWebPartStrings';
import Faqs from './components/Faqs';
import { IFaqsProps, IFaqCategory } from './components/IFaqsProps';
import '@pnp/polyfill-ie11';
import { sp } from '@pnp/sp';
import * as _ from 'lodash';

export interface IFaqsWebPartProps {
  title: string;
  subTitle: string;
  listName: string;
  questionField: string;
  answerField: string;
  categoryField: string;
  sortField: string;
  collapseCategories: boolean;
  collapseAnswers: boolean;
}

export default class FaqsWebPart extends BaseClientSideWebPart<IFaqsWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;

  private listFields: IPropertyPaneDropdownOption[];
  private listFieldsDropdownDisabled: boolean = true;

  private errorText: string = "";
  private categories: IFaqCategory[] = [];

  protected onInit(): Promise<void> {
    return new Promise<void>(
      (
        resolve: () => void, 
        reject: (error?: any) => void
      ) => {
        sp.setup({
          spfxContext: this.context,
          sp: {
            headers: {
              'Accept': 'application/json; odata=nometadata'
            }
          }
        });
        
        this.loadFaqs();
        resolve();
      }
    );
  }

  public render(): void {
    
    const element: React.ReactElement<IFaqsProps > = React.createElement(
      Faqs,
      {
        title: this.properties.title,
        subTitle: this.properties.subTitle,
        error: this.errorText,
        categories: this.categories,
        collapseCategories: this.properties.collapseCategories,
        collapseAnswers: this.properties.collapseAnswers
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private loadFaqs(): void {
    this.errorText = '';
    this.categories = [];

    if(!this.properties.listName) {
      this.errorText = 'Select the list to use for FAQs!';
      this.render();
      return;
    }

    if(!this.properties.questionField) {
      this.errorText = 'Select the questions field for FAQs!';
      this.render();
      return;
    }
    
    if(!this.properties.answerField) {
      this.errorText = 'Select the answers field for FAQs!';
      this.render();
      return;
    }

    let filter = 'ID ne 0';
    let selectFields = ['ID', this.properties.questionField, this.properties.answerField];
    let categoriesHash = [];

    if(this.properties.categoryField) {
      selectFields.push(this.properties.categoryField);
    }
    sp.web.lists.getByTitle(this.properties.listName)
    .items.filter(filter).select(selectFields.join(',')).orderBy(this.properties.sortField).get()
    .then((items: any[]): void => {
      let self = this;
      items.forEach((item) => {
        let thisCategory = self.properties.categoryField ? item[self.properties.categoryField] : 'General';
        let thisCategoryIndex = categoriesHash.indexOf(thisCategory);
        if(thisCategoryIndex < 0) {
          thisCategoryIndex = categoriesHash.length;
          categoriesHash.push(thisCategory);
          self.categories.push({title: thisCategory, faqs: []});
        }
        self.categories[thisCategoryIndex].faqs.push({
          question: item[self.properties.questionField],
          answer: item[self.properties.answerField]
        });

      });
      
      this.render();

    }, (error: any): void => {
      error.response.json().then((json) => {
        // error message where?
        this.errorText = json['odata.error'].message.value;
        this.render();
      });
    });
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        // Get custom lists from web
        sp.web.lists.select('Title').get()
        .then((data): void => {
          resolve(
            data.map(item => {
              return {
                key: item.Title,
                text: item.Title
              };
            })
          );
        });
      }
    );
  }

  private loadListFields(): Promise<IPropertyPaneDropdownOption[]> {
    if(!this.properties.listName) {
      return Promise.resolve();
    }
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        // Get custom lists from web
        sp.web.lists.getByTitle(this.properties.listName).fields
        .filter('ReadOnlyFIeld eq false').select('InternalName,Title').get()
        .then((data): void => {
          data.unshift({
            InternalName: "",
            Title: ""
          });
          resolve(
            data.map(item => {
              return {
                key: item.InternalName,
                text: item.Title
              };
            })
          );
        });
      }
    );
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropdownDisabled = !this.lists;
    if(this.lists) {
      return;
    }
    //this.context.statusRenderer.displayLoadingIndicator(this.domElement, strings.UpdatingText);
    this.loadLists().then((listOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
      this.lists = listOptions;
      this.listsDropdownDisabled = false;
      this.context.propertyPane.refresh();
      return this.loadListFields();
    })
    .then((listFieldOptions: IPropertyPaneDropdownOption[]): void => {
      this.listFields = listFieldOptions;
      this.listFieldsDropdownDisabled = this.listFields ? false : true;
      this.context.propertyPane.refresh();
      //this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    
    if(propertyPath === 'listName' && newValue) {
      // reset field properties
      let previousVal: string;
      previousVal = this.properties.questionField;
      this.onPropertyPaneFieldChanged('questionField', previousVal, this.properties.questionField = '');
      
      previousVal = this.properties.answerField;
      this.onPropertyPaneFieldChanged('answerField', previousVal, this.properties.answerField = '');
      
      previousVal = this.properties.categoryField;
      this.onPropertyPaneFieldChanged('categoryField', previousVal, this.properties.categoryField = '');
      
      previousVal = this.properties.sortField;
      this.onPropertyPaneFieldChanged('sortField', previousVal, this.properties.sortField = '');

      this.listFieldsDropdownDisabled = true;

      this.context.propertyPane.refresh();
      //this.context.statusRenderer.displayLoadingIndicator(this.domElement, strings.UpdatingText);

      this.loadListFields()
      .then((listFieldOptions: IPropertyPaneDropdownOption[]): void => {
        this.listFields = listFieldOptions;
        this.listFieldsDropdownDisabled = this.listFields ? false : true;
        this.context.propertyPane.refresh();
        //this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      });
    }
    this.loadFaqs();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.WebpartSettingsGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  value: this.properties.title
                }),
                PropertyPaneTextField('subTitle', {
                  label: strings.SubTitleFieldLabel,
                  value: this.properties.subTitle
                }),
                PropertyPaneToggle('collapseCategories', {
                  label: strings.CollapseCategoriesFieldLabel,
                  checked: this.properties.collapseCategories
                }),
                PropertyPaneToggle('collapseAnswers', {
                  label: strings.CollapseAnswersFieldLabel,
                  checked: this.properties.collapseAnswers
                })
              ]
            },
            {
              groupName: strings.ListSettingsGroupName,
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                }),
                PropertyPaneDropdown('questionField', {
                  label: strings.QuestionFieldLabel,
                  options: this.listFields,
                  selectedKey: this.properties.questionField,
                  disabled: this.listFieldsDropdownDisabled
                }),
                PropertyPaneDropdown('answerField', {
                  label: strings.AnswerFieldLabel,
                  options: this.listFields,
                  selectedKey: this.properties.answerField,
                  disabled: this.listFieldsDropdownDisabled
                }),
                PropertyPaneDropdown('categoryField', {
                  label: strings.CategoryFieldLabel,
                  options: this.listFields,
                  selectedKey: this.properties.categoryField,
                  disabled: this.listFieldsDropdownDisabled
                }),
                PropertyPaneDropdown('sortField', {
                  label: strings.SortFieldLabel,
                  options: this.listFields,
                  selectedKey: this.properties.sortField,
                  disabled: this.listFieldsDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
