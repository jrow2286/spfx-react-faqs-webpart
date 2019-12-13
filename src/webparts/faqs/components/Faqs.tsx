import * as React from 'react';
import styles from './Faqs.module.scss';
import { IFaqsProps } from './IFaqsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import classNames from 'classnames';


export default class Faqs extends React.Component<IFaqsProps, {}> {

  public toggleElement(id: string): void {
    let el = document.getElementById(id);
    el.classList.toggle(styles.collapsed);
  }

  public render(): React.ReactElement<IFaqsProps> {
    let categoryClass = classNames(styles.categoryWrap);
    if(this.props.collapseCategories) {
      categoryClass = classNames(styles.categoryWrap, styles.collapsed);
    }
    let faqWrapClass = classNames(styles.faqWrap);
    if(this.props.collapseAnswers) {
      faqWrapClass = classNames(styles.faqWrap, styles.collapsed);
    }
    
    return (
      <div className={ styles.faqs }>
        <div className={ styles.container }>
          {this.props.title &&
            <h1 className={ styles.title }>{escape(this.props.title)}</h1>
          }
          {this.props.subTitle &&
            <div className={ styles.subTitle }>{escape(this.props.subTitle)}</div>
          }
          {this.props.error &&
            <div className={ styles.error }>{escape(this.props.error)}</div>
          }

          {this.props.categories.map((category, catInd) => {
            return (
              <>
                <div id={'category-' + catInd}  className={ categoryClass }>
                  <div className={ styles.category} onClick={this.toggleElement.bind(this, 'category-' + catInd)}>
                    <span className={ styles.categoryTitle }>{escape(category.title)}</span>
                    <Icon className={ classNames(styles.chevron, styles.open) } iconName='ChevronRight' />
                    <Icon className={ classNames(styles.chevron, styles.closed) } iconName='ChevronDown' />
                  </div>
                  <div className={ styles.categoryFaqsWrap }>
                    {category.faqs.map((faq, faqInd) => {
                      return (
                        <>
                          <div id={'answer-' + catInd + '-' + faqInd} className={ faqWrapClass }>
                            <div className={ styles.question } onClick={this.toggleElement.bind(this, 'answer-' + catInd + '-' + faqInd)}>
                              <span>{faq.question}</span>
                              <Icon className={ classNames(styles.chevron, styles.open)  } iconName='ChevronRight' />
                              <Icon className={ classNames(styles.chevron, styles.closed) } iconName='ChevronDown' />
                            </div>
                            <div className={ styles.answer } dangerouslySetInnerHTML={{__html: faq.answer}}></div>
                          </div>
                        </>
                      );
                    })}
                  </div>
                </div>
              </>
            );
          })}
        </div>
      </div>
    );
  }
}
