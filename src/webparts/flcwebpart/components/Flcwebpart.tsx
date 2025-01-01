import * as React from 'react';
import styles from './Flcwebpart.module.scss';
import type { IFlcwebpartProps } from './IFlcwebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Flcwebpart extends React.Component<IFlcwebpartProps, {}> {
  public render(): React.ReactElement<IFlcwebpartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      // userDisplayName
      textField,
      multiLineTextField,
      linkField,
      dropdownField,
      choiceGroupField,
      sliderField,
      toggleField,
      checkboxField,
      buttonField

    } = this.props;

    return (
      <section className={`${styles.flcwebpart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <div>
            <h3>Welcome to SharePoint Framework!</h3>
          </div>
          <div>{environmentMessage}</div>
          <div className={`${styles.propertyBox}`}>Web part property value: <strong>{escape(description)}</strong></div>
          <div className={`${styles.propertyBox}`}> Text Field: <strong>{escape(textField)}</strong></div>
          <div className={`${styles.propertyBox}`}> Multi-line Text Field: <strong>{escape(multiLineTextField)}</strong></div>
          <div className={`${styles.propertyBox}`}> Link Field: <strong>{escape(linkField)}</strong></div>
          <div className={`${styles.propertyBox}`}> Dropdown Field: <strong>{escape(dropdownField)}</strong></div>
          <div className={`${styles.propertyBox}`}> Checkbox Field: <strong>{escape(checkboxField ? 'You have checked the checkbox field as true' : 'You have checked the checkbox field as false')}</strong></div>
          <div className={`${styles.propertyBox}`}> Choice Group Field: <strong>{escape(choiceGroupField)}</strong></div>
          <div className={`${styles.propertyBox}`}> Slider Field: <strong>{escape(sliderField.toString())}</strong></div>
          <div className={`${styles.propertyBox}`}> Toggle Field: <strong>{escape(toggleField ? 'Toggle is on' : 'Toggle is off')}</strong></div>
          <div className={`${styles.propertyBox}`}> Button Field: <strong>{escape(buttonField)}</strong></div>

        </div>
       
      </section>
    );
  }
}
