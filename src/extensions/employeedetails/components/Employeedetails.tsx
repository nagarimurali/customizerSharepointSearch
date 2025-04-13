import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { Label } from '@fluentui/react/lib/Label';
import styles from './Employeedetails.module.scss';
import FormContainer from './CustomForm/FormContainer';
//import styles from '../../dmsTranslatedocSet/components/DmsTranslatedocSet.module.scss';

export interface IEmployeedetailsProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'Employeedetails';

export default class Employeedetails extends React.Component<IEmployeedetailsProps, {}> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: Employeedetails mounted');

  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: Employeedetails unmounted');
  }

  public render(): React.ReactElement<{}> {
    const { displayMode } = this.props;
    let titleHeader = "";
    switch (displayMode) {
      case FormDisplayMode.New:
        titleHeader = "New";
        break;
      case FormDisplayMode.Edit:
        titleHeader = "Edit"
        break;
      case FormDisplayMode.Display:
        titleHeader = "View";
        break;
    }
    return <div className={styles.baselineForms} >
      <div className={styles.header}>
        <Label className={styles.titleHeader}>{titleHeader}</Label>
      </div>
      {displayMode === FormDisplayMode.New &&
     <FormContainer {...this.props} />
      }
      {displayMode === FormDisplayMode.Edit &&
      // 
        <>
         <FormContainer {...this.props} />
        </>
       
      }
      {displayMode === FormDisplayMode.Display &&
        <div>View</div>
      }
    </div>
      ;
  }
}
