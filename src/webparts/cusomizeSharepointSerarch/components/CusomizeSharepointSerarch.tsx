/* eslint-disable react/no-children-prop */
import * as React from 'react';
// import styles from './CusomizeSharepointSerarch.module.scss';
import type { ICusomizeSharepointSerarchProps } from './ICusomizeSharepointSerarchProps';
import SearchComponent from '../../../component/SearchComponent';
// import { SearchContextProvider } from '../../../context/SearchContext';
// import { escape } from '@microsoft/sp-lodash-subset';

export default class CusomizeSharepointSerarch extends React.Component<ICusomizeSharepointSerarchProps, {}> {
  public render(): React.ReactElement<ICusomizeSharepointSerarchProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;

    return (
      <>
        {this.props.context && <SearchComponent context={this.props.context} listName={'AsiaSourcingDMS'} />}

      </>

    );
  }
}
