import * as React from 'react';
import { IInputs, IOutputs } from './generated/ManifestTypes';
import { IAppWrapperProps } from './utilities/types';
import { AppWrapper } from './components/AppWrapper';
import { getProps, setContext } from './services/dataverseService';

export class FetchToSubgrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  public init(context: ComponentFramework.Context<IInputs>): void {
    setContext(context);
  }

  public updateView(): React.ReactElement {
    const props: IAppWrapperProps = { ...getProps() };
    return React.createElement(AppWrapper, props);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
  }
}
