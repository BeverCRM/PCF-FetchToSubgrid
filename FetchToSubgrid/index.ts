import { IInputs, IOutputs } from './generated/ManifestTypes';
import { AppWrapper, IAppWrapperProps } from './components/AppWrapper';
import * as React from 'react';
import { getProps, setContext } from './services/crmService';

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
