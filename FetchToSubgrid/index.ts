import * as React from 'react';
import { IInputs, IOutputs } from './generated/ManifestTypes';
import { IAppWrapperProps, IDataverseService } from './utilities/types';
import { AppWrapper } from './components/AppWrapper';
import { DataverseService } from './services/dataverseService';

export class FetchToSubgrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private _dataverseService: IDataverseService;

  public init(context: ComponentFramework.Context<IInputs>): void {
    this._dataverseService = new DataverseService(context);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    context.mode.trackContainerResize(true);
    const props: IAppWrapperProps = this._dataverseService.getProps();
    return React.createElement(AppWrapper, props);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
  }
}
