import './@types/string.extension';
import * as React from 'react';
import { IInputs, IOutputs } from './generated/ManifestTypes';
import { IAppWrapperProps, IDataverseService } from './@types/types';
import { AppWrapper } from './components/AppWrapper';
import { DataverseService } from './services/dataverseService';

export class FetchToSubgrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private _dataverseService: IDataverseService;

  public init(context: ComponentFramework.Context<IInputs>): void {
    context.mode.trackContainerResize(true);
    this._dataverseService = new DataverseService(context);
  }

  public updateView(): React.ReactElement {
    const props: IAppWrapperProps = this._dataverseService.getProps();
    return React.createElement(AppWrapper, props);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
  }
}
