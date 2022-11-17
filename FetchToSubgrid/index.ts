import { IInputs, IOutputs } from './generated/ManifestTypes';
import { FetchSubgrid, IFetchShbgridProps } from './Components/FetchSubgrid';
import * as React from 'react';
import FetchService from './Services/FetchService';

export class FetchToSubgrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;

    constructor() { }

    public init(
      context: ComponentFramework.Context<IInputs>,
      notifyOutputChanged: () => void,
    ): void {
      this.notifyOutputChanged = notifyOutputChanged;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
      FetchService.setContext(context);

      const props: IFetchShbgridProps = { inputValue: context.parameters.sampleProperty.raw };
      return React.createElement(
        FetchSubgrid, props,
      );
    }

    public getOutputs(): IOutputs {
      return {};
    }

    public destroy(): void {
    }
}
