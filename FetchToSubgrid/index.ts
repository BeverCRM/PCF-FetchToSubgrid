import { IInputs, IOutputs } from './generated/ManifestTypes';
import { FetchSubgrid } from './components/FetchSubgrid';
import * as React from 'react';
import { getProps, setContext } from './services/crmService';

export class FetchToSubgrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private Component: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;

    constructor() { }

    public init(
      context: ComponentFramework.Context<IInputs>,
      notifyOutputChanged: () => void,
    ): void {
      this.notifyOutputChanged = notifyOutputChanged;
      setContext(context);
    }

    public updateView(): React.ReactElement {
      return React.createElement(FetchSubgrid, getProps());
    }

    public getOutputs(): IOutputs {
      return {};
    }

    public destroy(): void {
    }
}
