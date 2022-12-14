import { IInputs, IOutputs } from './generated/ManifestTypes';
import { FetchSubgrid, IFetchSubgridProps } from './Components/FetchSubgrid';
import * as React from 'react';
import CrmService from './Services/CrmService';

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
      CrmService.setContext(context);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
      const props: IFetchSubgridProps = {
        numberOfRows: context.parameters.defaultNumberOfRows.raw,
        fetchXml: context.parameters.fetchXmlProperty.raw ??
          context.parameters.defaultFetchXmlProperty.raw,
      };

      return React.createElement(FetchSubgrid, props);
    }

    public getOutputs(): IOutputs {
      return {};
    }

    public destroy(): void {
    }
}
