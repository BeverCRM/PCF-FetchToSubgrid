import { IInputs, IOutputs } from './generated/ManifestTypes';
import { FetchSubgrid, IFetchSubgridProps } from './components/FetchSubgrid';
import * as React from 'react';
import { setContext } from './services/crmService';
import { parseString } from './utilities/utilities';

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

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
      const userParameters = context.parameters.userParameters.raw ||
      `{"NewVIsiblitiy":"true", "DeleteVisiblity":"true", "FetchXml":""}`;

      const props: IFetchSubgridProps = {
        numberOfRows: context.parameters.numberOfRows.raw,
        fetchXml: context.parameters.fetchXmlProperty.raw ??
        parseString(userParameters).FetchXml,
        userParameters,
      };

      return React.createElement(FetchSubgrid, props);
    }

    public getOutputs(): IOutputs {
      return {};
    }

    public destroy(): void {
    }
}
