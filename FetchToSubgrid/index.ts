import { IInputs, IOutputs } from './generated/ManifestTypes';
import { FetchSubgrid, IFetchSubgridProps } from './components/FetchSubgrid';
import * as React from 'react';
import { setContext } from './services/crmService';

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
      const fetchXml = context.parameters.fetchXmlProperty.raw;
      const pageSize = context.parameters.defaultPageSize.raw;

      try {
        const userParameters = JSON.parse(fetchXml ?? '');
        const props: IFetchSubgridProps = {
          fetchXml: userParameters.FetchXml || context.parameters.defaultFetchXmlProperty.raw,
          defaultPageSize: Number(userParameters.PageSize) || pageSize,
          newButtonVisibility: userParameters.NewButtonVIsiblitiy ||
           context.parameters.newButtonVisibility.raw,
          deleteButtonVisibility: userParameters.DeleteButtonVisiblity ||
           context.parameters.deleteButtonVisibility.raw,
          userParameters,
        };
        return React.createElement(FetchSubgrid, props);
      }
      catch {
        const props: IFetchSubgridProps = {
          fetchXml: fetchXml ?? context.parameters.defaultFetchXmlProperty.raw,
          defaultPageSize: pageSize,
          newButtonVisibility: context.parameters.newButtonVisibility.raw,
          deleteButtonVisibility: context.parameters.deleteButtonVisibility.raw,
          userParameters: {},
        };
        return React.createElement(FetchSubgrid, props);
      }
    }

    public getOutputs(): IOutputs {
      return {};
    }

    public destroy(): void {
    }
}
