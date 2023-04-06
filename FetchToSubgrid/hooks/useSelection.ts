import { IObjectWithKey, Selection } from '@fluentui/react';
import * as React from 'react';

export const useSelection = () => {
  const [selectedRecordIds, setSelectedRecordIds] = React.useState<string[]>([]);

  const selection = new Selection({
    onSelectionChanged: () => {
      const currentSelection: IObjectWithKey[] = selection.getSelection();
      const recordIds = currentSelection.map<string>((row: any) => row.id);
      setSelectedRecordIds(recordIds);
    },
  });

  return {
    selection,
    selectedRecordIds,
  };
};
