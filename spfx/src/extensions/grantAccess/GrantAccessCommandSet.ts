import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { HttpClient, type HttpClientResponse } from '@microsoft/sp-http';

const LOG_SOURCE = 'GrantAccessCommandSet';
const FUNCTION_URL = 'https://accessgrant-f7arb5hhe0h7cvf5.westeurope-01.azurewebsites.net/api/grant-access';

export default class GrantAccessCommandSet extends BaseListViewCommandSet<{}> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized');
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  private _onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    const command: Command = this.tryGetCommand('GRANT_ACCESS');
    if (command) {
      // Show button only when exactly one item is selected
      command.visible = this.context.listView.selectedRows?.length === 1;
    }
    this.raiseOnChange();
  };

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'GRANT_ACCESS':
        Log.info(LOG_SOURCE, 'Grant Access command executed');
        void this._grantAccess();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private async _grantAccess(): Promise<void> {
    const row = this.context.listView.selectedRows![0];
    const siteId = this.context.pageContext.site.id.toString();
    const listId = this.context.pageContext.list!.id.toString();
    const itemId = row.getValueByName('UniqueId');

    try {
      const response: HttpClientResponse = await this.context.httpClient.post(
        FUNCTION_URL,
        HttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ siteId, listId, itemId }),
        }
      );

      const result = await response.json();

      if (response.ok) {
        await Dialog.alert('Access granted successfully');
      } else {
        await Dialog.alert(`Error: ${result.error}`);
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      Log.warn(LOG_SOURCE, `Grant access failed: ${message}`);
      await Dialog.alert(`Error: ${message}`);
    }
  }
}
