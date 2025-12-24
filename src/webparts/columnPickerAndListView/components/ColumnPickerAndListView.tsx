import * as React from 'react';
import styles from './ColumnPickerAndListView.module.scss';
import type { IColumnPickerAndListViewProps } from './IColumnPickerAndListViewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ListPicker } from "@pnp/spfx-controls-react/lib/ListPicker";
import { Dropdown } from '@fluentui/react';

interface IState {
  selectedListId: string[] | null;
  items: any[];
  collectionData: any[];
  allColumns: any[];       // all columns of selected list
  selectedColumns: string[]; // columns user chooses to fetch
}
export default class ColumnPickerAndListView extends React.Component<IColumnPickerAndListViewProps, IState> {

  constructor(props: IColumnPickerAndListViewProps) {
    super(props);
    this.state = {
      selectedListId: null,
      items: [],
      collectionData: [],
      allColumns: [],     // all columns of selected list
      selectedColumns: [] // columns user chooses to fetch
    };

    this.onListPickerChange = this.onListPickerChange.bind(this);
  }
  private onListPickerChange(lists: string[]): void {
    if (!lists || lists.length === 0) {
      console.log("No list selected yet");
      return;
    }
    //const listId = lists[0]; // because multiSelect=false
    this.setState({ collectionData: [...this.state.collectionData, lists] });

    this.setState({ selectedListId: lists }, () => {
      void this.loadListItems();
    });
  }

  private async loadListItems() {
    if (!this.state.selectedListId) {
      return;
    }

    try {
      const items: any[] = await this.props.sp.web.lists.getById(this.state.selectedListId).items();
      this.setState({ items });
      this.setState({ items: items })
      // const userFields = items.filter(f => !f.Hidden && f.FromBaseType === false)
      const columnKeys = Object.keys(items[0]).filter(
        k => !k.startsWith("odata") && !k.startsWith("OData__")
      );
      this.setState({ allColumns: columnKeys });
      console.log("List items loaded:", columnKeys);
    } catch (error) {
      console.error("Error loading list items:", error);
    }
  }

  //This column picker and list view render method
  private async loadItemsBySelectedColumns(): Promise<void> {
    const { selectedColumns, selectedListId } = this.state;

    if (!selectedListId || selectedColumns.length === 0) {
      this.setState({ items: [] });
      return;
    }

    const items = await this.props.sp.web.lists
      .getById(selectedListId)
      .items
      .select(...selectedColumns)() // ⭐ dynamic column();

    this.setState({ items: items });
  }

  public render(): React.ReactElement<IColumnPickerAndListViewProps> {
    console.log(this.state.items)
    const {
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      context
    } = this.props;


    return (
      <section className={`${styles.columnPickerAndListView} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
        </div>
        <div>
          <ListPicker context={context}
            label="Select your list(s)"
            placeHolder="Select your list(s)"
            baseTemplate={100}
            // contentTypeId="0x0101"
            includeHidden={false}
            multiSelect={false}
            onSelectionChanged={this.onListPickerChange} />
        </div>
        <div>
          <Dropdown
            placeholder="Select columns"
            label="Select columns to display"
            multiSelect
            options={this.state.allColumns.map(col => ({
              key: col,
              text: col
            }))}
            selectedKeys={this.state.selectedColumns}
            onChange={(_, option) => {
              if (!option) return;

              const selectedKeys = option.selected
                ? [...this.state.selectedColumns, option.key as string]
                : this.state.selectedColumns.filter(
                  key => key !== option.key
                );

              // ✅ Call API AFTER state is updated
              this.setState({ selectedColumns: selectedKeys }, () => {
                void this.loadItemsBySelectedColumns();
              });
            }}
          />

        </div>

      </section>
    );
  }

}
