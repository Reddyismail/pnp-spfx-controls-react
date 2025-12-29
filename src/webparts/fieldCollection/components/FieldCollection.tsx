import * as React from 'react';
import styles from './FieldCollection.module.scss';
import type { IFieldCollectionProps } from './IFieldCollectionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { FieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-controls-react/lib/FieldCollectionData';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
// import { sp } from "@pnp/sp/presets/all";

interface IFeedbackItem {
  Title: string;
  Name: string;
  FeedBack: string;
  Email: string;
  Id?: number;
}

interface IState {
  collectionData: IFeedbackItem[];
  panelOpen: boolean; // Track panel open state
}

export default class FieldCollection extends React.Component<IFieldCollectionProps, IState> {

  constructor(props: IFieldCollectionProps) {
    super(props);
    this.state = {
      collectionData: [],
      panelOpen: true // open panel by default
    };

  }

  public componentDidMount() {
    void this.getListData();
  }

  private async getListData() {
    try {
      const response = await this.props.sp.web.lists.getByTitle('FeedBack').items.select(
        "Id",
        "Title",
        "Name",
        "FeedBack",
        "EmailDetails/EMail"
      ).expand("EmailDetails")();

      const data: IFeedbackItem[] = response.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Name: item.Name,
        FeedBack: item.FeedBack,
        Email: item.EmailDetails?.EMail || ''
      }));

      this.setState({ collectionData: data });
    } catch (error) {
      console.error("Error fetching list data:", error);
    }
  }

  private async saveListData(items: IFeedbackItem[]) {
    try {
      for (const item of items) {
        if (item.Id) {
          console.log("Updating item with ID:", item.Id);
          await this.props.sp.web.lists.getByTitle('FeedBack').items.getById(item.Id).update({
            Title: item.Title,
            Name: item.Name,
            FeedBack: item.FeedBack,
            EmaiId: item.Email
          });
        } else {
          console.log("Updating item with ID:", item.Id);
          await this.props.sp.web.lists.getByTitle('FeedBack').items.add({
            Title: item.Title,
            Name: item.Name,
            FeedBack: item.FeedBack,
            EmaiId: item.Email
          });
        }
      }
      void this.getListData();
    } catch (error) {
      console.error("Error saving list data:", error);
    }
  }

  public render(): React.ReactElement<IFieldCollectionProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    return (
      <section className={`${styles.fieldCollection} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img
            alt=""
            src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')}
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>

        <div>
          <FieldCollectionData
            key="FieldCollectionData"
            label="Fields Collection"
            manageBtnLabel="Manage"
            value={this.state.collectionData}
            panelHeader="Manage va  lues"
            itemsPerPage={5}
            context={this.props.context}
            // panelOpen={this.state.panelOpen} // Automatically open panel
            fields={[
              { id: "Title", title: "Title", type: CustomCollectionFieldType.string, required: true },
              { id: "Name", title: "Name", type: CustomCollectionFieldType.string },
              { id: "FeedBack", title: "Feedback", type: CustomCollectionFieldType.string },
              { id: "Email", title: "Email", type: CustomCollectionFieldType.string }
            ]}
            onChanged={async (value: IFeedbackItem[]) => {
              this.setState({ collectionData: value });
              await this.saveListData(value);
            }}
            executeFiltering={(searchFilter: string, item: any) => {
              return item.Title?.toLowerCase().includes(searchFilter.toLowerCase());
            }}
          />
        </div>
      </section>
    );
  }
}
