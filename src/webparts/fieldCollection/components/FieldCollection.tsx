import * as React from 'react';
import styles from './FieldCollection.module.scss';
import type { IFieldCollectionProps } from './IFieldCollectionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { FieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-controls-react/lib/FieldCollectionData';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
// import { sp } from "@pnp/sp/presets/all";

//Animation Related Imports
import { AnimatedDialog } from "@pnp/spfx-controls-react/lib/AnimatedDialog";
import { DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
// import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';

interface IFeedbackItem {
  Title: string;
  Name: string;
  FeedBack: string;
  Email: string;
  Id?: number;

}

interface IState {
  collectionData: IFeedbackItem[];
  oldCollectionData: IFeedbackItem[];
  panelOpen: boolean; // Track panel open state
  showAnimatedDialog?: boolean;
}
interface IDialogContentProps {
  type: number;
  title: string;
  subText: string;
}

interface IModalProps {
  isDarkOverlay: boolean;
}
export default class FieldCollection extends React.Component<IFieldCollectionProps, IState> {

  constructor(props: IFieldCollectionProps) {
    super(props);
    this.state = {
      collectionData: [],
      oldCollectionData: [],
      panelOpen: true,// open panel by default
      showAnimatedDialog: false
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
      if (this.state.collectionData.length > items.length) {
        const oldItems = this.state.collectionData;

        // 🔍 find deleted items
        const deletedItems = oldItems.filter(oldItem =>
          !items.some(items => items.Id === oldItem.Id)
        );
        for (const delItem of deletedItems) {
          await this.props.sp.web.lists.getByTitle('FeedBack').items.getById(delItem.Id).delete();
        }
      } else {
        for (const item of items) {
          if (item.Id) {
            console.log("Updating item with ID:", item.Id);
            await this.props.sp.web.lists.getByTitle('FeedBack').items.getById(item.Id).update({
              Title: item.Title,
              Name: item.Name,
              FeedBack: item.FeedBack,
              EmaiId: item.Email
            });
          }
          else {
            console.log("Updating item with ID:", item.Id);
            await this.props.sp.web.lists.getByTitle('FeedBack').items.add({
              Title: item.Title,
              Name: item.Name,
              FeedBack: item.FeedBack,
              EmaiId: item.Email
            });
          }
        }
      }



      void this.getListData();
      //this for dilogbox auto close
      setTimeout(() => {
        this.setState({ showAnimatedDialog: false });
      })
    } catch (error) {
      console.error("Error saving list data:", error);
    }
  }


  //Animated Dialog Close Handler

  private animatedDialogContentProps: IDialogContentProps = {

    type: DialogType.normal,
    title: 'Data Saved',
    subText: 'Succesfully saved the data to SharePoint list!',
  };

  private animatedModalProps: IModalProps = {
    isDarkOverlay: true
  };
  public render(): React.ReactElement<IFieldCollectionProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;
    //console.log(this.state.collectionData)
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
            panelHeader="Manage values"
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
              this.setState({ collectionData: value }, () => {
                this.setState({ showAnimatedDialog: true });
              });
              await this.saveListData(value);
            }}

            executeFiltering={(searchFilter: string, item: any) => {
              return item.Title?.toLowerCase().includes(searchFilter.toLowerCase());
            }}


          />
          <AnimatedDialog
            hidden={!this.state.showAnimatedDialog}
            onDismiss={() => { this.setState({ showAnimatedDialog: false }); }}
            dialogContentProps={this.animatedDialogContentProps}
            modalProps={this.animatedModalProps}
          >
            <DialogFooter>
              {/* <PrimaryButton onClick={() => { this.setState({ showAnimatedDialog: false }); }} text="Yes" />
              <DefaultButton onClick={() => { this.setState({ showAnimatedDialog: false }); }} text="No" /> */}
            </DialogFooter>
          </AnimatedDialog>
        </div>
      </section>
    );
  }
}
