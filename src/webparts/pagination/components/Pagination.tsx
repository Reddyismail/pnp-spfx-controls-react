import * as React from 'react';
import styles from './Pagination.module.scss';
import type { IPaginationProps } from './IPaginationProps';
import { escape } from '@microsoft/sp-lodash-subset';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import 'bootstrap/dist/css/bootstrap.min.css';

import { Spinner, SpinnerSize } from '@fluentui/react';
import { Pagination } from "@pnp/spfx-controls-react/lib/pagination";

interface IPaginationState {
  data: any[];
  isLoading: boolean;
  currentPage: number;
  itemsPerPage: number;
  totalItems: number; // Optional, we can estimate or fetch dynamically
}

export default class Paginations extends React.Component<IPaginationProps, IPaginationState> {

  constructor(props: IPaginationProps) {
    super(props);

    this.state = {
      data: [],
      isLoading: true,
      currentPage: 1,
      itemsPerPage: 10,
      totalItems: 0
    };
  }

  public componentDidMount(): void {
    void this.loadPage(this.state.currentPage);
    void this.loadTotalItemsCount();
  }

  // -----------------------------
  // GET TOTAL ITEMS COUNT (OPTIONAL)
  // -----------------------------
  private async loadTotalItemsCount(): Promise<void> {
    try {
      const list = await this.props.sp.web.lists.getByTitle("CustomerData_3k").select("ItemCount")();
      this.setState({ totalItems: list.ItemCount });
      const lists = await this.props.sp.web.lists.getByTitle("CustomerData_3k").select("ItemCount")();
      console.log(lists.ItemCount); // total items in the list
    } catch (err) {
      console.error("Error fetching total items count", err);
    }
  }

  // -----------------------------
  // LOAD ANY PAGE
  // -----------------------------
  private async loadPage(page: number): Promise<void> {
    try {
      this.setState({ isLoading: true });

      const skip = (page - 1) * this.state.itemsPerPage;
      console.log("Skip value:", skip);
      const items = await this.props.sp.web.lists
        .getByTitle("CustomerData_3k")
        .items
        .select("Title", "field_1", "field_2", "field_3", "field_4", "field_5")
        .top(this.state.itemsPerPage)
        .skip(skip)();

      this.setState({
        data: items,
        currentPage: page,
        isLoading: false
      });

    } catch (error) {
      console.error("Error fetching page data", error);
      this.setState({ isLoading: false });
    }
  }

  // -----------------------------
  // HANDLE PAGE CHANGE
  // -----------------------------
  private onPageChange = (page: number): void => {
    void this.loadPage(page);
  };

  // -----------------------------
  // RENDER
  // -----------------------------
  public render(): React.ReactElement<IPaginationProps> {
    console.log("Rendering Pagination component with state:", this.state.itemsPerPage);
    console.log("items", this.state.data);
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.pagination} ${hasTeamsContext ? styles.teams : ''}`}>

        {/* Header */}
        <div className={styles.welcome}>
          <img
            alt=""
            src={isDarkTheme
              ? require('../assets/welcome-dark.png')
              : require('../assets/welcome-light.png')}
            className={styles.welcomeImage}
          />
          <h2>Welcome, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part property value: <strong>{escape(description)}</strong>
          </div>
        </div>

        {/* Table */}
        <div className="mt-4">
          <table className="table table-bordered table-striped">
            <thead className="table-dark">
              <tr>
                <th>ID</th>
                <th>Name</th>
                <th>Email</th>
                <th>Address</th>
                <th>Salary</th>
                <th>Emp Code</th>
              </tr>
            </thead>

            <tbody>
              {this.state.isLoading ? (
                <tr>
                  <td colSpan={6} style={{ textAlign: 'center' }}>
                    <Spinner size={SpinnerSize.large} label="Loading data..." />
                  </td>
                </tr>
              ) : (
                this.state.data.map((item, index) => (
                  <tr key={index}>
                    <td>{item.Title}</td>
                    <td>{item.field_1}</td>
                    <td>{item.field_2}</td>
                    <td>{item.field_3}</td>
                    <td>{item.field_4}</td>
                    <td>{item.field_5}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>

          {/* Pagination */}
          {!this.state.isLoading && this.state.totalItems > 0 && (
            <Pagination
              currentPage={this.state.currentPage}
              totalPages={Math.ceil(this.state.totalItems / this.state.itemsPerPage)}
              onChange={this.onPageChange}
              limiter={3}
            />
          )}
        </div>
      </section>
    );
  }
}
