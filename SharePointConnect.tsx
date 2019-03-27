import React, {Component} from 'react';
import { Button, ButtonType, DatePicker, TextField, Dropdown, dropDownOptions } from 'office-ui-fabric-react';

export interface ISPConnectState {
  SP_ListArray: SPConnectValues[];
}

export interface SPConnectValues {
  name: string;
}

export default class SharePointConnect extends Component<{}, ISPConnectState> {
  private ListDropDownOptions: dropDownOptions[] = [];

  constructor(props: {}, state: ISPConnectState) {
    super(props);

    this.state = {
      SP_ListArray: []
    };
  }

  public componentDidMount() {
    this.getSharePointLists();
  }

  public getSharePointLists() {
    var self = this;
    var url = "https://binaryrepublik516.sharepoint.com/sites/UMC/_api/web/lists?$filter=Hidden eq false";

    fetch(url, {
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }
    })
      .then((res) => res.json())
      .then(
        (result) => {
          console.log(result.value);

          var listTitle = result.value.map((items) => { return items.Title; });
          self.LoadDropDownValues(listTitle);
        },
        (error) => {
          console.log(error);
        }
      );
  }

  private LoadDropDownValues(data): any {
    this.ListDropDownOptions = [];
    data.map((items) => {
      this.ListDropDownOptions.push({ key: items, text: items });
    });
  }

  render() {
    return (
      <div id="parentDiv">
        <Button buttonType={ButtonType.primary}>Get And Load Lists</Button>
      </div>
    );
  }
}