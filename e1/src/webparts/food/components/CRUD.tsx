
import * as React from 'react';
import styles from './Food.module.scss';
import { IFoodProps } from './IFoodProps';
import { escape } from '@microsoft/sp-lodash-subset';

// import { TextField } from 'office-ui-fabric-react/lib/TextField';
// import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";

// Document what this does
import "@pnp/sp/lists";
// Document what this does
import "@pnp/sp/items";

import { WebPartContext } from "@microsoft/sp-webpart-base";

// For submit
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

// To select person
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

// To Show dishes from Food List. (title field is used)
// import { ListPicker } from '@pnp/spfx-controls-react/lib/ListPicker';
import Select from 'react-select'


import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  


// Bad way of listing constants. Break out to resource file "strings" or something.
const ListNameAttendees: string = "Attendees";
const ListNameFoodMenu: string = "Food menu";
const FieldAttendeesPerson: string = "Attendee";
const FieldFoodMenuString: string = "Title";

export interface Dish{
  value:string;
  label:string;
}

// State for holding items
export interface ICrudStates {
    AvailableDishItems: any;
    UserObject: any;
    SelectedUserName: string;
    SelectedUserId: number;
    SelectedUserEmail: string;
    SelectedDish:string;
  }

export interface ICrudProps {
  description: string;
  context: WebPartContext;
  webURL:string;
}

export default class CRUD extends React.Component<IFoodProps, ICrudStates> {

    public async componentDidMount() {
      console.log("componentDidMount"); 
      await this.fetchDishesDataWithPnP();
    }

    public async fetchDishesDataWithPnP() {
        let web = Web(this.props.webUrl);
        const dishesItems: any[] = await web.lists.getByTitle(ListNameFoodMenu).items.select("*", "Title").get();
        console.log(dishesItems);
        if(dishesItems.length > 0){
          console.log("Found: dishesItems, extracting title and id");
          let xAvailableDishItems = dishesItems.map(({Title, ID}) => ({Title, ID}));
          console.log(xAvailableDishItems);
          this.setState({ AvailableDishItems: xAvailableDishItems });
        }
    }

    public _getPeoplePickerItems = async (items: any[]) => {
    
      if (items.length > 0) {
        console.log("User ID: " + items[0].id);
        this.setState({ UserObject: items[0] });
        this.setState({ SelectedUserEmail: items[0].secondaryText });
        this.setState({ SelectedUserName: items[0].text });
        this.setState({ SelectedUserId: items[0].id });
      }
      else {
        //ID=0;
        this.setState({ SelectedUserName: "" });
        this.setState({ SelectedUserId: 0 });
      }
    }

    private onListPickerChange (lists: string | string[]) {
      console.log("Lists:", lists);
    }

    private SelectedDish(){
      console.log("SelectedDish");
      console.log(this);
    }


    private async SaveDataWithPnP() {
      let web = Web(this.props.webUrl);
      
      await web.lists.getByTitle(ListNameAttendees).items.add({
  
        Title: "Testar användaren först",
        AttendeeId: { 
          results: [ this.state.SelectedUserId ] // allows multiple users
      }
  
      }).then(i => {
        console.log(i);
      });
      alert("Created Successfully With PNP");
    }

    private async SaveDataREST() {
      
      this._addListItem();
    }

    private _getItemEntityType(): Promise<string> {
      const endpoint: string = this.props.context.pageContext.web.absoluteUrl
                                    + `/_api/web/lists/getbytitle('` + ListNameAttendees + `')`
                                    + `?$select=ListItemEntityTypeFullName`
    
      return this.props.context.spHttpClient
          .get(endpoint, SPHttpClient.configurations.v1)
          .then(response => {
            return response.json();
          })
          .then(jsonResponse => {
            return jsonResponse.ListItemEntityTypeFullName;
          }) as Promise<string>;
    }

    private _addListItem(): Promise<SPHttpClientResponse> {
      return this._getItemEntityType()
        .then(spEntityType => {
          const request: any = {};

          console.log("spEntityType");

          console.log(spEntityType);


          request.body = JSON.stringify({
            Title: new Date().toUTCString(),
            AttendeeId: this.state.SelectedUserId,
            FoodChoice: "",
            '@odata.type': spEntityType
          });
    
          console.log(request.body);

          const endpoint: string = this.props.context.pageContext.web.absoluteUrl 
          + `/_api/web/lists/getbytitle('` + ListNameAttendees + `')/items`;
          
          return this.props.context.spHttpClient.post(
            endpoint, SPHttpClient.configurations.v1, request);
        });
    }

    public render(): React.ReactElement<IFoodProps> {
      let handler = this;

      const options = [
        { value: 'Fish', label: 'Fish' },
        { value: 'Meat', label: 'Meat' },
        { value: 'Vegan', label: 'Vegan' },
        { value: 'Vego', label: 'Vego' }
      ]

        return <div>
            <p>Exercise 1!</p>
            <h2>Dishes</h2>
            {/* <Select options={options} className="selectedItem" onChange={() => this.SelectedDish()} /> */}

            <h2>Who are you? </h2>
            <p>Please search on your name and select yourself. Max 1 Person. </p>
            <PeoplePicker
              context={this.props.context as any}
              personSelectionLimit={1}
              required={false}
              onChange={this._getPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
              ensureUser={true}
            />
             {/* <div><PrimaryButton text="Create SaveDataWithPnP" onClick={() =>this.SaveDataWithPnP()}/></div> */}
             <div><PrimaryButton text="Create SPFx REST API" onClick={() =>this.SaveDataREST()}/></div>
        </div>;
    }
}
  
  