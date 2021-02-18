import * as React from 'react';
import styles from './Spfxpnpjs.module.scss';
import { ISpfxpnpjsProps } from './ISpfxpnpjsProps';
import { escape, times } from '@microsoft/sp-lodash-subset';


import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
//import library

import {
  DefaultButton,
  Dropdown,
  TextField,
  Stack,
  IStackTokens,
  IDropdownStyles,
  PrimaryButton,
  IStyleSet,
  Slider,
} from "office-ui-fabric-react";

import { Label } from "office-ui-fabric-react/lib/Label"; 
import {
  IPivotStyles,
  Pivot,
  PivotItem,
  PivotLinkFormat,
  PivotLinkSize,
} from "office-ui-fabric-react/lib/Pivot"; 
import { sp } from "@pnp/sp/presets/all";  
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ISpfxpnpState } from './ISpfxpnpState';
import { spOperation } from '../Services/spService';

const pivotStyles: Partial<IStyleSet<IPivotStyles>> = {
  link: { width: "20%" },
  linkIsSelected: { width: "20%" },
};
const stackTokens: IStackTokens = { childrenGap: 100 };
const bigVertStack: IStackTokens = { childrenGap: 20 };
const SmallVertStack: IStackTokens = { childrenGap: 20 };
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 250 } };


export default class Spfxpnpjs extends React.Component<ISpfxpnpjsProps, ISpfxpnpState, {}> {
  public _spOps : spOperation;
  constructor(props : ISpfxpnpjsProps) {
    super(props);
    this._spOps = new spOperation();
    this.state = {
      customerNameList: [],
      productNameList:[],
      orderIDList:[],
      CustomerName: "" ,
      CustomerID: "",
      ProductID: "",
      ProductName: "",
      ProductUnitPrice: "",
      ProductType: "",
      NoofUnits:"",
      UnitPrice: "",
      SaleValue: "",
      OrderID: "",
      OrderStatus:"",
      whichButton: "Create"
    };

  }
  public componentDidMount() {
    console.log("Component Did Mount called!!");
    // getCustomerNameList
    this._spOps.getCustomerNameList(this.props.context).then((result: any) => {
      this.setState({ customerNameList: result });
      // console.log(this.state.customerNameList);
    });

    // getProductNameList
    this._spOps.getProductNameList(this.props.context).then((result: any) => {
      this.setState({ productNameList: result });
      // console.log(this.state.productNameList);
    });

    // getOrderList
    this._spOps.getOrderList(this.props.context).then((result: any) => {
      this.setState({ orderIDList: result });
    });
  }
  /**
   * getCustomerName
   * this function is called when a dropdown item is changed
   * To save the CustomerName and id to this.state
   */
  public getCustomerName = (event: any, data: any) => {
    console.log("getCustomerName called!!");
    // console.log(data);
    this.setState({
      CustomerName: data.text,
      CustomerID: data.key,
    });
  }

  /**
   * getProductName
   * this function is called when a dropdown item is changed
   * To save the productName,id,type,date and unit value to this.state
   * and show it in the form automatically
   */
  public getProductName = (event: any, data: any) => {
    console.log("getProductName called!!");
    // console.log(data);
    this._spOps
      .getProductDetails(this.props.context, data)
      .then((result: any) => {
        
        var totalValue: any;
        if (this.state.NoofUnits === "") {
          // Update Total Value when Number of Unit is not zero! otherwise don't update!!
          totalValue = this.state.SaleValue;
        } else if (this.state.NoofUnits === "0") {
          totalValue = this.state.SaleValue;
        } else {
          totalValue =
            result.UnitPrice * this.state.NoofUnits;
        }
        // console.log(results);
        // console.log(result.ProductExpiryDate);
        // console.log(date);
        this.setState({
          ProductName: data.text,
          ProductID: data.key,
          
          
          //ProductUnitPrice: result.Product_x0020_Unit_x0020_Price,
          SaleValue: totalValue,
        });
      });
  }
  /**
   * setNumberofUnits is called when number of units is changed to store the value in state.
   */
  public setNumberofUnits = (event: any, data: any) => {
    console.log("setNumberofUnits called!!");
    // console.log(this.state.ProductUnitPrice, data);
    var numberofUnits: any;
    var totalValue: any;
    if (data === "0") {
      numberofUnits = data;
      totalValue = "";
    } else if (data === "") {
      console.log(
        "setNumberofUnits called -> In ifelse -> data = '' statement!!"
      );
      numberofUnits = data;
      totalValue = "";
    } else if (this.state.ProductUnitPrice === "") {
      console.log(
        "setNumberofUnits called -> In ifelse -> UnitPrice = '' statement!!"
      );
      numberofUnits = parseInt(data);
      totalValue = "";
    } else {
      console.log("setNumberofUnits called -> In else statement!!");
      numberofUnits = parseInt(data);
      var priceofunit: number = parseInt(this.state.ProductUnitPrice);
      totalValue = numberofUnits * priceofunit;
    }
    // console.log(numberofUnits);
    // console.log(priceofunit);
    // console.log(totalValue);
    this.setState({
      NoofUnits: numberofUnits,
      SaleValue: totalValue,
    });
    return;
  }
   /**
   * validateItemAndAdd and upload the new item
   */
  public validateItemAndAdd = () => {
    console.log("validateItemAndAdd called!!");
    let myStateList = [
      this.state.CustomerID,
      this.state.CustomerName,
      this.state.ProductID,
      this.state.ProductName,
     
      this.state.ProductUnitPrice,
      this.state.NoofUnits,
      this.state.SaleValue,
    ];
    console.log(myStateList);
    for (let i = 0; i < myStateList.length; i++) {
      if (myStateList[i] === "") {
        this.setState({ OrderStatus: "Fill all Details!" });
        return;
      }
    }

    console.log("Validate Complete and Uploading Order Details");
    this._spOps
      .createItems(this.props.context, this.state)
      .then((result: string) => {
        this.setState({ OrderStatus: result });
      });
  }
  /**
   * validateItemAndModify
   */
  public validateItemAndModify = () => {
    if (this.state.OrderID === "" || this.state.OrderID === null) {
      this.setState({ OrderStatus: "Enter Order Id" });
      return;
    } else {
      this._spOps.updateItem(this.state).then((status) => {
        this.setState({ OrderStatus: status });
      });
    }
  }
  /**
   * validateAndDelete
   */
  public validateAndDelete = () => {
    if (this.state.OrderID === "" || this.state.OrderID === null) {
      this.setState({ OrderStatus: "Enter Order Id" });
      return;
    } else {
      this._spOps.deleteItem(this.state.OrderID).then((response) => {
        this.setState({ OrderStatus: response });
      });
    }
  }
  /**
   * getOrderDetailsToUpdate is called when order id field is changed to get the item details.
   */
  public getOrderDetailsToUpdate = (event: any, data: any) => {
    // Valid Order Id -> Not empty -> not zero
    console.log("getOrderDetailsToUpdate called!");
    if (data === "") {
      console.log("Empty data");
      return;
    }
    // Now get the Order List details from rest and call setstate to change the state
    this._spOps.getUpdateitem(this.props.context, data).then((results) => {
      var result = results.value[0];
      // console.log(result);
      // find customer name
      var customerId = result.Customer_x0020_IDId;
      var customerName;
      this.state.customerNameList.forEach((item) => {
        var flag = false;
        if (item.key === customerId && flag === false) {
          customerName = item.text;
          flag = true;
        }
      });

      var productId = result.ProductID;
      var productName;
      this.state.productNameList.forEach((item) => {
        var flag = false;
        if (item.key === productId && flag === false) {
          productName = item.text;
          flag = true;
        }
      });
       // Now find the product details to fill
       var data1 = { key: result.ProductID, text: productName };
       this.getProductName({}, data1);
 
       this.setState({
         OrderID: data.key,
         CustomerName: customerName,
         CustomerID: customerId,
         NoofUnits: result.UnitsSold,
       });
     });
   }
   /**
   * controlTabButton is called when tabs are changed to store which tab is active.
   */
  public controlTabButton = (data: any) => {
    console.log("Tab Changed");
    console.log(data);
    if (data.props.itemKey === "1") {
      // Add tab clicked
      // reset the tab and setstate for button
      this.setState({ whichButton: "Create" });
    } else if (data.props.itemKey === "2") {
      this.setState({ whichButton: "Update" });
    } else if (data.props.itemKey === "3") {
      this.setState({ whichButton: "Delete" });
    }
  }
  /**
   * resetForm
   */
  public resetForm = () => {
    // Will reset the state of disable text field - call setstate to change state
    // Will clear text for active text field -
    console.log("resetForm called!!");
    this.setState({
      OrderID: null,
      CustomerName: "",
      CustomerID: null,
      ProductID: null,
      ProductName: "",
      ProductUnitPrice: "",
      
     
      NoofUnits: "",
      SaleValue: "",
      OrderStatus: "Reset Done!!",
    });
    this.componentDidMount();
  }
  /**
   * renderButton is used to show active tab's button - eg: for ADD tab button should be SAVE button
   */
  public renderButton = () => {
    if (this.state.whichButton === "Create") {
      return (
        <PrimaryButton
          text="SAVE"
          onClick={this.validateItemAndAdd}
        ></PrimaryButton>
      );
    } else if (this.state.whichButton === "Update") {
      return (
        <PrimaryButton
          text="MODIFY"
          onClick={this.validateItemAndModify}
        ></PrimaryButton>
      );
    } else if (this.state.whichButton === "Delete") {
      return (
        <PrimaryButton
          text="DELETE"
          onClick={this.validateAndDelete}
          // onClick={() =>
          //   this._spOps
          //     .deleteItem(this.props.context, this.state.orderId)
          //     .then((status) => {
          //       this.setState({ status: status });
          //     })
          // }
        ></PrimaryButton>
      );
    }
  }

  public render(): React.ReactElement<ISpfxpnpjsProps> {
    return (
      <div className={ styles.spfxpnpjs }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Book Your Car Now!!</span>
              <br></br>
              <span className={ styles.subTitle }>Limited Stocks</span>
            </div>
          </div>
          <hr/>
          <div>
          <Pivot
              styles={pivotStyles}
              aria-label="Large Link Size Pivot Example"
              linkSize={PivotLinkSize.large}
              linkFormat={PivotLinkFormat.tabs}
              onLinkClick={this.controlTabButton}
            >
              <PivotItem headerText="ADD" itemKey="1">
                <Label>To add Orders Fill the form below.</Label>
                <div className={styles.emptyheight}></div>
              </PivotItem>
              <br/>
              <PivotItem headerText="EDIT" itemKey="2">
                <Label>Select Order Id below.</Label>
                <Dropdown
                  required
                  selectedKey={this.state.OrderID}
                  prefix="Order Id"
                  options={this.state.orderIDList}
                  onChange={this.getOrderDetailsToUpdate}
                ></Dropdown>
              </PivotItem>
              <br/>
              <PivotItem headerText="DELETE" itemKey="3">
                <Label>Select Order Id below.</Label>
                <Dropdown
                  required
                  selectedKey={this.state.OrderID}
                  prefix="Order Id"
                  options={this.state.orderIDList}
                  onChange={this.getOrderDetailsToUpdate}
                ></Dropdown>
              </PivotItem>
            </Pivot>

            <Dropdown
              required
              selectedKey={[this.state.CustomerID]}
              id="forReset1"
              label="Enter Customer Name"
              options={this.state.customerNameList}
              onChange={this.getCustomerName}
            ></Dropdown>
            <Dropdown
                  required
                  selectedKey={[this.state.ProductID]}
                  id="forReset2"
                  label="Enter Product Name"
                  options={this.state.productNameList}
                  onChange={this.getProductName}
                  //styles={dropdownStyles}
                ></Dropdown>

            <Stack horizontal wrap tokens={stackTokens}>
              <Stack tokens={bigVertStack}>
                {/*<Dropdown
                  required
                  selectedKey={[this.state.ProductID]}
                  id="forReset2"
                  label="Enter Product Name"
                  options={this.state.productNameList}
                  onChange={this.getProductName}
                  //styles={dropdownStyles}
                ></Dropdown>*/}
                 <TextField
                  id="forReset3"
                  label="Number of Units"
                  type="number"
                  min={1}
                  required
                  value={this.state.NoofUnits}
                  onChange={this.setNumberofUnits}
                /> 
                {/*<Slider
                  label="No of Units"
                  min={0}
                  max={50}
                  step={1}
                  onChanged={this.setNumberofUnits}
                  value={this.state.NoofUnits}
                />*/}
              </Stack>
              <Stack tokens={SmallVertStack}>
              <TextField
                  label="Product Type"
                  //disabled
                  placeholder={this.state.ProductType}
                />
                {/*<TextField
                  label="Product Expiry Date"
                  disabled
                  placeholder={
                    this.state.ProductExpiryDate === ""
                      ? ""
                      : new Date(this.state.ProductExpiryDate).toDateString()
                  }
                />*/}
                <TextField
                  label="Product Unit Price"
                  //readonly
                  placeholder={this.state.ProductUnitPrice}
                />
              </Stack>
            </Stack>

            <div>
              <Label>Total Sales Price</Label>
              <TextField
                ariaLabel="disabled Product Sales Price"
                readOnly
                prefix="Rs. "
                placeholder={this.state.SaleValue}
              />
            </div>
            <div className={ styles.emptyheight }>{this.state.OrderStatus}</div>
            <div className={ styles.emptyheight }></div>
            <Stack horizontal tokens={stackTokens}>
              {this.renderButton()}
              <DefaultButton
                text="CLEAR"
                onClick={() => this.resetForm()}
              ></DefaultButton>
            </Stack>
            <div className={styles.emptyheight}></div>
          </div>
        </div>
      </div>
    );
  }
}
