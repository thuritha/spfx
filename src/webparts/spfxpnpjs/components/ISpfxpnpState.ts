import { IDropdownOption } from "office-ui-fabric-react";
export interface ISpfxpnpState {
    customerNameList : IDropdownOption[];
    productNameList : IDropdownOption[];
    orderIDList: IDropdownOption[];
    CustomerName : any;
    CustomerID : any;
    ProductID : any;
    ProductName : any;
    ProductUnitPrice : any;
    NoofUnits : any;
    UnitPrice : any;
    SaleValue : any;
    OrderStatus : string;
    whichButton : string;
    OrderID: any;

}