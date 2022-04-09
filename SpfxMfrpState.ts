import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
export interface ISpfxMfrpState {
orderItems: IDropdownOption[];
customerItems: IDropdownOption[];
productItems: IDropdownOption[];
customerId: string;
productId: string;
orderId: any;
productType: string;
date: Date;
unitPrice: string;
numberOfUnits: string;
saleValue: string;
hideOrderId: boolean;
formateddate:string;
}

