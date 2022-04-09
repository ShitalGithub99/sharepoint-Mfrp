import * as React from 'react';//import react library from package.json
import styles from './SpfxMfrp.module.scss';//styles 
import { ISpfxMfrpProps } from './ISpfxMfrpProps';//importing interface
import { ISpfxMfrpState } from './SpfxMfrpState';// importing all the available states from SpfxMfrpState.ts
import { TextField } from 'office-ui-fabric-react'; // importing textfield from office ui
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';//import stack and stacktokens, provides padding between rows
import { Dropdown, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';// import dropdown from office ui
import { PrimaryButton } from 'office-ui-fabric-react';//import primary button from office ui
import { sp } from "@pnp/sp";//provides a fluent api for working with sharepoint REST
import "@pnp/sp/webs";//Webs serve as a container for lists, features, sub-webs, and all of the entity types.
import "@pnp/sp/lists";//list operations--add, get etc.
import "@pnp/sp/views";//defines the columns, ordering, and other details we see when we look at a list.
import "@pnp/sp/items";//get, add items from the list.
import { IItemAddResult } from '@pnp/sp/items';

//Declaration of dropdown array for required fields
var customerItems: IDropdownOption[] = [];
var productItems: IDropdownOption[] = [];
var orderItems: IDropdownOption[] = [];

//Declaration of required variables
var productName = '';
var customerName = '';
var orderId;

export default class SpfxMfrp extends React.Component<ISpfxMfrpProps, ISpfxMfrpState> {
  //Passing props and states to the constructor
  public constructor(props: ISpfxMfrpProps, state: ISpfxMfrpState) {
    super(props); //super is used to call the constructor of its parent class
    this.state = {
      //Assigned States from ISpfxMfrpState.tsx
      orderItems: [],
      orderId: '',
      customerItems: [],
      customerId: '',
      productItems: [],
      productType: '',
      date: new Date(),
      unitPrice: '',
      productId: '',
      numberOfUnits: '',
      saleValue: '',
      hideOrderId: false,
      formateddate:''
    };
    //Binding Functions
    // bind is used to send data as an arguments to the function  of class based components
    this.handleChange = this.handleChange.bind(this);
    this.handleChangeCust = this.handleChangeCust.bind(this);
    this.autoPopulate = this.autoPopulate.bind(this);
    this.handleUnitChange = this.handleUnitChange.bind(this);
    this.addToOrderList = this.addToOrderList.bind(this);
    this.handleChangeOrd = this.handleChangeOrd.bind(this);
    this.editOrderList = this.editOrderList.bind(this);
    this.deleteItem = this.deleteItem.bind(this);
    this.resetOrderList = this.resetOrderList.bind(this);
  }
  //initializing fields based on product name for autopopulate
  private handleChange(event): void {
    productName = event.key;
    this.setState({ numberOfUnits: '' }),
      this.setState({ saleValue: '' }),
      this.autoPopulate();
  }


  //Getting Vendor id when customer name is selected
  //async ensures that the function returns a promise. Other values are wrapped in a resolved promise automatically.
  //we used try/catch when we need to catch the error inside an event handler.
  //await makes javascript wait until that promise settles and returns its result.
  private async handleChangeCust(event): Promise<void> {
    try {
      customerName = event.key;
      let items = await sp.web.lists.getByTitle("Vendors").items.getPaged();//Getting the list items from the customers list
      for (let i = 0; i < items.results.length; i++) {
        if (items.results[i].VendorName == customerName) {
          this.setState({ customerId: items.results[i].VID});

        }
      }
    } catch (error) {
      console.error(error);
    }

  }

  //getting order details when OrderID is selected
  private async handleChangeOrd(event): Promise<void> {
    try {
      orderId = event.key;
      let items = await sp.web.lists.getByTitle("Orders").items.getPaged();//Getting the list items from the Orders list
      for (let i = 0; i < items.results.length; i++) {

        if (items.results[i].Order_x0020_ID == orderId) {
          this.setState({orderId: orderId});
          this.setState({ productId: items.results[i].ProductID});
          this.setState({ customerId: items.results[i].VendorID });
          this.setState({ numberOfUnits: items.results[i].UnitsSold});
          this.setState({ unitPrice: items.results[i].UnitPrice});
          this.setState({ saleValue: items.results[i].SaleValue});
        }
      }
      let customeritems = await sp.web.lists.getByTitle("Vendors").items.getPaged();
      for (let i = 0; i < customeritems.results.length; i++) {
        debugger;
        if (customeritems.results[i].VID == this.state.customerId) {
          customerName = customeritems.results[i].VendorName;
        }
      }
      let productitems = await sp.web.lists.getByTitle("Product").items.getPaged();
      for (let i = 0; i < productitems.results.length; i++) {
        if (productitems.results[i].PID == this.state.productId) {
          productName = productitems.results[i].ProductName;
          this.setState({ productType: productitems.results[i].ProductType });
          this.setState({ formateddate:String(productitems.results[i].ProductExpiryDate ).substring(0,10)});
        }
      }
    } catch (error) {
      console.error(error);
    }

  }

  //Autopopulating fields  when product name is selected
  private async autoPopulate(): Promise<void> {
    try {
      let items = await sp.web.lists.getByTitle("Product").items.getPaged();

      for (let i = 0; i < items.results.length; i++) {
        if (items.results[i].ProductName == productName) {
          this.setState({ productId: items.results[i].PID});
          this.setState({ productType: items.results[i].ProductType});
          this.setState({ unitPrice: items.results[i].ProductUnitPrice});
          this.setState({ formateddate:String(items.results[i].ProductExpiryDate ).substring(0,10)});
        }
      }
    } catch (error) {
      console.error(error);
    }
  }

  //Calculating Sale Value
  private handleUnitChange = (event) => {
    this.setState({ numberOfUnits: event.target.value.toString() });
    var units: number = parseInt(event.target.value);
    var unitPrice: number = parseInt(this.state.unitPrice);
    var calculate = units * unitPrice;
    this.setState({ saleValue: calculate.toString() });
    return event;
  }
  //Adding order to orders list
  private async addToOrderList(event): Promise<void> {
    try {
      debugger;
      var unitvalid = this.state.numberOfUnits;
      if (customerName == "" || productName == "") {
        alert("Please Select Vendor Name From Dropdown" + "\n" + "Please Select Product Name From Dropdown");
      }
      else if (this.state.numberOfUnits == "" || this.state.numberOfUnits <= "0") {
        alert("Please Enter Number Of Units");
      }
      else if (Number(unitvalid) !== parseInt(unitvalid) && Number(unitvalid) % 1 !== 0) {
        alert("Please enter No. of Units as integer value");
      }
      else {
        debugger;
        let item = await sp.web.lists.getByTitle('Orders').items.add({
          VendorID: this.state.customerId.toString(),
          ProductID: this.state.productId.toString(),
          UnitsSold: +this.state.numberOfUnits,
          UnitPrice: +this.state.unitPrice,
          SaleValue: +this.state.saleValue,
          ProductName:productName,
          VendorName:customerName
          //Title: "title"
        });
        alert("Order is added successfully in the list.");
        this.resetOrderList();
      }
    } catch (error) {
      console.error(error);
    }
  }



  // //Editing orders
  private async editOrderList(event): Promise<void> {

    try {
      this.setState({ hideOrderId: true });
      var unitvalid = this.state.numberOfUnits;

      if (orderId == undefined) {
        alert("Select an order ID for editing");
      }
      else if (this.state.numberOfUnits == "" && this.state.customerId == "") {
        this.setState({ hideOrderId: true });
      }
      else {
        if (this.state.numberOfUnits == "" || this.state.numberOfUnits == "0") {
          alert("Please Enter Valid Number Of Units or You have entered zero units");
        }
        else if (Number(unitvalid) !== parseInt(unitvalid) && Number(unitvalid) % 1 !== 0) {
          alert("Please enter No. of Units as integer value");
        }
        else {
          let id: any = orderId;//from input
          id = id.replace(/[^\d]/g, '');  //Extracting only integer.
          id = parseInt(id, 10);         //Trimming Leading Zeros.
          if (id > 0) {
            let editOrderList = await sp.web.lists.getByTitle("Orders").items.getById(id).update({
            VendorID: this.state.customerId.toString(),
            ProductID: this.state.productId.toString(),
            UnitsSold: +this.state.numberOfUnits,
            UnitPrice: +this.state.unitPrice,
            SaleValue: +this.state.saleValue,
            ProductName:productName,
            VendorName:customerName
            //Title: "title"

          });
          alert("ORDERID -" + orderId + " is updated successfully in the list.");
          this.resetOrderList();
        }
        }
      }
    }
    catch (error) {
      console.error(error);
    }
  }

  //Deleting order
  private async deleteItem(event): Promise<void> {
    try {
      this.setState({ hideOrderId: true });

      if (orderId == undefined) {
        alert("Select an order ID to delete");
      }
      let id: any = orderId;//from input
      id = id.replace(/[^\d]/g, '');  //Extracting only integer.
      id = parseInt(id, 10);         //Trimming Leading Zeros.

      if (id > 0) {
        let deleteItem = await sp.web.lists.getByTitle("Orders").items.getById(id).delete();
        console.log(deleteItem);
        alert(`Item ID: ${orderId} deleted successfully!`);
        this.resetOrderList();
      }
      else {
        alert(`Please enter a valid Order id.`);
      
      }
    } catch (error) {
      console.log(error);
    }
  }
  //reset order list
  private resetOrderList(): void {
    customerName = '';
    productName = '';
    orderId = '';

    this.setState({
      date: new Date(),
      productType: '',
      unitPrice: '',
      numberOfUnits: '',
      saleValue: ''
    });
  }


  public async componentDidMount(): Promise<void> {
    // get all the items from a sharepoint list
    var reacthandler = this;
    sp.web.lists.getByTitle("Vendors").items.select('VendorName').get().then((data) => {
      for (var k in data) {
        customerItems.push({ key: data[k].VendorName, text: data[k].VendorName });
      }
      reacthandler.setState({ customerItems });
      console.log(customerItems);
      return customerItems;
    });

    sp.web.lists.getByTitle("Product").items.select('ProductName').get().then((data) => {
      for (var k in data) {
        productItems.push({ key: data[k].ProductName, text: data[k].ProductName });
      }
      reacthandler.setState({ productItems });
      console.log(productItems);
      return productItems;
    });

    sp.web.lists.getByTitle("Orders").items.select('Order_x0020_ID').get().then((data) => {
      for (var k in data) {
        orderItems.push({ key: data[k].Order_x0020_ID, text: data[k].Order_x0020_ID});
      }
      reacthandler.setState({ orderItems: orderItems });
      console.log(orderItems);
      return orderItems;
    });
  }



  public render(): React.ReactElement<ISpfxMfrpProps> {
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 452 }
    };
    const stackTokens: IStackTokens = { childrenGap: 12 };
    return (
      <div className={styles.spfxMfrp}>
        <div className={styles.container}>
          <div>
            <div className={styles.header}>
              <header>
                <img src={require('./image/im6.jpg')} alt="Scar" width="65" height="40" />
                <h1 className={styles.heading}>Sales Order Form</h1>
              </header></div>
            <div className={styles.row}>
              <div className={styles.column}>
                <Stack tokens={stackTokens}>
                  <Dropdown
                    placeholder="Select Vendor Name"
                    label="Vendor Name"
                    selectedKey={customerName}
                    options={this.state.customerItems}
                    styles={dropdownStyles}
                    onChanged={this.handleChangeCust}
                  />
                </Stack>
                <Stack tokens={stackTokens}>
                  <Dropdown
                    placeholder="Select Product Name"
                    label="Product Name"
                    selectedKey={productName}
                    options={this.state.productItems}
                    styles={dropdownStyles}
                    onChanged={this.handleChange}
                  />
                </Stack>
                <TextField required={true}
                  placeholder="Product Type will Be Entered Automatically"
                  label="Product Type"
                  value={this.state.productType}
                  readOnly={true}
                  onChange={event => {
                    this.setState({ productType: this.state.productType });
                  }}
                /><br />
                <TextField label="Product Expiry Date"
                  placeholder=""
                  value={this.state.formateddate}
                  readOnly={true}
                />
                <TextField required={true}
                  placeholder="Product Unit Price"
                  label="Product Unit Price"
                  type="number"
                  value={this.state.unitPrice}
                  readOnly={true}
                  onChange={e => { this.setState({ unitPrice: this.state.unitPrice }); }}
                />
                <TextField required={true}
                  placeholder="Enter Number Of Units "
                  label="Number of units"
                  type="number"
                  value={this.state.numberOfUnits}
                  onChange={this.handleUnitChange}
                //onChange={e=>{this.setState({ numberOfUnits: 'e' })}}
                />
                <div className="HeadText">
                  <TextField required={true}
                    placeholder="Total Sale Value"
                    label="Sale Value"
                    type="number"
                    value={this.state.saleValue}
                    readOnly={true}
                    onChange={e => { this.setState({ saleValue: this.state.saleValue }); }}
                  />
                  <Stack tokens={stackTokens}>
                    {
                      this.state.hideOrderId ?
                        <Dropdown required={true}
                          placeholder="Select an Order ID to Edit or Delete"
                          label="Order ID"
                          selectedKey={orderId}
                          options={this.state.orderItems}
                          styles={dropdownStyles}
                          onChanged={this.handleChangeOrd} />
                        : null
                    }
                  </Stack><br></br>
                  <br></br>
                </div>
              </div>
              <div className={styles.column}>
                <hr />
                <PrimaryButton className={styles.add} onClick={this.addToOrderList} >ADD</PrimaryButton>&nbsp;&nbsp; {/*non breaking space*/}
                <PrimaryButton  className={styles.edit} onClick={this.editOrderList} >EDIT</PrimaryButton>&nbsp;&nbsp;
                <PrimaryButton className={styles.delete} onClick={this.deleteItem} >DELETE</PrimaryButton>&nbsp;&nbsp;
                <PrimaryButton className={styles.reset} onClick={this.resetOrderList}>RESET</PrimaryButton>&nbsp;&nbsp;
                <hr />
              </div>
            </div>
            <div className={styles.footer}>
              <footer>
                <section>
                  <h3 className={styles.footing}> &copy;Created By Shital@2022</h3>
                </section>
              </footer>
            </div>
          </div>
        </div>
      </div>
    );
  }
}