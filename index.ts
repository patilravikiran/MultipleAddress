import {IInputs, IOutputs} from "./generated/ManifestTypes";

enum AsciiKeyCode {
	Enter = 13,
	Up = 38,
	Down = 40
}

export class MultipleAddresses implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	//PCF variables
	private _context: ComponentFramework.Context<IInputs>;
	private notifyOutputChanged: () => void;
	private _container: HTMLDivElement;

	//HTML elements and events
	private _inputElement: HTMLInputElement;
	private _listDiv: HTMLDivElement;
	private _currentFocus: number;
	private _inputKeyPress: EventListenerOrEventListenerObject;
	private _listItemClick: EventListenerOrEventListenerObject;
	
	//Address variables
    private _addressname: string;
    private _addressline1: string;
    private _addressline2: string;
    private _addressline3: string;
    private _city: string;
    private _state: string;
    private _postalcode: string;
    private _country: string;
    
    //Empty constructor
	constructor()
	{	}

    /**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		//PCF variables initializations
		this._context = context;
		this.notifyOutputChanged = notifyOutputChanged;
		this._container = container;

		//HTML events
		this._currentFocus = -1;
        this._inputKeyPress = this.inputkeyPress.bind(this);
        this._listItemClick = this.listItemClick.bind(this);

		//Adding search input form control
		this._inputElement = document.createElement("input");
		this._inputElement.setAttribute("type", "text");
		this._inputElement.setAttribute("id", "MultipleAddressSearch");
		this._inputElement.setAttribute("class", "addresssearchinput");
		this._inputElement.setAttribute("placeholder", "Look for Company Address");
		this._inputElement.setAttribute("autocomplete", "off");
        this._inputElement.addEventListener("keyup", this._inputKeyPress);
        this._container.appendChild(this._inputElement);

		//Adding list DIV which will contain the parent company addresses
		this._listDiv = document.createElement("div");
		this._listDiv.setAttribute("id", "address-list");
		this._listDiv.setAttribute("class", "address-list-items");
		this._container.appendChild(this._listDiv);
		this._listDiv.hidden = true;
    }    

	/**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		//returning address data back to framework to update address fields against contact.
		return {
            addressname: this._addressname,
            addressline1: this._addressline1,
			addressline2: this._addressline2,
			addressline3: this._addressline3,
			city: this._city,
			state: this._state,
			postalcode: this._postalcode,
			country: this._country
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Deleting all the created HTML controls
		delete this._inputElement;
        delete this._container;
	}

	/** 
	 * Called when the address item is clicked from the list div. Used to update contact address.
	 * @param event HTML event
     */
	private listItemClick(event: Event): void {
        let currentHTMLElement = event.target as HTMLDivElement;
		this._inputElement.value = '';
		this.clearSearchResults();

		//Updating contact address fields
		if(currentHTMLElement !== null){
			let addressGuid = currentHTMLElement.id;
			if(addressGuid !== null){
				this.searchSelectedAddress(addressGuid);
			}
		}		
    }
	
	/** 
	 * Used to search selected address record in the system and update contact address.
	 * @param addressGuid Guid of selected address record
     */
    public async searchSelectedAddress(addressGuid: string) {
        await this._context.webAPI.retrieveRecord("customeraddress", addressGuid, '?$select=name,line1,line2,line3,city,stateorprovince,postalcode,country').then( addressRecord => {
			if(addressRecord.customeraddressid){
				this._addressname = addressRecord.name;
				this._addressname = addressRecord.name;
				this._addressline1 = addressRecord.line1;
				this._addressline2 = addressRecord.line2;
				this._addressline3 = addressRecord.line3;
				this._city = addressRecord.city;
				this._state = addressRecord.stateorprovince;
				this._postalcode = addressRecord.postalcode;                
				this._country = addressRecord.country;
				
				this.notifyOutputChanged();
			}
		});
    }
	
	/** 
	 * Called when any of keyboard event is captured on input element. Used to set/reset focus and populate address records
	 * @param event HTML event
     */
	private inputkeyPress(event: Event): void {
		let keyboardEvent = event as KeyboardEvent;

		if (keyboardEvent.keyCode == AsciiKeyCode.Down) { 
			this._currentFocus++;
			this.setActiveListItem();
		} else if (keyboardEvent.keyCode == AsciiKeyCode.Up) { 
			this._currentFocus--;
			this.setActiveListItem();
		} else if (keyboardEvent.keyCode == AsciiKeyCode.Enter) { 
			var listItem = this._listDiv.childNodes[this._currentFocus] as HTMLDivElement;
			listItem.click();
		}
		else {
			//populating address list
			this.populateAddressList();
		}
	}

	/** 
	 * Called when any of keyboard event excluding up, down and enter keys is captured on input element. Used to populate address list
     */
	private populateAddressList(): void {
		//Receiving parent company Guid from framework
		let companyGuid: string = "";
		if (typeof (this._context.parameters.companyguid) !== "undefined" &&
            typeof (this._context.parameters.companyguid.raw) !== "undefined" && this._context.parameters.companyguid.raw != null) {
				companyGuid = this._context.parameters.companyguid.raw;
        }		

		if(companyGuid !== ""){
            //forming query string
            let queryString: string = "?$select=name,line1,line2,line3,city,stateorprovince,postalcode,country&$filter=_parentid_value eq " + companyGuid + " and addressnumber ne 2";
            this.populateAddressRecords(queryString);
        }
    }
	
	/** 
	 * Used to populate address records on to input list
     */
    private async populateAddressRecords(queryString: string){        
        this.clearSearchResults();
		console.log("In populateAddressRecords");
        //Querying for address records
        await this._context.webAPI.retrieveMultipleRecords("customeraddress", queryString).then( addressList => {
            if(addressList.entities.length > 0){				
				console.log("addressList ", addressList);
				// Looping through each address record
                for (var i = 0; i < addressList.entities.length; i++) {
                    let listItemDiv = document.createElement("DIV");
					
					//forming complete address
					var completeAddress = "";
                    completeAddress = (addressList.entities[i].name !== null ? addressList.entities[i].name + ", " : "");
                    completeAddress = completeAddress + (addressList.entities[i].line1 !== null ? addressList.entities[i].line1 + ", " : "");
                    completeAddress = completeAddress + (addressList.entities[i].line2 !== null ? addressList.entities[i].line2 + ", " : "");
                    completeAddress = completeAddress + (addressList.entities[i].line3 !== null ? addressList.entities[i].line3 + ", " : "");
                    completeAddress = completeAddress + (addressList.entities[i].city !== null ? addressList.entities[i].city + ", " : "");
                    completeAddress = completeAddress + (addressList.entities[i].stateorprovince !== null ? addressList.entities[i].stateorprovince + ", " : "");
                    completeAddress = completeAddress + (addressList.entities[i].postalcode !== null ? addressList.entities[i].postalcode + ", " : "");
                    completeAddress = completeAddress + (addressList.entities[i].country !== null ? addressList.entities[i].country : "");

					//adding list item
                    listItemDiv.innerHTML = completeAddress;
                    listItemDiv.setAttribute("class", "address-list-item");
                    listItemDiv.id = addressList.entities[i].customeraddressid;
                    listItemDiv.addEventListener("click", this._listItemClick);                    
					this._listDiv.appendChild(listItemDiv);
				}
				this._listDiv.hidden = false;

				//let _hrNewAddress = document.createElement("hr");
				//_hrNewAddress.setAttribute("class", "hrNewAddress");
				//this._listDiv.appendChild(_hrNewAddress);

				let _btnNewAddress = document.createElement("button");
				_btnNewAddress.setAttribute("class","addressnewbutton");
				_btnNewAddress.innerHTML = "<hr class = 'hrNewAddress'> <b> + </b>  New Address";
				_btnNewAddress.addEventListener("click",() => {
					this.createNewAddress();
				});

				this._listDiv.appendChild(_btnNewAddress);
            }            
        });
    }
	

	/** 
	 * Called when any of keyboard event is captured on input element. Used to set/reset focus
     */
	private setActiveListItem(): void {
		
		//clear previous active item if exist
		this._listDiv.childNodes.forEach(element => {
			let htmlElement = element as HTMLDivElement;

			if (htmlElement.className !== "addressnewbutton" && htmlElement.className !== "hrNewAddress") {
				htmlElement.setAttribute("class", "address-list-item");
			}
		});

		//set focus logic for last and first item on list
		if (this._currentFocus >= this._listDiv.childNodes.length - 1){
			this._currentFocus = 0;
		}
		if (this._currentFocus < 0){
			this._currentFocus = (this._listDiv.childNodes.length - 2);
		}

		//set current item as an active item
		let htmlElement = this._listDiv.childNodes[this._currentFocus] as HTMLDivElement;
		htmlElement.setAttribute("class", "address-list-active");
	}

	/** 
	 * Used to clear search results
     */
	private clearSearchResults(): void {
        this._listDiv.innerHTML = '';
		this._listDiv.hidden = true;
	}

	private createNewAddress(): void{
		let companyGuid: string = "";
		if (typeof (this._context.parameters.companyguid) !== "undefined" &&
            typeof (this._context.parameters.companyguid.raw) !== "undefined" && this._context.parameters.companyguid.raw != null) {
				companyGuid = this._context.parameters.companyguid.raw;
        }		

		if(companyGuid !== ""){
			this.openCreateForm(companyGuid);
        }
	}

	private async openCreateForm(companyGuid: string) {
        //await this._context.webAPI.retrieveRecord("account", companyGuid, '?$select=name').then( accountRecord => {
		//	if(accountRecord.name){
				let entityFormOptions: any = {};
				let formParameters: any = {};
				//let companyLookup: any;

				//form options
				entityFormOptions["formId"] = "DF5BEF3A-9237-40BD-A27C-2CD7FD434706";
				entityFormOptions["entityName"] = "customeraddress";

				//fields to update
				//companyLookup = [{entityType: 'account', id:accountRecord.accountid, name:accountRecord.name}];
				formParameters["parentid_account@odata.bind"] = "/accounts(" + companyGuid + ")";
				//formParameters["_parentid_value@Microsoft.Dynamics.CRM.associatednavigationproperty"] = "parentid_account";
				//formParameters["_parentid_value@Microsoft.Dynamics.CRM.associatednavigationproperty"] = "account";
				formParameters["objecttypecode"] = "account";
				//formParameters["objecttypecode@OData.Community.Display.V1.FormattedValue"] = "Account";
				formParameters["addresstypecode"] = 2;

				// Open the form.
				this._context.navigation.openForm(entityFormOptions, formParameters).then(
					function (success) {
						console.log(success);
					},
					function (error) {
						console.log(error);	
					}
				);
		//	}
		//});
    }
}