  
import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { debounce, Options, Procedure } from 'ts-debounce';
import { stringify } from "querystring";

enum AsciiKeyCode {
	Enter = 13,
	Up = 38,
	Down = 40
}

interface Address {
    addressname: string;
    addressLine1: string;
	addressLine2: string;
	addressLine3: string;
	city: string;
	state: string;
	postalCode: string;
	country: string;
}


export class MultipleAddresses implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private inputElement: HTMLInputElement;
    private _iconElement: HTMLSpanElement;
	private inputKeyPress: EventListenerOrEventListenerObject;
    private listItemClick: EventListenerOrEventListenerObject;
    private _iconOnClick: EventListenerOrEventListenerObject;
	private context: ComponentFramework.Context<IInputs>;
	private notifyOutputChanged: () => void;
	private container: HTMLDivElement;
	private searchDiv: HTMLDivElement;
	private debouncedKeyPress: Procedure;

	private currentFocus: number;
	private address : Address;
    

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
        console.log("Inside init");
        this.context = context;
		this.notifyOutputChanged = notifyOutputChanged;
		this.container = container;
		this.currentFocus = -1;
        this.inputKeyPress = this.keyPress.bind(this);
        this.listItemClick = this.itemClick.bind(this);
        this._iconOnClick = this.keyPress.bind(this);

		/* add search input form control */
		this.inputElement = document.createElement("input");
		this.inputElement.setAttribute("type", "text");
		this.inputElement.setAttribute("id", "MultipleAddressSearch");
		this.inputElement.setAttribute("autocomplete", "off"); // tell firefox and chrome not to auto complete
		this.inputElement.setAttribute("name", "searchInputField");
		this.inputElement.setAttribute("class", "autocomplete");
        this.inputElement.addEventListener("keyup", this.inputKeyPress);

        this.container.appendChild(this.inputElement);
        this.container.classList.add("container");
		this.container.classList.add("iconPresent");
        this._iconElement = document.createElement("span");
		this._iconElement.setAttribute("id", "iconElement");
		this._iconElement.classList.add("ms-Icon");
		this._iconElement.classList.add("noOutline");
		this._iconElement.addEventListener("click", this._iconOnClick);
        //this.container.appendChild(this._iconElement);

		/* create a DIV element that will contain the items (values): */
		this.searchDiv = document.createElement("div");
		this.searchDiv.setAttribute("id", "01-autocomplete-list");
		this.searchDiv.setAttribute("class", "autocomplete-items");
		this.container.appendChild(this.searchDiv);
		this.searchDiv.hidden = true;

		/* add debounce variable for keypress */
		let debounceOptions: Options;
		debounceOptions = { isImmediate: false };
		this.debouncedKeyPress = debounce(this.keyPressDebounced, 200, debounceOptions);
    }    

    public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
        this.context = context;
        console.log("Inside update view");
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
        console.log("Inside getOutputs");
        console.log("Inside getOutputs", this.address.addressname);
        console.log("Inside getOutputs", this.address.addressLine1);
        return {
            addressname: this.address.addressname,
            addressline1: this.address.addressLine1,
			addressline2: this.address.addressLine2,
			addressline3: this.address.addressLine3,
			city: this.address.city,
			state: this.address.state,
			postalcode: this.address.postalCode,
			country: this.address.country
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
		delete this.inputElement;
        delete this.container;
        console.log("Inside destroy");
	}

	private itemClick(evt: Event): void {
        console.log("Inside itemClick");
        let divItem = evt.target as HTMLDivElement;
		this.inputElement.value = '';
		this.clearSearchResults();

        let _this = this;
		let addressGuid = divItem.id;
        console.log("addressGuid..", addressGuid);
        this.searchSelectedAddress(addressGuid).then((result) => {
            if(result.name){
                this.address.addressname = result.name;
                this.address.addressLine1 = result.line1;
                this.address.addressLine2 = result.line2;
                this.address.addressLine3 = result.line3;
                this.address.city = result.city;
                this.address.state = result.stateorprovince;
                this.address.postalCode = result.postalcode;                
                this.address.country = result.country;
                this.notifyOutputChanged();
            }
        });

		/*(async () => {

			this.searchSelectedAddress(addressGuid)
				.then(addressRecord => this.populateAddressData(addressRecord));
		})();*/
    }
    
    public async searchSelectedAddress(addressGuid: string) {
        console.log("Inside searchSelectedAddress");
        const response = await this.context.webAPI.retrieveRecord("customeraddress", addressGuid, '?$select=name,line1,line2,line3,city,stateorprovince,postalcode,country');
        console.log("Result: ", response);
        return response;
    }

	private populateAddressData(result: any): void {
        //if(addressRecord.name){
            this.address.addressname = result.name;
            this.address.addressLine1 = result.line1;
            this.address.addressLine2 = result.line2;
            this.address.addressLine3 = result.line3;
            this.address.city = result.city;
            this.address.state = result.stateorprovince;
            this.address.postalCode = result.postalcode;                
            this.address.country = result.country;
            this.notifyOutputChanged();
        //}      
	}
	
	private keyPress(evt: Event): void {
        console.log("Inside keypress");
		let keyboardEvent = evt as KeyboardEvent;
        console.log("keyboardEvent ", keyboardEvent);
		if (keyboardEvent.keyCode == AsciiKeyCode.Down) { 
			this.currentFocus++;
			this.setActiveElement();
		} else if (keyboardEvent.keyCode == AsciiKeyCode.Up) { 
			this.currentFocus--;
			this.setActiveElement();
		} else if (keyboardEvent.keyCode == AsciiKeyCode.Enter) { 
			var element1 = this.searchDiv.childNodes[this.currentFocus] as HTMLDivElement;
			element1.click();
		}
		else {
			/* only perform search if not keys pressed above */
			this.debouncedKeyPress();
		}
	}

	private keyPressDebounced(): void {
        console.log("Inside keyPressDebounced");
        var queryString = this.inputElement.value;
        let companyGuid: any = this.context.parameters.companyguid.raw;
        console.log("companyGuid: ",companyGuid);
        if(companyGuid !== "" || companyGuid != null)
        {
            //forming query string
            let queryString: string = "?$select=name,line1,line2,line3,city,stateorprovince,postalcode,country&$filter=_parentid_value eq " + companyGuid + " and addressnumber ne 2";
            this.populateAddressRecords(queryString).then( () => {
                console.log("Populating Address records finished...");
            });
        }
		/*(async () => {

			this.experianGlobalIntuitive.AddressSearch(queryString, ExperianCountry.Australia, ExperianDataset.AustraliaDataFusion)
				.then(function (response) {
					return response;
				})
				.then(address => this.populateData(address));
		})()*/
    }
    
    private async populateAddressRecords(queryString: string){        
        console.log("Inside populateAddressRecords");
        this.clearSearchResults();

        //Querying for address records
        await this.context.webAPI.retrieveMultipleRecords("customeraddress", queryString).then( addressResult => {
            if(addressResult.entities.length > 0){
                // Looping through each address record
                console.log("addressResult: ",addressResult);
                for (var i = 0; i < addressResult.entities.length; i++) {
                    //appeding address records onto table
                    let searchItemDiv = document.createElement("DIV");
                    var completeAddress = "";
                    completeAddress = (addressResult.entities[i].name !== null ? addressResult.entities[i].name + ", " : "");
                    completeAddress = completeAddress + (addressResult.entities[i].line1 !== null ? addressResult.entities[i].line1 + ", " : "");
                    completeAddress = completeAddress + (addressResult.entities[i].line2 !== null ? addressResult.entities[i].line2 + ", " : "");
                    completeAddress = completeAddress + (addressResult.entities[i].line3 !== null ? addressResult.entities[i].line3 + ", " : "");
                    completeAddress = completeAddress + (addressResult.entities[i].city !== null ? addressResult.entities[i].city + ", " : "");
                    completeAddress = completeAddress + (addressResult.entities[i].stateorprovince !== null ? addressResult.entities[i].stateorprovince + ", " : "");
                    completeAddress = completeAddress + (addressResult.entities[i].postalcode !== null ? addressResult.entities[i].postalcode + ", " : "");
                    completeAddress = completeAddress + (addressResult.entities[i].country !== null ? addressResult.entities[i].country : "");

                    searchItemDiv.innerHTML = completeAddress;
                    searchItemDiv.setAttribute("class", "autocomplete-item");
                    /* insert a input field that will hold the current array item's value */
                    searchItemDiv.id = addressResult.entities[i].customeraddressid;
                    /* execute a function when someone clicks on the item value (DIV element) */
                    searchItemDiv.addEventListener("click", this.listItemClick);
                    
                    this.searchDiv.appendChild(searchItemDiv);
                }
                console.log("Search Div : ", this.searchDiv);
            }            
        });
        
        let footer = document.createElement("DIV");
        footer.setAttribute("class", "list-footer");

        let footerWrapper = document.createElement("DIV");
        footerWrapper.setAttribute("class","list-footer-wrapper");
        footerWrapper.id = "list-footer-id";
        footerWrapper.appendChild(footer);
        this.searchDiv.appendChild(footerWrapper);
        this.searchDiv.hidden = false;
    }

	private clearSearchResults(): void {
        console.log("Inside clearSearchResults");
        this.searchDiv.innerHTML = '';
		this.searchDiv.hidden = true;
	}

	private setActiveElement(): void {
        console.log("Inside setActiveElement");
        /* clear active elements if they exist */
		this.searchDiv.childNodes.forEach(element => {
			let htmlElement = element as HTMLDivElement;

			if (htmlElement.id != "list-footer-id") {
				htmlElement.setAttribute("class", "autocomplete-item");
			}
		});

		if (this.currentFocus >= this.searchDiv.childNodes.length - 1) this.currentFocus = 0;
		if (this.currentFocus < 0) this.currentFocus = (this.searchDiv.childNodes.length - 2);

		let htmlElement = this.searchDiv.childNodes[this.currentFocus] as HTMLDivElement;
		htmlElement.setAttribute("class", "autocomplete-active");
	}
}