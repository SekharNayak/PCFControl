import {IInputs, IOutputs} from "./generated/ManifestTypes";
import axios from "axios";

export class DemoControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

   // Value of the field is stored and used inside the control 
		private _value: number;

		// PCF framework delegate which will be assigned to this object which would be called whenever any update happens. 
		private _notifyOutputChanged: () => void;

		// label element created as part of this control
		private label: HTMLInputElement;

		// button element created as part of this control
		private button: HTMLButtonElement;

		// Reference to the control container HTMLDivElement
		// This element contains all elements of our custom control example
		private _container: HTMLDivElement;

		private _innerDiv : HTMLDivElement;

		private  _form : HTMLFormElement;

		private _errorMessage : HTMLInputElement;

		/**
		 * Empty constructor.
		 */
		constructor()
		{

		}

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
			// Creating the label for the control and setting the relevant values.
			this.label = document.createElement("input");
			this.label.setAttribute("type", "label");
			this.label.addEventListener("blur", this.onInputBlur.bind(this));
			

			//Create a button to increment the value by 1.
			this.button = document.createElement("button");
			
			// Get the localized string from localized string 
			this.button.innerHTML = context.resources.getString("TS_IncrementControl_ButtonLabel");

			this.button.classList.add("SimpleIncrement_Button_Style");
			this._notifyOutputChanged = notifyOutputChanged;
			//this.button.addEventListener("click", (event) => { this._value = this._value + 1; this._notifyOutputChanged();});
			this.button.addEventListener("click", this.onButtonClick.bind(this));

			// Adding the label and button created to the container DIV.
			this._container = document.createElement("div");
			this._container.appendChild(this.label);
			this._container.appendChild(this.button);
			container.appendChild(this._container);
		}

		/**
		 * Button Event handler for the button created as part of this control
		 * @param event
		 */
		private onButtonClick(event: Event): void {

			let divToBeRemoved = document.getElementById("palceholder");
				if(divToBeRemoved != null){
					this._container.removeChild(divToBeRemoved)
				}

				let placeholderDiv = document.createElement("div");
				placeholderDiv.setAttribute("id","palceholder");
				

			const url = `https://customermanager.azurewebsites.net/api/v1/CustomerManager/${this._value}`;
		    axios.get(url,{
				headers: {
				  'x-api-key': `36b407e0-57f7-4b85-abae-8be89baedad9`
				}
			  }).then(response => {

				
				const { msfsi_EMailAddress : email ,
					 msfsi_FirstName : firstName,
					 msfsi_LastName : lastName,
					 msfsi_MobilePhone : phone,
					msfsi_LoanApplicantNumber : applicantNumber
					} = response.data;

				
				let innerDiv = this.BuildFormElement(email,firstName,lastName , phone , applicantNumber);
				placeholderDiv.appendChild(innerDiv);
				this._container.append(placeholderDiv);
			})
			.catch(error => {
				this._errorMessage = document.createElement("input");
				this._errorMessage.setAttribute("type", "label");
				this._errorMessage.classList.add("SimpleIncrement_Input_Error_Style");
				this._errorMessage.value = error.message;
				placeholderDiv.appendChild(this._errorMessage);
				this._container.appendChild(placeholderDiv);
			});
			this._notifyOutputChanged();
		}

		/**
		 * Input Blur Event handler for the input created as part of this control
		 * @param event
		 */
		private onInputBlur(event: Event): void {
			let inputNumber = Number(this.label.value);
			this._value = isNaN(inputNumber) ? (this.label.value as any) as number: inputNumber;
			this._notifyOutputChanged();
		}

		/**
		 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
		 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
		 */
		public updateView(context: ComponentFramework.Context<IInputs>): void
		{
			// This method would rerender the control with the updated values after we call NotifyOutputChanged
			//set the value of the field control to the raw value from the configured field
			this._value = context.parameters.value.raw!;
			this.label.value = this._value != null ? this._value.toString(): "";

			if(context.parameters.value.error)
			{
				this.label.classList.add("SimpleIncrement_Input_Error_Style");
			}
			else
			{
				this.label.classList.remove("SimpleIncrement_Input_Error_Style");
			}
		}

		/** 
		 * It is called by the framework prior to a control receiving new data. 
		 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
		 */
		public getOutputs(): IOutputs
		{
			// custom code goes here - remove the line below and return the correct output
			let result: IOutputs = {
				value: this._value
			};
			return result;
		}

		/** 
 		 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
		 * i.e. cancelling any pending remote calls, removing listeners, etc.
		 */
		public destroy(): void
		{
		}

		private BuildFormElement(email : string , firstName : string , lastName : string , phone :string , applicantNumber :string  ) : HTMLDivElement{
			let formElement = document.createElement("div");
			formElement.setAttribute("id","innerDiv");
			let para = document.createElement("p");
			para.innerText = "Customer search response ";
			formElement.appendChild(para);

			let line = document.createElement("hr");
			formElement.appendChild(line);

			//email 
			formElement.appendChild(this.CreateLableDiv("Email"));
			formElement.appendChild(this.CreateInputDiv(email));
		
			//first name 
			formElement.appendChild(this.CreateLableDiv("First Name "));
			formElement.appendChild(this.CreateInputDiv(firstName));
			//last name 
			formElement.appendChild(this.CreateLableDiv("Last name"));
			formElement.appendChild(this.CreateInputDiv(lastName));

			formElement.appendChild(this.CreateLableDiv("Mobile"));
			formElement.appendChild(this.CreateInputDiv(phone));

			formElement.appendChild(this.CreateLableDiv("Customer Number"));
			formElement.appendChild(this.CreateInputDiv(applicantNumber));
			return formElement;
		}

		private BuildEmptyDiv() : HTMLDivElement{
			let emptyDiv = document.createElement("div");
			return emptyDiv;
		}

		private CreateSingleColumn(labelName : string , inputValue : string ) : HTMLDivElement {
			let divElement = document.createElement("div");
			divElement.appendChild(this.CreateLableDiv(labelName));
			divElement.appendChild(this.CreateInputDiv(inputValue));
			return divElement;
		}

		private CreateLableDiv(name : string ){
			let label = document.createElement("label");
			label.innerText = name;
			return label;
		}
		private CreateInputDiv(inputValue : string ){
			let label = document.createElement("input");
			label.type = "text";
			label.classList.add("type-2");
			label.value = inputValue;
			return label;
		}
	}