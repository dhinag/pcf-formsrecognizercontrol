import {IInputs, IOutputs} from "./generated/ManifestTypes";
import {FormFill} from "./FormFill";

'use strict';

	// Define const here
	// Show Error css classname
	const ShowErrorClassName = "ShowError";	
	let formFill : FormFill;

	export class FormsRecognizerControl implements ComponentFramework.StandardControl<IInputs, IOutputs> 
	{
		// PCF framework context, "Input Properties" containing the parameters, control metadata and interface functions.
		private _context: ComponentFramework.Context<IInputs>;

		// PCF framework delegate which will be assigned to this object which would be called whenever any update happens. 
		private _notifyOutputChanged: () => void;

		// Control's container
		private controlContainer: HTMLDivElement;
		
		// button element created as part of this control
		private uploadButton: HTMLButtonElement;

		// label element created as part of this control
		private errorLabelElement: HTMLLabelElement;	

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
			this._context = context;
			this._notifyOutputChanged = notifyOutputChanged;
			this.controlContainer = document.createElement("div");

			//Create an upload button to call the forms recognizer api
			this.uploadButton = document.createElement("button");

			// Get the localized string from localized string 
			this.uploadButton.innerHTML = context.resources.getString("PCF_FormsRecognizerControl_Recognize_ButtonLabel");
			this.uploadButton.addEventListener("click", this.onUploadButtonClick.bind(this));

			// Create an error label element
			this.errorLabelElement = document.createElement("label");
			this.errorLabelElement.setAttribute("id","lblError");		
				
			// Adding the label and button created to the container DIV.
			this.controlContainer.appendChild(this.uploadButton);			
			this.controlContainer.appendChild(this.errorLabelElement);			
			container.appendChild(this.controlContainer);
		}

		/**
		 * This button event handler will allow the user to pick the file from the device
		 * @param event
		 */
		private onUploadButtonClick(event: Event): void 
		{	
			// context.device.pickFile(successCallback, errorCallback) is used to initiate the File Explorer
			this._context.device.pickFile().then(this.processFile.bind(this), this.showError.bind(this));			
		}

		//This method will train the model based on the uploaded pdf documents in the Azure Storage and the api result will present us with the model id, which will be used in the analyzeDocument api call.
		// private customTrainModel() : void
		// {
		// 	$.ajax({
		// 		method: "POST",
		// 		url: "https://formsrecognizerpcf.cognitiveservices.azure.com/formrecognizer/v1.0-preview/custom/train",
		// 		data: "{\n\"source\": \"https://yourstorage.blob.core.windows.net/formrecognizer?sv=2018-03-29&ss=b&srt=sco&sp=rl&st=2017-08-22T19%3A90%3A31Z&se=2019-08-23T19%3A10%3A31Z&sig=qj99tVTv0mbt8jA%3D\"\n}",
		// 		beforeSend: function(xhr) {
		// 			xhr.setRequestHeader("Ocp-Apim-Subscription-Key", "dc2e05751ac1505b"),
		// 			xhr.setRequestHeader("Content-Type", "application/json")
		// 		}
		// 	});
		// }

		private analyzeDocument(formData : FormData) : void
		{			
			let subscriptionKey = this._context.parameters.subscriptionKey.raw;
			let analyzeModelUrl = this._context.parameters.analyzeModelUrl.raw;
			$.ajax({
			   method: "POST",
			   url: analyzeModelUrl,
			   data: formData,
			   processData: false,
			   contentType: false,
			   mimeType: "multipart/form-data",
			   beforeSend: function(xhr) {
				   xhr.setRequestHeader("Ocp-Apim-Subscription-Key", subscriptionKey),
				   xhr.setRequestHeader("Content-Type", "application/pdf")				  
			   }, success: (data) => {
				   var obj = JSON.parse(data);
				   formFill = new FormFill();
				   obj.pages[0].keyValuePairs.forEach(function (element : any)
				   {
					   switch(element.key[0].text)
					   {
						   case "Last Name:" : {
							   formFill.lastName = element.value[0].text;
							   break;
						   }
						   case "First Name:" : {
							   formFill.firstName = element.value[0].text;
							   break;
						   }
						   case "Job Title:" : {
							   formFill.jobTitle = element.value[0].text;
							   break;
						   }
						   case "E-mail:" : {
							   formFill.email = element.value[0].text;
							   break;
						   }
						   case "Telephone:" : {
							   formFill.telephone = element.value[0].text;
							   break;
						   }									
						   default: {
							   break;
						   }
					   }							
				   });

				   this.hideError("");
				   if (Object.keys(formFill).length !== 0)
					{
						this.hideError("");
						this._notifyOutputChanged();
					}
					else 
					{
						this.showError("Please check the PDF and try again.");
					}
			   },
			   error: (jqXHR) => {
				   //Show the media unsupported error if the response status code is 415.
				   if(jqXHR.status === 415) {					

					   this.showError(this._context.resources.getString("PCF_FormsRecognizerControl_Media_UnSupported_Error"));
				   }
				   else {
					   this.showError(this._context.resources.getString("PCF_FormsRecognizerControl_General_Error"));
				   }
			   }					
	   });
	}
		
		/**
		 * 
		 * @param files 
		 */
		private processFile(files: ComponentFramework.FileObject[]): void
		{
			this.disableRecognizeButton();					
			if(files.length > 0)
			{
				let file: ComponentFramework.FileObject = files[0];					
				var formData = new FormData();
				formData.append("type", "application/pdf");
			
				let fileObj = this.EncodeToBase64("FileContent", file.fileContent);
				const url = fileObj;
				fetch(url)
				.then(res => res.blob())
				.then(blob => {
					const file = new File([blob], "Application");
					formData.append("form", file);
					this.analyzeDocument(formData);		
					
				}).catch(error => this.showError(this._context.resources.getString("PCF_FormsRecognizerControl_General_Error") + " " + error));				
			}
		}		
		
		/**
		 * 
		 * @param fileType file extension
		 * @param fileContent file content, base 64 format
		 */
		private EncodeToBase64(fileType: string, fileContent: string): string
		{
			return  "data:application/pdf;base64, " + fileContent;
		}

		/** 
		 *  Show Error Message
		 */
		private showError(errorText : string): void
		{
			this.enableRecognizeButton();	
			this.errorLabelElement.innerHTML = errorText;
			this.controlContainer.classList.add(ShowErrorClassName);
		}

		/** 
		 *  Hide Error Message
		*/
		private hideError(errorText : string): void
		{
			this.enableRecognizeButton();	
			this.errorLabelElement.innerHTML = "";
			this.controlContainer.classList.remove("ShowErrorClassName");
		}

		private enableRecognizeButton(): void
		{
			this.uploadButton.removeAttribute("disabled");
			this.uploadButton.innerHTML = this._context.resources.getString("PCF_FormsRecognizerControl_Recognize_ButtonLabel");	
		}

		private disableRecognizeButton(): void
		{
			this.uploadButton.innerHTML = this._context.resources.getString("PCF_FormsRecognizerControl_Recognizing_ButtonLabel");	
			this.uploadButton.setAttribute("disabled", "disabled");			
		}

		/** 
		 * It is called by the framework prior to a control receiving new data. 
		 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
		*/
		public getOutputs(): IOutputs
		{
			let result: IOutputs = 
			{
				lastName : formFill.lastName,
				firstName : formFill.firstName,
				jobTitle : formFill.jobTitle,
				email : formFill.email,
				telephone : formFill.telephone,
			};		
			return result;
		}

		/**
		 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
		 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
		 */
		public updateView(context: ComponentFramework.Context<IInputs>): void
		{
			this._context = context;
		}

		/** 
 		 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
		 * i.e. cancelling any pending remote calls, removing listeners, etc.
		 */
		public destroy(): void
		{
			this.uploadButton.removeEventListener("click", this.onUploadButtonClick.bind(this));
		}
	}