import * as React from "react";
import * as ReactDOM from "react-dom";
import { Log, FormDisplayMode } from "@microsoft/sp-core-library";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";

import "bootstrap/dist/css/bootstrap.min.css";  

import { IListFormInterfaceModel } from "../IListFormInterfaceModel";
import styles from "./ListFormCustomizer.module.scss";
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { add } from "lodash";

export interface IListFormCustomizerProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  // onSave?: () => void;
  // onClose?: () => void;
  items?:IListFormInterfaceModel;
  formSaved?:()=>void;
  formClosed?:()=>void;
  domElement:HTMLElement;
}

export interface IListFormCustomizerState{
  itemsState:IListFormInterfaceModel;
  FormFirstNameState?:string;
  FormLastNameState?:string;
  FormDOBState?:string;
  FormEmailState?:string;
  FormPhoneNumberState?:string;
  FormAddressState?:string;
  FormQualificationState?:string;
  FormExperienceState?:string;

}

const LOG_SOURCE: string = "ListFormCustomizer";

export default class ListFormCustomizer extends React.Component<
  IListFormCustomizerProps,
  IListFormCustomizerState
> {
    // Added for the item to show in the form; use with edit and view form
    private _item: IListFormInterfaceModel;
    this_item = {};
    // Added for item's etag to ensure integrity of the update; used with edit form
    private _etag?: string;

  constructor(props: IListFormCustomizerProps) {
    super(props);
    this.state = {
      itemsState: {},
      FormFirstNameState:"",
      FormLastNameState:"",
      FormDOBState:"",
      FormEmailState:"",
      FormPhoneNumberState:"",
      FormAddressState:"",
      FormQualificationState:"",
      FormExperienceState:"",

    };
  }

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, "React Element: ListFormCustomizer mounted");
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, "React Element: ListFormCustomizer unmounted");
  }

  
  private handleChangeFirstName = (e: { target: { value: any; }; }) => {
    this.setState({
     FormFirstNameState:e.target.value
    });
  };

  private handleChangeFormLastName = (e: { target: { value: any; }; }) => {
    this.setState({
     FormLastNameState:e.target.value
    });
  };

  private handleChangeFormDOB = (e: { target: { value: any; }; }) => {
    this.setState({
     FormDOBState:e.target.value
    });
  };

  private handleChangeFormEmail = (e: { target: { value: any; }; }) => {
    this.setState({
     FormEmailState:e.target.value
    });
  };

  private handleChangeFormPhoneNumber = (e: { target: { value: any; }; }) => {
    this.setState({
     FormPhoneNumberState:e.target.value
    });
  };

  private handleChangeFormAddress = (e: { target: { value: any; }; }) => {
    this.setState({
     FormAddressState:e.target.value
    });
  };


  private handleChangeFormQualification = (e: { target: { value: any; }; }) => {
    this.setState({
     FormQualificationState:e.target.value
    });
  };


  private handleChangeFormExperience = (e: { target: { value: any; }; }) => {
    this.setState({
     FormExperienceState:e.target.value
    });
  };



  public renderHtml(): JSX.Element{

    if (this.props.displayMode === FormDisplayMode.Display) {
     let displayHtml=( <>
      <section className="text-capitalize bg-light border-primary flex-fill">
        <h1 className="text-center text-capitalize">Application form</h1>
        <div className="container">
          <form id="application-form">
            <div className="form-group">
              <div className="form-row">
                <div className="col">
                  <p>
                    <strong>First Name</strong>
                    <span className="text-danger">*</span>
                  </p>
                  <input
                    type="text"
                    className="form-control"
                    required
                    placeholder="Ex. Praveen"
                    disabled
                    value={this.state.itemsState.FormFirstName}
                  />
                </div>
                <div className="col">
                  <p>
                    <strong>Last Name</strong>
                    <span className="text-danger">*</span>
                  </p>
                  <input
                    type="text"
                    className="form-control"
                    required
                    placeholder="Ex. Kumar"
                    disabled
                    value={this.state.itemsState.FormLastName}
                  />
                </div>
              </div>
            </div>
            <div className="form-group">
              <div className="form-row">
                <div className="col">
                  <p>
                    <strong>Date Of Birth</strong>
                    <span className="text-danger">*</span>
                  </p>
                  <input className="form-control" type="date"
                   required
                   disabled
                   value={this.state.itemsState.FormDOB}
                    />
                </div>
              </div>
            </div>
            <div className="form-group">
              <p>
                <strong>Email </strong>
                <span className="text-danger">*</span>
              </p>
              <input
                type="email"
                className="form-control"
                placeholder="user@domain.com"
                disabled
                value={this.state.itemsState.FormEmail}
              />
            </div>
            <div className="form-group">
              <p>
                <strong>Whatsapp Number </strong>
                <span className="text-danger">*</span>
              </p>
              <input
                type="number"
                className="form-control"
                placeholder="7777777777"
                disabled
                value={this.state.itemsState.FormPhoneNumber}
              />
            </div>
            <div className="form-group">
              <p>
                <strong>Address </strong>
                <span className="text-danger">*</span>
              </p>
              <input
                type="text"
                className="form-control"
                required
                placeholder="1st floor chennai"
                disabled
                value={this.state.itemsState.FormAddress}
              />
            </div>
            <div
              className="form-group"
              // style="box-shadow: inset 0px 0px;text-shadow: 0px 0px;"
            >
              <div className="form-row">
                <div className="col">
                  <p>Qualification</p>
                  <textarea
                    className="form-control"
                    //  type="text"
                    required
                    placeholder="Ex.  2006-2010"
                    name="Qualification"
                    disabled
                    value={this.state.itemsState.FormQualification}
                  ></textarea>
                </div>
                <div className="col">
                  <p>Experience</p>
                  <input                   
                    type="text"
                    className="form-control"
                    required
                    placeholder="Solution Architect"
                    disabled
                    value={this.state.itemsState.FormExperience}
                  />
                </div>
              </div>
            </div>
            <div className="form-group justify-content-center d-flex">
              <div id="submit-btn cancel">
                <div className="form-row">
                  <button
                    className="btn btn-primary btn-light m-0 rounded-pill px-4"
                    type="button"
                    // style="min-width: 500px;"
                    // action
                    // method="POST"
                    // target="hidden_iframe"
                  >
                    cancel
                  </button>
                </div>
              </div>
            </div>
          </form>
        </div>
      </section>
    </>
    );
    let cancelButton=document.getElementById('cancel');
    if(cancelButton)
    cancelButton.addEventListener('click', this._onClose.bind(this));
    return displayHtml;

    }
    else
    {
      let NeworEditHtml=( <>
        <section className="text-capitalize bg-light border-primary flex-fill">
          <h1 className="text-center text-capitalize">Application form</h1>
          <div className="container">
            <form id="application-form">
              <div className="form-group">
                <div className="form-row">
                  <div className="col">
                    <p>
                      <strong>First Name</strong>
                      <span className="text-danger">*</span>
                    </p>
                    <input
                    id="FormFirstName"
                      type="text"
                      className="form-control"
                      required
                      placeholder="Ex. Praveen"
                      onChange={this.handleChangeFirstName}
                      value={this.state.FormFirstNameState?this.state.FormFirstNameState:''}
                    />
                  </div>
                  <div className="col">
                    <p>
                      <strong>Last Name</strong>
                      <span className="text-danger">*</span>
                    </p>
                    <input
                    id="FormLastName"
                      type="text"
                      className="form-control"
                      required
                      placeholder="Ex. Kumar"
                      onChange={this.handleChangeFormLastName}
                      value={this.state.FormLastNameState?this.state.FormLastNameState:''}
                    />
                  </div>
                </div>
              </div>
              <div className="form-group">
                <div className="form-row">
                  <div className="col">
                    <p>
                      <strong>Date Of Birth</strong>
                      <span className="text-danger">*</span>
                    </p>
                    <input className="form-control"
                    id="FormDOB"
                     type="date" 
                     required
                     onChange={this.handleChangeFormDOB}
                     value={this.state.FormDOBState?this.state.FormDOBState:''} />
                  </div>
                </div>
              </div>
              <div className="form-group">
                <p>
                  <strong>Email </strong>
                  <span className="text-danger">*</span>
                </p>
                <input
                 id="FormEmail"
                  type="email"
                  className="form-control"
                  placeholder="user@domain.com"
                  onChange={this.handleChangeFormEmail}
                  value={this.state.FormEmailState?this.state.FormEmailState: ''}
                />
              </div>
              <div className="form-group">
                <p>
                  <strong>Whatsapp Number </strong>
                  <span className="text-danger">*</span>
                </p>
                <input
                 id="FormPhoneNumber"
                  type="number"
                  className="form-control"
                  placeholder="7777777777"
                  onChange={this.handleChangeFormPhoneNumber}
                  value={this.state.FormPhoneNumberState?this.state.FormPhoneNumberState: ''}
                />
              </div>
              <div className="form-group">
                <p>
                  <strong>Address </strong>
                  <span className="text-danger">*</span>
                </p>
                <input
                   id="FormAddress"
                  type="text"
                  className="form-control"
                  required
                  placeholder="1st floor chennai"
                  onChange={this.handleChangeFormAddress}
                  value={this.state.FormAddressState?this.state.FormAddressState: ''}
                />
              </div>
              <div
                className="form-group"
                // style="box-shadow: inset 0px 0px;text-shadow: 0px 0px;"
              >
                <div className="form-row">
                  <div className="col">
                    <p>Qualification</p>
                    <textarea
                    id="FormQualification"
                      className="form-control"
                      //  type="text"
                      required
                      placeholder="Ex.  2006-2010"
                      name="Qualification"
                      onChange={this.handleChangeFormQualification}
                      value={this.state.FormQualificationState?this.state.FormQualificationState: ''}
                    ></textarea>
                  </div>
                  <div className="col">
                    <p>Experience</p>
                    <input
                     id="FormExperience"
                      type="text"
                      className="form-control"
                      required
                      placeholder="Solution Architect"
                      onChange={this.handleChangeFormExperience}
                      value={this.state.FormExperienceState?this.state.FormExperienceState: ''}
                    />
                  </div>
                </div>
              </div>
              <div className="form-group justify-content-center d-flex">
                <div id="submit-btn">
                  <div className="form-row">
                    <button
                    id="save"
                      className="btn btn-primary btn-light m-0 rounded-pill px-4"
                      type="button"
                      onClick={(e) => {
                        this._onSave(e)
                      }}
                      // style="min-width: 500px;"
                      // action
                      // method="POST"
                      // target="hidden_iframe"
                    >
                      save
                    </button>
                    <button
                    id="cancel"
                      className="btn btn-primary btn-light m-0 rounded-pill px-4"
                      type="button"
                      // style="min-width: 500px;"
                      // action
                      // method="POST"
                      // target="hidden_iframe"
                    >
                      cancel
                    </button>
                  </div>
                </div>
              </div>
            </form>
          </div>
        </section>
      </>
      );
     // let saveButton= document.getElementById('save');
    //   if(saveButton)
    //  saveButton.addEventListener('click', this._onSave.bind(this));
    //  let cancelButton=  document.getElementById('cancel');
    //  if(cancelButton)
    // cancelButton.addEventListener('click', this._onClose.bind(this));
      return NeworEditHtml;
    }

  }

  
 public getItemValueFromAllControls():IListFormInterfaceModel
 {

  const FormFirstName: string = (document.getElementById("FormFirstName") as HTMLInputElement)
  .value;
  const FormLastName: string = (document.getElementById("FormLastName") as HTMLInputElement)
  .value;
  const FormDOB: string = (document.getElementById("FormDOB") as HTMLInputElement)
  .value;
  const FormEmail: string = (document.getElementById("FormEmail") as HTMLInputElement)
  .value;
  const FormPhoneNumber: string = (document.getElementById("FormPhoneNumber") as HTMLInputElement)
  .value;
  const FormAddress: string = (document.getElementById("FormAddress") as HTMLInputElement)
  .value;
  const FormQualification: string = (document.getElementById("FormQualification") as HTMLInputElement)
  .value;
  const FormExperience: string = (document.getElementById("FormExperience") as HTMLInputElement)
  .value;

  const newObjItems={
    FormFirstName:FormFirstName,
    FormLastName:FormLastName,
    FormDOB:FormDOB,
    FormEmail:FormEmail,
    FormPhoneNumber:FormPhoneNumber,
    FormAddress:FormAddress,
    FormQualification:FormQualification,
    FormExperience:FormExperience
  }
  

  this.setState({
  itemsState:newObjItems
   });

  // if(FormFirstName)
  // this.state.itemsState.FormFirstName=FormFirstName;
  // if(FormLastName)
  // this.state.itemsState.FormLastName=FormLastName;
  // if(FormDOB)
  // this.state.itemsState.FormDOB=FormDOB;
  // if(FormEmail)
  // this.state.itemsState.FormEmail=FormEmail;
  // if(FormPhoneNumber)
  // this.state.itemsState.FormPhoneNumber=FormPhoneNumber;
  // if(FormAddress)
  // this.state.itemsState.FormAddress=FormAddress;
  // if(FormQualification)
  // this.state.itemsState.FormQualification=FormQualification;
  // if(FormExperience)
  // this.state.itemsState.FormExperience=FormExperience;

  return newObjItems;
 }

 private _updateItem(items: IListFormInterfaceModel): Promise<SPHttpClientResponse> {
  return this.props.context.spHttpClient
  .post(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${this.props.context.list.title}')/items(${this.props.context.itemId})`, SPHttpClient.configurations.v1, {
    headers: {
      'content-type': 'application/json;odata.metadata=none',
      'if-match': this._etag,
      'x-http-method': 'MERGE'
    },
    body: JSON.stringify({
      FormFirstName: items.FormFirstName,
      FormLastName:items.FormLastName,
      FormDOB:items.FormDOB,
      FormEmail:items.FormEmail,
      FormPhoneNumber:items.FormPhoneNumber,
      FormAddress:items.FormAddress,
      FormQualification:items.FormQualification,
      FormExperience:items.FormExperience
    })
  });
}
private _createItem(items: IListFormInterfaceModel): Promise<SPHttpClientResponse> {
  return this.props.context.spHttpClient
  .post(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${this.props.context.list.title}')/items`, SPHttpClient.configurations.v1, {
    headers: {
      'content-type': 'application/json;odata.metadata=none'
    },
    body: JSON.stringify({
      FormFirstName: items.FormFirstName,
      FormLastName:items.FormLastName,
      FormDOB:items.FormDOB,
      FormEmail:items.FormEmail,
      FormPhoneNo:items.FormPhoneNumber,
      FormAddress:items.FormAddress,
      FormQualification:items.FormQualification,
      FormExperience:items.FormExperience
    })
  });
}

// e: React.MouseEvent<HTMLButtonElement, MouseEvent>

  private _onSave = async (e: React.MouseEvent<HTMLButtonElement, MouseEvent>): Promise<void> => {
   // e.preventDefault();
    // disable all input elements while we're saving the item
    this.props.domElement
      .querySelectorAll("input")
      .forEach((el) => el.setAttribute("disabled", "disabled"));
    // reset previous error message if any
    // this.domElement.querySelector(`.${styles.error}`).innerHTML = "";

    let request: Promise<SPHttpClientResponse>;
    let getItemValueFromAllControls=this.getItemValueFromAllControls();
    // const title: string = (document.getElementById("title") as HTMLInputElement)
    //   .value;

    switch (this.props.displayMode) {
      case FormDisplayMode.New:
        request = this._createItem(getItemValueFromAllControls);
        break;
      case FormDisplayMode.Edit:
        request = this._updateItem(getItemValueFromAllControls);
    }

    const res: SPHttpClientResponse = await request;

    if (res.ok) {
      // You MUST call this.formSaved() after you save the form.
      this.props.formSaved;
    } else {
      const error: { error: { message: string } } = await res.json();

      // this.domElement.querySelector(
      //   `.${styles.error}`
      // ).innerHTML = `An error has occurred while saving the item. Please try again. Error: ${error.error.message}`;
      this.props.domElement
        .querySelectorAll("input")
        .forEach((el) => el.removeAttribute("disabled"));
    }
    // You MUST call this.formSaved() after you save the form.
    this.props.formSaved;
  };

  private _onClose = (): void => {
    // You MUST call this.formClosed() after you close the form.
    this.props.formClosed;
  };

  public render(): React.ReactElement<{}> {
    return (
      this.renderHtml()      
    );
  }
}
