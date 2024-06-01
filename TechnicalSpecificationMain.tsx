import * as React from 'react';
import style from './TechnicalSpecificationMain.module.scss';
import { ITechnicalSpecificationMainProps } from './ITechnicalSpecificationMainProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { ITechnicalSpecificationMainState } from './ITechnicalSpecificationMainStates';
import { spListNames } from './ITechnicalSpecificationMainProvider';
import TechnicalSpecificationMainProvider, { ITechnicalSpecificationMainProvider } from './ITechnicalSpecificationMainProvider';

import { Fabric } from 'office-ui-fabric-react';
import { sp } from "@pnp/sp";
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { TextField, MaskedTextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { Link, MessageBar, MessageBarType } from '@fluentui/react';
import { DatePicker, IDatePickerStyles, defaultDatePickerStrings, SpinButton, ISpinButtonStyles, Position } from '@fluentui/react';
import { Separator } from '@fluentui/react/lib/Separator';
import { Icon } from '@fluentui/react/lib/Icon';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { CommandBarButton, IconButton, DefaultButton, CommandButton } from '@fluentui/react/lib/Button';
import { Label } from '@fluentui/react/lib/Label';
import 'bootstrap/dist/css/bootstrap.min.css';

import TechnicalSpecificationInputFormComponent from './InputForm/TechnicalSpecificationInputForm';
import TechnicalSpecificationSearchComponent from './Search/TechnicalSpecificationSearch'

require('./TechnicalSpecificationMain.module.scss');

export default class TechnicalSpecificationMain extends React.Component<ITechnicalSpecificationMainProps, ITechnicalSpecificationMainState, {}> {
  private importFileUploadRef: React.RefObject<HTMLInputElement>;
  private _Provider: ITechnicalSpecificationMainProvider;
  constructor(props: ITechnicalSpecificationMainProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context as any
    });

    this.state = {
      currentUser: null,
      title: null,
      description: null,
      isShowHomePage: false,
      isShowInputForm: false,
      isShowSearchForm:false
    };
    this._onPageLoad = this._onPageLoad.bind(this);
  }
  public componentWillMount(): void {
    this._Provider = new TechnicalSpecificationMainProvider();
  }
  public async componentDidMount() {
    try {
      //alert(wp);
      this._onPageLoad();
    }
    catch (error) {
      this.handlingException(error);
    }
  }
  public _onPageLoad() {
    // this.setState((state, props) => ({
    //   isShowInputForm:true,
    // }));
    let wp = this._Provider.getParameterByName('wp');
    if (wp === 'search') {
      this.setState((state, props) => ({
       
        isShowSearchForm: !state.isShowSearchForm, 
      }));
      console.log(this.state);
    }
    else{
      this.setState((state, props) => ({
        isShowInputForm: !state.isShowInputForm,
      }));
    }
  }
  public handlingException(error) {
    alert("There is an error\n" + error);
    console.log(error);
  }
  public render(): React.ReactElement<ITechnicalSpecificationMainProps> {
    const {
      userDisplayName,
      context,
      Title,
      CurrentUserAccessLevel,
    } = this.props;

    return (
      <div className="row">
        {this.state.isShowInputForm && (
          <TechnicalSpecificationInputFormComponent
            context={this.props.context}
            Title="Input Form"
            userDisplayName=""
          />
        ) ||
         this.state.isShowSearchForm && (
          <TechnicalSpecificationSearchComponent
            context={this.props.context}
            Title="Search"
            userDisplayName=""
          />
        )}

      </div>
    );
  }
}
