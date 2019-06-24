/**
 * Responsible for showing the edit panel for a componentgit
 */
import * as React from "react";
import * as strings from "ExtensionManagerWebPartStrings";
import {
  CommandBar,
  DefaultButton,
  Panel,
  PanelType,
  PrimaryButton,
  IContextualMenuItem
} from "office-ui-fabric-react";
import { ExtensionService } from '../services/ExtensionService';
import { unescape } from "@microsoft/sp-lodash-subset";
import AceEditor from "react-ace";
import * as ace from "brace";
import "brace/mode/json";
import "brace/theme/github";
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IUserCustomAction } from '../services/IUserCustomAction';
export interface IExtensionPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  isEdit: boolean;
  item?: IUserCustomAction;
  context: IWebPartContext
}

export interface IExtensionPanelState {
  item?: IUserCustomAction
}

const sampleObject: string = "{&quot;sampleTextOne&quot;:&quot;One item is selected in the list.&quot;,"
  + "&quot;sampleTextTwo&quot;:&quot;This command is always visible.&quot;}";

export class ExtensionPanel extends React.Component<IExtensionPanelProps, IExtensionPanelState> {
  public constructor(props) {
    super(props);
    this.state = {
      item: this.props.item
    };
    this.save = this.save.bind(this);
  }
  private _paneCommands: IContextualMenuItem[] = [
    {
      key: "saveItem",
      name: strings.SaveButton,
      icon: "Save",
      ariaLabel: strings.SaveButtonAriaLabel,
      ["data-automation-id"]: "saveButton",
      onClick: this.props.onDismiss,
    },
    {
      key: "cancelItem",
      name: strings.CancelButton,
      icon: "Cancel",
      ariaLabel: strings.CancelButtonAriaLabel,
      ["data-automation-id"]: "cancelButton",
      onClick: this.props.onDismiss,
    }
  ];
  private async editData(id, data) {
    const extServ = new ExtensionService(this.props.context);
    let updatedData = await extServ.editExtension(id, data);
    return updatedData;
  }
  private async save() {
    let { isEdit, item } = this.props;
    if (isEdit && item) {
      await this.editData(item.Id, { ClientSideComponentProperties: item.ClientSideComponentProperties })
    } else {
      //create new customactions
    }
    this.props.onDismiss();
  }
  public componentWillReceiveProps(newProps) {
    if (newProps.item) {
      this.setState({
        item: newProps.item
      });
    }
  }
  public render(): React.ReactElement<IExtensionPanelProps> {

    // automatically convert json string to an object so that we can format the string
    let { isEdit } = this.props;
    let { item } = this.state;
    let sampleObjectClean: string = unescape(sampleObject);
    let jsonObject: any = JSON.parse(sampleObjectClean);
    let jsonString: string = JSON.stringify(jsonObject, null, "\t");
    if (item && isEdit) {
      jsonString = item.ClientSideComponentProperties;
    }

    return (
      <Panel
        isOpen={this.props.isOpen}
        onDismiss={this.props.onDismiss}
        type={PanelType.medium}
        onRenderNavigation={this._onRenderNavigation}
        onRenderFooterContent={this._onRenderFooter}
        headerText={item ? item.Name : "New custom action"}
      >
        {
          this.props.isEdit ? <div>

          </div> : <div>
              New mode
          </div>
        }
        <AceEditor
          mode="json"
          theme="github"
          name="blah2"
          onChange={this._handleJsonChange}
          fontSize={14}
          showPrintMargin={true}
          showGutter={true}
          highlightActiveLine={true}
          value={jsonString}
          setOptions={{
            enableBasicAutocompletion: true,
            enableLiveAutocompletion: true,
            enableSnippets: false,
            showLineNumbers: true,
            tabSize: 2,
          }} />

      </Panel>
    );
  }

  private _handleJsonChange = (value): void => {
    let { item } = this.state;
    item.ClientSideComponentProperties = value;
    this.setState({
      item
    });
  }

  private _onRenderNavigation = (): JSX.Element => {
    return (
      <CommandBar
        isSearchBoxVisible={false}
        items={this._paneCommands}
      />
    );
  }

  private _onRenderFooter = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton
          onClick={this.save}
        >
          {strings.SaveButton}
        </PrimaryButton>
        <DefaultButton
          onClick={this.props.onDismiss}
        >
          {strings.CancelButton}
        </DefaultButton>
      </div>
    );
  }
}
