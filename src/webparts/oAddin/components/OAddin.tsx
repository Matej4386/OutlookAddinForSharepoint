import * as React from 'react';
import styles from './OAddin.module.scss';
import { IOAddinProps } from './IOAddinProps';
import { IOAddinState} from './IOAddinState';
import {sp} from '@pnp/sp';
import {PrimaryButton} from 'office-ui-fabric-react/lib/Button';
import OfficeJs from './modules/OfficeJs/OfficeJs';
import IOfficeJS from './modules/OfficeJs/OfficeJs';

export default class OAddin extends React.Component<IOAddinProps, IOAddinState> {
  private OfficeJs: IOfficeJS;
  constructor(props: IOAddinProps) {
    super(props);
    this.state = {
      attachments: {
        data: [],
        error: ''
      },
      sender: undefined,
      to: undefined,
      normalizedSubject: '',
      body: '',
    }
    this.OfficeJs = new OfficeJs();
  }
  public componentDidUpdate () {
    console.log ('Componet update');
    console.log (this.state);
  }
  public async componentWillMount () {
    const sender = await this.OfficeJs.getSender();
    const to = await this.OfficeJs.getTo();
    const subject = await this.OfficeJs.getSubject(true);
    const body = await this.OfficeJs.getBodyHtml();
    const attachments = await this.OfficeJs.getAttachments();

    this.setState({
      ...this.state,
      sender: sender.data,
      to: [...to.data],
      normalizedSubject: subject.data,
      body: body.data,
      attachments: attachments
    });
    
  }
  public render(): React.ReactElement<IOAddinProps> {
    return(
      <div>
        <div>
          From:
        </div>
        <div>
          {this.state.sender ? this.state.sender.displayName : 'N/A'}
        </div>
        <div>
          To:
        </div>
        <div>
        {
          this.state.to ? 
          this.state.to.map((to)=> {
            return to.displayName
          })
          :
          'N/A'
        }
        </div>
        <div>
          Subject:
        </div>
        <div>
          {this.state.normalizedSubject ? this.state.normalizedSubject : 'N/A'}
        </div>
        <div>
          Body:
        </div>
        <div>
          {
            this.state.body &&
            <div dangerouslySetInnerHTML={{__html: this.state.body}} />
          }
        </div>
        <div>
          Attachments:
        </div>
        <div>
          {
            this.state.attachments.data.length > 0 ?
            this.state.attachments.data.map((attach) => {
              return attach.attachmentsDetails + ','
            })
            :
            'N/A'
          }
        </div>
        <div>
          Attachments content:
        </div>
        <div>
          {
            this.state.attachments.data.length > 0 ? 
            this.state.attachments.data.map((attach) => {
              return attach.attachmentsContent.type
            })
            :
            'N/A'
          }
        </div>
        <PrimaryButton
          onClick={ () => this._save() }
          style={{marginRight: '2rem'}}
          iconProps={{iconName: 'Save'}}
        >
          Save
        </PrimaryButton>
      </div >
    );
  }
  private _save(): void {
      this.saveItem();
  }
  private async saveItem(): Promise<void> {
      try {
      const res = await sp.web.getFolderByServerRelativeUrl("/sites/matejdev/zmluvyent/Zdielane%20dokumenty/").files.add(this.state.attachments.data[0].attachmentsDetails.name, this.state.attachments.data[0].attachmentsContent.content, true);
      console.log (res);
    } catch (error) {
      console.log (error);
    }
  }

}
