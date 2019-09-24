import * as React from 'react';
import { DisplayMailsProps } from './IMyMailsProps';
import styles from './MyMails.module.scss';
import { Facepile, IFacepilePersona, IFacepileProps } from 'office-ui-fabric-react/lib/Facepile';
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Label } from 'office-ui-fabric-react/lib/Label';
//https://github.com/microsoftgraph/contoso-airlines-spfx-sample/tree/master/src/webparts/crewBadges/components

export class DisplayMails extends React.Component<DisplayMailsProps,{}> {
    constructor(props: DisplayMailsProps) {
        super(props);
    }

    public render(): React.ReactElement<DisplayMailsProps> {
        return (
            <div>
                {
                    this.props.readyToLoad && this.props.mailsToDisplay.length > 0
                    ?
                        this.props.mailsToDisplay.map((indMail) => 
                            <div className={styles.contentBox}>
                                <div className={styles.itemTitle} onClick={() => this._showMail(indMail.webLink)}> {indMail.subject}</div>
                                <Facepile  
                                        personas={[
                                            {
                                                personaName: indMail.from_Name, 
                                                data: indMail.from_Email,
                                                onClick: this.props.facePileClick
                                            }
                                        ]} 
                                        personaSize={PersonaSize.size40}
                                    />
                                <div className={styles.itemDate}>{this._formatDate(indMail.receivedDate)}</div>
                                <div className={styles.contentBoxLeft}>
                                    
                                </div>
                                <div className={styles.contentBoxRight}>
                                    
                                </div>
                            </div>
                        )
                        
                    :
                        <Label>Outlook is empty</Label>
                        /*<Spinner size={SpinnerSize.large} label="Loading..." ariaLive="assertive" />*/
                }
            </div>
        );
    }

    //On clicking of 'New Mail' button or individual mail item
    private _showMail = (event: any): any => {
        window.open(event, '_blank', 'location=yes,height=570,width=1000, top=150, left=300, scrollbars=yes,status=yes');
    }

    //To format the date
    private _formatDate(recDate: string): string {
        let date: Date = new Date(recDate);
        return date.toLocaleString();
    }
}