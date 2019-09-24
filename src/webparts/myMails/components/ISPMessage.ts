export interface ISPMessage {
    from_Name: string;
    from_Email: string;
    subject: string; //Email subject
    webLink: string; //Link to email message
    receivedDate: Date; 
}