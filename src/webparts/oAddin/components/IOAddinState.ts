export interface IOAddinState {
    attachments: {
        data: {
            attachmentsDetails: Office.AttachmentDetails;
            attachmentsContent: {
                type: Office.MailboxEnums.AttachmentContentFormat;
                content: any;
            };
        }[];
        error: string;
    };
    sender: Office.EmailAddressDetails;
    to: Office.EmailAddressDetails[];
    normalizedSubject: string;
    body: string;
}