/**
 * Add to config.json:
 * "externals": {
    "office": {
      "path": "https://appsforoffice.microsoft.com/lib/1/hosted/Office.js",
      "globalName": "office"
    }
  },

  to main .ts WebPart file insert:
  /// <reference path="../../../node_modules/@types/office-js/index.d.ts" />
 */
export interface IOfficeJS {
    getItemType(): Promise<{data: string, error: string}>;
    getItem(): Promise<{data: any, error: string}>;
    getSender(): Promise<{data: Office.EmailAddressDetails, error: string}>;
    getTo(): Promise<{data: Office.EmailAddressDetails[], error: string}>;
    getCc(): Promise<{data: Office.EmailAddressDetails[], error: string}>;
    getSubject(normalized: boolean): Promise<{data: string, error: string}>;
    getBodyHtml(): Promise<{data: string, error: string}>;
    getBodyText(): Promise<{data: string, error: string}>;
    getAttachments(): Promise<{
        data: {
            attachmentsDetails: Office.AttachmentDetails, 
            attachmentsContent: {
                type: Office.MailboxEnums.AttachmentContentFormat, 
                content: any
            }
        }[], 
        error: string
    }>;
}
export default class OfficeJs implements IOfficeJS {
    private init = (): Promise<any> => {
        return new Promise<any>((resolve, reject) => {
            require(['office'], () => {
                Office.onReady((info) => {
                    if (info.host === Office.HostType.Outlook) {
                        typeof Office.context === 'undefined' || typeof Office.context.mailbox === 'undefined' || typeof Office.context.mailbox.item === 'undefined' ?
                            reject('Office integration has not been initialized.')
                            : 
                            resolve('Success'); 
                    }
                });
            });
        });
    }
    private getBinaryFileContents (base64FileContents: string) {
        const raw = window.atob(base64FileContents);
        const rawLength = raw.length;
        const array = new Uint8Array(new ArrayBuffer(rawLength));
      
        for(let i: number = 0; i < rawLength; i++) {
          array[i] = raw.charCodeAt(i);
        }
        return array;
    }
    public getItemType = async (): Promise<{data: string, error: string}> => {
        try {
            await this.init();
            const item = Office.context.mailbox.item;            
            return {
                data: item.itemType,
                error: ''
            };
        } catch (error) {
            return {
                data: undefined,
                error: error
            };
        }
    }
    public getItem = async (): Promise<{data: any, error: string}> => {
        try {
            await this.init();
            const item = Office.context.mailbox.item;            
            return {
                data: item,
                error: ''
            };
        } catch (error) {
            return {
                data: undefined,
                error: error
            };
        }
    }
    public getSender = async (): Promise<{data: Office.EmailAddressDetails, error: string}> => {
        try {
            await this.init();
            const item = Office.context.mailbox.item;
            return {
                data: item.sender,
                error: ''
            };
            
        } catch (error) {
            return {
                data: undefined,
                error: error
            };
        }
    }

    public getTo = async (): Promise<{data: Office.EmailAddressDetails[], error: string}> => {
        try {
            await this.init();
            const item = Office.context.mailbox.item;
            return {
                data: [...item.to],
                error: ''
            };
        } catch (error) {
            return {
                data: undefined,
                error: error
            };
        }
    }
    public getCc = async (): Promise<{data: Office.EmailAddressDetails[], error: string}> => {
        try {
            await this.init();
            const item = Office.context.mailbox.item;
            return {
                data: [...item.cc],
                error: ''
            };
        } catch (error) {
            return {
                data: undefined,
                error: error
            };
        }
    }
    public getSubject = async (normalized: boolean): Promise<{data: string, error: string}> => {
        try {
            await this.init();
            const item = Office.context.mailbox.item;
            return {
                data: normalized ? item.normalizedSubject : item.subject,
                error: ''
            };
        } catch (error) {
            return {
                data: undefined,
                error: error
            };
        }
    }
    public getBodyHtml = async (): Promise<{data: string, error: string}> => {
        await this.init();
        return new Promise<any>((resolve, reject) => {
            try {
                const item = Office.context.mailbox.item;
                item.body.getAsync(
                    'html',
                    {asyncContext: undefined},
                    (bodyAsync: Office.AsyncResult<string>) => {
                        if (bodyAsync.status === Office.AsyncResultStatus.Succeeded) {
                            resolve ({
                                data: bodyAsync.value,
                                error: ''
                            });
                        } else {
                            reject({
                                data: undefined,
                                error: bodyAsync.error.message
                            });
                        }
                    }
                );
            } catch (error) {
                reject({
                    data: undefined,
                    error: error
                });
            }
        });
    }
    public getBodyText = async (): Promise<{data: string, error: string}> => {
        await this.init();
        return new Promise<any>((resolve, reject) => {
            try {
                const item = Office.context.mailbox.item;
                item.body.getAsync(
                    'text',
                    {asyncContext: undefined},
                    (bodyAsync: Office.AsyncResult<string>) => {
                        if (bodyAsync.status === Office.AsyncResultStatus.Succeeded) {
                            resolve ({
                                data: bodyAsync.value,
                                error: ''
                            });
                        } else {
                            reject({
                                data: undefined,
                                error: bodyAsync.error.message
                            });
                        }
                    }
                );
            } catch (error) {
                reject({
                    data: undefined,
                    error: error
                });
            }
        });
    }
    public getAttachments = async (): Promise<{data: {attachmentsDetails: Office.AttachmentDetails, attachmentsContent: {type: Office.MailboxEnums.AttachmentContentFormat, content: any}}[], error: string}> => {
        await this.init();
        return new Promise<any>((resolve, reject) => {
            try {
                const item = Office.context.mailbox.item;
                if (item.attachments.length > 0) {
                    const promises: any[] = [];
                    for (let i: number = 0; i < item.attachments.length; i++) {
                        promises.push(new Promise<any>((resolve2, reject2) => {item.getAttachmentContentAsync(
                            item.attachments[i].id,
                            {asyncContext: item.attachments[i]},
                            (result: Office.AsyncResult<Office.AttachmentContent>) => {
                                console.log (result);
                                if (result.status === Office.AsyncResultStatus.Succeeded) {
                                    switch (result.value.format) {
                                        case Office.MailboxEnums.AttachmentContentFormat.Base64:
                                            // Handle file attachment.
                                            resolve2({
                                                attachmentsDetails: result.asyncContext,
                                                attachmentsContent: {
                                                    type: Office.MailboxEnums.AttachmentContentFormat.Base64,
                                                    content: this.getBinaryFileContents(result.value.content)
                                                }
                                            });
                                            break;
                                        case Office.MailboxEnums.AttachmentContentFormat.Eml:
                                            // Handle email item attachment.
                                            resolve2({
                                                attachmentsDetails: result.asyncContext,
                                                attachmentsContent: {
                                                    type: Office.MailboxEnums.AttachmentContentFormat.Eml,
                                                    content: 'email'
                                                }
                                            });
                                            break;
                                        case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
                                            // Handle .icalender attachment.
                                            resolve2({
                                                attachmentsDetails: result.asyncContext,
                                                attachmentsContent: {
                                                    type: Office.MailboxEnums.AttachmentContentFormat.Eml,
                                                    content: 'icalender'
                                                }
                                            });
                                            break;
                                        case Office.MailboxEnums.AttachmentContentFormat.Url:
                                            // Handle cloud attachment.
                                            resolve2 ({
                                                attachmentsDetails: result.asyncContext,
                                                attachmentsContent: {
                                                    type: Office.MailboxEnums.AttachmentContentFormat.Url,
                                                    content: 'URL'
                                                }
                                            });
                                            break;
                                        default:
                                            resolve2({
                                                attachmentsDetails: result.asyncContext,
                                                attachmentsContent: undefined
                                            });
                                            break;
                                    }
                                } else {
                                    reject2({
                                        data: undefined,
                                        error: result.error.message
                                    });
                                }
                            } 
                        );}));
                    }
                    Promise.all(promises)
                        .then((results) => {
                            console.log (results);
                            resolve({
                                data: results,
                                error: ''
                            });
                        })
                        .catch((e) => {
                            reject({
                                data: undefined,
                                error: e
                            });
                        });
                } else {
                    resolve({
                        data: [],
                        error: ''
                    });
                }
            } catch (error) {
                reject({
                    data: undefined,
                    error: error
                });
            }
        });
    }
}