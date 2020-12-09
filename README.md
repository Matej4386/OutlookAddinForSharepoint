# OutlookAddinForSharepoint
Outlook addin example manifest for sharepoint 2019 spfx webpart It allows to for example to upload email attachment to Sharepoint library, or create list item with email properties (from, to, body, subject, cc, etc.)  It uses Office.js library. I have created typescript class to use Office.js base functions with await.

My setup is onpremise Exchange and onpremise Sharepoint 2019.

To debug outlook addin in windows 10 use Microsoft Edge DevTools Preview  https://www.microsoft.com/en-us/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot:overviewtab

This is starting point for this kind of solution. In manifest there is just reference to open sharepoint modern site with webpart.
In webpart in config.json add
"externals": {
    "office": {
      "path": "https://appsforoffice.microsoft.com/lib/1/hosted/Office.js",
      "globalName": "office"
    }
 }
 
 in Main webpart.ts add:
 // <reference path="../../../node_modules/@types/office-js/index.d.ts" />
 
 OfficeJs.ts is my class to implement Office.js function with callbacks as async function with await.
 Official documentation for Office.js is a little bit messy (for example to load attachments content and save it to sharepoint library it was a little bit confusing).
 
 This template read outlook email, show basic properties (from, to, body as HTML, subject and attachments) and allows to store first attachment in array to sharepoint library:
 await sp.web.getFolderByServerRelativeUrl("/sites/matejdev/zmluvyent/Zdielane%20dokumenty/").files.add(this.state.attachments.data[0].attachmentsDetails.name, this.state.attachments.data[0].attachmentsContent.content, true);
 
 Each function in OfficeJs is implemented as Promise, for example getAttachments reads all attachments for email and builds a object with attachments details and attachment content:
 data: {
            attachmentsDetails: Office.AttachmentDetails, 
            attachmentsContent: {
                type: Office.MailboxEnums.AttachmentContentFormat, 
                content: any
            }
        }[], 
        error: string
        
 Attachment content is Base64 encoded so before upload to sharepoint it must be decoded for example with this:
 
 
 private getBinaryFileContents (base64FileContents: string) {
        const raw = window.atob(base64FileContents);
        const rawLength = raw.length;
        const array = new Uint8Array(new ArrayBuffer(rawLength));     
        for(let i: number = 0; i < rawLength; i++) {
          array[i] = raw.charCodeAt(i);
        }
        return array;
    }
