import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { sp } from "@pnp/sp";
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as strings from 'CheckInCheckOutMultiFilesCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICheckInCheckOutMultiFilesCommandSetProperties {
  // This is an example; replace with your own properties
  //sampleTextOne: string;
  //sampleTextTwo: string;
}

const LOG_SOURCE: string = 'CheckInCheckOutMultiFilesCommandSet';

export default class CheckInCheckOutMultiFilesCommandSet extends BaseListViewCommandSet<ICheckInCheckOutMultiFilesCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CheckInCheckOutMultiFilesCommandSet');
    SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css');
    
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const CheckInCommand: Command = this.tryGetCommand('CheckIn_Files');
    const CheckOutCommand: Command = this.tryGetCommand('CheckOut_Files');
    
    
    let ShowCheckIn: boolean = false;
    let ShowCheckOut: boolean = false;
    if (event.selectedRows.length > 1) {
      for (let row of event.selectedRows) {
        let CheckoutUserID: string = row.getValueByName('CheckedOutUserId');
        
        if (CheckoutUserID === "") {
          ShowCheckIn = false;
          ShowCheckOut = true;
        }
        else if (CheckoutUserID !== "") {
          ShowCheckIn = true;
          ShowCheckOut = false;
          break;
        }
      }


      if (CheckInCommand) {
        if (ShowCheckIn) {
          CheckInCommand.visible = true;
          CheckOutCommand.visible = false;
        }
        else {
          CheckOutCommand.visible = true;
          CheckInCommand.visible = false;
        }
      }
    }
    else {
      CheckOutCommand.visible = false;
      CheckInCommand.visible = false;
    }

  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    let siteUrl: string = this.context.pageContext.web.absoluteUrl;
    let listName: string = `${this.context.pageContext.list.serverRelativeUrl}`.split("/").pop();
    let ItemurlArr: string[] = [];
    if (event.selectedRows.length > 1) {
      for (let row of event.selectedRows) {
        console.log(row);
        let itemName: string = row.getValueByName('FileLeafRef');
        let itemurl:string= row.getValueByName('FileRef');
        //let fullItemUrl: string = `/${listName}/${itemName}`;
        ItemurlArr.push(itemurl);
      }
    }

    switch (event.itemId) {
      case 'CheckIn_Files':
        this.CheckInFiles(siteUrl, listName, ItemurlArr)
        break;
      case 'CheckOut_Files':
        this.CheckOutFiles(siteUrl, listName, ItemurlArr);
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private CheckOutFiles(siteurl: string, listName: string, ItemurlArr: any) {
    let filecount = 0;
    for (let Item of ItemurlArr) {
      sp.web.getFileByServerRelativeUrl(Item).checkout().then(data => {
        filecount++;
        this.Refreshpage(filecount, ItemurlArr.length);

      }).catch(data => {
        filecount++;
        this.Refreshpage(filecount, ItemurlArr.length);
      });
    }
  }

  private CheckInFiles(siteurl: string, listName: string, ItemurlArr: any) {
    let filecount = 0;
    let checkInComment: string;
    let options = {

    }
    Dialog.prompt(`Enter Check In Comment:`).then((value: string) => {
      checkInComment = value;

      for (let Item of ItemurlArr) {
        sp.web.getFileByServerRelativeUrl(Item).checkin(checkInComment).then(data => {
          filecount++;
          this.Refreshpage(filecount, ItemurlArr.length);
        }).catch(data => {
          filecount++;
          this.Refreshpage(filecount, ItemurlArr.length);
        });
      }
    });
  }

  private GetselectedItems(event: IListViewCommandSetExecuteEventParameters, listName: string): Promise<string[]> {
    return new Promise((resolve) => {
      let ItemurlArr: string[] = [];
      if (event.selectedRows.length > 1) {
        for (let row of event.selectedRows) {
          let itemName: string = row.getValueByName('FileLeafRef');
          let itemurl:string= row.getValueByName('FileRef');
          ItemurlArr.push(itemurl);
        }
        resolve(ItemurlArr);
      }
    });
  }

  private Refreshpage(filecount: number, arrayno: number) {
    if (filecount === arrayno) {
      location.reload();
    }
  }
}
