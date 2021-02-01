import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset'; 
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http'; 
import { Dialog } from '@microsoft/sp-dialog/lib';
import styles from './FileuploadWebPart.module.scss';
import * as strings from 'FileuploadWebPartStrings';
import * as $ from 'jquery';
import 'jqueryui';
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require('./assets/pagestyles.css');
require('./assets/popup.css');
const saplogo: string = require('./assets/sap.png');
const saphome: string = require('./assets/sapnewhome.png');
export interface IFileuploadWebPartProps {
  description: string;
}
export default class FileuploadWebPart extends BaseClientSideWebPart<IFileuploadWebPartProps> {
  public file: File;
  public allfiles:any[];
  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.min.css');
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css');
    return super.onInit();
  }
  public render(): void {
    this.getuploadfiles();
    window['webPartContext']=this.context;
    this.domElement.innerHTML = `
     
    <div id="dialog" class="modal">
        <div class="header-bar">
            <div class="heading">
            <i class="fa fa-edit edit-icon"></i>
            <span>Archive from frontend</span>
        </div>
    </div>
    <div class="info-bar1">
        <span><u>Scenario: </u> &nbsp;&nbsp;&nbsp;&nbsp;Assign then Store</span>   
    </div>
    <div class="document-heading">
        <span class="heading">Document Type</span>
    </div>
   <div class="document-list">
        <div style="font-size: small;">
        <i class="fa fa-caret-down"></i>
        <i style="color: #e1ad01;" class="fa fa-folder-open" aria-hidden="true"></i>
        Standard Material
        <ul style="margin-top: 0px;" id="alldocuments"></ul>
    </div>
    <div class="file">
        <input style="font-size: small;" type="file" id="uploadFile" value="Upload File" />
        <input style="font-size: small;" type="button" class="uploadButton" value="Upload" /> 
    </div>
</div>
    <div class="end-icons">
        <i style="color: green;" class="fa fa-check-square-o" aria-hidden="true"></i>
        <i style="color: blue;" class="fa fa-undo"></i>
        <i style="color: orangered;" class="fa fa-times" aria-hidden="true"></i>
    </div>

</div>
      
       <!-- <div class="fixed-header">
        <div class="header-bar"></div>
        <div class="menu-bar">
            <ul class="menu-list">
                <i class="fa fa-edit edit-icon"></i>
                <li>Material</li>
                <li>Edit</li>
                <li>GoTo</li>
                <li>Environment</li>
                <li>System</li>
                <li>Help</li>
                <span class="window-icons">
                    <i class="fa fa-window-minimize" aria-hidden="true"></i>
                    <i class="fa fa-window-maximize"></i>
                    <i class="fa fa-times" aria-hidden="true"></i>
                </span>
            </ul>
        </div>
        <div>
            <img src="html-pages/header.jpg" height="60px" width="100%">
        </div> -->
    <div><img src='${saphome}' height="25px" width="100%">
    </div>
    <div class="info-bar">
    <div>
        <nav>
            <ul>
            <li class="submenu"><a href="#">menu</a>
                <ul>
                <li><a href="#">IT Solutions</a></li>
                <li class="submenu"><a href="#">create..</a>
                    <ul>
                        <li class="showpopup"><a href="#" ><span class="t">create Attachment</span></a></li>
                        <li>Scan Attachment</li>
                        <li>Create Note</li>
                        <li>Create external document (URL)</li>
                        <li>Store business document</li>
                        <li>Enter Bar Code</li>
                    </ul>
                </li>
                <li class="disable">Attachment List</li>
                <li>Linked Service Requests</li>
                <li>Private note</li>
                <li>Send <i class="fa fa-caret-right" aria-hidden="true"></i></li>
                <li class="disable">Relationships</li>
                <li>Workflow <i class="fa fa-caret-right" aria-hidden="true"></i></li>
                <li>My Objects <i class="fa fa-caret-right" aria-hidden="true"></i></li>
                <li>Help for object services</li>
                <li>Site Records</li>
                <li class="disable">Documents</li>
                </ul>
            </li>
            </ul>
        </nav>
    </div>
        <span class="info-text">Display Material 300-400 (Finished-product)</span>
    </div>
    <div class="info-bar second-info-bar">
        <ul class="info-list">
            <li><i class="fa fa-cube" aria-hidden="true"></i></i></li>
            <li><span><i class="fa fa-arrow-right" style="color: #3498DB;" aria-hidden="true"></i>Additional Data</span></li>
            <li><i style="color: green;" class="fa fa-copy"></i></i>Org. Level</li>
        </ul>
    </div>
    </div>

    <div class="form">
        <ul class="tab-list">
            <li><i class="fa fa-copy"></i>&nbsp;&nbsp;Basic Data 1</li>
            <li>Basic Data 2</li>
            <li>Sales: sales org.1</li>
            <li>Sales: sales org.1</li>
            <li>Sales: G...</li>
            &nbsp;
            <i class="fa fa-caret-left" aria-hidden="true"></i>&nbsp;
            <i class="fa fa-caret-right" aria-hidden="true"></i>&nbsp;
            <i class="fa fa-folder-o" aria-hidden="true"></i>
            </ul>
        <div>
            <label>Material</label>
            <input type="text" value="300-400">
            &nbsp;
            <i class="fa fa-copy"></i>
            <input style="width: 250px;" type="text" value="electronic (2-step packing material)">
            <i class="fa fa-info-circle" aria-hidden="true"></i>
        </div>
        <div class="general-data">
            <h5>General Data</h5>
            <hr>
            <div class="general-data-input">
                <div class="input">
                    <label class="input-name">Base Unit of Measure</label>
                    <input style="width: 30px;" type="text" value="PC">
                    <span>piece(s)</span>
                </div>
                <div class="input">
                    <label class="input-name">Material Group</label>
                    <input style="width: 100px;" type="text" value="001"><br>
                </div>
                <div class="input">

                    <label class="input-name">Old material number</label>
                    <input style="width: 100px;" type="text">
                </div>
                <div class="input">

                    <label class="input-name">Ext. Matl. Group</label>
                    <input style="width: 100px;" type="text"><br>
                </div>
                <div class="input">

                    <label class="input-name">Division</label>
                    <input style="width: 100px;" type="text" value="01">
                </div>
                <div class="input">

                    <label class="input-name">Lab/Office</label>
                    <input style="width: 100px;" type="text"><br>
                </div>
                <div class="input">

                    <label class="input-name">Product Allocation</label>
                    <input style="width: 100px;" type="text">
                </div>
                <div class="input">

                    <label class="input-name">Prod. Hierarchy</label>
                    <input style="width: 100px;" type="text" value="001001001050"><br>
                </div>
                <div class="input">

                    <label class="input-name">X-Plant Matl. Status</label>
                    <input style="width: 100px;" type="text">
                </div>
                <div class="input">

                    <label class="input-name">Valid from</label>
                    <input style="width: 100px;" type="text"><br>
                </div>
                <div class="input">
                    <input type="checkbox">
                    <label> Assign effect vals.</label><br>
                </div>
                <div class="input">

                    <label class="input-name">GenItemCatGroup</label>
                    <input style="width: 100px;" type="text" value="NORM">
                </div>
            </div>
        </div>

        <div class="authorization-group">
            <h5>Material Authorization Group</h5>
            <hr>
            <div class="input">
                <label class="input-name">Authorization Group</label>
                <input style="width: 50px;" type="text">
            </div>
        </div>

        <div class="dimensions">
            <h5>Dimensions/EANs</h5>
            <hr>
            <div class="general-data-input">

                <div class="input">

                    <label class="input-name">Gross Weight</label>
                    <input style="width: 100px;" type="text" value="1,800">
                </div>
                <div class="input">
                    <label class="input-name">Weight Unit</label>
                    <input style="width: 100px;" type="text" value="KG"><br>
                </div>
                <div class="input">
                    <label class="input-name">Net Weight</label>
                    <input style="width: 100px;" type="text" value="1,800"><br>
                </div>
                <div class="input">
                    <label class="input-name">Volume</label>
                    <input style="width: 100px;" type="text" value="1,000">
                </div>
                <div class="input">
                    <label class="input-name">Volume Unit</label>
                    <input style="width: 100px;" type="text">
                </div>
                <div class="input">
                    <label class="input-name">Size/Dimensions</label>
                    <input style="width: 100px;" type="text"><br>
                </div>
                <div class="input">
                    <label class="input-name">EAN/UPC</label>
                    <input style="width: 100px;" type="text">
                </div>
                <div class="input">
                    <label class="input-name">EAN Category</label>
                    <input style="width: 100px;" type="text">
                </div>
            </div>
        </div>

        <div class="packaging-material">
            <h5>Packaging material data</h5>
            <hr>
            <div class="input">
                <label class="input-name">Matl. Group Pack. Matls.</label>
                <input style="width: 30px;" type="text"><br>
            </div>
            <div class="input">
                <label class="input-name">Ref. mat for pack.</label>
                <input style="width: 100px;" type="text">
            </div>
        </div>

        <div class="data-texts">
            <h5>Basic Data Texts</h5>
            <hr>
            <div class="general-data-input">

                <div style="width: 280px; display: inline-flex;">
                    <label class="input-name">Languages Maitained</label>
                    <span style="width: 20px;">0</span>
                    <div class="input-group"> 
                    <input class="input-field" style="width: 100px;" type="text" value="Basic Data Text">
                    <i class="fa fa-copy icon"></i>
                    </div>
                </div>
                &nbsp;&nbsp;
                <div style="width: 220px;">
                    <label class="input-name">Language</label>
                    <select style="width: 80px;">
                        <option value=""></option>
                    </select>
                </div>
            </div>
        </div>
    </div>
    <div class="fixed-footer">
        <div class="footer-bar">
            <img class="center" src='${saplogo}' height="15px" width="30px">
            <div class="footer-menu-list">
            <li><i class="fa fa-angle-double-right"></i></li>
            <li><span>VCS(1)800&nbsp; <i class="fa fa-caret-down"></i> </span></li>
            <li><span>VENEDIG</span></li>
            <li><span>INS</span></li>
            <li><span>&nbsp;&nbsp;</span></li>
            <li><i class="fa fa-exchange-alt"></i></li>
            <li><i class="fa fa-lock"></i></li>
            </div>
        </div>
        <div class="header-bar"></div>
    </div>`;
      // tslint:disable-next-line: no-function-expression
      var self = this;
      $(document).ready(() => {
        //Dialog.alert('Hello world');
        $("#dialog" ).hide();
    });
      this.setButtonsEventHandlers(); 
  }
  private setButtonsEventHandlers(): void {  
    const webPart: FileuploadWebPart = this;  
    this.domElement.getElementsByClassName('uploadButton')[0].
    addEventListener('click', () => { this.UploadFiles(); });
    this.domElement.getElementsByClassName('showpopup')[0].
    addEventListener('click', () => { this.showpopup(); });
   
  } 
  public showpopup()
  {
    $("#dialog" ).show();
    $("#dialog" ).dialog({
      maxWidth:560,
      maxHeight: 450,
      width: 560,
      height: 450,
      modal: true
    });
  }
 public UploadFiles()
 {
  var files = (<HTMLInputElement>document.getElementById('uploadFile')).files;
  //in case of multiple files,iterate or else upload the first file.
  var file = files[0];
  if (file != undefined || file != null) {
    let spOpts : ISPHttpClientOptions  = {
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json"
      },
      body: file        
    };
  
    var url = `${this.context.pageContext.web.absoluteUrl}/_api/Web/Lists/getByTitle('Documents')/RootFolder/Files/Add(url='${file.name}', overwrite=true)`;
  
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
  
      console.log(`Status code: ${response.status}`);
      console.log(`Status text: ${response.statusText}`);
      this.getuploadfiles();
  
      response.json().then((responseJSON: JSON) => {
        console.log(responseJSON);
        this.clearfileupload();
      });
    });
  
  }
 }
 public clearfileupload()
 {
       //var control = document.getElementById('uploadFile');
       $("#uploadFile").val("");
    
 }
 public getuploadfiles()
 {
  let requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/Lists/GetByTitle('Documents')/items?$select=File&$expand=File`;   
  
  this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {  
          if (response.ok) {  
              response.json().then((responseJSON) => {  
                  if (responseJSON!=null && responseJSON.value!=null){  
                       this.allfiles = responseJSON.value;  
                       var treeview = ``;
                        treeview += `<ul>`;
                       this.allfiles.forEach(item => {
                        if (item) {
                          var filelink = item.File.LinkingUrl; 
                          filelink = filelink.replace(/ /g, '%20');
                          treeview += `<li><i class="fa fa-file-pdf-o" aria-hidden="true"></i> &nbsp;<a style="text-decoration:none; color: black;" href=`+ filelink +`>` + item.File.Name + `</a></li>`;
                        } 
                    });
                      treeview += `</ul>`;
                      $("#alldocuments").html('');
                      $("#alldocuments").append(treeview);
                  }  
              });  
          }  
      });
 }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

}
