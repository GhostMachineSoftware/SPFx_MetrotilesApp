import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {  
  SPHttpClient, 
  SPHttpClientResponse, 
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './MetrotilesWebPart.module.scss';
import * as strings from 'MetrotilesWebPartStrings';
//*** Custom Imports ***/
//import UsefulLinksHTML from './UsefulLinksHTML';

// import node module external libraries
require('popper.js');
import 'jquery';
import 'bootstrap';
import './styles/custom.css';

export interface IMetrotilesWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  TileName : string;
  TileOrder : number;
  TileDescription: string;
  TileImage: string;
  TileURL : string;
  TileBrowse : string;
}

export default class MetrotilesWebPart extends BaseClientSideWebPart <IMetrotilesWebPartProps> {
  
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Metrotiles')/Items?$orderby=TileOrder",SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private ButtonClick(oldVal: any): any {  
    let currentWebUrl = this.context.pageContext.web.absoluteUrl; 
    window.open(currentWebUrl+'/Lists/Metrotiles/AllItems.aspx','_blank');  
    //return "test"  
  }

  private _renderListAsync(): void {
    // Local environment
  if (Environment.type == EnvironmentType.SharePoint ||
          Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

  private _renderList(items: ISPList[]): void {
    let currentWebUrl = this.context.pageContext.web.absoluteUrl; 
    let tileHTML: string = '';
    let tileCount: number=1;
    const tilesContainer: Element = this.domElement.querySelector('#metrotiles');

    //console.log(currentWebUrl); 
    tilesContainer.innerHTML ='<div class="row">';

    items.forEach((item: ISPList) => {  
/*
      let tileName: string=item.TileName;
      let tileDesc: string=item.TileDescription;
      let tileOrder: number=Math.floor(item.TileOrder);
      let tileLink: string=item.TileURL;
      let tileBrowse: string=item.TileBrowse;      

      if (tileDesc === undefined) { tileDesc = "" };
*/
      tileHTML = `<a href="${item.TileURL}" target="${item.TileBrowse}" class="metrotile text-decoration-none text-center">
          <img class="img-responsive rounded" src="${item.TileImage}">
          <div class="overlay">
          <h5 class="font-weight-bolder text-decoration-none text-uppercase text-white text-center">${item.TileName}</h5>
          <p class="info text-decoration-none text-white font-weight-normal rounded">${item.TileDescription}</p>
          </div>
          </a>`;
      
      if (tileCount % 3 === 0) {
          tilesContainer.innerHTML += `</div>`;
          tilesContainer.innerHTML +='<div class="row">';
      }
      tilesContainer.innerHTML += tileHTML;
      tileCount++;
    });
  }

  public render(): void {
    //let currentWebUrl = this.context.pageContext.web.absoluteUrl; 
    let bootstrapCssURL = "https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css";
    let fontawesomeCssURL = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/regular.min.css";
    SPComponentLoader.loadCss(bootstrapCssURL);
    SPComponentLoader.loadCss(fontawesomeCssURL);
    
    this.domElement.innerHTML = `
      <div id="metrotiles"></div>
    `;
    this._renderListAsync();        
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
                }),            
                PropertyPaneButton('Edit Links', {
                  text: "Edit Links",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.ButtonClick.bind(this)  
                })    
              ]
            }
          ]
        }
      ]
    };
  }
}
