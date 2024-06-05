import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import styles from './SearchBarWebPart.module.scss';
import * as strings from 'SearchBarWebPartStrings';
export interface ISearchBarWebPartProps {
  description: string;
  documentLibrary: string;
}

export default class SearchBarWebPart extends BaseClientSideWebPart<ISearchBarWebPartProps> {

  private _libraryOptions: IPropertyPaneDropdownOption[] = [];

  public async onInit(): Promise<void> {
    await this.loadLibraryOptions();
    super.onInit();
  }

  private async loadLibraryOptions(): Promise<void> {
    try {
      const libraries = await this.getDocumentLibraries();
      this._libraryOptions = libraries.map(library => ({
        key: library,
        text: library
      }));
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Error loading library options:', error);
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.searchBar}">
      <input type="text" id="searchInput" placeholder="Enter your search term...">
      <button id="searchButton">Search</button>
      <div id="searchResults" class="${styles.searchresults}"></div>
    </div>`;

    const searchButton = this.domElement.querySelector('#searchButton');
    if (searchButton) {
      searchButton.addEventListener('click', () => this.executeSearch());
    }
  }

  private executeSearch(): void {
    const searchInput = this.domElement.querySelector('#searchInput') as HTMLInputElement;
    const searchTerm = searchInput.value.trim();
    if (searchTerm) {
      this.searchDocuments(searchTerm);
    } else {
      console.error('Search term is empty.');
    }
  }

  private searchDocuments(searchTerm: string): void {
    const documentLibrary = this.properties.documentLibrary;
     
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${documentLibrary}')/items?$select=FileLeafRef,Id,author0`;
    console.log(url);

    const DocumentsArray: [string, string,string,string][] = []; // Array of tuples
    // Use fetch to make the request
    fetch(url, {
            method: 'GET',
            headers: {
                'Accept': 'application/xml;odata=nometadata;charset=utf-8'
            }
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('Error fetching search results: ' + response.statusText);
            }
            return response.text(); // Get the response body as text
        })
        .then(data => {
            // Parse the XML response
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(data, 'text/xml');

        // Extract values of FileLeafRef from XML
        const entries = xmlDoc.getElementsByTagName('entry');
        for (let i = 0; i < entries.length; i++) {
            const entry = entries[i];
            const content = entry.getElementsByTagName('content')[0];
            if (content) {
                const properties = content.getElementsByTagName('m:properties')[0];
                if (properties) {
                    const fileLeafRef = properties.getElementsByTagName('d:FileLeafRef')[0].textContent;
                    var author = properties.getElementsByTagName('d:author0')[0].textContent;
                    var id = properties.getElementsByTagName('d:ID')[0].textContent;
                    if(!author){
                      author = "";
                    }
                    if(!id){
                      id = "";
                    }
                    if (fileLeafRef  ) {
                        const preview = this.getpreview(fileLeafRef);
                        console.log(preview);
                        DocumentsArray.push([fileLeafRef.trim(),author.trim(),id.trim(),preview]);
                    }
                }
            }
        }
        this.renderSearchResults(DocumentsArray);
    })
    .catch(error => {
        console.error('Error executing search:', error);
    });
  
   
    // Render search results
  }
  private getpreview(fileLeafRef: string): string {
    const filePath = encodeURIComponent(`/sites/msteams_274b5c/DocLib5/${fileLeafRef}`);
    const url = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/getpreview.ashx?path=${filePath}`
    return url ;
  }
  private renderSearchResults(DocumentsArray:[string, string,string,string][]): void {
    const searchResultsContainer = this.domElement.querySelector('#searchResults');
    if (!searchResultsContainer) {
      console.error('Search results container not found.');
      return;
    }
  
    let html = '';
  
    DocumentsArray.forEach((Document1) => {
      const fileLeafRef = Document1[0] ;
      // Construct the URL for each file
      const fileUrl = `${this.context.pageContext.web.absoluteUrl}/DocLib5/Forms/AllItems.aspx?id=%2Fsites%2Fmsteams_274b5c%2FDocLib5%2F${encodeURIComponent(fileLeafRef)}&parent=%2Fsites%2Fmsteams_274b5c%2FDocLib5`;
      const author = Document1[1] ;
      const preview = Document1[3] ;
      // Replace 'Document Title' with the actual document title
      const documentTitle = fileLeafRef ; // Replace this with the actual title
  
      // Replace 'Description or additional details' with the actual description
      const documentDescription = author; // Replace this with the actual description
  
      // Construct the HTML for each document
      html += `
        <div class="${styles.document}" id="document">
          <div class="${styles.preview}" id="preview">
            <img src="${preview}" alt="File Preview">
          </div>
          <div id="details" class="${styles.details}" >
            <a href="${fileUrl}" target="_blank">${documentTitle}</a>
            <p>${documentDescription}</p>
             </div>
        </div>
      `;
    });
  
    html += '</div>';
    searchResultsContainer.innerHTML = html;
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
                PropertyPaneDropdown('documentLibrary', {
                  label: 'Select Document Library',
                  options: this._libraryOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async getDocumentLibraries(): Promise<string[]> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 101`;

    try {
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        if (data && data.value) {
          return data.value.map((library: any) => library.Title);
        }
      } else {
        console.error('Error fetching document libraries:', response.statusText);
      }
    } catch (error) {
      console.error('Error fetching document libraries:', error);
    }
    
    return [];
  }
}
