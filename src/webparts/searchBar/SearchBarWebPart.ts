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
      <div id="searchResults"  class="${styles.searchresults}"></div>
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
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${documentLibrary}')/items?$select=FileLeafRef,Thumbnail`;
    console.log(url); 
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

            // Extract values of FileLeafRef and Thumbnail from XML
            const files: { name: string, thumbnailUrl: string }[] = [];
            const entries = xmlDoc.getElementsByTagName('entry');
            for (let i = 0; i < entries.length; i++) {
                const entry = entries[i];
                const content = entry.getElementsByTagName('content')[0];
                if (content) {
                    const properties = content.getElementsByTagName('m:properties')[0];
                    if (properties) {
                        const fileLeafRef = properties.getElementsByTagName('d:FileLeafRef')[0]?.textContent?.trim();
                        const idElement = entry.getElementsByTagName('id')[0];
                        if (fileLeafRef && idElement) {
                            const docId = encodeURIComponent(idElement.textContent?.trim()|| '');
                            const accessToken = 'eyJ0eXAiOiJKV1QiLCJub25jZSI6Im1OZHZYVFZtak5JX1h4bVVLN0hMcWluRFhNV195U0F6OW50SmY3ZVdlUTgiLCJhbGciOiJSUzI1NiIsIng1dCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCIsImtpZCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83ODgyMDg1Mi01NWZhLTQ1MGItOTA4ZC00NWMwZDkxMWU3NmIvIiwiaWF0IjoxNzE1NzY4NTYzLCJuYmYiOjE3MTU3Njg1NjMsImV4cCI6MTcxNTg1NTI2MywiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsidXJuOnVzZXI6cmVnaXN0ZXJzZWN1cml0eWluZm8iXSwiYWlvIjoiQVhRQWkvOFdBQUFBU05EMlNLL0pBTEhDY1JscUtyaHMxMm1MUTdMNGVYSmJTcFZETFFORWt2WkJ3WCtFRlJkM1I3Mzh6S3FoTmhVSW5Ca1VudVpaVC8wUjlnSFpmZStiL3lDVFBnbDNGU0N4TjA2eExYSWhDZU5BNXcvUTMyeWc2aFJhSHUzWGRhUklFbWplMG4yV2ZETDNxbUxHS0hibkRBPT0iLCJhbXIiOlsicHdkIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIEV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJjYXBvbGlkc19sYXRlYmluZCI6WyI3ZjA5YWQ4MC0xYzFmLTQ5MmItOTAzNC0zNzU0NjkxZGVkZjIiLCI0NjFmYjQwNy01NGI1LTQxZjQtYTU4OC01NjYwNmQ4ZDZjNTMiLCIwZTczNTFjZS0wYzc3LTQ1ZTgtODlmYi0zZjhkYzc3NjU3ZmMiLCIwODA0YTA2ZS02YWY1LTQ5YWEtYTgzZi01NmQ5ZGU3ZGY1NWMiLCI2MmMwZTNmYy1mNjU5LTQ2MzctOTIzZC1iZDJlODU4YWQ5Y2YiXSwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiODcuNzAuMjkuMjgiLCJuYW1lIjoi15DXk9edINeU15nXmdec15nXkteoIiwib2lkIjoiNGRlMDM1YTMtNGUwZi00NDJiLWE0MGItOGU5MmE2OTQwMWI5IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAzNUE3QkYxMkUiLCJyaCI6IjAuQVRVQVVnaUNlUHBWQzBXUWpVWEEyUkhuYXdNQUFBQUFBQUFBd0FBQUFBQUFBQUExQUtNLiIsInNjcCI6IkFQSUNvbm5lY3RvcnMuUmVhZC5BbGwgQXBwbGljYXRpb24uUmVhZFdyaXRlLkFsbCBDYWxlbmRhcnMuUmVhZFdyaXRlIENoYXQuUmVhZCBDaGF0LlJlYWRCYXNpYyBDb250YWN0cy5SZWFkV3JpdGUgRGV2aWNlTWFuYWdlbWVudEFwcHMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudEFwcHMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50UkJBQy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkLkFsbCBEaXJlY3RvcnkuUmVhZC5BbGwgRmlsZXMuUmVhZFdyaXRlLkFsbCBHcm91cC5SZWFkLkFsbCBHcm91cC5SZWFkV3JpdGUuQWxsIElkZW50aXR5Umlza0V2ZW50LlJlYWQuQWxsIE1haWwuUmVhZCBNYWlsLlJlYWRXcml0ZSBNYWlsYm94U2V0dGluZ3MuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgb3BlbmlkIFBlb3BsZS5SZWFkIFBsYWNlLlJlYWQgUG9saWN5LlJlYWQuQWxsIFByZXNlbmNlLlJlYWQgUHJlc2VuY2UuUmVhZC5BbGwgUHJpbnRlclNoYXJlLlJlYWRCYXNpYy5BbGwgUHJpbnRKb2IuQ3JlYXRlIFByaW50Sm9iLlJlYWRCYXNpYyBwcm9maWxlIFJlcG9ydHMuUmVhZC5BbGwgU2l0ZXMuUmVhZFdyaXRlLkFsbCBUYXNrcy5SZWFkV3JpdGUgVXNlci5SZWFkIFVzZXIuUmVhZC5BbGwgVXNlci5SZWFkQmFzaWMuQWxsIFVzZXIuUmVhZFdyaXRlIFVzZXIuUmVhZFdyaXRlLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IkpiOENteHlaMjhTYURVYzM4anM5eks4UklDMThHd0pVSnZGeE8yMlpVMk0iLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiI3ODgyMDg1Mi01NWZhLTQ1MGItOTA4ZC00NWMwZDkxMWU3NmIiLCJ1bmlxdWVfbmFtZSI6IjMyNjY4MTcwN0BpZGYuaWwiLCJ1cG4iOiIzMjY2ODE3MDdAaWRmLmlsIiwidXRpIjoiWjRoYUo3N2RnVUN3OXROUWthR3BBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19jYyI6WyJDUDEiXSwieG1zX3NzbSI6IjEiLCJ4bXNfc3QiOnsic3ViIjoibGJQM3ZfNklnWk8xcVdsakZQQkVzc0Q3dzZTX0R3SnEwRE9rOGZpUkdSMCJ9LCJ4bXNfdGNkdCI6MTU0OTg5MDQ4NH0.A9puGIUrK9-I_78gJpoNU_gzmpdhHesdUCWnNywJveAF3hqA3GYmsySBErIXgmaRciVS_U5DIkrPWyqvYqEuqeT0WxoIgtCNdSR_oaanpjakQv5N1ka-bmcSUyo_rw-aK6UCP6dphCWGRwK9o7Yo1naDwOC0wj1PK1MWFoqERnOjBmFsqqxklcRujKPWxCMBCamCWswRJHCsLcA8kdM0MPHlG21JAkvYWRIuPZOAvMw-G53nCpsYZhX67T-35PIxlcgV7Mo1jAzzpyeZvpUc96hdPpdTI3727qaYUo6oRzS9A4HwDjSou6HyH6zLcg7gN7r6qxKP3vI33uVD1mdR5A';
                            const thumbnailUrl = `https://southcentralus1-mediap.svc.ms/transform/thumbnail?provider=spo&inputFormat=pdf&cs=fFNQTw&docid=${docId}&access_token=${accessToken}&width=128&height=128`;
                            files.push({
                                name: fileLeafRef,
                                thumbnailUrl: thumbnailUrl
                            });
                        }
                    }
                }
            }

            // Render search results
            this.renderSearchResults(files);
        })
        .catch(error => {
            console.error('Error executing search:', error);
        });
}

  
private renderSearchResults(files: { name: string, thumbnailUrl: string }[]): void {
  const searchResultsContainer = this.domElement.querySelector('#searchResults');
  if (!searchResultsContainer) {
      console.error('Search results container not found.');
      return;
  }

  let html = '';

  files.forEach(file => {
      // Construct the URL for each file
      const fileUrl = `${this.context.pageContext.web.absoluteUrl}/DocLib5/Forms/AllItems.aspx?id=%2Fsites%2Fmsteams_274b5c%2FDocLib5%2F${encodeURIComponent(file.name)}&parent=%2Fsites%2Fmsteams_274b5c%2FDocLib5`;

      // Construct the HTML for each document with preview image
      html += `
          <div class="${styles.document}" id="document">
              <div class="${styles.preview}" id="preview">
                  <img src="${file.thumbnailUrl}" alt="File Preview">
              </div>
              <div id="details" class="${styles.details}" >
                  <a href="${fileUrl}" target="_blank">${file.name}</a>
                  <!-- You can add more details here if needed -->
              </div>
          </div>
      `;
  });

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
