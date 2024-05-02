import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface ISearchBarWebPartProps {
    description: string;
    documentLibrary: string;
}
export default class SearchBarWebPart extends BaseClientSideWebPart<ISearchBarWebPartProps> {
    private _libraryOptions;
    onInit(): Promise<void>;
    private loadLibraryOptions;
    render(): void;
    private executeSearch;
    private searchDocuments;
    private renderSearchResults;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    private getDocumentLibraries;
}
//# sourceMappingURL=SearchBarWebPart.d.ts.map