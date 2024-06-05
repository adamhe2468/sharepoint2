var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './SearchBarWebPart.module.scss';
import * as strings from 'SearchBarWebPartStrings';
var SearchBarWebPart = /** @class */ (function (_super) {
    __extends(SearchBarWebPart, _super);
    function SearchBarWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._libraryOptions = [];
        return _this;
    }
    SearchBarWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.loadLibraryOptions()];
                    case 1:
                        _a.sent();
                        _super.prototype.onInit.call(this);
                        return [2 /*return*/];
                }
            });
        });
    };
    SearchBarWebPart.prototype.loadLibraryOptions = function () {
        return __awaiter(this, void 0, void 0, function () {
            var libraries, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getDocumentLibraries()];
                    case 1:
                        libraries = _a.sent();
                        this._libraryOptions = libraries.map(function (library) { return ({
                            key: library,
                            text: library
                        }); });
                        this.context.propertyPane.refresh();
                        return [3 /*break*/, 3];
                    case 2:
                        error_1 = _a.sent();
                        console.error('Error loading library options:', error_1);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    SearchBarWebPart.prototype.render = function () {
        var _this = this;
        this.domElement.innerHTML = "\n    <div class=\"".concat(styles.searchBar, "\">\n      <input type=\"text\" id=\"searchInput\" placeholder=\"Enter your search term...\">\n      <button id=\"searchButton\">Search</button>\n      <div id=\"searchResults\" class=\"").concat(styles.searchresults, "\"></div>\n    </div>");
        var searchButton = this.domElement.querySelector('#searchButton');
        if (searchButton) {
            searchButton.addEventListener('click', function () { return _this.executeSearch(); });
        }
    };
    SearchBarWebPart.prototype.executeSearch = function () {
        var searchInput = this.domElement.querySelector('#searchInput');
        var searchTerm = searchInput.value.trim();
        if (searchTerm) {
            this.searchDocuments(searchTerm);
        }
        else {
            console.error('Search term is empty.');
        }
    };
    SearchBarWebPart.prototype.searchDocuments = function (searchTerm) {
        var _this = this;
        var documentLibrary = this.properties.documentLibrary;
        var url = "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists/getByTitle('").concat(documentLibrary, "')/items?$select=FileLeafRef,Id,author0");
        console.log(url);
        var DocumentsArray = []; // Array of tuples
        // Use fetch to make the request
        fetch(url, {
            method: 'GET',
            headers: {
                'Accept': 'application/xml;odata=nometadata;charset=utf-8'
            }
        })
            .then(function (response) {
            if (!response.ok) {
                throw new Error('Error fetching search results: ' + response.statusText);
            }
            return response.text(); // Get the response body as text
        })
            .then(function (data) {
            // Parse the XML response
            var parser = new DOMParser();
            var xmlDoc = parser.parseFromString(data, 'text/xml');
            // Extract values of FileLeafRef from XML
            var entries = xmlDoc.getElementsByTagName('entry');
            for (var i = 0; i < entries.length; i++) {
                var entry = entries[i];
                var content = entry.getElementsByTagName('content')[0];
                if (content) {
                    var properties = content.getElementsByTagName('m:properties')[0];
                    if (properties) {
                        var fileLeafRef = properties.getElementsByTagName('d:FileLeafRef')[0].textContent;
                        var author = properties.getElementsByTagName('d:author0')[0].textContent;
                        var id = properties.getElementsByTagName('d:ID')[0].textContent;
                        if (!author) {
                            author = "";
                        }
                        if (!id) {
                            id = "";
                        }
                        if (fileLeafRef) {
                            var preview = _this.getpreview(fileLeafRef);
                            console.log(preview);
                            DocumentsArray.push([fileLeafRef.trim(), author.trim(), id.trim(), preview]);
                        }
                    }
                }
            }
            _this.renderSearchResults(DocumentsArray);
        })
            .catch(function (error) {
            console.error('Error executing search:', error);
        });
        // Render search results
    };
    SearchBarWebPart.prototype.getpreview = function (fileLeafRef) {
        var filePath = encodeURIComponent("/sites/msteams_274b5c/DocLib5/".concat(fileLeafRef));
        var url = "".concat(this.context.pageContext.web.absoluteUrl, "/_layouts/15/getpreview.ashx?path=").concat(filePath);
        return url;
    };
    SearchBarWebPart.prototype.renderSearchResults = function (DocumentsArray) {
        var _this = this;
        var searchResultsContainer = this.domElement.querySelector('#searchResults');
        if (!searchResultsContainer) {
            console.error('Search results container not found.');
            return;
        }
        var html = '';
        DocumentsArray.forEach(function (Document1) {
            var fileLeafRef = Document1[0];
            // Construct the URL for each file
            var fileUrl = "".concat(_this.context.pageContext.web.absoluteUrl, "/DocLib5/Forms/AllItems.aspx?id=%2Fsites%2Fmsteams_274b5c%2FDocLib5%2F").concat(encodeURIComponent(fileLeafRef), "&parent=%2Fsites%2Fmsteams_274b5c%2FDocLib5");
            var author = Document1[1];
            var preview = Document1[3];
            // Replace 'Document Title' with the actual document title
            var documentTitle = fileLeafRef; // Replace this with the actual title
            // Replace 'Description or additional details' with the actual description
            var documentDescription = author; // Replace this with the actual description
            // Construct the HTML for each document
            html += "\n        <div class=\"".concat(styles.document, "\" id=\"document\">\n          <div class=\"").concat(styles.preview, "\" id=\"preview\">\n            <img src=\"").concat(preview, "\" alt=\"File Preview\">\n          </div>\n          <div id=\"details\" class=\"").concat(styles.details, "\" >\n            <a href=\"").concat(fileUrl, "\" target=\"_blank\">").concat(documentTitle, "</a>\n            <p>").concat(documentDescription, "</p>\n             </div>\n        </div>\n      ");
        });
        html += '</div>';
        searchResultsContainer.innerHTML = html;
    };
    Object.defineProperty(SearchBarWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    SearchBarWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    SearchBarWebPart.prototype.getDocumentLibraries = function () {
        return __awaiter(this, void 0, void 0, function () {
            var url, response, data, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists?$filter=BaseTemplate eq 101");
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 6, , 7]);
                        return [4 /*yield*/, this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) return [3 /*break*/, 4];
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        if (data && data.value) {
                            return [2 /*return*/, data.value.map(function (library) { return library.Title; })];
                        }
                        return [3 /*break*/, 5];
                    case 4:
                        console.error('Error fetching document libraries:', response.statusText);
                        _a.label = 5;
                    case 5: return [3 /*break*/, 7];
                    case 6:
                        error_2 = _a.sent();
                        console.error('Error fetching document libraries:', error_2);
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/, []];
                }
            });
        });
    };
    return SearchBarWebPart;
}(BaseClientSideWebPart));
export default SearchBarWebPart;
//# sourceMappingURL=SearchBarWebPart.js.map