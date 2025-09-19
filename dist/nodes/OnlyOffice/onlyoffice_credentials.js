"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.OnlyOfficeApi = void 0;
class OnlyOfficeApi {
    constructor() {
        this.name = 'onlyOfficeApi';
        this.displayName = 'OnlyOffice API';
        this.documentationUrl = 'https://api.onlyoffice.com/';
        this.properties = [
            {
                displayName: 'Base URL',
                name: 'baseUrl',
                type: 'string',
                default: '',
                placeholder: 'https://your-onlyoffice-instance.com',
                description: 'The base URL of your OnlyOffice instance',
                required: true,
            },
            {
                displayName: 'Access Token',
                name: 'token',
                type: 'string',
                typeOptions: {
                    password: true,
                },
                default: '',
                description: 'The API token for authentication',
                required: true,
            },
        ];
        this.authenticate = {
            type: 'generic',
            properties: {
                headers: {
                    Authorization: '=Bearer {{$credentials.token}}',
                },
            },
        };
    }
}
exports.OnlyOfficeApi = OnlyOfficeApi;
