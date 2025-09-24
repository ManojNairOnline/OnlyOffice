import {
  IExecuteFunctions,
  INodeExecutionData,
  INodeType,
  INodeTypeDescription,
  NodeOperationError,
  NodeConnectionType,
} from 'n8n-workflow';

export class OnlyOffice implements INodeType {
  description: INodeTypeDescription = {
    displayName: 'OnlyOffice',
    name: 'onlyOffice',
    icon: 'file:onlyoffice.svg',
    group: ['transform'],
    version: 1,
    subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
    description: 'Interact with OnlyOffice files and folders',
    defaults: {
      name: 'OnlyOffice',
    },
    inputs: [NodeConnectionType.Main],
    outputs: [NodeConnectionType.Main],
    credentials: [
      {
        name: 'onlyOfficeApi',
        required: true,
      },
    ],
    requestDefaults: {
      baseURL: '={{$credentials.baseUrl}}/api/2.0',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
    },
    properties: [
      {
        displayName: 'Resource',
        name: 'resource',
        type: 'options',
        noDataExpression: true,
        options: [
          {
            name: 'Folder',
            value: 'folder',
          },
          {
            name: 'File',
            value: 'file',
          },
        ],
        default: 'folder',
      },

      // Folder Operations
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: {
            resource: ['folder'],
          },
        },
        options: [
          {
            name: 'List',
            value: 'list',
            action: 'List folders',
          },
          {
            name: 'Create',
            value: 'create',
            action: 'Create folder',
          },
          {
            name: 'Rename',
            value: 'rename',
            action: 'Rename folder',
          },
          {
            name: 'Move',
            value: 'move',
            action: 'Move folder',
          },
          {
            name: 'Copy',
            value: 'copy',
            action: 'Copy folder',
          },
          {
            name: 'Delete',
            value: 'delete',
            action: 'Delete folder',
          },
        ],
        default: 'list',
      },

      // File Operations
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: {
            resource: ['file'],
          },
        },
        options: [
          {
            name: 'List',
            value: 'list',
            action: 'List files',
          },
          {
            name: 'Create',
            value: 'create',
            action: 'Create file',
          },
          {
            name: 'Rename',
            value: 'rename',
            action: 'Rename file',
          },
          {
            name: 'Move',
            value: 'move',
            action: 'Move file',
          },
          {
            name: 'Copy',
            value: 'copy',
            action: 'Copy file',
          },
          {
            name: 'Delete',
            value: 'delete',
            action: 'Delete file',
          },
        ],
        default: 'list',
      },

      // List Operations - Folder ID
      {
        displayName: 'Folder ID',
        name: 'folderId',
        type: 'string',
        default: '@my',
        required: true,
        displayOptions: {
          show: {
            operation: ['list'],
          },
        },
        description: 'ID of the folder to list contents from. Use @my for My Documents, @common for Common Documents.',
      },

      // Create Operations
      {
        displayName: 'Parent Folder ID',
        name: 'parentFolderId',
        type: 'string',
        default: '@my',
        required: true,
        displayOptions: {
          show: {
            operation: ['create'],
          },
        },
        description: 'ID of the parent folder where the new item will be created',
      },
      {
        displayName: 'Title',
        name: 'title',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['create'],
          },
        },
        description: 'Name of the new folder or file',
      },

      // File Type for Create File
      {
        displayName: 'File Type',
        name: 'fileType',
        type: 'options',
        options: [
          {
            name: 'Document (.docx)',
            value: 'docx',
          },
          {
            name: 'Spreadsheet (.xlsx)',
            value: 'xlsx',
          },
          {
            name: 'Presentation (.pptx)',
            value: 'pptx',
          },
        ],
        default: 'docx',
        required: true,
        displayOptions: {
          show: {
            resource: ['file'],
            operation: ['create'],
          },
        },
        description: 'Type of file to create',
      },

      // Operations requiring Item ID
      {
        displayName: 'Item ID',
        name: 'itemId',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['rename', 'move', 'copy', 'delete'],
          },
        },
        description: 'ID of the folder or file to operate on',
      },

      // Rename Operation
      {
        displayName: 'New Title',
        name: 'newTitle',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['rename'],
          },
        },
        description: 'New name for the folder or file',
      },

      // Move/Copy Operations
      {
        displayName: 'Destination Folder ID',
        name: 'destFolderId',
        type: 'string',
        default: '',
        required: true,
        displayOptions: {
          show: {
            operation: ['move', 'copy'],
          },
        },
        description: 'ID of the destination folder',
      },
      {
        displayName: 'Conflict Resolution',
        name: 'conflictResolveType',
        type: 'options',
        options: [
          {
            name: 'Skip',
            value: 'Skip',
          },
          {
            name: 'Overwrite',
            value: 'Overwrite',
          },
          {
            name: 'Duplicate',
            value: 'Duplicate',
          },
        ],
        default: 'Skip',
        displayOptions: {
          show: {
            operation: ['move', 'copy'],
          },
        },
        description: 'How to handle conflicts when moving or copying',
      },

      // Delete Options
      {
        displayName: 'Delete Immediately',
        name: 'deleteImmediately',
        type: 'boolean',
        default: false,
        displayOptions: {
          show: {
            operation: ['delete'],
          },
        },
        description: 'Whether to delete immediately or move to trash',
      },
    ],
  };

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData();
    const returnData: INodeExecutionData[] = [];

    for (let i = 0; i < items.length; i++) {
      try {
        const resource = this.getNodeParameter('resource', i) as string;
        const operation = this.getNodeParameter('operation', i) as string;

        let responseData;

        if (resource === 'folder') {
          responseData = await OnlyOffice.executeFolderOperation(this, operation, i);
        } else if (resource === 'file') {
          responseData = await OnlyOffice.executeFileOperation(this, operation, i);
        }

        if (Array.isArray(responseData)) {
          returnData.push(...responseData.map(item => ({ json: item })));
        } else {
          returnData.push({ json: responseData });
        }
      } catch (error) {
        if (this.continueOnFail()) {
          returnData.push({
            json: {
              error: error instanceof Error ? error.message : String(error),
            },
          });
          continue;
        }
        throw error;
      }
    }

    return [returnData];
  }

  private static async executeFolderOperation(context: IExecuteFunctions, operation: string, itemIndex: number): Promise<any> {
    const credentials = await context.getCredentials('onlyOfficeApi');
    const baseUrl = `${credentials.baseUrl}/api/2.0`;

    switch (operation) {
      case 'list':
        const folderId = context.getNodeParameter('folderId', itemIndex) as string;
        const response = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'GET',
            url: `${baseUrl}/files/${folderId}`,
          },
        );
        
        // Debug: Log the response to understand its structure
        console.log('OnlyOffice API Response Type:', typeof response);
        console.log('OnlyOffice API Response:', JSON.stringify(response, null, 2));
        
        // Also return debug info in the response for immediate visibility
        const debugInfo = {
          responseType: typeof response,
          responseLength: response ? (typeof response === 'string' ? response.length : Object.keys(response).length) : 'null',
          responsePreview: response ? (typeof response === 'string' ? response.substring(0, 200) : JSON.stringify(response).substring(0, 200)) : 'null'
        };
        
        // Extract the actual data from the API response
        // OnlyOffice API typically returns data in response.folders and response.files
        if (response && response.response) {
          const folders = response.response.folders || [];
          const files = response.response.files || [];
          return [...folders, ...files];
        }
        
        // Fallback: if response structure is different, try common alternatives
        if (response && response.data) {
          return response.data;
        }
        
        // If response has folders/files directly
        if (response && (response.folders || response.files)) {
          const folders = response.folders || [];
          const files = response.files || [];
          return [...folders, ...files];
        }
        
        // Handle case where response might be a string or unexpected format
        if (typeof response === 'string') {
          try {
            const parsedResponse = JSON.parse(response);
            if (parsedResponse && parsedResponse.response) {
              const folders = parsedResponse.response.folders || [];
              const files = parsedResponse.response.files || [];
              return [...folders, ...files];
            }
            if (parsedResponse && (parsedResponse.folders || parsedResponse.files)) {
              const folders = parsedResponse.folders || [];
              const files = parsedResponse.files || [];
              return [...folders, ...files];
            }
            return parsedResponse;
          } catch (e) {
            // If parsing fails, return the string as is
            return { error: 'Failed to parse response', rawResponse: response };
          }
        }
        
        // If response is an array, handle it properly
        if (Array.isArray(response)) {
          // Check if the first element is a JSON string
          if (response.length > 0 && typeof response[0] === 'string') {
            try {
              const parsedResponse = JSON.parse(response[0]);
              // Extract files and folders from the parsed response
              let extractedData = [];
              
              if (parsedResponse && parsedResponse.response) {
                const files = parsedResponse.response.files || [];
                const folders = parsedResponse.response.folders || [];
                extractedData = [...folders, ...files];
              } else if (parsedResponse && parsedResponse.files && parsedResponse.folders) {
                const files = parsedResponse.files || [];
                const folders = parsedResponse.folders || [];
                extractedData = [...folders, ...files];
              } else {
                extractedData = [parsedResponse];
              }
              
              return extractedData;
            } catch (e) {
              return [{ error: 'Failed to parse JSON response', rawResponse: response }];
            }
          } else {
            // If it's not a JSON string, return the array as is
            return response;
          }
        }
        
        // Parse the response correctly for OnlyOffice API
        let parsedResponse;
        
        // If response is a string, parse it as JSON
        if (typeof response === 'string') {
          try {
            parsedResponse = JSON.parse(response);
          } catch (e) {
            return [{ error: 'Failed to parse JSON response', rawResponse: response }];
          }
        } else {
          parsedResponse = response;
        }
        
        // Extract files and folders from the parsed response
        let extractedData = [];
        
        if (parsedResponse && parsedResponse.response) {
          const files = parsedResponse.response.files || [];
          const folders = parsedResponse.response.folders || [];
          extractedData = [...folders, ...files];
        } else if (parsedResponse && parsedResponse.files && parsedResponse.folders) {
          const files = parsedResponse.files || [];
          const folders = parsedResponse.folders || [];
          extractedData = [...folders, ...files];
        } else {
          extractedData = [parsedResponse];
        }
        
        return extractedData;

      case 'create':
        const parentFolderId = context.getNodeParameter('parentFolderId', itemIndex) as string;
        const title = context.getNodeParameter('title', itemIndex) as string;
        return await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'POST',
            url: `${baseUrl}/files/folder/${parentFolderId}`,
            body: { title },
          },
        );

      case 'rename':
        const itemId = context.getNodeParameter('itemId', itemIndex) as string;
        const newTitle = context.getNodeParameter('newTitle', itemIndex) as string;
        return await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'PUT',
            url: `${baseUrl}/files/folder/${itemId}`,
            body: { title: newTitle },
          },
        );

      case 'move':
      case 'copy':
        const moveItemId = context.getNodeParameter('itemId', itemIndex) as string;
        const destFolderId = context.getNodeParameter('destFolderId', itemIndex) as string;
        const conflictResolveType = context.getNodeParameter('conflictResolveType', itemIndex) as string;
        return await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'PUT',
            url: `${baseUrl}/files/fileops/${operation}`,
            body: {
              folderIds: [moveItemId],
              fileIds: [],
              destFolderId,
              conflictResolveType,
            },
          },
        );

      case 'delete':
        const deleteItemId = context.getNodeParameter('itemId', itemIndex) as string;
        const deleteImmediately = context.getNodeParameter('deleteImmediately', itemIndex) as boolean;
        return await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'DELETE',
            url: `${baseUrl}/files/folder/${deleteItemId}`,
            body: deleteImmediately ? { deleteAfter: true } : {},
          },
        );

      default:
        throw new NodeOperationError(context.getNode(), `Unknown folder operation: ${operation}`);
    }
  }

  private static async executeFileOperation(context: IExecuteFunctions, operation: string, itemIndex: number): Promise<any> {
    const credentials = await context.getCredentials('onlyOfficeApi');
    const baseUrl = `${credentials.baseUrl}/api/2.0`;

    switch (operation) {
      case 'list':
        const folderId = context.getNodeParameter('folderId', itemIndex) as string;
        const response = await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'GET',
            url: `${baseUrl}/files/${folderId}`,
          },
        );
        
        // Debug: Log the response to understand its structure
        console.log('OnlyOffice API Response Type:', typeof response);
        console.log('OnlyOffice API Response:', JSON.stringify(response, null, 2));
        
        // Also return debug info in the response for immediate visibility
        const debugInfo = {
          responseType: typeof response,
          responseLength: response ? (typeof response === 'string' ? response.length : Object.keys(response).length) : 'null',
          responsePreview: response ? (typeof response === 'string' ? response.substring(0, 200) : JSON.stringify(response).substring(0, 200)) : 'null'
        };
        
        // Extract the actual data from the API response
        // OnlyOffice API typically returns data in response.folders and response.files
        if (response && response.response) {
          const folders = response.response.folders || [];
          const files = response.response.files || [];
          return [...folders, ...files];
        }
        
        // Fallback: if response structure is different, try common alternatives
        if (response && response.data) {
          return response.data;
        }
        
        // If response has folders/files directly
        if (response && (response.folders || response.files)) {
          const folders = response.folders || [];
          const files = response.files || [];
          return [...folders, ...files];
        }
        
        // Handle case where response might be a string or unexpected format
        if (typeof response === 'string') {
          try {
            const parsedResponse = JSON.parse(response);
            if (parsedResponse && parsedResponse.response) {
              const folders = parsedResponse.response.folders || [];
              const files = parsedResponse.response.files || [];
              return [...folders, ...files];
            }
            if (parsedResponse && (parsedResponse.folders || parsedResponse.files)) {
              const folders = parsedResponse.folders || [];
              const files = parsedResponse.files || [];
              return [...folders, ...files];
            }
            return parsedResponse;
          } catch (e) {
            // If parsing fails, return the string as is
            return { error: 'Failed to parse response', rawResponse: response };
          }
        }
        
        // If response is an array, handle it properly
        if (Array.isArray(response)) {
          // Check if the first element is a JSON string
          if (response.length > 0 && typeof response[0] === 'string') {
            try {
              const parsedResponse = JSON.parse(response[0]);
              // Extract files and folders from the parsed response
              let extractedData = [];
              
              if (parsedResponse && parsedResponse.response) {
                const files = parsedResponse.response.files || [];
                const folders = parsedResponse.response.folders || [];
                extractedData = [...folders, ...files];
              } else if (parsedResponse && parsedResponse.files && parsedResponse.folders) {
                const files = parsedResponse.files || [];
                const folders = parsedResponse.folders || [];
                extractedData = [...folders, ...files];
              } else {
                extractedData = [parsedResponse];
              }
              
              return extractedData;
            } catch (e) {
              return [{ error: 'Failed to parse JSON response', rawResponse: response }];
            }
          } else {
            // If it's not a JSON string, return the array as is
            return response;
          }
        }
        
        // Parse the response correctly for OnlyOffice API
        let parsedResponse;
        
        // If response is a string, parse it as JSON
        if (typeof response === 'string') {
          try {
            parsedResponse = JSON.parse(response);
          } catch (e) {
            return [{ error: 'Failed to parse JSON response', rawResponse: response }];
          }
        } else {
          parsedResponse = response;
        }
        
        // Extract files and folders from the parsed response
        let extractedData = [];
        
        if (parsedResponse && parsedResponse.response) {
          const files = parsedResponse.response.files || [];
          const folders = parsedResponse.response.folders || [];
          extractedData = [...folders, ...files];
        } else if (parsedResponse && parsedResponse.files && parsedResponse.folders) {
          const files = parsedResponse.files || [];
          const folders = parsedResponse.folders || [];
          extractedData = [...folders, ...files];
        } else {
          extractedData = [parsedResponse];
        }
        
        return extractedData;

      case 'create':
        const parentFolderId = context.getNodeParameter('parentFolderId', itemIndex) as string;
        const title = context.getNodeParameter('title', itemIndex) as string;
        const fileType = context.getNodeParameter('fileType', itemIndex) as string;
        return await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'POST',
            url: `${baseUrl}/files/${parentFolderId}/file`,
            body: { 
              title: `${title}.${fileType}`,
              templateId: OnlyOffice.getTemplateId(fileType),
            },
          },
        );

      case 'rename':
        const itemId = context.getNodeParameter('itemId', itemIndex) as string;
        const newTitle = context.getNodeParameter('newTitle', itemIndex) as string;
        return await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'PUT',
            url: `${baseUrl}/files/file/${itemId}`,
            body: { title: newTitle },
          },
        );

      case 'move':
      case 'copy':
        const moveItemId = context.getNodeParameter('itemId', itemIndex) as string;
        const destFolderId = context.getNodeParameter('destFolderId', itemIndex) as string;
        const conflictResolveType = context.getNodeParameter('conflictResolveType', itemIndex) as string;
        return await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'PUT',
            url: `${baseUrl}/files/fileops/${operation}`,
            body: {
              folderIds: [],
              fileIds: [moveItemId],
              destFolderId,
              conflictResolveType,
            },
          },
        );

      case 'delete':
        const deleteItemId = context.getNodeParameter('itemId', itemIndex) as string;
        const deleteImmediately = context.getNodeParameter('deleteImmediately', itemIndex) as boolean;
        return await context.helpers.requestWithAuthentication.call(
          context,
          'onlyOfficeApi',
          {
            method: 'DELETE',
            url: `${baseUrl}/files/file/${deleteItemId}`,
            body: { deleteAfter: deleteImmediately },
          },
        );

      default:
        throw new NodeOperationError(context.getNode(), `Unknown file operation: ${operation}`);
    }
  }

  private static getTemplateId(fileType: string): number {
    // These are standard OnlyOffice template IDs
    switch (fileType) {
      case 'docx':
        return 1;
      case 'xlsx':
        return 2;
      case 'pptx':
        return 3;
      default:
        return 1;
    }
  }
}