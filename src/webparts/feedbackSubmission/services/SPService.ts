import { SPHttpClient } from '@microsoft/sp-http';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

export interface ISPList {
    Title: string;
    Id: string;
}

export interface ISPField {
    InternalName: string;
    Title: string;
    Hidden: boolean;
    ReadOnlyField: boolean;
}

export interface ISPItem {
    ID: number;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    [key: string]: any;
}

export class SPService {
    private readonly _spHttpClient: SPHttpClient;

    constructor(spHttpClient: SPHttpClient) {
        this._spHttpClient = spHttpClient;
    }

    public async getLists(currentWebUrl: string): Promise<IPropertyPaneDropdownOption[]> {
        try {
            const url = `${currentWebUrl}/_api/web/lists?$select=Id,Title&$filter=Hidden eq false`;
            const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "odata-version": ""
                }
            });
            const data = await response.json();
            if (data.d?.results) {
                return data.d.results.map((list: ISPList) => {
                    return { key: list.Title, text: list.Title };
                });
            }
            return [];
        } catch (e: unknown) {
            console.error('Error fetching lists', e);
            return [];
        }
    }

    public async getFields(currentWebUrl: string, listTitle: string): Promise<IPropertyPaneDropdownOption[]> {
        if (!listTitle) return [];
        try {
            const url = `${currentWebUrl}/_api/web/lists/getbytitle('${listTitle}')/fields?$select=InternalName,Title,Hidden,ReadOnlyField&$filter=Hidden eq false and ReadOnlyField eq false`;
            const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "odata-version": ""
                }
            });
            const data = await response.json();
            if (data.d?.results) {
                return data.d.results.map((field: ISPField) => {
                    return { key: field.InternalName, text: `${field.Title} [${field.InternalName}]` };
                });
            }
            return [];
        } catch (e: unknown) {
            console.error('Error fetching fields', e);
            return [];
        }
    }

    public async getQuestions(currentWebUrl: string, listTitle: string): Promise<ISPItem | undefined> {
        if (!listTitle) return undefined;
        try {
            const url = `${currentWebUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$top=1`;
            const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "odata-version": ""
                }
            });
            if (response.ok) {
                const data = await response.json();
                if (data.d?.results?.length > 0) {
                    return data.d.results[0];
                }
            }
            return undefined;
        } catch (error: unknown) {
            console.error("Error loading questions:", error);
            return undefined;
        }
    }

    public async getItem(currentWebUrl: string, listTitle: string, itemId: string, selectFields: string[] = ["UserAck"]): Promise<ISPItem | undefined> {
        try {
            const selectQuery = `ID,${selectFields.join(',')}`;
            const url = `${currentWebUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})?$select=${selectQuery}`;
            const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "odata-version": ""
                }
            });
            if (!response.ok) return undefined;
            const data = await response.json();
            return data.d;
        } catch (error: unknown) {
            console.error("Error getting item:", error);
            return undefined;
        }
    }

    public async submitFeedback(currentWebUrl: string, listTitle: string, itemId: string, values: Record<string, unknown>, itemType: string = "SP.ListItem"): Promise<boolean> {
        try {
            const url = `${currentWebUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`;

            const body: Record<string, unknown> = {
                "__metadata": { "type": itemType },
                ...values
            };

            const headers = {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",
                "odata-version": ""
            };

            const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
                headers: headers,
                body: JSON.stringify(body)
            });

            return response.ok;
        } catch (error: unknown) {
            console.error("Error submitting feedback:", error);
            return false;
        }
    }
}
