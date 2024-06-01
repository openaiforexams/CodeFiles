import { sp } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders/list";
import "@pnp/sp/files/folder";
import { IList } from "@pnp/sp/lists";
import { IFieldInfo } from "@pnp/sp/fields";
import { IItem } from "@pnp/sp/items";

import { isEmpty } from "@microsoft/sp-lodash-subset";
import TechnicalSpecificationMain from "./TechnicalSpecificationMain";
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';

export const spListNames = {
    lst_initiatives: "Initiatives",
    lst_initiativeDocuments: "Initiative Documents",
    lst_documentTypes: "Document Types",
    lst_IDCO:"IDCO",
    lst_approversOtherFunctions:"Approvers - Other functions",
    lst_approversCX:"Approvers - CX",
    lst_configurations:"Configurations"
}
export interface initiatives {
    ID?: number;
    Title: string;
    DocumentType: number,
    FileID: number;
    Revision: string;
    DocumentNumber:string;
    Description:string;
}
export interface initiativeDocuments {
    ID?: number;
    InitiativeID: number;
}
export interface ITechnicalSpecificationMainProvider {
    getcurrentUserGroups(filter): Promise<any[]>;
    getcurrentUserDetails(): Promise<any>;
    getSPListItems(filter, orderby, select, expand, listName): Promise<any>;
    getSPLookupColumnValues(listName): Promise<IDropdownOption[]>;
    getSPChoiceColumnOptions(listName, fieldName): Promise<IDropdownOption[]>;
    getParameterByName(name);
  
}
export default class TechnicalSpecificationMainProvider implements ITechnicalSpecificationMainProvider {
    public getcurrentUserGroups(filter): Promise<any[]> {
        return sp.web.currentUser.groups();
    }
    public getcurrentUserDetails(): Promise<any[]> {
        return sp.web.currentUser.get();
    }
    public getUserDetails(ID): Promise<any[]> {
        return sp.web.getUserById(ID).get();
    }
    public getSPListItems(filter, orderby, select, expand, listName): Promise<any> {
        let filterItem = (filter === null ? '' : filter);
        let selectItem=(select===null ? '*':select);
        return sp.web.lists
            .getByTitle(listName).items
            .select(selectItem)
            .filter(filterItem)
            .orderBy(orderby)
            .expand(expand)
            .getAll()
            .then(
                (items: any[]): any[] => {
                    return items;
                }
            );
    }
    public getSPLookupColumnValues(listName): Promise<IDropdownOption[]> {
        return sp.web.lists
            .getByTitle(listName).items
            .orderBy("Title")
            .getAll()
            .then(
                (items: IDropdownOption[]): IDropdownOption[] => {
                    const ArrItems: IDropdownOption[] = [];
                    if (items !== null) {
                        items.forEach(
                            (item: any): void => {
                                ArrItems.push({
                                    key: item.ID,
                                    text: item.Title
                                });
                            });
                    }
                    return ArrItems;
                }
            );
    }
 
    public getSPChoiceColumnOptions(listName, fieldName): Promise<IDropdownOption[]> {
        let dropdownOption: IDropdownOption[] = [];
        return sp.web.lists
            .getByTitle(listName).fields.getByInternalNameOrTitle(fieldName).select("Choices").get()
            .then(
                (fieldInfo: IFieldInfo & { Choices: string[] }) => {
                    fieldInfo.Choices.map(item => { dropdownOption.push({ key: item, text: item }); });
                    return dropdownOption;
                });
    }
    public getParameterByName(name, url = window.location.href) {
        name = name.replace(/[\[\]]/g, '\\$&');
        let regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
            results = regex.exec(url);
        if (!results) return null;
        else if (!results[2]) return null;
        else return decodeURIComponent(results[2].replace(/\+/g, ' '));
    }

}
