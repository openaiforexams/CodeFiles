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
import TechnicalSpecificationInputForm from "./TechnicalSpecificationInputForm";
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';

export const spListNames = {
    lst_initiatives: "Initiatives",
    lst_initiativeDocuments: "Initiative Documents",
    lst_revisonNumbers: "Revision Numbers",
    lst_documentTypes: "Document Types",
    lst_departments: "Departments",
    lst_IDCO: "IDCO",
    lst_approvalHistory: "Approval History",
    lst_approvalHistorySME: "Approval History - SME",
    lst_approversOtherFunctions: "Approvers - Other functions",
    lst_initiativeApproversOtherFunctions:"Initiative Approvers - Other functions",
    lst_approvalHistoryOtherFunctions: "Approval History - Other Functions",
    lst_approvalHistoryCX: "Approval History - CX",
    lst_approversCX: "Approvers - CX",
    lst_configurations: "Configurations",    
    lst_initiativeDiscussions:"Initiative Discussions"
}

export interface initiatives {
    ID?: number;
    Title: string;
    DocumentType: number,
    FileIDId: number;
    Revision: string;
    DocumentNumber: string;
    Description: string;
}
export interface IBatchCreationData {
    item: any;
    result: any;
    error: any;
    id:number;
}
export interface initiativeDocuments {
    ID?: number;
    ProjectID: number;
}
export interface ITechnicalSpecificationInputFormProvider {
    getcurrentUserGroups(filter): Promise<any[]>;
    getcurrentUserDetails(): Promise<any>;
    getSPListItems(filter, orderby, sortAsc, select, expand, listName): Promise<any>;
    getSPLookupColumnValues(listName, isLookup): Promise<IDropdownOption[]>;
    getSPChoiceColumnOptions(listName, fieldName): Promise<IDropdownOption[]>;
    getInitiativeDetails(ID): Promise<any>;
    getParameterByName(name);
    updateItem(listname, item, ID): Promise<any>;
    createItem(item): Promise<any>;
    createSPListItem(item,listname): Promise<any>;
    getDocument(ID): Promise<any>;
    updateDocument(fileupload, ID, filePath): Promise<any>;
    deleteAttachment(ID): Promise<any>;
    getUserProfileUrl(loginName): Promise<any>;
    processBatch(list, itemEntityType: string, data: IBatchCreationData[], index,type): Promise<IBatchCreationData[]>;
}
export default class TechnicalSpecificationInputFormProvider implements ITechnicalSpecificationInputFormProvider {
    public getcurrentUserGroups(filter): Promise<any[]> {
        return sp.web.currentUser.groups();
    }
    public getcurrentUserDetails(): Promise<any[]> {
        return sp.web.currentUser.get();
    }
    public getUserDetails(ID): Promise<any[]> {
        return sp.web.getUserById(ID).get();
    }
    public async getUserProfileUrl(loginName: string): Promise<any> {
        let userPictureUrl = await sp.profiles.getUserProfilePropertyFor(loginName, 'PictureURL');
        //console.log(userPictureUrl);
        return userPictureUrl.toString();
    }
    public async getSPListItems(filter, orderby, sortAsc, select, expand, listName): Promise<any> {
        let filterItem = (filter === null ? '' : filter);
        let selectItem = (select === null ? '*' : select);
        if(sortAsc){
            return sp.web.lists
            .getByTitle(listName).items
            .select(selectItem)
            .filter(filterItem)
            .orderBy(orderby,sortAsc)
            .expand(expand)
            .getAll()
            .then(
                (items: any[]): any[] => {
                    return items;
                }
            );
        }
        else{
            return sp.web.lists
            .getByTitle(listName).items
            .select(selectItem)
            .filter(filterItem)
            .orderBy(orderby,sortAsc)
            .expand(expand)
            .get()
            .then(
                (items: any[]): any[] => {
                    return items;
                }
            );
        }
        
    }
    public getSPLookupColumnValues(listName, isLookup): Promise<IDropdownOption[]> {
        return sp.web.lists
            .getByTitle(listName).items
            .orderBy("Title")
            .get()
            .then(
                (items: IDropdownOption[]): IDropdownOption[] => {
                    const ArrItems: IDropdownOption[] = [];
                    if (items !== null) {
                        items.forEach(
                            (item: any): void => {
                                if (isLookup) {
                                    ArrItems.push({
                                        key: item.ID,
                                        text: item.Title
                                    });
                                }
                                else {
                                    ArrItems.push({
                                        key: item.Title,
                                        text: item.Title
                                    });
                                }

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
    public getInitiativeDetails(ID): Promise<any> {
        return sp.web.lists
            .getByTitle(spListNames.lst_initiatives).items.filter("ID eq " + ID)
            .select('*', 'DocumentType/Title', 'DocumentType/ID',
                'Author/EMail', 'Author/Title', 'Author/Name','Requester/EMail','Requester/Title','Requester/Name', 'Editor/Title', 'Editor/EMail', 'Editor/Name', 'LastUpdatedBy/Title', 'LastUpdatedBy/EMail', 'LastUpdatedBy/Name', 'IDCOApprover/Name', 'IDCOApprover/EMail', 'SME/EMail', 'SME/Title', 'SME/Id'
            ).expand('DocumentType', 'Author', 'Editor', 'IDCOApprover', 'SME','LastUpdatedBy','Requester').getAll().then(
                (item) => {
                    //console.log(item);
                    return item;
                });
    }
    public createItem(item): Promise<any> {
        return sp.web.lists.
            getByTitle(spListNames.lst_initiatives).
            items.add(item).
            then((items): any => {
                return items;
            });
    }
    public createSPListItem(item,listname): Promise<any> {
        return sp.web.lists.
            getByTitle(listname).
            items.add(item).
            then((items): any => {
                return items;
            });
    }
    public updateItem(listname, item, ID): Promise<any> {
        return sp.web.lists.
            getByTitle(listname).
            items.getById(ID).update(item).
            then((items): any => {
                return items;
            });
    }
    public async processBatch (list: IList, itemEntityType: string, data: IBatchCreationData[], index = 0,type): Promise<IBatchCreationData[]> {        
        const batchSize = 100;
        if (data.length > index + 1) {
          let batch = sp.web.createBatch();
          for (let len = index + batchSize; index < len && index < data.length; index += 1) {
    
            let dataItem = data[index];        
            let success = (function (res) {
              this.result = res;
            }).bind(dataItem);
    
            let error = (function (err) {
              this.error = err;
              console.log(err);
            }).bind(dataItem);
    
            if(type==="Create")
            {
            list.items.inBatch(batch).add(dataItem.item, itemEntityType)
              .then(success)
              .catch(error);
            }
            if(type==="Update")
            {
              list.items.getById(dataItem.id).inBatch(batch).update(dataItem.item,'*',itemEntityType)
              .then(success)
              .catch(error);
            }
          }          
          console.log(`Processing (${index} of ${data.length})...`);
          let result = await batch.execute();
          return await this.processBatch(list, itemEntityType, data,index,type);
        } 
        else {
          //return new Promise(function (resolve, reject) {
            //resolve => resolve(data)
            true;
         // });
        }
        
      }
    public getParameterByName(name, url = window.location.href) {
        name = name.replace(/[\[\]]/g, '\\$&');
        let regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
            results = regex.exec(url);
        if (!results) return null;
        else if (!results[2]) return null;
        else return decodeURIComponent(results[2].replace(/\+/g, ' '));
    }
    public async getDocument(ID): Promise<any> {
        return sp.web.lists.getByTitle(spListNames.lst_initiativeDocuments).items.select("EncodedAbsUrl,FileLeafRef", "Created", "Modified", "File", "ID", "Author/Name", "Author/EMail", "Author/Title", "Editor/Name", "Editor/EMail", "Editor/Title").expand("File", "Author", "Editor").filter("InitiativeIDId eq " + ID).get().then(item => {
            //console.log(item);
            return item;
        });
    }
    public async updateDocument(fileupload, ID, filePath): Promise<any> {
        let FileName = fileupload.name;
        FileName = FileName.replace(/ /g, "_");
        const fileNamePath = filePath + encodeURI(FileName);
        //console.log(fileNamePath);
        const file = await sp.web.lists.getByTitle(spListNames.lst_initiativeDocuments).rootFolder.files.addUsingPath(fileNamePath, fileupload, { Overwrite: true });
        const item = await file.file.getItem();
        return await item.update({
            InitiativeIDId: ID
        });
    }
    public deleteAttachment(ID): Promise<any> {
        return sp.web.lists.getByTitle(spListNames.lst_initiativeDocuments).items.getById(ID).delete().then((item): any => {
            alert("Deleted successfully");
        }).catch((reason) => {
            alert("Error" + reason);
        });
    }
}
