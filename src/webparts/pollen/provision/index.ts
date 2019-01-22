import { each } from '@microsoft/sp-lodash-subset';
import { sp, Field, ItemAddResult, FieldAddResult, FieldCreationProperties, ListAddResult, FieldUpdateResult, ContentTypeAddResult, ViewAddResult } from '@pnp/sp';

import FieldSchema from './fields';
import ContentTypeSchema from './contentTypes';
import { View } from '@pnp/sp/src/views';

export default class Provisioner {
    public static ProvisionLists() {
        console.log('provisioning lists');
        // Provision Poll List
        return Provisioner.PollList().then(() => {
            console.log('poll list created');            
            // Provision Response List
            return Provisioner.ResponseList().then( () => {
                console.log('response list created');
            });
        });
    }

    private static PollList(): Promise<any> {
        return Provisioner.ProvisionFields(FieldSchema.PollFields).then(() => {            
            return Provisioner.ProvisionContentType(ContentTypeSchema.PollContentType).then((ct: ContentTypeAddResult) => {
                return Provisioner.ProvisionList('Pollen Polls', ct.data.Id.StringValue, FieldSchema.PollFields);
            });
        });
    }

    private static ResponseList(): Promise<any> {
        return Provisioner.ProvisionFields(FieldSchema.ResponseFields).then(() => {
            return Provisioner.ProvisionContentType(ContentTypeSchema.ResponseContentType).then((ct: ContentTypeAddResult) => {
               return Provisioner.ProvisionList('Pollen Responses', ct.data.Id.StringValue, FieldSchema.ResponseFields);
            });
        });
    }

    private static ProvisionFields(fieldSchemas): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            var fields = fieldSchemas.length;
            each(fieldSchemas, (fieldSchema) => {
                this.ProvisionField(fieldSchema).then((field) => {
                    fields--;
                    if(fields === 0){
                        resolve();
                    }
                }, (error) => {
                    fields--;
                    reject(error);
                });
            });
        });
    }

    private static ProvisionField(fieldSchema): Promise<FieldUpdateResult>{
        if(fieldSchema.fieldType === "SP.FieldLookup"){
            return sp.site.rootWeb.lists.getByTitle(fieldSchema.listName).get().then((list) => {
                console.log(list);
                return sp.site.rootWeb.fields.addLookup(fieldSchema.title, list.Id, "Id", fieldSchema.properties).then((field) => {
                    return field.field.update(fieldSchema.updateProperties);
                });
            });        
        } else {
            return sp.site.rootWeb.fields.add(fieldSchema.title, fieldSchema.fieldType, fieldSchema.properties).then((field) => {
                return field.field.update(fieldSchema.updateProperties);
            });
        }
    }

    private static ProvisionContentType(ctSchema): Promise<ContentTypeAddResult> {
        return sp.site.getWebUrlFromPageUrl(window.location.href).then(site => {
            return sp.site.rootWeb.contentTypes.add(ctSchema.Id, ctSchema.Name, ctSchema.Description, ctSchema.Group, null).then((ct) => {
                return new Promise<any>((resolve, reject) => {
                    var context = new SP.ClientContext(site); //.get_current();
                    var web = context.get_site().get_rootWeb();
                    var contentType = web.get_contentTypes().getById(ct.data.Id.StringValue);
                    var fieldLinks = contentType.get_fieldLinks();
    
                    each(ctSchema.Fields, (fieldName) => {
                        var field = web.get_fields().getByInternalNameOrTitle(fieldName);
                        var fieldLinkCreationInfo = new SP.FieldLinkCreationInformation();
                        fieldLinkCreationInfo.set_field(field);
                        fieldLinks.add(fieldLinkCreationInfo);
                        contentType.update(true);
                    });
    
                    context.executeQueryAsync((sender, args) => {
                        resolve(ct);
                    }, (sender, args) => {
                        reject(sender);
                    });
                });
            });
        });
    }

    private static SetDefaultContentType(list, contentTypeId): Promise<any>{
        return sp.site.getWebUrlFromPageUrl(window.location.href).then(site => {
            return new Promise<any>((resolve, reject) => {
                sp.web.lists.getByTitle(list).contentTypes.get();
                var context = new SP.ClientContext(site); //.get_current();
                var web = context.get_site().get_rootWeb();
                var contentTypes = web.get_lists().getByTitle(list).get_contentTypes();
                var rootFolder = web.get_lists().getByTitle(list).get_rootFolder();
                context.load(contentTypes);
                context.load(rootFolder);

                context.executeQueryAsync(
                    () => {
                        contentTypes = context.get_site().get_rootWeb().get_lists().getByTitle(list).get_contentTypes();
                        var ctOrder = new Array();
                        var ctEnum = contentTypes.getEnumerator();
                        
                        while (ctEnum.moveNext()) {
                            if(ctEnum.get_current().get_id() == contentTypeId)
                            {
                                ctOrder.push(ctEnum.get_current().get_id());
                            }
                        }

                        rootFolder.set_uniqueContentTypeOrder(ctOrder);
                        rootFolder.update();

                        context.executeQueryAsync(() => { resolve(); }, () => { reject(); } );
                    }, 
                    () => {
                        reject();
                    }
                );
            });
        });
    }

    private static ProvisionList(listName, contentTypeId, fieldSchema): Promise<any> {
        return sp.site.rootWeb.lists.add(listName, '', 100, true).then((l: ListAddResult) => {
            return l.list.contentTypes.addAvailableContentType(contentTypeId).then((ct: ContentTypeAddResult) => {
                // Set default content type on list
                return this.SetDefaultContentType(listName, ct.data.Id.StringValue).then(() => {
                    // Add fields to default view                
                    return fieldSchema.reduce((promiseChain, field) => {
                        return promiseChain.then(() => {
                            return l.list.defaultView.fields.add(field.title);
                        });
                    }, Promise.resolve([]));
                });
            });
        });
    }
}