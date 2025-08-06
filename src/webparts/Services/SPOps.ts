/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-async-promise-executor */
/* eslint-disable prefer-const */
import { sp } from '@pnp/sp/presets/all';
import { IDropdownOption } from 'office-ui-fabric-react';

export class SPOperation {

    public listArea1Topics: IDropdownOption[] = [];
    public listArea2Topics: IDropdownOption[] = [];
    public listTopics: IDropdownOption[] = [];

    public a: string;
    public b: string;
    public sp: any;

    public getAreaDropDownOptions(areaMap: Map<number, string>): Promise<IDropdownOption[]> {

        // let sp = spfi();

        let listAreaTitles: IDropdownOption[] = []
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {

            await sp.web.lists.getByTitle("Areas of Law").items.select("ID", "Title").get().then((results: any) => {

                results.map((result: any) => {

                    listAreaTitles.push({ key: result.ID, text: result.Title })
                    areaMap.set(result.ID, result.Title);
                })
                resolve(listAreaTitles);
            }, (error: any) => { reject("error occured") })

        })
    }

    public GetDropdownOptions(listLookup: Map<number, string>, listName: string): Promise<IDropdownOption[]> {

        let listTitles: IDropdownOption[] = []
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {

            // let sp = spfi();

            await sp.web.lists.getByTitle(listName).items.select("ID", "Title").top(5000).get().then((results: any) => {

                if (listName === "Years") {
                    results.sort((a: { Title: string; }, b: { Title: string; }) => {
                        if (a.Title === b.Title) {
                            return a.Title < b.Title ? 1 : -1
                        } else {
                            return a.Title < b.Title ? 1 : -1
                        }
                    })
                }
                else if (listName.indexOf("Libraries") === 0) {
                    results.sort((a: { Title: string; }, b: { Title: string; }) => {
                        if (a.Title === b.Title) {
                            return a.Title > b.Title ? 1 : -1
                        } else {
                            return a.Title > b.Title ? 1 : -1
                        }
                    })
                }

                results.map((result: any) => {
                    listTitles.push({ key: result.ID, text: result.Title })
                    listLookup.set(result.ID, result.Title);
                })
                resolve(listTitles);
            }, (error: any) => { reject("error occured") })

        })
    }

    public GetLibDropdownOptions(listLookup: Map<number, string>, listIdLookup: Map<string, number>, libLookupByFrag: Map<string, string>, fragLookupByLib: Map<string, string>, listName: string, selectedLibrary: string): Promise<IDropdownOption[]> {

        let listTitles: IDropdownOption[] = []
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {

            // let sp = spfi();

            await sp.web.lists.getByTitle(listName).items.select("ID", "Title", "LSB_UrlFragment").top(5000).get().then((results: any) => {
                results.sort((a: { Title: string; }, b: { Title: string; }) => {
                    if (a.Title === b.Title) {
                        return a.Title > b.Title ? 1 : -1
                    } else {
                        return a.Title > b.Title ? 1 : -1
                    }
                })

                let i = 0;

                listTitles.push({ key: 0, text: "All" })
                selectedLibrary = 'All';
                results.map((result: any) => {
                    i++;
                    listTitles.push({ key: i, text: result.Title })
                    listLookup.set(i, result.Title);
                    listIdLookup.set(result.Title, i);
                    libLookupByFrag.set(result.LSB_UrlFragment, result.Title);
                    fragLookupByLib.set(result.Title, result.LSB_UrlFragment);
                })
                resolve(listTitles);
            }, (error: any) => { reject("error occured") })

        })
    }

    public getArea1Topics(topicLookup: Map<number, string>, topicIdLookup: Map<string, number>, s: string): Promise<IDropdownOption[]> {

        // let sp = spfi();

        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            let query = `<View><Query><Where><Eq><FieldRef Name="AreaOfLaw"/><Value Type="Lookup">` + s + `</Value></Eq></Where></Query></View>`;
            await sp.web.lists.getByTitle("Topics of Law").getItemsByCAMLQuery({
                ViewXml: query,
            }).then((results: any) => {
                this.listArea1Topics = [];
                results.map((result: any) => {
                    this.listArea1Topics.push({ key: result.ID, text: result.Title })
                    topicLookup.set(result.ID, result.Title);
                    topicIdLookup.set(result.Title, result.ID);
                })
                resolve(this.listArea1Topics);
            }, (error: any) => { reject("error occured") })

        })
    }

    public getArea2Topics(topicLookup: Map<number, string>, topicIdLookup: Map<string, number>, s: string): Promise<IDropdownOption[]> {

        return new Promise<IDropdownOption[]>(async (resolve, reject) => {

            // let sp = spfi();

            let query = `<View><Query><Where><Eq><FieldRef Name="AreaOfLaw"/><Value Type="Lookup">` + s + `</Value></Eq></Where></Query></View>`;
            await sp.web.lists.getByTitle("Topics of Law").getItemsByCAMLQuery({
                ViewXml: query,
            }).then((results: any) => {
                this.listArea2Topics = [];
                results.map((result: any) => {
                    this.listArea2Topics.push({ key: result.ID, text: result.Title })
                    topicLookup.set(result.ID, result.Title);
                    topicIdLookup.set(result.Title, result.ID);
                })
                resolve(this.listArea2Topics);
            }, (error: any) => { reject("error occured") })

        })
    }

    public getAllTopics(topicLookup: Map<number, string>, topicIdLookup: Map<string, number>): Promise<IDropdownOption[]> {

        // let sp = spfi();

        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            let query = `<View><Query><Where><IsNotNull><FieldRef Name="Title"/></IsNotNull></Where></Query></View>`;
            sp.web.lists.getByTitle("Topics of Law").getItemsByCAMLQuery({
                ViewXml: query,
            }).then((results: any) => {
                results.map((result: any) => {
                    topicLookup.set(result.ID, result.Title);
                    topicIdLookup.set(result.Title, result.ID);
                })
                resolve(this.listTopics);
            }, (error: any) => { reject("error occured") })

        })
    }


}
