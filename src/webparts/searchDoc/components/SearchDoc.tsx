/* eslint-disable @rushstack/security/no-unsafe-regexp */
/* eslint-disable react/no-direct-mutation-state */
/* eslint-disable react/jsx-no-target-blank */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable no-var */
import * as React from 'react';
import styles from './SearchDoc.module.scss';

import { ISearchDocProps } from './ISearchDocProps';
import { ISearchDocState } from './ISearchDocState';
import { SPOperation } from "../../Services/SPOps";
import { ComboBox, PrimaryButton } from "office-ui-fabric-react";
import { sp } from '@pnp/sp';
import { Pagination } from "@pnp/spfx-controls-react/lib/pagination";

 

export default class SearchDoc extends React.Component<ISearchDocProps, ISearchDocState,
  {}>
{
  private _spService: SPOperation;
  public selectedListTitleArea1: string;
  public selectedListTitleTopic1: string;
  public selectedListTitleArea2: string;
  public selectedListTitleTopic2: string;
  public selectedListTitleLegislation: string;
  public selectedListTitleAuthor: string;
  public selectedListTitleYear: string;
  public selectedListTitleDocType: string;
  public selectedConflict: string;
  public selectedLibrary: string = 'All';
  public selectedSort: string = 'File Name';
  public keyPhraseValue: string;
  public queries: string[] = [];
  public defaultOption: string = '';
  public libraries: string[] = [];
  public libsToQuery: string[] = [];
  public itemList: any = [];
  public viewfields: string = '<ViewFields><FieldRef Name="Name"/><FieldRef Name="LSB_AreaOfLaw"/><FieldRef Name="LSB_AreaOfLaw2"/><FieldRef Name="LSB_TopicsOfLaw"/><FieldRef Name="LSB_TopicsOfLaw2"/><FieldRef Name="LSB_Legislation"/><FieldRef Name="LSB_Conflict"/><FieldRef Name="Document_x0020_Type"/><FieldRef Name="LSB_AuthorNames"/><FieldRef Name="Year"/><FieldRef Name="LSB_LegalFileNumber"/><FieldRef Name="RoutingRuleDescription"/></ViewFields>';
  public finalQuery: string = '';
  public areaLookup: Map<number, string> = new Map<number, string>();
  public areaIdLookup: Map<string, number> = new Map<string, number>();
  public topicLookup: Map<number, string> = new Map<number, string>();
  public topicIdLookup: Map<string, number> = new Map<string, number>();
  public legislationLookup: Map<number, string> = new Map<number, string>();
  public legislationIdLookup: Map<string, number> = new Map<string, number>();
  public docTypeLookup: Map<number, string> = new Map<number, string>();
  public authorLookup: Map<number, string> = new Map<number, string>();
  public yearLookup: Map<number, string> = new Map<number, string>();
  public lookupLibs: Map<number, string> = new Map<number, string>();
  public lookupLibByUrlFragment: Map<string, string> = new Map<string, string>();
  public lookupUrlFragmentByLibName: Map<string, string> = new Map<string, string>();

  public listRootUrl: string = "https://bcgov.sharepoint.com/";
  public sitePath: string = "/sites/AG-LSBKB/";
  public sortFields: any = [];
  public conflict: any = [];
  public baseState: any;
  public pageSize: number = 40;
  public pagesTotal: number = 5;
  public pagesMax: number = 5;
  public libsList = "Libraries";
  public statusFilter = "Approved";
  public finalCount: number = 0;
  public splitChars: string[] = [" ", ";"];

  constructor(props: ISearchDocProps) {
    super(props);
    this._initFormState();
    this.sortFields = [
      { key: "Name", text: "File Name", ascending: true },
      { key: "LSB_AreaOfLaw", text: "Area of Law (1)", ascending: true },
      { key: "Document_x0020_Type", text: "Doc Type", ascending: true },
      { key: "LSB_AuthorNames", text: "Author", ascending: true },
      { key: "Year", text: "Year", ascending: false },
      { key: "Library", text: "Library", ascending: true }
    ];
    this.conflict = [
      { key: "No", text: "No" },
      { key: "Yes", text: "Yes" }
    ];
  }

  private _initFormState() {
    this._spService = new SPOperation();
    this.state = {
      listTitlesArea1: [], listAreaTopics1: [], listTitlesArea2: [], listAreaTopics2: [], listTitlesLegislation: [], listTitlesDocType: [], listTitlesAuthor: [], listTitlesYear: [], listTitlesLibrary: [], keyPhrase: '',
      items: [
        {
          LSB_AreaOfLaw1: "",
          LSB_AreaOfLaw2: "",
          LSB_AuthorNames: "",
          LSB_Conflict: "",
          Document_x0020_Type: "",
          LSB_LegalFileNumber: "",
          LSB_Legislation: "",
          LSB_LibName: "",
          LSB_TopicsOfLaw1: "",
          LSB_TopicsOfLaw2: "",
          Year: "",
          Name: "",
          RoutingRuleDescription: "",
          URL: "",
          LSB_LibURL: ""
        }
      ],
      searchEnabled: true,
      firstTopicEnabled: false,
      secondAreaEnabled: false,
      secondTopicEnabled: false,
      selectedSort: this.selectedSort,
      selectedLibrary: this.selectedLibrary,
      paginatedItems: [
        {
          LSB_AreaOfLaw1: "",
          LSB_AreaOfLaw2: "",
          LSB_AuthorNames: "",
          LSB_Conflict: "",
          Document_x0020_Type: "",
          LSB_LegalFileNumber: "",
          LSB_Legislation: "",
          LSB_LibName: "",
          LSB_TopicsOfLaw1: "",
          LSB_TopicsOfLaw2: "",
          Year: "",
          Name: "",
          RoutingRuleDescription: "",
          URL: "",
          LSB_LibURL: ""
        }
      ],
    };
  }


  // Method to find a value in a string or string array
  public containsVal(inputObj: any, searchValue: string): boolean {
    let isContained = false;
    let arr: string[] = [];
    if (!Array.isArray(inputObj)) arr.push(inputObj);
    arr.forEach(function (value) {
      if (value.indexOf(searchValue) > -1) isContained = true;
      if (isContained) return;
    });
    return isContained;
  }

  public componentDidMount(): void {
    //Setup functions to load the dropdowns
    document.getElementById('dv_Table').style.display = 'none';
    document.getElementById('dv_pagination').style.display = 'none';
    this._spService.getAllTopics(this.topicLookup, this.topicIdLookup).then((result) => {
      this.setState({ listAreaTopics1: result });
      this.setState({ listAreaTopics2: result });
    });
    this._spService.getAreaDropDownOptions(this.areaLookup).then((result) => {
      this.setState({ listTitlesArea1: result });
      this.setState({ listTitlesArea2: result });
    });
    this._spService.GetDropdownOptions(this.legislationLookup, "Legislation").then((result) => {
      //console.log("Legislation options loaded:", result);
      this.setState({ listTitlesLegislation: result });
    });
    this._spService.GetDropdownOptions(this.docTypeLookup, "Document Types").then((result) => {
      this.setState({ listTitlesDocType: result });
    });
    this._spService.GetDropdownOptions(this.authorLookup, "Authors").then((result) => {
      this.setState({ listTitlesAuthor: result });
    });
    this._spService.GetDropdownOptions(this.yearLookup, "Years").then((result) => {
      this.setState({ listTitlesYear: result });
    });
    this._spService.GetLibDropdownOptions(this.lookupLibs, this.topicIdLookup, this.lookupLibByUrlFragment, this.lookupUrlFragmentByLibName, this.libsList, this.selectedLibrary).then((result) => {
      this.setState({ listTitlesLibrary: result });
    });

    let drLib = document.getElementById('drLib-input');
    if (drLib !== undefined && drLib !== null) {
      drLib.setAttribute('value', 'All');
      this.setState({ selectedLibrary: 'All' });
    }

    this.baseState = this.state;
  }

  private _onKeyPhraseChange = (e: React.ChangeEvent) => {
    const value = (e.target as HTMLInputElement).value;
    this.keyPhraseValue = value;
    this.setState({ selectedSort: this.selectedSort })
  }

  //Functions to get the selected value and implement the cascading dropdown functionality
  public getSelectedListTitleArea1 = (ev: any, data: any) => {
    this.selectedListTitleArea1 = data.text;
    this._spService.getArea1Topics(this.topicLookup, this.topicIdLookup, data.text).then((resultTopic: any) => {
      this.setState({ listAreaTopics1: resultTopic });
    });
    this.setState({ searchEnabled: true, firstTopicEnabled: true });
  }
  public getselectedListTitleTopic1 = (ev: any, data: any) => {
    this.selectedListTitleTopic1 = data.text;
    this.setState({ searchEnabled: true, secondAreaEnabled: true });
  }
  public getSelectedListTitleArea2 = (ev: any, data: any) => {
    this.selectedListTitleArea2 = data.text;
    this._spService.getArea2Topics(this.topicLookup, this.topicIdLookup, data.text).then((resultTopic: any) => {
      this.setState({ listAreaTopics2: resultTopic });
    });
    this.setState({ searchEnabled: true, secondTopicEnabled: true });
  }
  public getselectedListTitleTopic2 = (ev: any, data: any) => {
    this.selectedListTitleTopic2 = data.text;
  }

  public getselectedListTitleLegislation = (ev: any, data: any) => {
    this.selectedListTitleLegislation = data.text;
    this.setState({ searchEnabled: true });
  }
  public getSelectedListTitleAuthor = (ev: any, data: any) => {
    this.selectedListTitleAuthor = data.text;
    this.setState({ searchEnabled: true });
  }
  public getSelectedListTitleDocType = (ev: any, data: any) => {
    this.selectedListTitleDocType = data.text;
    this.setState({ searchEnabled: true });
  }
  public getSelectedListTitleYear = (ev: any, data: any) => {
    this.selectedListTitleYear = data.text;
    this.setState({ searchEnabled: true });
  }
  public getSelectedConflict = (ev: any, data: any) => {
    this.selectedConflict = data.text;
    this.setState({ searchEnabled: true });
  }
  public getSelectedLibrary = (ev: any, data: any) => {
    this.selectedLibrary = data.text;
    this.setState({ selectedLibrary: this.selectedLibrary });
  }
  public getSelectedSort = (ev: any, data: any) => {
    this.selectedSort = data.text;
    this.setState({ selectedSort: this.selectedSort })
  }
  public getkeyPhraseValue = (ev: any, data: any) => {
    this.keyPhraseValue = data.text;
    this.setState({ searchEnabled: true });
  }

  public getQueryOneFilter = (queries: string[]): string => {
    var con = queries[0];
    var query = '<View><Query><Where>' + con + '</Where></Query>' + this.viewfields + '</View>';
    return query;
  }
  public getQueryTwoFilters = (queries: string[]): string => {
    var cond1 = queries[0];
    var cond2 = queries[1];
    var query = '<View><Query><Where><And>' + cond1 + cond2 + '</And></Where></Query>' + this.viewfields + '</View>';
    return query;
  }

  public getQueryThreeFilters = (queries: string[]): string => {
    var cond1 = queries[0];
    var cond2 = queries[1];
    var cond3 = queries[2];
    var query = '<View><Query><Where><And>' + cond1 + '<And>' + cond2 + cond3 + '</And></And></Where></Query>' + this.viewfields + '</View>';
    return query;
  }
  public getQueryFourFilters = (queries: string[]): string => {
    var cond1 = queries[0];
    var cond2 = queries[1];
    var cond3 = queries[2];
    var cond4 = queries[3];
    var query = '<View><Query><Where><And>' + cond1 + '<And>' + cond2 + '<And>' + cond3 + cond4 + '</And></And></And></Where></Query>' + this.viewfields + '</View>';
    return query;
  }
  public getQueryFiveFilters = (queries: string[]): string => {
    var cond1 = queries[0];
    var cond2 = queries[1];
    var cond3 = queries[2];
    var cond4 = queries[3];
    var cond5 = queries[4];
    var query = '<View><Query><Where><And>' + cond1 + '<And>' + cond2 + '<And>' + cond3 + '<And>' + cond4 + cond5 + '</And></And></And></And></Where></Query>' + this.viewfields + '</View>';
    return query;
  }
  public getQuerySixFilters = (queries: string[]): string => {
    var cond1 = queries[0];
    var cond2 = queries[1];
    var cond3 = queries[2];
    var cond4 = queries[3];
    var cond5 = queries[4];
    var cond6 = queries[5];
    var query = '<View><Query><Where><And>' + cond1 + '<And>' + cond2 + '<And>' + cond3 + '<And>' + cond4 + '<And>' + cond5 + cond6 + '</And></And></And></And></And></Where></Query>' + this.viewfields + '</View>';
    return query;
  }
  public getQuerySevenFilters = (queries: string[]): string => {
    var cond1 = queries[0];
    var cond2 = queries[1];
    var cond3 = queries[2];
    var cond4 = queries[3];
    var cond5 = queries[4];
    var cond6 = queries[5];
    var cond7 = queries[6];
    var query = '<View><Query><Where><And>' + cond1 + '<And>' + cond2 + '<And>' + cond3 + '<And>' + cond4 + '<And>' + cond5 + '<And>' + cond6 + cond7 + '</And></And></And></And></And></And></Where></Query>' + this.viewfields + '</View>';
    return query;
  }
  public getQueryEightFilters = (queries: string[]): string => {
    var cond1 = queries[0];
    var cond2 = queries[1];
    var cond3 = queries[2];
    var cond4 = queries[3];
    var cond5 = queries[4];
    var cond6 = queries[5];
    var cond7 = queries[6];
    var cond8 = queries[7];
    var query = '<View><Query><Where><And>' + cond1 + '<And>' + cond2 + '<And>' + cond3 + '<And>' + cond4 + '<And>' + cond5 + '<And>' + cond6 + '<And>' + cond7 + cond8 + '</And></And></And></And></And></And></And></Where></Query>' + this.viewfields + '</View>';
    return query;
  }
  public getQueryNineFilters = (queries: string[]): string => {
    var cond1 = queries[0];
    var cond2 = queries[1];
    var cond3 = queries[2];
    var cond4 = queries[3];
    var cond5 = queries[4];
    var cond6 = queries[5];
    var cond7 = queries[6];
    var cond8 = queries[7];
    var cond9 = queries[8];
    var query = '<View><Query><Where><And>' + cond1 + '<And>' + cond2 + '<And>' + cond3 + '<And>' + cond4 + '<And>' + cond5 + '<And>' + cond6 + '<And>' + cond7 + '<And>' + cond8 + cond9 + '</And></And></And></And></And></And></And></And></Where></Query>' + this.viewfields + '</View>';
    return query;
  }

  private _closeElement(name: string, clause: string) {
    return (`${clause}</${name}>
           `);
  }

  private _splitkeyWords(str: string): string[] {
    let strArr: string[] = [];
    let strVal: string = "";
    var tempChar = this.splitChars[0]; // We can use the first token as a temporary join character
    for (var i = 0; i < this.splitChars.length; i++) {
      strVal = str.split(this.splitChars[i]).join(tempChar);
      strArr.push(strVal.trim());
    }
    strArr = str.split(tempChar);
    return strArr;
  }

  private _orCount(query: string) {
    var orOpen = (query.match(/<Or>/g) || []).length;
    var orClosed = (query.match(/<\/Or>/g) || []).length;
    return orOpen - orClosed;
  }

  private _onKeywordDoubleClick() {
    this.setState({ searchEnabled: true });
  }

  /***********************************************************************************************************************************/
  public SearchDoc(areaLookup: Map<number, string>, topicLookup: Map<number, string>, legislationLookup: Map<number, string>, docTypeLookup: Map<number, string>, yearLookup: Map<number, string>) {
   // console.log("Search Clicked");
    let queries = [];
    this.itemList = [];
    var n: number = 0;
    var filterCount = 0;
    var cond: string = "";
    var cond1: string = "";
    var cond2: string = "";
    var orCount: number = 0;
    var finalQuery: string = "";
    let libSelect: string = document.getElementById('drLib-input').getAttribute('value');
    this.finalCount = 0;
    let selArea1: boolean = this.selectedListTitleArea1 !== undefined;
    let selTopic1: boolean = this.selectedListTitleTopic1 !== undefined;
    let selArea2: boolean = this.selectedListTitleArea2 !== undefined;
    let selTopic2: boolean = this.selectedListTitleTopic2 !== undefined;
    document.getElementById('dv_pagination').style.display = 'none';
    this.setState({ selectedSort: 'File Name' });
    this.setState({ selectedLibrary: this.selectedLibrary });

    // First condition: Approved items only

    cond1 = `<Eq>
               <FieldRef Name="_ModerationStatus"/>
               <Value Type="ModStat">${this.statusFilter}</Value>
             </Eq>
             `;

    // Now, add filter for key prase or individual keywords
    if (this.keyPhraseValue !== undefined) {
      let keyWords: string[] = this._splitkeyWords(this.keyPhraseValue);
      let keyPhraseLength = this.keyPhraseValue.length;
      let firstChar = this.keyPhraseValue.substring(0, 1);
      let lastChar = this.keyPhraseValue.substring(keyPhraseLength - 1, keyPhraseLength);
      if ((firstChar === '"' && lastChar === '"') || (firstChar === "'" && lastChar === "'")) {
        let keyPhrase = this.keyPhraseValue.substring(1, keyPhraseLength - 1);
        cond2 = ` <Or>
                    <Contains>
                      <FieldRef Name="LinkFilename"/>
                      <Value Type="Text">${keyPhrase}</Value>
                    </Contains>
                    <Contains>
                      <FieldRef Name="RoutingRuleDescription"/>
                      <Value Type="Text">${keyPhrase}</Value>
                    </Contains>
                  </Or>
                `;
      } else if (keyWords.length === 1) {
        cond2 = ` <Or>
                    <Contains>
                      <FieldRef Name="LinkFilename"/>
                      <Value Type="Text">${keyWords[0]}</Value>
                    </Contains>
                    <Contains>
                      <FieldRef Name="RoutingRuleDescription"/>
                      <Value Type="Text">${keyWords[0]}</Value>
                    </Contains>
                  </Or>
                `;
      } else {
        cond2 = "";
        for (var i = 0; i < keyWords.length; i++) {
          var j = i + 1;
          if (j < keyWords.length) {
            cond2 = `${cond2}
                     <Or>
                       <Contains>
                         <FieldRef Name="LinkFilename"/>
                         <Value Type="Text">${keyWords[i]}</Value>
                       </Contains>
                       <Or>
                         <Contains>
                           <FieldRef Name="RoutingRuleDescription"/>
                           <Value Type="Text">${keyWords[i]}</Value>
                         </Contains>
              `;
          } else {
            cond2 = `${cond2}
                     <Or>
                       <Contains>
                         <FieldRef Name="LinkFilename"/>
                         <Value Type="Text">${keyWords[i]}</Value>
                       </Contains>
                       <Contains>
                         <FieldRef Name="RoutingRuleDescription"/>
                         <Value Type="Text">${keyWords[i]}</Value>
                       </Contains>
                    </Or>`;
          }
        }
      }

      orCount = this._orCount(cond2);
      if (orCount > 0) {
        for (var i = 1; i <= orCount; i++) {
          cond2 = this._closeElement('Or', cond2);
        }
      }

      cond = `<And>
                ${cond1}
                ${cond2}
              </And>
              `;
    } else {
      cond = cond1;
    }

    queries.push(cond);
    n++;
    filterCount++;
    cond = '';
    cond1 = '';
    cond2 = '';

    if (selArea1 && !selTopic1 && !selArea2 && !selTopic2) {
      cond = `
              <Or>
                <Contains>
                  <FieldRef Name="LSB_AreaOfLaw"/>
                  <Value Type="Lookup">${this.selectedListTitleArea1}</Value>
                </Contains>
                <Contains>
                  <FieldRef Name="LSB_AreaOfLaw2"/>
                  <Value Type="Lookup">${this.selectedListTitleArea1}</Value>
                </Contains>
              </Or>
              `;
      filterCount = 1;
    } else if (selTopic1 && !selArea2 && !selTopic2) {
      cond = `
              <And>
                <Or>
                  <Contains>
                    <FieldRef Name="LSB_AreaOfLaw"/>
                      <Value Type="Lookup">${this.selectedListTitleArea1}</Value>
                    </Contains>
                  <Contains>
                    <FieldRef Name="LSB_AreaOfLaw2"/>
                    <Value Type="Lookup">${this.selectedListTitleArea1}</Value>
                  </Contains>
                </Or>
                <Or>
                  <Contains>
                    <FieldRef Name="LSB_TopicsOfLaw"/>
                    <Value Type="Lookup">${this.selectedListTitleTopic1}</Value>
                  </Contains>
                  <Contains>
                    <FieldRef Name="LSB_TopicsOfLaw2"/>
                      <Value Type="Lookup">${this.selectedListTitleTopic1}</Value>
                    </Contains>
                  </Or>
                </And>
                `;
      filterCount = 2;
    } else if (selArea2 && !selTopic2) {
      cond = `
              <Or>
                <And>
                  <Or>
                    <Contains>
                      <FieldRef Name="LSB_AreaOfLaw"/>
                      <Value Type="Lookup">${this.selectedListTitleArea1}</Value>
                    </Contains>
                      <Contains>
                        <FieldRef Name="LSB_AreaOfLaw2"/>
                        <Value Type="Lookup">${this.selectedListTitleArea1}</Value>
                      </Contains>
                    </Or>
                    <Or>
                      <Contains>
                        <FieldRef Name="LSB_TopicsOfLaw"/>
                          <Value Type="Lookup">${this.selectedListTitleTopic1}</Value>
                          </Contains>
                        <Contains>
                          <FieldRef Name="LSB_TopicsOfLaw2"/>
                          <Value Type="Lookup">${this.selectedListTitleTopic1}</Value>
                        </Contains>
                      </Or>
                    </And>
                    <Or>
                      <Contains>
                        <FieldRef Name="LSB_AreaOfLaw"/>
                        <Value Type="Lookup">${this.selectedListTitleArea2}</Value>
                      </Contains>
                      <Contains>
                        <FieldRef Name="LSB_AreaOfLaw2"/>
                        <Value Type="Lookup">${this.selectedListTitleArea2}</Value>
                      </Contains>
                    </Or>
                  </Or>
                  `;
      filterCount = 3;
    } else if (selTopic2) {
      cond = `
              <Or>
                <And>
                  <Or>
                    <Contains>
                      <FieldRef Name="LSB_AreaOfLaw"/>
                      <Value Type="Lookup">${this.selectedListTitleArea1}</Value>
                    </Contains>
                    <Contains>
                      <FieldRef Name="LSB_AreaOfLaw2"/>
                        <Value Type="Lookup">${this.selectedListTitleArea1}</Value>
                      </Contains>
                  </Or>
                  <Or>
                    <Contains>
                      <FieldRef Name="LSB_TopicsOfLaw"/>
                      <Value Type="Lookup">${this.selectedListTitleTopic1}</Value>
                    </Contains>
                    <Contains>
                      <FieldRef Name="LSB_TopicsOfLaw2"/>
                        <Value Type="Lookup">${this.selectedListTitleTopic1}</Value>
                    </Contains>
                  </Or>
                </And>
                <And>
                  <Or>
                    <Contains>
                      <FieldRef Name="LSB_AreaOfLaw"/>
                      <Value Type="Lookup">${this.selectedListTitleArea2}</Value>
                    </Contains>
                    <Contains>
                      <FieldRef Name="LSB_AreaOfLaw2"/>
                      <Value Type="Lookup">${this.selectedListTitleArea2}</Value>
                    </Contains>
                  </Or>
                  <Or>
                    <Contains>
                      <FieldRef Name="LSB_TopicsOfLaw"/>
                      <Value Type="Lookup">${this.selectedListTitleTopic2}</Value>
                    </Contains>
                    <Contains>
                      <FieldRef Name="LSB_TopicsOfLaw2"/>
                      <Value Type="Lookup">${this.selectedListTitleTopic2}</Value>
                    </Contains>
                  </Or>
                </And>
              </Or>
              `;
      filterCount = 4;
    }

    if (cond !== '') {
      queries.push(cond);
      n++;
      cond = '';
    }


    if (this.selectedListTitleLegislation !== undefined) {
  const cond = `
  <Contains>
      <FieldRef Name="LSB_Legislation"/>
      <Value Type="Lookup">${this.selectedListTitleLegislation}</Value>
      </Contains>
  `;
      queries.push(cond);
      n++;
      filterCount++;
     // console.log("Legislation query:", cond); // Added console.log
    }
    if (this.selectedListTitleDocType !== undefined) {
      cond = `
              <Eq>
                <FieldRef Name="Document_x0020_Type"/>
                <Value Type="Text">${this.selectedListTitleDocType}</Value>
              </Eq>
              `;
      queries.push(cond);
      n++;
      filterCount++;
    }
    if (this.selectedListTitleAuthor !== undefined) {
      cond = `
              <Contains>
                <FieldRef Name="LSB_AuthorNames"/>
                <Value Type="Lookup">${this.selectedListTitleAuthor}</Value>
              </Contains>
              `;
      queries.push(cond);
      n++;
      filterCount++;
    }
    if (this.selectedListTitleYear !== undefined) {
      cond = `
              <Eq>
                <FieldRef Name="Year"/>
                <Value Type="Choice">${this.selectedListTitleYear}</Value>
              </Eq>
              `;
      queries.push(cond);
      n++;
      filterCount++;
    }
    if (this.selectedConflict !== undefined) {
      cond = `
              <Eq>
                <FieldRef Name="LSB_Conflict"/>
                <Value Type="Text">${this.selectedConflict}</Value>
              </Eq>
              `;
      queries.push(cond);
      n++;
      filterCount++;
    }

    // Call specific function based on number of parameters
    console.log("Filter count: " + filterCount + " (including Approval Status = 'Approved')");

    switch (n) {
      case 1: finalQuery = this.getQueryOneFilter(queries);
        break;
      case 2: finalQuery = this.getQueryTwoFilters(queries);
        break;
      case 3: finalQuery = this.getQueryThreeFilters(queries);
        break;
      case 4: finalQuery = this.getQueryFourFilters(queries);
        break;
      case 5: finalQuery = this.getQueryFiveFilters(queries);
        break;
      case 6: finalQuery = this.getQuerySixFilters(queries);
        break;
      case 7: finalQuery = this.getQuerySevenFilters(queries);
        break;
      case 8: finalQuery = this.getQueryEightFilters(queries);
        break;
      case 9: finalQuery = this.getQueryNineFilters(queries);
        break;
      default:
    }

    document.getElementById('dv_Table').style.display = 'none';

   // console.log("CAML Query: " + finalQuery);

    this.itemList = [];
    this.libsToQuery = [];

    libSelect = document.getElementById('drLib-input').getAttribute('value');
    if (libSelect === null || libSelect === '') libSelect = 'All';

    if (libSelect === 'All') {
      for (var i = 0; i <= this.lookupLibs.size; i++) {
        let libValue = this.lookupLibs.get(i);
        if (libValue !== undefined) this.libsToQuery.push(libValue);
      }
    } else {
      this.libsToQuery.push(libSelect);
    }

    for (var i = 0; i < this.libsToQuery.length; i++) {

      let lib = this.libsToQuery[i];
     // console.log("Doc library: " + lib);

      // Get the documents by passing the dynamic query using pnp.js
      sp.web.lists.getByTitle(lib).getItemsByCAMLQuery({
        ViewXml: finalQuery
      }, "File").then((results: any) => {
       // console.log("Results for library", lib, ":", results); // Added console.log
        results.map((result: any) => {

          let keep: boolean = true;

          let fil = result.File;
          let are = areaLookup.get(result.LSB_AreaOfLawId);
          if (!are) are = '';
          let topFound: boolean = false;
          let top = this._joinLookupArrayValues(result.LSB_TopicsOfLawId, topicLookup, '; ');
        //  console.log("Topics values:", top); // Debugging output
          if (!top) top = this._joinLookupArrayValues(result.LSB_TopicsOfLaw2Id, topicLookup, '; ')
          if (!top) top = '';
          if (selTopic1) topFound = this.containsVal(top, this.selectedListTitleTopic1);
          let ar2 = areaLookup.get(result.LSB_AreaOfLaw2Id);
          if (!ar2) ar2 = '';
          let to2Found: boolean = false;
          let to2 = this._joinLookupArrayValues(result.LSB_TopicsOfLaw2Id, topicLookup, '; ');
          if (!to2) to2 = this._joinLookupArrayValues(result.LSB_TopicsOfLawId, topicLookup, '; ')
          if (!to2) to2 = '';
          if (selTopic2) to2Found = this.containsVal(to2, this.selectedListTitleTopic2);

          if (selTopic1 || selTopic2 || topFound || to2Found) keep = true;
          else if (are !== undefined) keep = true;

          if (keep) {
            //let leg = legislationLookup.get(result.LSB_LegislationId);
            let leg = this._joinLookupArrayValues(result.LSB_LegislationId, legislationLookup, '; '); // Updated line
           // console.log("Legislation ID:", result.LSB_LegislationId, "Legislation Value:", leg); // Added console.log
            if (!leg) leg = '';
           // console.log("Legislation values:", leg); // Debugging output
            let aut = this._joinLookupArrayValues(result.LSB_AuthorNamesId, this.authorLookup, '; ');
            if (!aut) aut = '';
            let cfl = result.LSB_Conflict;
            if (!cfl) cfl = '';
            let typ = result.Document_x0020_Type;
            if (!typ) typ = '';
            let num = result.LSB_LegalFileNumber;
            let dsc = result.RoutingRuleDescription;
            if (!dsc) dsc = '';
            try {
              let url = fil.LinkingUri;
              if (url === null) url = fil.ServerRelativeUrl;
              let lbp = fil.ServerRelativeUrl.replace(this.sitePath, '');
              let iX = lbp.indexOf('/');
              if (iX === 0) iX = 1;
              let urf = lbp.substring(iX, lbp.indexOf('/', iX + 1));
              let lib = this.lookupLibByUrlFragment.get(urf);
              let lbu = this.listRootUrl + this.lookupUrlFragmentByLibName.get(lib);
              let yea = result.Year;
              if (!yea) yea = '';

              if (fil.Name) {
                let indItem = {
                  Name: fil.Name,
                  LSB_AreaOfLaw1: are,
                  LSB_TopicsOfLaw1: top,
                  LSB_AreaOfLaw2: ar2,
                  LSB_TopicsOfLaw2: to2,
                  LSB_Legislation: leg,
                  LSB_AuthorNames: aut,
                  LSB_Conflict: cfl,
                  Document_x0020_Type: typ,
                  LSB_LegalFileNumber: num,
                  RoutingRuleDescription: dsc,
                  LSB_LibName: lib,
                  LSB_LibURL: lbu,
                  URL: url,
                  Year: yea,
                }
                this.itemList.push(indItem);

                this.finalCount++;
               // console.log("Processed Item:", indItem); // Debugging output
                this.pagesTotal = Math.ceil(this.finalCount / this.pageSize);
                if (this.pagesTotal >= this.pagesMax) this.pagesTotal = this.pagesMax;
              }
            } catch { /* empty */ }
          }
        })

        if (this.selectedLibrary === undefined) {
          this.selectedLibrary = 'All';
          this.setState({ selectedLibrary: this.selectedLibrary })
        }

        if (this.selectedSort === undefined) {
          this.selectedSort = 'File Name';
          this.setState({ selectedSort: this.selectedSort })
        }

        switch (this.selectedSort) {
          case 'File Name':
            this.itemList.sort(function (a: any, b: any) {
              if (a.Name < b.Name) {
                return -1;
              }
              if (a.Name > b.Name) {
                return 1;
              }
              return 0;
            });
            break;
          case 'Area of Law (1)':
            this.itemList.sort(function (a: any, b: any) {
              if (a.LSB_AreaOfLaw1 < b.LSB_AreaOfLaw1) {
                return -1;
              }
              if (a.LSB_AreaOfLaw1 > b.LSB_AreaOfLaw1) {
                return 1;
              }
              return 0;
            });
            break;
          case 'Doc Type':
            this.itemList.sort(function (a: any, b: any) {
              if (a.Document_x0020_Type < b.Document_x0020_Type) {
                return -1;
              }
              if (a.Document_x0020_Type > b.Document_x0020_Type) {
                return 1;
              }
              return 0;
            });
            break;
          case 'Author':
            this.itemList.sort(function (a: any, b: any) {
              if (a.LSB_AuthorNames < b.LSB_AuthorNames) {
                return -1;
              }
              if (a.LSB_AuthorNames > b.LSB_AuthorNames) {
                return 1;
              }
              return 0;
            });
            break;
          case 'Year':
            this.itemList.sort(function (a: any, b: any) {
              if (a.Year > b.Year) {
                return -1;
              }
              if (a.Year < b.Year) {
                return 1;
              }
              return 0;
            });
            break;
          case 'Legislation':
            this.itemList.sort(function (a: any, b: any) {
              if (a.LSB_Legislation < b.LSB_Legislation) {
                return -1;
              }
              if (a.LSB_Legislation > b.LSB_Legislation) {
                return 1;
              }
              return 0;
            });
            break;
          case 'Library':
            this.itemList.sort(function (a: any, b: any) {
              if (a.LSB_LibName < b.LSB_LibName) {
                return -1;
              }
              if (a.LSB_LibName > b.LSB_LibName) {
                return 1;
              }
              return 0;
            });
            break;
          default:
        }

        //Results are bound to the state
        this.setState({ items: this.itemList });
        this._getPage(1);
        document.getElementById('dv_SearchResults').style.display = 'table-cell';
        document.getElementById('dv_tableCaption').style.display = 'block';
        document.getElementById('dv_Table').style.display = 'block';
        if (this.finalCount > this.pageSize) document.getElementById('dv_pagination').style.display = 'block';
      })
    }
  }

  private _clearForm() {
    window.location.reload();
  }

  private _joinLookupArrayValues(theArray: number[], theLookup: Map<number, string>, joiner: string) {

    let joined: string = '';
    let i = 0;
    if (theArray) {
      while (i < theArray.length) {
        if (i === 0) joined = theLookup.get(theArray[i]);
        else joined = joined + joiner + theLookup.get(theArray[i]);
        i++;
      }
    }

    return joined;
  }


  /***********************************************************************************************************************************/

  public render(): React.ReactElement<ISearchDocProps> {
    const hasTeamsContext = this.props.hasTeamsContext;
    // Import package version
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const packageSolution: any = require("../../../../config/package-solution.json");

    return (
      <section className={`${styles.searchDoc} ${hasTeamsContext ? styles.teams : ''}`} >

        <div className={styles.dv_ParentDic}>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Area of Law (1):</div>
            <ComboBox id="drpArea1" options={this.state.listTitlesArea1} autoComplete='on' required={false} placeholder={'Select Area of Law (1)'} onChange={this.getSelectedListTitleArea1} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Topics of Law (1): </div>
            <ComboBox id="drTopic1" options={this.state.listAreaTopics1} autoComplete='on' disabled={!this.state.firstTopicEnabled} placeholder={'Select Topics Of Law (1)'} onChange={this.getselectedListTitleTopic1} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Area of Law (2):</div>
            <ComboBox id="drpArea2" options={this.state.listTitlesArea2} autoComplete='on' disabled={!this.state.secondAreaEnabled} placeholder={'Select Area of Law (2)'} onChange={this.getSelectedListTitleArea2} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Topics of Law (2): </div>
            <ComboBox id="drTopic2" options={this.state.listAreaTopics2} autoComplete='on' disabled={!this.state.secondTopicEnabled} placeholder={'Select Topics Of Law (2)'} onChange={this.getselectedListTitleTopic2} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Legislation: </div>
            <ComboBox id="drLegislation" options={this.state.listTitlesLegislation} autoComplete='on' placeholder={'Select Legislation'} onChange={this.getselectedListTitleLegislation} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Document Type: </div>
            <ComboBox id="drDocType" options={this.state.listTitlesDocType} autoComplete='on' placeholder={'Select Document Type'} onChange={this.getSelectedListTitleDocType} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Author: </div>
            <ComboBox id="drAuthor" options={this.state.listTitlesAuthor} autoComplete='on' placeholder={'Select Author'} onChange={this.getSelectedListTitleAuthor} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Year: </div>
            <ComboBox id="drYear" options={this.state.listTitlesYear} autoComplete='on' placeholder={'Select Year'} onChange={this.getSelectedListTitleYear} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Restrictions: </div>
            <ComboBox id="drConflict" options={this.conflict} autoComplete='on' placeholder={'Select Yes or No'} onChange={this.getSelectedConflict} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Library Selection: </div>
            <ComboBox id="drLib" options={this.state.listTitlesLibrary} selectedKey={0} placeholder={'Select Library'} onChange={this.getSelectedLibrary} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Sort by: </div>
            <ComboBox id="drSort" options={this.sortFields} autoComplete='on' defaultValue={'File Name'} defaultSelectedKey={'Name'} placeholder={'Select Sort'} onChange={this.getSelectedSort} />
          </div>
          <div className={styles.inputBlock}>
            <div className={styles.inputHeader}> Keyword Search: </div>
            <div className={styles.descStyle}>
              <input id="keywordPhrase" name="keywordPhrase" className="keywordPhrase" autoComplete='on' type="text" placeholder={'Enter keywords or phrase (in quotes)'} onDoubleClick={this._onKeywordDoubleClick} onChange={this._onKeyPhraseChange} />
            </div>
          </div>
          <div className={styles.divButton}>
            <PrimaryButton id="searchButton" disabled={!this.state.searchEnabled} text="Search" onClick={() => this.SearchDoc(this.areaLookup, this.topicLookup, this.legislationLookup, this.docTypeLookup, this.yearLookup)} >Search</PrimaryButton>&nbsp;
            <PrimaryButton id="clearButton" disabled={!this.state.searchEnabled} text="Clear" alt='Clear form by refreshing page' onClick={() => this._clearForm()} >Clear</PrimaryButton>
          </div>
          <div className="packageVersion">{packageSolution.solution.version}</div>
        </div>
        <div id="dv_resultsContainer">
          <div id="dv_Table">
            <div id="dv_tableCaption" className={styles.tableCaptionStyle} >Document Search Results ({this.finalCount} &quot;{this.statusFilter}&quot; documents found)</div>
            <div id="dv_SearchResults" className={styles.tableStyle}>
              {
                this.state.paginatedItems.map(function (item, key) {
                 // console.log("Item being rendered:", item); // Debugging output
                  return (
                    <div className={styles.CardStyle} key={key}>
                      <div className={styles.DocPropStyle}>
                        <strong><a className={styles.DocLink} href={item.URL} target="_blank">{item.Name}</a></strong>
                      </div>
                      <div className={styles.DocPropStyle}><strong>Library:</strong>&nbsp;
                        <a href={item.LSB_LibURL} target="_blank">{item.LSB_LibName}</a>
                      </div>
                      <div className={styles.DocPropStyle}><strong>Area of Law (1):</strong> {item.LSB_AreaOfLaw1}</div>
                      <div className={styles.DocPropStyle}><strong>Topics (1):</strong> {item.LSB_TopicsOfLaw1}</div>
                      <div className={styles.DocPropStyle}><strong>Area of Law (2):</strong> {item.LSB_AreaOfLaw2}</div>
                      <div className={styles.DocPropStyle}><strong>Topics (2):</strong> {item.LSB_TopicsOfLaw2}</div>
                      <div className={styles.DocPropStyle}><strong>Legislation:</strong> {item.LSB_Legislation}</div>
                      <div className={styles.DocPropStyle}><strong>Restrictions:</strong> {item.LSB_Conflict}</div>
                      <div className={styles.DocPropStyle}><strong>Year:</strong> {item.Year}</div>
                      <div className={styles.DocPropStyle}><strong>Doc Type:</strong> {item.Document_x0020_Type}</div>
                      <div className={styles.DocPropStyle}><strong>Author(s):</strong> {item.LSB_AuthorNames}</div>
                      <div className={styles.DocPropStyle}><strong>Description:</strong> <i>{item.RoutingRuleDescription}</i></div>
                    </div>
                  )
                })
              }
            </div>
            <div id="dv_pagination" className={styles.DocPagination}>
              <Pagination
                currentPage={1}
                totalPages={this.pagesTotal}
                onChange={(page) => this._getPage(page)}
              />
            </div>
          </div>
        </div>
      </section >
    );
  }

  private _getPage(page: number) {
    // round a number up to the next integer.
    let roundUpPage: number = page - 1;

    this.setState({
      paginatedItems: this.state.items.slice(roundUpPage * this.pageSize, (roundUpPage * this.pageSize) + this.pageSize)
    });
  }
}
