/* eslint-disable @typescript-eslint/no-unused-vars */

import { IDropdownOption } from "office-ui-fabric-react";

export interface ISearchDocState {
  listTitlesArea1: IDropdownOption[],
  listTitlesArea2: IDropdownOption[],
  listAreaTopics1: IDropdownOption[],
  listAreaTopics2: IDropdownOption[],
  listTitlesLegislation: IDropdownOption[],
  listTitlesDocType: IDropdownOption[],
  listTitlesAuthor: IDropdownOption[], 
  listTitlesYear: IDropdownOption[],
  listTitlesLibrary: IDropdownOption[],
  keyPhrase: string,
  searchEnabled: boolean,
  firstTopicEnabled: boolean,
  secondAreaEnabled: boolean,
  secondTopicEnabled: boolean,
  selectedSort: string,
  selectedLibrary: string,
  items:
  {
    LSB_AreaOfLaw1?: string,
    LSB_AreaOfLaw2?: string,
    LSB_AuthorNames?: string,
    LSB_Conflict?: string,
    Document_x0020_Type?: string,
    LSB_LegalFileNumber?: string,
    LSB_Legislation?: string,
    LSB_LibName?: string,
    LSB_LibURL?: string,
    LSB_TopicsOfLaw1?: string,
    LSB_TopicsOfLaw2?: string,
    Year?: string,
    Name?: string,
    RoutingRuleDescription?: string,
    URL?: string,
  }[],
  paginatedItems:
  {
    LSB_AreaOfLaw1?: string,
    LSB_AreaOfLaw2?: string,
    LSB_AuthorNames?: string,
    LSB_Conflict?: string,
    Document_x0020_Type?: string,
    LSB_LegalFileNumber?: string,
    LSB_Legislation?: string,
    LSB_LibName?: string,
    LSB_LibURL?: string,
    LSB_TopicsOfLaw1?: string,
    LSB_TopicsOfLaw2?: string,
    Year?: string,
    Name?: string,
    RoutingRuleDescription?: string,
    URL?: string,
  }[]
}
