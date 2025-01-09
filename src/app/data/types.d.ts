export interface DiffChange {
  added: boolean,
  removed: boolean,
  value: string
}

//API REQUEST/RESPONSE
export interface IRequest {
  url: string;
  config: RequestInit | null;
}

export interface ICheckRequest {
  text: string;
  repeatedRequest: boolean;
}

export interface IEndpointError {
  status: number;
  message: string;
  request?: IRequest;
  responseSchema?: any;
}

export interface ResponseConfig {
  orthography: ConfigProperty;
}
export interface RequestConfig extends ResponseConfig {
  disabled_categories: string[];
}

export interface FilteredRequestConfig {
  disabled_categories: string[];
}
export interface ConfigProperty {
  value: string | string[] | boolean | number;
  status?: string;
}

//CHECK ENDPOINT
export interface ICheckResponse {
  results: ICheckResponseResult[];
  language: string;
  limit_reached: boolean;
  config_changed: boolean;
  notifications: number;
  gender_separator: string;
}

export interface ICheckResultLogItems {
  request__id: string;
  request__type: string;
  request__lang: string;
  request__client: string;
  request__text__length: number;
  request__is_auto_triggered: boolean;
  response__organizationId?: string;
  response__plan?: string;
  response__data__text: string;
  response__data__category: string;
  response__data__subcategory: string;
  response__data__start: number;
  response__data__end: number;
  response__data__alternatives: IAlternative[];
  response__data__label: string;
  response__data__explanation__text: string;
  response__data__explanation__icon: string;
  response__data__explanation__icon_image: string;
  response__data__explanation__url: string;
  response__data__gravity: number;
  response__language: string;
  response__limit_reached: boolean;
}

export interface ICheckResponseResult {
  text: string;
  context: string;
  category: string;
  subcategory: string;
  start: number;
  end: number;
  alternatives: IAlternative[];
  explanation: IExplanation;
  label: string;
  gravity: number;
  language: string;
  limit_reached: boolean;
  source: ISource;
}

export interface IRephrasingResult {
  sentence: string,
  results: Map<string, string>,
}

//AUTH ENDPOINT
export interface IAuthResponse {
  config: ResponseConfig;
  organization_config: ResponseConfig;
  plan: string;

  //private account
  id: string;
  name: string;
  domains: IDomains;
  config_hash: string;

  //organization account
  organization_id?: string;
  organization_name?: string;
  organization_domains: IDomains;
  organization_config_hash: string;
}

export interface IDomains {
  list: string[];
  type: string;
}
// export interface IRefreshTokenResponse {
//   email: string;
//   refresh_token: string;
//   access_token: string;
// }

//HIGHLIGHTS
export interface Position {
  top: number;
  left: number;
}
export interface Highlight {
  rects: DOMRect[];
  id: string;
  data: IAlertContentData;
  startOffset: number;
  endOffset: number;
  node: Node;
  plan?: string;
}

export type CustomInputElement =
  | HTMLTextAreaElement
  | HTMLInputElement
  | HTMLDivElement;

//ALERTS
export interface INodeWithAlerts {
  node: any;
  alerts: IAlert[];
  nodeIndex?: number;
}

export interface INodes {
  node: any;
  index: number;
  rawNode: Node;
}
export interface IAlert {
  id: string;
  startOffset: number;
  endOffset: number;
  popOverIsOpen: boolean;
  data: IAlertContentData;
  organizationId?: string;
  userId?: string;
  plan?: string;
  rect?: any;
}
export interface IAlertContentData {
  text: string;
  context: string;
  category: string;
  subcategory: string;
  alternatives: IAlternative[];
  label: string;
  explanation: IExplanation;
  language: string;
  gravity: number;
}

//POPOVER
export interface IAlternative {
  text: string;
  remove: boolean;
  inspiration: boolean;
  collective_noun?: ConstrainBooleanParameters;
  male_form?: string;
  female_form?: string;
  context: string;
  hovered?: boolean;
  url: string;
}

export interface IExplanation {
  text: string;
  icon: string;
  icon_image: string;
  url: string;
  context: string;
  content: string;
}

export interface ISource {
  text: string;
  url: string;
}

//ANALYTICS
export interface ILogItems {
  request__type: string;
  request__lang: string;
  request__client: string;
  response__id: string;
  response__startOffset: number;
  response__endOffset: number;
  response__popOverIsOpen: boolean;
  response__plan?: string;
  response__data__language: string;
  response__data__category: string;
  response__data__subcategory: string;
  response__data__context: string;
  response__data__text: string;
  response__data__label: string;
  response__data__explanation__text: string;
  response__data__explanation__icon: string;
  response__data__explanation__icon_image: string;
  response__data__explanation__url: string;
  response__data__source: ISource;
  response__data__alternatives: IAlternative[];
  response__data__gravity: number;
}
export interface IAlternativeLogItems extends ILogItems {
  request__alternative: string;
}
export interface IIgnoreLogItems extends ILogItems { }
export interface IVoteLogRequest {
  request__type: string;
  request__lang: string;
  request__client: string;
  vote__url: string;
}

export interface IDashboardLogRequest {
  request__type: string;
  request__lang: string;
  request__client: string;
  button__location: string;
}
export interface ICheckLogItems {
  request__id: string;
  request__type: string;
  request__lang: string;
  request__client: string;
  request__text__length: number;
  request__is_auto_triggered: boolean;
  response__organizationId?: string;
  response__plan?: string;
  response__results: ICheckResponseResult[];
  response__language: string;
  response__limit_reached: boolean;
}
export type DefaultConfigValue =
  | string
  | boolean
  | number
  | string[]
  | object
  | (() => string);

export interface IDomainRequest {
  domain: string;
  enabled: boolean;
}

export interface EnableWittyToggle {
  enabled: boolean;
  updateDashboard: boolean;
}

export type BaseUrl = {
  api: string;
  dashboard: string;
  plugin: string;
};

export interface IBaseUrls {
  [key: string]: BaseUrl;
}