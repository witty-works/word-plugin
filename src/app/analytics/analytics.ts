import {
    IAlert,
    IAlternativeLogItems,
    ICheckLogItems,
    ILogItems,
    IAuthResponse,
    ICheckResponse,
    ICheckResponseResult,
    ICheckResultLogItems,
  } from '../data/types';
import {
    captureEvent,
    getRequestData,
    getResponseData,
  } from './analyticsUtils';
  
  export const useAnalytics = () => {
    return {
      async checkLog(
        checkResponse: ICheckResponse,
        authResponse: IAuthResponse | null,
        inputLength: number,
        requestType: string,
        isAutoTriggered: boolean,
        checkLogEventId: string,
      ) {
        const checkResponseResultsWithoutContext = checkResponse.results.map(
          (result) => {
            const { context, ...resultWithoutContext } = result;
            return resultWithoutContext;
          },
        ) as ICheckResponseResult[];
  
        const checkLogItems: ICheckLogItems = {
          request__id: checkLogEventId,
          request__type: requestType,
          request__text__length: inputLength,
          request__is_auto_triggered: isAutoTriggered, //always false in word plugin
          ...getRequestData(),
          response__results: checkResponseResultsWithoutContext,
          response__language: checkResponse.language,
          response__limit_reached: checkResponse.limit_reached,
          response__organizationId: authResponse
            ? authResponse.organization_id
            : undefined,
          response__plan: authResponse ? authResponse.plan : undefined,
        };
  
        captureEvent(requestType, checkLogItems);
      },

      async checkResultLog(
        checkResponse: ICheckResponseResult,
        authResponse: IAuthResponse | null,
        inputLength: number,
        requestType: string,
        isAutoTriggered: boolean,
        checkLogEventId: string,
      ) {
        const checkLogItems: ICheckResultLogItems = {
          request__id: checkLogEventId,
          request__type: requestType,
          request__text__length: inputLength,
          request__is_auto_triggered: isAutoTriggered,
          ...getRequestData(),
          response__data__text: checkResponse.text,
          response__data__category: checkResponse.category,
          response__data__subcategory: checkResponse.subcategory,
          response__data__start: checkResponse.start,
          response__data__end: checkResponse.end,
          response__data__alternatives: checkResponse.alternatives,
          response__data__label: checkResponse.label,
          response__data__explanation__text: checkResponse.explanation?.text,
          response__data__explanation__icon: checkResponse.explanation?.icon,
          response__data__explanation__icon_image: checkResponse.explanation?.icon_image,
          response__data__explanation__url: checkResponse.explanation?.url,
          response__data__gravity: checkResponse.gravity,
          response__language: checkResponse.language,
          response__limit_reached: checkResponse.limit_reached,
          response__organizationId: authResponse
            ? authResponse.organization_id
            : undefined,
          response__plan: authResponse ? authResponse.plan : undefined,
        };
  
        captureEvent(requestType, checkLogItems);
      },
  
      async alternativeLog(logResponse: IAlert, alternative: string) {
        const alternativeLogItems: IAlternativeLogItems | any = {
          request__type: 'alternative',
          request__alternative: alternative,
          ...getRequestData(),
          ...getResponseData(logResponse),
        };
  
        captureEvent('alternative', alternativeLogItems);
      },
  
      async popoverLogs(logResponse: IAlert, logType: string) {
        const popoverLogItems: ILogItems | any= {
          request__type: logType,
          ...getRequestData(),
          ...getResponseData(logResponse),
        };
  
        captureEvent(logType, popoverLogItems);
      },

      async openLinkLog(logType: string) {
        const popoverLogItems = {
          request__type: logType,
        };
  
        captureEvent(logType, popoverLogItems);
      },
    };
  };