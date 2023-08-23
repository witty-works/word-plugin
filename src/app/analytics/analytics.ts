import {
    IAlert,
    IAlternativeLogItems,
    ICheckLogItems,
    ILogItems,
    IAuthResponse,
    ICheckResponse,
    ICheckResponseResult,
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
      ) {
        const checkResponseResultsWithoutContext = checkResponse.results.map(
          (result) => {
            const { context, ...resultWithoutContext } = result;
            return resultWithoutContext;
          },
        ) as ICheckResponseResult[];
  
        const checkLogItems: ICheckLogItems = {
          request__id: 'word-plugin', //no id needed for word plugin
          request__type: requestType,
          request__text__length: inputLength,
          request__is_auto_triggered: false, //no auto trigger for word plugin
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