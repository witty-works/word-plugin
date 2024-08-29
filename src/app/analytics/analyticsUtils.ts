import PostHog from 'posthog-js-lite'
import { IAlert } from '../data/types';
import { environment } from 'src/environments/environment';

export const DEV_ENV = window.location.hostname === 'localhost';

export const POSTHOG_API_KEY = DEV_ENV
  ? 'phc_QiISRw0yFAsndXqYD0HmfGvHaOBMxb57ZRIxlimvR64'
  : 'phc_i1tlvuh1iecIOSEr0QmTEIklrsSJGhULpUwUlf8fkkl';

export const POSTHOG_API_URL = DEV_ENV
  ? 'https://app.posthog.com'
  : 'https://eu.posthog.com';

export const captureEvent = (eventName: string, eventData: object) => {
    const userId = localStorage.getItem('user_id');
    const organizationId = localStorage.getItem('organization_id');
    const appId = 'test-app-id'

  const ph = new PostHog(POSTHOG_API_KEY, {
    host: POSTHOG_API_URL,
    bootstrap: {
      distinctId: userId || 'unknown',
    },
  })

    if (organizationId) {
      ph.capture(eventName, {
        ...eventData,
        request__app_id: appId, //@lukas what is this used for? 
        $groups: {
          organization: organizationId,
        },
      });
    } else {
      ph.capture(eventName, {
        ...eventData,
      });
    }
};

export const getResponseData = (logResponse: IAlert) => {
  if (!logResponse) return {};
  return {
    response__id: logResponse.id,
    response__organizationId: logResponse.organizationId,
    response__startOffset: logResponse.startOffset,
    response__endOffset: logResponse.endOffset,
    response__popOverIsOpen: logResponse.popOverIsOpen,
    response__plan: logResponse.plan,
    response__data__language: logResponse.data.language,
    response__data__category: logResponse.data.category,
    response__data__subcategory: logResponse.data.subcategory,
    response__data__context: logResponse.data.context,
    response__data__text: logResponse.data.text,
    response__data__label: logResponse.data.label,
    response__data__explanation__text: logResponse.data.explanation?.text,
    response__data__explanation__icon: logResponse.data.explanation?.icon,
    response__data__explanation__icon_image: logResponse.data.explanation?.icon_image,
    response__data__explanation__url: logResponse.data.explanation?.url,
    response__data__alternatives: logResponse.data.alternatives,
    response__data__gravity: logResponse.data.gravity,
  };
};

export const getRequestData = () => {
  return {
    request__lang: 'auto',
    request__client: environment.package_version,
  };
};