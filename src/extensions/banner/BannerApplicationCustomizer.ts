import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { sp, SPConfiguration, SearchQuery, SearchResults } from '@pnp/sp';

import * as strings from 'BannerApplicationCustomizerStrings';
import styles from './Banner.module.scss';

const LOG_SOURCE: string = 'ExternalSharingBanner.ApplicationCustomizer';
const CACHE_KEY: string = 'ExternalSharingBannerCache';

export interface IBannerApplicationCustomizerProperties {
  externalSharingText: string;
  textColor: string;
  backgroundColor: string;
  cacheEnabled: boolean;
  cacheLifetimeMins: number;
}

const DEFAULT_PROPERTIES: IBannerApplicationCustomizerProperties = {
  externalSharingText: "This site contains content which is viewable by external users",
  textColor: "#333",
  backgroundColor: "#ffffc6",
  cacheEnabled: true,
  cacheLifetimeMins: 15
};

/** A Custom Action which can be run during execution of a Client Side Application */
export default class BannerApplicationCustomizer
  extends BaseApplicationCustomizer<IBannerApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent;
  private _extensionProperties: IBannerApplicationCustomizerProperties;
  private _bannerIsShowing: boolean = false;

  @override
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

      // Initialize PnP SP library
      sp.setup({ spfxContext: this.context } as SPConfiguration);

      // Merge passed properties with default properties, overriding any defaults
      this._extensionProperties = { ...DEFAULT_PROPERTIES, ...this.properties };

      // Event handler to re-render banner on each page navigation (detect if we change sites)
      this.context.application.navigatedEvent.add(this, this.onNavigated);
    });
  }

  @override
  public onDispose(): void {
    if (this._topPlaceholder) {
      this._topPlaceholder.dispose();
    }
  }

  /**
   * Event handler that fires on every page load
   */
  private async onNavigated(): Promise<void> {
    // Get external sharing status
    const isExternalSharingEnabled = await this.getExternalSharingStatus();
    // If external sharing is enabled, render banner, otherwise reset/clear banner if it was showing
    isExternalSharingEnabled ? this.renderBanner() : this.hideBanner();
  }

  /**
   * Get external sharing status for the current site collection from cache or REST API
   */
  private async getExternalSharingStatus(): Promise<boolean> {
    const { cacheLifetimeMins } = this._extensionProperties;
    const siteId = this.context.pageContext.site.id.toString();

    // Check cache if it is enabled
    if (this._extensionProperties.cacheEnabled) {
      const cache = this.cacheGet(CACHE_KEY, siteId);
      if (null !== cache && undefined !== cache.isExternalSharingEnabled) {
        return cache.isExternalSharingEnabled;
      }
    }

    try {
      const isExternalSharingEnabled: boolean = await this.fetchExternalSharingStatus(siteId);

      // Check if cache is enabled, then cache result for the site with specified cache lifetime in minutes
      if (this._extensionProperties.cacheEnabled) {
        this.cacheSet(CACHE_KEY, siteId, { isExternalSharingEnabled }, (1000 * 60 * cacheLifetimeMins));
      }

      return isExternalSharingEnabled;
    }
    catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to retrieve current site from search REST API. Error: ${error}`));

      return false;
    }
  }

  /**
   * Fetch external sharing status from Search REST API by checking if any content exists on the site that is viewable by external users
   * @param siteId Unique ID of the site
   */
  private async fetchExternalSharingStatus(siteId: string): Promise<boolean> {
    try {
      // Construct search query using PnP SP
      // Use HiddenConstraints so we don't clutter native search analytics
      const searchQuery: SearchQuery = {
        Querytext: '*',
        QueryTemplate: `{searchterms} SiteId:"${siteId}" (ViewableByExternalUsers:1 OR ViewableByAnonymousUsers:1)`,
        SelectProperties: ['Title', 'SiteID', 'ViewableByExternalUsers'],
        ClientType: 'Custom',
        RowLimit: 1
      };
      const results: SearchResults = await sp.search(searchQuery);

      //If we got any results back, then the site has content which is viewable by external users
      return results.TotalRows > 0;
    }
    catch (error) {
      Log.error(LOG_SOURCE, new Error(`Failed to retrieve data from search REST API. Error: ${error}`));
      return false;
    }
  }

  /**
   * Render the 'content viewable by external users' banner on the current page
   */
  private renderBanner(): void {
    const { externalSharingText, backgroundColor, textColor } = this._extensionProperties;

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

      if (!this._topPlaceholder) {
        Log.error(LOG_SOURCE, new Error(`Unable to render Top placeholder`));
        return;
      }
    }

    //Only re-render banner if it isn't already showing
    if (!this._bannerIsShowing) {
      this._topPlaceholder.domElement.innerHTML = `
        <div style="background-color: ${backgroundColor};">
          <div class="${styles.BannerTextContainer}" style="height:0px;">
            <span style="color: ${textColor};">${externalSharingText}</span>
          </div>
        </div>
      `;

      //Animate expand
      setTimeout(() => {
        let elem = document.querySelector(`.${styles.BannerTextContainer}`);
        elem.className += ` ${styles.BannerTextContainerExpand}`;
        this._bannerIsShowing = true;
      }, 1);
    }
  }

  /**
   * Hide the 'content viewable by external users' banner if it is showing on the current page
   */
  private hideBanner(): void {
    if (this._topPlaceholder) {
      //Animate collapse
      let elem = document.querySelector(`.${styles.BannerTextContainer}`);
      elem.className = elem.className.replace(` ${styles.BannerTextContainerExpand}`, '');
    }
    this._bannerIsShowing = false;
  }

  /**
   * Helper function to get an item by key and subkey from session storage
   */
  private cacheGet = (cacheKey, cacheSubKey) => {
    try {
      const valueStr = sessionStorage.getItem(cacheKey);
      if (valueStr) {
        const val = JSON.parse(valueStr);
        const subVal = val[cacheSubKey];
        if (subVal) {
          return !(subVal.expiration && Date.now() > subVal.expiration) ? subVal.payload : null;
        }
      }
      return null;
    }
    catch (error) {
      Log.warn(LOG_SOURCE, `Unable to get cached data from sessionStorage with key ${cacheKey}. Error: ${error}`);
      return null;
    }
  }

  /**
   * Helper function to set an item by key and subkey into session storage
   */
  private cacheSet = (cacheKey, cacheSubKey, payload, expiresIn) => {
    try {
      const nowTicks = Date.now();
      const expiration = (expiresIn && nowTicks + expiresIn) || null;
      let cache = this.cacheGet(cacheKey, cacheSubKey) || {};
      cache[cacheSubKey] = { payload, expiration };
      sessionStorage.setItem(cacheKey, JSON.stringify(cache));
      return this.cacheGet(cacheKey, cacheSubKey);
    }
    catch (error) {
      Log.warn(LOG_SOURCE, `Unable to set cached data into sessionStorage with key ${cacheKey}. Error: ${error}`);
      return null;
    }
  }

  // Old way used to check the ViewableByExternalUsers property on the site result which appeared to be inconsisent
  // private async fetchExternalSharingStatus(siteId: string): Promise<boolean> {
  //   try {
  //     // Construct search query using PnP SP
  //     // Use HiddenConstraints so we don't clutter native search analytics
  //     const searchQuery: SearchQuery = {
  //       HiddenConstraints: `ContentClass:STS_Site AND SiteID:${siteId}`,
  //       SelectProperties: ['SiteID', 'ViewableByExternalUsers'],
  //       ClientType: 'Custom',
  //       RowLimit: 1
  //     };
  //     const results: SearchResults = await sp.search(searchQuery);

  //     // Check if we got one result back for the current site
  //     let isExternalSharingEnabled;
  //     if (results.PrimarySearchResults.length === 1) {
  //       isExternalSharingEnabled = results.PrimarySearchResults[0]["ViewableByExternalUsers"] === "true";
  //     }
  //     return undefined !== isExternalSharingEnabled ? isExternalSharingEnabled : false;
  //   }
  //   catch (error) {
  //     Log.error(LOG_SOURCE, new Error(`Failed to retrieve data from search REST API. Error: ${error}`));
  //     return false;
  //   }
  // }
}
